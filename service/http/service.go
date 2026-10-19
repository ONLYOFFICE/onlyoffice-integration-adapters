/**
 *
 * (c) Copyright Ascensio System SIA 2024
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */

// Package http provides an http service entry point.
//
// The http package's constructor is expected to be passed to a bootstrapper
// instance in order to be autoconfigured. A fully configured instance
// is ready to be used as a go-micro service.
package http

import (
	"context"
	"log"
	"net/http"
	"strings"
	"time"

	"github.com/ONLYOFFICE/onlyoffice-integration-adapters/config"
	plog "github.com/ONLYOFFICE/onlyoffice-integration-adapters/log"
	"github.com/ONLYOFFICE/onlyoffice-integration-adapters/messaging"
	"github.com/ONLYOFFICE/onlyoffice-integration-adapters/middleware"
	"github.com/ONLYOFFICE/onlyoffice-integration-adapters/resilience"
	chimiddleware "github.com/go-chi/chi/v5/middleware"
	"github.com/micro/plugins/v5/wrapper/breaker/hystrix"
	mclient "go-micro.dev/v5/client"
	"go-micro.dev/v5/registry"
	"go-micro.dev/v5/web"
	"go-micro.dev/v5/wrapper/trace/opentelemetry"
	"go.opentelemetry.io/otel"
	oteltrace "go.opentelemetry.io/otel/sdk/trace"
)

// ServerEngine provides main ways of interactions with
// any http server implementation (gorilla, chi, gin and so on).
type ServerEngine interface {
	ApplyMiddleware(middlewares ...func(http.Handler) http.Handler)
	NewHandler() interface {
		ServeHTTP(w http.ResponseWriter, req *http.Request)
	}
}

// A web.Service constructor. Called automatically by fx
// if provided to a bootstrapper instance.
//
// Returns a fully configured and ready to use http based web.Service.
// Panics on any error.
func NewService(
	engine ServerEngine,
	client mclient.Client,
	registry registry.Registry,
	broker messaging.BrokerWithOptions,
	tracer *oteltrace.TracerProvider,
	logger plog.Logger,
	serverConfig *config.ServerConfig,
	resilienceConfig *config.ResilienceConfig,
	corsConfig *config.CORSConfig,
	tracerConfig *config.TracerConfig,
) web.Service {
	if err := broker.Broker.Init(); err != nil {
		log.Fatalf("could not initialize a new broker instance: %s", err.Error())
	}

	if err := broker.Broker.Connect(); err != nil {
		log.Fatalf("broker connection error: %s", err.Error())
	}

	hystrix.ConfigureDefault(resilience.BuildHystrixCommandConfig(resilienceConfig))

	if err := client.Init(
		mclient.WrapCall(opentelemetry.NewCallWrapper(opentelemetry.WithTraceProvider(otel.GetTracerProvider()))),
	); err != nil {
		log.Fatalf("could not initialize a client instance: %s", err.Error())
	}

	if resilienceConfig.Resilience.RateLimiter.IPLimit > 0 {
		engine.ApplyMiddleware(middleware.NewRateLimiter(resilienceConfig.Resilience.RateLimiter.IPLimit, 1*time.Second, middleware.WithKeyFuncIP))
	}

	if resilienceConfig.Resilience.RateLimiter.Limit > 0 {
		engine.ApplyMiddleware(middleware.NewRateLimiter(resilienceConfig.Resilience.RateLimiter.Limit, 1*time.Second, middleware.WithKeyFuncAll))
	}

	engine.ApplyMiddleware(
		middleware.Log(logger),
		chimiddleware.RealIP,
		chimiddleware.RequestID,
		chimiddleware.StripSlashes,
		middleware.Version(serverConfig.Version),
		middleware.Cors(corsConfig.CORS.AllowedOrigins, corsConfig.CORS.AllowedMethods, corsConfig.CORS.AllowedHeaders, corsConfig.CORS.AllowCredentials),
	)

	if tracerConfig.Tracer.Enable {
		engine.ApplyMiddleware(
			middleware.TracePropagationMiddleware,
			middleware.LogTraceMiddleware,
		)
	}

	return web.NewService(
		web.Name(strings.Join([]string{serverConfig.Namespace, serverConfig.Name}, ":")),
		web.Version(serverConfig.Version),
		web.Address(serverConfig.Address),
		web.Context(context.Background()),
		web.Registry(registry),
		web.Handler(engine.NewHandler()),
		web.RegisterTTL(30*time.Second),
		web.RegisterInterval(10*time.Second),
		web.AfterStop(func() error {
			if tracer != nil {
				ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
				defer cancel()

				if err := tracer.Shutdown(ctx); err != nil {
					return err
				}

				return nil
			}

			return nil
		}),
	)
}

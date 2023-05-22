/**
 *
 * (c) Copyright Ascensio System SIA 2023
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

// Package config provides go-micro adapters' configuration structures
//
// The config package's structures are self-initialized by fx and bootstrapper.
// Fields are populated via yaml values or env variables. Env variables overwrite
// yaml configuration.
package config

import (
	"context"
	"os"
	"time"

	"github.com/sethvargo/go-envconfig"
	"gopkg.in/yaml.v2"
)

// A TracerConfig provides go-telemetry configuration for
// custom telemetry providers. This structure is expected to be
// initialized automatically by fx via yaml and env.
type TracerConfig struct {
	// Tracer is a nested structure used as a marker for yaml configuration.
	Tracer struct {
		// Name is tracer's name used to trace requests.
		Name string `yaml:"name" env:"TRACER_NAME,overwrite"`
		// Enable is flag to enable/disable tracing.
		Enable bool `yaml:"enable" env:"TRACER_ENABLE,overwrite"`
		// Address is tracer's instance address.
		Address string `yaml:"address" env:"TRACER_ADDRESS,overwrite"`
		// TracerType is a tracer provider selector.
		// 0 - Console.
		// 1 - Zipkin.
		//
		// By default - 0.
		TracerType    int     `yaml:"type" env:"TRACER_TYPE,overwrite"`
		FractionRatio float64 `yaml:"fraction" env:"TRACER_FRACTION_RATIO,overwrite"`
	} `yaml:"tracer"`
}

// Validate is called by fx and bootstrapper automatically after config initialization.
// It returns the first error encountered during validation.
//
// A successful Validate returns err == nil. Errors other than nil will
// cause application to panic
func (tc *TracerConfig) Validate() error {
	return nil
}

// A TracerConfig constructor. Called automatically by fx and
// bootstrapper with config path provided via cli.
//
// Returns a tracer configuration used to initialize a go-telemetry tracer provider
// and the first encountered error.
func BuildNewTracerConfig(path string) func() (*TracerConfig, error) {
	return func() (*TracerConfig, error) {
		var config TracerConfig
		config.Tracer.FractionRatio = 1
		if path != "" {
			file, err := os.Open(path)
			if err != nil {
				return nil, err
			}
			defer file.Close()

			decoder := yaml.NewDecoder(file)

			if err := decoder.Decode(&config); err != nil {
				return nil, err
			}
		}

		ctx, cancel := context.WithTimeout(context.Background(), 4*time.Second)
		defer cancel()
		if err := envconfig.Process(ctx, &config); err != nil {
			return nil, err
		}

		return &config, config.Validate()
	}
}

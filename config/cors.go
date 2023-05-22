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

// A CORSConfig provides configuration for cors middleware passed to a go-micro service.
// This structure is expected to be initialized automatically by fx via yaml and env.
type CORSConfig struct {
	// CORS is a nested structure used as a marker for yaml configuration.
	CORS struct {
		// AllowedOrigins is an http AllowedOrigins mapper
		//
		// By default - ["*"]
		AllowedOrigins []string `yaml:"origins" env:"ALLOWED_ORIGINS,overwrite"`
		// AllowedMethods is an http AllowedMethods mapper
		//
		// By default - ["*"]
		AllowedMethods []string `yaml:"methods" env:"ALLOWED_METHODS,overwrite"`
		// AllowedHeaders is an http AllowedHeaders mapper
		//
		// By default - ["*"]
		AllowedHeaders []string `yaml:"headers" env:"ALLOWED_HEADERS,overwrite"`
		// AllowCredentials is an http AllowCredentials mapper
		//
		// By default - false
		AllowCredentials bool `yaml:"credentials" env:"ALLOW_CREDENTIALS,overwrite"`
	} `yaml:"cors"`
}

// Validate is called by fx and bootstrapper automatically after config initialization.
// It returns the first error encountered during validation.
//
// A successful Validate returns err == nil. Errors other than nil will
// cause application to panic
func (cc *CORSConfig) Validate() error {
	return nil
}

// A CORSConfig constructor. Called automatically by fx and
// bootstrapper with config path provided via cli.
//
// Returns a cors configuration used to initialize a cors middleware
// and the first encountered error.
func BuildNewCorsConfig(path string) func() (*CORSConfig, error) {
	return func() (*CORSConfig, error) {
		var config CORSConfig
		config.CORS.AllowedOrigins = []string{"*"}
		config.CORS.AllowedMethods = []string{"*"}
		config.CORS.AllowedHeaders = []string{"*"}
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

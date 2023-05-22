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

// A WorkerConfig provides an async worker's configuration for
// custom worker providers. This structure is expected to be
// initialized automatically by fx via yaml and env.
type WorkerConfig struct {
	// Worker is a nested structure used as a marker for yaml configuration.
	Worker struct {
		// Enable is a worker's enable/disable flag.
		Enable bool `yaml:"enable" env:"WORKER_ENABLE,overwrite"`
		// Type is worker's implementation type.
		// 0 - Asynq worker.
		//
		// By default - 0.
		Type int `yaml:"type" env:"WORKER_TYPE,overwrite"`
		// MaxConcurrency is the maximum number of workers.
		MaxConcurrency int `yaml:"max_concurrency" env:"WORKER_MAX_CONCURRENCY,overwrite"`
		// RedisAddresses is redis instances addresses.
		RedisAddresses []string `yaml:"addresses" env:"WORKER_ADDRESS,overwrite"`
		// RedisUsername is redis basic auth username.
		RedisUsername string `yaml:"username" env:"WORKER_USERNAME,overwrite"`
		// RedisPassword is redis basic auth password.
		RedisPassword string `yaml:"password" env:"WORKER_PASSWORD,overwrite"`
		// RedisDatabase is worker's redis database for queueing.
		//
		// By default - 0.
		RedisDatabase int `yaml:"database" env:"WORKER_DATABASE,overwrite"`
	} `yaml:"worker"`
}

// Validate is called by fx and bootstrapper automatically after config initialization.
// It returns the first error encountered during validation.
//
// A successful Validate returns err == nil. Errors other than nil will
// cause application to panic
func (wc *WorkerConfig) Validate() error {
	if wc.Worker.Enable && len(wc.Worker.RedisAddresses) < 1 {
		return &InvalidConfigurationParameterError{
			Parameter: "Worker address",
			Reason:    "Should not be empty",
		}
	}

	return nil
}

// A WorkerConfig constructor. Called automatically by fx and
// bootstrapper with config path provided via cli.
//
// Returns a worker configuration used to initialize an async worker pool and
// the first encountered error.
func BuildNewWorkerConfig(path string) func() (*WorkerConfig, error) {
	return func() (*WorkerConfig, error) {
		var config WorkerConfig
		config.Worker.MaxConcurrency = 3
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

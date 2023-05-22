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

// A CacheConfig provides go-micro cache configuration for
// custom cache providers. This structure is expected to be
// initialized automatically by fx via yaml and env.
package config

import (
	"context"
	"os"
	"time"

	"github.com/sethvargo/go-envconfig"
	"gopkg.in/yaml.v2"
)

// A CacheConfig provides an entry point configuration for
// resilience patterns. This structure is expected to be
// initialized automatically by fx via yaml and env.
type ResilienceConfig struct {
	// Resilience is a nested structure used as a marker for yaml configuration.
	Resilience struct {
		// RateLimiter is a ratelimiter configuration.
		RateLimiter RateLimiterConfig `yaml:"rate_limiter"`
		// CircuitBreaker is a circuit breaker configuration.
		CircuitBreaker CircuitBreakerConfig `yaml:"circuit_breaker"`
	} `yaml:"resilience"`
}

// A ResilienceConfig constructor. Called automatically by fx and
// bootstrapper with config path provided via cli.
//
// Returns a resilience patterns configuration used to initialize a resilience middlewares
// and the first encountered error.
func BuildNewResilienceConfig(path string) func() (*ResilienceConfig, error) {
	return func() (*ResilienceConfig, error) {
		var config ResilienceConfig
		config.Resilience.RateLimiter.Limit = 3000
		config.Resilience.RateLimiter.IPLimit = 20
		config.Resilience.CircuitBreaker.Timeout = 5000
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

// A RateLimiterConfig provides rate-limiter's middleware configuration.
// This structure is expected to be initialized automatically by fx via yaml and env.
type RateLimiterConfig struct {
	// Limit is global requests limit cap
	//
	// By default - 3000
	Limit uint64 `yaml:"limit" env:"RATE_LIMIT,overwrite"`
	// IPLimit is ip specific requests limit cap
	//
	// By default - 20
	IPLimit uint64 `yaml:"iplimit" env:"RATE_LIMIT_IP,overwrite"`
}

// A CircuitBreakerConfig provides hystrix circuit breaker configuration.
// This structure is expected to be initialized automatically by fx via yaml and env.
type CircuitBreakerConfig struct {
	// Timeout is how long to wait for command to complete, in milliseconds
	//
	// By default - 1000
	Timeout int `yaml:"timeout" env:"CIRCUIT_TIMEOUT,overwrite"`
	// MaxConcurrent is how many commands of the same type can run at the same time
	//
	// By default - 10
	MaxConcurrent int `yaml:"max_concurrent" env:"CIRCUIT_MAX_CONCURRENT,overwrite"`
	// VolumeThreshold is the minimum number of requests needed before a circuit can be tripped due to health
	//
	// By default - 20
	VolumeThreshold int `yaml:"volume_threshold" env:"CIRCUIT_VOLUME_THRESHOLD,overwrite"`
	// SleepWindow is how long, in milliseconds, to wait after a circuit opens before testing for recovery
	//
	// By default - 5000
	SleepWindow int `yaml:"sleep_window" env:"CIRCUIT_SLEEP_WINDOW,overwrite"`
	// ErrorPercentThreshold causes circuits to open once the rolling measure of errors exceeds this percent of requests
	//
	// By default - 50
	ErrorPercentThreshold int `yaml:"error_percent_threshold" env:"CIRCUIT_ERROR_PERCENT_THRESHOLD,overwrite"`
}

// Validate is called by fx and bootstrapper automatically after config initialization.
// It returns the first error encountered during validation.
//
// A successful Validate returns err == nil. Errors other than nil will
// cause application to panic
func (rc *ResilienceConfig) Validate() error {
	return nil
}

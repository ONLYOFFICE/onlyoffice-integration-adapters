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

// A CacheConfig provides go-micro cache configuration for
// custom cache providers. This structure is expected to be
// initialized automatically by fx via yaml and env.
type CacheConfig struct {
	// Cache is a nested structure used as a marker for yaml configuration.
	Cache struct {
		// Type is gocache adapter type to be auto-configured.
		// 1 - Freecache.
		// 2 - Redis.
		//
		// By default - 1
		Type int `yaml:"type" env:"CACHE_TYPE,overwrite"`
		// Size is an optional field used to manually cache freecache
		// buffer size.
		//
		// By default - 10 * 1024 * 1024
		Size int `yaml:"size" env:"CACHE_SIZE,overwrite"`
		// Address is an optional field used to manually change redis
		// instance address.
		//
		// By default - 0.0.0.0:6379
		Address string `yaml:"address" env:"CACHE_ADDRESS,overwrite"`
		// Username is an optional field used to manually change redis
		// instance username
		//
		// By default - 'default'
		Username string `yaml:"username" env:"CACHE_USERNAME,overwrite"`
		// Password is an optional field used to manually cache redis
		// instance password
		//
		// By default - no password
		Password string `yaml:"password" env:"CACHE_PASSWORD,overwrite"`
		//
		Database int `yaml:"database" env:"CACHE_DATABASE,overwrite"`
	} `yaml:"cache"`
}

// Validate is called by fx and bootstrapper automatically after config initialization.
// It returns the first error encountered during validation.
//
// A successful Validate returns err == nil. Errors other than nil will
// cause application to panic
func (b *CacheConfig) Validate() error {
	switch b.Cache.Type {
	case 2:
		if b.Cache.Address == "" {
			return &InvalidConfigurationParameterError{
				Parameter: "Address",
				Reason:    "Redis cache must have a valid address",
			}
		}
		return nil
	default:
		return nil
	}
}

// A CacheConfig constructor. Called automatically by fx and
// bootstrapper with config path provided via cli.
//
// Returns a cache configuration used to initialize a go-micro cache
// and the first encountered error.
func BuildNewCacheConfig(path string) func() (*CacheConfig, error) {
	return func() (*CacheConfig, error) {
		var config CacheConfig
		config.Cache.Size = 10
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

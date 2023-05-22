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

// A RegistryConfig provides go-micro registry configuration for
// custom registry providers. This structure is expected to be
// initialized automatically by fx via yaml and env.
type RegistryConfig struct {
	// Registry is a nested structure used as a marker for yaml configuration.
	Registry struct {
		// Addresses is a list of registry instances.
		Addresses []string `yaml:"addresses" env:"REGISTRY_ADDRESSES,overwrite"`
		// CacheTTL is a service discovery interval.
		CacheTTL time.Duration `yaml:"cache_duration" env:"REGISTRY_CACHE_DURATION,overwrite"`
		// Type is a registry provider type
		// 1 - Kubernetes.
		// 2 - Consul.
		// 3 - ETCD.
		// 4 - MDNS.
		//
		// By default - 4
		Type int `yaml:"type" env:"REGISTRY_TYPE,overwrite"`
	} `yaml:"registry"`
}

// Validate is called by fx and bootstrapper automatically after config initialization.
// It returns the first error encountered during validation.
//
// A successful Validate returns err == nil. Errors other than nil will
// cause application to panic
func (r *RegistryConfig) Validate() error {
	switch r.Registry.Type {
	case 1:
		return nil
	default:
		if len(r.Registry.Addresses) <= 0 {
			return &InvalidConfigurationParameterError{
				Parameter: "Addresses",
				Reason:    "Length should be greater than zero",
			}
		}
		return nil
	}
}

// A RegistryConfig constructor. Called automatically by fx and
// bootstrapper with config path provided via cli.
//
// Returns a registry configuration used to initialize a go-micro registry
// and the first encountered error.
func BuildNewRegistryConfig(path string) func() (*RegistryConfig, error) {
	return func() (*RegistryConfig, error) {
		var config RegistryConfig
		config.Registry.CacheTTL = 10 * time.Second
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

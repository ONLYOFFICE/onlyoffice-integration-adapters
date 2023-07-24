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
	"strings"
	"time"

	"github.com/sethvargo/go-envconfig"
	"gopkg.in/yaml.v2"
)

// A ServerConfig provides go-micro service. This structure is expected to be
// initialized automatically by fx via yaml and env.
type ServerConfig struct {
	// Namespace is service specific namespace. Used to build requests to the other
	// services in the same namespace.
	Namespace string `yaml:"namespace" env:"SERVER_NAMESPACE,overwrite"`
	// Name is the name of the current service.
	Name string `yaml:"name" env:"SERVER_NAME,overwrite"`
	// Version is current version of the service. Used in version middleware
	// to set X-Version header.
	//
	// By default - 0.
	Version string `yaml:"version" env:"SERVER_VERSION,overwrite"`
	// Address is the service's address/port.
	Address string `yaml:"address" env:"SERVER_ADDRESS,overwrite"`
	// ReplAddress is system service's address.
	ReplAddress string `yaml:"repl_address" env:"REPL_ADDRESS,overwrite"`
	// Debug is flag to enable/disable debug features of the system's service.
	//
	// By default - false.
	Debug bool `yaml:"debug" env:"SERVER_DEBUG,overwrite"`
}

// Validate is called by fx and bootstrapper automatically after config initialization.
// It returns the first error encountered during validation.
//
// A successful Validate returns err == nil. Errors other than nil will
// cause application to panic
func (hs *ServerConfig) Validate() error {
	hs.Namespace = strings.TrimSpace(hs.Namespace)
	hs.Name = strings.TrimSpace(hs.Name)
	hs.Address = strings.TrimSpace(hs.Address)
	hs.ReplAddress = strings.TrimSpace(hs.ReplAddress)

	if hs.Namespace == "" {
		return &InvalidConfigurationParameterError{
			Parameter: "Namespace",
			Reason:    "Should not be empty",
		}
	}

	if hs.Name == "" {
		return &InvalidConfigurationParameterError{
			Parameter: "Name",
			Reason:    "Should not be empty",
		}
	}

	if hs.Address == "" {
		return &InvalidConfigurationParameterError{
			Parameter: "Address",
			Reason:    "Should not be empty",
		}
	}

	if hs.ReplAddress == "" {
		return &InvalidConfigurationParameterError{
			Parameter: "Repl Address",
			Reason:    "Should not be empty",
		}
	}

	return nil
}

// A ServerConfig constructor. Called automatically by fx and
// bootstrapper with config path provided via cli.
//
// Returns a server configuration used to initialize a go-micro service
// and the first encountered error.
func BuildNewServerConfig(path string) func() (*ServerConfig, error) {
	return func() (*ServerConfig, error) {
		var config ServerConfig
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

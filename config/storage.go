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

// A StorageConfig provides configuration for
// service level storage adapters. This structure is expected to be
// initialized automatically by fx via yaml and env.
// It is expected to be injected directly into services.
type StorageConfig struct {
	// Persistence is a nested structure used as a marker for yaml configuration.
	Storage struct {
		// Type is a persistence driver type.
		Type int `yaml:"type" env:"STORAGE_TYPE,overwrite"`
		// URL is a persistence driver adapter's url to connect to.
		URL string `yaml:"url" env:"STORAGE_URL,overwrite"`
		// DB is a database name to connect to.
		DB string `yaml:"db" env:"STORAGE_DB,overwrite"`
	} `yaml:"storage"`
}

// Validate is called by fx and bootstrapper automatically after config initialization.
// It returns the first error encountered during validation.
//
// A successful Validate returns err == nil. Errors other than nil will
// cause application to panic
func (p *StorageConfig) Validate() error {
	p.Storage.URL = strings.TrimSpace(p.Storage.URL)
	p.Storage.DB = strings.TrimSpace(p.Storage.DB)
	switch p.Storage.Type {
	case 1:
		if p.Storage.URL == "" {
			return &InvalidConfigurationParameterError{
				Parameter: "URL",
				Reason:    "MongoDB driver expects a valid url",
			}
		}
	default:
		if p.Storage.URL == "" {
			return &InvalidConfigurationParameterError{
				Parameter: "URL",
				Reason:    "MongoDB driver expects a valid url",
			}
		}
	}

	return nil
}

// A StorageConfig constructor. Called automatically by fx and
// bootstrapper with config path provided via cli.
//
// Returns a persistence configuration used inside adapters layers
// and the first encountered error.
func BuildNewStorageConfig(path string) func() (*StorageConfig, error) {
	return func() (*StorageConfig, error) {
		var config StorageConfig
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

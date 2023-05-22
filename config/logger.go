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

// A LoggerConfig provides go-micro logger configuration for
// custom logger providers. This structure is expected to be
// initialized automatically by fx via yaml and env.
type LoggerConfig struct {
	// Logger is a nested structure used as a marker for yaml configurations.
	Logger struct {
		// Name is used to configure logger name
		Name string `yaml:"name" env:"LOGGER_NAME,overwrite"`
		// Level is used to select logging level
		// 1 - Trace
		// 2 - Debug
		// 3 - Info
		// 4 - Warning
		// 5 - Error
		// 6 - Fatal
		//
		// By default - 4
		Level int `yaml:"level" env:"LOGGER_LEVEL,overwrite"`
		// Pretty is a flag used to enforce pretty log output
		//
		// By default - false
		Pretty bool `yaml:"pretty" env:"LOGGER_PRETTY,overwrite"`
		// Color is a flag used to enforce color indication of log levels
		//
		// By default - false
		Color bool `yaml:"color" env:"LOGGER_COLOR,overwrite"`
		// File is used to configure file log output
		//
		// By default - empty structure
		File FileLogConfig `yaml:"file"`
		// Elastic is used to configure elasticsearch log storage
		//
		// By default - empty structure
		Elastic ElasticLogConfig `yaml:"elastic"`
	} `yaml:"logger"`
}

// An ElasticLogConfig provides nested logger configuration for
// elastic logger providers. This structure is expected to be
// initialized automatically by fx via yaml and env.
type ElasticLogConfig struct {
	// Address is an elasticsearch instance address.
	Address string `yaml:"address" env:"ELASTIC_ADDRESS,overwrite"`
	// Index is an elasticsearch instance index.
	Index string `yaml:"index" env:"ELASTIC_INDEX,overwrite"`
	// Level is an elasticsearch instance logging level.
	Level int `yaml:"level" env:"ELASTIC_LEVEL,overwrite"`
	// Bulk is an elasticsearch instance bulk logging flag.
	Bulk bool `yaml:"bulk" env:"ELASTIC_BULK,overwrite"`
	// Async is an elasticsearch instance async logging flag.
	Async bool `yaml:"async" env:"ELASTIC_ASYNC,overwrite"`
	// HealthcheckEnabled is an elasticsearch instance flag to enable
	// periodical healthchecks.
	HealthcheckEnabled bool `yaml:"healthcheck" env:"ELASTIC_HEALTHCHECK,overwrite"`
	// BasicAuthUsername is an elasticsearch instance's basic auth username
	BasicAuthUsername string `yaml:"username" env:"ELASTIC_AUTH_USERNAME,overwrite"`
	// BasicAuthPassword is an elasticsearch instance's basic auth password
	BasicAuthPassword string `yaml:"password" env:"ELASTIC_AUTH_PASSWORD,overwrite"`
	// GzipEnabled is an elasticsearch instance flag to enable log zipping
	GzipEnabled bool `yaml:"gzip" env:"ELASTIC_GZIP_ENABLED,overwrite"`
}

// A FileLogConfig provides go-micro logger configuration for
// file logger providers. This structure is expected to be
// initialized automatically by fx via yaml and env.
type FileLogConfig struct {
	// Filename is a logging file to create/use to log to
	Filename string `yaml:"filename" env:"FILELOG_NAME,overwrite"`
	// MaxSize is the file's size limit before log rolling
	MaxSize int `yaml:"maxsize" env:"FILELOG_MAX_SIZE,overwrite"`
	// MaxAge is the file's age limit before log rolling
	MaxAge int `yaml:"maxage" env:"FILELOG_MAX_AGE,overwrite"`
	// MaxBackups is the maximum number of file's copies
	MaxBackups int `yaml:"maxbackups" env:"FILELOG_MAX_BACKUPS,overwrite"`
	// LocalTime is a flag to use local (server) time
	LocalTime bool `yaml:"localtime"`
	// Compress is a flag to compress log files on rolling
	Compress bool `yaml:"compress" env:"FILELOG_COMPRESS,overwrite"`
}

// Validate is called by fx and bootstrapper automatically after config initialization.
// It returns the first error encountered during validation.
//
// A successful Validate returns err == nil. Errors other than nil will
// cause application to panic
func (lc *LoggerConfig) Validate() error {
	return nil
}

// A LoggerConfig constructor. Called automatically by fx and
// bootstrapper with config path provided via cli.
//
// Returns a logger configuration used to initialize a new logger instance
// and the first encountered error.
func BuildNewLoggerConfig(path string) func() (*LoggerConfig, error) {
	return func() (*LoggerConfig, error) {
		var config LoggerConfig
		config.Logger.Name = "unknown"
		config.Logger.Level = 4
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

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

// A CryptoConfig provides configuration for
// built-in crypto providers. This structure is expected to be
// initialized automatically by fx via yaml and env.s
type CryptoConfig struct {
	// Crypto is a nested structure used as a marker for yaml configurations
	Crypto struct {
		// EncryptorType is an encryption algorithm type.
		// 1 - AES Gcm
		//
		// By default - 1
		EncryptorType int `yaml:"encryptor_type" env:"ENCRYPTOR_TYPE"`
		// JwtManagerType is a JWT library implementation type.
		// 1 - go-jwt/v5
		//
		// By default - 1
		JwtManagerType int `yaml:"jwt_manager_type" env:"JWT_MANAGER_TYPE"`
		// HasherType is a hash function implementation type.
		// 1 - md5
		//
		// By default - 1
		HasherType int `yaml:"hasher_type" env:"HASHER_TYPE"`
	} `yaml:"crypto"`
}

// A CryptoConfig constructor. Called automatically by fx and
// bootstrapper with config path provided via cli.
//
// Returns a crypto configuration used to initialize jwt, encryptor and hash managers
// and the first encountered error.
func BuildNewCryptoConfig(path string) func() (*CryptoConfig, error) {
	return func() (*CryptoConfig, error) {
		var config CryptoConfig
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

		return &config, nil
	}
}

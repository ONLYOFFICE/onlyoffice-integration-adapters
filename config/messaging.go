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

// A BrokerConfig provides go-micro broker configuration for
// custom broker providers. This structure is expected to be
// initialized automatically by fx via yaml and env.
type BrokerConfig struct {
	// Messaging is a nested structure used as a marker for yaml configuration.
	Messaging struct {
		// Enable is a broker's enable/disable flag.
		Enable bool `yaml:"enable" env:"BROKER_ENABLE,overwrite"`
		// Addrs is a list of broker instances.
		Addrs []string `yaml:"addresses" env:"BROKER_ADDRESSES,overwrite"`
		// Type is a broker type field.
		// 1 - RabbitMQ.
		// 2 - NATS.
		//
		// By default - Memory.
		Type int `yaml:"type" env:"BROKER_TYPE,overwrite"`
		// DisableAutoAck is an auto acknowledgement flag
		//
		// By default - false
		DisableAutoAck bool `yaml:"disable_auto_ack" env:"BROKER_DISABLE_AUTO_ACK,overwrite"`
		// Durable is a flag to make broker's queues durable
		//
		// By default - false
		Durable bool `yaml:"durable" env:"BROKER_DURABLE,overwrite"`
		// AckOnSuccess is an auto acknowledgement flag
		//
		// By default - false
		AckOnSuccess bool `yaml:"ack_on_success" env:"BROKER_ACK_ON_SUCCESS,overwrite"`
		// RequeueOnError is an auto requeue on error flag
		//
		// By default - false
		RequeueOnError bool `yaml:"requeue_on_error" env:"BROKER_REQUEUE_ON_ERROR,overwrite"`
	} `yaml:"messaging"`
}

// Validate is called by fx and bootstrapper automatically after config initialization.
// It returns the first error encountered during validation.
//
// A successful Validate returns err == nil. Errors other than nil will
// cause application to panic
func (b *BrokerConfig) Validate() error {
	if b.Messaging.Enable && len(b.Messaging.Addrs) == 0 && b.Messaging.Type > 0 {
		return &InvalidConfigurationParameterError{
			Parameter: "Addrs",
			Reason:    "Invalid number of addresses",
		}
	}

	return nil
}

// A BrokerConfig constructor. Called automatically by fx and
// bootstrapper with config path provided via cli.
//
// Returns a broker configuration used to initialize a go-micro broker
// and the first encountered error.
func BuildNewMessagingConfig(path string) func() (*BrokerConfig, error) {
	return func() (*BrokerConfig, error) {
		var config BrokerConfig
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

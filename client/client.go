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

// Package cache auto-injects and initializes a go-micro client for go-micro services
//
// The client package extracts from fx available registry and broker
// implementations initialized via yaml parameters or env variables.
// Client instance may be injected into any structure following fx injecton
// guidelines.
package client

import (
	"github.com/ONLYOFFICE/onlyoffice-integration-adapters/messaging"
	"go-micro.dev/v4/client"
	"go-micro.dev/v4/registry"
)

// A go-micro client.Client constructor. Called automatically by fx and
// bootstrapper.
//
// Returns a go-micro compliant client implementation based
// on registry and broker configuration. By default uses in-memory
// registry and broker implementations
func NewClient(
	registry registry.Registry, broker messaging.BrokerWithOptions,
) client.Client {
	return client.NewClient(
		client.Registry(registry),
		client.Broker(broker.Broker),
	)
}

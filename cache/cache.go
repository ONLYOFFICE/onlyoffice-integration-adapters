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

// Package cache provides caching adapters for go-micro
//
// The cache package should only be configured via yaml parameters or env variables.
// Cache instance should be accessed via micro client.Client and used to manually store
// and retreive cached values.
package cache

import (
	"context"
	"time"

	"github.com/ONLYOFFICE/onlyoffice-integration-adapters/config"
	"github.com/eko/gocache/lib/v4/marshaler"
	"github.com/eko/gocache/lib/v4/store"
	"go-micro.dev/v4/cache"
)

// A CustomCache provides go-micro compatible interface for
// custom cache providers. This structure is expected to be
// initialized automatically by fx.
type CustomCache struct {
	// Store is gocache provided marshaller.
	store *marshaler.Marshaler
	// Name is name for config based
	// initialization.
	name string
}

// Get retreives from a gocache provided store by key.
// It returns the value, extraction time and the first error
// encountered while extracting the value by key.
//
// A successful Get returns value != nil, time.Now() and err == nil.
func (c *CustomCache) Get(ctx context.Context, key string) (interface{}, time.Time, error) {
	var result interface{}
	_, err := c.store.Get(ctx, key, &result)
	return result, time.Now(), err
}

// Put stores into a gocache provided store by key, value and expiration date
// It returns the first error encountered while settings a new cache value.
//
// A successful Put returns err == nil.
func (c *CustomCache) Put(ctx context.Context, key string, val interface{}, d time.Duration) error {
	return c.store.Set(ctx, key, val, store.WithExpiration(d))
}

// Delete removes from a gocache provided store by key.
// It returns the first error encountered while removing a cache entry by key.
//
// A successful Delete returns err == nil.
func (c *CustomCache) Delete(ctx context.Context, key string) error {
	return c.store.Delete(ctx, key)
}

// String returns a gocache provided store name.
func (c *CustomCache) String() string {
	return c.name
}

// A CustomCache constructor. Called automatically by fx and
// bootstrapper.
//
// Returns a go-micro cache compliant implementation based
// on cache configuration. By default returns an in-memory
// implementation
func NewCache(config *config.CacheConfig) cache.Cache {
	switch config.Cache.Type {
	case 1:
		return &CustomCache{
			store: newMemory(config.Cache.Size),
			name:  "Freecache",
		}
	case 2:
		return &CustomCache{
			store: newRedis(
				config.Cache.Address, config.Cache.Username,
				config.Cache.Password, config.Cache.Database,
			),
			name: "Redis",
		}
	default:
		return &CustomCache{
			store: newMemory(config.Cache.Size),
			name:  "Freecache",
		}
	}
}

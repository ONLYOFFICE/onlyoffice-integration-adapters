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
	"time"

	"github.com/coocood/freecache"
	"github.com/eko/gocache/lib/v4/cache"
	"github.com/eko/gocache/lib/v4/marshaler"
	"github.com/eko/gocache/lib/v4/store"
	freecache_store "github.com/eko/gocache/store/freecache/v4"
)

// newMemory initializes an in-memory gocache store
// with the buffer size provided.
//
// Returns a new in-memory gocache compliant marshaler store
func newMemory(size int) *marshaler.Marshaler {
	freecacheStore := freecache_store.NewFreecache(
		freecache.NewCache(size*1024*1024),
		store.WithExpiration(10*time.Second),
	)
	cacheManage := cache.New[[]byte](freecacheStore)
	return marshaler.New(cacheManage.GetCodec().GetStore())
}

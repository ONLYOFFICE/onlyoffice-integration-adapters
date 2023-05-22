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
	"github.com/eko/gocache/lib/v4/cache"
	"github.com/eko/gocache/lib/v4/marshaler"
	redis_store "github.com/eko/gocache/store/redis/v4"
	"github.com/redis/go-redis/v9"
)

// newRedis initializes a redis gocache store
// with redis address, username, password and database
// credentials to establish a database connection
//
// Returns a new redis gocache compliant marshaler store
func newRedis(address, username, password string, db int) *marshaler.Marshaler {
	redisClient := redis.NewClient(&redis.Options{
		Username: username,
		Addr:     address,
		Password: password,
		DB:       db,
	})
	redisStore := redis_store.NewRedis(redisClient)
	cacheManager := cache.New[string](redisStore)
	marshaller := marshaler.New(cacheManager.GetCodec().GetStore())
	return marshaller
}

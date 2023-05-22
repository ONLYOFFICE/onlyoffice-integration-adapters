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

// Package storage provides a store wrapper over go-micro's store.Store and
// several implementations.
//
// The store package's structures are self-initialized by fx and bootstrapper.
// Fields are populated via yaml values or env variables. Env variables overwrite
// yaml configuration.
package storage

import (
	"context"

	"github.com/go-micro/plugins/v4/store/memory"
	"github.com/mitchellh/mapstructure"
	"go-micro.dev/v4/store"
)

type memoryStore struct {
	store store.Store
}

// A RefinedStore go-micro in-memory constructor. Called automatically by fx and
// bootstrapper.
func NewMemoryStore() RefinedStore {
	return &memoryStore{
		store: memory.NewStore(),
	}
}

// Initialize an in-memory adapter
func (s *memoryStore) Init(opts ...store.Option) error {
	return s.store.Init(opts...)
}

// List all the keys
func (s *memoryStore) List(ctx context.Context, opts ...ReadOption) error {
	var ops ReadOptions
	for _, o := range opts {
		o(&ops)
	}

	res, err := s.store.List(
		store.ListFrom(ops.Database, ops.Table),
		store.ListLimit(ops.Limit),
		store.ListOffset(ops.Offset),
		store.ListPrefix(ops.Prefix),
		store.ListSuffix(ops.Suffix),
	)

	if err != nil {
		return err
	}

	return mapstructure.Decode(res, ops.Result)
}

// Read a single record
func (s *memoryStore) Read(ctx context.Context, opts ...ReadOption) error {
	var ops ReadOptions
	for _, o := range opts {
		o(&ops)
	}

	res, err := s.store.Read(
		ops.Key,
		store.ReadFrom(ops.Database, ops.Table),
		store.ReadLimit(ops.Limit),
		store.ReadOffset(ops.Offset),
	)

	if err != nil {
		return err
	}

	return mapstructure.Decode(res, ops.Result)
}

// Write a record
func (s *memoryStore) Write(ctx context.Context, payload any, opts ...WriteOption) error {
	var ops WriteOptions
	for _, o := range opts {
		o(&ops)
	}

	p, ok := payload.(*store.Record)
	if !ok {
		return _errInvalidResultOption
	}

	return s.store.Write(
		p, store.WriteTo(ops.Database, ops.Table),
		store.WriteExpiry(ops.Expiry),
		store.WriteTTL(ops.TTL),
	)
}

// Write duplicate
func (s *memoryStore) Update(ctx context.Context, payload any, opts ...WriteOption) error {
	return s.Write(ctx, payload, opts...)
}

// Delete a record with a key
func (s *memoryStore) Delete(ctx context.Context, opts ...DeleteOption) error {
	var ops DeleteOptions
	for _, o := range opts {
		o(&ops)
	}

	return s.store.Delete(
		ops.Key,
		store.DeleteFrom(ops.Database, ops.Table),
	)
}

// Returns db options
func (s *memoryStore) Options() store.Options {
	return s.store.Options()
}

// Returns adapter name
func (s *memoryStore) String() string {
	return s.store.String()
}

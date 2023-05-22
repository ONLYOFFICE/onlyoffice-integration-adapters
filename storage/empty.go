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

	"go-micro.dev/v4/store"
)

type emptyStore struct {
}

// A RefinedStore mongo constructor. Called automatically by fx and
// bootstrapper.
func NewEmptyStore() RefinedStore {
	return &emptyStore{}
}

func (s *emptyStore) Init(opts ...store.Option) error {
	return nil
}

// List all the known records.
func (s *emptyStore) List(ctx context.Context, opts ...ReadOption) error {
	return nil
}

// Read a single key.
func (s *emptyStore) Read(ctx context.Context, opts ...ReadOption) error {
	return nil
}

// Write records.
func (s *emptyStore) Write(ctx context.Context, payload any, opts ...WriteOption) error {
	return nil
}

// Update records
func (s *emptyStore) Update(ctx context.Context, payload any, opts ...WriteOption) error {
	return nil
}

// Delete records with keys.
func (s *emptyStore) Delete(ctx context.Context, opts ...DeleteOption) error {
	return nil
}

// Returns db options.
func (s *emptyStore) Options() store.Options {
	return store.Options{}
}

// Returns adapter name.
func (s *emptyStore) String() string {
	return "empty"
}

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
// TODO: Refactor the interface
package storage

import (
	"context"
	"log"
	"time"

	"github.com/ONLYOFFICE/onlyoffice-integration-adapters/config"
	"go-micro.dev/v4/store"
)

type ReadOption func(l *ReadOptions)

type ReadOptions struct {
	// List from the following.
	Database, Table string
	// Key to read.
	Key string
	// Value to read (optional).
	Value string
	// Prefix returns all keys that are prefixed with key.
	Prefix string
	// Suffix returns all keys that end with key.
	Suffix string
	// Limit limits the number of returned keys.
	Limit uint
	// Offset when combined with Limit supports pagination.
	Offset uint
	// Result from the executed query.
	Result any
}

// Sets database and database table read options.
func ReadFrom(database, table string) ReadOption {
	return func(l *ReadOptions) {
		l.Database = database
		l.Table = table
	}
}

// Sets a filter field name.
func ReadKey(val string) ReadOption {
	return func(l *ReadOptions) {
		l.Key = val
	}
}

// Sets a filter field's name's value.
func ReadValue(val string) ReadOption {
	return func(l *ReadOptions) {
		l.Value = val
	}
}

// Sets a filter prefix value.
func ReadPrefix(val string) ReadOption {
	return func(l *ReadOptions) {
		l.Prefix = val
	}
}

// Sets a filter suffix value.
func ReadSuffix(val string) ReadOption {
	return func(l *ReadOptions) {
		l.Suffix = val
	}
}

// Sets select limit value.
func ReadLimit(val uint) ReadOption {
	return func(l *ReadOptions) {
		l.Limit = val
	}
}

// Sets select skip value.
func ReadOffset(val uint) ReadOption {
	return func(l *ReadOptions) {
		l.Offset = val
	}
}

// Sets a pointer to populate it with the result.
func ReadResult(val any) ReadOption {
	return func(l *ReadOptions) {
		l.Result = val
	}
}

type WriteOption func(w *WriteOptions)

// WriteOptions configures an individual Write operation
// If Expiry and TTL are set TTL takes precedence.
type WriteOptions struct {
	Database, Table string
	// Key is the key name to find and update the record (optional).
	Key string
	// Value is the key value to finad and update the record (optional).
	Value string
	// Expiry is the time the record expires.
	Expiry time.Time
	// TTL is the time until the record expires.
	TTL time.Duration
}

// Sets database and database table.
func WriteTo(database, table string) WriteOption {
	return func(w *WriteOptions) {
		w.Database = database
		w.Table = table
	}
}

// Sets a filter field name.
func WriteKey(val string) WriteOption {
	return func(w *WriteOptions) {
		w.Key = val
	}
}

// Sets a filter field's name's value.
func WriteValue(val string) WriteOption {
	return func(w *WriteOptions) {
		w.Value = val
	}
}

// Sets payload expiration.
func WriteExpiry(val time.Time) WriteOption {
	return func(w *WriteOptions) {
		w.Expiry = val
	}
}

// Sets payload ttl.
func WriteTTL(val time.Duration) WriteOption {
	return func(w *WriteOptions) {
		w.TTL = val
	}
}

// DeleteOptions configures an individual Delete operation.
type DeleteOptions struct {
	Database, Table string
	Key             string
	Value           string
}

// DeleteOption sets values in DeleteOptions.
type DeleteOption func(d *DeleteOptions)

// Sets database and database table.
func DeleteFrom(database, table string) DeleteOption {
	return func(d *DeleteOptions) {
		d.Database = database
		d.Table = table
	}
}

// Sets filter's field name.
func DeleteKey(val string) DeleteOption {
	return func(d *DeleteOptions) {
		d.Key = val
	}
}

// Sets filter's field value.
func DeleteValue(val string) DeleteOption {
	return func(d *DeleteOptions) {
		d.Value = val
	}
}

// RefinedStore is a go-micro store.Store wrapper
// to allow SQL and No-SQL database operations.
type RefinedStore interface {
	Init(opts ...store.Option) error
	List(ctx context.Context, opts ...ReadOption) error
	Read(ctx context.Context, opts ...ReadOption) error
	Write(ctx context.Context, payload any, opts ...WriteOption) error
	Update(ctx context.Context, payload any, opts ...WriteOption) error
	Delete(ctx context.Context, opts ...DeleteOption) error
	Options() store.Options
	String() string
}

// A RefinedStore constructor. Called automatically by fx and
// bootstrapper.
//
// Returns a RefinedStore compliant implementation based
// on persistence configuration.
//
// By default - empty adapter
func NewStorage(config *config.StorageConfig) RefinedStore {
	var s RefinedStore
	switch config.Storage.Type {
	case 1:
		s = NewMongoStore()
	default:
		s = NewEmptyStore()
	}

	if err := s.Init(
		store.Database(config.Storage.DB),
		store.Nodes(config.Storage.URL),
	); err != nil {
		log.Fatalln(err.Error())
	}

	return s
}

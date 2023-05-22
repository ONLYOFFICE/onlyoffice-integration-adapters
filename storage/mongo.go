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
	"errors"
	"time"

	"github.com/kamva/mgm/v3"
	"go-micro.dev/v4/store"
	"go.mongodb.org/mongo-driver/bson"
	"go.mongodb.org/mongo-driver/mongo"
	"go.mongodb.org/mongo-driver/mongo/options"
)

var _errInvalidResultOption = errors.New("expected to get a non-nil result option")

type mongoStore struct {
	options store.Options
}

// A RefinedStore mongo constructor. Called automatically by fx and
// bootstrapper.
func NewMongoStore() RefinedStore {
	return &mongoStore{}
}

func (s *mongoStore) configure() error {
	return mgm.SetDefaultConfig(
		&mgm.Config{CtxTimeout: 3 * time.Second}, s.options.Database,
		options.Client().ApplyURI(s.options.Nodes[0]),
	)
}

func (s *mongoStore) Init(opts ...store.Option) error {
	for _, o := range opts {
		o(&s.options)
	}

	return s.configure()
}

// List all the known records.
func (s *mongoStore) List(ctx context.Context, opts ...ReadOption) error {
	var ops ReadOptions
	for _, o := range opts {
		o(&ops)
	}

	col := mgm.CollectionByName(ops.Table)
	cur, err := col.Find(
		ctx, bson.D{},
		options.Find().SetSkip(int64(ops.Offset)).SetLimit(int64(ops.Limit)),
	)

	if err != nil {
		return err
	}

	if ops.Result == nil {
		return _errInvalidResultOption
	}

	if err := cur.All(ctx, ops.Result); err != nil {
		return err
	}

	return cur.Close(ctx)
}

// Read a single key.
func (s *mongoStore) Read(ctx context.Context, opts ...ReadOption) error {
	var options ReadOptions
	for _, o := range opts {
		o(&options)
	}

	col := mgm.CollectionByName(options.Table)
	sres := col.FindOne(ctx, bson.M{
		options.Key: options.Value,
	})

	if options.Result == nil {
		return _errInvalidResultOption
	}

	if err := sres.Decode(options.Result); err != nil {
		return err
	}

	return nil
}

// Write records.
func (s *mongoStore) Write(ctx context.Context, payload any, opts ...WriteOption) error {
	var options WriteOptions
	for _, o := range opts {
		o(&options)
	}

	return mgm.TransactionWithCtx(ctx, func(session mongo.Session, sc mongo.SessionContext) error {
		col := mgm.CollectionByName(options.Table)
		if _, err := col.InsertOne(ctx, payload); err != nil {
			return err
		}

		return session.CommitTransaction(sc)
	})
}

// Update records
func (s *mongoStore) Update(ctx context.Context, payload any, opts ...WriteOption) error {
	var options WriteOptions
	for _, o := range opts {
		o(&options)
	}

	return mgm.TransactionWithCtx(ctx, func(session mongo.Session, sc mongo.SessionContext) error {
		col := mgm.CollectionByName(options.Table)
		filter := bson.D{{Key: options.Key, Value: options.Value}}
		update := bson.D{{Key: "$set", Value: payload}}
		if _, err := col.UpdateOne(ctx, filter, update); err != nil {
			return err
		}

		return session.CommitTransaction(sc)
	})
}

// Delete records with keys.
func (s *mongoStore) Delete(ctx context.Context, opts ...DeleteOption) error {
	var options DeleteOptions
	for _, o := range opts {
		o(&options)
	}

	ctx, cancel := context.WithTimeout(context.Background(), 3*time.Second)
	defer cancel()
	return mgm.TransactionWithCtx(ctx, func(session mongo.Session, sc mongo.SessionContext) error {
		col := mgm.CollectionByName(options.Table)
		if _, err := col.DeleteOne(ctx, bson.M{
			options.Key: options.Value,
		}); err != nil {
			return err
		}

		return session.CommitTransaction(sc)
	})
}

// Returns db options.
func (s *mongoStore) Options() store.Options {
	return s.options
}

// Returns adapter name.
func (s *mongoStore) String() string {
	return "mongodb"
}

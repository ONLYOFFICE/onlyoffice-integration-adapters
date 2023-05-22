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

// Package events provides emitter adapters for services
//
// The events package's structures are self-initialized by fx and bootstrapper.
// Fields are populated via yaml values or env variables. Env variables overwrite
// yaml configuration.
package events

// An Event provides basic contracts for event handling.
// The implementation structure is expected to be initialized automatically by fx
// and bootstrapper.
type Event interface {
	// Name returns event name.
	Name() string
	// Get returns a payload by its key.
	Get(key string) any
	// Add adds a payload by its key.
	Add(key string, val any)
	// Abort interrupts event handling.
	Abort(bool)
	// IsAborted returns aborted flag.
	IsAborted() bool
}

// A Listener provides basic contracts for event listeners.
// The implementation structure is expected to be initialized automatically by fx
// and bootstrapper
type Listener interface {
	// Handle is an entry point for event handling.
	// Returns the first encountered error.
	//
	// A successful Handle return err == nil.
	Handle(e Event) error
}

// An Emitter provides basic contracts for event publishing/subscribing.
// The implementation structure is expected to be initialized automatically by fx
// and bootstrapper
type Emitter interface {
	// On is a subscription mechanism.
	// Takes an event name and a handler to process that event.
	On(name string, listener Listener)
	// Fire is a publication mechanism.
	// Takes an event name and a payload to be processed.
	Fire(name string, payload map[string]any)
}

// An Emitter constructor. Called automatically by fx and
// bootstrapper.
//
// Returns an emitter implementation based on configuration.
// By default returns a gokit emitter.
func NewEmitter() Emitter {
	return NewGoKitEmitter()
}

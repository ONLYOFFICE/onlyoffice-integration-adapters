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

// Package functions provides functional convenience structures
//
// The functional package should  be configured manually unlike the other packages from the module.
package functional

type action[T any] func(input T) (T, error)

// Pipe is a utility structure for functions composition.
type Pipe[T any] struct {
	chain []action[T]
}

// NewPipe initializes a new pipe for functions composition.
func NewPipe[T any]() *Pipe[T] {
	return &Pipe[T]{}
}

// Next appends a new handler function to the pipe.
func (p *Pipe[T]) Next(f action[T]) *Pipe[T] {
	p.chain = append(p.chain, f)
	return p
}

// Do starts chain execution.
func (p *Pipe[T]) Do() (T, error) {
	var res T
	var err error
	for _, fn := range p.chain {
		res, err = fn(res)
		if err != nil {
			break
		}
	}

	return res, err
}

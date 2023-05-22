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

// Package crypto provides basic cryptography wrappers and implementations for
// encryption, token management and hashing.
//
// The crypto package's structures are self-initialized by fx and bootstrapper.
// Fields are populated via yaml values or env variables. Env variables overwrite
// yaml configuration.
package crypto

import (
	"crypto/md5"
	"encoding/hex"
)

// md5Hasher is a basic Hasher implementation
type md5Hasher struct{}

// A Hasher constructor. Called automatically by fx and
// bootstrapper.
//
// Returns a Hasher compliant implementation based
// on cache configuration.
func newMD5Hasher() Hasher {
	return md5Hasher{}
}

// Hash transforms plaintext into a hashed text.
// It returns a hashed string.
//
// A successful Hash return a non-empty string.
func (h md5Hasher) Hash(text string) string {
	hash := md5.Sum([]byte(text))
	return hex.EncodeToString(hash[:])
}

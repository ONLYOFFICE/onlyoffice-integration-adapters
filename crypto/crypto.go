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
	"github.com/ONLYOFFICE/onlyoffice-integration-adapters/config"
	"github.com/golang-jwt/jwt/v5"
)

// An Encryptor provides basic contract for encryption types.
// The implementation structure is expected to be initialized automatically by fx and bootstrapper.
type Encryptor interface {
	Encrypt(text string, key []byte) (string, error)
	Decrypt(ciphertext string, key []byte) (string, error)
}

// An Encryptor constructor. Called automatically by fx and
// bootstrapper.
//
// Returns an encryptor implementation based
// on configuration. By default returns an AES GCM encryptor.
func NewEncryptor(config *config.CryptoConfig) Encryptor {
	switch config.Crypto.EncryptorType {
	case 1:
		return newAesEncryptor()
	default:
		return newAesEncryptor()
	}
}

// A JwtManager provides basic contract for jwt generation and verification.
// The implementation structure is expected to be intialized automatically by fx and bootstrapper.
type JwtManager interface {
	Sign(secret string, payload jwt.Claims) (string, error)
	Verify(secret, jwtToken string, body interface{}) error
}

// A JwtManager constructor. Called automatically by fx and
// bootstrapper.
//
// Returns a jwt manager implementation based
// on configuration.
func NewJwtManager(config *config.CryptoConfig) JwtManager {
	switch config.Crypto.JwtManagerType {
	case 1:
		return newOnlyofficeJwtManager()
	default:
		return newOnlyofficeJwtManager()
	}
}

// A Hasher provides basic contract for generating hash.
// The implementation structure is expected to be intialized automatically by fx and bootstrapper.
type Hasher interface {
	Hash(text string) string
}

// A Hasher constructor. Called automatically by fx and
// bootstrapper.
//
// Returns a hasher implementation based on configuration.
// By default returns an md5 hasher.
func NewHasher(config *config.CryptoConfig) Hasher {
	switch config.Crypto.HasherType {
	case 1:
		return newMD5Hasher()
	default:
		return newMD5Hasher()
	}
}

// A StateGenerator provides basic contract for generating a cryptographic state.
// The implementation structure is expected to be intialized automatically by fx and bootstrapper.
type StateGenerator interface {
	GenerateState(secret string) (string, error)
}

// A StateGenerator constructor. Called automatically by fx and
// bootstrapper.
//
// Returns a state generator implementation based on configuration.
func NewStateGenerator() StateGenerator {
	return newStateGenerator()
}

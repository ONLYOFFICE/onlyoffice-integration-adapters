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
	"crypto/hmac"
	"crypto/rand"
	"crypto/sha256"
	"encoding/base64"
	"encoding/hex"
	"net/url"
	"strings"
)

// stateGenerator is a basic StateGenerator implementation.
type stateGenerator struct {
}

func newStateGenerator() StateGenerator {
	return stateGenerator{}
}

// GenerateState takes a secret and generates an oauth2 state.
// It returns a newly generated state and the first encountered error.
//
// A successful GenerateState returns a state and err == nil.
func (sg stateGenerator) GenerateState(secret string) (string, error) {
	ts, err := randomHex(64)
	if err != nil {
		return "", err
	}

	hmac, err := hmacBase64(ts, secret)
	if err != nil {
		return "", err
	}

	return url.QueryEscape(strings.ReplaceAll(strings.Join([]string{hmac, ts}, "."), "+", "")), nil
}

// randomHex takes a buffer's length and outputs a random hex string.
// It returns a random hex string and the first encountered error.
//
// A successful randomHex returns a non-empty string and err == nil.
func randomHex(n int) (string, error) {
	bytes := make([]byte, n)
	if _, err := rand.Read(bytes); err != nil {
		return "", err
	}
	return hex.EncodeToString(bytes), nil
}

// hmacBase64 takes a plaintext and a secret and transforms them into a base64 hash.
// It returns a base64 hash and the first encountered error.
//
// A successful hmacBase64 returns a non-empty string and err == nil.
func hmacBase64(message string, secret string) (string, error) {
	key := []byte(secret)
	h := hmac.New(sha256.New, key)

	if _, err := h.Write([]byte(message)); err != nil {
		return "", err
	}

	return base64.StdEncoding.EncodeToString(h.Sum(nil)), nil
}

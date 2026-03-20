package client

import (
	"crypto/x509"
	"encoding/pem"
	"os"
	"path/filepath"

	"github.com/pkg/errors"
)

func CertPoolFromFile(filename string) (*x509.CertPool, error) {
	certs, err := certificatesFromFile(filename)
	if err != nil {
		return nil, err
	}

	pool := x509.NewCertPool()
	for _, cert := range certs {
		pool.AddCert(cert)
	}

	return pool, nil
}

func certificatesFromFile(file string) ([]*x509.Certificate, error) {
	if len(file) == 0 {
		return nil, errors.New("error reading certificates from an empty filename")
	}

	pemBlock, err := os.ReadFile(filepath.Clean(file))
	if err != nil {
		return nil, err
	}

	certs, err := CertsFromPEM(pemBlock)
	if err != nil {
		return nil, errors.Wrapf(err, "error reading %s", file)
	}

	return certs, nil
}

func CertsFromPEM(pemCerts []byte) ([]*x509.Certificate, error) {
	ok := false
	certs := []*x509.Certificate{}

	for len(pemCerts) > 0 {
		var block *pem.Block
		block, pemCerts = pem.Decode(pemCerts)

		if block == nil {
			break
		}

		if block.Type != "CERTIFICATE" || len(block.Headers) != 0 {
			continue
		}

		cert, err := x509.ParseCertificate(block.Bytes)
		if err != nil {
			return certs, err
		}

		certs = append(certs, cert)
		ok = true
	}

	if !ok {
		return certs, errors.New("could not read any certificates")
	}

	return certs, nil
}

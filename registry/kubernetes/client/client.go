package client

import (
	"crypto/tls"
	"errors"
	"fmt"
	"net/http"
	"os"
	"path"
	"strings"
	"sync"

	"go-micro.dev/v4/logger"

	"github.com/ONLYOFFICE/onlyoffice-integration-adapters/registry/kubernetes/client/api"
	"github.com/ONLYOFFICE/onlyoffice-integration-adapters/registry/kubernetes/client/watch"
)

var (
	serviceAccountPath = "/var/run/secrets/kubernetes.io/serviceaccount"

	ErrReadNamespace = errors.New("could not read namespace from service account secret")
)

type tokenFileTransport struct {
	tokenPath string
	wrapped   http.RoundTripper

	mu    sync.RWMutex
	token string
}

func (t *tokenFileTransport) cachedToken() string {
	t.mu.RLock()
	defer t.mu.RUnlock()
	return t.token
}

func (t *tokenFileTransport) refreshToken() (string, error) {
	raw, err := os.ReadFile(t.tokenPath)
	if err != nil {
		return "", fmt.Errorf("kubernetes: failed to refresh service account token: %w", err)
	}

	tok := strings.TrimSpace(string(raw))

	t.mu.Lock()
	t.token = tok
	t.mu.Unlock()

	return tok, nil
}

func (t *tokenFileTransport) RoundTrip(req *http.Request) (*http.Response, error) {
	r2 := req.Clone(req.Context())
	r2.Header.Set("Authorization", "Bearer "+t.cachedToken())

	resp, err := t.wrapped.RoundTrip(r2)
	if err != nil {
		return nil, err
	}

	if resp.StatusCode != http.StatusUnauthorized {
		return resp, nil
	}

	resp.Body.Close()

	newTok, err := t.refreshToken()
	if err != nil {
		return nil, err
	}

	r3 := req.Clone(req.Context())
	r3.Header.Set("Authorization", "Bearer "+newTok)

	if req.GetBody != nil {
		body, err := req.GetBody()
		if err != nil {
			return nil, fmt.Errorf("kubernetes: failed to rebuild request body for retry: %w", err)
		}
		r3.Body = body
	}

	return t.wrapped.RoundTrip(r3)
}

type client struct {
	opts *api.Options
}

func NewClientByHost(host string) Kubernetes {
	tr := &http.Transport{
		TLSClientConfig: &tls.Config{
			//nolint:gosec
			InsecureSkipVerify: true,
		},
		DisableCompression: true,
	}

	c := &http.Client{
		Transport: tr,
	}

	return &client{
		opts: &api.Options{
			Client:    c,
			Host:      host,
			Namespace: "default",
		},
	}
}

func NewClientInCluster() Kubernetes {
	host := "https://" + os.Getenv("KUBERNETES_SERVICE_HOST") + ":" + os.Getenv("KUBERNETES_SERVICE_PORT")

	s, err := os.Stat(serviceAccountPath)
	if err != nil {
		logger.Fatal(err)
	}

	if s == nil || !s.IsDir() {
		logger.Fatal(errors.New("no k8s service account found"))
	}

	tokenPath := path.Join(serviceAccountPath, "token")

	t, err := os.ReadFile(tokenPath)
	if err != nil {
		logger.Fatal(err)
	}

	ns, err := detectNamespace()
	if err != nil {
		logger.Fatal(err)
	}

	crt, err := CertPoolFromFile(path.Join(serviceAccountPath, "ca.crt"))
	if err != nil {
		logger.Fatal(err)
	}

	c := &http.Client{
		Transport: &tokenFileTransport{
			tokenPath: tokenPath,
			token:     strings.TrimSpace(string(t)),
			wrapped: &http.Transport{
				TLSClientConfig: &tls.Config{
					RootCAs:    crt,
					MinVersion: tls.VersionTLS12,
				},
				DisableCompression: true,
			},
		},
	}

	return &client{
		opts: &api.Options{
			Client:    c,
			Host:      host,
			Namespace: ns,
		},
	}
}

func (c *client) ListPods(labels map[string]string) (*PodList, error) {
	var pods PodList
	err := api.NewRequest(c.opts).Get().Resource("pods").Params(&api.Params{LabelSelector: labels}).Do().Decode(&pods)

	return &pods, err
}

func (c *client) UpdatePod(name string, p *Pod) (*Pod, error) {
	var pod Pod
	err := api.NewRequest(c.opts).Patch().Resource("pods").Name(name).Body(p).Do().Decode(&pod)

	return &pod, err
}

func (c *client) WatchPods(labels map[string]string) (watch.Watch, error) {
	return api.NewRequest(c.opts).Get().Resource("pods").Params(&api.Params{LabelSelector: labels}).Watch()
}

func detectNamespace() (string, error) {
	nsPath := path.Join(serviceAccountPath, "namespace")

	if s, err := os.Stat(nsPath); err != nil {
		return "", err
	} else if s.IsDir() {
		return "", ErrReadNamespace
	}

	ns, err := os.ReadFile(path.Clean(nsPath))
	if err != nil {
		return string(ns), err
	}

	return string(ns), nil
}

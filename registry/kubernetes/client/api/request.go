package api

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"net/url"

	"github.com/ONLYOFFICE/onlyoffice-integration-adapters/registry/kubernetes/client/watch"
)

type Request struct {
	client    *http.Client
	header    http.Header
	params    url.Values
	method    string
	host      string
	namespace string

	resource     string
	resourceName *string
	body         io.Reader

	err error
}

type Params struct {
	LabelSelector map[string]string
	Watch         bool
}

type Options struct {
	Host        string
	Namespace   string
	BearerToken *string
	Client      *http.Client
}

func NewRequest(opts *Options) *Request {
	req := Request{
		header:    make(http.Header),
		params:    make(url.Values),
		client:    opts.Client,
		namespace: opts.Namespace,
		host:      opts.Host,
	}

	if opts.BearerToken != nil {
		req.SetHeader("Authorization", "Bearer "+*opts.BearerToken)
	}

	return &req
}

func (r *Request) verb(method string) *Request {
	r.method = method
	return r
}

func (r *Request) Get() *Request {
	return r.verb("GET")
}

func (r *Request) Post() *Request {
	return r.verb("POST")
}

func (r *Request) Put() *Request {
	return r.verb("PUT")
}

func (r *Request) Patch() *Request {
	return r.verb("PATCH").SetHeader("Content-Type", "application/strategic-merge-patch+json")
}

func (r *Request) Delete() *Request {
	return r.verb("DELETE")
}

func (r *Request) Namespace(s string) *Request {
	r.namespace = s
	return r
}

func (r *Request) Resource(s string) *Request {
	r.resource = s
	return r
}

func (r *Request) Name(s string) *Request {
	r.resourceName = &s
	return r
}

func (r *Request) Body(in interface{}) *Request {
	b := new(bytes.Buffer)
	if err := json.NewEncoder(b).Encode(&in); err != nil {
		r.err = err
		return r
	}

	r.body = b

	return r
}

func (r *Request) Params(p *Params) *Request {
	for k, v := range p.LabelSelector {
		value := fmt.Sprintf("%s=%s", k, v)
		if label := r.params.Get("labelSelector"); len(label) > 0 {
			value = fmt.Sprintf("%s,%s", label, value)
		}
		r.params.Set("labelSelector", value)
	}

	return r
}

func (r *Request) SetHeader(key, value string) *Request {
	r.header.Add(key, value)
	return r
}

func (r *Request) request() (*http.Request, error) {
	rawURL := fmt.Sprintf("%s/api/v1/namespaces/%s/%s/", r.host, r.namespace, r.resource)

	if r.resourceName != nil {
		rawURL += *r.resourceName
	}

	if len(r.params) > 0 {
		rawURL += "?" + r.params.Encode()
	}

	req, err := http.NewRequest(r.method, rawURL, r.body)
	if err != nil {
		return nil, err
	}

	req.Header = r.header

	return req, nil
}

func (r *Request) Do() *Response {
	if r.err != nil {
		return &Response{err: r.err}
	}

	req, err := r.request()
	if err != nil {
		return &Response{err: err}
	}

	res, err := r.client.Do(req)
	if err != nil {
		return &Response{err: err}
	}

	return newResponse(res, err)
}

func (r *Request) Watch() (watch.Watch, error) {
	if r.err != nil {
		return nil, r.err
	}

	r.params.Set("watch", "true")

	req, err := r.request()
	if err != nil {
		return nil, err
	}

	return watch.NewBodyWatcher(req, r.client)
}

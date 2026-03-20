package api

import (
	"encoding/json"
	"io"
	"net/http"

	"github.com/pkg/errors"
	log "go-micro.dev/v4/logger"
)

var (
	ErrNoPodName = errors.New("no pod name provided")
	ErrNotFound  = errors.New("pod not found")
	ErrDecode    = errors.New("error decoding")
	ErrOther     = errors.New("unspecified error occurred in k8s registry")
)

type Response struct {
	res *http.Response
	err error
}

func (r *Response) Error() error {
	return r.err
}

func (r *Response) StatusCode() int {
	return r.res.StatusCode
}

func (r *Response) Decode(data interface{}) error {
	if r.err != nil {
		return r.err
	}

	var err error
	defer func() {
		nerr := r.res.Body.Close()
		if err == nil {
			err = nerr
		}
	}()

	if err = json.NewDecoder(r.res.Body).Decode(&data); err != nil {
		return errors.Wrap(ErrDecode, err.Error())
	}

	return r.err
}

func newResponse(r *http.Response, err error) *Response {
	resp := &Response{
		res: r,
		err: err,
	}

	if err != nil {
		return resp
	}

	s := resp.res.StatusCode
	if s == http.StatusOK || s == http.StatusCreated || s == http.StatusNoContent {
		return resp
	}

	if resp.res.StatusCode == http.StatusNotFound {
		resp.err = ErrNotFound
		return resp
	}

	log.Errorf("K8s: request failed with code %v", resp.res.StatusCode)

	b, rerr := io.ReadAll(resp.res.Body)
	if rerr == nil {
		log.Errorf("K8s: request failed with body: %s", string(b))
	}

	resp.err = ErrOther

	return resp
}

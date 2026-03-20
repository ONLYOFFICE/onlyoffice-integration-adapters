package watch

import (
	"bufio"
	"context"
	"encoding/json"
	"net/http"
	"sync/atomic"
	"time"

	"github.com/pkg/errors"
)

type bodyWatcher struct {
	ctx     context.Context
	stop    context.CancelFunc
	results chan Event
	res     *http.Response
	req     *http.Request
	client  *http.Client
}

func (wr *bodyWatcher) ResultChan() <-chan Event {
	return wr.results
}

func (wr *bodyWatcher) Stop() {
	select {
	case <-wr.ctx.Done():
		return
	default:
		wr.stop()
	}
}

func (wr *bodyWatcher) reconnect() error {
	req := wr.req.Clone(wr.ctx)

	//nolint:bodyclose
	res, err := wr.client.Do(req)
	if err != nil {
		return err
	}

	wr.res.Body.Close()
	wr.res = res

	return nil
}

func (wr *bodyWatcher) stream() {
	var ignore atomic.Bool

	ignore.Store(true)

	go func() {
		<-time.After(time.Second)
		ignore.Store(false)
	}()

	go func() {
		defer func() { wr.res.Body.Close() }()
	out:
		for {
			reader := bufio.NewReader(wr.res.Body)

			for {
				b, err := reader.ReadBytes('\n')
				if err != nil {
					break
				}

				if ignore.Load() {
					continue
				}

				var event Event
				if err := json.Unmarshal(b, &event); err != nil {
					continue
				}

				select {
				case <-wr.ctx.Done():
					break out
				case wr.results <- event:
				}
			}

			select {
			case <-wr.ctx.Done():
				break out
			case <-time.After(2 * time.Second):
			}

			if err := wr.reconnect(); err != nil {
				break out
			}
		}

		close(wr.results)
		wr.Stop()
	}()
}

func NewBodyWatcher(req *http.Request, client *http.Client) (Watch, error) {
	ctx, cancel := context.WithCancel(context.Background())

	req = req.WithContext(ctx)

	//nolint:bodyclose
	res, err := client.Do(req)
	if err != nil {
		cancel()
		return nil, errors.Wrap(err, "body watcher failed to make http request")
	}

	wr := &bodyWatcher{
		ctx:     ctx,
		results: make(chan Event),
		stop:    cancel,
		req:     req,
		res:     res,
		client:  client,
	}

	go wr.stream()

	return wr, nil
}

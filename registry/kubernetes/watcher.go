/**
 *
 * (c) Copyright Ascensio System SIA 2024
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

package kubernetes

import (
	"encoding/json"
	"errors"
	"strings"
	"sync"

	"go-micro.dev/v4/logger"
	"go-micro.dev/v4/registry"

	"github.com/ONLYOFFICE/onlyoffice-integration-adapters/registry/kubernetes/client"
	"github.com/ONLYOFFICE/onlyoffice-integration-adapters/registry/kubernetes/client/watch"
)

var (
	deleteAction = "delete"
)

type k8sWatcher struct {
	registry *kregistry
	watcher  watch.Watch
	next     chan *registry.Result

	sync.RWMutex
	pods map[string]*client.Pod
	sync.Once
}

func (k *k8sWatcher) updateCache() ([]*registry.Result, error) {
	podList, err := k.registry.client.ListPods(podSelector)
	if err != nil {
		return nil, err
	}

	var results []*registry.Result

	for _, p := range podList.Items {
		pod := p
		rslts := k.buildPodResults(&pod, nil)
		results = append(results, rslts...)

		k.Lock()
		k.pods[pod.Metadata.Name] = &pod
		k.Unlock()
	}

	return results, nil
}

func (k *k8sWatcher) buildPodResults(pod *client.Pod, cache *client.Pod) []*registry.Result {
	var results []*registry.Result

	ignore := make(map[string]bool)

	if pod.Metadata != nil {
		results, ignore = podBuildResult(pod, cache)
	}

	if cache != nil && cache.Metadata != nil {
		for annKey, annVal := range cache.Metadata.Annotations {
			if ignore[annKey] {
				continue
			}

			if !strings.HasPrefix(annKey, annotationServiceKeyPrefix) {
				continue
			}

			rslt := &registry.Result{Action: deleteAction}

			if err := json.Unmarshal([]byte(*annVal), &rslt.Service); err != nil {
				continue
			}

			results = append(results, rslt)
		}
	}

	return results
}

func (k *k8sWatcher) handleEvent(event watch.Event) {
	var pod client.Pod
	if err := json.Unmarshal([]byte(event.Object), &pod); err != nil {
		logger.Error("K8s Watcher: Couldnt unmarshal event object from pod")
		return
	}

	//nolint:exhaustive
	switch event.Type {
	case watch.Modified:
		k.RLock()
		cache := k.pods[pod.Metadata.Name]
		k.RUnlock()

		var results []*registry.Result

		if pod.Status.Phase == podRunning {
			results = k.buildPodResults(&pod, cache)
		} else {
			results = k.buildPodResults(&pod, nil)
		}

		for _, result := range results {
			if pod.Status.Phase != podRunning || pod.Metadata.DeletionTimestamp != "" {
				result.Action = deleteAction
			}
			k.next <- result
		}

		k.Lock()
		k.pods[pod.Metadata.Name] = &pod
		k.Unlock()

		return

	case watch.Deleted:
		results := k.buildPodResults(&pod, nil)

		for _, result := range results {
			result.Action = deleteAction
			k.next <- result
		}

		k.Lock()
		delete(k.pods, pod.Metadata.Name)
		k.Unlock()

		return
	}
}

func (k *k8sWatcher) Next() (*registry.Result, error) {
	r, ok := <-k.next
	if !ok {
		return nil, errors.New("result chan closed")
	}

	return r, nil
}

func (k *k8sWatcher) Stop() {
	k.watcher.Stop()

	select {
	case <-k.next:
		return
	default:
		k.Do(func() {
			close(k.next)
		})
	}
}

func newWatcher(kr *kregistry, opts ...registry.WatchOption) (registry.Watcher, error) {
	var wo registry.WatchOptions
	for _, o := range opts {
		o(&wo)
	}

	selector := podSelector
	if len(wo.Service) > 0 {
		selector = map[string]string{
			svcSelectorPrefix + serviceName(wo.Service): svcSelectorValue,
		}
	}

	watcher, err := kr.client.WatchPods(selector)
	if err != nil {
		return nil, err
	}

	k := &k8sWatcher{
		registry: kr,
		watcher:  watcher,
		next:     make(chan *registry.Result),
		pods:     make(map[string]*client.Pod),
	}

	if _, err := k.updateCache(); err != nil {
		return nil, err
	}

	go func() {
		for event := range watcher.ResultChan() {
			k.handleEvent(event)
		}

		k.Stop()
	}()

	return k, nil
}

func podBuildResult(pod *client.Pod, cache *client.Pod) ([]*registry.Result, map[string]bool) {
	results := make([]*registry.Result, 0, len(pod.Metadata.Annotations))
	ignore := make(map[string]bool)

	for annKey, annVal := range pod.Metadata.Annotations {
		if !strings.HasPrefix(annKey, annotationServiceKeyPrefix) {
			continue
		}

		if annVal == nil {
			continue
		}

		ignore[annKey] = true

		var (
			cacheExists bool
			cav         *string
		)

		if cache != nil && cache.Metadata != nil {
			cav, cacheExists = cache.Metadata.Annotations[annKey]
			if cacheExists && cav != nil && cav == annVal {
				continue
			}
		}

		rslt := &registry.Result{}
		if cacheExists {
			rslt.Action = "update"
		} else {
			rslt.Action = "create"
		}

		if err := json.Unmarshal([]byte(*annVal), &rslt.Service); err != nil {
			continue
		}

		results = append(results, rslt)
	}

	return results, ignore
}

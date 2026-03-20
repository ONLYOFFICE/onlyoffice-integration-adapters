package client

import "github.com/ONLYOFFICE/onlyoffice-integration-adapters/registry/kubernetes/client/watch"

type Kubernetes interface {
	ListPods(labels map[string]string) (*PodList, error)
	UpdatePod(podName string, pod *Pod) (*Pod, error)
	WatchPods(labels map[string]string) (watch.Watch, error)
}

type PodList struct {
	Items []Pod `json:"items"`
}

type Pod struct {
	Metadata *Meta   `json:"metadata"`
	Status   *Status `json:"status"`
}

type Meta struct {
	Name              string             `json:"name,omitempty"`
	Labels            map[string]*string `json:"labels,omitempty"`
	Annotations       map[string]*string `json:"annotations,omitempty"`
	DeletionTimestamp string             `json:"deletionTimestamp,omitempty"`
}

type Status struct {
	PodIP string `json:"podIP"`
	Phase string `json:"phase"`
}

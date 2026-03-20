package watch

import "encoding/json"

type Watch interface {
	Stop()
	ResultChan() <-chan Event
}

type EventType string

const (
	Added    EventType = "ADDED"
	Modified EventType = "MODIFIED"
	Deleted  EventType = "DELETED"
	Error    EventType = "ERROR"
)

type Event struct {
	Type   EventType       `json:"type"`
	Object json.RawMessage `json:"object"`
}

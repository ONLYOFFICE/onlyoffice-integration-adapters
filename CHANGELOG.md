# Change Log

### Changed
- migrated to go-micro v5 (`go-micro.dev/v5` v5.30.0)
- nats/rabbitmq brokers, consul/etcd registries and the opentelemetry wrapper now come from go-micro v5 core
- memory broker, mdns registry and memory store now use go-micro v5 built-in implementations
- hystrix, uber ratelimiter and roundrobin wrappers moved to `github.com/micro/plugins/v5`
- http services are built on `go-micro.dev/v5/web`: `service/http.NewService` returns `web.Service` instead of `micro.Service`
- bootstrapper runs either an rpc `micro.Service` or an http `web.Service`
- opentelemetry call wrapper is applied to the shared go-micro client in http services
- bumped Go to 1.26.7 and dependencies

### Fixed
- roundrobin and hystrix wrappers are applied to the injected `client.Client`

### Removed
- `github.com/go-micro/plugins/v4` dependencies, including the http server plugin
- `cache.Cache` parameter of `service/http.NewService`

## 0.1.6
### Added
- standalone kubernetes registry (`registry/kubernetes`)

### Changed
- bumped Go to 1.25.7 and dependencies

## 0.1.52
Also tagged as 0.1.51.

### Fixed
- kubernetes go-micro registry: replaced `go-micro/plugins/v4/registry/kubernetes` with a standalone implementation

## 0.1.5
### Changed
- bumped Go to 1.25.3 and dependencies

## 0.1.4
### Changed
- bumped Go to 1.23 and dependencies

## 0.1.3
### Changed
- bumped Go to 1.21 and dependencies

## 0.1.2 
### Fixed
- server version is a string (`SERVER_VERSION`)

## 0.1.1
### Added
- broker enable flag (`BROKER_ENABLE`)

### Fixed
- in-memory broker is used when the broker is disabled
- default broker configuration is no longer validated

## 0.1.0
### Added
- caching wrappers
- crypto module
- event emitters module
- go-micro messaging
- resilience patterns and middlewares
- distributed tracing
- storage adapters
- async workers module
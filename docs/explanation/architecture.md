---
dep:
  type: explanation
  audience: [contributor, maintainer]
  owner: "@checkfirst-ltd"
  created: 2026-04-03T12:00:00+03:00
  last_verified: 2026-04-03T12:00:00+03:00
  confidence: medium
  depends_on:
    - ../../src/microsoft-outlook.module.ts
    - ../../src/index.ts
  tags: [architecture, design, overview]
  links:
    - target: ../reference/module-configuration.md
      rel: EXPLAINS
    - target: ./oauth-flow.md
      rel: NEXT
    - target: ./delta-sync-and-rate-limiting.md
      rel: NEXT
---

# Architecture & Module Design

## Context

`@checkfirst/nestjs-outlook` is a NestJS library module that wraps the Microsoft Graph API for calendar and email operations. It is consumed by host NestJS applications (primarily `scheduleai-backend`) and designed as a self-contained, configurable module.

## Module Structure

The module follows NestJS's `ConfigurableModuleBuilder` pattern, providing `forRoot()` and `forRootAsync()` registration methods. This gives consumers full control over how configuration is injected.

### Layer Organization

```
src/
├── controllers/          # HTTP endpoints (auth callback, webhooks)
├── services/
│   ├── auth/             # OAuth flow, token management
│   ├── calendar/         # Calendar CRUD, webhooks, recurrence
│   ├── email/            # Email send, webhooks
│   ├── subscription/     # Subscription lifecycle (create, renew, cleanup)
│   └── shared/           # Cross-cutting: delta sync, rate limiter, user ID converter
├── entities/             # TypeORM entities (4 tables)
├── repositories/         # Custom TypeORM repositories
├── interfaces/           # TypeScript interfaces and config types
├── enums/                # Event types, permissions, resource types
├── dto/                  # Data transfer objects (webhook payloads)
├── types/                # Re-exported Microsoft Graph types
├── utils/                # Retry logic, Graph API executor
└── migrations/           # TypeORM database migrations
```

### Service Responsibilities

The module is organized around three domain areas with shared infrastructure:

**Auth** (`MicrosoftAuthService`): Owns the OAuth 2.0 authorization code flow. Manages token storage, refresh, and expiration detection. Emits `USER_AUTHENTICATED` when a user completes the consent flow.

**Calendar** (`CalendarService`, `RecurrenceService`): Calendar CRUD via Graph API, including batch operations (up to 20 per request). Webhook subscription creation and processing. Delta sync integration for efficient change detection. `RecurrenceService` handles series master/occurrence/exception classification and expansion.

**Email** (`EmailService`): Email sending via `/me/sendMail`. Email webhook subscription and notification processing.

**Shared** (`DeltaSyncService`, `GraphRateLimiterService`, `UserIdConverterService`): Cross-cutting concerns used by both calendar and email services. Delta sync provides incremental change tracking via Graph delta links. The rate limiter enforces Microsoft's per-user throttling limits (4 req/sec, 10K req/10min). `UserIdConverterService` maps between the host app's external user IDs and internal database IDs.

### Event-Driven Communication

The module communicates with the host application exclusively through NestJS `EventEmitter2`. This means:

- The host app has zero coupling to the module's internal services for receiving notifications
- All webhook processing results in typed events (`OutlookEventTypes` enum)
- Lifecycle events (reauthorization, subscription removal) let the host app react to subscription health changes

### User Identity Model

A key design decision is the two-layer user identity:

- **externalUserId** (string): The host application's user ID. This is how consumers reference users in all public API methods.
- **internalUserId** (number): The auto-generated primary key in the `microsoft_user` table. Used for internal database relationships only.

`UserIdConverterService` handles the mapping, keeping the host app decoupled from internal database structure.

## Tradeoffs

**Self-contained module vs. service library**: The module registers its own controllers (auth callback, webhook endpoints). This reduces integration effort but means consumers can't easily customize the webhook endpoint behavior. The `calendarWebhookPath` config option provides limited customization.

**TypeORM dependency**: The module requires TypeORM with MySQL. This was chosen for consistency with the primary consumer (`scheduleai-backend`) but limits adoption by applications using other ORMs or databases.

**Batch API usage**: Batch operations use Microsoft's $batch endpoint (max 20 requests per batch). This is an optimization for bulk operations but adds complexity in error handling — individual items in a batch can fail independently.

## Related

- [OAuth Flow & Token Management](./oauth-flow.md) — deep dive into authentication
- [Delta Sync & Rate Limiting](./delta-sync-and-rate-limiting.md) — change tracking and throttling
- [Module Configuration Reference](../reference/module-configuration.md) — all config options

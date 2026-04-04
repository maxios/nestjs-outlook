---
dep:
  type: reference
  audience: [ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-04-03T12:00:00+03:00
  last_verified: 2026-04-03T12:00:00+03:00
  confidence: medium
  depends_on:
    - ../../src/index.ts
    - ../../src/microsoft-outlook.module.ts
  tags: [ai, codebase, map, navigation]
  links:
    - target: ./ai-common-tasks.md
      rel: NEXT
    - target: ./module-configuration.md
      rel: NEXT
    - target: ../how-to/ai-add-graph-api-feature.md
      rel: NEXT
---

# Codebase Map

Quick-reference map of the `@checkfirst/nestjs-outlook` module for AI agents. Use this to orient before making changes.

---

## Identity

| Property | Value |
|----------|-------|
| Package | `@checkfirst/nestjs-outlook` |
| Framework | NestJS (TypeScript) |
| Database | MySQL via TypeORM |
| External API | Microsoft Graph API v1.0 |
| Publish target | npm (compiled to `dist/`) |
| Entry point | `src/index.ts` (re-exports everything public) |

---

## Directory Structure

| Path | Purpose | When to modify |
|------|---------|----------------|
| `src/microsoft-outlook.module.ts` | Root module definition, providers, exports | Adding/removing a service or controller |
| `src/index.ts` | Public API barrel file | Exposing new types, services, or controllers |
| `src/services/auth/` | OAuth flow, token management | Changing auth behavior or adding new auth methods |
| `src/services/calendar/` | Calendar CRUD, webhooks, recurrence | Adding calendar features or modifying event processing |
| `src/services/email/` | Email send, email webhooks | Adding email features |
| `src/services/subscription/` | Subscription lifecycle (create, renew, cleanup) | Changing subscription management logic |
| `src/services/shared/` | Delta sync, rate limiter, user ID converter | Cross-cutting concerns used by calendar and email |
| `src/controllers/` | HTTP endpoints (auth callback, calendar webhook, email webhook) | Changing request handling or adding endpoints |
| `src/entities/` | TypeORM entities (4 tables) | Schema changes (requires migration) |
| `src/repositories/` | Custom TypeORM repositories | Adding custom queries |
| `src/interfaces/` | Config, auth, recurrence interfaces | Adding new configuration or contract types |
| `src/enums/` | Event types, permissions, resource types, show-as | Adding new enum values |
| `src/dto/` | Webhook notification DTOs | Changing webhook payload validation |
| `src/types/` | Re-exported Microsoft Graph types + custom batch types | Adding new Graph type re-exports |
| `src/utils/` | Retry logic, Graph API executor, webhook validator | Changing retry/execution behavior |
| `src/migrations/` | TypeORM migrations | Schema changes |

---

## Service Dependency Graph

```
MicrosoftAuthService
  ├── CalendarService (forwardRef)
  ├── EmailService (forwardRef)
  ├── MicrosoftCsrfTokenRepository
  └── MicrosoftUser (TypeORM Repository)

CalendarService
  ├── MicrosoftAuthService (forwardRef)
  ├── OutlookWebhookSubscriptionRepository
  ├── OutlookDeltaLinkRepository
  ├── DeltaSyncService
  ├── UserIdConverterService
  ├── MicrosoftSubscriptionService
  ├── GraphRateLimiterService
  └── MicrosoftUser (TypeORM Repository)

EmailService
  ├── MicrosoftAuthService (forwardRef)
  ├── OutlookWebhookSubscriptionRepository
  ├── UserIdConverterService
  └── MicrosoftUser (TypeORM Repository)

MicrosoftSubscriptionService (no dependencies — standalone)

DeltaSyncService
  ├── OutlookDeltaLinkRepository
  ├── UserIdConverterService
  └── GraphRateLimiterService
```

---

## Key Patterns

| Pattern | Description |
|---------|-------------|
| `forwardRef(() => Service)` | Used between `MicrosoftAuthService` ↔ `CalendarService` ↔ `EmailService` to break circular deps |
| `externalUserId: string` | All public methods use the host app's user ID, never internal DB IDs |
| `EventEmitter2` | All notifications to the host app go through event emission, not return values |
| `executeGraphApiCall()` | Wrapper for Microsoft Graph API calls with error handling (in `src/utils/`) |
| `retryWithBackoff()` | Retry utility for transient failures (in `src/utils/retry.util.ts`) |
| `GraphRateLimiterService` | Must be consulted before Graph API calls (4 req/sec, 10K req/10min per user) |
| Batch API | Operations on multiple events use `$batch` endpoint (max 20 per request) |
| `@Cron` decorators | Automatic subscription renewal (hourly) and CSRF cleanup (every 5 min) |

---

## Entities & Tables

| Entity | Table | Key Columns |
|--------|-------|-------------|
| `MicrosoftUser` | `microsoft_user` | `externalUserId`, `accessToken`, `refreshToken`, `tokenExpiry`, `defaultCalendarId`, `isActive` |
| `OutlookWebhookSubscription` | `outlook_webhook_subscription` | `subscriptionId`, `externalUserId`, `resource`, `expirationDateTime` |
| `MicrosoftCsrfToken` | `microsoft_csrf_token` | `token`, `externalUserId`, `expiresAt` |
| `OutlookDeltaLink` | `outlook_delta_link` | `externalUserId`, `resourceType`, `deltaLink` |

---

## Enums Quick Reference

| Enum | Location | Values |
|------|----------|--------|
| `OutlookEventTypes` | `src/enums/event-types.enum.ts` | `USER_AUTHENTICATED`, `EVENT_CREATED`, `EVENT_UPDATED`, `EVENT_DELETED`, `EVENT_NOTIFICATION`, `IMPORT_COMPLETED`, `EMAIL_RECEIVED`, `EMAIL_UPDATED`, `EMAIL_DELETED`, `LIFECYCLE_*` |
| `PermissionScope` | `src/enums/permission-scope.enum.ts` | `CALENDAR_READ`, `CALENDAR_WRITE`, `EMAIL_READ`, `EMAIL_WRITE`, `EMAIL_SEND` |
| `ShowAsType` | `src/enums/show-as-type.enum.ts` | `UNKNOWN`, `FREE`, `TENTATIVE`, `BUSY`, `OOF`, `WORKING_ELSEWHERE` |
| `ResourceType` | `src/enums/resource-type.enum.ts` | Check file — used for delta link resource classification |

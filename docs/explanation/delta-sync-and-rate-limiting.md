---
dep:
  type: explanation
  audience: [contributor, maintainer]
  owner: "@checkfirst-ltd"
  created: 2026-04-03T12:00:00+03:00
  last_verified: 2026-04-03T12:00:00+03:00
  confidence: medium
  depends_on:
    - ../../src/services/shared/delta-sync.service.ts
    - ../../src/services/shared/graph-rate-limiter.service.ts
    - ../../src/entities/delta-link.entity.ts
  tags: [delta-sync, rate-limiting, graph-api, performance]
  links:
    - target: ../reference/calendar-service.md
      rel: EXPLAINS
    - target: ./architecture.md
      rel: REQUIRES
---

# Delta Sync & Rate Limiting

## Context

When processing webhook notifications, the module needs to fetch the actual event data from Microsoft Graph. Two shared services optimize this: `DeltaSyncService` for efficient change tracking and `GraphRateLimiterService` for staying within Microsoft's API limits.

## Delta Sync

### Problem

Microsoft webhook notifications only contain a resource URL and change type — not the full event data. The naive approach is to fetch each changed event individually, but this doesn't scale when users have many simultaneous changes.

### Solution: Delta Queries

Microsoft Graph supports delta queries that return only changes since the last sync. The flow:

1. **Initialization**: `CalendarService.initializeDeltaSync()` makes an initial delta query to `/me/calendarView/delta`, receiving all current events and a `deltaLink`.
2. **Storage**: The delta link is stored in the `outlook_delta_link` table, keyed by user ID and resource type.
3. **Subsequent syncs**: When a webhook fires, `handleOutlookWebhookV2()` uses the stored delta link to fetch only changes since the last sync.
4. **Pagination**: Delta responses may be paginated with `@odata.nextLink`. The service follows all pages before storing the new `@odata.deltaLink`.
5. **Error recovery**: If a delta link becomes invalid (Microsoft returns 410 Gone), the service falls back to a full re-sync.

### Delta Item Classification

Each item in a delta response can be:
- A new or updated item (has full properties)
- A deleted item (has `@removed` property with reason `"changed"` or `"deleted"`)

The `DeltaSyncService` provides typed responses (`DeltaEvent`, `DeltaMessage`) that include both the item data and removal metadata.

## Rate Limiting

### Microsoft's Limits

Microsoft Graph enforces per-user (mailbox) rate limits:
- **4 requests per second** per user
- **10,000 requests per 10 minutes** per user

Exceeding these returns HTTP 429 with a `Retry-After` header.

### GraphRateLimiterService

The rate limiter uses a sliding window algorithm with per-user tracking:

- **Short window**: Tracks timestamps of recent requests within a 1-second window (max 4)
- **Long window**: Tracks timestamps within a 10-minute window (max 10,000)
- **Cooldown**: When a 429 response is received, the `Retry-After` header sets a cooldown period during which no requests are sent for that user
- **Automatic cleanup**: A cron job prunes inactive user limiters (no activity for 30 minutes) to prevent memory leaks

Services call the rate limiter before making Graph API requests. If the limit would be exceeded, the request is delayed until a slot is available.

## Tradeoffs

**Delta sync complexity vs. efficiency**: Delta sync significantly reduces API calls but introduces state management complexity (storing delta links, handling invalidation, pagination). It's used in the V2 webhook handler while V1 fetches events individually — both are available for consumers to choose based on their needs.

**In-memory rate limiter**: The rate limiter stores state in memory, not in a database. This means each application instance tracks its own limits independently. In multi-instance deployments, the effective limit per user is multiplied by the number of instances. This was chosen for performance (no database overhead per API call) but may need revisiting for large-scale deployments.

**Retry strategy**: The `DeltaSyncService` uses a simple retry with fixed delays (max 3 retries, 1-second delay). Combined with the rate limiter's cooldown support, this handles most transient failures. More sophisticated backoff strategies could be added if needed.

## Related

- [Architecture](./architecture.md) — where these services fit in the module structure
- [Calendar Service Reference](../reference/calendar-service.md) — public methods that use delta sync

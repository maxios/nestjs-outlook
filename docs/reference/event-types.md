---
dep:
  type: reference
  audience: [consumer-dev, contributor]
  owner: "@checkfirst-ltd"
  created: 2026-04-03T12:00:00+03:00
  last_verified: 2026-04-03T12:00:00+03:00
  confidence: medium
  depends_on:
    - ../../src/enums/event-types.enum.ts
  tags: [events, enum, event-emitter]
  links:
    - target: ../how-to/handle-calendar-webhooks.md
      rel: NEXT
---

# Event Types Reference

All event types emitted by the module via NestJS `EventEmitter2`. Subscribe using `@OnEvent()` decorator.

Source: `src/enums/event-types.enum.ts`

---

## Authentication Events

| Enum Value | String | Description |
|------------|--------|-------------|
| `USER_AUTHENTICATED` | `microsoft.auth.user.authenticated` | Emitted after successful OAuth callback and token storage |

---

## Calendar Events

| Enum Value | String | Description |
|------------|--------|-------------|
| `EVENT_CREATED` | `outlook.event.created` | A new calendar event was created |
| `EVENT_UPDATED` | `outlook.event.updated` | An existing calendar event was modified |
| `EVENT_DELETED` | `outlook.event.deleted` | A calendar event was deleted |
| `EVENT_NOTIFICATION` | `outlook.event.notification` | Raw webhook notification received (before processing) |
| `IMPORT_COMPLETED` | `outlook.calendar.import.completed` | Bulk calendar import finished |

---

## Email Events

| Enum Value | String | Description |
|------------|--------|-------------|
| `EMAIL_RECEIVED` | `outlook.email.received` | A new email was received |
| `EMAIL_UPDATED` | `outlook.email.updated` | An existing email was modified |
| `EMAIL_DELETED` | `outlook.email.deleted` | An email was deleted |

---

## Lifecycle Events

| Enum Value | String | Description |
|------------|--------|-------------|
| `LIFECYCLE_REAUTHORIZATION_REQUIRED` | `outlook.lifecycle.reauthorization_required` | Microsoft requires the user to re-authorize the app |
| `LIFECYCLE_SUBSCRIPTION_REMOVED` | `outlook.lifecycle.subscription_removed` | A webhook subscription was removed by Microsoft |
| `LIFECYCLE_MISSED` | `outlook.lifecycle.missed` | A lifecycle notification was missed |

---

## Usage

```typescript
import { OnEvent } from '@nestjs/event-emitter';
import { OutlookEventTypes } from '@checkfirst/nestjs-outlook';

@OnEvent(OutlookEventTypes.EVENT_CREATED)
handleCreated(payload: any) { /* ... */ }
```

---
dep:
  type: reference
  audience: [consumer-dev, contributor]
  owner: "@checkfirst-ltd"
  created: 2026-04-03T12:00:00+03:00
  last_verified: 2026-04-03T12:00:00+03:00
  confidence: medium
  depends_on:
    - ../../src/services/calendar/calendar.service.ts
    - ../../src/types/microsoft-graph.types.ts
  tags: [calendar, service, api]
  links:
    - target: ./module-configuration.md
      rel: REQUIRES
    - target: ./event-types.md
      rel: USES
    - target: ./recurrence-types.md
      rel: NEXT
    - target: ../how-to/handle-calendar-webhooks.md
      rel: NEXT
---

# Calendar Service Reference

Public API of `CalendarService` for managing Outlook calendar events via Microsoft Graph API.

Source: `src/services/calendar/calendar.service.ts`

---

## `getDefaultCalendarId(externalUserId)`

| Property | Value |
|----------|-------|
| Parameters | `externalUserId: string` |
| Returns | `Promise<string>` |
| Description | Retrieves the user's default calendar ID. Caches the result in the `microsoft_user` table. |

---

## `createEvent(event, externalUserId, calendarId)`

| Property | Value |
|----------|-------|
| Parameters | `event: Partial<Event>`, `externalUserId: string`, `calendarId: string` |
| Returns | `Promise<{ event: Event }>` |
| Description | Creates a calendar event in the specified calendar. |

---

## `updateEvent(eventId, updates, externalUserId, calendarId)`

| Property | Value |
|----------|-------|
| Parameters | `eventId: string`, `updates: Partial<Event>`, `externalUserId: string`, `calendarId: string` |
| Returns | `Promise<{ event: Event }>` |
| Description | Updates an existing calendar event. |

---

## `getEventById(externalUserId, eventId)`

| Property | Value |
|----------|-------|
| Parameters | `externalUserId: string`, `eventId: string` |
| Returns | `Promise<Event \| null>` |
| Description | Fetches a single calendar event by its Microsoft Graph event ID. |

---

## `deleteEvent(event, externalUserId, calendarId)`

| Property | Value |
|----------|-------|
| Parameters | `event: Partial<Event>`, `externalUserId: string`, `calendarId: string` |
| Returns | `Promise<void>` |
| Description | Deletes a calendar event. |

---

## `createBatchEvents(events, externalUserId, calendarId)`

| Property | Value |
|----------|-------|
| Parameters | `events: Partial<Event>[]`, `externalUserId: string`, `calendarId: string` |
| Returns | `Promise<{ index: number; success: boolean; event?: Event; error?: string }[]>` |
| Description | Creates multiple events in a single batch request. Uses Microsoft Graph batch API (max 20 per batch). |

---

## `updateBatchEvents(events, externalUserId, calendarId)`

| Property | Value |
|----------|-------|
| Parameters | `events: { eventId: string; updates: Partial<Event> }[]`, `externalUserId: string`, `calendarId: string` |
| Returns | `Promise<{ index: number; success: boolean; event?: Event; error?: string }[]>` |
| Description | Updates multiple events in a single batch request. |

---

## `deleteBatchEvents(eventIds, externalUserId, calendarId)`

| Property | Value |
|----------|-------|
| Parameters | `eventIds: string[]`, `externalUserId: string`, `calendarId: string` |
| Returns | `Promise<{ index: number; success: boolean; error?: string }[]>` |
| Description | Deletes multiple events in a single batch request. |

---

## `createWebhookSubscription(externalUserId)`

| Property | Value |
|----------|-------|
| Parameters | `externalUserId: string` |
| Returns | `Promise<void>` |
| Description | Creates a Microsoft Graph webhook subscription for calendar change notifications. Stores the subscription in the database. |

---

## `renewWebhookSubscription(subscriptionId)`

| Property | Value |
|----------|-------|
| Parameters | `subscriptionId: string` |
| Returns | `Promise<void>` |
| Description | Renews an existing webhook subscription before it expires. |

---

## `deleteWebhookSubscription(externalUserId)`

| Property | Value |
|----------|-------|
| Parameters | `externalUserId: string` |
| Returns | `Promise<void>` |
| Description | Deletes the active webhook subscription for a user. |

---

## `handleOutlookWebhook(validationToken, notifications)`

| Property | Value |
|----------|-------|
| Parameters | `validationToken: string \| undefined`, `notifications: ChangeNotification[]` |
| Returns | `Promise<string \| void>` |
| Description | Processes incoming webhook notifications. Returns the validation token for subscription validation requests. Emits `EVENT_CREATED`, `EVENT_UPDATED`, or `EVENT_DELETED` events. |

---

## `handleOutlookWebhookV2(validationToken, notifications)`

| Property | Value |
|----------|-------|
| Parameters | `validationToken: string \| undefined`, `notifications: ChangeNotification[]` |
| Returns | `Promise<string \| void>` |
| Description | V2 webhook handler using delta sync for change detection instead of individual event fetching. |

---

## `initializeDeltaSync(externalUserId)`

| Property | Value |
|----------|-------|
| Parameters | `externalUserId: string` |
| Returns | `Promise<void>` |
| Description | Initializes delta sync for a user by fetching the initial delta link. |

---

## `getEventsBatch(externalUserId, eventIds)`

| Property | Value |
|----------|-------|
| Parameters | `externalUserId: string`, `eventIds: string[]` |
| Returns | `Promise<Event[]>` |
| Description | Fetches multiple events by ID using the batch API. |

---

## Cron Jobs

| Method | Schedule | Description |
|--------|----------|-------------|
| `renewSubscriptions()` | Every hour | Automatically renews webhook subscriptions nearing expiration |

---
dep:
  type: how-to
  audience: [consumer-dev]
  owner: "@checkfirst-ltd"
  created: 2026-04-03T12:00:00+03:00
  last_verified: 2026-04-03T12:00:00+03:00
  confidence: medium
  depends_on:
    - ../../src/services/calendar/calendar.service.ts
    - ../../src/controllers/calendar.controller.ts
    - ../../src/enums/event-types.enum.ts
  tags: [calendar, webhooks, events]
  links:
    - target: ../tutorials/getting-started.md
      rel: REQUIRES
    - target: ../reference/calendar-service.md
      rel: USES
    - target: ../reference/event-types.md
      rel: USES
    - target: ./manage-subscriptions.md
      rel: NEXT
---

# How-To: Handle Calendar Webhooks

**Goal**: Receive and process real-time calendar change notifications from Microsoft Outlook.

## Prerequisites

- Module installed and configured (see [Getting Started](../tutorials/getting-started.md))
- User has completed OAuth flow and has active tokens

## Steps

1. Create a webhook subscription for the user:

   ```typescript
   await calendarService.createWebhookSubscription('user-123');
   ```

2. The module's built-in `CalendarController` exposes the webhook endpoint at the configured `calendarWebhookPath` (default: `/calendar/webhook`). Microsoft sends POST requests to this endpoint when calendar events change.

3. Listen for processed events in your application:

   ```typescript
   import { OnEvent } from '@nestjs/event-emitter';
   import { OutlookEventTypes } from '@checkfirst/nestjs-outlook';

   @Injectable()
   export class CalendarListener {
     @OnEvent(OutlookEventTypes.EVENT_CREATED)
     onCreated(payload: { externalUserId: string; event: Event }) {
       // New event created in user's calendar
     }

     @OnEvent(OutlookEventTypes.EVENT_UPDATED)
     onUpdated(payload: { externalUserId: string; event: Event }) {
       // Existing event modified
     }

     @OnEvent(OutlookEventTypes.EVENT_DELETED)
     onDeleted(payload: { externalUserId: string; eventId: string }) {
       // Event removed from calendar
     }
   }
   ```

4. Handle lifecycle events for subscription health:

   ```typescript
   @OnEvent(OutlookEventTypes.LIFECYCLE_REAUTHORIZATION_REQUIRED)
   onReauth(payload: { externalUserId: string }) {
     // Prompt user to re-authenticate
   }

   @OnEvent(OutlookEventTypes.LIFECYCLE_SUBSCRIPTION_REMOVED)
   onRemoved(payload: { externalUserId: string }) {
     // Re-create subscription
   }
   ```

## Verification

- Create/update/delete a calendar event in the user's Outlook calendar
- Verify your `@OnEvent` handlers are called with the correct payload
- Check application logs for `CalendarService` entries confirming webhook processing

## Related

- [Calendar Service Reference](../reference/calendar-service.md) — full API details
- [Event Types Reference](../reference/event-types.md) — all emitted event types
- [Manage Subscriptions](./manage-subscriptions.md) — subscription lifecycle management

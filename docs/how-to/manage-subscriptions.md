---
dep:
  type: how-to
  audience: [consumer-dev]
  owner: "@checkfirst-ltd"
  created: 2026-04-03T12:00:00+03:00
  last_verified: 2026-04-03T12:00:00+03:00
  confidence: medium
  depends_on:
    - ../../src/services/subscription/microsoft-subscription.service.ts
    - ../../src/services/calendar/calendar.service.ts
  tags: [subscriptions, webhooks, cleanup]
  links:
    - target: ../tutorials/getting-started.md
      rel: REQUIRES
    - target: ../reference/calendar-service.md
      rel: USES
    - target: ../reference/email-service.md
      rel: USES
    - target: ./handle-calendar-webhooks.md
      rel: REQUIRES
---

# How-To: Manage Subscriptions

**Goal**: Create, renew, clean up, and monitor Microsoft Graph webhook subscriptions.

## Prerequisites

- Module installed and configured (see [Getting Started](../tutorials/getting-started.md))
- User has completed OAuth flow

## Steps

### Create Subscriptions

1. Create a calendar webhook subscription:

   ```typescript
   await calendarService.createWebhookSubscription('user-123');
   ```

2. Create an email webhook subscription:

   ```typescript
   await emailService.createWebhookSubscription('user-123');
   ```

### Automatic Renewal

The module automatically renews subscriptions via a cron job that runs every hour. No action required.

### Manual Cleanup

3. Clean up all subscriptions for a user using `MicrosoftSubscriptionService`:

   ```typescript
   import { MicrosoftSubscriptionService } from '@checkfirst/nestjs-outlook';

   const result = await subscriptionService.cleanupSubscriptions({
     accessToken: token,
   });
   // result: { totalFound, successfullyDeleted, failedToDelete, ... }
   ```

4. Clean up with a filter (e.g., only calendar subscriptions):

   ```typescript
   const result = await subscriptionService.cleanupSubscriptions({
     accessToken: token,
     filter: (sub) => sub.resource.includes('/events'),
   });
   ```

### Full Cleanup (Disconnect User)

5. Revoke tokens and remove all subscriptions:

   ```typescript
   await subscriptionService.fullCleanup(refreshToken, accessToken);
   ```

## Verification

- Call `subscriptionService.getActiveSubscriptions(accessToken)` to list current subscriptions
- Verify the subscription count matches expectations
- Check application logs for renewal and cleanup activity

## Related

- [Calendar Service Reference](../reference/calendar-service.md) — calendar subscription methods
- [Email Service Reference](../reference/email-service.md) — email subscription methods
- [Handle Calendar Webhooks](./handle-calendar-webhooks.md) — processing webhook notifications

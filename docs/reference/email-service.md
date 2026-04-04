---
dep:
  type: reference
  audience: [consumer-dev, contributor]
  owner: "@checkfirst-ltd"
  created: 2026-04-03T12:00:00+03:00
  last_verified: 2026-04-03T12:00:00+03:00
  confidence: medium
  depends_on:
    - ../../src/services/email/email.service.ts
  tags: [email, service, api]
  links:
    - target: ./module-configuration.md
      rel: REQUIRES
    - target: ./event-types.md
      rel: USES
    - target: ../how-to/send-emails.md
      rel: NEXT
---

# Email Service Reference

Public API of `EmailService` for sending and managing emails via Microsoft Graph API.

Source: `src/services/email/email.service.ts`

---

## `sendEmail(message, externalUserId)`

| Property | Value |
|----------|-------|
| Parameters | `message: Message`, `externalUserId: string` |
| Returns | `Promise<any>` |
| Description | Sends an email using the Microsoft Graph `/me/sendMail` endpoint. The `Message` type is from `@microsoft/microsoft-graph-types`. |

---

## `createWebhookSubscription(externalUserId)`

| Property | Value |
|----------|-------|
| Parameters | `externalUserId: string` |
| Returns | `Promise<void>` |
| Description | Creates a Microsoft Graph webhook subscription for email change notifications (new mail, updates, deletions). |

---

## `deleteWebhookSubscription(externalUserId)`

| Property | Value |
|----------|-------|
| Parameters | `externalUserId: string` |
| Returns | `Promise<void>` |
| Description | Deletes the active email webhook subscription for a user. |

---

## `handleEmailWebhook(validationToken, notifications)`

| Property | Value |
|----------|-------|
| Parameters | `validationToken: string \| undefined`, `notifications: ChangeNotification[]` |
| Returns | `Promise<string \| void>` |
| Description | Processes incoming email webhook notifications. Returns validation token for subscription validation. Emits `EMAIL_RECEIVED`, `EMAIL_UPDATED`, or `EMAIL_DELETED` events. |

---
dep:
  type: how-to
  audience: [consumer-dev]
  owner: "@checkfirst-ltd"
  created: 2026-04-03T12:00:00+03:00
  last_verified: 2026-04-03T12:00:00+03:00
  confidence: medium
  depends_on:
    - ../../src/services/email/email.service.ts
  tags: [email, send, graph-api]
  links:
    - target: ../tutorials/getting-started.md
      rel: REQUIRES
    - target: ../reference/email-service.md
      rel: USES
    - target: ../reference/permission-scopes.md
      rel: USES
---

# How-To: Send Emails

**Goal**: Send an email on behalf of an authenticated Microsoft user via Microsoft Graph API.

## Prerequisites

- Module installed and configured (see [Getting Started](../tutorials/getting-started.md))
- User has completed OAuth flow with `EMAIL_SEND` permission scope

## Steps

1. Inject `EmailService` into your service or controller:

   ```typescript
   import { EmailService } from '@checkfirst/nestjs-outlook';
   import { Message } from '@microsoft/microsoft-graph-types';

   @Injectable()
   export class NotificationService {
     constructor(private readonly emailService: EmailService) {}
   }
   ```

2. Compose a `Message` object following the Microsoft Graph `Message` type:

   ```typescript
   const message: Message = {
     subject: 'Meeting Confirmation',
     body: {
       contentType: 'html',
       content: '<p>Your meeting has been confirmed.</p>',
     },
     toRecipients: [
       {
         emailAddress: {
           address: 'recipient@example.com',
           name: 'Recipient Name',
         },
       },
     ],
   };
   ```

3. Send the email:

   ```typescript
   await this.emailService.sendEmail(message, 'user-123');
   ```

## Verification

- Check the user's Sent Items folder in Outlook for the sent email
- Verify the recipient received the email
- Check application logs for `EmailService` entries confirming the send

## Related

- [Email Service Reference](../reference/email-service.md) — full API details
- [Permission Scopes Reference](../reference/permission-scopes.md) — required scopes

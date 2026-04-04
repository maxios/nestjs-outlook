---
dep:
  type: tutorial
  audience: [consumer-dev]
  owner: "@checkfirst-ltd"
  created: 2026-04-03T12:00:00+03:00
  last_verified: 2026-04-03T12:00:00+03:00
  confidence: medium
  depends_on:
    - ../../src/microsoft-outlook.module.ts
    - ../../src/interfaces/config/outlook-config.interface.ts
  tags: [getting-started, setup, installation]
  links:
    - target: ../reference/module-configuration.md
      rel: TEACHES
    - target: ../reference/event-types.md
      rel: TEACHES
    - target: ../how-to/handle-calendar-webhooks.md
      rel: NEXT
    - target: ../how-to/send-emails.md
      rel: NEXT
---

# Tutorial: Set Up nestjs-outlook in Your NestJS Application

## Prerequisites

- A NestJS application with TypeORM configured (MySQL)
- A Microsoft Azure AD app registration with client ID and secret
- Node.js 18+

## What You'll Build

A NestJS application that authenticates users with Microsoft, receives calendar webhook notifications, and can send emails via Microsoft Graph API.

## Steps

### Step 1 — Install the Package

```bash
npm install @checkfirst/nestjs-outlook
```

**Expected result**: The package is added to your `package.json` dependencies.

### Step 2 — Register the Module

Import `MicrosoftOutlookModule` in your root or feature module using `forRoot()` or `forRootAsync()`:

```typescript
import { MicrosoftOutlookModule } from '@checkfirst/nestjs-outlook';

@Module({
  imports: [
    MicrosoftOutlookModule.forRoot({
      clientId: process.env.MICROSOFT_CLIENT_ID,
      clientSecret: process.env.MICROSOFT_CLIENT_SECRET,
      redirectPath: 'auth/microsoft/callback',
      backendBaseUrl: process.env.BACKEND_BASE_URL,
      basePath: 'api/v1',
      calendarWebhookPath: '/calendar/webhook',
    }),
  ],
})
export class AppModule {}
```

For async configuration (recommended for production):

```typescript
MicrosoftOutlookModule.forRootAsync({
  imports: [ConfigModule],
  useFactory: (configService: ConfigService) => ({
    clientId: configService.get('MICROSOFT_CLIENT_ID'),
    clientSecret: configService.get('MICROSOFT_CLIENT_SECRET'),
    redirectPath: 'auth/microsoft/callback',
    backendBaseUrl: configService.get('BACKEND_BASE_URL'),
    basePath: 'api/v1',
  }),
  inject: [ConfigService],
}),
```

**Expected result**: The module registers its controllers, services, and TypeORM entities. No errors on application startup.

### Step 3 — Run Database Migrations

The module ships with TypeORM migrations that create the required tables (`outlook_webhook_subscription`, `microsoft_csrf_token`, `microsoft_user`, `outlook_delta_link`).

Add the module's migration path to your TypeORM data source configuration:

```typescript
migrations: [
  'dist/migrations/*.js',
  'node_modules/@checkfirst/nestjs-outlook/dist/migrations/*.js',
],
```

Then run migrations:

```bash
npm run typeorm:run
```

**Expected result**: Four tables are created in your database.

### Step 4 — Listen for Events

The module emits events via NestJS `EventEmitter2`. Subscribe to them in your own services:

```typescript
import { OnEvent } from '@nestjs/event-emitter';
import { OutlookEventTypes } from '@checkfirst/nestjs-outlook';

@Injectable()
export class CalendarSyncService {
  @OnEvent(OutlookEventTypes.EVENT_CREATED)
  handleEventCreated(payload: any) {
    // Handle new calendar event
  }

  @OnEvent(OutlookEventTypes.USER_AUTHENTICATED)
  handleUserAuthenticated(payload: any) {
    // Handle successful OAuth flow completion
  }
}
```

**Expected result**: Your service methods are called when the corresponding Microsoft events occur.

### Step 5 — Initiate OAuth Flow

Use `MicrosoftAuthService` to generate the OAuth URL and redirect users:

```typescript
import { MicrosoftAuthService } from '@checkfirst/nestjs-outlook';

@Controller('auth')
export class AuthController {
  constructor(private readonly msAuth: MicrosoftAuthService) {}

  @Get('microsoft/connect')
  async connect(@Res() res: Response) {
    const url = await this.msAuth.getAuthorizationUrl('user-123');
    res.redirect(url);
  }
}
```

**Expected result**: Users are redirected to Microsoft's consent screen. After granting permissions, they are redirected back to your `redirectPath`, and a `USER_AUTHENTICATED` event is emitted.

## What You Built

You have a NestJS application that:
- Authenticates users with Microsoft OAuth 2.0
- Stores tokens and manages refresh automatically
- Listens for calendar and email change events
- Is ready to use `CalendarService` and `EmailService` for Graph API operations

## Next Steps

- [Handle Calendar Webhooks](../how-to/handle-calendar-webhooks.md) — process real-time calendar notifications
- [Send Emails](../how-to/send-emails.md) — send emails via Graph API
- [Module Configuration Reference](../reference/module-configuration.md) — all configuration options

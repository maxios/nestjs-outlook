---
dep:
  type: reference
  audience: [consumer-dev, contributor]
  owner: "@checkfirst-ltd"
  created: 2026-04-03T12:00:00+03:00
  last_verified: 2026-04-03T12:00:00+03:00
  confidence: medium
  depends_on:
    - ../../src/interfaces/config/outlook-config.interface.ts
    - ../../src/microsoft-outlook.module.ts
    - ../../src/constants.ts
  tags: [configuration, module, setup]
  links:
    - target: ../tutorials/getting-started.md
      rel: REQUIRES
---

# Module Configuration Reference

Configuration options for `MicrosoftOutlookModule.forRoot()` and `MicrosoftOutlookModule.forRootAsync()`.

---

## `MicrosoftOutlookConfig`

Source: `src/interfaces/config/outlook-config.interface.ts`

| Field | Type | Required | Default | Description |
|-------|------|----------|---------|-------------|
| `clientId` | `string` | Yes | — | Microsoft Azure AD application (client) ID |
| `clientSecret` | `string` | Yes | — | Microsoft Azure AD application client secret |
| `redirectPath` | `string` | Yes | — | Path segment of the OAuth redirect URI (e.g., `auth/microsoft/callback`) |
| `backendBaseUrl` | `string` | Yes | — | Base URL of the backend server (e.g., `https://api.example.com`) |
| `basePath` | `string` | No | `undefined` | API base path prefix (e.g., `api/v1`). Prepended to controller routes |
| `calendarWebhookPath` | `string` | No | `/calendar/webhook` | Path for the calendar webhook endpoint |

---

## Registration Methods

### `forRoot(config)`

Synchronous registration. Pass a `MicrosoftOutlookConfig` object directly.

```typescript
MicrosoftOutlookModule.forRoot({
  clientId: '...',
  clientSecret: '...',
  redirectPath: 'auth/microsoft/callback',
  backendBaseUrl: 'https://api.example.com',
})
```

### `forRootAsync(options)`

Asynchronous registration. Supports `useFactory`, `useClass`, and `useExisting` patterns for dependency injection.

```typescript
MicrosoftOutlookModule.forRootAsync({
  imports: [ConfigModule],
  useFactory: (config: ConfigService) => ({
    clientId: config.get('MS_CLIENT_ID'),
    clientSecret: config.get('MS_CLIENT_SECRET'),
    redirectPath: 'auth/microsoft/callback',
    backendBaseUrl: config.get('BACKEND_URL'),
  }),
  inject: [ConfigService],
})
```

---

## Injection Token

| Token | Value | Description |
|-------|-------|-------------|
| `MICROSOFT_CONFIG` | `'MICROSOFT_CONFIG'` | Injection token for the resolved config. Used internally by services. |

---

## Registered Entities

The module registers these TypeORM entities via `TypeOrmModule.forFeature()`:

| Entity | Table | Description |
|--------|-------|-------------|
| `OutlookWebhookSubscription` | `outlook_webhook_subscription` | Tracks active webhook subscriptions |
| `MicrosoftCsrfToken` | `microsoft_csrf_token` | Stores CSRF tokens for OAuth state validation |
| `MicrosoftUser` | `microsoft_user` | Stores Microsoft user tokens and metadata |
| `OutlookDeltaLink` | `outlook_delta_link` | Stores delta sync links for incremental change tracking |

---

## Exported Services

| Service | Description |
|---------|-------------|
| `CalendarService` | Calendar CRUD, webhook handling, delta sync |
| `RecurrenceService` | Recurring event processing and expansion |
| `EmailService` | Send and manage emails |
| `MicrosoftAuthService` | OAuth flow, token management |
| `DeltaSyncService` | Incremental change tracking via delta links |
| `UserIdConverterService` | Convert between external and internal user IDs |
| `MicrosoftSubscriptionService` | Webhook subscription lifecycle management |
| `GraphRateLimiterService` | Microsoft Graph API rate limit handling |

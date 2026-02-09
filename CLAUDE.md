# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is `@checkfirst/nestjs-outlook`, an opinionated NestJS module for Microsoft Outlook integration via Microsoft Graph API. It provides authentication, calendar management, email sending, and real-time webhook notifications.

## Development Commands

```bash
# Build the library
npm run build

# Development with hot reload (watches and pushes to yalc)
npm run dev

# Linting
npm run lint              # Check for errors
npm run lint:fix          # Auto-fix issues

# Testing
npm run test              # Run all tests
npm run test:watch        # Watch mode
npm run test:cov          # With coverage
```

### Sample Application

The sample app in `samples/nestjs-outlook-example/` demonstrates integration patterns:

```bash
cd samples/nestjs-outlook-example

# Run with yalc-linked library
npm run start:dev:yalc

# Database management
npm run migration:run     # Run TypeORM migrations
npm run reset-and-migrate # Reset SQLite and migrate
```

### Local Development with yalc

1. Build and publish library: `npm run build && yalc publish`
2. In sample app: `yalc add @checkfirst/nestjs-outlook`
3. Run `npm run dev` in library root (watches and pushes changes)
4. Run `npm run start:dev:yalc` in sample app

## Architecture

### Module Structure

The library exports a single configurable NestJS module:

```
src/
├── microsoft-outlook.module.ts   # Main module with forRoot() configuration
├── controllers/
│   ├── microsoft-auth.controller.ts  # OAuth callback handling
│   ├── calendar.controller.ts        # Webhook endpoint + calendar ops
│   └── email.controller.ts           # Email webhook endpoint
├── services/
│   ├── auth/microsoft-auth.service.ts      # OAuth flow, token management
│   ├── calendar/calendar.service.ts        # Calendar CRUD, delta sync, webhooks
│   ├── email/email.service.ts              # Email sending + webhooks
│   ├── subscription/microsoft-subscription.service.ts  # Webhook subscription management
│   └── shared/
│       ├── delta-sync.service.ts           # Microsoft Graph delta API handling
│       └── user-id-converter.service.ts    # External↔Internal ID mapping
├── entities/                        # TypeORM entities
├── repositories/                    # Data access layer
├── migrations/                      # Database migrations (must run in order)
└── enums/                           # Permission scopes, event types
```

### Key Concepts

**User ID Terminology:**
- `externalUserId` (string): ID from the host application using this library
- `internalUserId` (number): Auto-generated primary key in `MicrosoftUser` entity

**Event-Driven Architecture:**
The module emits events via `@nestjs/event-emitter`:
- `USER_AUTHENTICATED` - OAuth flow completed
- `EVENT_CREATED/UPDATED/DELETED` - Calendar changes via delta sync
- `EMAIL_RECEIVED/UPDATED/DELETED` - Email notifications

**Delta Sync:**
Calendar and email services use Microsoft Graph's delta API for efficient change tracking. The `DeltaSyncService` manages delta links and provides both buffering (fetch all then process) and streaming (page-by-page) modes.

### Configuration

```typescript
MicrosoftOutlookModule.forRoot({
  clientId: 'MICROSOFT_CLIENT_ID',
  clientSecret: 'MICROSOFT_CLIENT_SECRET',
  redirectPath: 'auth/microsoft/callback',
  backendBaseUrl: 'https://your-api.example.com',
  basePath: 'api/v1',  // optional
})
```

### Database Tables

Migrations create these tables (run in order):
1. `outlook_webhook_subscriptions` - Webhook subscription tracking
2. `microsoft_csrf_tokens` - OAuth CSRF protection
3. `microsoft_users` - Token storage and user mapping

## ESLint Configuration

Uses TypeScript ESLint strict config with these notable rules:
- `@typescript-eslint/no-unused-vars` - Prefix with `_` to ignore
- `eslint-comments/require-description` - ESLint disable comments need explanations
- `import/no-absolute-path` - Enforce relative imports

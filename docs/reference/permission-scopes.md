---
dep:
  type: reference
  audience: [consumer-dev]
  owner: "@checkfirst-ltd"
  created: 2026-04-03T12:00:00+03:00
  last_verified: 2026-04-03T12:00:00+03:00
  confidence: medium
  depends_on:
    - ../../src/enums/permission-scope.enum.ts
  tags: [permissions, scopes, oauth]
  links:
    - target: ../explanation/oauth-flow.md
      rel: NEXT
---

# Permission Scopes Reference

Provider-agnostic permission scopes that map to Microsoft Graph API scopes.

Source: `src/enums/permission-scope.enum.ts`

---

## `PermissionScope` Enum

| Enum Value | String | Microsoft Graph Scope | Description |
|------------|--------|----------------------|-------------|
| `CALENDAR_READ` | `'CALENDAR_READ'` | `Calendars.Read` | Read access to user's calendars |
| `CALENDAR_WRITE` | `'CALENDAR_WRITE'` | `Calendars.ReadWrite` | Read/write access to user's calendars |
| `EMAIL_READ` | `'EMAIL_READ'` | `Mail.Read` | Read access to user's email |
| `EMAIL_WRITE` | `'EMAIL_WRITE'` | `Mail.ReadWrite` | Read/write access to user's email |
| `EMAIL_SEND` | `'EMAIL_SEND'` | `Mail.Send` | Permission to send email as the user |

---

## Default Scopes

The `MicrosoftAuthService` always requests these scopes in addition to any custom scopes:

| Scope | Purpose |
|-------|---------|
| `offline_access` | Enables refresh token issuance |
| `User.Read` | Read basic user profile information |

The default permission scopes (requested unless overridden) are: `CALENDAR_READ`, `CALENDAR_WRITE`, `EMAIL_SEND`, `EMAIL_READ`, `EMAIL_WRITE`.

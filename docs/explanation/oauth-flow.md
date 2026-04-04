---
dep:
  type: explanation
  audience: [contributor, maintainer]
  owner: "@checkfirst-ltd"
  created: 2026-04-03T12:00:00+03:00
  last_verified: 2026-04-03T12:00:00+03:00
  confidence: medium
  depends_on:
    - ../../src/services/auth/microsoft-auth.service.ts
    - ../../src/entities/microsoft-user.entity.ts
    - ../../src/entities/csrf-token.entity.ts
  tags: [oauth, authentication, tokens]
  links:
    - target: ../reference/module-configuration.md
      rel: EXPLAINS
    - target: ../reference/permission-scopes.md
      rel: EXPLAINS
    - target: ./architecture.md
      rel: REQUIRES
---

# OAuth Flow & Token Management

## Context

The module uses Microsoft's OAuth 2.0 authorization code flow to obtain access and refresh tokens for individual users. This explanation covers how the flow works, how tokens are stored and refreshed, and how CSRF protection is implemented.

## Authorization Code Flow

The flow involves three parties: the host app's end user, this module, and Microsoft's identity platform.

1. **Initiation**: The host app calls `MicrosoftAuthService.getLoginUrl(externalUserId, scopes?)` to generate a Microsoft consent URL. The method creates a CSRF token, stores it in the `microsoft_csrf_token` table with a 30-minute expiry, and encodes it along with the `externalUserId` in the OAuth `state` parameter.

2. **Consent**: The user is redirected to Microsoft's consent screen. The scopes requested include the `requiredScopes` (`offline_access`, `User.Read`) plus any configured `PermissionScope` values.

3. **Callback**: Microsoft redirects to the configured `redirectPath` with an authorization code. The `MicrosoftAuthController` handles this, validating the CSRF token from the state parameter against the stored token.

4. **Token Exchange**: `MicrosoftAuthService.exchangeCodeForToken()` sends the code to Microsoft's token endpoint and receives an access token, refresh token, and expiry timestamp.

5. **Storage**: Tokens are stored in the `microsoft_user` entity, keyed by `externalUserId`. If a user already exists, tokens are updated.

6. **Event Emission**: A `USER_AUTHENTICATED` event is emitted, allowing the host app to perform post-authentication setup (e.g., creating webhook subscriptions).

## Token Refresh

Access tokens expire (typically after 1 hour). The module handles refresh transparently:

- `getUserAccessToken()` checks `isTokenExpired()` with a 5-minute buffer before the actual expiry
- If expired, `refreshAccessToken()` exchanges the refresh token for a new access token
- The new tokens are saved to the database
- If the refresh token itself has expired or been revoked, the method throws — the host app should listen for this and prompt re-authentication

A `retryWithBackoff` utility wraps token operations to handle transient Microsoft API failures.

## CSRF Protection

The CSRF token mechanism prevents authorization code injection:

- Tokens are generated using `crypto.randomBytes()`
- Stored in a dedicated database table with an expiry timestamp
- Validated during the callback before any token exchange occurs
- A cron job (`cleanupExpiredTokens`) runs every 5 minutes to prune expired tokens

## Tradeoffs

**Database-backed CSRF tokens**: Using a database table instead of in-memory storage makes the module work in multi-instance deployments (e.g., behind a load balancer). The cost is a database round-trip for each auth initiation and callback.

**Tenant ID fixed to "common"**: The module uses `common` as the Azure AD tenant, meaning it supports personal, work, and school Microsoft accounts. This is the most permissive setting. Organizations requiring single-tenant restrictions would need to fork or extend this service.

**Transparent token refresh**: Token refresh is automatic but can mask failures. If a refresh token is revoked by Microsoft (e.g., password change), the error only surfaces when a service tries to make an API call.

## Related

- [Architecture](./architecture.md) — overall module structure
- [Permission Scopes Reference](../reference/permission-scopes.md) — available OAuth scopes
- [Module Configuration Reference](../reference/module-configuration.md) — OAuth-related config fields

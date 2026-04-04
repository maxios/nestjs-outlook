---
dep:
  type: reference
  audience: [ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-04-03T12:00:00+03:00
  last_verified: 2026-04-03T12:00:00+03:00
  confidence: medium
  depends_on:
    - ../../src/index.ts
    - ../../src/microsoft-outlook.module.ts
    - ../../src/enums/event-types.enum.ts
  tags: [ai, tasks, modification-guide]
  links:
    - target: ./ai-codebase-map.md
      rel: REQUIRES
    - target: ../how-to/ai-add-graph-api-feature.md
      rel: NEXT
---

# Common Tasks Reference

File-level checklist for common modifications. Follow these to avoid missing required touchpoints.

---

## Add a New Event Type

| Step | File | Action |
|------|------|--------|
| 1 | `src/enums/event-types.enum.ts` | Add new value to `OutlookEventTypes` enum |
| 2 | Service that emits it | Add `this.eventEmitter.emit(OutlookEventTypes.NEW_TYPE, payload)` |
| 3 | No export needed | Enum is already exported via `src/index.ts` |

---

## Add a New Public Service Method

| Step | File | Action |
|------|------|--------|
| 1 | `src/services/<domain>/<service>.ts` | Add the method |
| 2 | Use `externalUserId` | Never accept `internalUserId` in public methods |
| 3 | Use `executeGraphApiCall()` | For any Microsoft Graph API calls |
| 4 | Consult rate limiter | Call `GraphRateLimiterService` before Graph requests if doing bulk operations |

---

## Add a New Service

| Step | File | Action |
|------|------|--------|
| 1 | `src/services/<domain>/<service>.ts` | Create the service with `@Injectable()` |
| 2 | `src/microsoft-outlook.module.ts` | Add to `providers` array |
| 3 | `src/microsoft-outlook.module.ts` | Add to `exports` array if consumers need it |
| 4 | `src/index.ts` | Add `export * from './services/<domain>/<service>'` |

---

## Add a New Controller Endpoint

| Step | File | Action |
|------|------|--------|
| 1 | `src/controllers/<controller>.ts` | Add the route handler |
| 2 | Controller is already registered | Check `src/microsoft-outlook.module.ts` controllers array |

---

## Add a New Entity / Schema Change

| Step | File | Action |
|------|------|--------|
| 1 | `src/entities/<entity>.ts` | Create/modify the entity |
| 2 | `src/microsoft-outlook.module.ts` | Add to `TypeOrmModule.forFeature([...])` if new entity |
| 3 | `src/migrations/` | Create migration: `{timestamp}-{DescriptiveName}.ts` |
| 4 | `src/index.ts` | Export the entity if consumers need it |

---

## Add a New Enum Value

| Step | File | Action |
|------|------|--------|
| 1 | `src/enums/<enum>.ts` | Add the value |
| 2 | No export needed | Enums are already barrel-exported via `src/index.ts` |

---

## Add a New Interface / Type

| Step | File | Action |
|------|------|--------|
| 1 | `src/interfaces/<domain>/<interface>.ts` or `src/types/` | Create the type |
| 2 | `src/index.ts` | Add export if consumers need it |

---

## Add a New Permission Scope

| Step | File | Action |
|------|------|--------|
| 1 | `src/enums/permission-scope.enum.ts` | Add value to `PermissionScope` |
| 2 | `src/services/auth/microsoft-auth.service.ts` | Add scope mapping to Microsoft Graph scope string |
| 3 | Optionally update `defaultScopes` | If it should be requested by default |

---

## Modify Webhook Processing

| Step | File | Action |
|------|------|--------|
| 1 | `src/services/calendar/calendar.service.ts` | Modify `handleOutlookWebhook()` or `handleOutlookWebhookV2()` |
| 2 | `src/services/email/email.service.ts` | Modify `handleEmailWebhook()` for email changes |
| 3 | `src/dto/outlook-webhook-notification.dto.ts` | Modify DTO if payload structure changes |

---

## Common Pitfalls

| Pitfall | Guidance |
|---------|----------|
| Circular dependency | `MicrosoftAuthService` ↔ `CalendarService` ↔ `EmailService` use `forwardRef()`. New cross-service deps may need the same pattern. |
| Forgetting `src/index.ts` | Any new public type, service, or controller must be exported from the barrel file. |
| Using `internalUserId` in public API | Always use `externalUserId` (string). Use `UserIdConverterService` internally. |
| Skipping rate limiter | All Graph API calls in bulk/batch operations must go through `GraphRateLimiterService`. |
| Missing migration | Any entity change requires a new migration file in `src/migrations/`. |
| Batch size limit | Microsoft Graph `$batch` supports max 20 requests per batch. Split larger sets. |

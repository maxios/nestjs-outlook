---
dep:
  type: how-to
  audience: [ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-04-03T12:00:00+03:00
  last_verified: 2026-04-03T12:00:00+03:00
  confidence: medium
  depends_on:
    - ../../src/microsoft-outlook.module.ts
    - ../../src/index.ts
    - ../../src/services/calendar/calendar.service.ts
  tags: [ai, how-to, new-feature, graph-api]
  links:
    - target: ../reference/ai-codebase-map.md
      rel: REQUIRES
    - target: ../reference/ai-common-tasks.md
      rel: REQUIRES
    - target: ../reference/module-configuration.md
      rel: USES
---

# How-To: Add a New Microsoft Graph API Feature

**Goal**: Add a new Graph API capability (e.g., contacts, tasks, or a new calendar operation) following the module's established patterns.

## Prerequisites

- Read [Codebase Map](../reference/ai-codebase-map.md) for orientation
- Read [Common Tasks](../reference/ai-common-tasks.md) for checklist

## Steps

1. **Create or extend the service** in the appropriate domain folder:

   ```
   src/services/<domain>/<service>.ts
   ```

   Follow this template for a new public method:

   ```typescript
   async newMethod(
     externalUserId: string,  // Always use externalUserId for public API
     ...params: any[]
   ): Promise<ReturnType> {
     const accessToken = await this.microsoftAuthService.getUserAccessToken({ externalUserId });

     const client = Client.init({
       authProvider: (done) => { done(null, accessToken); },
     });

     return executeGraphApiCall(async () => {
       return client.api('/me/<resource>').get();
     });
   }
   ```

2. **Add any new types** to the appropriate location:
   - Microsoft Graph type re-exports → `src/types/microsoft-graph.types.ts`
   - Custom interfaces → `src/interfaces/<domain>/`
   - New enums → `src/enums/`

3. **Emit events** for any async notifications to the host app:

   ```typescript
   // Add to src/enums/event-types.enum.ts
   NEW_EVENT = 'outlook.<domain>.new_event',

   // Emit in your service
   this.eventEmitter.emit(OutlookEventTypes.NEW_EVENT, { externalUserId, data });
   ```

4. **Register in the module** (`src/microsoft-outlook.module.ts`):
   - Add to `providers` array
   - Add to `exports` array if consumers need direct access
   - Add new entities to `TypeOrmModule.forFeature([])` if applicable

5. **Export from barrel** (`src/index.ts`):

   ```typescript
   export * from './services/<domain>/<service>';
   ```

6. **Handle rate limiting** for bulk operations:

   ```typescript
   await this.rateLimiter.waitForSlot(externalUserId);
   // then make the Graph API call
   ```

7. **Add batch support** if the feature involves multiple items:
   - Use Microsoft's `$batch` endpoint
   - Max 20 requests per batch
   - Handle per-item failures in the batch response

8. **Create a migration** if you added or modified entities:
   - File: `src/migrations/{timestamp}-{DescriptiveName}.ts`
   - Use TypeORM migration API

## Verification

- [ ] Service compiles: `npm run build`
- [ ] Tests pass: `npm run test`
- [ ] Lint passes: `npm run lint`
- [ ] New public types/services are exported from `src/index.ts`
- [ ] Service is registered in `src/microsoft-outlook.module.ts` providers and exports
- [ ] All public methods use `externalUserId`, not `internalUserId`
- [ ] Graph API calls use `executeGraphApiCall()` wrapper
- [ ] Bulk operations respect rate limiter

## Related

- [Codebase Map](../reference/ai-codebase-map.md) — file locations and patterns
- [Common Tasks](../reference/ai-common-tasks.md) — checklists for other modifications

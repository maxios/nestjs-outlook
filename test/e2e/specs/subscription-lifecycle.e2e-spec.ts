/**
 * Test cases:
 *   G-01            create webhook subscription
 *   H-01, H-02, H-03, H-05  renew webhook subscription
 *   I-01, I-02, I-03  delete webhook subscription
 *   K-01, K-02       getActiveSubscriptionForUser
 *   L-02             cleanupSubscriptions continues past failures
 *   N-02, N-03, N-04, N-05, N-07  health check
 *
 * G-04 (P1: 5xx propagation through full retry path) is deferred to a follow-up
 * spec because exercising it requires either fake-timer plumbing or a 7-retry
 * HTTP wait that's noisy without much extra signal.
 */
import nock from 'nock';

import { createTestApp, E2ETestHarness } from '../helpers/test-module';
import {
  setupNock,
  teardownNock,
  clearNock,
  mockCreateSubscription,
  mockPatchSubscription,
  mockDeleteSubscription,
  mockListSubscriptions,
  mockBatchSubscriptionVerify,
} from '../helpers/graph-nock';
import { silenceLibraryLogs, restoreLibraryLogs } from '../helpers/silence';
import { shortCircuitDelay, restoreDelay } from '../helpers/fast-delay';
import { seedMicrosoftUser, seedSubscription } from '../helpers/seed';
import { MicrosoftUserStatus } from '../../../src/enums/microsoft-user-status.enum';
import { OutlookEventTypes } from '../../../src/enums/event-types.enum';
import { OutlookWebhookSubscription } from '../../../src/entities/outlook-webhook-subscription.entity';

describe('subscription: full lifecycle', () => {
  let harness: E2ETestHarness;

  beforeAll(() => {
    silenceLibraryLogs();
    setupNock();
    shortCircuitDelay();
  });

  afterAll(() => {
    restoreLibraryLogs();
    teardownNock();
    restoreDelay();
  });

  beforeEach(async () => {
    clearNock();
    harness = await createTestApp();
  });

  afterEach(async () => {
    await harness.close();
    expect(nock.pendingMocks()).toEqual([]);
  });

  // ─── G. createWebhookSubscription ─────────────────────────────────────

  it('G-01: builds correct payload and persists local subscription row', async () => {
    const user = await seedMicrosoftUser(harness, { externalUserId: 'ext-1' });
    const create = mockCreateSubscription({ id: 'sub-cal' });

    const before = Date.now();
    await harness.subscriptionService.createWebhookSubscription('ext-1');
    const after = Date.now();

    expect(create.capturedBodies).toHaveLength(1);
    const body = create.capturedBodies[0]!;
    expect(body.changeType).toBe('created,updated,deleted');
    expect(body.resource).toBe('/me/events');
    expect(body.notificationUrl).toBe('https://app.test/api/calendar/webhook');
    expect(body.lifecycleNotificationUrl).toBe('https://app.test/api/calendar/webhook');
    expect(typeof body.clientState).toBe('string');
    expect(body.clientState as string).toMatch(
      new RegExp(`^user_${user.id}_[0-9a-f-]{36}$`),
    );
    const expiration = new Date(body.expirationDateTime as string).getTime();
    expect(expiration - before).toBeGreaterThan(71 * 3600 * 1000);
    expect(expiration - after).toBeLessThan(73 * 3600 * 1000);

    const row = await harness.dataSource
      .getRepository(OutlookWebhookSubscription)
      .findOne({ where: { subscriptionId: 'sub-cal' } });
    expect(row).not.toBeNull();
    expect(row?.userId).toBe(user.id);
    expect(row?.resource).toBe('/me/events');
  });

  // ─── H. renewWebhookSubscription ──────────────────────────────────────

  it('H-01: happy renewal updates local expiry', async () => {
    const user = await seedMicrosoftUser(harness, { externalUserId: 'ext-1' });
    const sub = await seedSubscription(harness, {
      userId: user.id,
      subscriptionId: 'sub-renew',
      expirationDateTime: new Date(Date.now() + 24 * 3600 * 1000),
    });
    const newExpiration = new Date(Date.now() + 72 * 3600 * 1000).toISOString();
    mockPatchSubscription({ id: 'sub-renew', expirationDateTime: newExpiration });

    await harness.subscriptionService.renewWebhookSubscription('sub-renew', user.id);

    const updated = await harness.dataSource
      .getRepository(OutlookWebhookSubscription)
      .findOne({ where: { id: sub.id } });
    expect(updated?.expirationDateTime.toISOString()).toBe(newExpiration);
    expect(updated?.isActive).toBe(true);
  });

  it('H-02: missing/inactive user → deactivates sub locally and throws (no PATCH issued)', async () => {
    const user = await seedMicrosoftUser(harness, { externalUserId: 'ext-1', isActive: false });
    await seedSubscription(harness, { userId: user.id, subscriptionId: 'sub-orphan' });

    await expect(
      harness.subscriptionService.renewWebhookSubscription('sub-orphan', user.id),
    ).rejects.toThrow(/Cannot renew subscription/);

    const after = await harness.dataSource
      .getRepository(OutlookWebhookSubscription)
      .findOne({ where: { subscriptionId: 'sub-orphan' } });
    expect(after?.isActive).toBe(false);
  });

  it('H-03: PATCH 404 → deactivates local, recreates, emits SUBSCRIPTION_RECREATED', async () => {
    const user = await seedMicrosoftUser(harness, { externalUserId: 'ext-1' });
    await seedSubscription(harness, { userId: user.id, subscriptionId: 'sub-gone' });
    mockPatchSubscription({ id: 'sub-gone', status: 404, body: { error: { code: 'NotFound' } } });
    mockCreateSubscription({ id: 'sub-fresh' });

    await harness.subscriptionService.renewWebhookSubscription('sub-gone', user.id);

    const oldRow = await harness.dataSource
      .getRepository(OutlookWebhookSubscription)
      .findOne({ where: { subscriptionId: 'sub-gone' } });
    expect(oldRow?.isActive).toBe(false);

    const newRow = await harness.dataSource
      .getRepository(OutlookWebhookSubscription)
      .findOne({ where: { subscriptionId: 'sub-fresh' } });
    expect(newRow?.isActive).toBe(true);
    expect(newRow?.userId).toBe(user.id);

    const recreatedEvents = harness.events.filter((e) => e.name === OutlookEventTypes.SUBSCRIPTION_RECREATED);
    expect(recreatedEvents).toHaveLength(1);
    expect(recreatedEvents[0]?.args[0]).toMatchObject({
      subscriptionId: 'sub-gone',
      userId: user.id,
      reason: 'renewal_404',
    });
  });

  it('H-05: PATCH 401 → deactivates sub, emits SUBSCRIPTION_AUTH_FAILED, throws', async () => {
    const user = await seedMicrosoftUser(harness, { externalUserId: 'ext-1' });
    await seedSubscription(harness, { userId: user.id, subscriptionId: 'sub-401' });
    mockPatchSubscription({ id: 'sub-401', status: 401, body: { error: { code: 'Unauthorized' } } });

    await expect(
      harness.subscriptionService.renewWebhookSubscription('sub-401', user.id),
    ).rejects.toThrow(/Failed to renew webhook subscription/);

    const row = await harness.dataSource
      .getRepository(OutlookWebhookSubscription)
      .findOne({ where: { subscriptionId: 'sub-401' } });
    expect(row?.isActive).toBe(false);

    const authFailEvents = harness.events.filter((e) => e.name === OutlookEventTypes.SUBSCRIPTION_AUTH_FAILED);
    expect(authFailEvents).toHaveLength(1);
    expect(authFailEvents[0]?.args[0]).toMatchObject({
      subscriptionId: 'sub-401',
      userId: user.id,
      statusCode: 401,
    });
  });

  // ─── I. deleteWebhookSubscription ─────────────────────────────────────

  it('I-01: last subscription deleted → user marked inactive', async () => {
    const user = await seedMicrosoftUser(harness, { externalUserId: 'ext-1' });
    await seedSubscription(harness, { userId: user.id, subscriptionId: 'sub-only' });
    mockDeleteSubscription({ id: 'sub-only', status: 204 });

    const ok = await harness.subscriptionService.deleteWebhookSubscription('sub-only', 'ext-1');
    expect(ok).toBe(true);

    const sub = await harness.dataSource
      .getRepository(OutlookWebhookSubscription)
      .findOne({ where: { subscriptionId: 'sub-only' } });
    expect(sub?.isActive).toBe(false);

    const refreshed = await harness.microsoftUserRepo.findOne({ where: { id: user.id } });
    expect(refreshed?.isActive).toBe(false);
  });

  it('I-02: user keeps active when other subscriptions remain', async () => {
    const user = await seedMicrosoftUser(harness, { externalUserId: 'ext-1' });
    await seedSubscription(harness, { userId: user.id, subscriptionId: 'sub-a' });
    await seedSubscription(harness, { userId: user.id, subscriptionId: 'sub-b' });
    mockDeleteSubscription({ id: 'sub-a', status: 204 });

    await harness.subscriptionService.deleteWebhookSubscription('sub-a', 'ext-1');

    const refreshed = await harness.microsoftUserRepo.findOne({ where: { id: user.id } });
    expect(refreshed?.isActive).toBe(true);

    const remaining = await harness.dataSource
      .getRepository(OutlookWebhookSubscription)
      .findOne({ where: { subscriptionId: 'sub-b' } });
    expect(remaining?.isActive).toBe(true);
  });

  it('I-03: Graph 404 still cleans local row and returns true', async () => {
    const user = await seedMicrosoftUser(harness, { externalUserId: 'ext-1' });
    await seedSubscription(harness, { userId: user.id, subscriptionId: 'sub-404' });
    mockDeleteSubscription({ id: 'sub-404', status: 404 });

    const ok = await harness.subscriptionService.deleteWebhookSubscription('sub-404', 'ext-1');
    expect(ok).toBe(true);

    const row = await harness.dataSource
      .getRepository(OutlookWebhookSubscription)
      .findOne({ where: { subscriptionId: 'sub-404' } });
    expect(row?.isActive).toBe(false);
  });

  // ─── K. getActiveSubscriptionForUser ──────────────────────────────────

  it('K-01: active user with active subscription returns the subscription id', async () => {
    const user = await seedMicrosoftUser(harness, { externalUserId: 'ext-1' });
    await seedSubscription(harness, { userId: user.id, subscriptionId: 'sub-current' });

    const id = await harness.subscriptionService.getActiveSubscriptionForUser('ext-1');
    expect(id).toBe('sub-current');
  });

  it('K-02: CORRUPTED user returns null (so UI re-prompts OAuth)', async () => {
    const user = await seedMicrosoftUser(harness, {
      externalUserId: 'ext-1',
      status: MicrosoftUserStatus.CORRUPTED,
    });
    await seedSubscription(harness, { userId: user.id, subscriptionId: 'sub-stale' });

    const id = await harness.subscriptionService.getActiveSubscriptionForUser('ext-1');
    expect(id).toBeNull();
  });

  // ─── L. cleanupSubscriptions ──────────────────────────────────────────

  it('L-02: cleanupSubscriptions continues past individual delete failures', async () => {
    await seedMicrosoftUser(harness, { externalUserId: 'ext-1' });
    mockListSubscriptions([
      { id: 'sub-1', resource: '/me/events', clientState: 'user_1_a' },
      { id: 'sub-2', resource: '/me/events', clientState: 'user_1_b' },
      { id: 'sub-3', resource: '/me/events', clientState: 'user_1_c' },
    ]);
    mockDeleteSubscription({ id: 'sub-1', status: 204 });
    // 401 is in the lib's non-retryable set, so `executeGraphApiCall` throws on
    // the first attempt without entering the ~127s exponential-backoff loop.
    // The behavior under test — cleanup continues past the failed delete — is
    // identical regardless of which non-204 status sub-2 returns.
    mockDeleteSubscription({ id: 'sub-2', status: 401 });
    mockDeleteSubscription({ id: 'sub-3', status: 204 });

    const result = await harness.subscriptionService.cleanupSubscriptions({ accessToken: 'fake' });

    expect(result.totalFound).toBe(3);
    expect(result.successfullyDeleted).toBe(2);
    expect(result.failedToDelete).toBe(1);
    expect(result.deletedSubscriptionIds.sort()).toEqual(['sub-1', 'sub-3']);
    expect(result.errors).toHaveLength(1);
    expect(result.errors[0]?.subscriptionId).toBe('sub-2');
  });

  // ─── N. verifySubscriptionHealth ──────────────────────────────────────

  it('N-02: only subscriptions expiring within 24h enter the /$batch payload', async () => {
    const user = await seedMicrosoftUser(harness, { externalUserId: 'ext-1' });
    await seedSubscription(harness, {
      userId: user.id,
      subscriptionId: 'sub-soon',
      expirationDateTime: new Date(Date.now() + 6 * 3600 * 1000),
    });
    await seedSubscription(harness, {
      userId: user.id,
      subscriptionId: 'sub-far',
      expirationDateTime: new Date(Date.now() + 48 * 3600 * 1000),
    });

    const batch = mockBatchSubscriptionVerify({
      responses: [
        {
          id: '0',
          status: 200,
          body: {
            id: 'sub-soon',
            expirationDateTime: new Date(Date.now() + 24 * 3600 * 1000).toISOString(),
          },
        },
      ],
    });

    await harness.subscriptionService.verifySubscriptionHealth();

    expect(batch.capturedBodies).toHaveLength(1);
    const requests = (batch.capturedBodies[0]?.requests ?? []) as Array<{ url: string }>;
    expect(requests).toHaveLength(1);
    expect(requests[0]?.url).toBe('/subscriptions/sub-soon');
  });

  it('N-03: batch 200 with comfortable expiry → no renewal, no recreate, no events', async () => {
    const user = await seedMicrosoftUser(harness, { externalUserId: 'ext-1' });
    await seedSubscription(harness, {
      userId: user.id,
      subscriptionId: 'sub-soon',
      expirationDateTime: new Date(Date.now() + 6 * 3600 * 1000),
    });

    mockBatchSubscriptionVerify({
      responses: [
        {
          id: '0',
          status: 200,
          body: {
            id: 'sub-soon',
            // Comfortable expiry — no forced renewal.
            expirationDateTime: new Date(Date.now() + 24 * 3600 * 1000).toISOString(),
          },
        },
      ],
    });

    await harness.subscriptionService.verifySubscriptionHealth();

    expect(harness.events.filter((e) => String(e.name).startsWith('subscription'))).toHaveLength(0);
  });

  it('N-04: batch 200 with expirationDateTime <12h → forced renewal via PATCH', async () => {
    const user = await seedMicrosoftUser(harness, { externalUserId: 'ext-1' });
    await seedSubscription(harness, {
      userId: user.id,
      subscriptionId: 'sub-soon',
      expirationDateTime: new Date(Date.now() + 6 * 3600 * 1000),
    });

    const renewedExpiration = new Date(Date.now() + 72 * 3600 * 1000).toISOString();
    mockBatchSubscriptionVerify({
      responses: [
        {
          id: '0',
          status: 200,
          body: {
            id: 'sub-soon',
            // <12h forces renewal.
            expirationDateTime: new Date(Date.now() + 6 * 3600 * 1000).toISOString(),
          },
        },
      ],
    });
    mockPatchSubscription({ id: 'sub-soon', expirationDateTime: renewedExpiration });

    await harness.subscriptionService.verifySubscriptionHealth();

    const updated = await harness.dataSource
      .getRepository(OutlookWebhookSubscription)
      .findOne({ where: { subscriptionId: 'sub-soon' } });
    expect(updated?.expirationDateTime.toISOString()).toBe(renewedExpiration);
  });

  it('N-05: batch 404 → deactivates, recreates, emits SUBSCRIPTION_RECREATED', async () => {
    const user = await seedMicrosoftUser(harness, { externalUserId: 'ext-1' });
    await seedSubscription(harness, {
      userId: user.id,
      subscriptionId: 'sub-gone',
      expirationDateTime: new Date(Date.now() + 6 * 3600 * 1000),
    });

    mockBatchSubscriptionVerify({
      responses: [
        { id: '0', status: 404, body: { error: { code: 'NotFound' } } },
      ],
    });
    mockCreateSubscription({ id: 'sub-replacement' });

    await harness.subscriptionService.verifySubscriptionHealth();

    const oldRow = await harness.dataSource
      .getRepository(OutlookWebhookSubscription)
      .findOne({ where: { subscriptionId: 'sub-gone' } });
    expect(oldRow?.isActive).toBe(false);

    const newRow = await harness.dataSource
      .getRepository(OutlookWebhookSubscription)
      .findOne({ where: { subscriptionId: 'sub-replacement' } });
    expect(newRow?.isActive).toBe(true);

    const recreated = harness.events.filter((e) => e.name === OutlookEventTypes.SUBSCRIPTION_RECREATED);
    expect(recreated).toHaveLength(1);
    expect(recreated[0]?.args[0]).toMatchObject({
      subscriptionId: 'sub-gone',
      userId: user.id,
      reason: 'health_check_404',
    });
  });

  it('N-07: subscriptions owned by CORRUPTED users are excluded from /$batch payloads', async () => {
    const goodUser = await seedMicrosoftUser(harness, { externalUserId: 'ext-good' });
    const corruptedUser = await seedMicrosoftUser(harness, {
      externalUserId: 'ext-corrupt',
      status: MicrosoftUserStatus.CORRUPTED,
    });
    await seedSubscription(harness, {
      userId: goodUser.id,
      subscriptionId: 'sub-good',
      expirationDateTime: new Date(Date.now() + 6 * 3600 * 1000),
    });
    await seedSubscription(harness, {
      userId: corruptedUser.id,
      subscriptionId: 'sub-corrupt',
      expirationDateTime: new Date(Date.now() + 6 * 3600 * 1000),
    });

    const batch = mockBatchSubscriptionVerify({
      responses: [
        {
          id: '0',
          status: 200,
          body: {
            id: 'sub-good',
            expirationDateTime: new Date(Date.now() + 24 * 3600 * 1000).toISOString(),
          },
        },
      ],
    });

    await harness.subscriptionService.verifySubscriptionHealth();

    expect(batch.capturedBodies).toHaveLength(1);
    const requests = (batch.capturedBodies[0]?.requests ?? []) as Array<{ url: string }>;
    expect(requests).toHaveLength(1);
    expect(requests[0]?.url).toBe('/subscriptions/sub-good');
  });
});

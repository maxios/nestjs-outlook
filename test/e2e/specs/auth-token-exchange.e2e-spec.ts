/**
 * Test cases C-01, C-02, C-03, C-04, C-05, C-07, C-08, C-09 (token exchange via service)
 * and E-01, E-02, E-03, E-04 (OAuth callback via HTTP).
 */
import request from 'supertest';
import nock from 'nock';

import { createTestApp, E2ETestHarness } from '../helpers/test-module';
import {
  setupNock,
  teardownNock,
  clearNock,
  mockTokenEndpointSuccess,
  mockTokenEndpointError,
  mockMailboxOk,
  mockMailboxInactive,
  mockListSubscriptions,
  mockCreateSubscription,
} from '../helpers/graph-nock';
import { silenceLibraryLogs, restoreLibraryLogs } from '../helpers/silence';
import { buildValidState, encodeState } from '../helpers/state';
import { seedMicrosoftUser } from '../helpers/seed';
import { PermissionScope } from '../../../src/enums/permission-scope.enum';
import { OutlookEventTypes } from '../../../src/enums/event-types.enum';
import { MicrosoftUserStatus } from '../../../src/enums/microsoft-user-status.enum';
import { MailboxInactiveError } from '../../../src/errors/mailbox-inactive.error';
import { MicrosoftUser } from '../../../src/entities/microsoft-user.entity';
import { OutlookWebhookSubscription } from '../../../src/entities/outlook-webhook-subscription.entity';

describe('auth: token exchange + OAuth callback', () => {
  let harness: E2ETestHarness;

  beforeAll(() => {
    silenceLibraryLogs();
    setupNock();
  });

  afterAll(() => {
    restoreLibraryLogs();
    teardownNock();
  });

  beforeEach(async () => {
    clearNock();
    harness = await createTestApp();
  });

  afterEach(async () => {
    await harness.close();
    expect(nock.pendingMocks()).toEqual([]);
  });

  // ─── C. exchangeCodeForToken ─────────────────────────────────────────

  it('C-01: happy path persists user, creates calendar+email webhooks, emits USER_AUTHENTICATED', async () => {
    const { state } = await buildValidState(harness, 'ext-1');

    mockTokenEndpointSuccess({ accessToken: 'fresh-access', refreshToken: 'fresh-refresh', expiresIn: 3600 });
    mockMailboxOk();
    mockListSubscriptions([], 2);
    const calendarMock = mockCreateSubscription({ id: 'sub-cal' });
    const emailMock = mockCreateSubscription({ id: 'sub-mail' });

    const result = await harness.authService.exchangeCodeForToken('AUTH_CODE', state);
    expect(result).toEqual({
      access_token: 'fresh-access',
      refresh_token: 'fresh-refresh',
      expires_in: 3600,
    });

    const user = await harness.microsoftUserRepo.findOne({ where: { externalUserId: 'ext-1' } });
    expect(user).not.toBeNull();
    expect(user?.isActive).toBe(true);
    expect(user?.status).toBe(MicrosoftUserStatus.ACTIVE);
    expect(user?.accessToken).toBe('fresh-access');
    expect(user?.refreshToken).toBe('fresh-refresh');
    expect((user?.tokenExpiry.getTime() ?? 0) - Date.now()).toBeGreaterThan(3500 * 1000);

    const subs = await harness.dataSource
      .getRepository(OutlookWebhookSubscription)
      .find({ where: { userId: user!.id } });
    expect(subs.map((s) => s.resource).sort()).toEqual(['/me/events', '/me/messages']);

    const authEvents = harness.events.filter((e) => e.name === OutlookEventTypes.USER_AUTHENTICATED);
    expect(authEvents).toHaveLength(1);
    expect(authEvents[0]?.args[0]).toBe('ext-1');
    expect(authEvents[0]?.args[1]).toMatchObject({
      externalUserId: 'ext-1',
      scopes: expect.arrayContaining([PermissionScope.CALENDAR_READ, PermissionScope.EMAIL_READ]),
    });

    expect(calendarMock.capturedBodies).toHaveLength(1);
    expect(emailMock.capturedBodies).toHaveLength(1);
  });

  it('C-02: invalid state → throws "Invalid state parameter"', async () => {
    await expect(harness.authService.exchangeCodeForToken('AUTH_CODE', 'not-base64')).rejects.toThrow(
      /Invalid state parameter/,
    );
    expect(await harness.microsoftUserRepo.count()).toBe(0);
  });

  it('C-03: CSRF mismatch → throws "CSRF validation failed"', async () => {
    const fabricated = encodeState({
      userId: 'ext-1',
      csrf: 'forged-csrf-not-in-db',
      timestamp: Date.now(),
      requestedScopes: [PermissionScope.CALENDAR_READ],
    });

    await expect(harness.authService.exchangeCodeForToken('AUTH_CODE', fabricated)).rejects.toThrow(
      /CSRF validation failed/,
    );
    expect(await harness.microsoftUserRepo.count()).toBe(0);
  });

  it('C-04: token endpoint 5xx → throws and no user is persisted', async () => {
    const { state } = await buildValidState(harness, 'ext-1');
    mockTokenEndpointError({ status: 500, description: 'boom' });

    await expect(harness.authService.exchangeCodeForToken('AUTH_CODE', state)).rejects.toThrow(
      /Failed to exchange code for token/,
    );
    expect(await harness.microsoftUserRepo.count()).toBe(0);
  });

  it('C-05: MailboxNotEnabledForRESTAPI → throws MailboxInactiveError and deactivates user', async () => {
    const { state } = await buildValidState(harness, 'ext-1');
    mockTokenEndpointSuccess();
    mockMailboxInactive();

    await expect(harness.authService.exchangeCodeForToken('AUTH_CODE', state)).rejects.toBeInstanceOf(
      MailboxInactiveError,
    );

    const user = await harness.microsoftUserRepo.findOne({ where: { externalUserId: 'ext-1' } });
    expect(user).not.toBeNull();
    expect(user?.isActive).toBe(false);
  });

  it('C-07: reconnection reuses CORRUPTED user row and flips status to ACTIVE', async () => {
    const seeded = await seedMicrosoftUser(harness, {
      externalUserId: 'ext-1',
      isActive: false,
      status: MicrosoftUserStatus.CORRUPTED,
      accessToken: 'old-access',
      refreshToken: 'old-refresh',
    });

    const { state } = await buildValidState(harness, 'ext-1');
    mockTokenEndpointSuccess({ accessToken: 'new-access', refreshToken: 'new-refresh' });
    mockMailboxOk();
    mockListSubscriptions([], 2);
    mockCreateSubscription({ id: 'sub-cal' });
    mockCreateSubscription({ id: 'sub-mail' });

    await harness.authService.exchangeCodeForToken('AUTH_CODE', state);

    const allUsers = await harness.microsoftUserRepo.find();
    expect(allUsers).toHaveLength(1);
    expect(allUsers[0]?.id).toBe(seeded.id);
    expect(allUsers[0]?.isActive).toBe(true);
    expect(allUsers[0]?.status).toBe(MicrosoftUserStatus.ACTIVE);
    expect(allUsers[0]?.accessToken).toBe('new-access');
    expect(allUsers[0]?.refreshToken).toBe('new-refresh');
  });

  it('C-08: only EMAIL_SEND requested → no webhook subscriptions created, no /subscriptions HTTP calls', async () => {
    const { state } = await buildValidState(harness, 'ext-1', [PermissionScope.EMAIL_SEND]);
    mockTokenEndpointSuccess();
    mockMailboxOk();
    // Intentionally no /subscriptions stubs — any call would fail with "no match".

    await harness.authService.exchangeCodeForToken('AUTH_CODE', state);

    const subs = await harness.dataSource.getRepository(OutlookWebhookSubscription).find();
    expect(subs).toHaveLength(0);
  });

  it('C-09: CALENDAR_READ requested → calendar webhook created, email webhook NOT created', async () => {
    const { state } = await buildValidState(harness, 'ext-1', [PermissionScope.CALENDAR_READ]);
    mockTokenEndpointSuccess();
    mockMailboxOk();
    mockListSubscriptions([], 1);
    const calendarMock = mockCreateSubscription({ id: 'sub-cal' });

    await harness.authService.exchangeCodeForToken('AUTH_CODE', state);

    const subs = await harness.dataSource.getRepository(OutlookWebhookSubscription).find();
    expect(subs).toHaveLength(1);
    expect(subs[0]?.resource).toBe('/me/events');
    expect(calendarMock.capturedBodies[0]?.resource).toBe('/me/events');
  });

  // ─── E. OAuth callback HTTP endpoint ─────────────────────────────────

  it('E-01: GET /auth/microsoft/callback returns 200 + success HTML on happy path', async () => {
    const { state } = await buildValidState(harness, 'ext-1');
    mockTokenEndpointSuccess();
    mockMailboxOk();
    mockListSubscriptions([], 2);
    mockCreateSubscription({ id: 'sub-cal' });
    mockCreateSubscription({ id: 'sub-mail' });

    const response = await request(harness.app.getHttpServer())
      .get('/auth/microsoft/callback')
      .query({ code: 'AUTH_CODE', state });

    expect(response.status).toBe(200);
    expect(response.text).toContain('Authorization successful!');
    expect(response.text).toContain("postMessage('microsoft-auth-success'");
  });

  it('E-02: missing `code` returns 400', async () => {
    const response = await request(harness.app.getHttpServer())
      .get('/auth/microsoft/callback')
      .query({ state: 'something' });

    expect(response.status).toBe(400);
    expect(response.text).toBe('Missing required parameters');
  });

  it('E-03: missing `state` returns 400', async () => {
    const response = await request(harness.app.getHttpServer())
      .get('/auth/microsoft/callback')
      .query({ code: 'AUTH_CODE' });

    expect(response.status).toBe(400);
    expect(response.text).toBe('Missing required parameters');
  });

  it('E-04: MailboxInactiveError renders 200 + failure HTML', async () => {
    const { state } = await buildValidState(harness, 'ext-1');
    mockTokenEndpointSuccess();
    mockMailboxInactive();

    const response = await request(harness.app.getHttpServer())
      .get('/auth/microsoft/callback')
      .query({ code: 'AUTH_CODE', state });

    expect(response.status).toBe(200);
    expect(response.text).toContain('Calendar Connection Failed');
    expect(response.text).toContain('microsoft-auth-failed');

    const user = await harness.microsoftUserRepo.findOne({ where: { externalUserId: 'ext-1' } });
    expect(user?.isActive).toBe(false);
  });
});

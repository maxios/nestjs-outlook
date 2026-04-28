/**
 * Test cases D-01, D-02, D-04, D-05, D-08, D-09 (token refresh via getUserAccessToken).
 */
import nock from 'nock';

import { createTestApp, E2ETestHarness } from '../helpers/test-module';
import {
  setupNock,
  teardownNock,
  clearNock,
  mockTokenEndpointSuccess,
  mockTokenEndpointError,
} from '../helpers/graph-nock';
import { silenceLibraryLogs, restoreLibraryLogs } from '../helpers/silence';
import { seedMicrosoftUser } from '../helpers/seed';
import { MicrosoftUserStatus } from '../../../src/enums/microsoft-user-status.enum';
import { MicrosoftRefreshTokenInvalidError } from '../../../src/errors/microsoft-refresh-token-invalid.error';

describe('auth: token refresh', () => {
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

  it('D-01: non-expired token is returned without any HTTP call', async () => {
    await seedMicrosoftUser(harness, {
      externalUserId: 'ext-1',
      accessToken: 'still-fresh',
      refreshToken: 'rt-1',
      tokenExpiry: new Date(Date.now() + 60 * 60 * 1000),
    });

    const token = await harness.authService.getUserAccessToken({ externalUserId: 'ext-1' });

    expect(token).toBe('still-fresh');
    // The afterEach already asserts no pending mocks; here we additionally assert
    // no nock interceptors have been consumed (we didn't register any).
    expect(nock.activeMocks()).toEqual([]);
  });

  it('D-02: token within the 5-minute expiry buffer triggers a refresh, persists new tokens', async () => {
    const seeded = await seedMicrosoftUser(harness, {
      externalUserId: 'ext-1',
      accessToken: 'about-to-expire',
      refreshToken: 'rt-old',
      tokenExpiry: new Date(Date.now() + 4 * 60 * 1000),
    });

    mockTokenEndpointSuccess({
      accessToken: 'fresh-access-token',
      refreshToken: 'rt-new',
      expiresIn: 3600,
    });

    const token = await harness.authService.getUserAccessToken({ externalUserId: 'ext-1' });
    expect(token).toBe('fresh-access-token');

    const updated = await harness.microsoftUserRepo.findOne({ where: { id: seeded.id } });
    expect(updated?.accessToken).toBe('fresh-access-token');
    expect(updated?.refreshToken).toBe('rt-new');
    expect((updated?.tokenExpiry.getTime() ?? 0) - Date.now()).toBeGreaterThan(3500 * 1000);
  });

  it('D-04: refresh 400 invalid_grant flips user CORRUPTED and throws MicrosoftRefreshTokenInvalidError', async () => {
    const seeded = await seedMicrosoftUser(harness, {
      externalUserId: 'ext-1',
      tokenExpiry: new Date(Date.now() - 60 * 1000),
    });

    mockTokenEndpointError({ status: 400, errorCode: 'invalid_grant', description: 'token revoked' });

    await expect(
      harness.authService.getUserAccessToken({ externalUserId: 'ext-1' }),
    ).rejects.toBeInstanceOf(MicrosoftRefreshTokenInvalidError);

    const updated = await harness.microsoftUserRepo.findOne({ where: { id: seeded.id } });
    expect(updated?.status).toBe(MicrosoftUserStatus.CORRUPTED);
  });

  it('D-05: refresh 400 consent_required flips user CORRUPTED + typed error', async () => {
    const seeded = await seedMicrosoftUser(harness, {
      externalUserId: 'ext-1',
      tokenExpiry: new Date(Date.now() - 60 * 1000),
    });

    mockTokenEndpointError({ status: 400, errorCode: 'consent_required' });

    await expect(
      harness.authService.getUserAccessToken({ externalUserId: 'ext-1' }),
    ).rejects.toBeInstanceOf(MicrosoftRefreshTokenInvalidError);

    const updated = await harness.microsoftUserRepo.findOne({ where: { id: seeded.id } });
    expect(updated?.status).toBe(MicrosoftUserStatus.CORRUPTED);
  });

  it('D-08: refresh 400 invalid_client throws generic error and does NOT corrupt the user', async () => {
    const seeded = await seedMicrosoftUser(harness, {
      externalUserId: 'ext-1',
      tokenExpiry: new Date(Date.now() - 60 * 1000),
    });

    mockTokenEndpointError({ status: 400, errorCode: 'invalid_client' });

    await expect(
      harness.authService.getUserAccessToken({ externalUserId: 'ext-1' }),
    ).rejects.toThrow(/Failed to get valid access token/);

    const updated = await harness.microsoftUserRepo.findOne({ where: { id: seeded.id } });
    expect(updated?.status).toBe(MicrosoftUserStatus.ACTIVE);
  });

  it('D-09: refresh 500 server_error throws generic error and does NOT corrupt the user', async () => {
    const seeded = await seedMicrosoftUser(harness, {
      externalUserId: 'ext-1',
      tokenExpiry: new Date(Date.now() - 60 * 1000),
    });

    mockTokenEndpointError({ status: 500, description: 'transient' });

    await expect(
      harness.authService.getUserAccessToken({ externalUserId: 'ext-1' }),
    ).rejects.toThrow(/Failed to get valid access token/);

    const updated = await harness.microsoftUserRepo.findOne({ where: { id: seeded.id } });
    expect(updated?.status).toBe(MicrosoftUserStatus.ACTIVE);
  });
});

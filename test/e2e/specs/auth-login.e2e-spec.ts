/**
 * Test cases A-01, A-02, A-03 (login URL) and B-01, B-02, B-03 (CSRF validation).
 *
 * See: ~/.claude/plans/we-need-to-init-sequential-hummingbird.md
 */
import { createTestApp, E2ETestHarness } from '../helpers/test-module';
import { setupNock, teardownNock, clearNock } from '../helpers/graph-nock';
import { silenceLibraryLogs, restoreLibraryLogs } from '../helpers/silence';
import { PermissionScope } from '../../../src/enums/permission-scope.enum';
import { MicrosoftCsrfToken } from '../../../src/entities/csrf-token.entity';

describe('auth: login URL + CSRF validation', () => {
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
  });

  // ─── A. getLoginUrl ───────────────────────────────────────────────────

  it('A-01: default scopes produce a well-formed authorize URL', async () => {
    const url = await harness.authService.getLoginUrl('ext-1');

    const parsed = new URL(url);
    expect(parsed.origin).toBe('https://login.microsoftonline.com');
    expect(parsed.pathname).toBe('/common/oauth2/v2.0/authorize');

    const params = parsed.searchParams;
    expect(params.get('client_id')).toBe('test-client');
    expect(params.get('response_type')).toBe('code');
    expect(params.get('response_mode')).toBe('query');
    expect(params.get('redirect_uri')).toBe('https://app.test/api/auth/microsoft/callback');

    const scopeParam = params.get('scope');
    expect(scopeParam).not.toBeNull();
    const scopeTokens = (scopeParam ?? '').split(' ');
    expect(scopeTokens).toEqual(
      expect.arrayContaining([
        'offline_access',
        'User.Read',
        'Calendars.Read',
        'Calendars.ReadWrite',
        'Mail.Read',
        'Mail.ReadWrite',
        'Mail.Send',
      ]),
    );
    expect(params.get('state')).toBeTruthy();
  });

  it('A-02: state decodes to {userId, csrf, timestamp, requestedScopes} and CSRF row is persisted', async () => {
    const before = Date.now();
    const url = await harness.authService.getLoginUrl('ext-1', [PermissionScope.CALENDAR_READ]);
    const after = Date.now();

    const state = new URL(url).searchParams.get('state');
    expect(state).toBeTruthy();
    const decoded = harness.authService.parseState(state ?? '');
    expect(decoded).not.toBeNull();
    expect(decoded?.userId).toBe('ext-1');
    expect(decoded?.requestedScopes).toEqual([PermissionScope.CALENDAR_READ]);
    expect(typeof decoded?.csrf).toBe('string');
    expect((decoded?.csrf ?? '').length).toBe(64);
    expect(/^[0-9a-f]{64}$/.test(decoded?.csrf ?? '')).toBe(true);
    expect(decoded?.timestamp).toBeGreaterThanOrEqual(before);
    expect(decoded?.timestamp).toBeLessThanOrEqual(after);

    const csrfRow = await harness.dataSource.getRepository(MicrosoftCsrfToken).findOne({
      where: { token: decoded?.csrf },
    });
    expect(csrfRow).not.toBeNull();
    expect(csrfRow?.userId).toBe('ext-1');
    const expiresMs = csrfRow?.expires.getTime() ?? 0;
    expect(expiresMs - before).toBeGreaterThan(29 * 60 * 1000);
    expect(expiresMs - after).toBeLessThan(31 * 60 * 1000);
  });

  it('A-03: custom scopes only include their Microsoft equivalents (plus required)', async () => {
    const url = await harness.authService.getLoginUrl('ext-1', [PermissionScope.EMAIL_SEND]);
    const scopeParam = new URL(url).searchParams.get('scope') ?? '';
    const tokens = scopeParam.split(' ').sort();

    expect(tokens).toEqual(['Mail.Send', 'User.Read', 'offline_access'].sort());
    expect(tokens).not.toContain('Calendars.Read');
    expect(tokens).not.toContain('Calendars.ReadWrite');
    expect(tokens).not.toContain('Mail.Read');
    expect(tokens).not.toContain('Mail.ReadWrite');
  });

  // ─── B. validateCsrfToken ─────────────────────────────────────────────

  it('B-01: empty token returns "Missing CSRF token"', async () => {
    const result = await harness.authService.validateCsrfToken('');
    expect(result).toBe('Missing CSRF token');
  });

  it('B-02: unknown token returns "Invalid or expired CSRF token"', async () => {
    const result = await harness.authService.validateCsrfToken('not-in-db');
    expect(result).toBe('Invalid or expired CSRF token');
  });

  it('B-03: a freshly seeded token validates as null', async () => {
    const token = 'a'.repeat(64);
    await harness.csrfRepo.saveToken(token, 'ext-1', 30 * 60 * 1000);

    const result = await harness.authService.validateCsrfToken(token);
    expect(result).toBeNull();

    // Per the repository contract, validation is one-time-use and removes the row.
    const remaining = await harness.dataSource.getRepository(MicrosoftCsrfToken).count();
    expect(remaining).toBe(0);
  });
});

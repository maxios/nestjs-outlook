import { createTestApp, E2ETestHarness } from '../helpers/test-module';
import { setupNock, teardownNock } from '../helpers/graph-nock';

describe('e2e harness smoke', () => {
  let harness: E2ETestHarness;

  beforeAll(() => {
    setupNock();
  });

  afterAll(() => {
    teardownNock();
  });

  beforeEach(async () => {
    harness = await createTestApp();
  });

  afterEach(async () => {
    await harness.close();
  });

  it('boots the module and resolves all expected services', () => {
    expect(harness.authService).toBeDefined();
    expect(harness.subscriptionService).toBeDefined();
    expect(harness.emailService).toBeDefined();
    expect(harness.calendarService).toBeDefined();
    expect(harness.rateLimiter).toBeDefined();
    expect(harness.dataSource.isInitialized).toBe(true);
  });

  it('persists a row in microsoft_csrf_tokens via the repository', async () => {
    await harness.csrfRepo.saveToken('a'.repeat(64), 'ext-1', 30 * 60 * 1000);
    const found = await harness.dataSource.getRepository('microsoft_csrf_tokens').findOne({
      where: { userId: 'ext-1' },
    });
    expect(found).not.toBeNull();
  });
});

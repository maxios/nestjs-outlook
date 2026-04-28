import nock from 'nock';

export const GRAPH_BASE = 'https://graph.microsoft.com';
export const GRAPH_V1 = `${GRAPH_BASE}/v1.0`;
export const MS_LOGIN_BASE = 'https://login.microsoftonline.com';
export const TOKEN_PATH = '/common/oauth2/v2.0/token';
export const LOGOUT_PATH = '/common/oauth2/v2.0/logout';

export function setupNock(): void {
  nock.disableNetConnect();
  // Allow loopback for supertest, which spins up a local server.
  nock.enableNetConnect((host) => /^(127\.0\.0\.1|::1|localhost)(?::\d+)?$/.test(host));
}

export function teardownNock(): void {
  nock.cleanAll();
  nock.enableNetConnect();
}

export function clearNock(): void {
  nock.cleanAll();
}

/** Stub the Microsoft Identity token endpoint with a successful token response. */
export function mockTokenEndpointSuccess(opts: {
  accessToken?: string;
  refreshToken?: string;
  expiresIn?: number;
  times?: number;
} = {}): nock.Scope {
  const body = {
    access_token: opts.accessToken ?? 'access-token-123',
    refresh_token: opts.refreshToken ?? 'refresh-token-123',
    expires_in: opts.expiresIn ?? 3600,
    token_type: 'Bearer',
    scope: 'offline_access User.Read',
  };
  return nock(MS_LOGIN_BASE).post(TOKEN_PATH).times(opts.times ?? 1).reply(200, body);
}

export function mockTokenEndpointError(opts: {
  status: number;
  errorCode?: string;
  description?: string;
  times?: number;
}): nock.Scope {
  const body = opts.errorCode
    ? { error: opts.errorCode, error_description: opts.description ?? '' }
    : { error_description: opts.description ?? 'unknown' };
  return nock(MS_LOGIN_BASE).post(TOKEN_PATH).times(opts.times ?? 1).reply(opts.status, body);
}

/** Mailbox validation succeeds. */
export function mockMailboxOk(times = 1): nock.Scope {
  return nock(GRAPH_BASE).get('/v1.0/me/mailboxSettings').times(times).reply(200, {
    timeZone: 'UTC',
  });
}

/** Mailbox validation returns the disabled-mailbox error. */
export function mockMailboxInactive(times = 1): nock.Scope {
  return nock(GRAPH_BASE)
    .get('/v1.0/me/mailboxSettings')
    .times(times)
    .reply(400, {
      error: {
        code: 'MailboxNotEnabledForRESTAPI',
        message: 'The mailbox is not enabled for REST API',
      },
    });
}

/** Mailbox validation returns a transient 5xx. */
export function mockMailboxTransient(status = 503, times = 1): nock.Scope {
  return nock(GRAPH_BASE).get('/v1.0/me/mailboxSettings').times(times).reply(status, 'transient');
}

/** Stub `GET /v1.0/subscriptions` (used by cleanup discovery). */
export function mockListSubscriptions(value: unknown[] = [], times = 1): nock.Scope {
  return nock(GRAPH_BASE).get('/v1.0/subscriptions').times(times).reply(200, { value });
}

export function mockListSubscriptionsError(status = 500, times = 1): nock.Scope {
  return nock(GRAPH_BASE).get('/v1.0/subscriptions').times(times).reply(status, 'err');
}

export interface CreatedSubscriptionMock {
  scope: nock.Scope;
  capturedBodies: Array<Record<string, unknown>>;
}

/**
 * Stub `POST /v1.0/subscriptions` returning a synthesized subscription body
 * that echoes back the resource and clientState the caller submitted.
 */
export function mockCreateSubscription(opts: {
  id?: string;
  times?: number;
} = {}): CreatedSubscriptionMock {
  const capturedBodies: Array<Record<string, unknown>> = [];
  let counter = 0;
  const baseId = opts.id ?? 'sub';
  const scope = nock(GRAPH_BASE)
    .post('/v1.0/subscriptions', (body: Record<string, unknown>) => {
      capturedBodies.push(body);
      return true;
    })
    .times(opts.times ?? 1)
    .reply(201, () => {
      const body = capturedBodies[capturedBodies.length - 1] ?? {};
      counter++;
      const id = opts.times && opts.times > 1 ? `${baseId}-${counter}` : baseId;
      return {
        id,
        resource: body.resource,
        changeType: body.changeType,
        clientState: body.clientState,
        notificationUrl: body.notificationUrl,
        expirationDateTime: body.expirationDateTime,
      };
    });
  return { scope, capturedBodies };
}

export function mockCreateSubscriptionError(status = 500, times = 1): nock.Scope {
  return nock(GRAPH_BASE).post('/v1.0/subscriptions').times(times).reply(status, 'err');
}

/** Stub `PATCH /v1.0/subscriptions/{id}` returning a renewed subscription body. */
export function mockPatchSubscription(opts: {
  id: string;
  expirationDateTime?: string;
  status?: number;
  body?: unknown;
}): nock.Scope {
  const status = opts.status ?? 200;
  const body =
    opts.body ??
    {
      id: opts.id,
      expirationDateTime: opts.expirationDateTime ?? new Date(Date.now() + 72 * 3600 * 1000).toISOString(),
    };
  return nock(GRAPH_BASE).patch(`/v1.0/subscriptions/${opts.id}`).reply(status, body);
}

/** Stub `DELETE /v1.0/subscriptions/{id}`. */
export function mockDeleteSubscription(opts: { id: string; status?: number; times?: number }): nock.Scope {
  return nock(GRAPH_BASE)
    .delete(`/v1.0/subscriptions/${opts.id}`)
    .times(opts.times ?? 1)
    .reply(opts.status ?? 204, '');
}

/** Stub `POST /v1.0/$batch` returning the supplied responses array. */
export function mockBatchSubscriptionVerify(opts: {
  responses: Array<{ id: string; status: number; body?: unknown }>;
  times?: number;
}): { scope: nock.Scope; capturedBodies: Array<Record<string, unknown>> } {
  const capturedBodies: Array<Record<string, unknown>> = [];
  const scope = nock(GRAPH_BASE)
    .post('/v1.0/$batch', (body: Record<string, unknown>) => {
      capturedBodies.push(body);
      return true;
    })
    .times(opts.times ?? 1)
    .reply(200, {
      responses: opts.responses.map((r) => ({
        id: r.id,
        status: r.status,
        body: r.body ?? {},
      })),
    });
  return { scope, capturedBodies };
}

/** Stub the Microsoft logout endpoint. */
export function mockLogoutEndpoint(opts: { status?: number; times?: number } = {}): nock.Scope {
  return nock(MS_LOGIN_BASE)
    .post(LOGOUT_PATH)
    .times(opts.times ?? 1)
    .reply(opts.status ?? 200, '');
}

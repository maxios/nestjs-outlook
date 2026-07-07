import { Client } from '@microsoft/microsoft-graph-client';
import { DeltaEvent, DeltaResponse, TokenProvider } from './types';

// Mock the Graph client factory so runDeltaSync drives a stubbed client
// instead of hitting the network.
jest.mock('./graph-client');
import { createGraphClient } from './graph-client';
import { runDeltaSync } from './run-delta-sync';

const mockedCreateGraphClient = createGraphClient as jest.MockedFunction<
  typeof createGraphClient
>;

/**
 * Build a fake Graph client whose `.api(url).header(...).get()` returns the
 * given pages in sequence (mimicking @odata.nextLink pagination).
 */
function fakeClient(pages: DeltaResponse<DeltaEvent>[]) {
  let call = 0;
  const get = jest.fn(() => Promise.resolve(pages[call++]));
  const header = jest.fn(() => ({ get }));
  const api = jest.fn(() => ({ header }));
  return { client: { api } as unknown as Client, api, header, get };
}

const tokenProvider: TokenProvider = { getAccessToken: jest.fn(async () => 'token') };

function ev(id: string): DeltaEvent {
  return { id, subject: `event-${id}` } as DeltaEvent;
}

describe('runDeltaSync', () => {
  beforeEach(() => jest.clearAllMocks());

  it('follows pagination across pages and accumulates into a final result', async () => {
    const pages: DeltaResponse<DeltaEvent>[] = [
      {
        value: [ev('a'), ev('b')],
        '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/events/delta?$skiptoken=PAGE2',
      },
      {
        value: [{ id: 'a', '@removed': { reason: 'deleted' } } as DeltaEvent, ev('c')],
        '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/me/events/delta?$deltatoken=NEW',
      },
    ];
    const { client, get } = fakeClient(pages);
    mockedCreateGraphClient.mockReturnValue(client);

    const result = await runDeltaSync({ userId: 'user-1', tokenProvider });

    // Both pages were fetched.
    expect(get).toHaveBeenCalledTimes(2);
    // 'a' was created then removed → net deleted; 'b' and 'c' survive.
    expect(result.events.map((e) => e.id).sort()).toEqual(['b', 'c']);
    expect(result.deletedIds).toEqual(['a']);
    // Terminal cursor is returned for the caller to persist.
    expect(result.deltaLink).toBe(
      'https://graph.microsoft.com/v1.0/me/events/delta?$deltatoken=NEW',
    );
    expect(result.stats.pages).toBe(2);
    expect(result.stats.totalChanges).toBe(4);
  });

  it('starts from the provided deltaLink when given one (incremental)', async () => {
    const pages: DeltaResponse<DeltaEvent>[] = [
      { value: [ev('x')], '@odata.deltaLink': 'DELTA_NEXT' },
    ];
    const { client, api } = fakeClient(pages);
    mockedCreateGraphClient.mockReturnValue(client);

    const result = await runDeltaSync({
      userId: 'user-1',
      tokenProvider,
      deltaLink: 'DELTA_PRIOR',
    });

    // The incremental cursor is used as the first request URL.
    expect(api).toHaveBeenCalledWith('DELTA_PRIOR');
    expect(result.deltaLink).toBe('DELTA_NEXT');
    expect(result.events.map((e) => e.id)).toEqual(['x']);
  });
});

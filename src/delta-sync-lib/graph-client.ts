import { Client } from '@microsoft/microsoft-graph-client';
import { TokenProvider } from './types';

/**
 * Build an authenticated Microsoft Graph client backed by a pluggable
 * {@link TokenProvider}. Mirrors `calendar.service.ts#getAuthenticatedClient`,
 * bridging the provider's async `getAccessToken` to the SDK's callback-style
 * `authProvider(done)` contract.
 *
 * The token is resolved lazily on each request the SDK makes, so a provider
 * that refreshes/caches tokens transparently keeps long streams alive.
 */
export function createGraphClient(tokenProvider: TokenProvider, userId: string): Client {
  return Client.init({
    authProvider: (done) => {
      tokenProvider
        .getAccessToken(userId)
        .then((token) => {
          done(null, token);
        })
        .catch((err: unknown) => {
          done(err instanceof Error ? err : new Error(String(err)), null);
        });
    },
  });
}

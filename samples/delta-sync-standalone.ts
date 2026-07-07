/**
 * Standalone delta-accumulation demo.
 *
 * Runs the framework-free delta library end-to-end against a real mailbox: it
 * streams the whole Microsoft Graph calendar delta, folds it into a final
 * result, and prints the accumulated events, deletions, and the new cursor.
 *
 * Auth is pluggable — here we use the simplest possible TokenProvider that
 * returns a static access token from the environment. In production you would
 * plug in your own refresh-grant / caching logic behind `getAccessToken`.
 *
 * Run (needs a valid delegated Graph token with Calendars.Read):
 *
 *   GRAPH_ACCESS_TOKEN=eyJ0... npx ts-node samples/delta-sync-standalone.ts
 *
 * Or, after `npm run build`, point the import at ../dist and run with node.
 *
 * Stateless-loop demo: the returned `deltaLink` is fed straight back into a
 * second call to show an incremental (near-empty) sync.
 */
import { runDeltaSync, TokenProvider } from '../src/delta-sync-lib';

async function main(): Promise<void> {
  const accessToken = process.env.GRAPH_ACCESS_TOKEN;
  if (!accessToken) {
    console.error('Set GRAPH_ACCESS_TOKEN to a valid Microsoft Graph access token.');
    process.exit(1);
  }

  const userId = process.env.GRAPH_USER_ID ?? 'demo-user';

  // The whole auth contract: hand back a valid Bearer token for this user.
  const tokenProvider: TokenProvider = {
    getAccessToken: async () => accessToken,
  };

  const logger = {
    log: (m: string) => console.log(m),
    warn: (m: string) => console.warn(m),
    error: (m: string, e?: unknown) => console.error(m, e ?? ''),
  };

  console.log('--- Run 1: initialize baseline (no prior cursor) ---');
  const first = await runDeltaSync({ userId, tokenProvider, logger });
  console.log({
    events: first.events.length,
    deleted: first.deletedIds.length,
    deltaLink: first.deltaLink ? `${first.deltaLink.slice(0, 60)}...` : null,
    stats: first.stats,
  });

  if (!first.deltaLink) {
    return;
  }

  console.log('\n--- Run 2: incremental (feed the returned cursor back in) ---');
  const second = await runDeltaSync({
    userId,
    tokenProvider,
    deltaLink: first.deltaLink,
    logger,
  });
  console.log({
    events: second.events.length,
    deleted: second.deletedIds.length,
    deltaLink: second.deltaLink ? `${second.deltaLink.slice(0, 60)}...` : null,
    stats: second.stats,
  });
}

main().catch((err) => {
  console.error('delta-sync-standalone failed:', err);
  process.exit(1);
});

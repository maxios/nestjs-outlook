#!/usr/bin/env node
/**
 * delta-accumulate — standalone, single-file CLI.
 *
 * Uses the framework-free delta-sync-lib to request the whole Microsoft Graph
 * calendar delta stream for a user and print the ACCUMULATED result: the net
 * created/updated events, the deleted ids, and the new delta cursor.
 *
 * Inputs (flags or env; flags win):
 *   --access-token   / GRAPH_ACCESS_TOKEN    (required) delegated Graph token
 *   --client-id      / GRAPH_CLIENT_ID       (required) Azure app registration id
 *   --client-secret  / GRAPH_CLIENT_SECRET   (required) app client secret
 *   --refresh-token  / GRAPH_REFRESH_TOKEN   (optional) enables auto-refresh
 *   --delta-link     / GRAPH_DELTA_LINK      (optional) prior cursor → incremental
 *   --user-id        / GRAPH_USER_ID         (optional) label passed to provider
 *   --resource-path  / GRAPH_RESOURCE_PATH   (optional) default /me/events/delta
 *   --scopes         / GRAPH_SCOPES          (optional) refresh scopes
 *   --json                                    print the full result as JSON
 *   --tenant         / GRAPH_TENANT          (optional) default "common"
 *
 * The access token is used for the Graph calls. client-id/client-secret are the
 * OAuth app credentials used to refresh that token via the refresh_token grant
 * when a --refresh-token is supplied (and when Graph returns 401 mid-stream).
 *
 * Run:
 *   npx ts-node samples/delta-accumulate.ts --access-token eyJ0... \
 *     --client-id <id> --client-secret <secret>
 *
 *   # or, after `npm run build`, against the compiled lib:
 *   node dist-sample/delta-accumulate.js ...
 */
import { runDeltaSync, TokenProvider, DeltaSyncResult } from '../src/delta-sync-lib';

// ── tiny arg parser ────────────────────────────────────────────────────────
function parseArgs(argv: string[]): Record<string, string | boolean> {
  const out: Record<string, string | boolean> = {};
  for (let i = 0; i < argv.length; i++) {
    const token = argv[i];
    if (!token.startsWith('--')) continue;
    const key = token.slice(2);
    const next = argv[i + 1];
    if (next === undefined || next.startsWith('--')) {
      out[key] = true; // boolean flag
    } else {
      out[key] = next;
      i++;
    }
  }
  return out;
}

function pick(
  args: Record<string, string | boolean>,
  flag: string,
  env: string,
): string | undefined {
  const v = args[flag];
  if (typeof v === 'string' && v.length > 0) return v;
  const e = process.env[env];
  return e && e.length > 0 ? e : undefined;
}

// ── refresh_token grant (only invoked when a refresh token is provided) ──────
async function refreshAccessToken(params: {
  tenant: string;
  clientId: string;
  clientSecret: string;
  refreshToken: string;
  scopes: string;
}): Promise<string> {
  const url = `https://login.microsoftonline.com/${params.tenant}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: params.clientId,
    client_secret: params.clientSecret,
    grant_type: 'refresh_token',
    refresh_token: params.refreshToken,
    scope: params.scopes,
  });

  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString(),
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Token refresh failed (${res.status}): ${text}`);
  }

  const json = (await res.json()) as { access_token?: string };
  if (!json.access_token) {
    throw new Error('Token refresh response did not contain an access_token');
  }
  return json.access_token;
}

/**
 * Build the pluggable token provider the library needs. Serves the supplied
 * access token; if it has been invalidated and a refresh token is available,
 * transparently mints a new one with the client credentials.
 */
function makeTokenProvider(cfg: {
  accessToken: string;
  clientId: string;
  clientSecret: string;
  refreshToken?: string;
  tenant: string;
  scopes: string;
}): TokenProvider & { invalidate(): void } {
  let current: string | null = cfg.accessToken;
  return {
    invalidate() {
      current = null;
    },
    async getAccessToken(): Promise<string> {
      if (current) return current;
      if (!cfg.refreshToken) {
        throw new Error(
          'Access token is invalid and no --refresh-token was provided to refresh it.',
        );
      }
      current = await refreshAccessToken({
        tenant: cfg.tenant,
        clientId: cfg.clientId,
        clientSecret: cfg.clientSecret,
        refreshToken: cfg.refreshToken,
        scopes: cfg.scopes,
      });
      return current;
    },
  };
}

function die(message: string): never {
  console.error(`error: ${message}`);
  console.error('run with --help-less; see the header comment for usage.');
  process.exit(1);
}

async function main(): Promise<void> {
  const args = parseArgs(process.argv.slice(2));

  const accessToken = pick(args, 'access-token', 'GRAPH_ACCESS_TOKEN');
  const clientId = pick(args, 'client-id', 'GRAPH_CLIENT_ID');
  const clientSecret = pick(args, 'client-secret', 'GRAPH_CLIENT_SECRET');

  if (!accessToken) die('missing --access-token (or GRAPH_ACCESS_TOKEN)');
  if (!clientId) die('missing --client-id (or GRAPH_CLIENT_ID)');
  if (!clientSecret) die('missing --client-secret (or GRAPH_CLIENT_SECRET)');

  const refreshToken = pick(args, 'refresh-token', 'GRAPH_REFRESH_TOKEN');
  const deltaLink = pick(args, 'delta-link', 'GRAPH_DELTA_LINK') ?? null;
  const userId = pick(args, 'user-id', 'GRAPH_USER_ID') ?? 'cli-user';
  const resourcePath = pick(args, 'resource-path', 'GRAPH_RESOURCE_PATH');
  const tenant = pick(args, 'tenant', 'GRAPH_TENANT') ?? 'common';
  const scopes =
    pick(args, 'scopes', 'GRAPH_SCOPES') ??
    'https://graph.microsoft.com/.default offline_access';
  const asJson = args.json === true;

  const tokenProvider = makeTokenProvider({
    accessToken,
    clientId,
    clientSecret,
    refreshToken,
    tenant,
    scopes,
  });

  const logger = {
    log: (m: string) => console.error(m), // progress → stderr, keep stdout clean
    warn: (m: string) => console.error(m),
    error: (m: string, e?: unknown) => console.error(m, e ?? ''),
  };

  let result: DeltaSyncResult;
  try {
    result = await runDeltaSync({
      userId,
      tokenProvider,
      deltaLink,
      resourcePath,
      logger,
    });
  } catch (err) {
    // One retry after a forced refresh if the initial token was rejected.
    if (refreshToken && isUnauthorized(err)) {
      logger.warn('[delta-accumulate] token rejected — refreshing and retrying once');
      tokenProvider.invalidate();
      result = await runDeltaSync({
        userId,
        tokenProvider,
        deltaLink,
        resourcePath,
        logger,
      });
    } else {
      throw err;
    }
  }

  // ── output on stdout ─────────────────────────────────────────────────────
  if (asJson) {
    process.stdout.write(JSON.stringify(result, null, 2) + '\n');
    return;
  }

  const { events, deletedIds, deltaLink: newCursor, stats } = result;
  console.log('');
  console.log('Accumulated delta result');
  console.log('────────────────────────');
  console.log(
    `pages=${stats.pages}  rawChanges=${stats.totalChanges}  ` +
      `creates=${stats.creates} updates=${stats.updates} ` +
      `deletes=${stats.deletes} recreates=${stats.recreates}`,
  );
  console.log(`final events: ${events.length}   deleted ids: ${deletedIds.length}`);
  console.log('');

  for (const e of events) {
    const when = e.lastModifiedDateTime ?? e.createdDateTime ?? '';
    console.log(`  ● ${e.id}  ${when}  ${e.subject ?? '(no subject)'}`);
  }
  for (const id of deletedIds) {
    console.log(`  ✖ ${id}  (deleted)`);
  }

  console.log('');
  console.log(`new deltaLink: ${newCursor ?? '(none)'}`);
  console.log('(persist the deltaLink and pass it as --delta-link next time)');
}

function isUnauthorized(err: unknown): boolean {
  if (!err || typeof err !== 'object') return false;
  const anyErr = err as { statusCode?: number; response?: { status?: number } };
  return anyErr.statusCode === 401 || anyErr.response?.status === 401;
}

main().catch((err: unknown) => {
  console.error('delta-accumulate failed:', err instanceof Error ? err.message : err);
  process.exit(1);
});

import { Client } from '@microsoft/microsoft-graph-client';
import { is410Error } from '../utils/retry.util';
import { accumulateDeltaChanges } from './accumulate';
import { createGraphClient } from './graph-client';
import { buildInitialUrl, streamDeltaPages } from './delta-stream';
import {
  AccumulatorState,
  DeltaLogger,
  DeltaSyncInput,
  DeltaSyncResult,
  DeltaSyncStats,
} from './types';

const DEFAULT_RESOURCE_PATH = '/me/events/delta';

/**
 * Request the full Microsoft Graph calendar delta stream for one user and fold
 * it into a single result: the accumulated final events, the net deletions,
 * and the new cursor to persist.
 *
 * This is the standalone, framework-free counterpart to
 * `reconciliation.service.ts#fetchRemoteMetadata`. It is:
 *
 * - **auth-agnostic** — obtains tokens via {@link DeltaSyncInput.tokenProvider};
 * - **stateless** — takes an optional prior `deltaLink` and returns the new one;
 *   it never reads or writes any store;
 * - **accumulation only** — it does not compare against local state or detect
 *   drift.
 *
 * On a 410 (expired cursor) while using an existing `deltaLink`, it transparently
 * restarts a full sync from a fresh baseline and returns that result.
 */
export async function runDeltaSync(input: DeltaSyncInput): Promise<DeltaSyncResult> {
  const { userId, tokenProvider, logger } = input;
  const resourcePath = input.resourcePath ?? DEFAULT_RESOURCE_PATH;

  const client = createGraphClient(tokenProvider, userId);

  const initialUrl = buildInitialUrl(resourcePath, input.dateRange);
  const hasCursor = typeof input.deltaLink === 'string' && input.deltaLink.length > 0;
  const startUrl = hasCursor ? (input.deltaLink as string) : initialUrl;

  logger?.log?.(
    `[runDeltaSync] user=${userId} mode=${hasCursor ? 'incremental' : 'initialize'}`,
  );

  try {
    return await accumulateStream(client, startUrl, logger);
  } catch (error) {
    if (hasCursor && is410Error(error)) {
      logger?.warn?.(
        '[runDeltaSync] Delta cursor expired (410) — restarting full sync from a fresh baseline',
      );
      return await accumulateStream(client, initialUrl, logger);
    }
    throw error;
  }
}

/**
 * Drive {@link streamDeltaPages}, folding every page through
 * {@link accumulateDeltaChanges}, and materialize the terminal result.
 */
async function accumulateStream(
  client: Client,
  startUrl: string,
  logger?: DeltaLogger,
): Promise<DeltaSyncResult> {
  let acc: AccumulatorState = {
    final: new Map(),
    deleted: new Set(),
  };

  const stats: DeltaSyncStats = {
    pages: 0,
    totalChanges: 0,
    creates: 0,
    updates: 0,
    deletes: 0,
    recreates: 0,
  };

  const generator = streamDeltaPages(client, startUrl, { logger });
  let result = await generator.next();

  while (!result.done) {
    const batch = result.value;
    stats.pages++;
    stats.totalChanges += batch.length;

    const out = accumulateDeltaChanges(batch, acc, logger);
    acc = { final: out.final, deleted: out.deleted };
    stats.creates += out.counts.creates;
    stats.updates += out.counts.updates;
    stats.deletes += out.counts.deletes;
    stats.recreates += out.counts.recreates;

    result = await generator.next();
  }

  const deltaLink = result.value;

  logger?.log?.(
    `[runDeltaSync] done: ${acc.final.size} final events, ${acc.deleted.size} deletions ` +
      `across ${stats.pages} pages (${stats.totalChanges} raw changes)`,
  );

  return {
    events: Array.from(acc.final.values()),
    deletedIds: Array.from(acc.deleted),
    deltaLink,
    stats,
  };
}

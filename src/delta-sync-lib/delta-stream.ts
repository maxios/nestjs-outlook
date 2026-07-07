import { Client } from '@microsoft/microsoft-graph-client';
import { delay } from '../utils/retry.util';
import { executeGraphApiCall } from './graph-executor';
import { DeltaEvent, DeltaItem, DeltaLogger, DeltaResponse } from './types';

const DEFAULT_INTER_PAGE_DELAY_MS = 200;

/**
 * Compose the initial delta request URL, optionally constrained by a date
 * window. Pure — no I/O. Ported from `delta-sync.service.ts#buildInitialUrl`.
 */
export function buildInitialUrl(
  resourcePath: string,
  dateRange?: { startDate: Date; endDate: Date },
): string {
  if (!dateRange) {
    return resourcePath;
  }
  const { startDate, endDate } = dateRange;
  return `${resourcePath}?startDateTime=${startDate.toISOString()}&endDateTime=${endDate.toISOString()}`;
}

/**
 * Sort a batch by modification/creation time (oldest → newest) so that the
 * accumulator observes changes in causal order within a page. Ported from
 * `delta-sync.service.ts#sortDeltaItems`.
 */
export function sortDeltaItems<T extends DeltaItem>(items: T[]): T[] {
  return [...items].sort((a, b) => {
    const aTime = a.lastModifiedDateTime ?? a.createdDateTime ?? '';
    const bTime = b.lastModifiedDateTime ?? b.createdDateTime ?? '';
    return new Date(aTime).getTime() - new Date(bTime).getTime();
  });
}

export interface StreamOptions {
  logger?: DeltaLogger;
  /** Delay between page requests (default: 200ms). Set 0 to disable. */
  interPageDelayMs?: number;
}

/**
 * Async generator over a delta query: follows `@odata.nextLink` page by page,
 * yielding each page's items (sorted), and returns the terminal
 * `@odata.deltaLink` cursor once the last page is reached.
 *
 * Ported from `delta-sync.service.ts#fetchDeltaPagesCore` + `streamSortedPages`,
 * minus the rate limiter and event emitter. Each request carries
 * `Prefer: IdType="ImmutableId"` so ids survive mailbox migrations.
 *
 * Throws on a non-recoverable error (including 410 Gone — the caller decides
 * whether to restart from a fresh baseline; see run-delta-sync.ts).
 */
export async function* streamDeltaPages(
  client: Client,
  startUrl: string,
  opts: StreamOptions = {},
): AsyncGenerator<DeltaEvent[], string | null, unknown> {
  const { logger } = opts;
  const interPageDelayMs = opts.interPageDelayMs ?? DEFAULT_INTER_PAGE_DELAY_MS;

  let currentUrl = startUrl;
  let pageCount = 0;
  let finalDeltaLink: string | null = null;

  while (currentUrl) {
    pageCount++;

    const response = (await executeGraphApiCall(
      () => client.api(currentUrl).header('Prefer', 'IdType="ImmutableId"').get(),
      { logger, resourceName: `deltaPage-${pageCount}` },
    )) as DeltaResponse<DeltaEvent>;

    const sorted = sortDeltaItems(response.value);
    logger?.debug?.(`[streamDeltaPages] Page ${pageCount}: ${sorted.length} items`);

    // The delta link only appears on the final page.
    if (response['@odata.deltaLink']) {
      finalDeltaLink = response['@odata.deltaLink'];
    }

    yield sorted;

    currentUrl = response['@odata.nextLink'] || '';
    if (currentUrl && interPageDelayMs > 0) {
      await delay(interPageDelayMs);
    }
  }

  return finalDeltaLink;
}

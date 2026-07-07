import { AccumulatorState, BatchCounts, DeltaEvent, DeltaLogger } from './types';

/**
 * Fold one batch of delta changes into the running accumulator.
 *
 * Ported from `reconciliation.service.ts#accumulateDeltaChanges` (the NestJS
 * logger swapped for the optional {@link DeltaLogger}). Same semantics:
 *
 * - Deletions (`@removed` or `isCancelled`) remove the id from `final` and add
 *   it to `deleted`.
 * - Creates/updates set the id in `final` (last-write-wins across batches).
 * - A create/update for an id that was `deleted` earlier in the stream is a
 *   "recreate": the id is pulled back out of `deleted` and restored to `final`.
 *
 * Pure w.r.t. inputs other than the (optionally mutated) `base` accumulator:
 * when `base` is supplied it is mutated in place and returned; otherwise a
 * fresh accumulator is created.
 */
export function accumulateDeltaChanges(
  changes: DeltaEvent[],
  base?: AccumulatorState,
  logger?: DeltaLogger,
): AccumulatorState & { counts: BatchCounts } {
  const accumulator: AccumulatorState = base ?? {
    final: new Map<string, DeltaEvent>(),
    deleted: new Set<string>(),
  };

  const counts: BatchCounts = { creates: 0, updates: 0, deletes: 0, recreates: 0 };

  for (const change of changes) {
    const eventId = change.id;

    if (!eventId) {
      logger?.warn?.('[accumulateDeltaChanges] Delta change missing id, skipping');
      continue;
    }

    // Handle deletions (@removed or isCancelled)
    if (change['@removed'] || change.isCancelled) {
      const wasPresent = accumulator.final.delete(eventId);
      accumulator.deleted.add(eventId);
      counts.deletes++;
      logger?.debug?.(`[accumulateDeltaChanges] Deleted event ${eventId} (wasPresent=${wasPresent})`);
      continue;
    }

    // Handle create/update
    const wasDeleted = accumulator.deleted.has(eventId);
    if (wasDeleted) {
      // Event was deleted earlier in stream, now recreated
      accumulator.deleted.delete(eventId);
      counts.recreates++;
      logger?.debug?.(`[accumulateDeltaChanges] Recreated event ${eventId} (was deleted earlier)`);
    }

    const wasPresent = accumulator.final.has(eventId);
    accumulator.final.set(eventId, change);

    if (wasPresent) {
      counts.updates++;
    } else {
      counts.creates++;
    }
  }

  logger?.debug?.(
    `[accumulateDeltaChanges] Batch processed: creates=${counts.creates}, ` +
      `updates=${counts.updates}, deletes=${counts.deletes}, recreates=${counts.recreates}`,
  );

  return { ...accumulator, counts };
}

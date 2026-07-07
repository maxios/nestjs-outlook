/**
 * Public contracts for the standalone delta-accumulation library.
 *
 * This library is framework-free (no NestJS, no TypeORM). It reproduces the
 * "request delta + accumulate the whole stream" half of the reconciliation
 * flow, returning the folded final state without comparing against any local
 * store. See ./run-delta-sync.ts for the entry point.
 */
import { Event, Message } from '../types';

/**
 * A single item in a Microsoft Graph delta response.
 *
 * Delta queries return FULL entities, so a `DeltaEvent` is a Graph `Event`
 * plus the delta bookkeeping fields below. Deletions arrive as an item that
 * carries `@removed` (and/or `isCancelled`) instead of the full body.
 */
export interface DeltaItem {
  id?: string;
  lastModifiedDateTime?: string;
  createdDateTime?: string;
  changeKey?: string;
  isCancelled?: boolean;
  '@removed'?: {
    reason: 'changed' | 'deleted';
  };
}

export type DeltaEvent = Event & DeltaItem;
export type DeltaMessage = Message & DeltaItem;

/** Envelope of a Microsoft Graph delta page. */
export interface DeltaResponse<T> {
  '@odata.nextLink'?: string;
  '@odata.deltaLink'?: string;
  value: T[];
}

/**
 * Pluggable access-token source. The library is auth-agnostic: it only needs
 * a way to obtain a valid Bearer token for the given opaque user id. The
 * caller owns OAuth (refresh grants, on-behalf-of, caching, etc.).
 */
export interface TokenProvider {
  getAccessToken(userId: string): Promise<string>;
}

/** Optional logger. Every method is optional; omitted ones are no-ops. */
export interface DeltaLogger {
  log?(message: string): void;
  debug?(message: string): void;
  warn?(message: string): void;
  error?(message: string, error?: unknown): void;
}

/** Input to {@link runDeltaSync}. */
export interface DeltaSyncInput {
  /** Opaque user identifier passed straight to the token provider. */
  userId: string;
  /** Source of a valid Graph access token for this user. */
  tokenProvider: TokenProvider;
  /**
   * Prior delta cursor. When present, an incremental sync runs from it.
   * When absent/empty, the library initializes a fresh delta baseline.
   * The library is stateless — it never reads or writes this anywhere;
   * the caller persists {@link DeltaSyncResult.deltaLink} for next time.
   */
  deltaLink?: string | null;
  /**
   * Optional window applied ONLY when initializing (no `deltaLink`). Ignored
   * for incremental syncs, which resume from the cursor's own watermark.
   */
  dateRange?: { startDate: Date; endDate: Date };
  /** Graph delta resource. Defaults to `/me/events/delta`. */
  resourcePath?: string;
  /** Optional logger. */
  logger?: DeltaLogger;
}

/** Per-run counters describing what the accumulation observed. */
export interface DeltaSyncStats {
  pages: number;
  totalChanges: number;
  creates: number;
  updates: number;
  deletes: number;
  recreates: number;
}

/** Result of {@link runDeltaSync} — the folded final state of the stream. */
export interface DeltaSyncResult {
  /** Accumulated final state: full event objects, last-write-wins per id. */
  events: DeltaEvent[];
  /** Ids explicitly removed across the stream (net of any recreate). */
  deletedIds: string[];
  /** New cursor to persist and feed back on the next call (null if none). */
  deltaLink: string | null;
  stats: DeltaSyncStats;
}

/** In-flight accumulator state threaded across delta batches. */
export interface AccumulatorState {
  final: Map<string, DeltaEvent>;
  deleted: Set<string>;
}

/** Counts produced by folding a single batch. */
export interface BatchCounts {
  creates: number;
  updates: number;
  deletes: number;
  recreates: number;
}

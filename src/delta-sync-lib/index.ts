/**
 * Standalone delta-accumulation library.
 *
 * Framework-free (no NestJS/TypeORM). Requests the Microsoft Graph calendar
 * delta stream for a user and folds the whole stream into a final result:
 * `{ events, deletedIds, deltaLink, stats }`. Auth is pluggable
 * ({@link TokenProvider}) and the delta cursor is stateless (in/out).
 *
 * Primary entry point: {@link runDeltaSync}.
 *
 * @example
 * ```ts
 * import { runDeltaSync } from '@checkfirst-ltd/microsoft-outlook'; // or './delta-sync-lib'
 *
 * const tokenProvider = { getAccessToken: async () => process.env.GRAPH_ACCESS_TOKEN! };
 * let cursor: string | null = null;
 *
 * // First run: initialize a baseline.
 * const first = await runDeltaSync({ userId: 'user-123', tokenProvider });
 * cursor = first.deltaLink;
 *
 * // Later: incremental — feed the previous cursor back in.
 * const next = await runDeltaSync({ userId: 'user-123', tokenProvider, deltaLink: cursor });
 * cursor = next.deltaLink;
 * ```
 */
export { runDeltaSync } from './run-delta-sync';
export { accumulateDeltaChanges } from './accumulate';
export { createGraphClient } from './graph-client';
export { executeGraphApiCall } from './graph-executor';
export { buildInitialUrl, sortDeltaItems, streamDeltaPages } from './delta-stream';
export * from './types';

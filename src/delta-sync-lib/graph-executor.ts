import {
  delay,
  extractRetryAfterSeconds,
  is404Error,
  is429Error,
  is503Error,
  retryWithBackoff,
} from '../utils/retry.util';
import { DeltaLogger } from './types';

/** Options for {@link executeGraphApiCall}. */
export interface GraphExecOptions {
  /** Maximum retry attempts (default: 3). */
  maxRetries?: number;
  /** Base delay in ms for exponential backoff (default: 1000). */
  retryDelayMs?: number;
  /** Optional logger. */
  logger?: DeltaLogger;
  /** Label used in log lines. */
  resourceName?: string;
  /** When true, a 404 resolves to `null` instead of throwing (default: false). */
  return404AsNull?: boolean;
}

/**
 * Execute a Microsoft Graph call with retry + rate-limit-aware backoff.
 *
 * Slimmed port of `outlook-api-executor.util.ts#executeGraphApiCall` with the
 * NestJS `GraphRateLimiterService` dependency removed (the delta stream paces
 * itself with an inter-page delay instead). Retains the important behavior:
 *
 * - Honors the `Retry-After` header on 429 (rate limit) and 503 responses.
 * - Exponential backoff for other transient/5xx/network errors.
 * - Does NOT retry permanent client errors (401/403/404/410) — see
 *   `isNonRetryableError` in retry.util.
 */
export async function executeGraphApiCall<T>(
  operation: () => Promise<T>,
  options: GraphExecOptions = {},
): Promise<T | null> {
  const {
    maxRetries = 3,
    retryDelayMs = 1000,
    logger,
    resourceName = 'resource',
    return404AsNull = false,
  } = options;

  try {
    return await retryWithBackoff(
      async () => {
        try {
          return await operation();
        } catch (error) {
          // 429: prefer the server's Retry-After over blind backoff.
          if (is429Error(error)) {
            const retryAfterSeconds = extractRetryAfterSeconds(error);
            if (retryAfterSeconds !== null) {
              logger?.warn?.(
                `Rate limited on ${resourceName}, waiting ${retryAfterSeconds}s as per Retry-After header`,
              );
              await delay(retryAfterSeconds * 1000);
            }
          }

          // 503: also honor Retry-After when present.
          if (is503Error(error)) {
            const retryAfterSeconds = extractRetryAfterSeconds(error);
            if (retryAfterSeconds !== null) {
              logger?.warn?.(
                `Service unavailable (503) on ${resourceName}, waiting ${retryAfterSeconds}s as per Retry-After header`,
              );
              await delay(retryAfterSeconds * 1000);
            }
          }

          throw error;
        }
      },
      {
        maxRetries,
        retryDelayMs,
        logger: logger?.warn
          ? { warn: (message: string) => logger.warn?.(message) }
          : undefined,
        operationName: resourceName,
      },
    );
  } catch (error) {
    if (is404Error(error)) {
      if (return404AsNull) {
        logger?.warn?.(`Resource not found (likely deleted): ${resourceName}`);
        return null;
      }
      logger?.error?.(`Resource not found: ${resourceName}`, error);
    }
    throw error;
  }
}

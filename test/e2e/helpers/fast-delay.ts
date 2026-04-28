import * as retryUtil from '../../../src/utils/retry.util';

/**
 * Replace the library's `delay` utility with a no-op for the duration of a spec
 * file. Lets retry-loop tests run in milliseconds rather than the ~127s of
 * exponential backoff that `executeGraphApiCall(maxRetries: 7)` would normally
 * require.
 *
 * Use in `beforeAll` and pair with `restoreDelay` in `afterAll`.
 */
let original: typeof retryUtil.delay | null = null;

export function shortCircuitDelay(): void {
  if (original) return;
  original = retryUtil.delay;
  (retryUtil as { delay: typeof retryUtil.delay }).delay = async () => undefined;
}

export function restoreDelay(): void {
  if (original) {
    (retryUtil as { delay: typeof retryUtil.delay }).delay = original;
    original = null;
  }
}

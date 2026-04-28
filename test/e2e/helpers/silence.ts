/**
 * Silences `console.log` noise emitted by the library's constructors / debug paths
 * during e2e runs. Errors and warnings are kept so genuine problems still surface.
 *
 * Import for side effects in a spec file's top-level scope, or call `silenceLibraryLogs()`
 * inside a `beforeAll`.
 */
let originalLog: typeof console.log | null = null;

export function silenceLibraryLogs(): void {
  if (originalLog) return;
  originalLog = console.log;
  console.log = (): void => undefined;
}

export function restoreLibraryLogs(): void {
  if (originalLog) {
    console.log = originalLog;
    originalLog = null;
  }
}

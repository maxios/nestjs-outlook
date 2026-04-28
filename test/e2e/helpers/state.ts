import { E2ETestHarness } from './test-module';
import { PermissionScope } from '../../../src/enums/permission-scope.enum';

export interface BuiltState {
  state: string;
  csrf: string;
  timestamp: number;
}

/**
 * Encode a state object the same way `getLoginUrl` does (base64, padding stripped).
 */
export function encodeState(stateObj: object): string {
  return Buffer.from(JSON.stringify(stateObj)).toString('base64').replace(/=/g, '');
}

/**
 * Build a valid OAuth state string and seed a matching CSRF row in the DB.
 *
 * The result is suitable for `exchangeCodeForToken(code, state)` and
 * `GET /auth/microsoft/callback?code=...&state=...`.
 */
export async function buildValidState(
  harness: E2ETestHarness,
  externalUserId: string,
  requestedScopes: PermissionScope[] = [
    PermissionScope.CALENDAR_READ,
    PermissionScope.CALENDAR_WRITE,
    PermissionScope.EMAIL_SEND,
    PermissionScope.EMAIL_READ,
    PermissionScope.EMAIL_WRITE,
  ],
): Promise<BuiltState> {
  const csrf = `csrf-${Math.random().toString(16).slice(2)}${Math.random().toString(16).slice(2)}`.padEnd(64, '0').slice(0, 64);
  const timestamp = Date.now();

  await harness.csrfRepo.saveToken(csrf, externalUserId, 30 * 60 * 1000);

  const stateObj = {
    userId: externalUserId,
    csrf,
    timestamp,
    requestedScopes,
  };

  return { state: encodeState(stateObj), csrf, timestamp };
}

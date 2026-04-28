import { E2ETestHarness } from './test-module';
import { MicrosoftUser } from '../../../src/entities/microsoft-user.entity';
import { MicrosoftUserStatus } from '../../../src/enums/microsoft-user-status.enum';
import { OutlookWebhookSubscription } from '../../../src/entities/outlook-webhook-subscription.entity';

export async function seedMicrosoftUser(
  harness: E2ETestHarness,
  overrides: Partial<MicrosoftUser> = {},
): Promise<MicrosoftUser> {
  const user = new MicrosoftUser();
  user.externalUserId = overrides.externalUserId ?? 'ext-1';
  user.accessToken = overrides.accessToken ?? 'seeded-access-token';
  user.refreshToken = overrides.refreshToken ?? 'seeded-refresh-token';
  user.tokenExpiry = overrides.tokenExpiry ?? new Date(Date.now() + 60 * 60 * 1000);
  user.scopes = overrides.scopes ?? 'offline_access User.Read Calendars.Read Calendars.ReadWrite Mail.Send Mail.Read Mail.ReadWrite';
  user.isActive = overrides.isActive ?? true;
  user.status = overrides.status ?? MicrosoftUserStatus.ACTIVE;
  user.defaultCalendarId = overrides.defaultCalendarId ?? null;
  return harness.microsoftUserRepo.save(user);
}

export async function seedSubscription(
  harness: E2ETestHarness,
  overrides: Partial<OutlookWebhookSubscription> & { userId: number },
): Promise<OutlookWebhookSubscription> {
  return harness.webhookSubRepo.saveSubscription({
    subscriptionId: overrides.subscriptionId ?? `sub-${Math.random().toString(16).slice(2, 8)}`,
    userId: overrides.userId,
    resource: overrides.resource ?? '/me/events',
    changeType: overrides.changeType ?? 'created,updated,deleted',
    clientState: overrides.clientState ?? `user_${overrides.userId}_state`,
    notificationUrl: overrides.notificationUrl ?? 'https://app.test/api/calendar/webhook',
    expirationDateTime: overrides.expirationDateTime ?? new Date(Date.now() + 72 * 3600 * 1000),
    isActive: overrides.isActive ?? true,
    lastNotificationAt: overrides.lastNotificationAt ?? null,
  });
}

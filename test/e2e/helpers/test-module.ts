import { Test, TestingModule } from '@nestjs/testing';
import { INestApplication } from '@nestjs/common';
import { TypeOrmModule } from '@nestjs/typeorm';
import { EventEmitter2 } from '@nestjs/event-emitter';
import { DataSource, Repository } from 'typeorm';

import { MicrosoftOutlookModule } from '../../../src/microsoft-outlook.module';
import { MicrosoftOutlookConfig } from '../../../src/interfaces/config/outlook-config.interface';
import { MicrosoftAuthService } from '../../../src/services/auth/microsoft-auth.service';
import { MicrosoftSubscriptionService } from '../../../src/services/subscription/microsoft-subscription.service';
import { EmailService } from '../../../src/services/email/email.service';
import { CalendarService } from '../../../src/services/calendar/calendar.service';
import { GraphRateLimiterService } from '../../../src/services/shared/graph-rate-limiter.service';
import { MicrosoftCsrfTokenRepository } from '../../../src/repositories/microsoft-csrf-token.repository';
import { OutlookWebhookSubscriptionRepository } from '../../../src/repositories/outlook-webhook-subscription.repository';
import { MicrosoftUser } from '../../../src/entities/microsoft-user.entity';
import { MicrosoftCsrfToken } from '../../../src/entities/csrf-token.entity';
import { OutlookWebhookSubscription } from '../../../src/entities/outlook-webhook-subscription.entity';
import { OutlookDeltaLink } from '../../../src/entities/delta-link.entity';

export const DEFAULT_TEST_CONFIG: MicrosoftOutlookConfig = {
  clientId: 'test-client',
  clientSecret: 'test-secret',
  redirectPath: '/auth/microsoft/callback',
  backendBaseUrl: 'https://app.test',
  basePath: 'api',
  calendarWebhookPath: '/calendar/webhook',
};

export interface CapturedEvent {
  name: string;
  args: unknown[];
}

export interface E2ETestHarness {
  module: TestingModule;
  app: INestApplication;
  dataSource: DataSource;
  config: MicrosoftOutlookConfig;
  authService: MicrosoftAuthService;
  subscriptionService: MicrosoftSubscriptionService;
  emailService: EmailService;
  calendarService: CalendarService;
  rateLimiter: GraphRateLimiterService;
  csrfRepo: MicrosoftCsrfTokenRepository;
  webhookSubRepo: OutlookWebhookSubscriptionRepository;
  microsoftUserRepo: Repository<MicrosoftUser>;
  events: CapturedEvent[];
  close: () => Promise<void>;
}

export async function createTestApp(
  configOverride: Partial<MicrosoftOutlookConfig> = {},
): Promise<E2ETestHarness> {
  const config: MicrosoftOutlookConfig = { ...DEFAULT_TEST_CONFIG, ...configOverride };

  const module: TestingModule = await Test.createTestingModule({
    imports: [
      TypeOrmModule.forRoot({
        type: 'better-sqlite3',
        database: ':memory:',
        dropSchema: true,
        synchronize: true,
        entities: [MicrosoftUser, MicrosoftCsrfToken, OutlookWebhookSubscription, OutlookDeltaLink],
      }),
      MicrosoftOutlookModule.forRoot(config),
    ],
  }).compile();

  const app = module.createNestApplication({ logger: false });
  await app.init();

  const dataSource = module.get(DataSource);
  const eventEmitter = module.get(EventEmitter2);

  const events: CapturedEvent[] = [];
  eventEmitter.onAny((name: string | string[], ...args: unknown[]) => {
    events.push({ name: Array.isArray(name) ? name.join('.') : name, args });
  });

  const harness: E2ETestHarness = {
    module,
    app,
    dataSource,
    config,
    authService: module.get(MicrosoftAuthService),
    subscriptionService: module.get(MicrosoftSubscriptionService),
    emailService: module.get(EmailService),
    calendarService: module.get(CalendarService),
    rateLimiter: module.get(GraphRateLimiterService),
    csrfRepo: module.get(MicrosoftCsrfTokenRepository),
    webhookSubRepo: module.get(OutlookWebhookSubscriptionRepository),
    microsoftUserRepo: dataSource.getRepository(MicrosoftUser),
    events,
    close: async () => {
      await app.close();
    },
  };

  return harness;
}

/**
 * Microsoft Graph API subscription structure
 */
export interface MicrosoftSubscription {
  id: string;
  resource: string;
  changeType?: string;
  clientState?: string;
  notificationUrl?: string;
  expirationDateTime?: string;
  creatorId?: string;
}

/**
 * Filter function for subscriptions
 */
export type SubscriptionFilter = (subscription: MicrosoftSubscription) => boolean;

/**
 * Options for subscription cleanup
 */
export interface SubscriptionCleanupOptions {
  accessToken: string;
  filter?: SubscriptionFilter;
}

/**
 * Result of subscription cleanup operation
 */
export interface SubscriptionCleanupResult {
  totalFound: number;
  successfullyDeleted: number;
  failedToDelete: number;
  deletedSubscriptionIds: string[];
  errors: Array<{ subscriptionId: string; error: string }>;
}

/**
 * Options for creating a new subscription
 */
export interface CreateSubscriptionOptions {
  accessToken: string;
  resource: string;
  notificationUrl: string;
  userId: number;
  expirationDateTime: Date;
  correlationId?: string;
}
---
dep:
  type: reference
  audience: [consumer-dev, contributor]
  owner: "@checkfirst-ltd"
  created: 2026-04-03T12:00:00+03:00
  last_verified: 2026-04-03T12:00:00+03:00
  confidence: medium
  depends_on:
    - ../../src/interfaces/recurrence.interfaces.ts
    - ../../src/services/calendar/recurrence.service.ts
  tags: [recurrence, calendar, types]
  links:
    - target: ./calendar-service.md
      rel: REQUIRES
---

# Recurrence Types Reference

Interfaces for processing Microsoft Outlook recurring calendar events.

Source: `src/interfaces/recurrence.interfaces.ts`

---

## `OutlookEventType`

Union type classifying calendar events.

| Value | Description |
|-------|-------------|
| `'singleInstance'` | A non-recurring event |
| `'seriesMaster'` | The root event defining a recurring series |
| `'occurrence'` | A single instance within a recurring series |
| `'exception'` | A modified instance that deviates from the series pattern |

---

## `RecurrenceRule`

Mirrors Microsoft Graph `PatternedRecurrence`. Stored as JSON on series master rows.

### `pattern`

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `type` | `'daily' \| 'weekly' \| 'absoluteMonthly' \| 'relativeMonthly' \| 'absoluteYearly' \| 'relativeYearly'` | Yes | Recurrence pattern type |
| `interval` | `number` | Yes | Frequency interval (e.g., every 2 weeks) |
| `daysOfWeek` | `string[]` | No | Days of the week (for weekly patterns) |
| `dayOfMonth` | `number` | No | Day of the month (for monthly/yearly patterns) |
| `month` | `number` | No | Month (for yearly patterns) |
| `firstDayOfWeek` | `string` | No | First day of the week for the pattern |
| `index` | `string` | No | Ordinal position: `'first'`, `'second'`, `'third'`, `'fourth'`, `'last'` |

### `range`

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `type` | `'endDate' \| 'noEnd' \| 'numbered'` | Yes | How the recurrence ends |
| `startDate` | `string` | Yes | Series start date |
| `endDate` | `string` | No | Series end date (when `type` is `'endDate'`) |
| `numberOfOccurrences` | `number` | No | Total occurrences (when `type` is `'numbered'`) |
| `recurrenceTimeZone` | `string` | No | Time zone for the recurrence |

---

## `ProcessedOutlookEvent`

Enriched event produced by `RecurrenceService.processEvent()`. Ready for calendar-hub consumption.

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `externalId` | `string` | Yes | Microsoft Graph event ID |
| `eventType` | `OutlookEventType` | Yes | Event classification |
| `start` | `{ dateTime: string; timeZone: string }` | Yes | Event start |
| `end` | `{ dateTime: string; timeZone: string }` | Yes | Event end |
| `subject` | `string` | Yes | Event title |
| `bodyPreview` | `string` | Yes | Truncated body text |
| `location` | `string` | No | Event location |
| `showAs` | `string` | No | Free/busy status |
| `changeKey` | `string` | No | Change tracking key |
| `seriesMasterId` | `string` | No | Parent series master ID (for occurrences/exceptions) |
| `transactionId` | `string` | No | Maps to `iCalUId` |
| `recurrenceRule` | `RecurrenceRule` | No | Only set on `seriesMaster` events |
| `originalStart` | `{ dateTime: string; timeZone: string }` | No | Only set on `exception` events |

---

## `ExpansionWindow`

| Field | Type | Description |
|-------|------|-------------|
| `startDate` | `Date` | Start of the expansion date range |
| `endDate` | `Date` | End of the expansion date range |

---

## `ExpandRecurringSeriesOptions`

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `existingExternalIds` | `string[]` | No | Currently stored occurrence IDs for stale detection |

---

## `RecurringEventExpansionResult`

Returned by `RecurrenceService.expandRecurringSeries()`.

| Field | Type | Description |
|-------|------|-------------|
| `seriesMaster` | `ProcessedOutlookEvent` | The series master event |
| `instances` | `ProcessedOutlookEvent[]` | Expanded occurrence/exception instances |
| `expansionWindow` | `ExpansionWindow` | Date range used for expansion |
| `staleExternalIds` | `string[]` | IDs that existed before but were not returned |

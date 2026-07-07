import { accumulateDeltaChanges } from './accumulate';
import { AccumulatorState, DeltaEvent } from './types';

/** Minimal delta event helper. */
function ev(id: string, extra: Partial<DeltaEvent> = {}): DeltaEvent {
  return { id, subject: `event-${id}`, ...extra } as DeltaEvent;
}

/** Removal marker helper. */
function removed(id: string): DeltaEvent {
  return { id, '@removed': { reason: 'deleted' } } as DeltaEvent;
}

describe('accumulateDeltaChanges', () => {
  it('creates then updates keep the latest version (last-write-wins)', () => {
    const acc = accumulateDeltaChanges([
      ev('a', { subject: 'v1' }),
      ev('a', { subject: 'v2' }),
    ]);

    expect(acc.final.size).toBe(1);
    expect((acc.final.get('a') as DeltaEvent).subject).toBe('v2');
    expect(acc.deleted.size).toBe(0);
    expect(acc.counts).toMatchObject({ creates: 1, updates: 1, deletes: 0, recreates: 0 });
  });

  it('@removed deletes an event and records the id', () => {
    const acc = accumulateDeltaChanges([ev('a'), removed('a')]);

    expect(acc.final.has('a')).toBe(false);
    expect(acc.deleted.has('a')).toBe(true);
    expect(acc.counts).toMatchObject({ creates: 1, deletes: 1 });
  });

  it('treats isCancelled as a deletion', () => {
    const acc = accumulateDeltaChanges([
      ev('a'),
      ev('a', { isCancelled: true }),
    ]);

    expect(acc.final.has('a')).toBe(false);
    expect(acc.deleted.has('a')).toBe(true);
    expect(acc.counts.deletes).toBe(1);
  });

  it('recreate: create -> delete -> create ends in final, not deleted', () => {
    const acc = accumulateDeltaChanges([
      ev('a', { subject: 'v1' }),
      removed('a'),
      ev('a', { subject: 'v3' }),
    ]);

    expect(acc.final.has('a')).toBe(true);
    expect((acc.final.get('a') as DeltaEvent).subject).toBe('v3');
    expect(acc.deleted.has('a')).toBe(false);
    expect(acc.counts.recreates).toBe(1);
  });

  it('accumulates across multiple batches via the base accumulator', () => {
    const base: AccumulatorState = { final: new Map(), deleted: new Set() };

    const first = accumulateDeltaChanges([ev('a'), ev('b')], base);
    const second = accumulateDeltaChanges(
      [removed('a'), ev('c')],
      { final: first.final, deleted: first.deleted },
    );

    expect([...second.final.keys()].sort()).toEqual(['b', 'c']);
    expect(second.deleted.has('a')).toBe(true);
  });

  it('skips changes without an id', () => {
    const acc = accumulateDeltaChanges([{ subject: 'orphan' } as DeltaEvent, ev('a')]);

    expect(acc.final.size).toBe(1);
    expect(acc.final.has('a')).toBe(true);
  });
});

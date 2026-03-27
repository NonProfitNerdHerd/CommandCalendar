/**
 * Outlook-style day/week time grid: fixed hours, vertical duration, overlap columns.
 */

export const CALENDAR_GRID_START_HOUR = 6;
export const CALENDAR_GRID_END_HOUR = 20;
/** Pixel height per hour in the scrollable grid (readability over compactness). */
export const CALENDAR_PX_PER_HOUR = 54;

export function setDayHour(d: Date, hour: number, min: number = 0): Date {
  const x = new Date(d.getFullYear(), d.getMonth(), d.getDate(), hour, min, 0, 0);
  return x;
}

/** Clip [start,end] to this calendar day's visible grid window; null if no intersection. */
export function clipIntervalToDayGrid(
  day: Date,
  rangeStart: Date,
  rangeEndInclusive: Date
): { startMs: number; endMs: number } | null {
  const winStart = setDayHour(day, CALENDAR_GRID_START_HOUR, 0);
  const winEnd = setDayHour(day, CALENDAR_GRID_END_HOUR, 0);
  const s = Math.max(rangeStart.getTime(), winStart.getTime());
  const e = Math.min(rangeEndInclusive.getTime(), winEnd.getTime());
  if (e - s < 60 * 1000) {
    return null;
  }
  return { startMs: s, endMs: e };
}

/** Minutes from grid start (6:00 AM) for a timestamp on the same local day. */
export function minutesFromGridStart(day: Date, ms: number): number {
  const grid0 = setDayHour(day, CALENDAR_GRID_START_HOUR, 0);
  return (ms - grid0.getTime()) / 60000;
}

export const CALENDAR_GRID_TOTAL_MINUTES: number =
  (CALENDAR_GRID_END_HOUR - CALENDAR_GRID_START_HOUR) * 60;

/**
 * Greedy lane assignment for overlapping intervals, then per-event laneCount among simultaneous overlaps only.
 */
export function assignTimeOverlapLanes<T extends { startMs: number; endMs: number }>(
  items: T[]
): (T & { lane: number; laneCount: number })[] {
  if (items.length === 0) {
    return [];
  }
  const sorted: T[] = [...items].sort(
    (a: T, b: T) => a.startMs - b.startMs || a.endMs - b.endMs
  );
  const laneEndMs: number[] = [];
  const laneByIndex: number[] = [];

  for (const it of sorted) {
    let lane: number = -1;
    for (let i: number = 0; i < laneEndMs.length; i++) {
      if (it.startMs >= laneEndMs[i]) {
        lane = i;
        laneEndMs[i] = it.endMs;
        break;
      }
    }
    if (lane < 0) {
      lane = laneEndMs.length;
      laneEndMs.push(it.endMs);
    }
    laneByIndex.push(lane);
  }

  const withLane: (T & { lane: number })[] = sorted.map((it: T, i: number) => ({
    ...it,
    lane: laneByIndex[i]
  }));

  return withLane.map((it: T & { lane: number }) => {
    const concurrent: (T & { lane: number })[] = withLane.filter(
      (o: T & { lane: number }) => !(o.endMs <= it.startMs || o.startMs >= it.endMs)
    );
    const maxLane: number = concurrent.reduce(
      (m: number, o: T & { lane: number }) => Math.max(m, o.lane),
      0
    );
    return { ...it, laneCount: maxLane + 1 };
  });
}

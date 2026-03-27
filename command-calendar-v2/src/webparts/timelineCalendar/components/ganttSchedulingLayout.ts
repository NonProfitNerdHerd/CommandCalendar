/**
 * Vertical lane layout for a single Gantt resource row (one swim lane).
 *
 * ## Lane assignment (greedy, interval scheduling on a line)
 * Items are sorted by start time. Each item is assigned the lowest lane index k such that
 * the previous item in lane k has **lane blocking end** strictly before this item's start.
 *
 * ## Bar-only vs visual footprint
 * By default, `laneBlockingEndMs` extends the bar's clipped end in the time domain to include
 * space for a **right-side label**: we map bar right edge + gap + estimated label width (px)
 * back to milliseconds on the same linear time axis as the track. Short bars with long titles
 * therefore reserve more horizontal "occupancy" and are pushed into additional vertical lanes
 * when another event would start before that occupancy clears.
 *
 * ## Vertical geometry
 * Each lane is one horizontal strip tall enough for the bar and a same-row label
 * (`max(barHeight, labelLineHeight)` + padding). Lane gaps add separation between strips.
 */

export interface IGanttLayoutInputItem<T = unknown> {
  id: string;
  startMs: number;
  endMsInclusive: number;
  /**
   * End time used only for lane stacking (max of bar end and label footprint in time).
   * If omitted, `endMsInclusive` is used.
   */
  laneBlockingEndMs?: number;
  data: T;
}

export interface IGanttScheduledPlacement<T = unknown> {
  data: T;
  laneIndex: number;
  /** Clipped bar bounds for drawing */
  layoutStartMs: number;
  layoutEndMsInclusive: number;
}

export interface IGanttLayoutMetrics {
  barHeightPx: number;
  /** Line box for right-side label (single line); also drives font-size estimate */
  labelLineHeightPx: number;
  /** Horizontal gap between bar end and label start */
  barToLabelGapPx: number;
  /** Legacy vertical gap (used by below-bar fallback layout only) */
  intraLaneGapPx: number;
  /** Vertical gap between stacked lanes */
  laneGapPx: number;
  paddingTopPx: number;
  paddingBottomPx: number;
}

export interface IGanttLaneLayoutResult<T = unknown> {
  placements: IGanttScheduledPlacement<T>[];
  laneCount: number;
  trackHeightPx: number;
}

export interface IGanttLaneVerticalGeometry {
  laneTopPx: (laneIndex: number) => number;
  laneHeightsPx: number[];
  trackHeightPx: number;
}

const DEFAULT_METRICS: IGanttLayoutMetrics = {
  barHeightPx: 22,
  labelLineHeightPx: 16,
  barToLabelGapPx: 8,
  intraLaneGapPx: 6,
  laneGapPx: 14,
  paddingTopPx: 10,
  paddingBottomPx: 12
};

/** Heuristic label width for lane footprint (cap avoids absurd ms spread). */
export function estimateGanttLabelWidthPx(title: string, fontSizePx: number, maxPx: number = 560): number {
  const t: string = title || '';
  const avgChar: number = fontSizePx * 0.52;
  return Math.min(maxPx, Math.ceil(t.length * avgChar) + 12);
}

/** Horizontal padding reserved inside the bar when comparing title width to bar width (px). */
export const GANTT_IN_BAR_TITLE_PADDING_PX = 14;

/**
 * True when the estimated single-line title width fits inside the visible bar (DABAL in-bar label).
 * If false, render the title to the right of the bar and use label footprint for lane stacking.
 */
export function ganttTitleFitsInsideBar(
  title: string,
  fontSizePx: number,
  barWidthPx: number
): boolean {
  if (!barWidthPx || barWidthPx < GANTT_IN_BAR_TITLE_PADDING_PX + 6) {
    return false;
  }
  const usable: number = barWidthPx - GANTT_IN_BAR_TITLE_PADDING_PX;
  const est: number = estimateGanttLabelWidthPx(title, fontSizePx, 2000);
  return est <= usable;
}

/**
 * Map right edge of bar+label footprint (px) to a time >= bar end so greedy lanes treat
 * long right-side titles like extended horizontal occupancy.
 */
export function computeLaneBlockingEndMs(
  startMs: number,
  endMsInclusive: number,
  trackStartMs: number,
  trackEndMs: number,
  trackWidthPx: number,
  labelWidthPx: number,
  barToLabelGapPx: number
): number {
  if (trackWidthPx <= 0 || trackEndMs <= trackStartMs) {
    return endMsInclusive;
  }

  const span: number = trackEndMs - trackStartMs;
  const W: number = trackWidthPx;
  const clipEnd: number = Math.min(endMsInclusive, trackEndMs);
  const barRightPx: number = ((clipEnd - trackStartMs) / span) * W;
  const occupiedRightPx: number = Math.min(W, barRightPx + barToLabelGapPx + labelWidthPx);
  const occupiedEndMs: number = trackStartMs + (occupiedRightPx / W) * span;
  return Math.max(endMsInclusive, Math.ceil(occupiedEndMs));
}

/** One lane row: bar + right label share vertical center; readability over compactness. */
export function defaultLaneStripHeightPx(metrics: IGanttLayoutMetrics): number {
  const core: number = Math.max(metrics.barHeightPx, metrics.labelLineHeightPx);
  return metrics.paddingTopPx + core + metrics.paddingBottomPx;
}

/** Below-bar fallback strip when the track is too narrow for a stable right-side label row. */
export function ganttBelowBarFallbackStripHeightPx(metrics: IGanttLayoutMetrics): number {
  return metrics.barHeightPx + metrics.intraLaneGapPx + Math.max(40, metrics.labelLineHeightPx * 2);
}

/** In-bar fallback: colored bar with wrapped/ellipsis title inside. */
export function ganttInBarFallbackStripHeightPx(metrics: IGanttLayoutMetrics): number {
  return Math.max(metrics.barHeightPx + 8, metrics.labelLineHeightPx * 3 + 8);
}

export function intervalsOverlapInclusive(
  aStart: number,
  aEndInclusive: number,
  bStart: number,
  bEndInclusive: number
): boolean {
  return aStart <= bEndInclusive && bStart <= aEndInclusive;
}

export function clipIntervalToWindow(
  itemStart: number,
  itemEndInclusive: number,
  winStart: number,
  winEndInclusive: number
): { startMs: number; endMsInclusive: number } | null {
  const cs: number = Math.max(itemStart, winStart);
  const ce: number = Math.min(itemEndInclusive, winEndInclusive);
  if (cs > ce) {
    return null;
  }

  return { startMs: cs, endMsInclusive: ce };
}

export function laneTopOffsetPx(laneIndex: number, metrics: IGanttLayoutMetrics): number {
  const strip: number = defaultLaneStripHeightPx(metrics);
  return metrics.paddingTopPx + laneIndex * (strip + metrics.laneGapPx);
}

export function computeTrackHeightPx(laneCount: number, metrics: IGanttLayoutMetrics): number {
  const strip: number = defaultLaneStripHeightPx(metrics);
  const m: IGanttLayoutMetrics = metrics;
  if (laneCount <= 0) {
    return m.paddingTopPx + strip + m.paddingBottomPx;
  }

  return (
    m.paddingTopPx +
    laneCount * strip +
    (laneCount - 1) * m.laneGapPx +
    m.paddingBottomPx
  );
}

export function computeLaneVerticalGeometry<T>(
  laneCount: number,
  placements: IGanttScheduledPlacement<T>[],
  getStripHeightPx: (p: IGanttScheduledPlacement<T>) => number,
  metrics: IGanttLayoutMetrics
): IGanttLaneVerticalGeometry {
  const fallbackStrip: number = defaultLaneStripHeightPx(metrics);
  const n: number = Math.max(0, laneCount);

  if (n === 0) {
    return {
      laneTopPx: () => 0,
      laneHeightsPx: [],
      trackHeightPx: metrics.paddingTopPx + fallbackStrip + metrics.paddingBottomPx
    };
  }

  const heights: number[] = [];
  for (let i: number = 0; i < n; i++) {
    heights.push(0);
  }
  for (const p of placements) {
    if (p.laneIndex >= 0 && p.laneIndex < n) {
      const h: number = getStripHeightPx(p);
      heights[p.laneIndex] = Math.max(heights[p.laneIndex], h);
    }
  }

  for (let i: number = 0; i < n; i++) {
    if (heights[i] === 0) {
      heights[i] = fallbackStrip;
    }
  }

  const topByLane: number[] = new Array(n);
  let y: number = metrics.paddingTopPx;
  for (let i: number = 0; i < n; i++) {
    topByLane[i] = y;
    y += heights[i];
    if (i < n - 1) {
      y += metrics.laneGapPx;
    }
  }

  return {
    laneTopPx: (laneIndex: number) => topByLane[laneIndex] ?? 0,
    laneHeightsPx: heights,
    trackHeightPx: y + metrics.paddingBottomPx
  };
}

function blockingEndForItem<T>(it: IGanttLayoutInputItem<T>): number {
  return it.laneBlockingEndMs != null ? it.laneBlockingEndMs : it.endMsInclusive;
}

/**
 * Assign lanes using **laneBlockingEndMs** (label footprint) when provided.
 */
export function computeGanttLaneLayout<T>(
  items: IGanttLayoutInputItem<T>[],
  metrics: IGanttLayoutMetrics = DEFAULT_METRICS
): IGanttLaneLayoutResult<T> {
  if (!items || items.length === 0) {
    return {
      placements: [],
      laneCount: 0,
      trackHeightPx: computeTrackHeightPx(0, metrics)
    };
  }

  const sorted: IGanttLayoutInputItem<T>[] = [...items].sort(
    (a: IGanttLayoutInputItem<T>, b: IGanttLayoutInputItem<T>) => {
      const ds: number = a.startMs - b.startMs;
      if (ds !== 0) {
        return ds;
      }

      return blockingEndForItem(b) - blockingEndForItem(a);
    }
  );

  const laneLastBlockingEnd: number[] = [];
  const placements: IGanttScheduledPlacement<T>[] = [];

  for (const it of sorted) {
    const blockEnd: number = blockingEndForItem(it);
    let laneIndex: number = -1;
    for (let k: number = 0; k < laneLastBlockingEnd.length; k++) {
      if (laneLastBlockingEnd[k] < it.startMs) {
        laneIndex = k;
        laneLastBlockingEnd[k] = blockEnd;
        break;
      }
    }

    if (laneIndex < 0) {
      laneIndex = laneLastBlockingEnd.length;
      laneLastBlockingEnd.push(blockEnd);
    }

    placements.push({
      data: it.data,
      laneIndex,
      layoutStartMs: it.startMs,
      layoutEndMsInclusive: it.endMsInclusive
    });
  }

  const laneCount: number = laneLastBlockingEnd.length;

  return {
    placements,
    laneCount,
    trackHeightPx: computeTrackHeightPx(laneCount, metrics)
  };
}

export function getDefaultGanttLayoutMetrics(): IGanttLayoutMetrics {
  return { ...DEFAULT_METRICS };
}

export function getGanttLayoutMetricsForVisibleRange(visibleDays: number): IGanttLayoutMetrics {
  const d: number = Math.max(GANTT_ZOOM_MIN_DAYS_REF, Math.min(visibleDays, GANTT_ZOOM_MAX_DAYS_REF));
  const density: number = Math.min(1, 14 / d);
  const scale: number = 1 + (1 - density) * 0.95;

  return {
    barHeightPx: Math.round(22 * scale),
    labelLineHeightPx: Math.round(16 * scale),
    barToLabelGapPx: Math.round(8 * scale),
    intraLaneGapPx: Math.round(6 * scale),
    laneGapPx: Math.round(14 * scale),
    paddingTopPx: Math.round(10 * scale),
    paddingBottomPx: Math.round(12 * scale)
  };
}

const GANTT_ZOOM_MIN_DAYS_REF = 7;
const GANTT_ZOOM_MAX_DAYS_REF = 365 * 5;

import * as React from 'react';
import TimelineCalendar from './TimelineCalendar';
import { ITimelineCalendarProps } from './ITimelineCalendarProps';
import { CalendarTimeGrid } from './CalendarTimeGrid';
import { clipIntervalToDayGrid } from './calendarTimeGridLayout';
import { downloadElementAsPdf } from './calendarPdfExport';
import {
  clipIntervalToWindow,
  computeGanttLaneLayout,
  computeLaneBlockingEndMs,
  computeLaneVerticalGeometry,
  defaultLaneStripHeightPx,
  estimateGanttLabelWidthPx,
  ganttBelowBarFallbackStripHeightPx,
  ganttTitleFitsInsideBar,
  getGanttLayoutMetricsForVisibleRange,
  IGanttLaneLayoutResult,
  IGanttLaneVerticalGeometry,
  IGanttLayoutInputItem,
  IGanttLayoutMetrics,
  IGanttScheduledPlacement
} from './ganttSchedulingLayout';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { TooltipHost } from 'office-ui-fabric-react/lib/Tooltip';

type TViewKey = 'gantt' | 'ganttZoom' | 'calendar' | 'horizon';
type TCalendarMode = 'day' | 'week5' | 'week7' | 'month';

interface ITimelineItem {
  id: string;
  title: string;
  start: Date;
  end: Date;
  categoryKey: string;
  categoryLabel: string;
  categoryColor: string;
  location: string;
  categoryText: string;
  description: string;
  author: string;
  editor: string;
  modified: string;
  eventUrl: string;
  encodedAbsUrl: string;
  spId: string;
  objType: string;
  calendarKey: string;
  calendarLabel: string;
  staffSectionKey: string;
  staffSectionLabel: string;
  ccMergedDaily?: boolean;
  ccDailyStartOffsetMs?: number;
  ccDailyDurationMs?: number;
  /** True for expanded recurrence slices (SP RecurrenceID / Graph occurrence) — no multi-day span bar */
  ccIsExpandedRecurrenceInstance?: boolean;
}

interface IWeekBarPlaced {
  event: ITimelineItem;
  startCol: number;
  span: number;
  lane: number;
}

interface ICategoryMeta {
  key: string;
  label: string;
  color: string;
}

interface ICalendarMeta {
  key: string;
  label: string;
}

interface IStaffSectionMeta {
  key: string;
  label: string;
}

interface ISwimLaneDef {
  key: string;
  label: string;
}

interface IHorizonBlock {
  label: string;
  startOffsetDays: number;
  endOffsetDays: number;
}

const DAY_LABELS: string[] = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const WEEKDAY_LABELS: string[] = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
const HORIZON_BLOCKS: IHorizonBlock[] = [
  { label: 'Next 30 Days', startOffsetDays: 0, endOffsetDays: 30 },
  { label: 'Days 31-60', startOffsetDays: 31, endOffsetDays: 60 },
  { label: 'Days 61-90', startOffsetDays: 61, endOffsetDays: 90 },
  { label: 'Days 91-120', startOffsetDays: 91, endOffsetDays: 120 }
];

type TGanttZoomGranularity = 'day' | 'week' | 'year';

const GANTT_ZOOM_MIN_DAYS = 7;
const GANTT_ZOOM_MAX_DAYS = 365 * 5;

function ganttZoomGranularity(visibleDays: number): TGanttZoomGranularity {
  if (visibleDays <= 14) {
    return 'day';
  }

  if (visibleDays <= 365) {
    return 'week';
  }

  return 'year';
}

interface IGanttTimeColumn {
  start: Date;
  end: Date;
  label: string;
}

function buildGanttZoomColumns(
  windowStart: Date,
  windowEnd: Date,
  granularity: TGanttZoomGranularity
): IGanttTimeColumn[] {
  const columns: IGanttTimeColumn[] = [];
  const we: number = endOfDay(windowEnd).getTime();

  if (granularity === 'day') {
    const lastDay: Date = endOfDay(windowEnd);
    for (let d: Date = startOfDay(windowStart); d.getTime() <= lastDay.getTime(); d = addDays(d, 1)) {
      columns.push({
        start: startOfDay(d),
        end: endOfDay(d),
        label: `${d.getMonth() + 1}/${d.getDate()}`
      });
    }
    return columns;
  }

  if (granularity === 'week') {
    let cursor: Date = startOfWeek(startOfDay(windowStart), true);
    while (cursor.getTime() <= we) {
      const weekEnd: Date = endOfDay(addDays(cursor, 6));
      columns.push({
        start: startOfDay(cursor),
        end: weekEnd.getTime() > we ? new Date(we) : weekEnd,
        label: `W ${cursor.getMonth() + 1}/${cursor.getDate()}`
      });
      cursor = addDays(cursor, 7);
    }
    return columns;
  }

  const startYear: number = startOfDay(windowStart).getFullYear();
  const endYear: number = startOfDay(windowEnd).getFullYear();
  for (let y: number = startYear; y <= endYear; y++) {
    const yStart: Date = new Date(y, 0, 1, 0, 0, 0, 0);
    const yEnd: Date = new Date(y, 11, 31, 23, 59, 59, 999);
    columns.push({
      start: yStart,
      end: yEnd,
      label: String(y)
    });
  }
  return columns;
}

function ganttEventBarPercent(
  eventItem: ITimelineItem,
  trackStartMs: number,
  trackEndMs: number
): { leftPct: number; widthPct: number } | null {
  const evStart: number = eventItem.start.getTime();
  const evEnd: number = getInclusiveEventEnd(eventItem).getTime();
  if (evEnd < trackStartMs || evStart > trackEndMs) {
    return null;
  }

  const clipStart: number = Math.max(evStart, trackStartMs);
  const clipEnd: number = Math.min(evEnd, trackEndMs);
  const span: number = trackEndMs - trackStartMs;
  if (span <= 0) {
    return null;
  }

  const leftPct: number = ((clipStart - trackStartMs) / span) * 100;
  const widthPct: number = Math.max(0.35, ((clipEnd - clipStart) / span) * 100);
  return { leftPct, widthPct };
}

/**
 * Lane inputs with `laneBlockingEndMs`: bar end extended by estimated right-label width (px → time),
 * so greedy lanes stack vertically when titles would collide horizontally.
 */
function buildGanttLayoutItemsWithLabelFootprint(
  laneEvents: ITimelineItem[],
  trackStartMs: number,
  trackEndMs: number,
  trackWidthPx: number,
  metrics: IGanttLayoutMetrics
): IGanttLayoutInputItem<ITimelineItem>[] {
  const items: IGanttLayoutInputItem<ITimelineItem>[] = [];
  const fontPx: number = metrics.labelLineHeightPx;

  laneEvents.forEach((eventItem: ITimelineItem) => {
    const rawEnd: number = getInclusiveEventEnd(eventItem).getTime();
    const clipped = clipIntervalToWindow(
      eventItem.start.getTime(),
      rawEnd,
      trackStartMs,
      trackEndMs
    );

    if (!clipped) {
      return;
    }

    let startMs: number = clipped.startMs;
    let endMsInclusive: number = clipped.endMsInclusive;
    if (startMs === endMsInclusive) {
      endMsInclusive = startMs + 1;
    }

    const spanMs: number = trackEndMs - trackStartMs;
    const barWidthPx: number =
      trackWidthPx > 0 && spanMs > 0
        ? ((endMsInclusive - startMs) / spanMs) * trackWidthPx
        : 0;
    const labelInsideBar: boolean = ganttTitleFitsInsideBar(eventItem.title, fontPx, barWidthPx);
    const labelWidthPx: number = estimateGanttLabelWidthPx(eventItem.title, fontPx);
    const laneBlockingEndMs: number =
      trackWidthPx > 0 && !labelInsideBar
        ? computeLaneBlockingEndMs(
            startMs,
            endMsInclusive,
            trackStartMs,
            trackEndMs,
            trackWidthPx,
            labelWidthPx,
            metrics.barToLabelGapPx
          )
        : endMsInclusive;

    items.push({
      id: String(eventItem.id),
      startMs,
      endMsInclusive,
      laneBlockingEndMs,
      data: eventItem
    });
  });

  return items;
}

/** Primary layout: label to the right of the bar in the same lane. Below-bar only when track is extremely narrow. */
const GANTT_MIN_TRACK_WIDTH_RIGHT_LABEL_PX = 160;

type TGanttCellLayoutMode = 'right' | 'below-fallback';

function ganttCellLayoutMode(trackWidthPx: number): TGanttCellLayoutMode {
  if (trackWidthPx > 0 && trackWidthPx < GANTT_MIN_TRACK_WIDTH_RIGHT_LABEL_PX) {
    return 'below-fallback';
  }
  return 'right';
}

const ZoomableGanttChart: React.FC<{ events: ITimelineItem[]; swimLanes: ISwimLaneDef[] }> = ({ events, swimLanes }) => {
  const [visibleDays, setVisibleDays] = React.useState<number>(14);
  const [rangeCenter, setRangeCenter] = React.useState<Date>(() => startOfDay(new Date()));

  const panDays: number = React.useMemo(
    () => Math.max(1, Math.floor(visibleDays * 0.85)),
    [visibleDays]
  );

  const { windowStart, windowEnd, trackStartMs, trackEndMs } = React.useMemo(() => {
    const halfFloor: number = Math.floor((visibleDays - 1) / 2);
    const halfCeil: number = visibleDays - 1 - halfFloor;
    const wStart: Date = startOfDay(addDays(rangeCenter, -halfFloor));
    const wEnd: Date = endOfDay(addDays(rangeCenter, halfCeil));
    return {
      windowStart: wStart,
      windowEnd: wEnd,
      trackStartMs: wStart.getTime(),
      trackEndMs: wEnd.getTime()
    };
  }, [visibleDays, rangeCenter]);

  const granularity: TGanttZoomGranularity = ganttZoomGranularity(visibleDays);
  const columns: IGanttTimeColumn[] = React.useMemo(
    () => buildGanttZoomColumns(windowStart, windowEnd, granularity),
    [windowStart, windowEnd, granularity]
  );

  const laneDefinitions: ISwimLaneDef[] = React.useMemo(
    () => (swimLanes.length > 0 ? swimLanes : [{ key: '', label: 'Unassigned' }]),
    [swimLanes]
  );

  const ganttLayoutMetrics: IGanttLayoutMetrics = React.useMemo(
    () => getGanttLayoutMetricsForVisibleRange(visibleDays),
    [visibleDays]
  );

  const eventsByStaffKey: Map<string, ITimelineItem[]> = React.useMemo(() => {
    const map: Map<string, ITimelineItem[]> = new Map<string, ITimelineItem[]>();
    events.forEach((eventItem: ITimelineItem) => {
      const key: string = eventItem.staffSectionKey ? String(eventItem.staffSectionKey) : '';
      const bucket: ITimelineItem[] = map.get(key);
      if (bucket) {
        bucket.push(eventItem);
      } else {
        map.set(key, [eventItem]);
      }
    });
    return map;
  }, [events]);

  const trackMeasureRef = React.useRef<HTMLDivElement | null>(null);
  const [trackWidthPx, setTrackWidthPx] = React.useState<number>(0);

  const swimLaneTrackLayouts: IGanttLaneLayoutResult<ITimelineItem>[] = React.useMemo(() => {
    return laneDefinitions.map((lane: ISwimLaneDef) => {
      const candidates: ITimelineItem[] = eventsByStaffKey.get(lane.key) || [];
      const laneEvents: ITimelineItem[] = candidates.filter((eventItem: ITimelineItem) => {
        const evEnd: number = getInclusiveEventEnd(eventItem).getTime();
        return !(evEnd < trackStartMs || eventItem.start.getTime() > trackEndMs);
      });

      const layoutInputs: IGanttLayoutInputItem<ITimelineItem>[] = buildGanttLayoutItemsWithLabelFootprint(
        laneEvents,
        trackStartMs,
        trackEndMs,
        trackWidthPx,
        ganttLayoutMetrics
      );

      return computeGanttLaneLayout(layoutInputs, ganttLayoutMetrics);
    });
  }, [
    laneDefinitions,
    eventsByStaffKey,
    trackStartMs,
    trackEndMs,
    ganttLayoutMetrics,
    trackWidthPx
  ]);

  const cellLayoutMode: TGanttCellLayoutMode = ganttCellLayoutMode(trackWidthPx);

  const swimLaneGeometries: IGanttLaneVerticalGeometry[] = React.useMemo(() => {
    return swimLaneTrackLayouts.map((layout: IGanttLaneLayoutResult<ITimelineItem>) => {
      const getStripHeightPx = (p: IGanttScheduledPlacement<ITimelineItem>): number => {
        const barInfo = ganttEventBarPercent(p.data, trackStartMs, trackEndMs);
        if (!barInfo) {
          return defaultLaneStripHeightPx(ganttLayoutMetrics);
        }
        return cellLayoutMode === 'below-fallback'
          ? ganttBelowBarFallbackStripHeightPx(ganttLayoutMetrics)
          : defaultLaneStripHeightPx(ganttLayoutMetrics);
      };
      return computeLaneVerticalGeometry(layout.laneCount, layout.placements, getStripHeightPx, ganttLayoutMetrics);
    });
  }, [swimLaneTrackLayouts, trackStartMs, trackEndMs, ganttLayoutMetrics, cellLayoutMode]);

  const anyEventsInWindow: boolean = React.useMemo(
    () => swimLaneTrackLayouts.some((layout: IGanttLaneLayoutResult<ITimelineItem>) => layout.placements.length > 0),
    [swimLaneTrackLayouts]
  );

  const colCount: number = Math.max(columns.length, 1);
  const gridTemplate: string = `minmax(140px, 200px) repeat(${colCount}, minmax(24px, 1fr))`;

  React.useLayoutEffect(() => {
    const el: HTMLDivElement | null = trackMeasureRef.current;
    if (!el || typeof ResizeObserver === 'undefined') {
      return undefined;
    }
    const updateWidth = (): void => {
      const w: number = el.getBoundingClientRect().width;
      if (w > 0) {
        setTrackWidthPx(Math.floor(w));
      }
    };
    updateWidth();
    const ro: ResizeObserver = new ResizeObserver(() => updateWidth());
    ro.observe(el);
    return () => {
      ro.disconnect();
    };
  }, [colCount, laneDefinitions.length, anyEventsInWindow]);

  return (
    <div style={{ border: '1px solid #edebe9', borderRadius: '4px', background: '#fff', overflow: 'auto' }}>
      <div
        style={{
          display: 'flex',
          flexWrap: 'wrap',
          justifyContent: 'flex-end',
          alignItems: 'center',
          gap: '8px',
          padding: '8px 10px',
          borderBottom: '1px solid #edebe9',
          background: '#faf9f8'
        }}
      >
        <span style={{ fontSize: '12px', color: '#605e5c', marginRight: 'auto' }}>
          {windowStart.toLocaleDateString()} – {windowEnd.toLocaleDateString()} ·{' '}
          {granularity === 'day' ? 'By day' : granularity === 'week' ? 'By week' : 'By year'}
        </span>
        <SmallButton label="◀ Back" onClick={() => setRangeCenter((c: Date) => startOfDay(addDays(c, -panDays)))} />
        <SmallButton label="Forward ▶" onClick={() => setRangeCenter((c: Date) => startOfDay(addDays(c, panDays)))} />
        <SmallButton label="Today" onClick={() => setRangeCenter(startOfDay(new Date()))} isPrimary />
        <SmallButton
          label="Zoom in"
          onClick={() => {
            setVisibleDays((d: number) => Math.max(GANTT_ZOOM_MIN_DAYS, Math.round(d / 1.35)));
          }}
          disabled={visibleDays <= GANTT_ZOOM_MIN_DAYS}
        />
        <SmallButton
          label="Zoom out"
          onClick={() => {
            setVisibleDays((d: number) => Math.min(GANTT_ZOOM_MAX_DAYS, Math.round(d * 1.35)));
          }}
          disabled={visibleDays >= GANTT_ZOOM_MAX_DAYS}
        />
      </div>

      <div style={{ minWidth: `${280 + colCount * 28}px` }}>
        <div style={{ display: 'grid', gridTemplateColumns: gridTemplate, background: '#f3f2f1', borderBottom: '1px solid #edebe9' }}>
          <div style={{ padding: '6px 8px', fontWeight: 600, fontSize: '12px', color: '#605e5c' }} aria-hidden>
            {'\u00a0'}
          </div>
          {columns.map((col: IGanttTimeColumn, index: number) => (
            <div
              key={`gantt-h-${index}-${col.label}`}
              style={{
                padding: '6px 2px',
                fontSize: '10px',
                fontWeight: 600,
                textAlign: 'center',
                borderLeft: '1px solid #edebe9',
                overflow: 'hidden',
                textOverflow: 'ellipsis',
                whiteSpace: 'nowrap'
              }}
              title={col.label}
            >
              {col.label}
            </div>
          ))}
        </div>

        {!anyEventsInWindow && (
          <div style={{ padding: '24px', fontSize: '13px', color: '#605e5c' }}>No events in this time range.</div>
        )}

        {laneDefinitions.map((lane: ISwimLaneDef, laneRowIndex: number) => {
          const layout: IGanttLaneLayoutResult<ITimelineItem> = swimLaneTrackLayouts[laneRowIndex];
          const geo: IGanttLaneVerticalGeometry = swimLaneGeometries[laneRowIndex];
          const trackHeightPx: number = geo.trackHeightPx;

          return (
            <div
              key={`gantt-lane-${lane.key || 'blank'}`}
              style={{
                display: 'grid',
                gridTemplateColumns: gridTemplate,
                borderBottom: '1px solid #edebe9',
                alignItems: 'stretch',
                minHeight: 0
              }}
            >
              <div
                style={{
                  padding: '8px',
                  fontSize: '12px',
                  fontWeight: 600,
                  background: '#faf9f8',
                  borderRight: '1px solid #edebe9',
                  display: 'flex',
                  alignItems: 'flex-start',
                  wordBreak: 'break-word',
                  minHeight: `${trackHeightPx}px`,
                  boxSizing: 'border-box'
                }}
                title={lane.label}
              >
                {lane.label}
              </div>
              <div
                ref={laneRowIndex === 0 ? trackMeasureRef : undefined}
                style={{
                  gridColumn: `2 / span ${colCount}`,
                  position: 'relative',
                  minHeight: `${trackHeightPx}px`,
                  margin: '4px 6px',
                  overflow: 'visible'
                }}
              >
                {layout.placements.map((p: IGanttScheduledPlacement<ITimelineItem>) => {
                  const eventItem: ITimelineItem = p.data;
                  const bar = ganttEventBarPercent(eventItem, trackStartMs, trackEndMs);
                  if (!bar) {
                    return null;
                  }

                  const laneTopPx: number = geo.laneTopPx(p.laneIndex);
                  const stripHeightPx: number =
                    cellLayoutMode === 'below-fallback'
                      ? ganttBelowBarFallbackStripHeightPx(ganttLayoutMetrics)
                      : defaultLaneStripHeightPx(ganttLayoutMetrics);

                  return (
                    <GanttChartEventCell
                      key={`gantt-${lane.key || 'row'}-${eventItem.id}-L${p.laneIndex}`}
                      eventItem={eventItem}
                      bar={bar}
                      metrics={ganttLayoutMetrics}
                      laneTopPx={laneTopPx}
                      stripHeightPx={stripHeightPx}
                      layoutMode={cellLayoutMode}
                      trackWidthPx={trackWidthPx}
                    />
                  );
                })}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
};

const TimelineCalendarTabbed: React.FC<ITimelineCalendarProps> = (props: ITimelineCalendarProps) => {
  const [activeView, setActiveView] = React.useState<TViewKey>('calendar');
  const [events, setEvents] = React.useState<ITimelineItem[]>([]);
  const [isLoadingData, setIsLoadingData] = React.useState<boolean>(true);
  const [selectedStaffSectionKeys, setSelectedStaffSectionKeys] = React.useState<string[]>([]);
  const [selectedCategoryKeys, setSelectedCategoryKeys] = React.useState<string[]>([]);
  const [selectedCalendarKeys, setSelectedCalendarKeys] = React.useState<string[]>([]);
  const [titleSearchText, setTitleSearchText] = React.useState<string>('');
  const [calendarReferenceDate, setCalendarReferenceDate] = React.useState<Date>(startOfDay(new Date()));
  const [calendarMode, setCalendarMode] = React.useState<TCalendarMode>('month');
  const calendarPrintRef = React.useRef<HTMLDivElement | null>(null);
  const [calendarPrintBusy, setCalendarPrintBusy] = React.useState<boolean>(false);
  const eventsSignatureRef = React.useRef<string>('');

  const timelineElement = React.useMemo((): JSX.Element => (
    <TimelineCalendar
      {...props}
      selectedStaffSectionKeys={selectedStaffSectionKeys}
      selectedCategoryKeys={selectedCategoryKeys}
      selectedCalendarKeys={selectedCalendarKeys}
      selectedTextFilter={titleSearchText}
      hideLegendBar
    />
  ), [props, selectedStaffSectionKeys, selectedCategoryKeys, selectedCalendarKeys, titleSearchText]);

  const categoryMetaByKey = React.useMemo((): Map<string, ICategoryMeta> => {
    const nextMap: Map<string, ICategoryMeta> = new Map<string, ICategoryMeta>();
    const categories: any[] = props.categories || [];

    categories.forEach((category: any) => {
      if (!category || !category.name) {
        return;
      }

      const key: string = props.ensureValidClassName(category.name);
      if (!key) {
        return;
      }

      const stylesText: string = String(props.buildDivStyles(category) || '');
      const categoryColor: string = extractCategoryColor(stylesText) || category.bgColor || category.borderColor || '#8a8886';
      nextMap.set(key, {
        key,
        label: category.name,
        color: categoryColor
      });
    });

    return nextMap;
  }, [props.categories, props.buildDivStyles, props.ensureValidClassName]);

  const staffSectionMetaByKey = React.useMemo((): Map<string, IStaffSectionMeta> => {
    const nextMap: Map<string, IStaffSectionMeta> = new Map<string, IStaffSectionMeta>();
    const groups: any[] = props.groups || [];
    groups.forEach((group: any) => {
      if (!group || !group.uniqueId || !group.name) {
        return;
      }

      nextMap.set(String(group.uniqueId), {
        key: String(group.uniqueId),
        label: String(group.name)
      });
    });

    return nextMap;
  }, [props.groups]);

  const captureEvents = React.useCallback((): void => {
    if (isLoadingData) {
      return;
    }

    const timelineGlobal: any = (window as any).TC;
    if (!timelineGlobal || !timelineGlobal.eventsDataSet || !timelineGlobal.eventsDataSet.get) {
      return;
    }

    const rawItems: any[] = timelineGlobal.eventsDataSet.get({
      filter: function(item: any): boolean {
        return item.className !== 'weekend';
      }
    }) || [];

    const mappedItems: ITimelineItem[] = rawItems
      .filter((item: any) => item && item.start)
      .map((item: any) => {
        const start: Date = new Date(item.start);
        const parsedEnd: Date = item.end ? new Date(item.end) : new Date(item.start);
        const end: Date = isNaN(parsedEnd.getTime()) || parsedEnd.getTime() < start.getTime() ? start : parsedEnd;
        const title: string = stripHtml(item.content || item.title || '(Untitled)');
        const categoryFromItem: ICategoryMeta = resolveEventCategory(item, categoryMetaByKey, props.ensureValidClassName);
        const calendarFromItem = resolveEventCalendar(item);
        const staffFromItem: IStaffSectionMeta = resolveEventStaffSection(item, staffSectionMetaByKey);
        const location: string = stripHtml(String(item.Location || ''));
        const categoryText: string = stripHtml(String(item.Category || item.category || categoryFromItem.label || ''));
        const description: string = stripHtml(String(item.Description || ''));
        const author: string = stripHtml(String(item.Author || ''));
        const editor: string = stripHtml(String(item.Editor || ''));
        const modified: string = stripHtml(String(item.Modified || ''));

        return {
          id: String(item.id || `${start.toISOString()}-${title}`),
          title,
          start,
          end,
          categoryKey: categoryFromItem.key,
          categoryLabel: categoryFromItem.label,
          categoryColor: categoryFromItem.color,
          location,
          categoryText,
          description,
          author,
          editor,
          modified,
          eventUrl: extractEventUrlFromRaw(item),
          encodedAbsUrl: String(item.encodedAbsUrl || ''),
          spId: String(item.spId || ''),
          objType: String(item.objType || ''),
          calendarKey: calendarFromItem.key,
          calendarLabel: calendarFromItem.label,
          staffSectionKey: staffFromItem.key,
          staffSectionLabel: staffFromItem.label,
          ccMergedDaily: !!item.ccMergedDaily,
          ccDailyStartOffsetMs: Number(item.ccDailyStartOffsetMs || 0),
          ccDailyDurationMs: Number(item.ccDailyDurationMs || 0),
          ccIsExpandedRecurrenceInstance: !!item.ccIsExpandedRecurrenceInstance
        };
      })
      .sort((firstEvent: ITimelineItem, secondEvent: ITimelineItem) => {
        const startDiff: number = firstEvent.start.getTime() - secondEvent.start.getTime();
        if (startDiff !== 0) {
          return startDiff;
        }

        return String(firstEvent.id).localeCompare(String(secondEvent.id));
      });

    const nextSignature: string = mappedItems
      .map((eventItem: ITimelineItem) => (
        `${eventItem.id}|${eventItem.start.getTime()}|${eventItem.end.getTime()}|${eventItem.title}|${eventItem.staffSectionKey}|${eventItem.categoryKey}|${eventItem.calendarKey}|${eventItem.calendarLabel}|${eventItem.modified}`
      ))
      .join('~');

    if (eventsSignatureRef.current === nextSignature) {
      return;
    }

    eventsSignatureRef.current = nextSignature;
    setEvents(mappedItems);
  }, [categoryMetaByKey, props.ensureValidClassName, isLoadingData, staffSectionMetaByKey]);

  const categoryFilterOptions: IDropdownOption[] = React.useMemo(() => {
    const optionMap: Map<string, ICategoryMeta> = new Map<string, ICategoryMeta>();

    categoryMetaByKey.forEach((meta: ICategoryMeta, key: string) => {
      optionMap.set(key, meta);
    });

    events.forEach((eventItem: ITimelineItem) => {
      if (!eventItem.categoryKey || optionMap.has(eventItem.categoryKey)) {
        return;
      }

      optionMap.set(eventItem.categoryKey, {
        key: eventItem.categoryKey,
        label: eventItem.categoryLabel || eventItem.categoryKey,
        color: eventItem.categoryColor || '#8a8886'
      });
    });

    const optionValues: ICategoryMeta[] = [];
    optionMap.forEach((meta: ICategoryMeta) => {
      optionValues.push(meta);
    });

    optionValues.sort((firstMeta: ICategoryMeta, secondMeta: ICategoryMeta) => firstMeta.label.localeCompare(secondMeta.label));

    return optionValues.map((meta: ICategoryMeta) => ({
      key: meta.key,
      text: meta.label,
      data: {
        color: meta.color
      }
    }));
  }, [categoryMetaByKey, events]);

  /** Events used to populate Staff Section list: narrowed by Category when categories are selected. */
  const eventsForStaffFilterOptions: ITimelineItem[] = React.useMemo(() => {
    if (selectedCategoryKeys.length === 0) {
      return events;
    }

    const categorySet: Set<string> = new Set<string>(selectedCategoryKeys.map((k: string) => String(k)));
    return events.filter(
      (eventItem: ITimelineItem) => eventItem.categoryKey && categorySet.has(String(eventItem.categoryKey))
    );
  }, [events, selectedCategoryKeys]);

  /**
   * Calendar options use the same event pool as Staff (Category → Staff), then narrow by selected staff.
   * Previously calendars were taken from all events when staff was selected, so Category+Staff could show
   * calendars from that staff in *other* categories (felt like the cascade was backwards).
   */
  const eventsForCalendarFilterOptions: ITimelineItem[] = React.useMemo(() => {
    const base: ITimelineItem[] = eventsForStaffFilterOptions;
    if (selectedStaffSectionKeys.length === 0) {
      return base;
    }

    const staffSet: Set<string> = new Set<string>(selectedStaffSectionKeys.map((k: string) => String(k)));
    return base.filter(
      (eventItem: ITimelineItem) =>
        eventItem.staffSectionKey && staffSet.has(String(eventItem.staffSectionKey))
    );
  }, [eventsForStaffFilterOptions, selectedStaffSectionKeys]);

  const staffSectionFilterOptions: IDropdownOption[] = React.useMemo(() => {
    const optionMap: Map<string, string> = new Map<string, string>();

    eventsForStaffFilterOptions.forEach((eventItem: ITimelineItem) => {
      if (!eventItem.staffSectionKey) {
        return;
      }

      const metaLabel: string | undefined = staffSectionMetaByKey.get(eventItem.staffSectionKey)?.label;
      const label: string = metaLabel || eventItem.staffSectionLabel || eventItem.staffSectionKey;
      if (!optionMap.has(eventItem.staffSectionKey)) {
        optionMap.set(eventItem.staffSectionKey, label);
      }
    });

    const options: IDropdownOption[] = [];
    optionMap.forEach((label: string, key: string) => {
      options.push({
        key,
        text: label
      });
    });

    options.sort((firstOption: IDropdownOption, secondOption: IDropdownOption) =>
      String(firstOption.text).localeCompare(String(secondOption.text))
    );
    return options;
  }, [eventsForStaffFilterOptions, staffSectionMetaByKey]);

  const calendarFilterOptions: IDropdownOption[] = React.useMemo(() => {
    const optionMap: Map<string, string> = new Map<string, string>();

    eventsForCalendarFilterOptions.forEach((eventItem: ITimelineItem) => {
      if (!eventItem.calendarKey || !eventItem.calendarLabel) {
        return;
      }

      if (!optionMap.has(eventItem.calendarKey)) {
        optionMap.set(eventItem.calendarKey, eventItem.calendarLabel);
      }
    });

    const options: IDropdownOption[] = [];
    optionMap.forEach((label: string, key: string) => {
      options.push({
        key,
        text: label
      });
    });

    options.sort((firstOption: IDropdownOption, secondOption: IDropdownOption) =>
      String(firstOption.text).localeCompare(String(secondOption.text))
    );
    return options;
  }, [eventsForCalendarFilterOptions]);

  const filteredEvents: ITimelineItem[] = React.useMemo(() => {
    const selectedStaffSectionKeySet: Set<string> = new Set<string>(selectedStaffSectionKeys.map((k: string) => String(k)));
    const selectedCategoryKeySet: Set<string> = new Set<string>(selectedCategoryKeys.map((k: string) => String(k)));
    const selectedCalendarKeySet: Set<string> = new Set<string>(selectedCalendarKeys.map((k: string) => String(k)));
    const normalizedTitleSearch = titleSearchText.trim().toLowerCase();

    return events.filter((eventItem: ITimelineItem) => {
      const staffSectionMatch = selectedStaffSectionKeys.length === 0 ||
        (eventItem.staffSectionKey && selectedStaffSectionKeySet.has(String(eventItem.staffSectionKey)));
      const categoryMatch = selectedCategoryKeys.length === 0 ||
        (eventItem.categoryKey && selectedCategoryKeySet.has(String(eventItem.categoryKey)));
      const calendarMatch = selectedCalendarKeys.length === 0 ||
        (eventItem.calendarKey && selectedCalendarKeySet.has(String(eventItem.calendarKey)));
      const titleMatch = normalizedTitleSearch.length === 0 ||
        eventItem.title.toLowerCase().indexOf(normalizedTitleSearch) !== -1;
      return staffSectionMatch && categoryMatch && calendarMatch && titleMatch;
    });
  }, [events, selectedStaffSectionKeys, selectedCategoryKeys, selectedCalendarKeys, titleSearchText]);

  const nonGanttDisplayEvents: ITimelineItem[] = React.useMemo(() => {
    return expandMergedDailyEventsForDisplay(filteredEvents);
  }, [filteredEvents]);

  const ganttZoomSwimLanes: ISwimLaneDef[] = React.useMemo((): ISwimLaneDef[] => {
    const groups: any[] = props.groups || [];
    let lanes: ISwimLaneDef[];

    if (groups.length > 0) {
      lanes = groups.map((g: any) => ({
        key: String(g.uniqueId),
        label: String(g.name || g.uniqueId)
      }));
    } else {
      const map: Map<string, string> = new Map<string, string>();
      nonGanttDisplayEvents.forEach((ev: ITimelineItem) => {
        const k: string = ev.staffSectionKey ? String(ev.staffSectionKey) : '';
        if (!map.has(k)) {
          map.set(k, ev.staffSectionLabel || (k === '' ? 'Unassigned' : k));
        }
      });
      lanes = [];
      map.forEach((label: string, key: string) => {
        lanes.push({ key, label });
      });
      lanes.sort((a: ISwimLaneDef, b: ISwimLaneDef) => a.label.localeCompare(b.label));
    }

    const hasUnassigned: boolean = nonGanttDisplayEvents.some((e: ITimelineItem) => !e.staffSectionKey);
    if (hasUnassigned && !lanes.some((l: ISwimLaneDef) => l.key === '')) {
      return [...lanes, { key: '', label: 'Unassigned' }];
    }

    return lanes;
  }, [props.groups, nonGanttDisplayEvents]);

  const onCategoryFilterChange = React.useCallback((event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) {
      return;
    }

    const optionKey: string = String(option.key);
    setSelectedCategoryKeys((previousKeys: string[]) => {
      if (option.selected) {
        if (previousKeys.indexOf(optionKey) > -1) {
          return previousKeys;
        }

        return [...previousKeys, optionKey];
      }

      return previousKeys.filter((key: string) => key !== optionKey);
    });
  }, []);

  const onStaffSectionFilterChange = React.useCallback((event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) {
      return;
    }

    const optionKey: string = String(option.key);
    setSelectedStaffSectionKeys((previousKeys: string[]) => {
      if (option.selected) {
        if (previousKeys.indexOf(optionKey) > -1) {
          return previousKeys;
        }

        return [...previousKeys, optionKey];
      }

      return previousKeys.filter((key: string) => key !== optionKey);
    });
  }, []);

  const onCalendarFilterChange = React.useCallback((event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) {
      return;
    }

    const optionKey: string = String(option.key);
    setSelectedCalendarKeys((previousKeys: string[]) => {
      if (option.selected) {
        if (previousKeys.indexOf(optionKey) > -1) {
          return previousKeys;
        }

        return [...previousKeys, optionKey];
      }

      return previousKeys.filter((key: string) => key !== optionKey);
    });
  }, []);

  const validStaffOptionKeys: string = React.useMemo(
    () => staffSectionFilterOptions.map((o: IDropdownOption) => String(o.key)).sort().join('|'),
    [staffSectionFilterOptions]
  );

  const validCalendarOptionKeys: string = React.useMemo(
    () => calendarFilterOptions.map((o: IDropdownOption) => String(o.key)).sort().join('|'),
    [calendarFilterOptions]
  );

  React.useEffect(() => {
    const valid: Set<string> = new Set<string>(
      staffSectionFilterOptions.map((o: IDropdownOption) => String(o.key))
    );
    setSelectedStaffSectionKeys((previousKeys: string[]) => {
      const next: string[] = previousKeys.filter((key: string) => valid.has(key));
      return next.length === previousKeys.length ? previousKeys : next;
    });
  }, [validStaffOptionKeys]);

  React.useEffect(() => {
    const valid: Set<string> = new Set<string>(
      calendarFilterOptions.map((o: IDropdownOption) => String(o.key))
    );
    setSelectedCalendarKeys((previousKeys: string[]) => {
      const next: string[] = previousKeys.filter((key: string) => valid.has(key));
      return next.length === previousKeys.length ? previousKeys : next;
    });
  }, [validCalendarOptionKeys]);

  React.useEffect(() => {
    if (activeView === 'gantt') {
      return;
    }

    captureEvents();
    const intervalId: number = window.setInterval(() => {
      captureEvents();
    }, 2000);

    return () => {
      window.clearInterval(intervalId);
    };
  }, [activeView, captureEvents]);

  React.useEffect(() => {
    if (activeView === 'gantt') {
      window.setTimeout(() => {
        window.dispatchEvent(new Event('resize'));
      }, 100);
    } else {
      captureEvents();
    }
  }, [activeView, captureEvents]);

  React.useEffect(() => {
    const handleLoadingEvent = (event: Event): void => {
      const customEvent = event as CustomEvent;
      if (!customEvent.detail || customEvent.detail.instanceId !== props.instanceId) {
        return;
      }

      setIsLoadingData(!!customEvent.detail.isLoading);
    };

    window.addEventListener('command-calendar-loading', handleLoadingEvent);
    return () => {
      window.removeEventListener('command-calendar-loading', handleLoadingEvent);
    };
  }, [props.instanceId]);

  const calendarPrintBusyRef = React.useRef<boolean>(false);
  const handleCalendarPrintPdf = React.useCallback(async (): Promise<void> => {
    const root: HTMLDivElement | null = calendarPrintRef.current;
    if (!root || calendarPrintBusyRef.current) {
      return;
    }
    calendarPrintBusyRef.current = true;
    setCalendarPrintBusy(true);
    try {
      const stamp: string = new Date().toISOString().slice(0, 10);
      await downloadElementAsPdf(root, `Command-Calendar-${calendarMode}-${stamp}.pdf`);
    } finally {
      calendarPrintBusyRef.current = false;
      setCalendarPrintBusy(false);
    }
  }, [calendarMode]);

  return (
    <div style={{ position: 'relative' }}>
      <div
        style={{
          background: '#f5f5f5',
          border: '1px solid #edebe9',
          borderRadius: '6px',
          padding: '10px',
          marginBottom: '10px'
        }}
      >
        <div
          style={{
            display: 'inline-flex',
            flexWrap: 'wrap',
            gap: '4px',
            padding: '5px',
            marginBottom: '12px',
            background: 'linear-gradient(180deg, #e8e8e8 0%, #d8d8d8 100%)',
            borderRadius: '12px',
            boxShadow: 'inset 0 1px 1px rgba(255,255,255,0.65), inset 0 -1px 1px rgba(0,0,0,0.06)',
            border: '1px solid #c8c8c8'
          }}
        >
          <TabButton isActive={activeView === 'calendar'} label="Calendar" onClick={() => setActiveView('calendar')} />
          <TabButton isActive={activeView === 'gantt'} label="Gantt Chart Timeline" onClick={() => setActiveView('gantt')} />
          <TabButton isActive={activeView === 'ganttZoom'} label="Gantt Chart DABAL" onClick={() => setActiveView('ganttZoom')} />
          <TabButton isActive={activeView === 'horizon'} label="30-60-90-120" onClick={() => setActiveView('horizon')} />
        </div>

        <div style={{ display: 'flex', gap: '8px', alignItems: 'flex-end', flexWrap: 'wrap' }}>
          <Dropdown
            label="Category"
            placeholder="All categories"
            multiSelect
            options={categoryFilterOptions}
            selectedKeys={selectedCategoryKeys}
            onChange={onCategoryFilterChange}
            styles={{ dropdown: { minWidth: 280 } }}
          />
          <Dropdown
            label="Staff Section"
            placeholder={selectedCategoryKeys.length === 0 ? 'All staff sections' : 'Staff in selected categories'}
            multiSelect
            options={staffSectionFilterOptions}
            selectedKeys={selectedStaffSectionKeys}
            onChange={onStaffSectionFilterChange}
            styles={{ dropdown: { minWidth: 260 } }}
            disabled={staffSectionFilterOptions.length === 0}
          />
          <Dropdown
            label="Calendar"
            placeholder={
              selectedStaffSectionKeys.length === 0
                ? 'All calendars'
                : 'Calendars for selected staff sections'
            }
            multiSelect
            options={calendarFilterOptions}
            selectedKeys={selectedCalendarKeys}
            onChange={onCalendarFilterChange}
            styles={{ dropdown: { minWidth: 280 } }}
            disabled={calendarFilterOptions.length === 0}
          />
          <SmallButton
            label="Clear Filters"
            onClick={() => {
              setSelectedStaffSectionKeys([]);
              setSelectedCategoryKeys([]);
              setSelectedCalendarKeys([]);
              setTitleSearchText('');
            }}
            disabled={selectedStaffSectionKeys.length === 0 && selectedCategoryKeys.length === 0 && selectedCalendarKeys.length === 0 && titleSearchText.trim().length === 0}
            isPrimary
          />
        </div>
        <div style={{ marginTop: '8px' }}>
          <TextField
            label="Title Search Filter"
            placeholder="Search event titles"
            value={titleSearchText}
            onChange={(_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
              setTitleSearchText(newValue || '');
            }}
            styles={{ fieldGroup: { maxWidth: 520 } }}
          />
        </div>
      </div>

      {activeView !== 'gantt' && (
        <CategoryLegend options={categoryFilterOptions} selectedCategoryKeys={selectedCategoryKeys} />
      )}

      <div style={{ display: activeView === 'calendar' ? 'block' : 'none' }}>
        <div ref={calendarPrintRef} style={{ background: '#fff' }}>
          <div
            style={{
              display: 'grid',
              gridTemplateColumns: '1fr auto 1fr',
              alignItems: 'center',
              gap: '8px',
              marginBottom: '10px'
            }}
          >
            <div style={{ display: 'flex', gap: '8px', justifySelf: 'start', flexWrap: 'wrap' }}>
              <SmallButton
                label={
                  calendarMode === 'month'
                    ? 'Last Month'
                    : calendarMode === 'day'
                      ? 'Previous Day'
                      : 'Previous Week'
                }
                onClick={() => {
                  if (calendarMode === 'month') {
                    setCalendarReferenceDate(addMonths(calendarReferenceDate, -1));
                  } else if (calendarMode === 'day') {
                    setCalendarReferenceDate(addDays(calendarReferenceDate, -1));
                  } else {
                    setCalendarReferenceDate(addDays(calendarReferenceDate, -7));
                  }
                }}
                isPrimary
              />
              <SmallButton label="Today" onClick={() => setCalendarReferenceDate(startOfDay(new Date()))} isPrimary />
              <SmallButton
                label={
                  calendarMode === 'month'
                    ? 'Next Month'
                    : calendarMode === 'day'
                      ? 'Next Day'
                      : 'Next Week'
                }
                onClick={() => {
                  if (calendarMode === 'month') {
                    setCalendarReferenceDate(addMonths(calendarReferenceDate, 1));
                  } else if (calendarMode === 'day') {
                    setCalendarReferenceDate(addDays(calendarReferenceDate, 1));
                  } else {
                    setCalendarReferenceDate(addDays(calendarReferenceDate, 7));
                  }
                }}
                isPrimary
              />
            </div>
            <div
              style={{
                fontWeight: 700,
                fontSize: '3em',
                textAlign: 'center',
                whiteSpace: 'nowrap',
                lineHeight: 1.1
              }}
            >
              {buildCalendarTitle(calendarReferenceDate, calendarMode)}
            </div>
            <div style={{ display: 'flex', gap: '8px', justifySelf: 'end', flexWrap: 'wrap', alignItems: 'center' }}>
              <SmallButton label="Day" onClick={() => setCalendarMode('day')} isActive={calendarMode === 'day'} />
              <SmallButton label="5-Day Week" onClick={() => setCalendarMode('week5')} isActive={calendarMode === 'week5'} />
              <SmallButton label="7-Day Week" onClick={() => setCalendarMode('week7')} isActive={calendarMode === 'week7'} />
              <SmallButton label="Month" onClick={() => setCalendarMode('month')} isActive={calendarMode === 'month'} />
              <SmallButton
                label={calendarPrintBusy ? 'Preparing PDF…' : 'Print PDF'}
                onClick={() => {
                  void handleCalendarPrintPdf();
                }}
                disabled={calendarPrintBusy}
              />
            </div>
          </div>
          <CalendarView events={nonGanttDisplayEvents} referenceDate={calendarReferenceDate} mode={calendarMode} />
        </div>
      </div>

      <div style={{ display: activeView === 'gantt' ? 'block' : 'none' }}>
        {timelineElement}
      </div>
      <div style={{ display: activeView === 'ganttZoom' ? 'block' : 'none' }}>
        <ZoomableGanttChart events={nonGanttDisplayEvents} swimLanes={ganttZoomSwimLanes} />
      </div>
      <div style={{ display: activeView === 'horizon' ? 'block' : 'none' }}>
        <HorizonView events={nonGanttDisplayEvents} />
      </div>

      {isLoadingData && (
        <div
          style={{
            position: 'absolute',
            top: 0,
            right: 0,
            bottom: 0,
            left: 0,
            zIndex: 1000,
            background: 'rgba(245, 245, 245, 0.86)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center'
          }}
        >
          <div
            style={{
              background: '#ffffff',
              border: '1px solid #d0d0d0',
              borderRadius: '6px',
              padding: '14px 20px',
              fontWeight: 700
            }}
          >
            loading data, please wait
          </div>
        </div>
      )}
    </div>
  );
};

const TabButton: React.FC<{ isActive: boolean; label: string; onClick: () => void }> = ({ isActive, label, onClick }) => (
  <button
    type="button"
    onClick={onClick}
    style={{
      border: isActive ? '1px solid #b0b0b0' : '1px solid transparent',
      background: isActive ? '#f5f5f5' : 'transparent',
      color: isActive ? '#323130' : '#605e5c',
      borderRadius: '9px',
      padding: '7px 14px',
      fontSize: '13px',
      fontWeight: isActive ? 700 : 600,
      letterSpacing: '0.01em',
      cursor: 'pointer',
      boxShadow: isActive ? '0 1px 2px rgba(0, 0, 0, 0.08), inset 0 1px 0 rgba(255,255,255,0.9)' : 'none',
      transition: 'background 0.15s ease, color 0.15s ease, box-shadow 0.15s ease'
    }}
  >
    {label}
  </button>
);

const SmallButton: React.FC<{ label: string; onClick: () => void; disabled?: boolean; isActive?: boolean; isPrimary?: boolean }> = ({
  label,
  onClick,
  disabled,
  isActive,
  isPrimary
}) => (
  <button
    type="button"
    disabled={disabled}
    onClick={onClick}
    style={{
      border: isPrimary ? '1px solid #0078d4' : (isActive ? '1px solid #0078d4' : '1px solid #c8c6c4'),
      background: isPrimary ? '#0078d4' : (isActive ? '#eff6fc' : '#fff'),
      color: isPrimary ? '#ffffff' : (isActive ? '#0078d4' : '#323130'),
      borderRadius: '4px',
      height: '32px',
      padding: '0 12px',
      cursor: disabled ? 'default' : 'pointer',
      opacity: disabled ? 0.6 : 1
    }}
  >
    {label}
  </button>
);

const CategoryLegend: React.FC<{ options: IDropdownOption[]; selectedCategoryKeys: string[] }> = ({ options, selectedCategoryKeys }) => {
  if (!options || options.length === 0) {
    return null;
  }

  return (
    <div style={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: '12px', marginBottom: '10px' }}>
      <span style={{ fontWeight: 600, fontSize: '12px' }}>Legend:</span>
      {options.map((option: IDropdownOption) => {
        const optionColor: string = option.data && option.data.color ? String(option.data.color) : '#8a8886';
        const isDimmed: boolean = selectedCategoryKeys.length > 0 && selectedCategoryKeys.indexOf(String(option.key)) === -1;
        return (
          <span
            key={`legend-${String(option.key)}`}
            style={{
              fontSize: '12px',
              color: '#323130',
              opacity: isDimmed ? 0.45 : 1
            }}
          >
            <CategoryDot color={optionColor} />
            {option.text}
          </span>
        );
      })}
    </div>
  );
};

const CalendarView: React.FC<{ events: ITimelineItem[]; referenceDate: Date; mode: TCalendarMode }> = ({
  events,
  referenceDate,
  mode
}) => {
  if (mode === 'month') {
    return <CalendarMonthGrid events={events} referenceDate={referenceDate} />;
  }

  return <CalendarWeekTimeView events={events} referenceDate={referenceDate} mode={mode} />;
};

const CalendarMonthGrid: React.FC<{ events: ITimelineItem[]; referenceDate: Date }> = ({ events, referenceDate }) => {
  const monthStart: Date = new Date(referenceDate.getFullYear(), referenceDate.getMonth(), 1);
  const gridStart: Date = startOfWeek(monthStart, false);
  const [expandedDayKeys, setExpandedDayKeys] = React.useState<string[]>([]);
  const gridDays: Date[] = [];

  for (let index = 0; index < 42; index++) {
    gridDays.push(addDays(gridStart, index));
  }

  const weekChunks: Date[][] = [];
  for (let w = 0; w < gridDays.length; w += 7) {
    weekChunks.push(gridDays.slice(w, w + 7));
  }

  return (
    <div style={{ border: '1px solid #edebe9' }}>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, minmax(120px, 1fr))', background: '#faf9f8', borderBottom: '1px solid #edebe9' }}>
        {DAY_LABELS.map((label: string) => (
          <div key={label} style={{ padding: '6px', fontWeight: 600, fontSize: '12px' }}>{label}</div>
        ))}
      </div>
      {weekChunks.map((weekDates: Date[], weekIndex: number) => {
        const barPlacements: IWeekBarPlaced[] = computeWeekOutlookBars(weekDates, events);
        const maxLane: number = barPlacements.reduce((acc: number, p: IWeekBarPlaced) => Math.max(acc, p.lane), -1);
        const barRowCount: number = maxLane + 1;

        return (
          <div key={`month-week-${weekIndex}`}>
            {barRowCount > 0 && (
              <div
                style={{
                  display: 'grid',
                  gridTemplateColumns: 'repeat(7, minmax(0, 1fr))',
                  columnGap: '1px',
                  rowGap: '3px',
                  padding: '4px 0 2px',
                  borderBottom: '1px solid #edebe9',
                  background: '#faf9f8',
                  gridTemplateRows: `repeat(${barRowCount}, minmax(18px, auto))`
                }}
              >
                {barPlacements.map((p: IWeekBarPlaced) => (
                  <div
                    key={`bar-${p.event.id}-w${weekIndex}-l${p.lane}-c${p.startCol}`}
                    style={{
                      gridColumn: `${p.startCol + 1} / span ${p.span}`,
                      gridRow: p.lane + 1,
                      margin: '0 1px',
                      minHeight: '18px',
                      display: 'flex',
                      alignItems: 'center',
                      overflow: 'hidden'
                    }}
                  >
                    <EventLinkWithTooltip
                      event={p.event}
                      compact
                      showDateRange={false}
                      renderAsBar
                      barStartsHere
                      barEndsHere
                    />
                  </div>
                ))}
              </div>
            )}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, minmax(120px, 1fr))' }}>
              {weekDates.map((day: Date) => {
                const dayEvents: ITimelineItem[] = eventsForDate(events, day).filter(
                  (eventItem: ITimelineItem) => !shouldShowAsOutlookSpanBar(eventItem)
                );
                const isCurrentMonth: boolean = day.getMonth() === monthStart.getMonth();
                const dayKey = `${day.getFullYear()}-${day.getMonth()}-${day.getDate()}`;
                const isExpanded = expandedDayKeys.indexOf(dayKey) > -1;
                const visibleEvents = isExpanded ? dayEvents : dayEvents.slice(0, 4);
                return (
                  <div
                    key={`month-${day.toISOString()}`}
                    style={{
                      borderRight: '1px solid #f3f2f1',
                      borderBottom: '1px solid #f3f2f1',
                      minHeight: '100px',
                      padding: '6px',
                      background: isCurrentMonth ? '#fff' : '#faf9f8'
                    }}
                  >
                    <div style={{ fontWeight: 600, fontSize: '12px', marginBottom: '4px' }}>{day.getDate()}</div>
                    {visibleEvents.map((eventItem: ITimelineItem) => (
                      <EventLinkWithTooltip
                        key={`month-${day.toISOString()}-${eventItem.id}`}
                        event={eventItem}
                        compact
                        showDateRange
                      />
                    ))}
                    {dayEvents.length > 4 && (
                      <button
                        type="button"
                        onClick={() => {
                          setExpandedDayKeys((previousKeys: string[]) => {
                            if (previousKeys.indexOf(dayKey) > -1) {
                              return previousKeys.filter((key: string) => key !== dayKey);
                            }

                            return [...previousKeys, dayKey];
                          });
                        }}
                        style={{
                          border: 'none',
                          background: 'transparent',
                          color: '#0078d4',
                          cursor: 'pointer',
                          fontSize: '11px',
                          padding: 0
                        }}
                      >
                        {isExpanded ? 'Show less' : `+${dayEvents.length - 4} more`}
                      </button>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        );
      })}
    </div>
  );
};

const HorizonView: React.FC<{ events: ITimelineItem[] }> = ({ events }) => {
  const todayStart: Date = startOfDay(new Date());
  const endOfToday: Date = endOfDay(new Date());

  return (
    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, minmax(190px, 1fr))', gap: '10px' }}>
      {HORIZON_BLOCKS.map((block: IHorizonBlock) => {
        const rangeStart: Date = addDays(todayStart, block.startOffsetDays);
        const rangeEnd: Date = addDays(endOfToday, block.endOffsetDays);
        const blockEvents: ITimelineItem[] = events
          .filter((eventItem: ITimelineItem) => eventItem.end.getTime() >= rangeStart.getTime() && eventItem.start.getTime() <= rangeEnd.getTime())
          .sort((firstEvent: ITimelineItem, secondEvent: ITimelineItem) => firstEvent.start.getTime() - secondEvent.start.getTime());

        return (
          <div key={block.label} style={{ border: '1px solid #edebe9', borderRadius: '6px', padding: '10px' }}>
            <div style={{ fontSize: '16px', fontWeight: 600 }}>{block.label}</div>
            <div style={{ fontSize: '28px', fontWeight: 700, margin: '8px 0' }}>{blockEvents.length}</div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
              {blockEvents.map((eventItem: ITimelineItem) => (
                <div key={`${block.label}-${eventItem.id}`} style={{ fontSize: '12px' }}>
                  <EventLinkWithTooltip event={eventItem} />
                  <div style={{ color: '#605e5c', marginLeft: '14px' }}>
                    {eventItem.start.toLocaleString()} - {eventItem.end.toLocaleString()}
                  </div>
                </div>
              ))}
              {blockEvents.length === 0 && (
                <div style={{ fontSize: '12px', color: '#605e5c' }}>No events in this range.</div>
              )}
            </div>
          </div>
        );
      })}
    </div>
  );
};

const EventLinkWithTooltip: React.FC<{
  event: ITimelineItem;
  compact?: boolean;
  showDateRange?: boolean;
  renderAsBar?: boolean;
  barStartsHere?: boolean;
  barEndsHere?: boolean;
  hideCategoryDot?: boolean;
  ganttTitleBesideBar?: boolean;
  /** Title below Gantt bar — full title wraps (no ellipsis); tooltip still has details */
  ganttTrackLabel?: boolean;
  /** Single-line title to the right of the bar (ellipsis; tooltip has full title) */
  ganttRightOfBarLabel?: boolean;
  /** In-bar title on DABAL: wrap full text inside colored bar */
  ganttWrapTitleInBar?: boolean;
  omitSurfaceZIndex?: boolean;
}> = ({
  event,
  compact,
  showDateRange,
  renderAsBar,
  barStartsHere,
  barEndsHere,
  hideCategoryDot,
  ganttTitleBesideBar,
  ganttTrackLabel,
  ganttRightOfBarLabel,
  ganttWrapTitleInBar,
  omitSurfaceZIndex
}) => {
  const eventUrl: string = getEventUrl(event);
  const tooltipContent: JSX.Element = (
    <div>
      <div style={{ fontWeight: 700, marginBottom: '4px' }}>{event.title}</div>
      <div><b>Location:</b> {event.location || '-'}</div>
      <div><b>Category:</b> {event.categoryText || event.categoryLabel || '-'}</div>
      <div><b>Start:</b> {event.start.toLocaleString()}</div>
      <div><b>End:</b> {event.end.toLocaleString()}</div>
      <div><b>Calendar:</b> {event.calendarLabel || '-'}</div>
      <div><b>Description:</b> {limitText(event.description || '-', 220)}</div>
      {(event.author || event.editor || event.modified) && (
        <div style={{ marginTop: '6px', paddingTop: '4px', borderTop: '1px solid #edebe9' }}>
          {event.author && <div><b>Created By:</b> {event.author}</div>}
          {event.editor && <div><b>Modified By:</b> {event.editor}</div>}
          {event.modified && <div><b>Modified On:</b> {event.modified}</div>}
        </div>
      )}
      {eventUrl && <div style={{ marginTop: '6px', color: '#605e5c' }}>Click title to open event</div>}
    </div>
  );

  const baseTextStyle: React.CSSProperties = compact ? {
    display: 'block',
    fontSize: '11px',
    marginBottom: ganttTitleBesideBar || ganttTrackLabel || ganttRightOfBarLabel ? 0 : 2,
    lineHeight: ganttTitleBesideBar ? '18px' : ganttTrackLabel ? '1.25' : ganttRightOfBarLabel ? '16px' : '1.2'
  } : {
    display: 'block',
    fontSize: '12px',
    lineHeight: '1.2'
  };

  const textStyle: React.CSSProperties = renderAsBar ? {
    ...baseTextStyle,
    color: getContrastColor(event.categoryColor || '#8a8886'),
    background: event.categoryColor || '#8a8886',
    borderTopLeftRadius: barStartsHere ? '3px' : '0px',
    borderBottomLeftRadius: barStartsHere ? '3px' : '0px',
    borderTopRightRadius: barEndsHere ? '3px' : '0px',
    borderBottomRightRadius: barEndsHere ? '3px' : '0px',
    padding: `2px ${barEndsHere ? '6px' : '2px'} 2px ${barStartsHere ? '6px' : '2px'}`,
    width: '100%',
    boxSizing: 'border-box',
    position: 'relative',
    ...(ganttWrapTitleInBar
      ? {
          whiteSpace: 'normal',
          overflow: 'visible',
          wordBreak: 'break-word',
          textOverflow: 'clip',
          minHeight: '100%',
          alignSelf: 'stretch'
        }
      : {
          whiteSpace: 'nowrap',
          overflow: 'hidden',
          textOverflow: 'ellipsis'
        }),
    ...(omitSurfaceZIndex ? {} : { zIndex: 2 })
  } : ganttRightOfBarLabel ? {
    ...baseTextStyle,
    width: '100%',
    minWidth: 0,
    maxWidth: '100%',
    boxSizing: 'border-box',
    fontWeight: 600,
    color: '#323130',
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    textAlign: 'left'
  } : ganttTrackLabel ? {
    ...baseTextStyle,
    width: '100%',
    minWidth: 0,
    maxWidth: '100%',
    boxSizing: 'border-box',
    fontWeight: 600,
    color: '#323130',
    whiteSpace: 'normal',
    overflow: 'visible',
    wordBreak: 'break-word',
    textAlign: 'left'
  } : {
    ...baseTextStyle,
    ...(ganttTitleBesideBar ? { fontWeight: 600 as const } : {})
  };

  const dateRangeTextStyle: React.CSSProperties = compact ? {
    display: 'block',
    fontSize: '10px',
    color: '#605e5c',
    marginLeft: '14px',
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis'
  } : {
    display: 'block',
    fontSize: '11px',
    color: '#605e5c',
    marginLeft: '14px'
  };

  return (
    <TooltipHost content={tooltipContent} closeDelay={250}>
      <div
        style={
          ganttTrackLabel || ganttRightOfBarLabel
            ? { width: '100%', minWidth: 0, boxSizing: 'border-box' }
            : undefined
        }
      >
      {eventUrl ? (
        <button
          type="button"
          onClick={() => {
            const popupWindow = window.open(eventUrl, '_blank', 'noopener,noreferrer');
            if (popupWindow) {
              popupWindow.opener = null;
            }
          }}
          style={{
            ...textStyle,
            textDecoration: 'none',
            border: 'none',
            textAlign: 'left',
            cursor: 'pointer',
            ...(renderAsBar
              ? {}
              : {
                  color: '#0078d4',
                  background: 'transparent',
                  padding: 0,
                  margin: 0
                })
          }}
        >
          {!renderAsBar && !hideCategoryDot && <CategoryDot color={event.categoryColor} />}
          {renderAsBar && !barStartsHere ? '\u00A0' : event.title}
        </button>
      ) : (
          <span style={{ ...textStyle, color: renderAsBar ? getContrastColor(event.categoryColor || '#8a8886') : '#323130' }}>
          {!renderAsBar && !hideCategoryDot && <CategoryDot color={event.categoryColor} />}
            {renderAsBar && !barStartsHere ? '\u00A0' : event.title}
          </span>
        )}
        {showDateRange && !renderAsBar && (
          <span style={dateRangeTextStyle}>
            {event.start.toLocaleString()} - {event.end.toLocaleString()}
          </span>
        )}
      </div>
    </TooltipHost>
  );
};

const CalendarWeekTimeView: React.FC<{
  events: ITimelineItem[];
  referenceDate: Date;
  mode: 'day' | 'week5' | 'week7';
}> = ({ events, referenceDate, mode }) => {
  const isDay: boolean = mode === 'day';
  const isWorkWeek: boolean = mode === 'week5';
  const dayCount: number = isDay ? 1 : isWorkWeek ? 5 : 7;
  const weekStart: Date = isDay ? startOfDay(referenceDate) : startOfWeek(referenceDate, isWorkWeek);
  const days: Date[] = [];

  for (let index = 0; index < dayCount; index++) {
    days.push(addDays(weekStart, index));
  }

  const columnHeaderLabels: string[] = days.map((day: Date, index: number) => {
    if (isDay) {
      return day.toLocaleDateString(undefined, { weekday: 'short', month: 'short', day: 'numeric' });
    }
    const lab: string = isWorkWeek ? WEEKDAY_LABELS[index] : DAY_LABELS[index];
    return `${lab} ${day.getMonth() + 1}/${day.getDate()}`;
  });

  const timedEventsByDay: ITimelineItem[][] = days.map((day: Date) =>
    events.filter((eventItem: ITimelineItem) => {
      if (shouldShowAsOutlookSpanBar(eventItem)) {
        return false;
      }
      if (!eventOccursOnDate(eventItem, day)) {
        return false;
      }
      return clipIntervalToDayGrid(day, eventItem.start, getInclusiveEventEnd(eventItem)) !== null;
    })
  );

  const barPlacements: IWeekBarPlaced[] = computeWeekOutlookBars(days, events);
  const maxLane: number = barPlacements.reduce((acc: number, p: IWeekBarPlaced) => Math.max(acc, p.lane), -1);
  const barRowCount: number = maxLane + 1;

  const multiDayBarRow: React.ReactNode =
    barRowCount > 0 ? (
      <div
        style={{
          display: 'grid',
          gridTemplateColumns: `repeat(${dayCount}, minmax(0, 1fr))`,
          columnGap: '1px',
          rowGap: '3px',
          padding: '4px 0 2px',
          borderBottom: '1px solid #edebe9',
          background: '#faf9f8',
          gridTemplateRows: `repeat(${barRowCount}, minmax(18px, auto))`
        }}
      >
        {barPlacements.map((p: IWeekBarPlaced) => (
          <div
            key={`tg-span-${p.event.id}-l${p.lane}-c${p.startCol}`}
            style={{
              gridColumn: `${p.startCol + 1} / span ${p.span}`,
              gridRow: p.lane + 1,
              margin: '0 1px',
              minHeight: '18px',
              display: 'flex',
              alignItems: 'center',
              overflow: 'hidden'
            }}
          >
            <EventLinkWithTooltip
              event={p.event}
              compact
              showDateRange={false}
              renderAsBar
              barStartsHere
              barEndsHere
            />
          </div>
        ))}
      </div>
    ) : undefined;

  return (
    <CalendarTimeGrid<ITimelineItem>
      days={days}
      columnHeaderLabels={columnHeaderLabels}
      timedEventsByDay={timedEventsByDay}
      getEventStart={(e: ITimelineItem) => e.start}
      getEventEndInclusive={getInclusiveEventEnd}
      renderEvent={(e: ITimelineItem) => (
        <div
          style={{
            height: '100%',
            display: 'flex',
            flexDirection: 'column',
            justifyContent: 'flex-start',
            borderLeft: `4px solid ${e.categoryColor || '#0078d4'}`,
            background: '#ffffff',
            padding: '1px 2px 2px 4px',
            overflow: 'hidden',
            boxSizing: 'border-box'
          }}
        >
          <EventLinkWithTooltip event={e} compact showDateRange={false} hideCategoryDot />
        </div>
      )}
      multiDayBarRow={multiDayBarRow}
    />
  );
};

const GanttChartEventCell: React.FC<{
  eventItem: ITimelineItem;
  bar: { leftPct: number; widthPct: number };
  metrics: IGanttLayoutMetrics;
  laneTopPx: number;
  stripHeightPx: number;
  layoutMode: TGanttCellLayoutMode;
  trackWidthPx: number;
}> = ({ eventItem, bar, metrics, laneTopPx, stripHeightPx, layoutMode, trackWidthPx }) => {
  const color: string = eventItem.categoryColor || '#0078d4';
  const parentRemainPct: number = Math.max(0.01, 100 - bar.leftPct);
  const barWidthAsPctOfParent: number = Math.min(100, (bar.widthPct / parentRemainPct) * 100);
  const barWidthPx: number = trackWidthPx > 0 ? (bar.widthPct / 100) * trackWidthPx : 0;
  const titleInsideBar: boolean =
    layoutMode === 'right' && ganttTitleFitsInsideBar(eventItem.title, metrics.labelLineHeightPx, barWidthPx);
  if (layoutMode === 'below-fallback') {
    return (
      <div
        style={{
          position: 'absolute',
          left: `${bar.leftPct}%`,
          top: `${laneTopPx}px`,
          width: `calc(${parentRemainPct}% - 6px)`,
          minHeight: `${stripHeightPx}px`,
          boxSizing: 'border-box',
          display: 'flex',
          flexDirection: 'column',
          justifyContent: 'flex-start',
          overflow: 'visible'
        }}
      >
        <div
          style={{
            marginBottom: `${metrics.intraLaneGapPx}px`,
            height: `${metrics.barHeightPx}px`,
            width: `${barWidthAsPctOfParent}%`,
            minWidth: '2px',
            flexShrink: 0,
            borderRadius: '4px',
            background: color,
            border: '1px solid rgba(0,0,0,0.08)',
            boxSizing: 'border-box'
          }}
          aria-hidden
        />
        <div style={{ flex: '1 1 auto', minWidth: 0, width: '100%', display: 'block' }}>
          <EventLinkWithTooltip
            event={eventItem}
            compact
            showDateRange={false}
            hideCategoryDot
            ganttTrackLabel
          />
        </div>
      </div>
    );
  }

  if (titleInsideBar) {
    return (
      <div
        style={{
          position: 'absolute',
          left: `${bar.leftPct}%`,
          top: `${laneTopPx}px`,
          width: `${bar.widthPct}%`,
          minWidth: '3px',
          height: `${stripHeightPx}px`,
          boxSizing: 'border-box',
          display: 'flex',
          flexDirection: 'row',
          alignItems: 'center',
          overflow: 'hidden'
        }}
      >
        <div
          style={{
            display: 'flex',
            flexDirection: 'row',
            alignItems: 'center',
            width: '100%',
            height: `${metrics.barHeightPx}px`,
            minWidth: 0,
            borderRadius: '4px',
            background: color,
            border: '1px solid rgba(0,0,0,0.08)',
            boxSizing: 'border-box',
            overflow: 'hidden'
          }}
        >
          <EventLinkWithTooltip
            event={eventItem}
            compact
            showDateRange={false}
            hideCategoryDot
            renderAsBar
            barStartsHere={true}
            barEndsHere={true}
            omitSurfaceZIndex
          />
        </div>
      </div>
    );
  }

  return (
    <div
      style={{
        position: 'absolute',
        left: `${bar.leftPct}%`,
        top: `${laneTopPx}px`,
        width: `calc(${parentRemainPct}% - 6px)`,
        height: `${stripHeightPx}px`,
        boxSizing: 'border-box',
        display: 'flex',
        flexDirection: 'row',
        alignItems: 'center',
        overflow: 'hidden'
      }}
    >
      <div
        style={{
          height: `${metrics.barHeightPx}px`,
          width: `${barWidthAsPctOfParent}%`,
          minWidth: '2px',
          flexShrink: 0,
          borderRadius: '4px',
          background: color,
          border: '1px solid rgba(0,0,0,0.08)',
          boxSizing: 'border-box'
        }}
        aria-hidden
      />
      <div
        style={{
          flex: '1 1 auto',
          minWidth: 0,
          marginLeft: `${metrics.barToLabelGapPx}px`,
          display: 'flex',
          alignItems: 'center'
        }}
      >
        <EventLinkWithTooltip
          event={eventItem}
          compact
          showDateRange={false}
          hideCategoryDot
          ganttRightOfBarLabel
        />
      </div>
    </div>
  );
};

const CategoryDot: React.FC<{ color: string }> = ({ color }) => (
  <span
    style={{
      display: 'inline-block',
      width: '8px',
      height: '8px',
      borderRadius: '50%',
      backgroundColor: color || '#8a8886',
      marginRight: '6px',
      verticalAlign: 'middle'
    }}
  />
);

function buildCalendarTitle(referenceDate: Date, mode: TCalendarMode): string {
  if (mode === 'month') {
    return referenceDate.toLocaleDateString(undefined, { month: 'long', year: 'numeric' });
  }

  if (mode === 'day') {
    return referenceDate.toLocaleDateString(undefined, {
      weekday: 'long',
      month: 'long',
      day: 'numeric',
      year: 'numeric'
    });
  }

  const isWorkWeek: boolean = mode === 'week5';
  const startDate: Date = startOfWeek(referenceDate, isWorkWeek);
  const endDate: Date = addDays(startDate, isWorkWeek ? 4 : 6);
  return `${startDate.toLocaleDateString()} - ${endDate.toLocaleDateString()}`;
}

function eventsForDate(events: ITimelineItem[], day: Date): ITimelineItem[] {
  return events.filter((eventItem: ITimelineItem) => eventOccursOnDate(eventItem, day));
}

function eventOccursOnDate(eventItem: ITimelineItem, day: Date): boolean {
  const dayStart = startOfDay(day);
  const dayEnd = endOfDay(day);
  const eventStart = eventItem.start.getTime();
  const eventEndInclusive = getInclusiveEventEnd(eventItem).getTime();
  return eventStart <= dayEnd.getTime() && eventEndInclusive >= dayStart.getTime();
}

function maxOutlookDate(firstDate: Date, secondDate: Date): Date {
  return firstDate.getTime() >= secondDate.getTime() ? firstDate : secondDate;
}

function minOutlookDate(firstDate: Date, secondDate: Date): Date {
  return firstDate.getTime() <= secondDate.getTime() ? firstDate : secondDate;
}

function shouldShowAsOutlookSpanBar(eventItem: ITimelineItem): boolean {
  if (eventItem.ccIsExpandedRecurrenceInstance) {
    return false;
  }
  if (eventItem.ccMergedDaily) {
    return false;
  }
  const startDay: number = startOfDay(eventItem.start).getTime();
  const endDay: number = startOfDay(getInclusiveEventEnd(eventItem)).getTime();
  return endDay > startDay;
}

function calendarDayColumnInOutlookWeek(weekDates: Date[], day: Date): number {
  const target: number = startOfDay(day).getTime();
  for (let i: number = 0; i < weekDates.length; i++) {
    if (startOfDay(weekDates[i]).getTime() === target) {
      return i;
    }
  }
  return -1;
}

function intervalsOverlapOutlookCols(s1: number, e1: number, s2: number, e2: number): boolean {
  return !(e1 < s2 || e2 < s1);
}

function assignOutlookBarLanes(
  segments: { event: ITimelineItem; startCol: number; span: number }[]
): IWeekBarPlaced[] {
  const sorted: { event: ITimelineItem; startCol: number; span: number }[] = [...segments].sort(
    (a, b) =>
      a.startCol - b.startCol ||
      b.span - a.span ||
      a.event.start.getTime() - b.event.start.getTime()
  );
  const laneIntervals: { startCol: number; endCol: number }[][] = [];
  const placed: IWeekBarPlaced[] = [];

  for (const seg of sorted) {
    const s: number = seg.startCol;
    const e: number = seg.startCol + seg.span - 1;
    let laneIdx: number = 0;

    for (; laneIdx < laneIntervals.length; laneIdx++) {
      const conflicts: boolean = laneIntervals[laneIdx].some((interval: { startCol: number; endCol: number }) =>
        intervalsOverlapOutlookCols(s, e, interval.startCol, interval.endCol)
      );
      if (!conflicts) {
        break;
      }
    }

    if (laneIdx === laneIntervals.length) {
      laneIntervals.push([]);
    }
    laneIntervals[laneIdx].push({ startCol: s, endCol: e });
    placed.push({ event: seg.event, startCol: seg.startCol, span: seg.span, lane: laneIdx });
  }

  return placed;
}

function computeWeekOutlookBars(weekDates: Date[], allEvents: ITimelineItem[]): IWeekBarPlaced[] {
  if (weekDates.length === 0) {
    return [];
  }

  const weekStart: Date = startOfDay(weekDates[0]);
  const weekEnd: Date = endOfDay(weekDates[weekDates.length - 1]);

  const raw: { event: ITimelineItem; startCol: number; span: number }[] = [];

  allEvents.forEach((eventItem: ITimelineItem) => {
    if (!shouldShowAsOutlookSpanBar(eventItem)) {
      return;
    }
    if (eventItem.end.getTime() < weekStart.getTime() || eventItem.start.getTime() > weekEnd.getTime()) {
      return;
    }

    const clipStart: Date = maxOutlookDate(startOfDay(eventItem.start), weekStart);
    const clipEndDay: Date = minOutlookDate(
      startOfDay(getInclusiveEventEnd(eventItem)),
      startOfDay(weekDates[weekDates.length - 1])
    );
    const startCol: number = calendarDayColumnInOutlookWeek(weekDates, clipStart);
    const endCol: number = calendarDayColumnInOutlookWeek(weekDates, clipEndDay);

    if (startCol < 0 || endCol < 0) {
      return;
    }

    raw.push({ event: eventItem, startCol, span: endCol - startCol + 1 });
  });

  return assignOutlookBarLanes(raw);
}

function expandMergedDailyEventsForDisplay(events: ITimelineItem[]): ITimelineItem[] {
  const expandedEvents: ITimelineItem[] = [];

  events.forEach((eventItem: ITimelineItem) => {
    if (!eventItem.ccMergedDaily) {
      expandedEvents.push(eventItem);
      return;
    }

    const rangeStart = startOfDay(eventItem.start);
    const rangeEnd = startOfDay(getInclusiveEventEnd(eventItem));
    const startOffsetMs = (eventItem.ccDailyStartOffsetMs || 0);
    let durationMs = (eventItem.ccDailyDurationMs || 0);
    if (durationMs <= 0) {
      durationMs = 60 * 60 * 1000;
    }

    for (let currentDay = new Date(rangeStart.getTime()); currentDay.getTime() <= rangeEnd.getTime(); currentDay = addDays(currentDay, 1)) {
      const dayStart = currentDay.getTime() + startOffsetMs;
      const dayEnd = dayStart + durationMs;
      expandedEvents.push({
        ...eventItem,
        id: `${eventItem.id}__${currentDay.toISOString().substring(0, 10)}`,
        start: new Date(dayStart),
        end: new Date(dayEnd),
        ccMergedDaily: false
      });
    }
  });

  expandedEvents.sort((firstEvent: ITimelineItem, secondEvent: ITimelineItem) => {
    const timeDiff = firstEvent.start.getTime() - secondEvent.start.getTime();
    if (timeDiff !== 0) {
      return timeDiff;
    }

    return firstEvent.title.localeCompare(secondEvent.title);
  });

  return expandedEvents;
}

function getInclusiveEventEnd(eventItem: ITimelineItem): Date {
  const endDate = new Date(eventItem.end.getTime());
  const isMidnight = endDate.getHours() === 0 && endDate.getMinutes() === 0 && endDate.getSeconds() === 0 && endDate.getMilliseconds() === 0;
  if (endDate.getTime() > eventItem.start.getTime() && isMidnight) {
    endDate.setMilliseconds(endDate.getMilliseconds() - 1);
  }
  return endDate;
}

function getContrastColor(hexOrCssColor: string): string {
  if (!hexOrCssColor || hexOrCssColor.indexOf('#') !== 0) {
    return '#ffffff';
  }

  const hex = hexOrCssColor.replace('#', '');
  if (hex.length !== 6) {
    return '#ffffff';
  }

  const red = parseInt(hex.substring(0, 2), 16);
  const green = parseInt(hex.substring(2, 4), 16);
  const blue = parseInt(hex.substring(4, 6), 16);
  const brightness = ((red * 299) + (green * 587) + (blue * 114)) / 1000;
  return brightness > 155 ? '#202020' : '#ffffff';
}

function extractEventUrlFromRaw(item: any): string {
  if (item.calEventWebLink) {
    return String(item.calEventWebLink);
  }

  if (item.encodedAbsUrl && item.spId != null) {
    const encodedAbsUrl: string = String(item.encodedAbsUrl);
    const slashIndex: number = encodedAbsUrl.lastIndexOf('/');
    if (slashIndex > -1) {
      const itemUrl: string = encodedAbsUrl.substring(0, slashIndex);
      const formsSegment: string = String(item.objType || '') === '1' ? '/Forms' : '';
      return `${itemUrl}${formsSegment}/DispForm.aspx?ID=${encodeURIComponent(String(item.spId))}`;
    }
  }

  if (item.encodedAbsUrl) {
    return String(item.encodedAbsUrl);
  }

  return '';
}

function getEventUrl(eventItem: ITimelineItem): string {
  if (eventItem.eventUrl) {
    return eventItem.eventUrl;
  }

  if (eventItem.encodedAbsUrl && eventItem.spId) {
    const slashIndex: number = eventItem.encodedAbsUrl.lastIndexOf('/');
    if (slashIndex > -1) {
      const itemUrl: string = eventItem.encodedAbsUrl.substring(0, slashIndex);
      const formsSegment: string = eventItem.objType === '1' ? '/Forms' : '';
      return `${itemUrl}${formsSegment}/DispForm.aspx?ID=${encodeURIComponent(eventItem.spId)}`;
    }
  }

  return '';
}

function limitText(value: string, maxLength: number): string {
  if (!value || value.length <= maxLength) {
    return value || '';
  }

  return `${value.substring(0, maxLength)}...`;
}

function stripHtml(value: string): string {
  if (!value) {
    return '';
  }

  const element: HTMLDivElement = document.createElement('div');
  element.innerHTML = value;
  return (element.textContent || element.innerText || '').trim();
}

function startOfDay(date: Date): Date {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0, 0);
}

function endOfDay(date: Date): Date {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59, 999);
}

function startOfWeek(date: Date, mondayFirst: boolean): Date {
  const current: Date = startOfDay(date);
  const dayNumber: number = current.getDay();
  const offset: number = mondayFirst ? (dayNumber + 6) % 7 : dayNumber;
  return addDays(current, -offset);
}

function addDays(date: Date, days: number): Date {
  const nextDate: Date = new Date(date.getTime());
  nextDate.setDate(nextDate.getDate() + days);
  return nextDate;
}

function addMonths(date: Date, months: number): Date {
  return new Date(date.getFullYear(), date.getMonth() + months, date.getDate(), 0, 0, 0, 0);
}

function extractCategoryColor(stylesText: string): string {
  if (!stylesText) {
    return '';
  }

  const bgMatch: RegExpMatchArray = stylesText.match(/background-color\s*:\s*([^;]+)/i);
  if (bgMatch && bgMatch[1]) {
    return bgMatch[1].trim();
  }

  const borderMatch: RegExpMatchArray = stylesText.match(/border-color\s*:\s*([^;]+)/i);
  if (borderMatch && borderMatch[1]) {
    return borderMatch[1].trim();
  }

  return '';
}

function resolveEventCategory(
  item: any,
  categoryMetaByKey: Map<string, ICategoryMeta>,
  ensureValidClassName: (value: string) => string
): ICategoryMeta {
  const rawClassName: string = String(item.className || '');
  const classTokens: string[] = rawClassName.split(/\s+/).filter((value: string) => !!value && value !== 'vis-selected');

  let classMatch: string = '';
  classTokens.some((token: string) => {
    if (categoryMetaByKey.has(token)) {
      classMatch = token;
      return true;
    }

    return false;
  });

  if (classMatch) {
    const matchedCategory: ICategoryMeta = categoryMetaByKey.get(classMatch);
    if (matchedCategory) {
      return matchedCategory;
    }
  }

  const rawCategoryText: string = String(item.Category || item.category || '').split(',')[0].trim();
  if (rawCategoryText) {
    const normalizedCategoryKey: string = ensureValidClassName(rawCategoryText);
    if (categoryMetaByKey.has(normalizedCategoryKey)) {
      const matchedCategory: ICategoryMeta = categoryMetaByKey.get(normalizedCategoryKey);
      if (matchedCategory) {
        return matchedCategory;
      }
    }

    return {
      key: normalizedCategoryKey || rawCategoryText,
      label: rawCategoryText,
      color: '#8a8886'
    };
  }

  if (classTokens.length > 0) {
    const fallbackKey: string = classTokens[0];
    if (categoryMetaByKey.has(fallbackKey)) {
      const fallbackCategory: ICategoryMeta = categoryMetaByKey.get(fallbackKey);
      if (fallbackCategory) {
        return fallbackCategory;
      }
    }

    return {
      key: fallbackKey,
      label: fallbackKey,
      color: '#8a8886'
    };
  }

  return {
    key: '',
    label: '',
    color: '#8a8886'
  };
}

function resolveEventCalendar(item: any): ICalendarMeta {
  const sourceObj: any = item && item.sourceObj ? item.sourceObj : {};
  if (sourceObj.siteUrl && sourceObj.list) {
    const displayLabel = String(sourceObj.displayName || sourceObj.listName || sourceObj.list || 'SharePoint Calendar');
    return {
      key: `sp|${sourceObj.siteUrl}|${sourceObj.list}`,
      label: displayLabel
    };
  }

  if (sourceObj.resource) {
    const personaMail: string = sourceObj.persona && sourceObj.persona.length > 0 ? String(sourceObj.persona[0].mail || '') : '';
    const label: string = String(sourceObj.displayName || sourceObj.resourceName || sourceObj.name || sourceObj.resource || personaMail || 'Outlook Calendar');
    return {
      key: `graph|${sourceObj.resource}|${personaMail}`,
      label
    };
  }

  const fallbackLabel: string = String(sourceObj.listName || sourceObj.group || sourceObj.name || 'Calendar');
  return {
    key: `name|${fallbackLabel}`,
    label: fallbackLabel
  };
}

function resolveEventStaffSection(item: any, staffSectionMetaByKey: Map<string, IStaffSectionMeta>): IStaffSectionMeta {
  const groupKey: string = item && item.group ? String(item.group) : '';
  if (!groupKey) {
    return {
      key: '',
      label: ''
    };
  }

  const existingMeta = staffSectionMetaByKey.get(groupKey);
  if (existingMeta) {
    return existingMeta;
  }

  return {
    key: groupKey,
    label: groupKey
  };
}

export default TimelineCalendarTabbed;

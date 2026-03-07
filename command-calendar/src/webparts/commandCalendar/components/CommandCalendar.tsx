import * as React from 'react';
import {
  DefaultButton,
  DirectionalHint,
  Dropdown,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  Pivot,
  PivotItem,
  Spinner,
  SpinnerSize,
  TooltipHost
} from '@fluentui/react';
import styles from './CommandCalendar.module.scss';
import type { ICommandCalendarProps } from './ICommandCalendarProps';
import { ICalendarEvent, ICalendarSourceConfig } from '../models/CalendarModels';
import { CalendarDataService, parseCalendarSources } from '../services/CalendarDataService';

const DAY_LABELS: string[] = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const GANTT_ZOOM_ORDER: ('day' | 'week' | 'month')[] = ['day', 'week', 'month'];

const CommandCalendar: React.FC<ICommandCalendarProps> = (props: ICommandCalendarProps) => {
  const [events, setEvents] = React.useState<ICalendarEvent[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [errorMessage, setErrorMessage] = React.useState<string | undefined>();
  const [selectedCategories, setSelectedCategories] = React.useState<string[]>([]);
  const [activeView, setActiveView] = React.useState<string>('gantt');
  const [calendarAnchor, setCalendarAnchor] = React.useState<Date>(startOfMonth(new Date()));
  const [reloadToken, setReloadToken] = React.useState<number>(0);
  const [ganttZoomIndex, setGanttZoomIndex] = React.useState<number>(1);

  const dataService: CalendarDataService = React.useMemo(
    () => new CalendarDataService(props.spHttpClient),
    [props.spHttpClient]
  );

  const parsedSources = React.useMemo(
    () => parseCalendarSources(props.calendarSources || ''),
    [props.calendarSources]
  );

  React.useEffect(() => {
    let cancelled: boolean = false;

    const loadEvents = async (): Promise<void> => {
      if (parsedSources.sources.length === 0) {
        setEvents([]);
        setErrorMessage(undefined);
        return;
      }

      setIsLoading(true);
      setErrorMessage(undefined);

      try {
        const today: Date = new Date();
        const rangeStart: Date = addDays(startOfDay(today), -Math.max(0, props.lookBackDays));
        const rangeEnd: Date = addDays(endOfDay(today), Math.max(120, props.lookAheadDays));
        const loadedEvents: ICalendarEvent[] = await dataService.getEvents(
          parsedSources.sources,
          rangeStart,
          rangeEnd
        );
        if (!cancelled) {
          setEvents(loadedEvents);
        }
      } catch (error) {
        if (!cancelled) {
          setEvents([]);
          setErrorMessage(
            error instanceof Error ? error.message : 'An unknown error occurred while loading events.'
          );
        }
      } finally {
        if (!cancelled) {
          setIsLoading(false);
        }
      }
    };

    loadEvents().catch(() => undefined);

    return () => {
      cancelled = true;
    };
  }, [
    dataService,
    parsedSources.sources,
    props.lookAheadDays,
    props.lookBackDays,
    reloadToken
  ]);

  const categoryOptions: IDropdownOption[] = React.useMemo(() => {
    const categoryMap: Record<string, string> = {};
    events.forEach((event: ICalendarEvent) => {
      event.categories.forEach((category: string) => {
        const normalized: string = category.toLocaleLowerCase();
        if (!categoryMap[normalized]) {
          categoryMap[normalized] = category;
        }
      });
    });

    return Object.keys(categoryMap)
      .sort((a: string, b: string) => categoryMap[a].localeCompare(categoryMap[b]))
      .map((key: string) => ({
        key: categoryMap[key],
        text: categoryMap[key]
      }));
  }, [events]);

  const filteredEvents: ICalendarEvent[] = React.useMemo(() => {
    if (selectedCategories.length === 0) {
      return events;
    }

    const selectedLookup: Record<string, boolean> = {};
    selectedCategories.forEach((selectedCategory: string) => {
      selectedLookup[selectedCategory.toLocaleLowerCase()] = true;
    });

    return events.filter((event: ICalendarEvent) =>
      event.categories.some((eventCategory: string) => selectedLookup[eventCategory.toLocaleLowerCase()])
    );
  }, [events, selectedCategories]);

  const zoomKey: 'day' | 'week' | 'month' = GANTT_ZOOM_ORDER[ganttZoomIndex];
  const baseSpanDays: number = Math.max(30, props.lookAheadDays + Math.max(0, props.lookBackDays));
  const timelineStart: Date = React.useMemo(
    () => addDays(startOfDay(new Date()), -Math.max(0, props.lookBackDays)),
    [props.lookBackDays]
  );
  const timelineDays: number = getTimelineSpanDays(zoomKey, baseSpanDays);
  const timelineEnd: Date = addDays(timelineStart, timelineDays);
  const ganttTicks: IGanttTick[] = React.useMemo(
    () => buildGanttTicks(timelineStart, timelineDays, zoomKey),
    [timelineDays, timelineStart, zoomKey]
  );

  const monthGridDates: Date[] = React.useMemo(() => {
    const monthStart: Date = startOfMonth(calendarAnchor);
    const gridStart: Date = addDays(monthStart, -monthStart.getDay());
    return Array.from({ length: 42 }, (_: unknown, index: number) => addDays(gridStart, index));
  }, [calendarAnchor]);

  const sourceWarnings: string[] = parsedSources.errors;
  const hasNoSourcesConfigured: boolean = parsedSources.sources.length === 0;

  const onCategoryChanged = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (!option) {
      return;
    }

    const optionKey: string = String(option.key);
    setSelectedCategories((previousSelections: string[]) => {
      if (option.selected) {
        return previousSelections.indexOf(optionKey) > -1
          ? previousSelections
          : previousSelections.concat(optionKey);
      }

      return previousSelections.filter((selection: string) => selection !== optionKey);
    });
  };

  const onGanttWheel = (ev: React.WheelEvent<HTMLDivElement>): void => {
    if (Math.abs(ev.deltaY) < 4) {
      return;
    }

    ev.preventDefault();
    setGanttZoomIndex((currentIndex: number) => {
      if (ev.deltaY < 0) {
        return Math.max(0, currentIndex - 1);
      }

      return Math.min(GANTT_ZOOM_ORDER.length - 1, currentIndex + 1);
    });
  };

  const onEventClick = (event: ICalendarEvent): void => {
    if (!event.itemUrl) {
      return;
    }

    window.open(event.itemUrl, '_blank');
  };

  const renderGanttView = (): JSX.Element => {
    const visibleEvents: ICalendarEvent[] = filteredEvents.filter((event: ICalendarEvent) =>
      rangesOverlap(event.start, event.end, timelineStart, timelineEnd)
    );

    if (visibleEvents.length === 0) {
      return <div className={styles.emptyState}>No events in the current timeline window.</div>;
    }

    return (
      <div className={styles.ganttView} onWheel={onGanttWheel}>
        <div className={styles.ganttToolbar}>
          <div className={styles.ganttRangeSummary}>
            {formatDate(timelineStart)} - {formatDate(timelineEnd)}
          </div>
          <div className={styles.ganttZoomLegend}>
            Scroll over the chart to zoom ({zoomKey})
          </div>
          <div className={styles.ganttZoomButtons}>
            {GANTT_ZOOM_ORDER.map((option: 'day' | 'week' | 'month', index: number) => (
              <button
                type="button"
                key={option}
                className={`${styles.ganttZoomButton} ${
                  index === ganttZoomIndex ? styles.ganttZoomButtonActive : ''
                }`}
                onClick={() => setGanttZoomIndex(index)}
              >
                {option}
              </button>
            ))}
          </div>
        </div>
        <div className={styles.ganttHeader}>
          <div className={styles.ganttEventColumn}>Event</div>
          <div className={styles.ganttTimelineColumn}>Timeline</div>
        </div>
        <div className={styles.ganttAxisRow}>
          <div />
          <div className={styles.ganttAxisTrack}>
            {ganttTicks.map((tick: IGanttTick) => (
              <div key={tick.key} className={styles.ganttTick} style={{ left: `${tick.leftPercent}%` }}>
                <span className={styles.ganttTickLabel}>{tick.label}</span>
              </div>
            ))}
          </div>
        </div>
        {visibleEvents.map((event: ICalendarEvent) => {
          const clampedStart: Date = maxDate(event.start, timelineStart);
          const clampedEnd: Date = minDate(event.end, timelineEnd);
          const startOffsetDays: number = dateDiffInDays(timelineStart, clampedStart);
          const durationDays: number = Math.max(1, dateDiffInDays(clampedStart, clampedEnd) + 1);
          const leftPercent: number = (startOffsetDays / timelineDays) * 100;
          const widthPercent: number = Math.max((durationDays / timelineDays) * 100, 1);
          const tooltipContent: JSX.Element = renderEventTooltip(event);

          return (
            <div className={styles.ganttRow} key={event.id}>
              <div className={styles.ganttEventMeta}>
                <span className={styles.eventTitle}>{event.title}</span>
                <span className={styles.sourcePill} style={{ borderColor: event.sourceColor }}>
                  {event.sourceName}
                </span>
                <span className={styles.eventTime}>{formatDateRange(event.start, event.end)}</span>
              </div>
              <div className={styles.ganttTrack}>
                {event.isRecurringInstance ? (
                  <TooltipHost
                    content={tooltipContent}
                    directionalHint={DirectionalHint.bottomCenter}
                  >
                    <button
                      type="button"
                      className={styles.ganttDotButton}
                      onClick={() => onEventClick(event)}
                      style={{
                        left: `${leftPercent}%`,
                        borderColor: event.sourceColor,
                        backgroundColor: event.sourceColor
                      }}
                      title={event.title}
                    />
                  </TooltipHost>
                ) : (
                  <TooltipHost
                    content={tooltipContent}
                    directionalHint={DirectionalHint.bottomCenter}
                  >
                    <button
                      type="button"
                      className={styles.ganttBarButton}
                      onClick={() => onEventClick(event)}
                      style={{
                        left: `${leftPercent}%`,
                        width: `${widthPercent}%`,
                        backgroundColor: event.sourceColor
                      }}
                      title={event.title}
                    />
                  </TooltipHost>
                )}
              </div>
            </div>
          );
        })}
      </div>
    );
  };

  const renderCalendarView = (): JSX.Element => {
    return (
      <div className={styles.calendarView}>
        <div className={styles.calendarHeader}>
          <div className={styles.calendarTitle}>
            {calendarAnchor.toLocaleString(undefined, { month: 'long', year: 'numeric' })}
          </div>
          <div className={styles.calendarActions}>
            <DefaultButton
              text="Previous"
              onClick={() => setCalendarAnchor(addMonths(calendarAnchor, -1))}
            />
            <DefaultButton text="Today" onClick={() => setCalendarAnchor(startOfMonth(new Date()))} />
            <DefaultButton text="Next" onClick={() => setCalendarAnchor(addMonths(calendarAnchor, 1))} />
          </div>
        </div>
        <div className={styles.calendarGrid}>
          {DAY_LABELS.map((dayLabel: string) => (
            <div key={dayLabel} className={`${styles.calendarCell} ${styles.calendarDayLabel}`}>
              {dayLabel}
            </div>
          ))}
          {monthGridDates.map((gridDate: Date) => {
            const dayEvents: ICalendarEvent[] = filteredEvents
              .filter((event: ICalendarEvent) =>
                rangesOverlap(event.start, event.end, startOfDay(gridDate), endOfDay(gridDate))
              )
              .sort((a: ICalendarEvent, b: ICalendarEvent) => a.start.getTime() - b.start.getTime());

            const isCurrentMonth: boolean = gridDate.getMonth() === calendarAnchor.getMonth();
            return (
              <div
                key={gridDate.toISOString()}
                className={`${styles.calendarCell} ${isCurrentMonth ? '' : styles.calendarCellMuted}`}
              >
                <div className={styles.calendarDate}>{gridDate.getDate()}</div>
                <div className={styles.calendarItems}>
                  {dayEvents.slice(0, 3).map((event: ICalendarEvent) => (
                    <TooltipHost
                      key={`${event.id}-${gridDate.toISOString()}`}
                      content={renderEventTooltip(event)}
                      directionalHint={DirectionalHint.bottomLeftEdge}
                    >
                      <button
                        type="button"
                        className={styles.calendarItemButton}
                        onClick={() => onEventClick(event)}
                        title={event.title}
                      >
                        <span
                          className={styles.calendarItemDot}
                          style={{ backgroundColor: event.sourceColor }}
                        />
                        <span className={styles.calendarItemText}>{event.title}</span>
                      </button>
                    </TooltipHost>
                  ))}
                  {dayEvents.length > 3 && (
                    <div className={styles.moreItems}>+{dayEvents.length - 3} more</div>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    );
  };

  const renderHorizonView = (): JSX.Element => {
    const horizons: number[] = [30, 60, 90, 120];
    const today: Date = startOfDay(new Date());

    return (
      <div className={styles.quadView}>
        {horizons.map((days: number) => {
          const horizonEnd: Date = addDays(endOfDay(today), days);
          const horizonEvents: ICalendarEvent[] = filteredEvents
            .filter((event: ICalendarEvent) =>
              rangesOverlap(event.start, event.end, today, horizonEnd)
            )
            .sort((a: ICalendarEvent, b: ICalendarEvent) => a.start.getTime() - b.start.getTime());

          return (
            <div className={styles.quadCard} key={days}>
              <div className={styles.quadHeader}>Next {days} Days</div>
              <div className={styles.quadCount}>{horizonEvents.length}</div>
              <div className={styles.quadList}>
                {horizonEvents.slice(0, 5).map((event: ICalendarEvent) => (
                  <div className={styles.quadItem} key={`${days}-${event.id}`}>
                    <span className={styles.quadItemTitle}>{event.title}</span>
                    <span className={styles.quadItemMeta}>
                      {formatDateRange(event.start, event.end)} | {event.sourceName}
                    </span>
                  </div>
                ))}
                {horizonEvents.length === 0 && (
                  <div className={styles.emptyInline}>No events in this horizon.</div>
                )}
              </div>
            </div>
          );
        })}
      </div>
    );
  };

  return (
    <section className={styles.commandCalendar}>
      <div className={styles.headerArea}>
        <div>
          <h2 className={styles.title}>Command Calendar</h2>
          {props.description && <div className={styles.subtitle}>{props.description}</div>}
        </div>
        <DefaultButton text="Refresh" onClick={() => setReloadToken((token: number) => token + 1)} />
      </div>

      {sourceWarnings.length > 0 && (
        <MessageBar messageBarType={MessageBarType.warning}>
          {sourceWarnings.join(' ')}
        </MessageBar>
      )}

      {errorMessage && (
        <MessageBar messageBarType={MessageBarType.error}>
          {errorMessage}
        </MessageBar>
      )}

      {hasNoSourcesConfigured && (
        <MessageBar messageBarType={MessageBarType.info}>
          Configure one or more calendar sources in the property pane using:
          <br />
          siteUrl|calendarListTitle|Display Name|#Color
        </MessageBar>
      )}

      <div className={styles.sourcesLegend}>
        {parsedSources.sources.map((source: ICalendarSourceConfig) => (
          <span className={styles.sourceLegendItem} key={source.key}>
            <span className={styles.sourceLegendDot} style={{ backgroundColor: source.color }} />
            {source.displayName}
          </span>
        ))}
      </div>

      <div className={styles.filterBar}>
        <Dropdown
          label="Filter categories (applies to all views)"
          placeholder="Select one or more categories"
          multiSelect
          selectedKeys={selectedCategories}
          options={categoryOptions}
          onChange={onCategoryChanged}
          className={styles.categoryDropdown}
        />
        <DefaultButton
          text="Clear filters"
          onClick={() => setSelectedCategories([])}
          disabled={selectedCategories.length === 0}
        />
      </div>

      {isLoading ? (
        <Spinner size={SpinnerSize.large} label="Loading calendar events..." />
      ) : (
        <Pivot
          selectedKey={activeView}
          onLinkClick={(item?: PivotItem) => setActiveView(item?.props.itemKey || 'gantt')}
        >
          <PivotItem headerText="Gantt View" itemKey="gantt">
            {renderGanttView()}
          </PivotItem>
          <PivotItem headerText="Calendar View" itemKey="calendar">
            {renderCalendarView()}
          </PivotItem>
          <PivotItem headerText="30/60/90/120" itemKey="horizon">
            {renderHorizonView()}
          </PivotItem>
        </Pivot>
      )}
    </section>
  );
};

interface IGanttTick {
  key: string;
  leftPercent: number;
  label: string;
}

function getTimelineSpanDays(zoomKey: 'day' | 'week' | 'month', baseSpanDays: number): number {
  if (zoomKey === 'day') {
    return Math.max(28, Math.floor(baseSpanDays * 0.6));
  }

  if (zoomKey === 'month') {
    return Math.max(180, Math.floor(baseSpanDays * 2));
  }

  return Math.max(90, baseSpanDays);
}

function buildGanttTicks(
  timelineStart: Date,
  timelineDays: number,
  zoomKey: 'day' | 'week' | 'month'
): IGanttTick[] {
  if (zoomKey === 'month') {
    return buildMonthTicks(timelineStart, timelineDays);
  }

  const stepDays: number = zoomKey === 'week' ? 7 : getDayStep(timelineDays);
  const ticks: IGanttTick[] = [];

  for (let offset = 0; offset <= timelineDays; offset += stepDays) {
    const tickDate: Date = addDays(timelineStart, offset);
    ticks.push({
      key: `${tickDate.toISOString()}-${offset}`,
      leftPercent: (offset / timelineDays) * 100,
      label:
        zoomKey === 'week'
          ? `Wk ${getWeekNumber(tickDate)}`
          : tickDate.toLocaleDateString(undefined, { month: 'short', day: 'numeric' })
    });
  }

  return ticks;
}

function buildMonthTicks(timelineStart: Date, timelineDays: number): IGanttTick[] {
  const ticks: IGanttTick[] = [];
  const timelineEnd: Date = addDays(timelineStart, timelineDays);
  let cursor: Date = new Date(timelineStart.getFullYear(), timelineStart.getMonth(), 1);

  if (cursor.getTime() < timelineStart.getTime()) {
    cursor = addMonths(cursor, 1);
  }

  while (cursor.getTime() <= timelineEnd.getTime()) {
    const offsetDays: number = dateDiffInDays(timelineStart, cursor);
    ticks.push({
      key: cursor.toISOString(),
      leftPercent: Math.max(0, (offsetDays / timelineDays) * 100),
      label: cursor.toLocaleDateString(undefined, { month: 'short', year: '2-digit' })
    });
    cursor = addMonths(cursor, 1);
  }

  if (ticks.length === 0) {
    ticks.push({
      key: timelineStart.toISOString(),
      leftPercent: 0,
      label: timelineStart.toLocaleDateString(undefined, { month: 'short', day: 'numeric' })
    });
  }

  return ticks;
}

function getDayStep(timelineDays: number): number {
  if (timelineDays <= 35) {
    return 1;
  }
  if (timelineDays <= 70) {
    return 2;
  }
  if (timelineDays <= 120) {
    return 5;
  }

  return 10;
}

function getWeekNumber(date: Date): number {
  const utcDate: Date = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum: number = utcDate.getUTCDay() || 7;
  utcDate.setUTCDate(utcDate.getUTCDate() + 4 - dayNum);
  const yearStart: Date = new Date(Date.UTC(utcDate.getUTCFullYear(), 0, 1));
  return Math.ceil((((utcDate.getTime() - yearStart.getTime()) / 86400000) + 1) / 7);
}

function renderEventTooltip(event: ICalendarEvent): JSX.Element {
  return (
    <div className={styles.eventTooltip}>
      <div className={styles.tooltipTitle}>{event.title}</div>
      <div><strong>Category:</strong> {event.categories.join(', ')}</div>
      <div><strong>Start:</strong> {formatDateTime(event.start, event.isAllDay)}</div>
      <div><strong>End:</strong> {formatDateTime(event.end, event.isAllDay)}</div>
      <div><strong>Description:</strong> {event.description || 'N/A'}</div>
      {event.itemUrl && <div className={styles.tooltipHint}>Click to open item</div>}
    </div>
  );
}

function addDays(date: Date, days: number): Date {
  const nextDate: Date = new Date(date.getTime());
  nextDate.setDate(nextDate.getDate() + days);
  return nextDate;
}

function addMonths(date: Date, months: number): Date {
  const nextDate: Date = new Date(date.getTime());
  nextDate.setMonth(nextDate.getMonth() + months);
  return startOfMonth(nextDate);
}

function startOfDay(date: Date): Date {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0, 0);
}

function endOfDay(date: Date): Date {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59, 999);
}

function startOfMonth(date: Date): Date {
  return new Date(date.getFullYear(), date.getMonth(), 1, 0, 0, 0, 0);
}

function dateDiffInDays(startDate: Date, endDate: Date): number {
  const millisecondsPerDay: number = 24 * 60 * 60 * 1000;
  const utcStart: number = Date.UTC(
    startDate.getFullYear(),
    startDate.getMonth(),
    startDate.getDate()
  );
  const utcEnd: number = Date.UTC(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
  return Math.floor((utcEnd - utcStart) / millisecondsPerDay);
}

function rangesOverlap(
  firstStart: Date,
  firstEnd: Date,
  secondStart: Date,
  secondEnd: Date
): boolean {
  return firstStart.getTime() <= secondEnd.getTime() && secondStart.getTime() <= firstEnd.getTime();
}

function minDate(firstDate: Date, secondDate: Date): Date {
  return firstDate.getTime() <= secondDate.getTime() ? firstDate : secondDate;
}

function maxDate(firstDate: Date, secondDate: Date): Date {
  return firstDate.getTime() >= secondDate.getTime() ? firstDate : secondDate;
}

function formatDate(date: Date): string {
  return date.toLocaleDateString(undefined, {
    month: 'short',
    day: 'numeric',
    year: 'numeric'
  });
}

function formatDateTime(date: Date, isAllDay: boolean): string {
  if (isAllDay) {
    return `${formatDate(date)} (All day)`;
  }

  return date.toLocaleString(undefined, {
    month: 'short',
    day: 'numeric',
    year: 'numeric',
    hour: 'numeric',
    minute: '2-digit'
  });
}

function formatDateRange(startDate: Date, endDate: Date): string {
  const startText: string = formatDate(startDate);
  const endText: string = formatDate(endDate);
  return startText === endText ? startText : `${startText} - ${endText}`;
}

export default CommandCalendar;

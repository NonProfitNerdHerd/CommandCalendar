import * as React from 'react';
import {
  DefaultButton,
  Dropdown,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  Pivot,
  PivotItem,
  Spinner,
  SpinnerSize
} from '@fluentui/react';
import styles from './CommandCalendar.module.scss';
import type { ICommandCalendarProps } from './ICommandCalendarProps';
import { ICalendarEvent, ICalendarSourceConfig } from '../models/CalendarModels';
import { CalendarDataService, parseCalendarSources } from '../services/CalendarDataService';

const DAY_LABELS: string[] = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

const CommandCalendar: React.FC<ICommandCalendarProps> = (props: ICommandCalendarProps) => {
  const [events, setEvents] = React.useState<ICalendarEvent[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [errorMessage, setErrorMessage] = React.useState<string | undefined>();
  const [selectedCategories, setSelectedCategories] = React.useState<string[]>([]);
  const [activeView, setActiveView] = React.useState<string>('gantt');
  const [calendarAnchor, setCalendarAnchor] = React.useState<Date>(startOfMonth(new Date()));
  const [reloadToken, setReloadToken] = React.useState<number>(0);

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

  const timelineStart: Date = React.useMemo(() => startOfDay(new Date()), []);
  const timelineDays: number = Math.max(120, props.lookAheadDays);
  const timelineEnd: Date = addDays(timelineStart, timelineDays);

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

  const renderGanttView = (): JSX.Element => {
    const visibleEvents: ICalendarEvent[] = filteredEvents.filter((event: ICalendarEvent) =>
      rangesOverlap(event.start, event.end, timelineStart, timelineEnd)
    );

    if (visibleEvents.length === 0) {
      return <div className={styles.emptyState}>No events in the current timeline window.</div>;
    }

    return (
      <div className={styles.ganttView}>
        <div className={styles.ganttHeader}>
          <div className={styles.ganttEventColumn}>Event</div>
          <div className={styles.ganttTimelineColumn}>
            {formatDate(timelineStart)} - {formatDate(timelineEnd)}
          </div>
        </div>
        {visibleEvents.map((event: ICalendarEvent) => {
          const clampedStart: Date = maxDate(event.start, timelineStart);
          const clampedEnd: Date = minDate(event.end, timelineEnd);
          const startOffsetDays: number = dateDiffInDays(timelineStart, clampedStart);
          const durationDays: number = Math.max(1, dateDiffInDays(clampedStart, clampedEnd) + 1);
          const leftPercent: number = (startOffsetDays / timelineDays) * 100;
          const widthPercent: number = Math.max((durationDays / timelineDays) * 100, 1);

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
                <div
                  className={styles.ganttBar}
                  style={{
                    left: `${leftPercent}%`,
                    width: `${widthPercent}%`,
                    backgroundColor: event.sourceColor
                  }}
                />
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
                    <div className={styles.calendarItem} key={`${event.id}-${gridDate.toISOString()}`}>
                      <span className={styles.calendarItemDot} style={{ backgroundColor: event.sourceColor }} />
                      <span className={styles.calendarItemText}>{event.title}</span>
                    </div>
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

function formatDateRange(startDate: Date, endDate: Date): string {
  const startText: string = formatDate(startDate);
  const endText: string = formatDate(endDate);
  return startText === endText ? startText : `${startText} - ${endText}`;
}

export default CommandCalendar;

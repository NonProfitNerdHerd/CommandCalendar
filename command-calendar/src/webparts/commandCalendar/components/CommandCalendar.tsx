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
const GANTT_RANGE_OPTIONS: number[] = [30, 60, 90, 120];
const HOUR_SLOTS: number[] = Array.from({ length: 24 }, (_: unknown, index: number) => index);
const DEFAULT_CATEGORY_COLORS: string[] = [
  '#0078d4',
  '#107c10',
  '#d83b01',
  '#5c2d91',
  '#c239b3',
  '#0099bc',
  '#8764b8',
  '#038387'
];
type CalendarMode = 'day' | 'week' | 'month';
const BUSINESS_HOUR_START: number = 8;
const BUSINESS_HOUR_END: number = 18;
interface IHorizonBlock {
  label: string;
  startOffsetDays: number;
  endOffsetDays: number;
}
interface ISwimLaneDefinition {
  order: number;
  laneName: string;
  calendarNames: string[];
}

const HORIZON_BLOCKS: IHorizonBlock[] = [
  { label: 'Next 30 Days', startOffsetDays: 0, endOffsetDays: 30 },
  { label: 'Days 31-60', startOffsetDays: 31, endOffsetDays: 60 },
  { label: 'Days 61-90', startOffsetDays: 61, endOffsetDays: 90 },
  { label: 'Days 91-120', startOffsetDays: 91, endOffsetDays: 120 }
];

const CommandCalendar: React.FC<ICommandCalendarProps> = (props: ICommandCalendarProps) => {
  const [events, setEvents] = React.useState<ICalendarEvent[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [errorMessage, setErrorMessage] = React.useState<string | undefined>();
  const [selectedCategories, setSelectedCategories] = React.useState<string[]>([]);
  const [activeView, setActiveView] = React.useState<string>('gantt');
  const [calendarFocusDate, setCalendarFocusDate] = React.useState<Date>(startOfDay(new Date()));
  const [calendarMode, setCalendarMode] = React.useState<CalendarMode>('month');
  const [showFullDayHours, setShowFullDayHours] = React.useState<boolean>(false);
  const [reloadToken, setReloadToken] = React.useState<number>(0);
  const [ganttRangeDays, setGanttRangeDays] = React.useState<number>(60);

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

  const categoryColorMap: Record<string, string> = React.useMemo(() => {
    const configuredMap: Record<string, string> = parseCategoryColors(props.categoryColorMappings || '');
    const categoryKeyToDisplayName: Record<string, string> = {};
    events.forEach((event: ICalendarEvent) => {
      event.categories.forEach((category: string) => {
        const normalized: string = normalizeCategoryKey(category);
        if (!categoryKeyToDisplayName[normalized]) {
          categoryKeyToDisplayName[normalized] = category;
        }
      });
    });

    Object.keys(categoryKeyToDisplayName).forEach((categoryKey: string, index: number) => {
      if (!configuredMap[categoryKey]) {
        configuredMap[categoryKey] = DEFAULT_CATEGORY_COLORS[index % DEFAULT_CATEGORY_COLORS.length];
      }
    });

    return configuredMap;
  }, [events, props.categoryColorMappings]);

  const sourceByKey: Record<string, ICalendarSourceConfig> = React.useMemo(() => {
    const lookup: Record<string, ICalendarSourceConfig> = {};
    parsedSources.sources.forEach((source: ICalendarSourceConfig) => {
      lookup[source.key] = source;
    });
    return lookup;
  }, [parsedSources.sources]);

  const parsedSwimLaneMappings: ISwimLaneDefinition[] = React.useMemo(
    () => parseSwimLaneMappings(props.swimLaneMappings || ''),
    [props.swimLaneMappings]
  );

  const sourceToLaneMap: Record<string, string> = React.useMemo(() => {
    const mapping: Record<string, string> = {};
    parsedSwimLaneMappings.forEach((laneDefinition: ISwimLaneDefinition) => {
      laneDefinition.calendarNames.forEach((calendarName: string) => {
        const normalizedCalendarName: string = normalizeCategoryKey(calendarName);
        parsedSources.sources.forEach((source: ICalendarSourceConfig) => {
          if (
            normalizeCategoryKey(source.displayName) === normalizedCalendarName ||
            normalizeCategoryKey(source.listTitle) === normalizedCalendarName
          ) {
            mapping[source.key] = laneDefinition.laneName;
          }
        });
      });
    });
    return mapping;
  }, [parsedSources.sources, parsedSwimLaneMappings]);

  const laneOrderMap: Record<string, number> = React.useMemo(() => {
    const map: Record<string, number> = {};
    parsedSwimLaneMappings.forEach((laneDefinition: ISwimLaneDefinition) => {
      if (typeof map[laneDefinition.laneName] === 'undefined') {
        map[laneDefinition.laneName] = laneDefinition.order;
      }
    });
    return map;
  }, [parsedSwimLaneMappings]);

  const timelineStart: Date = startOfDay(new Date());
  const timelineDays: number = ganttRangeDays;
  const timelineEnd: Date = addDays(timelineStart, timelineDays);
  const ganttTicks: IGanttTick[] = React.useMemo(
    () => buildGanttTicks(timelineStart, timelineDays),
    [timelineDays, timelineStart]
  );

  const monthGridDates: Date[] = React.useMemo(() => {
    const monthStart: Date = startOfMonth(calendarFocusDate);
    const gridStart: Date = addDays(monthStart, -monthStart.getDay());
    return Array.from({ length: 42 }, (_: unknown, index: number) => addDays(gridStart, index));
  }, [calendarFocusDate]);

  const weekDates: Date[] = React.useMemo(() => {
    const weekStart: Date = startOfWeek(calendarFocusDate);
    return Array.from({ length: 7 }, (_: unknown, index: number) => addDays(weekStart, index));
  }, [calendarFocusDate]);

  const visibleHourSlots: number[] = React.useMemo(() => {
    if (showFullDayHours) {
      return HOUR_SLOTS;
    }

    return HOUR_SLOTS.filter((hour: number) => hour >= BUSINESS_HOUR_START && hour <= BUSINESS_HOUR_END);
  }, [showFullDayHours]);

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

  const onEventClick = (event: ICalendarEvent): void => {
    if (!event.itemUrl) {
      return;
    }

    window.open(event.itemUrl, '_blank');
  };

  const shiftCalendar = (delta: number): void => {
    setCalendarFocusDate((currentDate: Date) => {
      if (calendarMode === 'day') {
        return addDays(currentDate, delta);
      }
      if (calendarMode === 'week') {
        return addDays(currentDate, delta * 7);
      }
      return addMonths(currentDate, delta);
    });
  };

  const onCalendarModeSelected = (mode: CalendarMode): void => {
    setCalendarMode(mode);
  };

  const renderCalendarEventButton = (
    event: ICalendarEvent,
    className: string,
    showDot: boolean
  ): JSX.Element => {
    const eventColor: string = getEventColor(event, categoryColorMap);
    return (
      <TooltipHost
        content={renderEventTooltip(event)}
        directionalHint={DirectionalHint.bottomLeftEdge}
      >
        <button
          type="button"
          className={className}
          onClick={() => onEventClick(event)}
          title={event.title}
        >
          {showDot && (
            <span
              className={styles.calendarItemDot}
              style={{ backgroundColor: eventColor }}
            />
          )}
          <span className={styles.calendarItemText}>{event.title}</span>
        </button>
      </TooltipHost>
    );
  };

  const renderGanttView = (): JSX.Element => {
    const visibleEvents: ICalendarEvent[] = filteredEvents.filter((event: ICalendarEvent) =>
      rangesOverlap(event.start, event.end, timelineStart, timelineEnd)
    );

    if (visibleEvents.length === 0) {
      return <div className={styles.emptyState}>No events in the current timeline window.</div>;
    }

    const laneMap: Record<string, Record<string, ICalendarEvent[]>> = {};
    visibleEvents.forEach((event: ICalendarEvent) => {
      const source: ICalendarSourceConfig | undefined = sourceByKey[event.sourceKey];
      const laneName: string =
        sourceToLaneMap[event.sourceKey] || (source ? source.staffGroup : 'Unassigned');
      if (!laneMap[laneName]) {
        laneMap[laneName] = {};
      }
      if (!laneMap[laneName][event.sourceKey]) {
        laneMap[laneName][event.sourceKey] = [];
      }
      laneMap[laneName][event.sourceKey].push(event);
    });

    const laneNames: string[] = Object.keys(laneMap).sort((a: string, b: string) => {
      const aOrder: number = typeof laneOrderMap[a] === 'number' ? laneOrderMap[a] : 9999;
      const bOrder: number = typeof laneOrderMap[b] === 'number' ? laneOrderMap[b] : 9999;
      if (aOrder !== bOrder) {
        return aOrder - bOrder;
      }
      return a.localeCompare(b);
    });

    return (
      <div className={styles.ganttView}>
        <div className={styles.ganttToolbar}>
          <div className={styles.ganttRangeSummary}>
            {formatDate(timelineStart)} - {formatDate(timelineEnd)}
          </div>
          <div className={styles.ganttZoomLegend}>
            Select timeline range
          </div>
          <div className={styles.ganttZoomButtons}>
            {GANTT_RANGE_OPTIONS.map((rangeDays: number) => (
              <button
                type="button"
                key={rangeDays}
                className={`${styles.ganttZoomButton} ${
                  rangeDays === ganttRangeDays ? styles.ganttZoomButtonActive : ''
                }`}
                onClick={() => setGanttRangeDays(rangeDays)}
              >
                {rangeDays} days
              </button>
            ))}
          </div>
        </div>
        <div className={styles.ganttHeader}>
          <div className={styles.ganttEventColumn}>Staff / Calendar</div>
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
        {laneNames.map((laneName: string) => {
          const sourceKeys: string[] = Object.keys(laneMap[laneName]).sort((a: string, b: string) => {
            const aName: string = sourceByKey[a] ? sourceByKey[a].displayName : a;
            const bName: string = sourceByKey[b] ? sourceByKey[b].displayName : b;
            return aName.localeCompare(bName);
          });

          return (
            <div className={styles.swimLaneSection} key={laneName}>
              <div className={styles.swimLaneHeader}>{laneName}</div>
              {sourceKeys.map((sourceKey: string) => {
                const sourceEvents: ICalendarEvent[] = laneMap[laneName][sourceKey]
                  .slice()
                  .sort((a: ICalendarEvent, b: ICalendarEvent) => a.start.getTime() - b.start.getTime());
                const sourceName: string = sourceByKey[sourceKey]
                  ? sourceByKey[sourceKey].displayName
                  : sourceKey;

                return (
                  <div className={styles.ganttRow} key={`${laneName}-${sourceKey}`}>
                    <div className={styles.ganttRowLabel}>{sourceName}</div>
                    <div className={styles.ganttTrack}>
                      {sourceEvents.map((event: ICalendarEvent) => {
                        const clampedStart: Date = maxDate(event.start, timelineStart);
                        const clampedEnd: Date = minDate(event.end, timelineEnd);
                        const startOffsetDays: number = dateDiffInDays(timelineStart, clampedStart);
                        const durationDays: number = Math.max(1, dateDiffInDays(clampedStart, clampedEnd) + 1);
                        const leftPercent: number = (startOffsetDays / timelineDays) * 100;
                        const widthPercent: number = Math.max((durationDays / timelineDays) * 100, 1);
                        const tooltipContent: JSX.Element = renderEventTooltip(event);
                        const eventColor: string = getEventColor(event, categoryColorMap);

                        if (event.isRecurringInstance) {
                          return (
                            <TooltipHost
                              key={event.id}
                              content={tooltipContent}
                              directionalHint={DirectionalHint.bottomCenter}
                            >
                              <button
                                type="button"
                                className={styles.ganttDotButton}
                                onClick={() => onEventClick(event)}
                                style={{
                                  left: `${leftPercent}%`,
                                  borderColor: eventColor,
                                  backgroundColor: eventColor
                                }}
                                title={event.title}
                              />
                            </TooltipHost>
                          );
                        }

                        return (
                          <TooltipHost
                            key={event.id}
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
                                backgroundColor: eventColor
                              }}
                              title={event.title}
                            />
                          </TooltipHost>
                        );
                      })}
                    </div>
                  </div>
                );
              })}
            </div>
          );
        })}
      </div>
    );
  };

  const renderCalendarView = (): JSX.Element => {
    const calendarTitle: string = getCalendarTitle(calendarFocusDate, calendarMode);

    const renderMonthView = (): JSX.Element => (
      <div className={styles.calendarGrid}>
        {DAY_LABELS.map((dayLabel: string) => (
          <div key={dayLabel} className={`${styles.calendarCell} ${styles.calendarDayLabel}`}>
            {dayLabel}
          </div>
        ))}
        {monthGridDates.map((gridDate: Date) => {
          const dayEvents: ICalendarEvent[] = filteredEvents
            .filter((event: ICalendarEvent) => eventOccursOnDate(event, gridDate))
            .sort((a: ICalendarEvent, b: ICalendarEvent) => a.start.getTime() - b.start.getTime());

          const isCurrentMonth: boolean = gridDate.getMonth() === calendarFocusDate.getMonth();
          return (
            <div
              key={gridDate.toISOString()}
              className={`${styles.calendarCell} ${isCurrentMonth ? '' : styles.calendarCellMuted}`}
            >
              <div className={styles.calendarDate}>{gridDate.getDate()}</div>
              <div className={styles.calendarItems}>
                {dayEvents.slice(0, 3).map((event: ICalendarEvent) => (
                  <div key={`${event.id}-${gridDate.toISOString()}`}>
                    {renderCalendarEventButton(event, styles.calendarItemButton, true)}
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
    );

    const renderDayView = (): JSX.Element => {
      const dayStart: Date = startOfDay(calendarFocusDate);
      const dayEnd: Date = endOfDay(calendarFocusDate);
      const dayEvents: ICalendarEvent[] = filteredEvents
        .filter((event: ICalendarEvent) => rangesOverlap(event.start, event.end, dayStart, dayEnd))
        .sort((a: ICalendarEvent, b: ICalendarEvent) => a.start.getTime() - b.start.getTime());
      const allDayEvents: ICalendarEvent[] = dayEvents.filter((event: ICalendarEvent) => event.isAllDay);

      return (
        <div className={styles.agendaView}>
          {allDayEvents.length > 0 && (
            <div className={styles.agendaAllDayRow}>
              <div className={styles.agendaHourLabel}>All day</div>
              <div className={styles.agendaHourContent}>
                {allDayEvents.map((event: ICalendarEvent) => (
                  <div className={styles.agendaEventChip} key={`${event.id}-allday`}>
                    {renderCalendarEventButton(event, styles.agendaEventButton, true)}
                  </div>
                ))}
              </div>
            </div>
          )}
          {visibleHourSlots.map((hour: number) => {
            const slotStart: Date = new Date(
              dayStart.getFullYear(),
              dayStart.getMonth(),
              dayStart.getDate(),
              hour,
              0,
              0,
              0
            );
            const slotEnd: Date = addHours(slotStart, 1);
            const slotEvents: ICalendarEvent[] = dayEvents.filter((event: ICalendarEvent) =>
              isEventStartingInSlot(event, slotStart, slotEnd)
            );

            return (
              <div className={styles.agendaHourRow} key={`${dayStart.toISOString()}-${hour}`}>
                <div className={styles.agendaHourLabel}>{formatHourLabel(hour)}</div>
                <div className={styles.agendaHourContent}>
                  {slotEvents.map((event: ICalendarEvent) => (
                    <div className={styles.agendaEventChip} key={`${event.id}-${hour}`}>
                      {renderCalendarEventButton(event, styles.agendaEventButton, true)}
                    </div>
                  ))}
                </div>
              </div>
            );
          })}
        </div>
      );
    };

    const renderWeekView = (): JSX.Element => {
      const weekStart: Date = startOfWeek(calendarFocusDate);
      const weekEnd: Date = endOfDay(addDays(weekStart, 6));
      const weekEvents: ICalendarEvent[] = filteredEvents
        .filter((event: ICalendarEvent) => rangesOverlap(event.start, event.end, weekStart, weekEnd))
        .sort((a: ICalendarEvent, b: ICalendarEvent) => a.start.getTime() - b.start.getTime());

      return (
        <div className={styles.weekView}>
          <div className={styles.weekHeaderRow}>
            <div className={styles.weekTimeColumnHeader} />
            {weekDates.map((weekDate: Date) => (
              <div className={styles.weekDayHeader} key={`header-${weekDate.toISOString()}`}>
                {weekDate.toLocaleDateString(undefined, {
                  weekday: 'short',
                  month: 'short',
                  day: 'numeric'
                })}
              </div>
            ))}
          </div>
          {visibleHourSlots.map((hour: number) => (
            <div className={styles.weekHourRow} key={`week-hour-${hour}`}>
              <div className={styles.weekTimeColumn}>{formatHourLabel(hour)}</div>
              {weekDates.map((weekDate: Date) => {
                const slotStart: Date = new Date(
                  weekDate.getFullYear(),
                  weekDate.getMonth(),
                  weekDate.getDate(),
                  hour,
                  0,
                  0,
                  0
                );
                const slotEnd: Date = addHours(slotStart, 1);
                const slotEvents: ICalendarEvent[] = weekEvents.filter((event: ICalendarEvent) =>
                  isEventStartingInSlot(event, slotStart, slotEnd)
                );

                return (
                  <div className={styles.weekHourCell} key={`${weekDate.toISOString()}-${hour}`}>
                    {slotEvents.map((event: ICalendarEvent) => (
                      <div className={styles.weekEventWrapper} key={`${event.id}-${hour}`}>
                        {renderCalendarEventButton(event, styles.weekEventButton, false)}
                      </div>
                    ))}
                  </div>
                );
              })}
            </div>
          ))}
        </div>
      );
    };

    return (
      <div className={styles.calendarView}>
        <div className={styles.calendarHeader}>
          <div className={styles.calendarTitle}>{calendarTitle}</div>
          <div className={styles.calendarActions}>
            <DefaultButton
              text="Previous"
              onClick={() => shiftCalendar(-1)}
            />
            <DefaultButton text="Today" onClick={() => setCalendarFocusDate(startOfDay(new Date()))} />
            <DefaultButton text="Next" onClick={() => shiftCalendar(1)} />
            <DefaultButton
              text={showFullDayHours ? '8am-6pm' : '24 Hours'}
              onClick={() => setShowFullDayHours((currentValue: boolean) => !currentValue)}
            />
          </div>
          <div className={styles.calendarModeButtons}>
            {(['day', 'week', 'month'] as CalendarMode[]).map((mode: CalendarMode) => (
              <button
                type="button"
                key={mode}
                className={`${styles.calendarModeButton} ${
                  mode === calendarMode ? styles.calendarModeButtonActive : ''
                }`}
                onClick={() => onCalendarModeSelected(mode)}
              >
                {mode}
              </button>
            ))}
          </div>
        </div>
        {calendarMode === 'month' && renderMonthView()}
        {calendarMode === 'day' && renderDayView()}
        {calendarMode === 'week' && renderWeekView()}
      </div>
    );
  };

  const renderHorizonView = (): JSX.Element => {
    const today: Date = startOfDay(new Date());

    return (
      <div className={styles.quadView}>
        {HORIZON_BLOCKS.map((block: IHorizonBlock) => {
          const horizonStart: Date = addDays(startOfDay(today), block.startOffsetDays);
          const horizonEnd: Date = addDays(endOfDay(today), block.endOffsetDays);
          const horizonEvents: ICalendarEvent[] = filteredEvents
            .filter((event: ICalendarEvent) =>
              event.start.getTime() >= horizonStart.getTime() &&
              event.start.getTime() <= horizonEnd.getTime()
            )
            .sort((a: ICalendarEvent, b: ICalendarEvent) => a.start.getTime() - b.start.getTime());

          return (
            <div className={styles.quadCard} key={block.label}>
              <div className={styles.quadHeader}>{block.label}</div>
              <div className={styles.quadCount}>{horizonEvents.length}</div>
              <div className={styles.quadList}>
                {horizonEvents.slice(0, 5).map((event: ICalendarEvent) => (
                  <div className={styles.quadItem} key={`${block.label}-${event.id}`}>
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
          siteUrl|calendarListTitle|Display Name|Staff Group
        </MessageBar>
      )}

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

function buildGanttTicks(timelineStart: Date, timelineDays: number): IGanttTick[] {
  if (timelineDays <= 30) {
    return buildDailyTicks(timelineStart, timelineDays);
  }

  if (timelineDays === 60 || timelineDays === 90) {
    return buildWeeklyMondayTicks(timelineStart, timelineDays);
  }

  return buildMonthlyTicks(timelineStart, timelineDays);
}

function buildDailyTicks(timelineStart: Date, timelineDays: number): IGanttTick[] {
  const ticks: IGanttTick[] = [];
  for (let offset = 0; offset <= timelineDays; offset += 1) {
    const tickDate: Date = addDays(timelineStart, offset);
    ticks.push({
      key: `${tickDate.toISOString()}-${offset}`,
      leftPercent: (offset / timelineDays) * 100,
      label: tickDate.toLocaleDateString(undefined, { month: 'short', day: 'numeric' })
    });
  }
  return ticks;
}

function buildWeeklyMondayTicks(timelineStart: Date, timelineDays: number): IGanttTick[] {
  const ticks: IGanttTick[] = [];
  const timelineEnd: Date = addDays(timelineStart, timelineDays);
  for (
    let cursor: Date = firstMondayOnOrAfter(timelineStart);
    cursor.getTime() <= timelineEnd.getTime();
    cursor = addDays(cursor, 7)
  ) {
    const offsetDays: number = dateDiffInDays(timelineStart, cursor);
    ticks.push({
      key: cursor.toISOString(),
      leftPercent: (offsetDays / timelineDays) * 100,
      label: cursor.toLocaleDateString(undefined, { month: 'short', day: 'numeric' })
    });
  }

  if (ticks.length === 0) {
    ticks.push({
      key: `fallback-${timelineStart.toISOString()}`,
      leftPercent: 0,
      label: timelineStart.toLocaleDateString(undefined, { month: 'short', day: 'numeric' })
    });
  }
  return ticks;
}

function buildMonthlyTicks(timelineStart: Date, timelineDays: number): IGanttTick[] {
  const ticks: IGanttTick[] = [];
  for (let monthOffset = 0; monthOffset < 4; monthOffset += 1) {
    const monthStart: Date = new Date(
      timelineStart.getFullYear(),
      timelineStart.getMonth() + monthOffset,
      1
    );
    const offsetDays: number = dateDiffInDays(timelineStart, monthStart);
    const leftPercent: number = monthOffset === 0 ? 0 : (offsetDays / timelineDays) * 100;
    ticks.push({
      key: monthStart.toISOString(),
      leftPercent: Math.max(0, leftPercent),
      label: monthStart.toLocaleDateString(undefined, { month: 'long' })
    });
  }
  return ticks;
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
  return startOfDay(nextDate);
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

function startOfWeek(date: Date): Date {
  const dayStart: Date = startOfDay(date);
  return addDays(dayStart, -dayStart.getDay());
}

function firstMondayOnOrAfter(date: Date): Date {
  const dayStart: Date = startOfDay(date);
  const dayOfWeek: number = dayStart.getDay();
  const daysUntilMonday: number = dayOfWeek === 1 ? 0 : ((8 - dayOfWeek) % 7);
  return addDays(dayStart, daysUntilMonday);
}

function addHours(date: Date, hours: number): Date {
  const nextDate: Date = new Date(date.getTime());
  nextDate.setHours(nextDate.getHours() + hours);
  return nextDate;
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

function formatHourLabel(hour: number): string {
  const dateForLabel: Date = new Date(2000, 0, 1, hour, 0, 0, 0);
  return dateForLabel.toLocaleTimeString(undefined, {
    hour: 'numeric',
    minute: '2-digit'
  });
}

function getCalendarTitle(calendarDate: Date, mode: CalendarMode): string {
  if (mode === 'day') {
    return calendarDate.toLocaleDateString(undefined, {
      weekday: 'long',
      month: 'long',
      day: 'numeric',
      year: 'numeric'
    });
  }

  if (mode === 'week') {
    const weekStart: Date = startOfWeek(calendarDate);
    const weekEnd: Date = addDays(weekStart, 6);
    return `${formatDate(weekStart)} - ${formatDate(weekEnd)}`;
  }

  return calendarDate.toLocaleString(undefined, { month: 'long', year: 'numeric' });
}

function isEventStartingInSlot(event: ICalendarEvent, slotStart: Date, slotEnd: Date): boolean {
  return event.start.getTime() >= slotStart.getTime() && event.start.getTime() < slotEnd.getTime();
}

function eventOccursOnDate(event: ICalendarEvent, targetDate: Date): boolean {
  const dayStart: Date = startOfDay(targetDate);
  const dayEnd: Date = endOfDay(targetDate);

  if (event.isAllDay) {
    return rangesOverlap(event.start, event.end, dayStart, dayEnd);
  }

  return isSameCalendarDay(event.start, targetDate);
}

function isSameCalendarDay(firstDate: Date, secondDate: Date): boolean {
  return (
    firstDate.getFullYear() === secondDate.getFullYear() &&
    firstDate.getMonth() === secondDate.getMonth() &&
    firstDate.getDate() === secondDate.getDate()
  );
}

function parseCategoryColors(rawMappings: string): Record<string, string> {
  const mapping: Record<string, string> = {};
  rawMappings
    .split('\n')
    .map((line: string) => line.trim())
    .filter((line: string) => line.length > 0)
    .forEach((line: string) => {
      const parts: string[] = line.split('|').map((part: string) => part.trim());
      if (parts.length < 2) {
        return;
      }
      const categoryName: string = parts[0];
      const color: string = parts[1];
      if (!categoryName || !/^#?[0-9a-f]{6}$/i.test(color)) {
        return;
      }
      const normalizedColor: string = color.charAt(0) === '#' ? color : `#${color}`;
      mapping[normalizeCategoryKey(categoryName)] = normalizedColor;
    });
  return mapping;
}

function parseSwimLaneMappings(rawMappings: string): ISwimLaneDefinition[] {
  const definitions: ISwimLaneDefinition[] = [];
  rawMappings
    .split('\n')
    .map((line: string) => line.trim())
    .filter((line: string) => line.length > 0)
    .forEach((line: string, index: number) => {
      const parts: string[] = line.split('|').map((part: string) => part.trim());
      if (parts.length < 2) {
        return;
      }

      let order: number = index + 1;
      let laneName: string = '';
      let calendarsRaw: string = '';
      if (parts.length >= 3 && /^\d+$/.test(parts[0])) {
        order = parseInt(parts[0], 10);
        laneName = parts[1];
        calendarsRaw = parts[2];
      } else if (parts.length >= 3 && /^\d+$/.test(parts[1])) {
        laneName = parts[0];
        order = parseInt(parts[1], 10);
        calendarsRaw = parts[2];
      } else {
        laneName = parts[0];
        calendarsRaw = parts[1];
      }

      if (!laneName || !calendarsRaw) {
        return;
      }

      const calendarNames: string[] = calendarsRaw
        .split(',')
        .map((value: string) => value.trim())
        .filter((value: string) => value.length > 0);

      if (calendarNames.length === 0) {
        return;
      }

      definitions.push({
        order,
        laneName,
        calendarNames
      });
    });

  return definitions;
}

function normalizeCategoryKey(category: string): string {
  return category.toLocaleLowerCase().trim();
}

function getEventColor(event: ICalendarEvent, categoryColorMap: Record<string, string>): string {
  const firstCategory: string = event.categories.length > 0 ? event.categories[0] : 'uncategorized';
  const mappedColor: string = categoryColorMap[normalizeCategoryKey(firstCategory)];
  return mappedColor || '#605e5c';
}

export default CommandCalendar;

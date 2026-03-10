import * as React from 'react';
import TimelineCalendar from './TimelineCalendar';
import { ITimelineCalendarProps } from './ITimelineCalendarProps';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { TooltipHost } from 'office-ui-fabric-react/lib/Tooltip';

type TViewKey = 'gantt' | 'calendar' | 'horizon';
type TCalendarMode = 'month' | 'week7' | 'week5';

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

const TimelineCalendarTabbed: React.FC<ITimelineCalendarProps> = (props: ITimelineCalendarProps) => {
  const [activeView, setActiveView] = React.useState<TViewKey>('calendar');
  const [events, setEvents] = React.useState<ITimelineItem[]>([]);
  const [isLoadingData, setIsLoadingData] = React.useState<boolean>(true);
  const [selectedCategoryKeys, setSelectedCategoryKeys] = React.useState<string[]>([]);
  const [selectedCalendarKeys, setSelectedCalendarKeys] = React.useState<string[]>([]);
  const [calendarReferenceDate, setCalendarReferenceDate] = React.useState<Date>(startOfDay(new Date()));
  const [calendarMode, setCalendarMode] = React.useState<TCalendarMode>('month');
  const eventsSignatureRef = React.useRef<string>('');

  const timelineElement = React.useMemo((): JSX.Element => (
    <TimelineCalendar
      {...props}
      selectedCategoryKeys={selectedCategoryKeys}
      selectedCalendarKeys={selectedCalendarKeys}
      hideLegendBar
    />
  ), [props, selectedCategoryKeys, selectedCalendarKeys]);

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
          calendarLabel: calendarFromItem.label
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
        `${eventItem.id}|${eventItem.start.getTime()}|${eventItem.end.getTime()}|${eventItem.title}|${eventItem.categoryKey}|${eventItem.calendarKey}|${eventItem.calendarLabel}|${eventItem.modified}`
      ))
      .join('~');

    if (eventsSignatureRef.current === nextSignature) {
      return;
    }

    eventsSignatureRef.current = nextSignature;
    setEvents(mappedItems);
  }, [categoryMetaByKey, props.ensureValidClassName, isLoadingData]);

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

  const calendarFilterOptions: IDropdownOption[] = React.useMemo(() => {
    const optionMap: Map<string, string> = new Map<string, string>();

    events.forEach((eventItem: ITimelineItem) => {
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
  }, [events]);

  const filteredEvents: ITimelineItem[] = React.useMemo(() => {
    const selectedCategoryKeySet: Set<string> = new Set<string>(selectedCategoryKeys);
    const selectedCalendarKeySet: Set<string> = new Set<string>(selectedCalendarKeys);

    return events.filter((eventItem: ITimelineItem) => {
      const categoryMatch = selectedCategoryKeys.length === 0 ||
        (eventItem.categoryKey && selectedCategoryKeySet.has(eventItem.categoryKey));
      const calendarMatch = selectedCalendarKeys.length === 0 ||
        (eventItem.calendarKey && selectedCalendarKeySet.has(eventItem.calendarKey));
      return categoryMatch && calendarMatch;
    });
  }, [events, selectedCategoryKeys, selectedCalendarKeys]);

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
        <div style={{ display: 'flex', gap: '8px', marginBottom: '10px', flexWrap: 'wrap' }}>
          <TabButton isActive={activeView === 'calendar'} label="Calendar View" onClick={() => setActiveView('calendar')} />
          <TabButton isActive={activeView === 'gantt'} label="Gnatt Chart View" onClick={() => setActiveView('gantt')} />
          <TabButton isActive={activeView === 'horizon'} label="30-60-90-120 View" onClick={() => setActiveView('horizon')} />
        </div>

        <div style={{ display: 'flex', gap: '8px', alignItems: 'flex-end', flexWrap: 'wrap' }}>
          <Dropdown
            label="Category filter"
            placeholder="All categories"
            multiSelect
            options={categoryFilterOptions}
            selectedKeys={selectedCategoryKeys}
            onChange={onCategoryFilterChange}
            styles={{ dropdown: { minWidth: 280 } }}
          />
          <Dropdown
            label="Calendar filter"
            placeholder="All calendars"
            multiSelect
            options={calendarFilterOptions}
            selectedKeys={selectedCalendarKeys}
            onChange={onCalendarFilterChange}
            styles={{ dropdown: { minWidth: 280 } }}
          />
          <SmallButton
            label="Clear Filters"
            onClick={() => {
              setSelectedCategoryKeys([]);
              setSelectedCalendarKeys([]);
            }}
            disabled={selectedCategoryKeys.length === 0 && selectedCalendarKeys.length === 0}
          />
        </div>
      </div>

      {activeView !== 'gantt' && (
        <CategoryLegend options={categoryFilterOptions} selectedCategoryKeys={selectedCategoryKeys} />
      )}

      {activeView === 'calendar' && (
        <div
          style={{
            display: 'grid',
            gridTemplateColumns: '1fr auto 1fr',
            alignItems: 'center',
            gap: '8px',
            marginBottom: '10px'
          }}
        >
          <div style={{ display: 'flex', gap: '8px', justifySelf: 'start' }}>
            <SmallButton
              label={calendarMode === 'month' ? 'Last Month' : 'Previous Week'}
              onClick={() => {
                if (calendarMode === 'month') {
                  setCalendarReferenceDate(addMonths(calendarReferenceDate, -1));
                } else {
                  setCalendarReferenceDate(addDays(calendarReferenceDate, -7));
                }
              }}
            />
            <SmallButton label="Today" onClick={() => setCalendarReferenceDate(startOfDay(new Date()))} />
            <SmallButton
              label={calendarMode === 'month' ? 'Next Month' : 'Next Week'}
              onClick={() => {
                if (calendarMode === 'month') {
                  setCalendarReferenceDate(addMonths(calendarReferenceDate, 1));
                } else {
                  setCalendarReferenceDate(addDays(calendarReferenceDate, 7));
                }
              }}
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
          <div style={{ display: 'flex', gap: '8px', justifySelf: 'end' }}>
            <SmallButton label="Month" onClick={() => setCalendarMode('month')} isActive={calendarMode === 'month'} />
            <SmallButton label="7-Day Week" onClick={() => setCalendarMode('week7')} isActive={calendarMode === 'week7'} />
            <SmallButton label="5-Day Week" onClick={() => setCalendarMode('week5')} isActive={calendarMode === 'week5'} />
          </div>
        </div>
      )}

      <div style={{ display: activeView === 'gantt' ? 'block' : 'none' }}>
        {timelineElement}
      </div>
      <div style={{ display: activeView === 'calendar' ? 'block' : 'none' }}>
        <CalendarView events={filteredEvents} referenceDate={calendarReferenceDate} mode={calendarMode} />
      </div>
      <div style={{ display: activeView === 'horizon' ? 'block' : 'none' }}>
        <HorizonView events={filteredEvents} />
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
      border: isActive ? '1px solid #0078d4' : '1px solid #c8c6c4',
      background: isActive ? '#eff6fc' : '#fff',
      color: isActive ? '#0078d4' : '#323130',
      borderRadius: '4px',
      padding: '10px 14px',
      fontSize: '1.5em',
      fontWeight: 700,
      cursor: 'pointer'
    }}
  >
    {label}
  </button>
);

const SmallButton: React.FC<{ label: string; onClick: () => void; disabled?: boolean; isActive?: boolean }> = ({
  label,
  onClick,
  disabled,
  isActive
}) => (
  <button
    type="button"
    disabled={disabled}
    onClick={onClick}
    style={{
      border: isActive ? '1px solid #0078d4' : '1px solid #c8c6c4',
      background: isActive ? '#eff6fc' : '#fff',
      color: isActive ? '#0078d4' : '#323130',
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

  return <CalendarWeekGrid events={events} referenceDate={referenceDate} mode={mode} />;
};

const CalendarMonthGrid: React.FC<{ events: ITimelineItem[]; referenceDate: Date }> = ({ events, referenceDate }) => {
  const monthStart: Date = new Date(referenceDate.getFullYear(), referenceDate.getMonth(), 1);
  const gridStart: Date = startOfWeek(monthStart, false);
  const [expandedDayKeys, setExpandedDayKeys] = React.useState<string[]>([]);
  const gridDays: Date[] = [];

  for (let index = 0; index < 42; index++) {
    gridDays.push(addDays(gridStart, index));
  }

  return (
    <div style={{ border: '1px solid #edebe9' }}>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, minmax(120px, 1fr))', background: '#faf9f8', borderBottom: '1px solid #edebe9' }}>
        {DAY_LABELS.map((label: string) => (
          <div key={label} style={{ padding: '6px', fontWeight: 600, fontSize: '12px' }}>{label}</div>
        ))}
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, minmax(120px, 1fr))' }}>
        {gridDays.map((day: Date) => {
          const dayEvents: ITimelineItem[] = eventsForDate(events, day);
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
};

const CalendarWeekGrid: React.FC<{ events: ITimelineItem[]; referenceDate: Date; mode: TCalendarMode }> = ({ events, referenceDate, mode }) => {
  const isWorkWeek: boolean = mode === 'week5';
  const weekStart: Date = startOfWeek(referenceDate, isWorkWeek);
  const dayCount: number = isWorkWeek ? 5 : 7;
  const labels: string[] = isWorkWeek ? WEEKDAY_LABELS : DAY_LABELS;
  const days: Date[] = [];

  for (let index = 0; index < dayCount; index++) {
    days.push(addDays(weekStart, index));
  }

  return (
    <div style={{ border: '1px solid #edebe9' }}>
      <div style={{ display: 'grid', gridTemplateColumns: `repeat(${dayCount}, minmax(150px, 1fr))`, background: '#faf9f8', borderBottom: '1px solid #edebe9' }}>
        {days.map((day: Date, index: number) => (
          <div key={`week-label-${day.toISOString()}`} style={{ padding: '6px', fontWeight: 600, fontSize: '12px' }}>
            {labels[index]} {day.getMonth() + 1}/{day.getDate()}
          </div>
        ))}
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: `repeat(${dayCount}, minmax(150px, 1fr))` }}>
        {days.map((day: Date) => {
          const dayEvents: ITimelineItem[] = eventsForDate(events, day);
          return (
            <div
              key={`week-day-${day.toISOString()}`}
              style={{
                borderRight: '1px solid #f3f2f1',
                borderBottom: '1px solid #f3f2f1',
                minHeight: '220px',
                padding: '6px',
                background: '#fff'
              }}
            >
              {dayEvents.length === 0 && <div style={{ fontSize: '11px', color: '#8a8886' }}>No events</div>}
              {dayEvents.map((eventItem: ITimelineItem) => (
                <EventLinkWithTooltip
                  key={`week-${day.toISOString()}-${eventItem.id}`}
                  event={eventItem}
                  compact={false}
                  showDateRange
                />
              ))}
            </div>
          );
        })}
      </div>
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
          .filter((eventItem: ITimelineItem) => eventItem.start.getTime() >= rangeStart.getTime() && eventItem.start.getTime() <= rangeEnd.getTime())
          .sort((firstEvent: ITimelineItem, secondEvent: ITimelineItem) => firstEvent.start.getTime() - secondEvent.start.getTime());

        return (
          <div key={block.label} style={{ border: '1px solid #edebe9', borderRadius: '6px', padding: '10px' }}>
            <div style={{ fontSize: '16px', fontWeight: 600 }}>{block.label}</div>
            <div style={{ fontSize: '28px', fontWeight: 700, margin: '8px 0' }}>{blockEvents.length}</div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
              {blockEvents.map((eventItem: ITimelineItem) => (
                <div key={`${block.label}-${eventItem.id}`} style={{ fontSize: '12px' }}>
                  <EventLinkWithTooltip event={eventItem} />
                  <div style={{ color: '#605e5c', marginLeft: '14px' }}>{eventItem.start.toLocaleDateString()}</div>
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

const EventLinkWithTooltip: React.FC<{ event: ITimelineItem; compact?: boolean; showDateRange?: boolean }> = ({ event, compact, showDateRange }) => {
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

  const textStyle: React.CSSProperties = compact ? {
    display: 'block',
    fontSize: '11px',
    marginBottom: '2px',
    lineHeight: '1.2'
  } : {
    display: 'block',
    fontSize: '12px',
    lineHeight: '1.2'
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
      <div>
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
            color: '#0078d4',
            textDecoration: 'none',
            border: 'none',
            background: 'transparent',
            padding: 0,
            margin: 0,
            textAlign: 'left',
            cursor: 'pointer'
          }}
        >
          <CategoryDot color={event.categoryColor} />
          {event.title}
        </button>
      ) : (
          <span style={{ ...textStyle, color: '#323130' }}>
            <CategoryDot color={event.categoryColor} />
            {event.title}
          </span>
        )}
        {showDateRange && (
          <span style={dateRangeTextStyle}>
            {event.start.toLocaleString()} - {event.end.toLocaleString()}
          </span>
        )}
      </div>
    </TooltipHost>
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

  const isWorkWeek: boolean = mode === 'week5';
  const startDate: Date = startOfWeek(referenceDate, isWorkWeek);
  const endDate: Date = addDays(startDate, isWorkWeek ? 4 : 6);
  return `${startDate.toLocaleDateString()} - ${endDate.toLocaleDateString()}`;
}

function eventsForDate(events: ITimelineItem[], day: Date): ITimelineItem[] {
  return events.filter((eventItem: ITimelineItem) => eventOccursOnDate(eventItem, day));
}

function eventOccursOnDate(eventItem: ITimelineItem, day: Date): boolean {
  const dayStart: number = startOfDay(day).getTime();
  const dayEnd: number = endOfDay(day).getTime();
  return eventItem.start.getTime() <= dayEnd && eventItem.end.getTime() >= dayStart;
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

export default TimelineCalendarTabbed;

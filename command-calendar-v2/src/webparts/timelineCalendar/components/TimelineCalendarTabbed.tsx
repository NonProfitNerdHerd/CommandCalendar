import * as React from 'react';
import TimelineCalendar from './TimelineCalendar';
import { ITimelineCalendarProps } from './ITimelineCalendarProps';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

type TViewKey = 'gantt' | 'calendar' | 'agenda' | 'horizon';

interface ITimelineItem {
  id: string;
  title: string;
  start: Date;
  end: Date;
  categoryKey: string;
  categoryLabel: string;
  categoryColor: string;
}

interface ICategoryMeta {
  key: string;
  label: string;
  color: string;
}

interface IHorizonBlock {
  label: string;
  startOffsetDays: number;
  endOffsetDays: number;
}

const DAY_LABELS: string[] = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const HORIZON_BLOCKS: IHorizonBlock[] = [
  { label: 'Next 30 Days', startOffsetDays: 0, endOffsetDays: 30 },
  { label: 'Days 31-60', startOffsetDays: 31, endOffsetDays: 60 },
  { label: 'Days 61-90', startOffsetDays: 61, endOffsetDays: 90 },
  { label: 'Days 91-120', startOffsetDays: 91, endOffsetDays: 120 }
];

const TimelineCalendarTabbed: React.FC<ITimelineCalendarProps> = (props: ITimelineCalendarProps) => {
  const [activeView, setActiveView] = React.useState<TViewKey>('gantt');
  const [events, setEvents] = React.useState<ITimelineItem[]>([]);
  const [selectedCategoryKeys, setSelectedCategoryKeys] = React.useState<string[]>([]);
  const eventsSignatureRef = React.useRef<string>('');

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

        return {
          id: String(item.id || `${start.toISOString()}-${title}`),
          title,
          start,
          end,
          categoryKey: categoryFromItem.key,
          categoryLabel: categoryFromItem.label,
          categoryColor: categoryFromItem.color
        };
      })
      .sort((a: ITimelineItem, b: ITimelineItem) => a.start.getTime() - b.start.getTime());

    const nextSignature: string = mappedItems
      .map((event: ITimelineItem) => `${event.id}|${event.start.getTime()}|${event.end.getTime()}|${event.title}|${event.categoryKey}`)
      .join('~');

    if (eventsSignatureRef.current === nextSignature) {
      return;
    }

    eventsSignatureRef.current = nextSignature;
    setEvents(mappedItems);
  }, [categoryMetaByKey, props.ensureValidClassName]);

  const categoryFilterOptions: IDropdownOption[] = React.useMemo(() => {
    const optionMap: Map<string, ICategoryMeta> = new Map<string, ICategoryMeta>();

    categoryMetaByKey.forEach((meta: ICategoryMeta, key: string) => {
      optionMap.set(key, meta);
    });

    events.forEach((event: ITimelineItem) => {
      if (!event.categoryKey || optionMap.has(event.categoryKey)) {
        return;
      }

      optionMap.set(event.categoryKey, {
        key: event.categoryKey,
        label: event.categoryLabel || event.categoryKey,
        color: event.categoryColor || '#8a8886'
      });
    });

    const optionValues: ICategoryMeta[] = [];
    optionMap.forEach((meta: ICategoryMeta) => {
      optionValues.push(meta);
    });

    optionValues.sort((first: ICategoryMeta, second: ICategoryMeta) => first.label.localeCompare(second.label));

    return optionValues.map((meta: ICategoryMeta) => ({
      key: meta.key,
      text: meta.label
    }));
  }, [categoryMetaByKey, events]);

  const filteredEvents: ITimelineItem[] = React.useMemo(() => {
    if (selectedCategoryKeys.length === 0) {
      return events;
    }

    const selectedKeySet: Set<string> = new Set<string>(selectedCategoryKeys);
    return events.filter((event: ITimelineItem) => event.categoryKey && selectedKeySet.has(event.categoryKey));
  }, [events, selectedCategoryKeys]);

  const onCategoryFilterChange = React.useCallback((event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) {
      return;
    }

    const optionKey: string = String(option.key);
    setSelectedCategoryKeys((prevKeys: string[]) => {
      if (option.selected) {
        if (prevKeys.indexOf(optionKey) > -1) {
          return prevKeys;
        }

        return [...prevKeys, optionKey];
      }

      return prevKeys.filter((key: string) => key !== optionKey);
    });
  }, []);

  React.useEffect(() => {
    if (activeView === 'gantt') {
      return;
    }

    const intervalId: number = window.setInterval(captureEvents, 1500);
    captureEvents();

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

  return (
    <div>
      <div style={{ display: 'flex', gap: '8px', marginBottom: '8px' }}>
        <TabButton
          isActive={activeView === 'gantt'}
          label="Gnatt Chart View"
          onClick={() => setActiveView('gantt')}
        />
        <TabButton
          isActive={activeView === 'calendar'}
          label="Calendar View"
          onClick={() => setActiveView('calendar')}
        />
        <TabButton
          isActive={activeView === 'agenda'}
          label="Agenda View"
          onClick={() => setActiveView('agenda')}
        />
        <TabButton
          isActive={activeView === 'horizon'}
          label="30-60-90-120 View"
          onClick={() => setActiveView('horizon')}
        />
      </div>

      {activeView !== 'gantt' && (
        <div style={{ display: 'flex', alignItems: 'flex-end', gap: '8px', marginBottom: '10px' }}>
          <Dropdown
            label="Category filter"
            placeholder="All categories"
            multiSelect
            options={categoryFilterOptions}
            selectedKeys={selectedCategoryKeys}
            onChange={onCategoryFilterChange}
            styles={{ dropdown: { minWidth: 280 } }}
          />
          <button
            type="button"
            onClick={() => setSelectedCategoryKeys([])}
            disabled={selectedCategoryKeys.length === 0}
            style={{
              border: '1px solid #c8c6c4',
              background: '#fff',
              color: '#323130',
              borderRadius: '4px',
              height: '32px',
              padding: '0 12px',
              cursor: selectedCategoryKeys.length === 0 ? 'default' : 'pointer',
              opacity: selectedCategoryKeys.length === 0 ? 0.6 : 1
            }}
          >
            Clear
          </button>
        </div>
      )}

      <div style={{ display: activeView === 'gantt' ? 'block' : 'none' }}>
        <TimelineCalendar {...props} />
      </div>
      <div style={{ display: activeView === 'calendar' ? 'block' : 'none' }}>
        <CalendarMonthGrid events={filteredEvents} />
      </div>
      <div style={{ display: activeView === 'agenda' ? 'block' : 'none' }}>
        <AgendaView events={filteredEvents} />
      </div>
      <div style={{ display: activeView === 'horizon' ? 'block' : 'none' }}>
        <HorizonView events={filteredEvents} />
      </div>
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
      padding: '6px 12px',
      fontWeight: isActive ? 600 : 400,
      cursor: 'pointer'
    }}
  >
    {label}
  </button>
);

const CalendarMonthGrid: React.FC<{ events: ITimelineItem[] }> = ({ events }) => {
  const currentDate: Date = new Date();
  const monthStart: Date = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);
  const gridStart: Date = new Date(monthStart.getTime());
  gridStart.setDate(gridStart.getDate() - gridStart.getDay());
  const gridDays: Date[] = [];

  for (let i = 0; i < 42; i++) {
    const dateCell: Date = new Date(gridStart.getTime());
    dateCell.setDate(gridStart.getDate() + i);
    gridDays.push(dateCell);
  }

  return (
    <div style={{ border: '1px solid #edebe9' }}>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, minmax(110px, 1fr))', background: '#faf9f8', borderBottom: '1px solid #edebe9' }}>
        {DAY_LABELS.map((label: string) => (
          <div key={label} style={{ padding: '6px', fontWeight: 600, fontSize: '12px' }}>{label}</div>
        ))}
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, minmax(110px, 1fr))' }}>
        {gridDays.map((day: Date) => {
          const dayEvents: ITimelineItem[] = events.filter((event: ITimelineItem) => isSameDay(event.start, day));
          const isCurrentMonth: boolean = day.getMonth() === currentDate.getMonth();
          return (
            <div
              key={day.toISOString()}
              style={{
                borderRight: '1px solid #f3f2f1',
                borderBottom: '1px solid #f3f2f1',
                minHeight: '95px',
                padding: '6px',
                background: isCurrentMonth ? '#fff' : '#faf9f8'
              }}
            >
              <div style={{ fontWeight: 600, fontSize: '12px', marginBottom: '4px' }}>{day.getDate()}</div>
              {dayEvents.slice(0, 3).map((event: ITimelineItem) => (
                <div key={`${day.toISOString()}-${event.id}`} style={{ fontSize: '11px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                  <CategoryDot color={event.categoryColor} />
                  {event.title}
                </div>
              ))}
              {dayEvents.length > 3 && <div style={{ fontSize: '11px', color: '#605e5c' }}>+{dayEvents.length - 3} more</div>}
            </div>
          );
        })}
      </div>
    </div>
  );
};

const AgendaView: React.FC<{ events: ITimelineItem[] }> = ({ events }) => {
  const now: Date = new Date();
  const endRange: Date = new Date(now.getTime());
  endRange.setDate(endRange.getDate() + 30);
  const upcomingEvents: ITimelineItem[] = events
    .filter((event: ITimelineItem) => event.start.getTime() >= now.getTime() && event.start.getTime() <= endRange.getTime())
    .sort((a: ITimelineItem, b: ITimelineItem) => a.start.getTime() - b.start.getTime());

  return (
    <div style={{ border: '1px solid #edebe9', padding: '10px' }}>
      {upcomingEvents.length === 0 && (
        <div style={{ color: '#605e5c' }}>No upcoming events in next 30 days.</div>
      )}
      {upcomingEvents.map((event: ITimelineItem) => (
        <div key={`agenda-${event.id}`} style={{ padding: '8px 0', borderBottom: '1px solid #f3f2f1' }}>
          <div style={{ fontWeight: 600 }}>
            <CategoryDot color={event.categoryColor} />
            {event.title}
          </div>
          <div style={{ fontSize: '12px', color: '#605e5c' }}>
            {event.start.toLocaleString()} - {event.end.toLocaleString()}
          </div>
        </div>
      ))}
    </div>
  );
};

const HorizonView: React.FC<{ events: ITimelineItem[] }> = ({ events }) => {
  const today: Date = new Date();
  const todayStart: Date = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 0, 0, 0, 0);

  return (
    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, minmax(190px, 1fr))', gap: '10px' }}>
      {HORIZON_BLOCKS.map((block: IHorizonBlock) => {
        const rangeStart: Date = addDays(todayStart, block.startOffsetDays);
        const rangeEnd: Date = addDays(new Date(todayStart.getFullYear(), todayStart.getMonth(), todayStart.getDate(), 23, 59, 59, 999), block.endOffsetDays);
        const blockEvents: ITimelineItem[] = events
          .filter((event: ITimelineItem) => event.start.getTime() >= rangeStart.getTime() && event.start.getTime() <= rangeEnd.getTime())
          .sort((a: ITimelineItem, b: ITimelineItem) => a.start.getTime() - b.start.getTime());

        return (
          <div key={block.label} style={{ border: '1px solid #edebe9', borderRadius: '6px', padding: '10px' }}>
            <div style={{ fontSize: '16px', fontWeight: 600 }}>{block.label}</div>
            <div style={{ fontSize: '28px', fontWeight: 700, margin: '8px 0' }}>{blockEvents.length}</div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
              {blockEvents.slice(0, 5).map((event: ITimelineItem) => (
                <div key={`${block.label}-${event.id}`} style={{ fontSize: '12px' }}>
                  <div style={{ fontWeight: 600 }}>
                    <CategoryDot color={event.categoryColor} />
                    {event.title}
                  </div>
                  <div style={{ color: '#605e5c' }}>{event.start.toLocaleDateString()}</div>
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

const CategoryDot: React.FC<{ color: string }> = ({ color }) => (
  <span
    style={{
      display: 'inline-block',
      width: '8px',
      height: '8px',
      borderRadius: '50%',
      backgroundColor: color || '#8a8886',
      marginRight: '6px'
    }}
  />
);

function stripHtml(value: string): string {
  if (!value) {
    return '';
  }
  const element: HTMLDivElement = document.createElement('div');
  element.innerHTML = value;
  return (element.textContent || element.innerText || '').trim();
}

function isSameDay(firstDate: Date, secondDate: Date): boolean {
  return (
    firstDate.getFullYear() === secondDate.getFullYear() &&
    firstDate.getMonth() === secondDate.getMonth() &&
    firstDate.getDate() === secondDate.getDate()
  );
}

function addDays(date: Date, days: number): Date {
  const nextDate: Date = new Date(date.getTime());
  nextDate.setDate(nextDate.getDate() + days);
  return nextDate;
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

function resolveEventCategory(item: any, categoryMetaByKey: Map<string, ICategoryMeta>, ensureValidClassName: (value: string) => string): ICategoryMeta {
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
    return categoryMetaByKey.get(classMatch);
  }

  const rawCategoryText: string = String(item.Category || item.category || '').split(',')[0].trim();
  if (rawCategoryText) {
    const normalizedCategoryKey: string = ensureValidClassName(rawCategoryText);
    if (categoryMetaByKey.has(normalizedCategoryKey)) {
      return categoryMetaByKey.get(normalizedCategoryKey);
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
      return categoryMetaByKey.get(fallbackKey);
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

export default TimelineCalendarTabbed;

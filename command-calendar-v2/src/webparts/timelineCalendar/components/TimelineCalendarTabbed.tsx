import * as React from 'react';
import TimelineCalendar from './TimelineCalendar';
import { ITimelineCalendarProps } from './ITimelineCalendarProps';

type TViewKey = 'gantt' | 'calendar' | 'agenda' | 'horizon';

interface ITimelineItem {
  id: string;
  title: string;
  start: Date;
  end: Date;
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
  const eventsSignatureRef = React.useRef<string>('');

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
        return {
          id: String(item.id || `${start.toISOString()}-${title}`),
          title,
          start,
          end
        };
      })
      .sort((a: ITimelineItem, b: ITimelineItem) => a.start.getTime() - b.start.getTime());

    const nextSignature: string = mappedItems
      .map((event: ITimelineItem) => `${event.id}|${event.start.getTime()}|${event.end.getTime()}|${event.title}`)
      .join('~');

    if (eventsSignatureRef.current === nextSignature) {
      return;
    }

    eventsSignatureRef.current = nextSignature;
    setEvents(mappedItems);
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

      <div style={{ display: activeView === 'gantt' ? 'block' : 'none' }}>
        <TimelineCalendar {...props} />
      </div>
      <div style={{ display: activeView === 'calendar' ? 'block' : 'none' }}>
        <CalendarMonthGrid events={events} />
      </div>
      <div style={{ display: activeView === 'agenda' ? 'block' : 'none' }}>
        <AgendaView events={events} />
      </div>
      <div style={{ display: activeView === 'horizon' ? 'block' : 'none' }}>
        <HorizonView events={events} />
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
          <div style={{ fontWeight: 600 }}>{event.title}</div>
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
                  <div style={{ fontWeight: 600 }}>{event.title}</div>
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

export default TimelineCalendarTabbed;

import * as React from 'react';
import TimelineCalendar from './TimelineCalendar';
import { ITimelineCalendarProps } from './ITimelineCalendarProps';

type TViewKey = 'gantt' | 'calendar';

interface ITimelineItem {
  id: string;
  title: string;
  start: Date;
  end: Date;
}

const DAY_LABELS: string[] = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

const TimelineCalendarTabbed: React.FC<ITimelineCalendarProps> = (props: ITimelineCalendarProps) => {
  const [activeView, setActiveView] = React.useState<TViewKey>('gantt');
  const [events, setEvents] = React.useState<ITimelineItem[]>([]);

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

    setEvents(mappedItems);
  }, []);

  React.useEffect(() => {
    const intervalId: number = window.setInterval(captureEvents, 1500);
    captureEvents();
    return () => {
      window.clearInterval(intervalId);
    };
  }, [captureEvents]);

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
      </div>

      <div style={{ display: activeView === 'gantt' ? 'block' : 'none' }}>
        <TimelineCalendar {...props} />
      </div>
      <div style={{ display: activeView === 'calendar' ? 'block' : 'none' }}>
        <CalendarMonthGrid events={events} />
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

export default TimelineCalendarTabbed;

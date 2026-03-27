import * as React from 'react';
import {
  assignTimeOverlapLanes,
  CALENDAR_GRID_END_HOUR,
  CALENDAR_GRID_START_HOUR,
  CALENDAR_GRID_TOTAL_MINUTES,
  CALENDAR_PX_PER_HOUR,
  clipIntervalToDayGrid,
  minutesFromGridStart
} from './calendarTimeGridLayout';

const TIME_GUTTER_WIDTH_PX = 52;
const COLUMN_MIN_WIDTH_PX = 120;
const EVENT_BLOCK_GAP_PX = 1;

export interface ICalendarTimeGridProps<T> {
  days: Date[];
  columnHeaderLabels: string[];
  timedEventsByDay: T[][];
  getEventStart: (item: T) => Date;
  getEventEndInclusive: (item: T) => Date;
  renderEvent: (item: T) => React.ReactNode;
  /** Multi-day span bars (Outlook-style strip above the hourly area). */
  multiDayBarRow?: React.ReactNode;
}

export function CalendarTimeGrid<T>(props: ICalendarTimeGridProps<T>): JSX.Element {
  const {
    days,
    columnHeaderLabels,
    timedEventsByDay,
    getEventStart,
    getEventEndInclusive,
    renderEvent,
    multiDayBarRow
  } = props;

  const gridBodyHeightPx: number = (CALENDAR_GRID_TOTAL_MINUTES / 60) * CALENDAR_PX_PER_HOUR;
  const hours: number[] = [];
  for (let h: number = CALENDAR_GRID_START_HOUR; h < CALENDAR_GRID_END_HOUR; h++) {
    hours.push(h);
  }

  return (
    <div style={{ border: '1px solid #edebe9', background: '#fff', overflowX: 'auto' }}>
      {multiDayBarRow}
      <div style={{ display: 'flex', flexDirection: 'row', alignItems: 'stretch', minWidth: 0 }}>
        <div
          style={{
            width: `${TIME_GUTTER_WIDTH_PX}px`,
            flexShrink: 0,
            borderRight: '1px solid #edebe9',
            background: '#faf9f8'
          }}
        >
          <div
            style={{
              height: '40px',
              borderBottom: '1px solid #edebe9',
              boxSizing: 'border-box'
            }}
          />
          <div style={{ position: 'relative', height: `${gridBodyHeightPx}px` }}>
            {hours.map((hour: number) => {
              const label =
                hour === 0
                  ? '12 AM'
                  : hour < 12
                    ? `${hour} AM`
                    : hour === 12
                      ? '12 PM'
                      : `${hour - 12} PM`;
              return (
                <div
                  key={`tg-${hour}`}
                  style={{
                    position: 'absolute',
                    left: 0,
                    right: 4,
                    top: `${(hour - CALENDAR_GRID_START_HOUR) * CALENDAR_PX_PER_HOUR}px`,
                    fontSize: '11px',
                    color: '#605e5c',
                    textAlign: 'right',
                    lineHeight: 1,
                    transform: 'translateY(-50%)'
                  }}
                >
                  {label}
                </div>
              );
            })}
          </div>
        </div>
        <div style={{ flex: '1 1 auto', display: 'flex', flexDirection: 'row', minWidth: 0 }}>
          {days.map((day: Date, dayIndex: number) => {
            const dayEvents: T[] = timedEventsByDay[dayIndex] || [];
            const clipped = dayEvents
              .map((item: T) => {
                const clip = clipIntervalToDayGrid(day, getEventStart(item), getEventEndInclusive(item));
                if (!clip) {
                  return null;
                }
                return { item, startMs: clip.startMs, endMs: clip.endMs };
              })
              .filter((x): x is { item: T; startMs: number; endMs: number } => x !== null);

            const placed = assignTimeOverlapLanes(clipped);

            return (
              <div
                key={`day-col-${day.toISOString()}`}
                style={{
                  flex: '1 1 0',
                  minWidth: `${COLUMN_MIN_WIDTH_PX}px`,
                  borderRight: '1px solid #edebe9',
                  display: 'flex',
                  flexDirection: 'column'
                }}
              >
                <div
                  style={{
                    height: '40px',
                    padding: '6px 4px',
                    fontWeight: 600,
                    fontSize: '12px',
                    borderBottom: '1px solid #edebe9',
                    background: '#faf9f8',
                    boxSizing: 'border-box',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    textAlign: 'center'
                  }}
                >
                  {columnHeaderLabels[dayIndex] || ''}
                </div>
                <div
                  style={{
                    position: 'relative',
                    height: `${gridBodyHeightPx}px`,
                    background:
                      'repeating-linear-gradient(to bottom, #faf9f8 0px, #faf9f8 1px, transparent 1px, transparent ' +
                      `${CALENDAR_PX_PER_HOUR}px)`
                  }}
                >
                  {hours.map((hour: number) => (
                    <div
                      key={`hl-${dayIndex}-${hour}`}
                      aria-hidden
                      style={{
                        position: 'absolute',
                        left: 0,
                        right: 0,
                        top: `${(hour - CALENDAR_GRID_START_HOUR) * CALENDAR_PX_PER_HOUR}px`,
                        borderTop: '1px solid #e1dfdd',
                        pointerEvents: 'none'
                      }}
                    />
                  ))}
                  {placed.map((pe: { item: T; startMs: number; endMs: number; lane: number; laneCount: number }, evIdx: number) => {
                    const topMin: number = minutesFromGridStart(day, pe.startMs);
                    const endMin: number = minutesFromGridStart(day, pe.endMs);
                    const durMin: number = Math.max(endMin - topMin, 12);
                    const topPx: number = (topMin / 60) * CALENDAR_PX_PER_HOUR;
                    const heightPx: number = Math.max((durMin / 60) * CALENDAR_PX_PER_HOUR - EVENT_BLOCK_GAP_PX, 20);
                    const pct: number = 100 / pe.laneCount;
                    const leftPct: number = pe.lane * pct;
                    return (
                      <div
                        key={`ev-${dayIndex}-${evIdx}-${String((pe.item as any)?.id ?? evIdx)}`}
                        style={{
                          position: 'absolute',
                          left: `calc(${leftPct}% + 2px)`,
                          width: `calc(${pct}% - 4px)`,
                          top: `${topPx}px`,
                          height: `${heightPx}px`,
                          boxSizing: 'border-box',
                          borderRadius: '3px',
                          border: '1px solid rgba(0,0,0,0.12)',
                          overflow: 'hidden',
                          background: 'rgba(255,255,255,0.94)',
                          fontSize: '10px',
                          lineHeight: 1.2,
                          zIndex: 1
                        }}
                      >
                        {renderEvent(pe.item)}
                      </div>
                    );
                  })}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

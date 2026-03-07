export interface ICalendarSourceConfig {
  key: string;
  siteUrl: string;
  listTitle: string;
  displayName: string;
  color: string;
}

export interface ICalendarEvent {
  id: string;
  title: string;
  start: Date;
  end: Date;
  categories: string[];
  sourceKey: string;
  sourceName: string;
  sourceColor: string;
  location?: string;
  description?: string;
  isAllDay: boolean;
}

export interface ICalendarSourceParseResult {
  sources: ICalendarSourceConfig[];
  errors: string[];
}

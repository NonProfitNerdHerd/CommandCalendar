import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {
  ICalendarEvent,
  ICalendarSourceConfig,
  ICalendarSourceParseResult
} from '../models/CalendarModels';

const SOURCE_COLORS: string[] = [
  '#0078d4',
  '#107c10',
  '#5c2d91',
  '#d83b01',
  '#c239b3',
  '#0099bc',
  '#8764b8',
  '#038387'
];

export function parseCalendarSources(calendarSources: string): ICalendarSourceParseResult {
  const sources: ICalendarSourceConfig[] = [];
  const errors: string[] = [];
  const lines: string[] = calendarSources
    .split('\n')
    .map((line: string) => line.trim())
    .filter((line: string) => line.length > 0);

  lines.forEach((line: string, index: number) => {
    const parts: string[] = line.split('|').map((part: string) => part.trim());
    if (parts.length < 2) {
      errors.push(`Line ${index + 1} must include at least "siteUrl|listTitle".`);
      return;
    }

    const siteUrl: string = parts[0];
    const listTitle: string = parts[1];
    const displayName: string = parts[2] || listTitle;
    const color: string = parts[3] || SOURCE_COLORS[index % SOURCE_COLORS.length];

    if (!/^https?:\/\//i.test(siteUrl)) {
      errors.push(`Line ${index + 1} has an invalid site URL: "${siteUrl}".`);
      return;
    }

    sources.push({
      key: `${siteUrl}::${listTitle}`,
      siteUrl: siteUrl.replace(/\/$/, ''),
      listTitle,
      displayName,
      color
    });
  });

  return { sources, errors };
}

export class CalendarDataService {
  public constructor(private readonly _spHttpClient: SPHttpClient) {}

  public async getEvents(
    sources: ICalendarSourceConfig[],
    rangeStart: Date,
    rangeEnd: Date
  ): Promise<ICalendarEvent[]> {
    const eventsPerSource: ICalendarEvent[][] = await Promise.all(
      sources.map((source: ICalendarSourceConfig) =>
        this._getEventsFromSource(source, rangeStart, rangeEnd)
      )
    );

    return eventsPerSource
      .reduce((all: ICalendarEvent[], sourceEvents: ICalendarEvent[]) => all.concat(sourceEvents), [])
      .sort((a: ICalendarEvent, b: ICalendarEvent) => a.start.getTime() - b.start.getTime());
  }

  private async _getEventsFromSource(
    source: ICalendarSourceConfig,
    rangeStart: Date,
    rangeEnd: Date
  ): Promise<ICalendarEvent[]> {
    const escapedListTitle: string = source.listTitle.replace(/'/g, "''");
    const startIso: string = rangeStart.toISOString();
    const endIso: string = rangeEnd.toISOString();
    const filter: string = `EndDate ge datetime'${startIso}' and EventDate le datetime'${endIso}'`;
    const query: string =
      `$select=Id,Title,EventDate,EndDate,Category,Location,Description,fAllDayEvent&` +
      `$filter=${encodeURIComponent(filter)}&$top=5000`;

    let requestUrl: string =
      `${source.siteUrl}/_api/web/lists/getByTitle('${escapedListTitle}')/items?${query}`;
    const collectedEvents: ICalendarEvent[] = [];

    while (requestUrl) {
      const response: SPHttpClientResponse = await this._spHttpClient.get(
        requestUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata'
          }
        }
      );

      if (!response.ok) {
        throw new Error(
          `Unable to load "${source.displayName}" (${response.status} ${response.statusText}).`
        );
      }

      const payload: { value?: unknown[]; ['@odata.nextLink']?: string } = await response.json();
      const rows: unknown[] = payload.value || [];
      rows.forEach((row: unknown) => {
        const mappedEvent: ICalendarEvent | undefined = this._mapEvent(row, source);
        if (mappedEvent) {
          collectedEvents.push(mappedEvent);
        }
      });

      requestUrl = payload['@odata.nextLink'] || '';
    }

    return collectedEvents;
  }

  private _mapEvent(row: unknown, source: ICalendarSourceConfig): ICalendarEvent | undefined {
    if (!row || typeof row !== 'object') {
      return undefined;
    }

    const item: Record<string, unknown> = row as Record<string, unknown>;
    const start: Date = new Date(String(item.EventDate || ''));
    const end: Date = new Date(String(item.EndDate || ''));

    if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) {
      return undefined;
    }

    return {
      id: `${source.key}-${String(item.Id || '')}`,
      title: String(item.Title || '(Untitled Event)'),
      start,
      end: end.getTime() >= start.getTime() ? end : start,
      categories: this._extractCategories(item.Category),
      sourceKey: source.key,
      sourceName: source.displayName,
      sourceColor: source.color,
      location: item.Location ? String(item.Location) : undefined,
      description: item.Description ? String(item.Description) : undefined,
      isAllDay: Boolean(item.fAllDayEvent)
    };
  }

  private _extractCategories(rawCategory: unknown): string[] {
    if (!rawCategory) {
      return ['Uncategorized'];
    }

    const categoryText: string = String(rawCategory).trim();
    if (!categoryText) {
      return ['Uncategorized'];
    }

    if (categoryText.indexOf(';#') > -1) {
      const tokens: string[] = categoryText
        .split(';#')
        .map((token: string) => token.trim())
        .filter((token: string) => token.length > 0 && !/^\d+$/.test(token));
      return tokens.length > 0 ? tokens : ['Uncategorized'];
    }

    const commaTokens: string[] = categoryText
      .split(',')
      .map((token: string) => token.trim())
      .filter((token: string) => token.length > 0);

    return commaTokens.length > 0 ? commaTokens : ['Uncategorized'];
  }
}

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
    const query: string =
      '$select=Id,Title,EventDate,EndDate,Category,Location,Description,fAllDayEvent&' +
      '$orderby=EventDate asc&$top=5000';

    let requestUrl: string =
      `${source.siteUrl}/_api/web/lists/getByTitle('${escapedListTitle}')/items?${query}`;
    const collectedEvents: ICalendarEvent[] = [];

    while (requestUrl) {
      try {
        const payload: unknown = await this._requestJsonWithFallback(requestUrl);
        const rows: unknown[] = this._extractRows(payload);
        rows.forEach((row: unknown) => {
          const mappedEvent: ICalendarEvent | undefined = this._mapEvent(row, source);
          if (
            mappedEvent &&
            mappedEvent.end.getTime() >= rangeStart.getTime() &&
            mappedEvent.start.getTime() <= rangeEnd.getTime()
          ) {
            collectedEvents.push(mappedEvent);
          }
        });

        requestUrl = this._extractNextLink(payload);
      } catch (error) {
        const message: string =
          error instanceof Error ? error.message : 'Unknown error while loading events.';
        throw new Error(`Unable to load "${source.displayName}". ${message}`);
      }
    }

    return collectedEvents;
  }

  private async _requestJsonWithFallback(requestUrl: string): Promise<unknown> {
    const acceptHeaders: string[] = [
      'application/json;odata=nometadata',
      'application/json;odata=minimalmetadata',
      'application/json;odata=verbose',
      'application/json'
    ];

    let lastResponse: SPHttpClientResponse | undefined;
    for (const acceptHeader of acceptHeaders) {
      const response: SPHttpClientResponse = await this._spHttpClient.get(
        requestUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: acceptHeader
          }
        }
      );

      if (response.ok) {
        return response.json();
      }

      lastResponse = response;
      if (response.status !== 406) {
        break;
      }
    }

    throw new Error(
      `Unable to load calendar items (${lastResponse ? `${lastResponse.status} ${lastResponse.statusText}` : 'unknown error'}). ` +
      'Verify the calendar list title exactly matches the list display name and that you have read access.'
    );
  }

  private _extractRows(payload: unknown): unknown[] {
    if (!payload || typeof payload !== 'object') {
      return [];
    }

    const normalizedPayload: Record<string, unknown> = payload as Record<string, unknown>;
    if (Array.isArray(normalizedPayload.value)) {
      return normalizedPayload.value;
    }

    const dValue: unknown = normalizedPayload.d;
    if (dValue && typeof dValue === 'object') {
      const dObject: Record<string, unknown> = dValue as Record<string, unknown>;
      if (Array.isArray(dObject.results)) {
        return dObject.results;
      }
    }

    return [];
  }

  private _extractNextLink(payload: unknown): string {
    if (!payload || typeof payload !== 'object') {
      return '';
    }

    const normalizedPayload: Record<string, unknown> = payload as Record<string, unknown>;
    const odataNextLink: unknown = normalizedPayload['@odata.nextLink'];
    if (typeof odataNextLink === 'string') {
      return odataNextLink;
    }

    const dValue: unknown = normalizedPayload.d;
    if (dValue && typeof dValue === 'object') {
      const dObject: Record<string, unknown> = dValue as Record<string, unknown>;
      const legacyNext: unknown = dObject.__next;
      if (typeof legacyNext === 'string') {
        return legacyNext;
      }
    }

    return '';
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

    if (typeof rawCategory === 'object') {
      const categoryObject: Record<string, unknown> = rawCategory as Record<string, unknown>;
      const results: unknown = categoryObject.results;
      if (Array.isArray(results)) {
        const categories: string[] = results
          .map((entry: unknown) => String(entry || '').trim())
          .filter((entry: string) => entry.length > 0);
        return categories.length > 0 ? categories : ['Uncategorized'];
      }
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

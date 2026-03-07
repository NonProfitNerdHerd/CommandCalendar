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
    const potentialFourth: string = parts[3] || '';
    const potentialFifth: string = parts[4] || '';
    const color: string = /^#/i.test(potentialFourth)
      ? potentialFourth
      : SOURCE_COLORS[index % SOURCE_COLORS.length];
    const staffGroup: string = potentialFifth || (/^#/i.test(potentialFourth) ? 'General' : (potentialFourth || 'General'));

    if (!/^https?:\/\//i.test(siteUrl)) {
      errors.push(`Line ${index + 1} has an invalid site URL: "${siteUrl}".`);
      return;
    }

    sources.push({
      key: `${siteUrl}::${listTitle}`,
      siteUrl: siteUrl.replace(/\/$/, ''),
      listTitle,
      displayName,
      staffGroup,
      color
    });
  });

  return { sources, errors };
}

export class CalendarDataService {
  private readonly _displayFormUrlCache: Record<string, string> = {};

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

    const mergedEvents: ICalendarEvent[] = eventsPerSource
      .reduce((all: ICalendarEvent[], sourceEvents: ICalendarEvent[]) => all.concat(sourceEvents), [])
      .sort((a: ICalendarEvent, b: ICalendarEvent) => a.start.getTime() - b.start.getTime());

    const dedupedByKey: Record<string, ICalendarEvent> = {};
    mergedEvents.forEach((event: ICalendarEvent) => {
      const dedupeKey: string =
        `${event.sourceKey}|${event.title}|${event.start.toISOString()}|${event.end.toISOString()}|${event.itemUrl || ''}`;
      if (!dedupedByKey[dedupeKey]) {
        dedupedByKey[dedupeKey] = event;
      }
    });

    return Object.keys(dedupedByKey)
      .map((key: string) => dedupedByKey[key])
      .sort((a: ICalendarEvent, b: ICalendarEvent) => a.start.getTime() - b.start.getTime());
  }

  private async _getEventsFromSource(
    source: ICalendarSourceConfig,
    rangeStart: Date,
    rangeEnd: Date
  ): Promise<ICalendarEvent[]> {
    const displayFormUrl: string = await this._getDisplayFormUrl(source);
    const escapedListTitle: string = source.listTitle.replace(/'/g, "''");
    const query: string =
      '$select=Id,Title,EventDate,EndDate,Duration,Category,Location,Description,fAllDayEvent,fRecurrence,RecurrenceData,RecurrenceID&' +
      '$orderby=EventDate asc&$top=5000';

    let requestUrl: string =
      `${source.siteUrl}/_api/web/lists/getByTitle('${escapedListTitle}')/items?${query}`;
    const collectedEvents: ICalendarEvent[] = [];

    while (requestUrl) {
      try {
        const payload: unknown = await this._requestJsonWithFallback(requestUrl);
        const rows: unknown[] = this._extractRows(payload);
        rows.forEach((row: unknown) => {
          const mappedEvents: ICalendarEvent[] = this._mapEvents(
            row,
            source,
            rangeStart,
            rangeEnd,
            displayFormUrl
          );
          if (mappedEvents.length > 0) {
            mappedEvents.forEach((mappedEvent: ICalendarEvent) => {
              collectedEvents.push(mappedEvent);
            });
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

  private async _getDisplayFormUrl(source: ICalendarSourceConfig): Promise<string> {
    const cached: string = this._displayFormUrlCache[source.key];
    if (cached) {
      return cached;
    }

    const escapedListTitle: string = source.listTitle.replace(/'/g, "''");
    const requestUrl: string =
      `${source.siteUrl}/_api/web/lists/getByTitle('${escapedListTitle}')?$select=DefaultDisplayFormUrl`;
    const payload: unknown = await this._requestJsonWithFallback(requestUrl);
    const normalizedPayload: Record<string, unknown> = payload as Record<string, unknown>;

    let relativeDisplayFormUrl: string = '';
    if (typeof normalizedPayload.DefaultDisplayFormUrl === 'string') {
      relativeDisplayFormUrl = normalizedPayload.DefaultDisplayFormUrl;
    } else if (normalizedPayload.d && typeof normalizedPayload.d === 'object') {
      const legacyObject: Record<string, unknown> = normalizedPayload.d as Record<string, unknown>;
      if (typeof legacyObject.DefaultDisplayFormUrl === 'string') {
        relativeDisplayFormUrl = legacyObject.DefaultDisplayFormUrl;
      }
    }

    const absoluteUrl: string = this._toAbsoluteUrl(
      source.siteUrl,
      relativeDisplayFormUrl,
      source.listTitle
    );
    this._displayFormUrlCache[source.key] = absoluteUrl;
    return absoluteUrl;
  }

  private _toAbsoluteUrl(siteUrl: string, relativeUrl: string, listTitle: string): string {
    if (!relativeUrl) {
      return `${siteUrl}/Lists/${encodeURIComponent(listTitle)}/DispForm.aspx`;
    }

    if (/^https?:\/\//i.test(relativeUrl)) {
      return relativeUrl;
    }

    if (relativeUrl.charAt(0) === '/') {
      const parsedSiteUrl: URL = new URL(siteUrl);
      return `${parsedSiteUrl.origin}${relativeUrl}`;
    }

    return `${siteUrl}/${relativeUrl.replace(/^\//, '')}`;
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

  private _mapEvents(
    row: unknown,
    source: ICalendarSourceConfig,
    rangeStart: Date,
    rangeEnd: Date,
    displayFormUrl: string
  ): ICalendarEvent[] {
    if (!row || typeof row !== 'object') {
      return [];
    }

    const item: Record<string, unknown> = row as Record<string, unknown>;
    const recurrenceIdValue: string = String(item.RecurrenceID || '').trim();
    const isExpandedOccurrence: boolean =
      recurrenceIdValue.length > 0 && recurrenceIdValue.toLowerCase() !== 'null';
    const isRecurringMaster: boolean = Boolean(item.fRecurrence) && !isExpandedOccurrence;
    if (isRecurringMaster) {
      return this._expandRecurringEvent(
        item,
        source,
        rangeStart,
        rangeEnd,
        displayFormUrl
      );
    }

    const mappedEvent: ICalendarEvent | undefined = this._mapSingleEvent(
      item,
      source,
      displayFormUrl,
      isExpandedOccurrence
    );
    if (!mappedEvent) {
      return [];
    }

    if (
      mappedEvent.end.getTime() >= rangeStart.getTime() &&
      mappedEvent.start.getTime() <= rangeEnd.getTime()
    ) {
      return [mappedEvent];
    }

    return [];
  }

  private _mapSingleEvent(
    item: Record<string, unknown>,
    source: ICalendarSourceConfig,
    displayFormUrl: string,
    isRecurringInstance: boolean,
    occurrenceStart?: Date,
    occurrenceEnd?: Date
  ): ICalendarEvent | undefined {
    const fallbackStart: Date = new Date(String(item.EventDate || ''));
    const fallbackEnd: Date = new Date(String(item.EndDate || ''));
    const start: Date = occurrenceStart || fallbackStart;
    const durationOverrideMs: number | undefined = this._getDurationMsFromItem(item);
    let end: Date = occurrenceEnd || fallbackEnd;
    if (durationOverrideMs && (!occurrenceEnd || end.getTime() < start.getTime())) {
      end = new Date(start.getTime() + durationOverrideMs);
    }

    if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) {
      return undefined;
    }

    const itemId: string = String(item.Id || '');
    const eventId: string = isRecurringInstance
      ? `${source.key}-${itemId}-${start.toISOString()}`
      : `${source.key}-${itemId}`;
    const categories: string[] = this._extractCategories(item.Category);

    return {
      id: eventId,
      title: String(item.Title || '(Untitled Event)'),
      start,
      end: end.getTime() >= start.getTime() ? end : new Date(start.getTime() + 60 * 60 * 1000),
      categories,
      sourceKey: source.key,
      sourceName: source.displayName,
      sourceColor: source.color,
      itemUrl: this._buildItemUrl(displayFormUrl, itemId),
      location: item.Location ? String(item.Location) : undefined,
      description: this._toPlainText(item.Description ? String(item.Description) : ''),
      isAllDay: Boolean(item.fAllDayEvent),
      isRecurringInstance
    };
  }

  private _expandRecurringEvent(
    item: Record<string, unknown>,
    source: ICalendarSourceConfig,
    rangeStart: Date,
    rangeEnd: Date,
    displayFormUrl: string
  ): ICalendarEvent[] {
    const recurrenceXmlRaw: string = String(item.RecurrenceData || '').trim();
    const baseEvent: ICalendarEvent | undefined = this._mapSingleEvent(
      item,
      source,
      displayFormUrl,
      true
    );
    if (!baseEvent) {
      return [];
    }

    if (!recurrenceXmlRaw) {
      return [];
    }

    const parsedRule: IParsedRecurrenceRule | undefined = this._parseRecurrenceRule(
      recurrenceXmlRaw,
      baseEvent.start
    );
    if (!parsedRule) {
      return [];
    }

    const durationFromItemMs: number | undefined = this._getDurationMsFromItem(item);
    const fallbackDurationMs: number = Math.max(1, baseEvent.end.getTime() - baseEvent.start.getTime());
    const durationMs: number = durationFromItemMs || this._normalizeRecurringDuration(
      fallbackDurationMs,
      baseEvent.isAllDay
    );
    const seriesEndByRule: Date = parsedRule.windowEnd || rangeEnd;
    const seriesEnd: Date =
      seriesEndByRule.getTime() < rangeEnd.getTime() ? seriesEndByRule : rangeEnd;
    const occurrences: ICalendarEvent[] = [];
    let matchCount: number = 0;

    for (
      let cursorDate: Date = startOfDay(baseEvent.start);
      cursorDate.getTime() <= seriesEnd.getTime();
      cursorDate = addDays(cursorDate, 1)
    ) {
      if (!this._matchesRecurrenceDate(cursorDate, baseEvent.start, parsedRule)) {
        continue;
      }

      matchCount += 1;
      if (parsedRule.repeatInstances && matchCount > parsedRule.repeatInstances) {
        break;
      }

      const occurrenceStart: Date = setTimeFrom(baseEvent.start, cursorDate);
      const occurrenceEnd: Date = new Date(occurrenceStart.getTime() + durationMs);

      if (
        occurrenceEnd.getTime() < rangeStart.getTime() ||
        occurrenceStart.getTime() > rangeEnd.getTime()
      ) {
        continue;
      }

      const mappedEvent: ICalendarEvent | undefined = this._mapSingleEvent(
        item,
        source,
        displayFormUrl,
        true,
        occurrenceStart,
        occurrenceEnd
      );
      if (mappedEvent) {
        occurrences.push(mappedEvent);
      }
    }

    return occurrences;
  }

  private _getDurationMsFromItem(item: Record<string, unknown>): number | undefined {
    const rawDuration: unknown = item.Duration;
    if (typeof rawDuration === 'undefined' || rawDuration === '') {
      return undefined;
    }

    const durationMinutes: number = parseInt(String(rawDuration), 10);
    if (Number.isNaN(durationMinutes) || durationMinutes <= 0) {
      return undefined;
    }

    return durationMinutes * 60 * 1000;
  }

  private _normalizeRecurringDuration(durationMs: number, isAllDay: boolean): number {
    const minimumMs: number = isAllDay ? 24 * 60 * 60 * 1000 : 15 * 60 * 1000;
    const maximumMs: number = isAllDay ? 7 * 24 * 60 * 60 * 1000 : 24 * 60 * 60 * 1000;
    if (durationMs < minimumMs) {
      return minimumMs;
    }
    if (durationMs > maximumMs) {
      return minimumMs;
    }
    return durationMs;
  }

  private _buildItemUrl(displayFormUrl: string, itemId: string): string | undefined {
    if (!displayFormUrl || !itemId) {
      return undefined;
    }

    return displayFormUrl.indexOf('?') > -1
      ? `${displayFormUrl}&ID=${encodeURIComponent(itemId)}`
      : `${displayFormUrl}?ID=${encodeURIComponent(itemId)}`;
  }

  private _parseRecurrenceRule(recurrenceXmlRaw: string, masterStart: Date): IParsedRecurrenceRule | undefined {
    const recurrenceXml: string = decodeHtmlEntities(recurrenceXmlRaw);
    let documentRoot: Document;

    try {
      const parser: DOMParser = new DOMParser();
      documentRoot = parser.parseFromString(recurrenceXml, 'text/xml');
    } catch {
      return undefined;
    }

    const parserErrors: HTMLCollectionOf<Element> = documentRoot.getElementsByTagName('parsererror');
    if (parserErrors.length > 0) {
      return undefined;
    }

    const ruleNode: Element | null = documentRoot.getElementsByTagName('rule')[0] || null;
    if (!ruleNode) {
      return undefined;
    }

    const repeatNode: Element | null = ruleNode.getElementsByTagName('repeat')[0] || null;
    if (!repeatNode || !repeatNode.firstElementChild) {
      return undefined;
    }

    const repeatPatternNode: Element = repeatNode.firstElementChild;
    const repeatPatternName: string = repeatPatternNode.tagName.toLowerCase();
    const interval: number = this._extractInterval(repeatPatternNode, repeatPatternName);
    const weekDays: number[] = this._extractWeekDays(repeatPatternNode, masterStart.getDay());

    const windowEndNode: Element | null = ruleNode.getElementsByTagName('windowEnd')[0] || null;
    const repeatInstancesNode: Element | null =
      ruleNode.getElementsByTagName('repeatInstances')[0] || null;
    const repeatInstances: number | undefined = repeatInstancesNode
      ? parseInt((repeatInstancesNode.textContent || '').trim(), 10)
      : undefined;
    const windowEnd: Date | undefined = windowEndNode
      ? parseSharePointDate((windowEndNode.textContent || '').trim())
      : undefined;

    let monthOfYear: number | undefined;
    const monthAttribute: string | null = repeatPatternNode.getAttribute('month');
    if (monthAttribute) {
      monthOfYear = parseInt(monthAttribute, 10) - 1;
    }

    return {
      patternName: repeatPatternName,
      interval: interval > 0 ? interval : 1,
      weekDays,
      dayOfMonth: parseOptionalInteger(repeatPatternNode.getAttribute('day') || undefined),
      monthOfYear,
      weekdayOfMonth: (repeatPatternNode.getAttribute('weekdayOfMonth') || '').toLowerCase(),
      repeatInstances:
        repeatInstances && !Number.isNaN(repeatInstances) && repeatInstances > 0
          ? repeatInstances
          : undefined,
      windowEnd
    };
  }

  private _extractInterval(repeatPatternNode: Element, patternName: string): number {
    if (patternName === 'daily') {
      return parseOptionalInteger(repeatPatternNode.getAttribute('dayFrequency') || undefined) || 1;
    }

    if (patternName === 'weekly') {
      return parseOptionalInteger(repeatPatternNode.getAttribute('weekFrequency') || undefined) || 1;
    }

    if (patternName === 'monthly' || patternName === 'monthlybyday') {
      return parseOptionalInteger(repeatPatternNode.getAttribute('monthFrequency') || undefined) || 1;
    }

    if (patternName === 'yearly' || patternName === 'yearlybyday') {
      return parseOptionalInteger(repeatPatternNode.getAttribute('yearFrequency') || undefined) || 1;
    }

    return 1;
  }

  private _extractWeekDays(repeatPatternNode: Element, fallbackDay: number): number[] {
    const dayAttributes: Record<string, number> = {
      su: 0,
      mo: 1,
      tu: 2,
      we: 3,
      th: 4,
      fr: 5,
      sa: 6
    };
    const selected: number[] = [];
    Object.keys(dayAttributes).forEach((key: string) => {
      const attributeValue: string | null = repeatPatternNode.getAttribute(key);
      if (attributeValue && attributeValue.toUpperCase() === 'TRUE') {
        selected.push(dayAttributes[key]);
      }
    });

    if (selected.length === 0) {
      selected.push(fallbackDay);
    }

    return selected;
  }

  private _matchesRecurrenceDate(
    date: Date,
    seriesStart: Date,
    rule: IParsedRecurrenceRule
  ): boolean {
    if (date.getTime() < startOfDay(seriesStart).getTime()) {
      return false;
    }

    const daysDiff: number = diffDays(startOfDay(seriesStart), startOfDay(date));
    const monthsDiff: number = diffMonths(seriesStart, date);
    const yearsDiff: number = date.getFullYear() - seriesStart.getFullYear();

    if (rule.patternName === 'daily') {
      return daysDiff % rule.interval === 0;
    }

    if (rule.patternName === 'weekly') {
      if (rule.weekDays.indexOf(date.getDay()) === -1) {
        return false;
      }
      const weeksDiff: number = Math.floor(daysDiff / 7);
      return weeksDiff % rule.interval === 0;
    }

    if (rule.patternName === 'monthly') {
      if (monthsDiff < 0 || monthsDiff % rule.interval !== 0) {
        return false;
      }
      const expectedDay: number = rule.dayOfMonth || seriesStart.getDate();
      const maxDayInMonth: number = daysInMonth(date.getFullYear(), date.getMonth());
      const targetDay: number = Math.min(expectedDay, maxDayInMonth);
      return date.getDate() === targetDay;
    }

    if (rule.patternName === 'monthlybyday') {
      if (monthsDiff < 0 || monthsDiff % rule.interval !== 0) {
        return false;
      }
      return isNthWeekdayOfMonth(date, rule.weekDays[0], rule.weekdayOfMonth || 'first');
    }

    if (rule.patternName === 'yearly') {
      if (yearsDiff < 0 || yearsDiff % rule.interval !== 0) {
        return false;
      }
      const targetMonth: number =
        typeof rule.monthOfYear === 'number' ? rule.monthOfYear : seriesStart.getMonth();
      const targetDay: number = rule.dayOfMonth || seriesStart.getDate();
      return date.getMonth() === targetMonth && date.getDate() === targetDay;
    }

    if (rule.patternName === 'yearlybyday') {
      if (yearsDiff < 0 || yearsDiff % rule.interval !== 0) {
        return false;
      }
      const targetMonth: number =
        typeof rule.monthOfYear === 'number' ? rule.monthOfYear : seriesStart.getMonth();
      if (date.getMonth() !== targetMonth) {
        return false;
      }
      return isNthWeekdayOfMonth(date, rule.weekDays[0], rule.weekdayOfMonth || 'first');
    }

    return false;
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

  private _toPlainText(value: string): string {
    if (!value) {
      return '';
    }

    return value
      .replace(/<[^>]+>/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }
}

interface IParsedRecurrenceRule {
  patternName: string;
  interval: number;
  weekDays: number[];
  dayOfMonth?: number;
  monthOfYear?: number;
  weekdayOfMonth?: string;
  repeatInstances?: number;
  windowEnd?: Date;
}

function parseOptionalInteger(value: string | undefined): number | undefined {
  if (!value) {
    return undefined;
  }

  const parsed: number = parseInt(value, 10);
  return Number.isNaN(parsed) ? undefined : parsed;
}

function decodeHtmlEntities(value: string): string {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, '&');
}

function parseSharePointDate(value: string): Date | undefined {
  if (!value) {
    return undefined;
  }

  const parsedDate: Date = new Date(value);
  return Number.isNaN(parsedDate.getTime()) ? undefined : parsedDate;
}

function startOfDay(date: Date): Date {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0, 0);
}

function addDays(date: Date, days: number): Date {
  const nextDate: Date = new Date(date.getTime());
  nextDate.setDate(nextDate.getDate() + days);
  return nextDate;
}

function diffDays(start: Date, end: Date): number {
  const millisecondsPerDay: number = 24 * 60 * 60 * 1000;
  const utcStart: number = Date.UTC(start.getFullYear(), start.getMonth(), start.getDate());
  const utcEnd: number = Date.UTC(end.getFullYear(), end.getMonth(), end.getDate());
  return Math.floor((utcEnd - utcStart) / millisecondsPerDay);
}

function diffMonths(start: Date, end: Date): number {
  return (end.getFullYear() - start.getFullYear()) * 12 + (end.getMonth() - start.getMonth());
}

function daysInMonth(year: number, month: number): number {
  return new Date(year, month + 1, 0).getDate();
}

function setTimeFrom(sourceDateTime: Date, targetDate: Date): Date {
  return new Date(
    targetDate.getFullYear(),
    targetDate.getMonth(),
    targetDate.getDate(),
    sourceDateTime.getHours(),
    sourceDateTime.getMinutes(),
    sourceDateTime.getSeconds(),
    sourceDateTime.getMilliseconds()
  );
}

function isNthWeekdayOfMonth(date: Date, weekday: number, nth: string): boolean {
  if (date.getDay() !== weekday) {
    return false;
  }

  const occurrenceNumber: number = Math.floor((date.getDate() - 1) / 7) + 1;
  const normalizedNth: string = (nth || '').toLowerCase();

  if (normalizedNth === 'last') {
    const nextWeekSameWeekday: Date = new Date(date.getTime());
    nextWeekSameWeekday.setDate(nextWeekSameWeekday.getDate() + 7);
    return nextWeekSameWeekday.getMonth() !== date.getMonth();
  }

  if (normalizedNth === 'first') {
    return occurrenceNumber === 1;
  }
  if (normalizedNth === 'second') {
    return occurrenceNumber === 2;
  }
  if (normalizedNth === 'third') {
    return occurrenceNumber === 3;
  }
  if (normalizedNth === 'fourth') {
    return occurrenceNumber === 4;
  }

  return occurrenceNumber === 1;
}

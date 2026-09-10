export interface ICalendarDirectorySource {
  siteUrl?: string;
  list: string;
  listName?: string;
  displayName?: string;
}

export interface ICalendarDirectoryLink {
  label: string;
  url?: string;
  error?: string;
}

export interface ICalendarDirectoryResponse {
  ok: boolean;
  json(): Promise<any>;
}

/** Resolve against the source web, not the App Catalog or the page URL. */
export function absoluteCalendarUrl(path: string, webUrl: string): string {
  const url = new URL(path, webUrl.replace(/\/+$/, '') + '/');
  if (url.protocol !== 'https:' && url.protocol !== 'http:') {
    throw new Error('Unsupported calendar URL');
  }
  return url.href;
}

export async function resolveCalendarDirectoryLink(
  source: ICalendarDirectorySource,
  currentWebUrl: string,
  get: (url: string) => Promise<ICalendarDirectoryResponse>
): Promise<ICalendarDirectoryLink> {
  const label = source.displayName || source.listName || source.list || 'SharePoint Calendar';
  try {
    const webUrl = absoluteCalendarUrl(source.siteUrl || currentWebUrl, currentWebUrl).replace(/\/+$/, '');
    // Existing configurations store the immutable list GUID, even when a list is renamed.
    const listId = String(source.list || '').replace(/[{}]/g, '');
    if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(listId)) {
      throw new Error('Missing calendar list ID');
    }
    const response = await get(`${webUrl}/_api/web/lists('${listId}')?$select=DefaultViewUrl`);
    if (!response.ok) {
      throw new Error('Calendar lookup failed');
    }
    const body = await response.json();
    const metadata = body.d || body;
    if (typeof metadata.DefaultViewUrl !== 'string' || !metadata.DefaultViewUrl.trim()) {
      throw new Error('Calendar URL unavailable');
    }
    return { label, url: absoluteCalendarUrl(metadata.DefaultViewUrl, webUrl) };
  } catch (_) {
    // Do not invent a clickable path when the list is deleted or inaccessible.
    return { label, error: 'Calendar link unavailable. Check access and the configured source, then retry.' };
  }
}

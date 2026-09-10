import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import { ITimelineCalendarProps } from './ITimelineCalendarProps';
import { ICalendarItem } from './IConfigurationItems';
import { ICalendarDirectoryLink, ICalendarDirectorySource, resolveCalendarDirectoryLink } from './calendarDirectoryLinks';

export const CalendarDirectory: React.FC<ITimelineCalendarProps> = (props) => {
  const [links, setLinks] = React.useState<ICalendarDirectoryLink[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [retry, setRetry] = React.useState(0);
  // Property-pane editors can mutate source objects in place. Snapshot the saved fields.
  const sourcesJson = JSON.stringify((props.lists || []).map((source: ICalendarDirectorySource) => ({
    siteUrl: source.siteUrl,
    list: source.list,
    listName: source.listName,
    displayName: source.displayName
  })));
  const webUrl = props.context.pageContext.web.absoluteUrl;
  const client = props.context.spHttpClient;

  React.useEffect(() => {
    let cancelled = false;
    setLoading(true);
    const sources: ICalendarDirectorySource[] = JSON.parse(sourcesJson);
    Promise.all(sources.map((source) => resolveCalendarDirectoryLink(
      source, webUrl, (url) => client.get(url, SPHttpClient.configurations.v1)
    ))).then((resolved) => {
      if (!cancelled) {
        setLinks(resolved);
        setLoading(false);
      }
    }).catch(() => {
      if (!cancelled) {
        setLinks(sources.map((source) => ({
          label: source.displayName || source.listName || source.list,
          error: 'Calendar link unavailable. Please retry.'
        })));
        setLoading(false);
      }
    });
    return () => { cancelled = true; };
  }, [sourcesJson, webUrl, client, retry]);

  const outlookSources: ICalendarItem[] = props.calsAndPlans || [];
  const cellStyle: React.CSSProperties = { padding: '12px 0', borderBottom: '1px solid #edebe9', textAlign: 'left', verticalAlign: 'top' };

  return (
    <section aria-label="Calendar directory" style={{ padding: '16px', background: '#fff', border: '1px solid #edebe9', borderRadius: '6px' }}>
      <p>Configured data sources from the web part. Use the link to open each SharePoint calendar.</p>
      {loading ? <p role="status">Loading calendar links...</p> : (
        <>
          {links.length === 0 && outlookSources.length === 0 ? <p>No calendars are configured.</p> : (
            <table style={{ width: '100%', borderCollapse: 'collapse', tableLayout: 'fixed' }}>
              <thead><tr>
                <th scope="col" style={{ ...cellStyle, width: '38%' }}>Calendar Display Name</th>
                <th scope="col" style={cellStyle}>Link</th>
              </tr></thead>
              <tbody>
                {links.map((link, index) => <tr key={index}>
                  <td style={cellStyle}>{link.label}</td>
                  <td style={{ ...cellStyle, overflowWrap: 'anywhere' }}>
                    {link.url ? <a href={link.url} target="_blank" rel="noopener noreferrer">{link.url}</a> : link.error}
                  </td>
                </tr>)}
                {outlookSources.map((source, index) => <tr key={`outlook-${index}`}>
                  <td style={cellStyle}>{source.displayName || (source.persona && source.persona[0] && source.persona[0].mail) || 'Outlook Calendar'}</td>
                  <td style={cellStyle}>Open this calendar in Outlook.</td>
                </tr>)}
              </tbody>
            </table>
          )}
          {links.some((link) => !!link.error) && <button type="button" onClick={() => setRetry(retry + 1)} style={{ marginTop: '12px' }}>Retry links</button>}
        </>
      )}
    </section>
  );
};

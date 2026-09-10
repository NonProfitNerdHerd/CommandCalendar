# Command Calendar V2

SharePoint Framework (SPFx) web part that aggregates **SharePoint** list/classic calendar data and **Microsoft 365 / Outlook** calendars (including **group** calendars) into a single, filterable experience. This project extends the open-source [**Timeline Calendar**](https://github.com/spsprinkles/timeline-calendar) with a **tabbed UI**, **Gantt-style timelines**, **Outlook-like calendar grids**, **upcoming-items horizons**, **rich filtering**, and **print-to-PDF** for calendar views.

![SPFx 1.17.2](https://img.shields.io/badge/SPFx-1.17.2-green.svg)
![Node.js v16](https://img.shields.io/badge/Node.js-v16-green.svg)
![SPO](https://img.shields.io/badge/SharePoint%20Online-Compatible-green.svg)

---

## Summary

### Release 2.1.11.0 — Calendar directory URLs

- The **Calendar directory** tab reads the existing SharePoint calendar sources configured in the web part. It shows each display name and the full, clickable calendar URL, including the hostname.
- Links are resolved at runtime by querying the saved list GUID for SharePoint's `DefaultViewUrl`. The directory does not construct paths from display names or insert `/Lists/` or `calendar.aspx`. Legitimate `/Lists/` paths are preserved.
- Relative source URLs use the current web part's SharePoint web context. Each returned calendar path is resolved against its source web, including calendars on another site. There is no hardcoded SharePoint tenant or App Catalog URL. Absolute configured sources retain their original host.
- Existing sources are resolved without opening or saving the property pane again. The directory includes calendars with no events and is independent of the event filters. Failed lookups show an unavailable message with **Retry links**, rather than a guessed link.
- Outlook sources are listed by name with an instruction to open them in Outlook; this release resolves direct links for SharePoint sources. Outlook event loading is unchanged.
- The solution ID, feature ID, web part ID, serialized property schema, and data version are unchanged. Calendar configuration is read without rewriting it.
- Restored missing optional TypeScript declarations for the existing staff-section and title-search filter props so the production build compiles. This does not alter persisted settings or filter behavior.

#### Upgrade and test

1. Download **[Command-Calendar-V2.sppkg](../Command-Calendar-V2.sppkg)** (app version **2.1.11.0**, JavaScript package version **2.1.11**).
2. Upload it to the **same tenant or site collection App Catalog** as the existing app, replace the existing package, and deploy. Keep the current app and web parts installed so their saved calendar settings remain in place.
3. If the site-installed app offers an update in **Site contents**, apply that update. For tenant-wide deployment, existing web part instances use the updated assets after the package is deployed. See Microsoft's [deployment and update guidance](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/tenant-scoped-deployment).
4. Reload the existing page and open **Calendar directory**. Confirm that the configured display names remain and that the link text includes `https://` and the correct source hostname. Click each link and confirm it opens the intended calendar.
5. Test a calendar on the current site, one on another site, and a calendar whose URL does not contain `/Lists/`. Repeat in another test tenant with sources configured for that tenant. Deploying the package does not copy calendar data or grant cross-tenant access.

Automated validation: `npm run test:directory` checks dynamic tenant resolution, source-site selection, root sites, list renames, custom views, encoding, and lookup failures. `npm run package` compiles and bundles the production app and produces `sharepoint/solution/command-calendar-v2.sppkg`. The root download is a copy of that build output. Live navigation and preservation of the deployed page's settings must also be checked in SharePoint.

Build note: SPFx treats output on stderr as failure. This release uses the locked dependencies with `BROWSERSLIST_IGNORE_OLD_DATA=true` to suppress the outdated browser-data notice (in PowerShell, set `$env:BROWSERSLIST_IGNORE_OLD_DATA = 'true'` before running the package command). The existing gulp configuration disables legacy lint; the release checks are the URL regression suite and production TypeScript/bundle/package tasks.

SharePoint returns the list's actual server-relative default-view address through [`DefaultViewUrl`](https://learn.microsoft.com/en-us/previous-versions/office/sharepoint-visio/jj246956(v=office.15)); the directory turns that address into an absolute URL.

- **Multi-source timeline** — Combine events from SharePoint lists and Graph-backed calendars in one place.
- **Tabbed navigation** — Switch between **Calendar**, **Gantt Chart Timeline**, **Gantt Chart DABAL** (zoomed Gantt), and **30-60-90-120** (upcoming horizons) without leaving the web part.
- **Calendar views** — **Day**, **5-Day Week**, **7-Day Week**, and **Month**. Day and week views use a **time grid** (6:00 AM–8:00 PM) with events positioned by start/end time, duration-based height, and side-by-side layout for overlaps—similar to Outlook or Google Calendar—not a simple list.
- **Filtering & search** — **Category**, **Staff Section**, and **Calendar** filters with **cascading** behavior, plus **title search** to narrow events quickly.
- **Readable Gantt** — Improved **label spacing** and layout logic to reduce overlap; **pagination** and **zoom** variants for long schedules.
- **Calendar display names** — Friendly names for connected calendars where configured.
- **Print / PDF** — **Print PDF** captures the **current calendar view** (grid layout and event placement) via a browser snapshot; includes a **point-in-time disclaimer** and generation timestamp in the PDF footer.

For web part property-pane options and advanced JSON configuration, the upstream [**user/setup guide (wiki)**](https://github.com/spsprinkles/timeline-calendar/wiki) remains the best reference for shared Timeline Calendar behaviors.

---

## User-facing features

### Tabs

| Tab | Purpose |
| --- | ------- |
| **Calendar** | Month grid or day/week **time grids** with navigation (previous/next/today), view switcher, and **Print PDF**. |
| **Gantt Chart Timeline** | Full-width Gantt-style timeline with pagination and layout tuned for readable labels. |
| **Gantt Chart DABAL** | Alternate Gantt view with different zoom/scale (suited to dense or long-range planning). |
| **30-60-90-120** | Upcoming-items lists grouped by **30 / 60 / 90 / 120** day horizons. |
| **Calendar directory** | Configured calendar display names and full SharePoint calendar links resolved dynamically from each source site. |

### Calendar views and defaults

- View buttons (left to right): **Day** → **5-Day Week** → **7-Day Week** → **Month**.
- Opening the **Calendar** tab defaults to **Month** view.
- **Day** shows one column; **5-Day** and **7-Day** show five or seven day columns.
- **Time grid** (Day / 5-Day / 7-Day):
  - Vertical axis from **6:00 AM** to **8:00 PM**.
  - Hour lines are evenly spaced; events align to their scheduled times.
  - Event **height reflects duration**; **overlapping** events are **stacked in columns** within the same day.
- **Month** view uses a classic **month grid** (weeks × days).

### Filters

- **Category** — Filter by event category (with legend when applicable).
- **Staff Section** — Narrow by staff/section metadata where used in your configuration.
- **Calendar** — Limit to specific connected calendars; behavior is **cascading** with other filters so selections stay consistent.
- **Title search** — Free-text filter on event titles.

### Print PDF (Calendar tab only)

- Click **Print PDF** to generate a **downloadable PDF** of what you see in the calendar area (title + active view).
- The export is **not** a plain event list: it preserves **grid structure**, **event positions**, and **spacing** as rendered on screen (rasterized via [html2canvas](https://html2canvas.hertzen.com/) and packaged with [jsPDF](https://github.com/parallax/jsPDF)).
- Every PDF includes footer text: *This document is a point-in-time snapshot generated on [date/time] and may not reflect any changes made after that moment.*

> **Note:** Very tall views (for example a dense **Month** grid) are scaled to fit the PDF page; text may appear smaller than on screen. If you need poster-size or multi-page exports, that would be a future enhancement.

---

## Developers

### Prerequisites

- **Node.js 16.x** (SPFx 1.17.x requirement). An `.nvmrc` is provided in this folder for [nvm](https://github.com/nvm-sh/nvm) / compatible version managers.

### Commands

```bash
cd command-calendar-v2
npm install
npm run build      # debug bundle
npm run package    # ship build + produce .sppkg for the App Catalog
```

The SharePoint package is written to:

`sharepoint/solution/command-calendar-v2.sppkg`

Upload that file to your tenant **App Catalog** (Apps for SharePoint), replace an existing package if needed, and deploy. Approve **Microsoft Graph** permission requests in the SharePoint admin center if prompted.

### Notable dependencies (calendar PDF)

- **html2canvas** — DOM snapshot for print fidelity.
- **jspdf** (pinned to **2.x** for SPFx/webpack compatibility; v3 pulled dependencies that failed in this toolchain).

### Source layout (high level)

| Area | Role |
| --- | ---- |
| `TimelineCalendarTabbed.tsx` | Tabs, filters, calendar chrome, **Print PDF** trigger, view state. |
| `CalendarTimeGrid.tsx` + `calendarTimeGridLayout.ts` | Time grid rendering, 6 AM–8 PM geometry, overlap lanes. |
| `calendarPdfExport.ts` | PDF generation, footer disclaimer, reserved footer space. |
| `ganttSchedulingLayout.ts` | Gantt label/row layout helpers. |

---

## Graph API permissions

The web part requests the same **delegated** Microsoft Graph scopes as upstream Timeline Calendar. Approve them in the SharePoint admin experience (or Azure AD) or calendar and directory features will be limited.

| Permission | Reason |
| ---------- | ------ |
| **User.Read** | Resolve the current user’s group memberships (e.g. M365 groups for calendar pickers). |
| **User.Read.All** | Search users and shared mailboxes in people-style pickers. |
| **Group.Read.All** | Search groups and read group calendar events the user can access. |
| **Calendars.Read.Shared** | Read calendars shared with the current user. |

---

## Content Security Policy (CSP)

Loading external scripts in SharePoint Online is subject to [CSP / trusted script sources](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/content-securty-policy-trusted-script-sources). This web part uses the **Monaco Editor** for advanced JSON editing. It tries CDNs in order (unless your tenant uses a variant such as DoD365-Sec, which is handled differently):

1. `https://cdnjs.cloudflare.com/`
2. `https://cdn.jsdelivr.net/`

Add one or both (with the **trailing `/`**) under **Trusted script sources** in the admin portal, per [Microsoft’s guidance](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/content-securty-policy-trusted-script-sources#managing-the-content-security-policy-rules-in-sharepoint-online).

For a narrower allow-list, you can scope to Monaco’s folder only (again include the trailing `/`):

1. `https://cdnjs.cloudflare.com/ajax/libs/monaco-editor/`
2. `https://cdn.jsdelivr.net/npm/monaco-editor@0.47.0/`

These are **base paths** for scripts the editor loads, not pages you browse to. If no trusted source is configured, Monaco may fail to load and a **plain textarea** fallback is used (as in upstream Timeline Calendar).

See also the [upstream wiki / project](https://github.com/spsprinkles/timeline-calendar/wiki) for property and environment notes.

---

## Upstream & versioning

- **Upstream project:** [spsprinkles/timeline-calendar](https://github.com/spsprinkles/timeline-calendar)  
- **This package:** `command-calendar-v2` (see `package.json` / `config/package-solution.json` for current version).  
- Historical release notes for the base product: [timeline-calendar releases](https://github.com/spsprinkles/timeline-calendar/releases).

---

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

Upstream Timeline Calendar is similarly provided as-is; this fork adds features on top of that baseline.

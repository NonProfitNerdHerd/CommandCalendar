# Command Calendar (SPFx)

Command Calendar is an SPFx React web part that aggregates events from multiple SharePoint calendar lists and renders them in three switchable tab views:

1. **Gantt View** (timeline bars across upcoming days)
2. **Calendar View** (month grid with event markers)
3. **30/60/90/120 View** (quad horizon cards)

A **multi-select Category filter** applies globally to all views.

---

## SharePoint Framework Version

- **SPFx 1.22.2**

---

## Features

- Pulls events from **multiple SharePoint site calendars**
- Configurable source list in property pane
- Tabbed navigation between all required views
- Cross-view category multi-filtering
- Source legend and manual refresh support
- Production package output (`.sppkg`) ready for App Catalog deployment

---

## Calendar Source Configuration

In the web part property pane, configure **Calendar sources** as one source per line:

```text
siteUrl|calendarListTitle|Display Name|Staff Group
```

Examples:

```text
https://contoso.sharepoint.com/sites/command|Calendar|Command Site|G1
https://contoso.sharepoint.com/sites/ops|Operations Calendar|Ops|G2
https://contoso.sharepoint.com/sites/projects|Project Calendar|Projects|G357
```

Only `siteUrl|calendarListTitle` is required. `Display Name` and `Staff Group` are optional.

Optional advanced mapping for Gantt swim lanes:

```text
Order|Lane Name|Calendar Display Name,Another Calendar
```

Example:

```text
1|G1|Command Site,Ops
2|G2|Projects
```

Optional category coloring:

```text
Category|#Color
```

---

## Build and Package

From the `command-calendar` folder:

```bash
npm install
npm run build
```

The deployable package is created at:

```text
sharepoint/solution/command-calendar.sppkg
```

---

## Deploy to SharePoint

1. Upload `sharepoint/solution/command-calendar.sppkg` to your tenant App Catalog.
2. Approve/deploy the app (tenant-wide deployment is enabled in this solution).
3. Add the app to the target site (if needed by your tenant policy).
4. Edit a modern SharePoint page and add the **CommandCalendar** web part.
5. Open property pane and configure:
   - Header description
   - Look-back / look-ahead days
   - Calendar source lines

---

## Development

Start local workbench/dev server:

```bash
npm run start
```

---

## Notes

- Users need read access to the configured source calendars.
- The web part uses standard SharePoint REST APIs through SPFx `SPHttpClient`.
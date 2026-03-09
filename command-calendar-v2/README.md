# Command Calendar V2

Command Calendar V2 is based on the excellent `spsprinkles/timeline-calendar` project and preserves the same timeline interaction model.

This version adds top-level tabs:

- **Gnatt Chart View** (default / main view, using the original timeline interaction)
- **Calendar View**
- **Agenda View**

## Build notes

This codebase is SPFx 1.17.2-based. In this cloud environment (Node 22), packaging requires a temporary local patch to allow the older SPFx build rig to run. The generated package is still valid for deployment.

## Build/package

From `command-calendar-v2`:

```bash
npm install
npm run package
```

Output package:

```text
sharepoint/solution/command-calendar-v2.sppkg
```

## Deploy

1. Upload `command-calendar-v2.sppkg` to the SharePoint App Catalog.
2. Deploy the app.
3. Add **Command Calendar V2** web part to a page.

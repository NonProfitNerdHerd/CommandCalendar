const assert = require('assert').strict;
const fs = require('fs');
const path = require('path');
const ts = require('typescript');
const Module = require('module');

const filename = path.resolve(__dirname, '../src/webparts/timelineCalendar/components/calendarDirectoryLinks.ts');
const compiled = ts.transpileModule(fs.readFileSync(filename, 'utf8'), {
  compilerOptions: { target: ts.ScriptTarget.ES2018, module: ts.ModuleKind.CommonJS }
});
const service = new Module(filename, module);
service._compile(compiled.outputText, filename);
const { resolveCalendarDirectoryLink } = service.exports;
const list = '12345678-1234-1234-1234-123456789abc';
let passed = 0;

async function check(name, run) {
  await run();
  passed++;
  console.log(`PASS ${name}`);
}

async function main() {
  for (const tenant of ['https://tenant-a.sharepoint.com', 'https://tenant-b.sharepoint.us']) {
    await check(`actual path without Lists on ${tenant}`, async () => {
      const source = Object.freeze({ siteUrl: '/sites/CRM', list, listName: 'Renamed Calendar', displayName: 'G1' });
      let requested;
      const result = await resolveCalendarDirectoryLink(source, `${tenant}/sites/Hosting`, async (url) => {
        requested = url;
        return { ok: true, json: async () => ({ DefaultViewUrl: '/sites/CRM/Events/calendar.aspx' }) };
      });
      assert.equal(requested, `${tenant}/sites/CRM/_api/web/lists('${list}')?$select=DefaultViewUrl`);
      assert.deepEqual(result, { label: 'G1', url: `${tenant}/sites/CRM/Events/calendar.aspx` });
      assert.equal(source.listName, 'Renamed Calendar');
    });
  }

  const cases = [
    ['valid Lists path and custom view', '/sites/CRM', '/sites/CRM/Lists/Original%20Name/custom.aspx', 'https://tenant.sharepoint.com/sites/CRM/Lists/Original%20Name/custom.aspx'],
    ['root site', '/', '/Events/calendar.aspx', 'https://tenant.sharepoint.com/Events/calendar.aspx'],
    ['current web fallback', undefined, 'Events/calendar.aspx', 'https://tenant.sharepoint.com/sites/Hosting/Events/calendar.aspx'],
    ['web-relative source', 'Child', 'Events/calendar.aspx', 'https://tenant.sharepoint.com/sites/Hosting/Child/Events/calendar.aspx'],
    ['absolute source on another host', 'https://other.sharepoint.com/sites/Source/', '/sites/Source/Events/calendar.aspx', 'https://other.sharepoint.com/sites/Source/Events/calendar.aspx'],
    ['absolute metadata URL', '/sites/CRM', 'https://tenant.sharepoint.com/sites/CRM/Calendar/custom.aspx?view=1', 'https://tenant.sharepoint.com/sites/CRM/Calendar/custom.aspx?view=1'],
    ['spaces and existing encoding', '/sites/CRM', '/sites/CRM/Team Events/view.aspx?label=A%20B', 'https://tenant.sharepoint.com/sites/CRM/Team%20Events/view.aspx?label=A%20B']
  ];
  for (const [name, siteUrl, returnedPath, expected] of cases) {
    await check(name, async () => {
      const result = await resolveCalendarDirectoryLink({ siteUrl, list: `{${list}}`, listName: 'Existing source' }, 'https://tenant.sharepoint.com/sites/Hosting', async () => ({
        ok: true, json: async () => ({ d: { DefaultViewUrl: returnedPath } })
      }));
      assert.equal(result.url, expected);
      assert.equal(result.label, 'Existing source');
    });
  }

  for (const [name, get] of [
    ['access denied or deleted list', async () => ({ ok: false })],
    ['network failure', async () => { throw new Error('Offline'); }],
    ['missing metadata', async () => ({ ok: true, json: async () => ({}) })],
    ['unsafe URL', async () => ({ ok: true, json: async () => ({ DefaultViewUrl: 'javascript:alert(1)' }) })],
    ['invalid JSON', async () => ({ ok: true, json: async () => { throw new Error('Invalid JSON'); } })]
  ]) {
    await check(name, async () => {
      const result = await resolveCalendarDirectoryLink({ siteUrl: '/sites/CRM', list }, 'https://tenant.sharepoint.com', get);
      assert.equal(result.url, undefined);
      assert.ok(result.error);
    });
  }
  await check('invalid list ID never issues a request', async () => {
    let requests = 0;
    const result = await resolveCalendarDirectoryLink({ siteUrl: '/', list: 'Events' }, 'https://tenant.sharepoint.com', async () => {
      requests++;
      return { ok: false };
    });
    assert.equal(requests, 0, 'A title must not be substituted for the saved list GUID');
    assert.ok(result.error);
    assert.equal(result.url, undefined);
  });
  console.log(`${passed} directory checks passed.`);
}

main().catch((error) => { console.error(error); process.exitCode = 1; });

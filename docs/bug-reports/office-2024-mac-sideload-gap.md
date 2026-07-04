# Bug: `office-addin-debugging` v5.x / v6.x cannot sideload dev add-ins on Office LTSC 2024 Mac

## Summary

On Office LTSC Standard for Mac 2024 (version 16.110.2, build 26062818),
`office-addin-debugging start` (5.1.6 and 6.1.1) successfully registers a
catalog node in the SQL-backed `MicrosoftRegistrationDB.reg`, but the
add-in never appears in the Office UI, no `WefProcess` is spawned, no
HTTP request reaches the dev server, and no log line mentions the add-in
GUID or the dev server URL.

## Environment

- Office LTSC Standard for Mac 2024, version 16.110.2, build 26062818
- macOS 15.x (tested on 15.5)
- `office-addin-debugging@5.1.6` and `6.1.1` (both reproduce)
- `office-addin-dev-settings@2.3.6` and `3.1.1` (both produce byte-equivalent Mac write path)
- Excel process is the only Office process; no WefProcess child is ever spawned

## What we observed

### Registration half-works

After `npx office-addin-debugging start manifest.xml --app excel ...` runs
on LTSC 2024 Mac, the catalog node IS written to
`~/Library/Group Containers/UBF8T346G9.Office/MicrosoftRegistrationDB/MicrosoftRegistrationDB_19413899015.reg`
at `HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Wef\Providers\<catalog-hash>`.
Values written:

- `Entitlements` = REG_QWORD FILETIME
- `UniqueId` = `"developer"`

What is **not** written (compare to working Omex catalog `geLWvPTNy9pxk4xxzFEGHw==`):

- `BlockForMinor` = 0 (REG_DWORD)
- `EntitlementsInvalid` = `{<guid>}` (REG_SZ)
- `AppStates` child node + per-addin value (REG_QWORD FILETIME)
- `AddinLifecycle` entries (`OMEX_<catalog>_<addin>_<N>` = lifecycle action)

### On-disk layout is empty

Working Omex catalogs have a full on-disk layout at
`~/Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft/Office/16.0/Wef/{EC993EB3-E226-5544-8239-1F3370D3A9CD}/Omex/<catalog-hash>/`
with 10 subdirectories (`AppStates`, `AppDetails`, `ExtendedManifest`,
`ExtendedManifestResources`, `Html`, `Javascript`, `Manifests`,
`Metadata`, `Tokens`, `TrustedApps`, `WellKnownResources`) and
per-addin files. The dev catalog has **no on-disk directory at all**.

The legacy `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`
path that `office-addin-dev-settings@2.x` writes to (`fs.ensureLinkSync`)
is also empty on a stock LTSC 2024 Mac install.

### CustomUIValidationCache has the entry

The manifest IS being validated — the SQL table has
`e504fb41-..._developer.Microsoft.Excel.Workbook = 1099998353` (REG_DWORD).
The failure is downstream in the WEF on-disk init step.

### After manually backfilling all 5 missing preconditions

We manually created the full Omex-equivalent structure for the dev
catalog (BlockForMinor, EntitlementsInvalid, AppStates, AddinLifecycle,
on-disk 10 subdirs + per-addin files including the manifest XML).
**Excel still does not spawn a WefProcess and still does not log any
reference to the dev catalog.** This suggests the WEF scanner has an
additional gate we have not identified — possibly related to
`UniqueId="developer"` (Omex uses the user's LiveId, not `"developer"`),
or a hidden registry key, or a signed manifest requirement.

## Why we think this is a tooling gap, not a user error

1. The same project on a Microsoft 365 Excel install sideloads normally
   (dev catalog gets auto-injected as a custom ribbon tab).
2. `office-addin-debugging@5.1.6` writes the minimum required SQL rows
   and the legacy `~/Documents/wef/` symlink. LTSC 2024 Mac ignores
   the legacy path entirely, and does not promote the dev catalog
   through the WEF on-disk init pipeline.
3. Manually backfilling the Omex-equivalent layout is insufficient —
   the dev catalog remains invisible to LTSC 2024 Mac's WEF scanner.

## Reproduction

```bash
# 1. New project
mkdir repro && cd repro
# (initialize a taskpane add-in via Yo Office with manifest id e504fb41-...)

# 2. Start dev server
npm run dev-server &  # listens on https://127.0.0.1:3000

# 3. Sideload
./node_modules/.bin/office-addin-debugging start manifest.xml \
  --app excel --source-bundle-url-host 127.0.0.1 --source-bundle-url-port 3000 --no-debug

# 4. Excel opens. Wait 30s.
# 5. Observe: no WefProcess child of Microsoft Excel.
# 6. Dev server logs: zero requests from WefProcess.
# 7. ~/Library/Containers/com.microsoft.Excel/Data/Library/Logs/Diagnostics/EXCEL/Primary*.log
#    contains zero references to the add-in GUID.
# 8. SQL: catalog is registered (2 values), but no on-disk mirror.
# 9. Insert > Add-ins > My Add-ins in the Excel ribbon does NOT list
#    the dev add-in.
```

## Suggested fix

`office-addin-debugging` should detect LTSC 2024+ Mac installs (e.g.
`Excel.app/Contents/Info.plist` CFBundleVersion >= `16.110`) and write
the new on-disk layout + full SQL row set:

1. Create `Wef/{EC993EB3-...}/<catalog-hash>/` (no `Omex/` parent) with
   the 10 subdirectories and per-addin files (manifest XML, AppStates
   XML, Tokens XML, Metadata JSON, Html redirector, empty TrustedApps
   marker).
2. Write `BlockForMinor=0` (REG_DWORD).
3. Write `EntitlementsInvalid={unique-guid}` (REG_SZ). The exact GUID
   for the dev path needs to be sourced from the Office dev-channel
   entitlement service.
4. Write `AppStates` child node + `<aid>_<ver>` = FILETIME (REG_QWORD).
5. Write `AddinLifecycle` entries `OMEX_<catalog>_<aid>_<1|2|4>` with
   the lifecycle action codes (1=install, 2=uninstall, 4=update).
6. Update `UniqueId` to match the user's identity, or add a fallback
   scan that matches the dev catalog regardless of UniqueId value.

## What we tried and ruled out

- `office-addin-debugging@5.1.6` → `6.1.1` upgrade: no change.
  `dev-settings-mac.js` is byte-equivalent between `office-addin-dev-settings@2.3.6`
  and `3.1.1` (verified by direct tarball diff — only `tslib` import
  syntax changes, no behavioral diff).
- `DebugAddins=1` registry flip at `HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Options`:
  no change in dev catalog visibility.
- `WEFOfficeAddinTrustedDomains` plist key: dev server reachable, not
  the bottleneck.
- Disabling system proxy (Clash at 127.0.0.1:7897): no change.
- Re-registering after deleting legacy `Documents/wef`: no change.
- Rebooting Excel: no change.
- Manual backfill of all 5 preconditions (SQL + on-disk): WefProcess
  still doesn't spawn.

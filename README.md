# AVD Usage Tracker

Automates the daily AVD uptime tracking that was being done by hand in
`Daily_and_Monthly_Usage_*.xlsx`. It reads each VM's powered-on time from Azure
Monitor and writes `"H hrs M mins"` into the day's column, flagging any VM that
ran past the scheduled shutdown with `(Extended)` + a yellow highlight.

## How uptime is measured

An Azure VM only emits platform metrics while it is **powered on**. The tool
queries the `Percentage CPU` metric at 1-minute (`PT1M`) granularity over the UTC
day and counts the non-null buckets — that count is the number of minutes the VM
was running. This:

- needs **no agent** inside the VM,
- naturally **sums multiple on/off cycles** in the same day,
- counts up to "now" if the VM is **still running**,
- is timezone-independent for the number itself (only the day window is UTC).

This is more accurate than a manual spot-check, which only captures uptime up to
the moment someone looked.

## "Extended" rule

Shutdown is driven by Azure DevOps pipelines (per team/region/shift), so there is
no single Azure schedule to read. Instead, a row is flagged `(Extended)` when its
uptime exceeds a configurable threshold (default **540 min = 9h**). Tune it in
`config.json` globally or per region / resource-group / VM:

```json
"extendedThresholdMinutes": 540,
"thresholdOverrides": {
  "byRegion": { "westeurope": 540 },
  "byResourceGroup": { "RRWVDS-RG1": 540 },
  "byVm": { "rrwvd-8": 600 }
}
```

Most specific match wins: `byVm` > `byResourceGroup` > `byRegion` > default.

## Setup (once)

```bash
cd eng/Scripts/AVDs/usage-tracker
npm install            # installs exceljs
az login               # to the tenant that owns the ReefReview_MCA_WVDs subscription
```

Edit `config.json` → set `workbook` to the real path of the usage spreadsheet.

## Run

```bash
# Dry run — compute + print, touch nothing (do this first):
node update-avd-usage.js

# Update the workbook in place (a timestamped .bak is written first):
node update-avd-usage.js --write

# Write to a separate copy, leaving the original untouched:
node update-avd-usage.js --out C:/AVD/usage-2026-06-30.xlsx

# Reconcile the UTC day that just ended (recommended for the scheduled run):
node update-avd-usage.js --write --yesterday

# Target a specific UTC day (default = today):
node update-avd-usage.js --write --date 2026-06-29

# Override workbook / sheet:
node update-avd-usage.js --write --workbook "C:/AVD/usage.xlsx" --sheet "Jun 26"
```

The sheet is auto-picked from the target month (e.g. `Jun 26`). The day's column
is found by matching the date in row 1; if today isn't there yet, a new column is
**appended** at the right (header date + `Daily/Avg time` sub-header), inheriting
the previous day's formatting. Reason columns are never auto-filled — that stays a
human decision.

## Deleted / deprovisioned VMs (Activity Log fallback)

When a sheet row's VM is no longer in the live inventory (deprovisioned), the
platform-metrics API returns `ResourceNotFound`. Instead of skipping it, the tool
falls back to the **Azure Activity Log**, which keeps the VM's `start` / `deallocate`
/ `delete` events for ~90 days. It reconstructs uptime by pairing those events:

- **multiple start triggers in a day are summed** — every powered-on interval counts;
- it checks the **end-of-day deallocate** actually happened — if the VM was still
  running at the UTC-day boundary it's reported as **not deallocated** (shutdown
  pipeline missed it / it crossed midnight);
- reconstructed cells get an Excel **note** ("Reconstructed from Activity Log …") so
  the value is auditable, and the run prints `via activity-log` + a `recovered` count.

Configured under `deletedVmLookup` in `config.json` (which resource groups to search).
If the VM was deleted longer ago than the retention window (or never ran), the row is
reported as `no Activity Log data` and left untouched.

## Not handled (left untouched, reported)

- `rrsgncusg-win1s…win6s` rows — these are **not** plain VMs in the subscription
  (scaling-plan / ephemeral hosts; the sheet sometimes shows "VM not found").
  Configured via `skipPrefixes` in `config.json`.

## Scheduling (important — run AFTER the day closes)

A live "today" run only captures uptime up to the moment it runs, so anything after
it (late workers, an extension till 21:00) is missed and never corrected. Instead,
run **once after the UTC day has ended, targeting that closed day**:

```bash
node update-avd-usage.js --write --yesterday
```

`--yesterday` uses the **full 00:00–24:00 UTC window** of the day that just ended,
so the figure is final and includes any extension. It overwrites that day's column
in place and applies the `(Extended)` highlight. The run log shows each VM's detected
boot–stop window (e.g. `[07:30Z-20:16Z]`) so the real stop time is auditable; a VM
that was still running at the UTC boundary is marked `ran to/over midnight`.

Recommended trigger: **daily ~00:45 UTC** (a buffer after shutdown pipelines + metric
lag) via Windows Task Scheduler or an Azure DevOps scheduled pipeline.

> Cross-midnight sessions are split at the UTC boundary — each day's column gets its
> own slice. If a whole overnight session should instead be attributed to one day,
> that rule needs defining.

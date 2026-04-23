# PaperCut MF — Manager Email Alert on Over-Threshold Print Jobs

Sends an email to a user's line manager (per Active Directory) when they submit a print job that exceeds a configurable page threshold.

## Components

| File | Role |
|------|------|
| `manager-email-sync.ps1` | Scheduled PowerShell script. Reads each AD user's `manager` attribute, resolves the manager's `mail` attribute, and writes it into a PaperCut MF global config key. |
| `papercut MF print script.txt` | PaperCut MF print script (`printJobHook`). When a job exceeds the page threshold, reads the manager email for the submitting user from the config key and sends a notification. Falls back to a configured email address if no manager email is set. |

## How it works

1. **PowerShell sync (scheduled, e.g. nightly)**
   - Enumerates enabled AD users with a `manager` attribute set.
   - For each user, resolves the manager's `mail`.
   - Writes a config key via `server-command set-config`:
     ```
     script.user-defined.user-custom-property.manager-email.<sAMAccountName>
     ```
   - Logs results to `C:\PaperCutScripts\Logs\sync-manager-email-YYYY-MM-DD.log`.
2. **Print script (per job)**
   - Waits for job analysis to complete.
   - If `totalPages` is at or below the threshold, logs and exits.
   - Otherwise reads:
     ```js
     inputs.utils.getProperty("user-custom-property.manager-email." + inputs.user.username)
     ```
   - If empty, falls back to `FALLBACK_EMAIL`.
   - Sends an email to the resolved address with job details.

## Setup

### 1. PowerShell sync script

**Requirements**
- Runs on the PaperCut MF application server (so it can call `server-command.exe` locally).
- PowerShell ActiveDirectory module (`RSAT: Active Directory` feature).
- Run under an account with read access to AD and permission to execute `server-command.exe`.

**Install**
1. Copy `manager-email-sync.ps1` to `C:\PaperCutScripts\` (or similar).
2. Edit the `--- Config ---` block at the top if needed:
   - `$ServerCommandPath` — path to `server-command.exe`.
   - `$SearchBase` — defaults to the current domain DN; override to scope to an OU.
   - `$LogDir` / `$LogRetentionDays` — logging options.
3. Test-run from an elevated PowerShell:
   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File C:\PaperCutScripts\manager-email-sync.ps1
   ```
4. Verify keys appear in PaperCut admin UI → **Options → Actions → Config editor (Advanced)** → search `script.user-defined.user-custom-property.manager-email`.
5. Schedule via **Task Scheduler** — e.g. daily at 02:00.

### 2. Print script

1. In PaperCut admin UI, go to **Printers → \<your printer\> → Scripting**.
2. Enable scripting.
3. Paste the contents of `papercut MF print script.txt` into the editor.
4. Edit the `--- Config ---` block at the top if needed:
   - `PAGE_THRESHOLD` — strictly greater-than; default `100` (101+ pages triggers alert).
   - `FALLBACK_EMAIL` — used when no manager email is known for the user.
5. Click **Apply**.

### 3. Verify

- Pick a user whose manager email is populated (check the Config editor).
- Submit a print job larger than the threshold.
- Expect a PaperCut App Log entry:
  ```
  Sent over-threshold alert for <user> (<pages> pages) to manager <email>
  ```
- If you see a fallback entry instead, the user has no config key written — re-run the PS script and check the log for `SKIP`/`FAIL` lines.

## Verification commands

### List AD users with their manager and manager's email

```powershell
Get-ADUser -Filter { Enabled -eq $true } -Properties manager, mail |
    Where-Object { $_.manager } |
    ForEach-Object {
        $mgr = Get-ADUser -Identity $_.manager -Properties mail
        [PSCustomObject]@{
            User         = $_.sAMAccountName
            UserEmail    = $_.mail
            Manager      = $mgr.sAMAccountName
            ManagerEmail = $mgr.mail
        }
    } |
    Sort-Object User |
    Format-Table -AutoSize
```

### Diff AD manager email vs PaperCut config (confirm sync is current)

Any `InSync = False` row = AD manager changed after the last sync run, or the write failed.

```powershell
$ServerCommandPath = "C:\Program Files\PaperCut MF\server\bin\win\server-command.exe"

Get-ADUser -Filter { Enabled -eq $true } -Properties manager, mail |
    Where-Object { $_.manager } |
    ForEach-Object {
        $mgr = Get-ADUser -Identity $_.manager -Properties mail
        $configKey = "script.user-defined.user-custom-property.manager-email.$($_.sAMAccountName)"
        $pcValue = & $ServerCommandPath get-config $configKey 2>$null
        [PSCustomObject]@{
            User            = $_.sAMAccountName
            AD_Manager      = $mgr.sAMAccountName
            AD_ManagerEmail = $mgr.mail
            PC_ManagerEmail = $pcValue
            InSync          = ($mgr.mail -eq $pcValue)
        }
    } |
    Sort-Object User |
    Format-Table -AutoSize
```

> **Note:** this uses `get-config` against the visible config key created by `manager-email-sync.ps1`. The hidden-variant path (`get-user-property <user> print-script-property.manager-email`) will NOT return anything for this deployment.

## Troubleshooting

| Symptom | Likely cause |
|---------|--------------|
| `No manager-email set for user X — falling back to …` | No config key for user. Either user wasn't in scope of the PS run, their AD `manager` is empty, or the manager has no `mail` attribute. Check the PS log. |
| Script throws `Cannot find function getConfigProperty …` | Print script is using the wrong API. Correct call is `inputs.utils.getProperty(...)` — not `actions.utils.getConfigProperty`. |
| `server-command not found` in PS log | `$ServerCommandPath` wrong, or PaperCut not installed on this host. |
| Keys never appear in Config editor | PS is running on a host without `server-command` or under an account that lacks permission. Must run on the app server. |
| PaperCut still runs old script behaviour after edits | Script is stored server-side. Paste the updated file into the Scripting editor and click **Apply**. Editing the local file does not push changes. |

## Notes

- The PS script writes one config key per user. Large user bases will make the Config editor busy — search rather than browse.
- Config key values have a 1000-character limit — plenty for an email address.
- The threshold check uses `<=` so `PAGE_THRESHOLD = 100` means 100 pages does NOT alert, 101 does.

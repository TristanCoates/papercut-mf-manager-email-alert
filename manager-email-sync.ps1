 # =============================================================================
# sync-manager-email-visible.ps1
# Syncs each AD user's manager's email into a PaperCut MF global config key
# that is visible and searchable in Options > Actions > Config editor (Advanced).
# Run as a scheduled task on the PaperCut app server.
#
# Config keys created: script.user-defined.user-custom-property.manager-email.<username>
# Print script reads via: inputs.utils.getProperty("user-custom-property.manager-email." + inputs.user.username)
# =============================================================================

Import-Module ActiveDirectory

# --- Config ---
$ServerCommandPath = "C:\Program Files\PaperCut MF\server\bin\win\server-command.exe"
$PropertyName      = "user-custom-property.manager-email"
$SearchBase        = (Get-ADDomain).DistinguishedName   # auto-detect domain
$LogDir            = "C:\PaperCutScripts\Logs"
$LogRetentionDays  = 30
# --------------

if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}
$LogFile = Join-Path $LogDir ("sync-manager-email-" + (Get-Date -Format "yyyy-MM-dd") + ".log")

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp  [$Level]  $Message" | Tee-Object -FilePath $LogFile -Append
}

Get-ChildItem -Path $LogDir -Filter "sync-manager-email-*.log" -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogRetentionDays) } |
    Remove-Item -Force -ErrorAction SilentlyContinue

Write-Log "=== Starting manager-email sync (visible config variant) ==="
Write-Log "Using SearchBase: $SearchBase"

if (-not (Test-Path $ServerCommandPath)) {
    Write-Log "FATAL: server-command not found at $ServerCommandPath" "ERROR"
    exit 1
}

try {
    $users = Get-ADUser -Filter { Enabled -eq $true } `
                        -SearchBase $SearchBase `
                        -Properties sAMAccountName, manager |
             Where-Object { $_.manager }
    Write-Log "Found $($users.Count) enabled users with a manager attribute."
}
catch {
    Write-Log "FATAL: Could not query AD - $($_.Exception.Message)" "ERROR"
    exit 1
}

if ($users.Count -eq 0) {
    Write-Log "No eligible users found. Exiting." "WARN"
    exit 0
}

$updated = 0; $skipped = 0; $failed = 0
$managerCache = @{}

foreach ($user in $users) {
    try {
        $managerDN = $user.manager

        if (-not $managerCache.ContainsKey($managerDN)) {
            try {
                $managerCache[$managerDN] = Get-ADUser -Identity $managerDN -Properties mail
            }
            catch {
                Write-Log "SKIP: $($user.sAMAccountName) - manager DN could not be resolved ($managerDN)" "WARN"
                $managerCache[$managerDN] = $null
            }
        }

        $manager = $managerCache[$managerDN]

        if ($null -eq $manager) { $skipped++; continue }

        $managerEmail = $manager.mail

        if ([string]::IsNullOrWhiteSpace($managerEmail)) {
            Write-Log "SKIP: $($user.sAMAccountName) - manager $($manager.sAMAccountName) has no mail attribute." "WARN"
            $skipped++; continue
        }

        # Build the fully-qualified config key and write via set-config.
        $configKey = "script.user-defined.$PropertyName.$($user.sAMAccountName)"

        & $ServerCommandPath set-config $configKey $managerEmail | Out-Null

        if ($LASTEXITCODE -eq 0) {
            Write-Log "OK: $($user.sAMAccountName) -> $managerEmail"
            $updated++
        }
        else {
            Write-Log "FAIL: $($user.sAMAccountName) - server-command exit code $LASTEXITCODE" "ERROR"
            $failed++
        }
    }
    catch {
        Write-Log "ERROR: $($user.sAMAccountName) - $($_.Exception.Message)" "ERROR"
        $failed++
    }
}

Write-Log "=== Done. Updated: $updated, Skipped: $skipped, Failed: $failed ==="

if ($failed -gt 0) { exit 2 }
exit 0 

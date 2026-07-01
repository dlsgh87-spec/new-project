[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$TaskName = "OSP Inventory Weekday 0915 Sync"
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
$SyncScript = Join-Path $ProjectRoot "src\sync_inventory.py"

if (-not (Test-Path -LiteralPath $Python)) {
    throw "Python virtual environment was not found: $Python"
}

if (-not (Test-Path -LiteralPath $SyncScript)) {
    throw "Sync script was not found: $SyncScript"
}

$Action = New-ScheduledTaskAction `
    -Execute $Python `
    -Argument "`"$SyncScript`"" `
    -WorkingDirectory $ProjectRoot

$Now = Get-Date
$StartAt = Get-Date -Hour 9 -Minute 15 -Second 0

$Trigger = New-ScheduledTaskTrigger `
    -Weekly `
    -DaysOfWeek Monday, Tuesday, Wednesday, Thursday, Friday `
    -At $StartAt

$CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$Principal = New-ScheduledTaskPrincipal `
    -UserId $CurrentUser `
    -LogonType Interactive `
    -RunLevel Limited

$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -MultipleInstances IgnoreNew

$LegacyTaskNames = @("OSP Inventory Hourly Sync", "OSP Inventory Daily Noon Sync")
foreach ($LegacyTaskName in $LegacyTaskNames) {
    if ($TaskName -ne $LegacyTaskName -and (Get-ScheduledTask -TaskName $LegacyTaskName -ErrorAction SilentlyContinue)) {
        Unregister-ScheduledTask -TaskName $LegacyTaskName -Confirm:$false
        Write-Host "Legacy task removed: $LegacyTaskName"
    }
}

if ($PSCmdlet.ShouldProcess($TaskName, "Register weekday 09:15 inventory sync task")) {
    Register-ScheduledTask `
        -TaskName $TaskName `
        -Action $Action `
        -Trigger $Trigger `
        -Principal $Principal `
        -Settings $Settings `
        -Description "Runs OSP inventory sync on weekdays at 09:15." `
        -Force | Out-Null

    Write-Host "Scheduled task installed: $TaskName"
    Write-Host "Start time: $($StartAt.ToString('HH:mm:ss'))"
    Write-Host "Repeat: weekdays at 09:15"
}

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$TaskName = "OSP Inventory Weekday 0915 Sync"
)

$ErrorActionPreference = "Stop"

$TaskNames = @($TaskName, "OSP Inventory Daily Noon Sync", "OSP Inventory Hourly Sync") | Select-Object -Unique
$Removed = $false

foreach ($Name in $TaskNames) {
    if (-not (Get-ScheduledTask -TaskName $Name -ErrorAction SilentlyContinue)) {
        continue
    }

    if ($PSCmdlet.ShouldProcess($Name, "Unregister inventory sync task")) {
        Unregister-ScheduledTask -TaskName $Name -Confirm:$false
        Write-Host "Scheduled task removed: $Name"
        $Removed = $true
    }
}

if (-not $Removed) {
    Write-Host "Scheduled task was not found: $TaskName"
}

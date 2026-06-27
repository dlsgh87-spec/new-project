$ErrorActionPreference = "Stop"

$RepoPath = "C:\Users\Administrator\Documents\New project"
$RemoteName = "origin"
$RemoteUrl = "https://github.com/dlsgh87-spec/new-project.git"
$Branch = "main"
$CommitMessage = "Auto sync from CHOIIH"
$LogDir = Join-Path $env:LOCALAPPDATA "CodexAutoSync\new-project"
$LogPath = Join-Path $LogDir "sync.log"
$LockPath = Join-Path $LogDir "sync.lock"

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

function Write-Log {
    param([string] $Message)
    $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogPath -Value "[$stamp] $Message"
}

function Invoke-Git {
    param(
        [Parameter(Mandatory = $true)]
        [string[]] $Arguments,
        [switch] $AllowFailure
    )

    Write-Log "git $($Arguments -join ' ')"
    $previousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        $output = & git @Arguments 2>&1
        $code = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }

    foreach ($line in $output) {
        Write-Log "  $line"
    }

    if ($code -ne 0 -and -not $AllowFailure) {
        throw "Git command failed: git $($Arguments -join ' ')"
    }

    return [pscustomobject]@{
        Code = $code
        Output = $output
    }
}

function Test-GitPath {
    param([string] $RelativePath)
    $gitDir = (Invoke-Git @("rev-parse", "--git-dir")).Output | Select-Object -First 1
    return Test-Path (Join-Path $gitDir $RelativePath)
}

function Stop-IfConflictOrOperationInProgress {
    if (Test-GitPath "MERGE_HEAD") {
        throw "Merge is already in progress. Stopping without overwriting anything."
    }

    if ((Test-GitPath "rebase-merge") -or (Test-GitPath "rebase-apply")) {
        throw "Rebase is already in progress. Stopping without overwriting anything."
    }

    $unmergedResult = Invoke-Git @("diff", "--name-only", "--diff-filter=U") -AllowFailure
    $unmerged = @(
        $unmergedResult.Output |
            Where-Object { $_ -and ($_ -notmatch "^warning: ") }
    )
    if ($unmerged.Count -gt 0) {
        throw "Unmerged files exist. Stopping without overwriting anything: $($unmerged -join ', ')"
    }
}

function Repair-BrokenOrigHead {
    $gitDir = (Invoke-Git @("rev-parse", "--git-dir")).Output | Select-Object -First 1
    $origHeadPath = Join-Path $gitDir "ORIG_HEAD"

    if (-not (Test-Path $origHeadPath)) {
        return
    }

    $verify = Invoke-Git @("rev-parse", "--verify", "--quiet", "ORIG_HEAD") -AllowFailure
    if ($verify.Code -ne 0) {
        Write-Log "Removing broken ORIG_HEAD metadata."
        Remove-Item -LiteralPath $origHeadPath -Force
    }
}

function Test-WorkingTreeHasChanges {
    $status = @(
        (Invoke-Git @("status", "--porcelain")).Output |
            Where-Object { $_ -and ($_ -notmatch "^warning: ") }
    )
    return $status.Count -gt 0
}

try {
    if (Test-Path $LockPath) {
        $lockAge = (Get-Date) - (Get-Item $LockPath).LastWriteTime
        if ($lockAge.TotalMinutes -lt 30) {
            Write-Log "Another sync appears to be running. Exiting."
            exit 0
        }

        Write-Log "Removing stale lock file."
        Remove-Item -LiteralPath $LockPath -Force
    }

    New-Item -ItemType File -Path $LockPath -Force | Out-Null
    Write-Log "==== Sync started ===="

    Set-Location -LiteralPath $RepoPath

    $inside = (Invoke-Git @("rev-parse", "--is-inside-work-tree")).Output | Select-Object -First 1
    if ($inside -ne "true") {
        throw "$RepoPath is not a Git repository."
    }

    $currentRemoteUrl = (Invoke-Git @("remote", "get-url", $RemoteName) -AllowFailure).Output | Select-Object -First 1
    if ($LASTEXITCODE -ne 0 -or $currentRemoteUrl -ne $RemoteUrl) {
        Invoke-Git @("remote", "remove", $RemoteName) -AllowFailure | Out-Null
        Invoke-Git @("remote", "add", $RemoteName, $RemoteUrl) | Out-Null
    }

    $currentBranch = (Invoke-Git @("branch", "--show-current")).Output | Select-Object -First 1
    if ($currentBranch -ne $Branch) {
        if (Test-WorkingTreeHasChanges) {
            throw "Current branch is $currentBranch, not $Branch, and local changes exist. Stopping."
        }

        Invoke-Git @("checkout", $Branch) | Out-Null
    }

    Invoke-Git @("fetch", $RemoteName, $Branch) | Out-Null
    Invoke-Git @("branch", "--set-upstream-to=$RemoteName/$Branch", $Branch) -AllowFailure | Out-Null

    Stop-IfConflictOrOperationInProgress

    Repair-BrokenOrigHead

    if (Test-WorkingTreeHasChanges) {
        Write-Log "Local changes found. Committing before sync."
        Invoke-Git @("add", "-A") | Out-Null

        $staged = (Invoke-Git @("diff", "--cached", "--name-only")).Output
        if ($staged.Count -gt 0) {
            Invoke-Git @("commit", "-m", $CommitMessage) | Out-Null
        }
    }
    else {
        Write-Log "No local file changes found."
    }

    Stop-IfConflictOrOperationInProgress

    Repair-BrokenOrigHead

    $pull = Invoke-Git @("pull", "--rebase", $RemoteName, $Branch) -AllowFailure
    if ($pull.Code -ne 0) {
        Write-Log "Pull/rebase failed. This may be a conflict. Stopping without push."
        exit 2
    }

    Stop-IfConflictOrOperationInProgress

    $push = Invoke-Git @("push", $RemoteName, $Branch) -AllowFailure
    if ($push.Code -ne 0) {
        Write-Log "Push failed. Stopping."
        exit 3
    }

    Write-Log "Sync completed."
}
catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    exit 1
}
finally {
    if (Test-Path $LockPath) {
        Remove-Item -LiteralPath $LockPath -Force -ErrorAction SilentlyContinue
    }
    Write-Log "==== Sync ended ===="
}

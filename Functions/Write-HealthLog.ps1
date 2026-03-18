function Write-HealthLog {
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Config,
        [PSCustomObject]$HealthMetrics,
        [PSCustomObject[]]$ServiceResults,
        [PSCustomObject[]]$EventResults,
        [PSCustomObject]$OverallStatus
    )
    $logLines = @()
    $logLines += "========== System Health Monitor Log =========="
    $logLines += "Run Time: $($Config.RunTime)"
    $logLines += "Computer Name: $($Config.ComputerName)"
    $logLines += "Overall Status: $($OverallStatus.OverallHealth)"
    $logLines += ""

    $logLines += "Status Reasons:"
    foreach ($reason in $OverallStatus.Reasons) {
        $logLines += " - $reason"
    }
    $logLines += ""

    $logLines += "System Health Metrics:"
    $logLines += "CPU Usage: $($HealthMetrics.CpuUsage)% [$($HealthMetrics.CpuStatus)]"
    $logLines += "Memory Usage: $($HealthMetrics.MemoryUsage)% [$($HealthMetrics.MemoryStatus)]"
    $logLines += ""

    $logLines += "Disk Usage:"
    foreach ($disk in $HealthMetrics.DiskResults) {
        $logLines += " - Drive $($disk.DriveLetter): $($disk.UsedGB)GB used of $($disk.TotalGB)GB ($($disk.FreePercent)% free) `
        | [$($disk.Status)]"
    }
    $logLines += ""

    $logLines += "Service Results:"
    foreach ($service in $ServiceResults) {
        $logLines += "Service: $($service.ServiceName) | Original Status: $($service.OriginalStatus) | Current Status: $($service.Status) `
        | Remediation Needed: $($service.RemediationNeeded) | Remediation Attempted: $($service.RemediationAttempted) `
        | Remediation Success: $($service.RemediationSuccess) | Notes: $($service.Notes)"
    }
    $logLines += ""

    $logLines += "Recent Event Errors:"
    if ($EventResults.Count -ge 0) {
              foreach ($event in $EventResults) {
            $logLines += "Log: $($event.LogName) | Time: $($event.Time) | ID: $($event.EventId) | Provider: $($event.ProviderName) `
            | Message: $($event.Message)"
        }
    } 
    else { 
        $logLines += "No recent event errors found."
    }
    $logLines += ""
    $logLines += "=============================================="

    $logLines | Out-File -Append "$($Config.LogFile)\SystemHealthLog.txt"
}
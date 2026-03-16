<# 
.SYNOPSIS
    PowerShell System Health Monitor and Auto-Remediation Tool

.DESCRIPTION
    Checks local system health metrics, verifies monitored services,
    attempts basic remediation, reviews recent event log errors,
    writes a log file, and generates an HTML report.

.NOTES
    Author: Josh Frometa
    Version: 1.0
    For use in Windows PowerShell 5.1 and later (including PowerShell Core on Windows)
#>

# Testing Code
# Use $PWD instead of $PSScriptRoot for testing, but switch back to $PSScriptRoot for production use

# Testing Initialize-HealthMonitor function
Get-ChildItem "$PWD\Functions\*.ps1" | ForEach-Object { . $_.FullName }
# Get-Command Initialize-HealthMonitor
# $config = Initialize-HealthMonitor
# $config | Format-List

#Testing Get-HealthMetrics function
$config = Initialize-HealthMonitor
$healthMetrics = Get-HealthMetrics -Config $config
$healthMetrics | Format-List
$healthMetrics.DiskResults | Format-Table -AutoSize

# Main Execution Block
try {
    $config = Initialize-HealthMonitor
    $healthMetrics = Get-HealthMetrics -Config $config
    $serviceResults = Get-ServiceHealth -ServiceNames $config.MonitoredServices
    $serviceResults = Invoke-ServiceRemediation -ServiceResults $serviceResults
}
catch {
    Write-Error "System Health monitor failed: $_"
    exit 1
}
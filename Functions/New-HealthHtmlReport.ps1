function New-HealthHtmlReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Config,

        [Parameter(Mandatory = $true)]
        [PSCustomObject]$HealthMetrics,

        [Parameter(Mandatory = $true)]
        [PSCustomObject[]]$ServiceResults,

        [PSCustomObject[]]$EventResults,

        [Parameter(Mandatory = $true)]
        [PSCustomObject]$OverallStatus
    )

    function Get-StatusColor {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Status
        )

        switch ($Status) {
            "Healthy"  { "#198754" } # green
            "Warning"  { "#ffc107" } # yellow
            "Critical" { "#dc3545" } # red
            default    { "#6c757d" } # gray
        }
    }

    function ConvertTo-HtmlEncoded {
        param(
            [AllowNull()]
            [string]$Text
        )

        if ([string]::IsNullOrWhiteSpace($Text)) {
            return ""
        }

        return [System.Net.WebUtility]::HtmlEncode($Text)
    }

    $overallColor = Get-StatusColor -Status $OverallStatus.Status

    $reasonItems = foreach ($reason in $OverallStatus.Reasons) {
        "<li>$(ConvertTo-HtmlEncoded $reason)</li>"
    }

    $diskRows = foreach ($disk in $HealthMetrics.DiskResults) {
        $diskColor = Get-StatusColor -Status $disk.Status
        @"
<tr>
    <td>$(ConvertTo-HtmlEncoded $disk.DriveLetter)</td>
    <td>$(ConvertTo-HtmlEncoded $disk.VolumeName)</td>
    <td>$($disk.SizeGB)</td>
    <td>$($disk.FreeGB)</td>
    <td>$($disk.PercentFree)%</td>
    <td>$($disk.PercentUsed)%</td>
    <td><span class='status-badge' style='background-color: $diskColor;'>$(ConvertTo-HtmlEncoded $disk.Status)</span></td>
</tr>
"@
    }

    $serviceRows = foreach ($service in $ServiceResults) {
        $serviceState = if ($service.CurrentStatus -eq "Running") {
            "Healthy"
        }
        elseif ($service.CurrentStatus -eq "Not Found") {
            "Warning"
        }
        elseif ($service.RemediationAttempted -and -not $service.RemediationSucceeded) {
            "Critical"
        }
        elseif ($service.NeedsRemediation) {
            "Warning"
        }
        else {
            "Healthy"
        }

        $serviceColor = Get-StatusColor -Status $serviceState

        @"
<tr>
    <td>$(ConvertTo-HtmlEncoded $service.ServiceName)</td>
    <td>$(ConvertTo-HtmlEncoded $service.DisplayName)</td>
    <td>$(ConvertTo-HtmlEncoded $service.OriginalStatus)</td>
    <td>$(ConvertTo-HtmlEncoded $service.CurrentStatus)</td>
    <td>$($service.NeedsRemediation)</td>
    <td>$($service.RemediationAttempted)</td>
    <td>$($service.RemediationSucceeded)</td>
    <td>$(ConvertTo-HtmlEncoded $service.Notes)</td>
    <td><span class='status-badge' style='background-color: $serviceColor;'>$serviceState</span></td>
</tr>
"@
    }

    $eventRows = if ($EventResults.Count -gt 0) {
        foreach ($event in $EventResults) {
            @"
<tr>
    <td>$(ConvertTo-HtmlEncoded $event.LogName)</td>
    <td>$(ConvertTo-HtmlEncoded ($event.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")))</td>
    <td>$(ConvertTo-HtmlEncoded $event.Id)</td>
    <td>$(ConvertTo-HtmlEncoded $event.ProviderName)</td>
    <td><pre>$(ConvertTo-HtmlEncoded $event.Message)</pre></td>
</tr>
"@
        }
    }
    else {
        @"
<tr>
    <td colspan='5'>No recent event errors found.</td>
</tr>
"@
    }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>System Health Report - $(ConvertTo-HtmlEncoded $Config.ComputerName)</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
            background-color: #f4f6f8;
            color: #212529;
        }

        h1, h2 {
            margin-bottom: 10px;
        }

        .section {
            background-color: #ffffff;
            padding: 16px;
            margin-bottom: 20px;
            border-radius: 8px;
            box-shadow: 0 1px 4px rgba(0,0,0,0.08);
        }

        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 12px;
        }

        .summary-card {
            background-color: #f8f9fa;
            border-left: 6px solid #dee2e6;
            padding: 12px;
            border-radius: 6px;
        }

        .status-badge {
            color: white;
            padding: 4px 10px;
            border-radius: 999px;
            font-weight: bold;
            display: inline-block;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 12px;
        }

        th, td {
            border: 1px solid #dee2e6;
            padding: 10px;
            text-align: left;
            vertical-align: top;
        }

        th {
            background-color: #e9ecef;
        }

        pre {
            margin: 0;
            white-space: pre-wrap;
            word-wrap: break-word;
            font-family: Arial, sans-serif;
        }

        ul {
            margin-top: 8px;
        }
    </style>
</head>
<body>
    <h1>System Health Report</h1>

    <div class="section">
        <h2>Run Summary</h2>
        <div class="summary-grid">
            <div class="summary-card">
                <strong>Computer Name:</strong><br>
                $(ConvertTo-HtmlEncoded $Config.ComputerName)
            </div>
            <div class="summary-card">
                <strong>Run Time:</strong><br>
                $(ConvertTo-HtmlEncoded ($Config.RunTime.ToString("yyyy-MM-dd HH:mm:ss")))
            </div>
            <div class="summary-card">
                <strong>Overall Status:</strong><br>
                <span class="status-badge" style="background-color: $overallColor;">$(ConvertTo-HtmlEncoded $OverallStatus.Status)</span>
            </div>
        </div>
    </div>

    <div class="section">
        <h2>Status Reasons</h2>
        <ul>
            $($reasonItems -join "`n")
        </ul>
    </div>

    <div class="section">
        <h2>System Health Metrics</h2>
        <div class="summary-grid">
            <div class="summary-card">
                <strong>CPU Usage:</strong><br>
                $($HealthMetrics.CpuPercent)%<br>
                <span class="status-badge" style="background-color: $(Get-StatusColor -Status $HealthMetrics.CpuStatus);">
                    $(ConvertTo-HtmlEncoded $HealthMetrics.CpuStatus)
                </span>
            </div>
            <div class="summary-card">
                <strong>Memory Usage:</strong><br>
                $($HealthMetrics.MemoryPercent)%<br>
                <span class="status-badge" style="background-color: $(Get-StatusColor -Status $HealthMetrics.MemoryStatus);">
                    $(ConvertTo-HtmlEncoded $HealthMetrics.MemoryStatus)
                </span>
            </div>
        </div>
    </div>

    <div class="section">
        <h2>Disk Usage</h2>
        <table>
            <thead>
                <tr>
                    <th>Drive</th>
                    <th>Volume Name</th>
                    <th>Size (GB)</th>
                    <th>Free (GB)</th>
                    <th>% Free</th>
                    <th>% Used</th>
                    <th>Status</th>
                </tr>
            </thead>
            <tbody>
                $($diskRows -join "`n")
            </tbody>
        </table>
    </div>

    <div class="section">
        <h2>Service Results</h2>
        <table>
            <thead>
                <tr>
                    <th>Service Name</th>
                    <th>Display Name</th>
                    <th>Original Status</th>
                    <th>Current Status</th>
                    <th>Needs Remediation</th>
                    <th>Remediation Attempted</th>
                    <th>Remediation Succeeded</th>
                    <th>Notes</th>
                    <th>Health</th>
                </tr>
            </thead>
            <tbody>
                $($serviceRows -join "`n")
            </tbody>
        </table>
    </div>

    <div class="section">
        <h2>Recent Event Errors</h2>
        <table>
            <thead>
                <tr>
                    <th>Log Name</th>
                    <th>Time Created</th>
                    <th>Event ID</th>
                    <th>Provider</th>
                    <th>Message</th>
                </tr>
            </thead>
            <tbody>
                $($eventRows -join "`n")
            </tbody>
        </table>
    </div>
</body>
</html>
"@

    $html | Out-File -FilePath $Config.HtmlReportFile -Encoding utf8
}
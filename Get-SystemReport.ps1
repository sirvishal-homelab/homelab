param (
    [string[]]$ComputerName,
    [switch]$Hotfix,
    [switch]$HotfixDetail,
    [switch]$OS,
    [switch]$BIOS,
    [switch]$Hardware,
    [switch]$Apps,
    [switch]$CompareApps,
    [switch]$ExportCsv,
    [switch]$ExportHtml
)

$report = @{}
$appMatrix = @{}
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

foreach ($computer in $ComputerName) {
    Write-Host "Processing $computer..."
    $row = [PSCustomObject]@{ Hostname = $computer }

    # HOTFIX
    if ($Hotfix) {
        try {
            $hotfixes = Get-HotFix -ComputerName $computer -ErrorAction Stop
            $row.HotfixCount = $hotfixes.Count

            if ($HotfixDetail) {
                $hotfixes | Select HotFixID, InstalledOn, Description |
                    Export-Csv "HotfixDetails_$computer.csv" -NoTypeInformation
            }
        } catch {
            Write-Warning "Hotfix query failed for $computer: $_"
            $row.HotfixCount = "Error"
        }
    }

    # OS INFO
    if ($OS) {
        try {
            $os = Get-CimInstance Win32_OperatingSystem -ComputerName $computer -ErrorAction Stop
            $row.OSVersion = $os.Caption
            $row.BuildNumber = $os.BuildNumber
            $row.InstallDate = $os.InstallDate
            $row.LastBoot = $os.LastBootUpTime
        } catch {
            Write-Warning "OS query failed for $computer: $_"
            $row.OSVersion = "Error"
        }
    }

    # BIOS INFO
    if ($BIOS) {
        try {
            $bios = Get-CimInstance Win32_BIOS -ComputerName $computer -ErrorAction Stop
            $row.BIOSVersion = ($bios.BIOSVersion -join ", ")
            $row.BIOSManufacturer = $bios.Manufacturer
            $row.BIOSReleaseDate = $bios.ReleaseDate
        } catch {
            Write-Warning "BIOS query failed for $computer: $_"
            $row.BIOSVersion = "Error"
        }
    }

    # HARDWARE INFO
    if ($Hardware) {
        try {
            $sys = Get-CimInstance Win32_ComputerSystem -ComputerName $computer -ErrorAction Stop
            $row.Manufacturer = $sys.Manufacturer
            $row.Model = $sys.Model
            $row.RAMGB = [math]::Round($sys.TotalPhysicalMemory / 1GB, 2)

            $cpu = Get-CimInstance Win32_Processor -ComputerName $computer -ErrorAction Stop
            $row.CPU = $cpu.Name
        } catch {
            Write-Warning "Hardware query failed for $computer: $_"
            $row.Manufacturer = "Error"
        }
    }

    # APPS INFO
    if ($Apps -or $CompareApps) {
        try {
            $paths = @(
                "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
                "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
            )

            $apps = foreach ($path in $paths) {
                Get-ItemProperty $path -ErrorAction SilentlyContinue | Where-Object {
                    $_.DisplayName -and $_.DisplayVersion
                } | Select-Object DisplayName, DisplayVersion
            }

            $row.AppCount = $apps.Count

            if ($Apps) {
                $apps | Export-Csv "InstalledApps_$computer.csv" -NoTypeInformation
            }

            if ($CompareApps) {
                $appMatrix[$computer] = @{}
                foreach ($app in $apps) {
                    $name = $app.DisplayName.Trim()
                    $version = $app.DisplayVersion
                    $appMatrix[$computer][$name] = $version
                }
            }

        } catch {
            Write-Warning "App query failed for $computer: $_"
            $row.AppCount = "Error"
        }
    }

    $report[$computer] = $row
}

# Output system report
$report.Values | Format-Table -AutoSize

# Export system report
if ($ExportCsv) {
    $report.Values | Export-Csv "SystemReport_$timestamp.csv" -NoTypeInformation
    Write-Host "System report exported to SystemReport_$timestamp.csv"
}

# Compare apps across servers
if ($CompareApps) {
    $allApps = $appMatrix.Values | ForEach-Object { $_.Keys } | Select-Object -Unique
    $comparison = @()

    foreach ($app in $allApps) {
        $row = [ordered]@{ AppName = $app }
        foreach ($server in $ComputerName) {
            $row[$server] = $appMatrix[$server][$app] ?? "Not Installed"
        }
        $comparison += [PSCustomObject]$row
    }

    $comparison | Format-Table -AutoSize

    if ($ExportCsv) {
        $comparison | Export-Csv "AppComparison_$timestamp.csv" -NoTypeInformation
        Write-Host "App comparison exported to AppComparison_$timestamp.csv"
    }
}

# Export HTML report
if ($ExportHtml) {
    $templatePath = ".\report-template.html"
    $template = if (Test-Path $templatePath) {
        Get-Content $templatePath -Raw
    } else {
@"
<html>
<head>
    <style>
        body { font-family: Segoe UI, sans-serif; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
    </style>
</head>
<body>
    <h2>System Inventory Report</h2>
    {{ReportTable}}
</body>
</html>
"@
    }

    $htmlTable = $report.Values | ConvertTo-Html -Fragment -PreContent "<h3>System Summary</h3>"
    $finalHtml = $template -replace "{{ReportTable}}", $htmlTable
    $finalHtml | Out-File "SystemReport_$timestamp.html" -Encoding UTF8
    Write-Host "HTML report saved to SystemReport_$timestamp.html"
}

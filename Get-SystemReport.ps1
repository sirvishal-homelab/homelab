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

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logPath = "SystemInventory_$timestamp.log"
$report = @{}
$appMatrix = @{}

function Write-Log {
    param ([string]$Message)
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logPath -Value "$TimeStamp - $Message"
}

# Load config.json
$configPath = ".\config.json"
if (Test-Path $configPath) {
    try {
        $config = Get-Content $configPath -Raw | ConvertFrom-Json
        Write-Log "✅ Loaded config.json"
    } catch {
        Write-Warning "Failed to parse config.json: $_"
        Write-Log "❌ Failed to parse config.json"
    }
} else {
    Write-Warning "config.json not found"
    Write-Log "⚠️ config.json not found"
}

# Apply config overrides
$outputDir = $config.Export.OutputDirectory ?? "."
$templatePath = $config.Template.Path ?? ".\report-template.html"
$enableCsv = $config.Export.EnableCsv
$enableHtml = $config.Export.EnableHtml
$compareApps = $CompareApps -or ($config.Apps.CompareAcrossServers -eq $true)

Write-Log "Using output directory: $outputDir"
Write-Log "Using template path: $templatePath"
Write-Log "CompareApps enabled: $compareApps"

Write-Log "🔍 Starting system inventory for $($ComputerName.Count) machines"

foreach ($computer in $ComputerName) {
    Write-Host "Processing $computer..."
    Write-Log "Processing $computer"
    $row = [PSCustomObject]@{ Hostname = $computer }

    # HOTFIX
    if ($Hotfix) {
        try {
            $hotfixes = Get-HotFix -ComputerName $computer -ErrorAction Stop
            $row.HotfixCount = $hotfixes.Count
            Write-Log "$computer: Hotfix count = $($hotfixes.Count)"

            if ($HotfixDetail) {
                $file = Join-Path $outputDir "HotfixDetails_$computer.csv"
                $hotfixes | Select HotFixID, InstalledOn, Description | Export-Csv $file -NoTypeInformation
                Write-Log "$computer: Hotfix details exported to $file"
            }
        } catch {
            Write-Warning "Hotfix query failed for $computer: $_"
            Write-Log "$computer: Hotfix query failed - $_"
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
            Write-Log "$computer: OS info collected"
        } catch {
            Write-Warning "OS query failed for $computer: $_"
            Write-Log "$computer: OS query failed - $_"
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
            Write-Log "$computer: BIOS info collected"
        } catch {
            Write-Warning "BIOS query failed for $computer: $_"
            Write-Log "$computer: BIOS query failed - $_"
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
            Write-Log "$computer: Hardware info collected"
        } catch {
            Write-Warning "Hardware query failed for $computer: $_"
            Write-Log "$computer: Hardware query failed - $_"
            $row.Manufacturer = "Error"
        }
    }

    # APPS INFO
    if ($Apps -or $compareApps) {
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
            Write-Log "$computer: App count = $($apps.Count)"

            if ($Apps -or ($config.Apps.EnableExport -eq $true)) {
                $file = Join-Path $outputDir "InstalledApps_$computer.csv"
                $apps | Export-Csv $file -NoTypeInformation
                Write-Log "$computer: App list exported to $file"
            }

            if ($compareApps) {
                $appMatrix[$computer] = @{}
                foreach ($app in $apps) {
                    $name = $app.DisplayName.Trim()
                    $version = $app.DisplayVersion
                    $appMatrix[$computer][$name] = $version
                }
                Write-Log "$computer: App matrix data collected"
            }

        } catch {
            Write-Warning "App query failed for $computer: $_"
            Write-Log "$computer: App query failed - $_"
            $row.AppCount = "Error"
        }
    }

    $report[$computer] = $row
}

# Output system report
$report.Values | Format-Table -AutoSize

# Export system report
if ($ExportCsv -or $enableCsv) {
    $file = Join-Path $outputDir "SystemReport_$timestamp.csv"
    $report.Values | Export-Csv $file -NoTypeInformation
    Write-Log "System report exported to $file"
    Write-Host "System report exported to $file"
}

# Compare apps across servers
if ($compareApps) {
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

    if ($ExportCsv -or $enableCsv) {
        $file = Join-Path $outputDir "AppComparison_$timestamp.csv"
        $comparison | Export-Csv $file -NoTypeInformation
        Write-Log "App comparison exported to $file"
        Write-Host "App comparison exported to $file"
    }
}

# Export HTML report
if ($ExportHtml -or $enableHtml) {
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
    $file = Join-Path $outputDir "SystemReport_$timestamp.html"
    $finalHtml | Out-File $file -Encoding UTF8

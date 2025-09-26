# Laptop Inspection Script - Generates HTML Report
# Created for refurbished laptop testing
# Run as Administrator for best results

$OutputPath = "$env:USERPROFILE\Desktop\Laptop_Inspection_Report.html"
$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

Write-Host "Starting Laptop Inspection..." -ForegroundColor Green
Write-Host "Report will be saved to: $OutputPath" -ForegroundColor Yellow

# Initialize HTML Report
$HTML = @"
<!DOCTYPE html>
<html>
<head>
    <title>Laptop Inspection Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .header { background-color: #2c3e50; color: white; padding: 20px; border-radius: 5px; margin-bottom: 20px; }
        .section { background-color: white; padding: 20px; margin-bottom: 20px; border-radius: 5px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
        .section h2 { color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px; }
        .pass { color: #27ae60; font-weight: bold; }
        .warning { color: #f39c12; font-weight: bold; }
        .fail { color: #e74c3c; font-weight: bold; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #3498db; color: white; }
        .cmd-output { background-color: #2c3e50; color: #ecf0f1; padding: 10px; border-radius: 3px; font-family: monospace; white-space: pre-wrap; overflow-x: auto; }
    </style>
</head>
<body>
    <div class="header">
        <h1>🔍 Laptop Inspection Report</h1>
        <p>Generated on: $Timestamp</p>
        <p>Inspection of: AMD Ryzen 5 3500U Refurbished Laptop</p>
    </div>
"@

Write-Host "1. Collecting System Information..." -ForegroundColor Cyan

# System Information
try {
    $SystemInfo = Get-ComputerInfo | Select-Object WindowsProductName, WindowsVersion, WindowsBuildLabEx, TotalPhysicalMemory, CsProcessors
    $CPU = Get-WmiObject -Class Win32_Processor | Select-Object Name, MaxClockSpeed, NumberOfCores, NumberOfLogicalProcessors
    $Memory = [math]::Round((Get-WmiObject -Class Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)

    $HTML += @"
    <div class="section">
        <h2>💻 System Information</h2>
        <table>
            <tr><th>Component</th><th>Details</th><th>Status</th></tr>
            <tr><td>Operating System</td><td>$($SystemInfo.WindowsProductName)</td><td class="pass">✓ PASS</td></tr>
            <tr><td>Windows Version</td><td>$($SystemInfo.WindowsVersion)</td><td class="pass">✓ PASS</td></tr>
            <tr><td>CPU</td><td>$($CPU.Name)</td><td class="pass">✓ PASS</td></tr>
            <tr><td>CPU Cores</td><td>$($CPU.NumberOfCores) cores, $($CPU.NumberOfLogicalProcessors) threads</td><td class="pass">✓ PASS</td></tr>
            <tr><td>Total RAM</td><td>$Memory GB</td><td class="$(if($Memory -ge 15){'pass'}else{'warning'})">$(if($Memory -ge 15){'✓ PASS'}else{'⚠ WARNING'})</td></tr>
        </table>
    </div>
"@
} catch {
    $HTML += "<div class='section'><h2>💻 System Information</h2><p class='fail'>❌ ERROR: Could not retrieve system information</p></div>"
}

Write-Host "2. Checking Storage Drives..." -ForegroundColor Cyan

# Storage Information
try {
    $Disks = Get-WmiObject -Class Win32_DiskDrive | Select-Object Model, Size, Status
    $LogicalDisks = Get-WmiObject -Class Win32_LogicalDisk | Where-Object {$_.DriveType -eq 3} | Select-Object DeviceID, Size, FreeSpace, FileSystem

    $HTML += "<div class='section'><h2>💾 Storage Information</h2><table><tr><th>Drive</th><th>Model</th><th>Size</th><th>Status</th></tr>"

    foreach ($disk in $Disks) {
        $SizeGB = [math]::Round($disk.Size / 1GB, 2)
        $StatusClass = if($disk.Status -eq "OK") {"pass"} else {"fail"}
        $StatusText = if($disk.Status -eq "OK") {"✓ PASS"} else {"❌ FAIL"}
        $HTML += "<tr><td>Physical Drive</td><td>$($disk.Model)</td><td>$SizeGB GB</td><td class='$StatusClass'>$StatusText</td></tr>"
    }
    $HTML += "</table>"

    $HTML += "<h3>Logical Drives</h3><table><tr><th>Drive</th><th>Total Size</th><th>Free Space</th><th>File System</th></tr>"
    foreach ($drive in $LogicalDisks) {
        $TotalGB = [math]::Round($drive.Size / 1GB, 2)
        $FreeGB = [math]::Round($drive.FreeSpace / 1GB, 2)
        $HTML += "<tr><td>$($drive.DeviceID)</td><td>$TotalGB GB</td><td>$FreeGB GB</td><td>$($drive.FileSystem)</td></tr>"
    }
    $HTML += "</table></div>"

} catch {
    $HTML += "<div class='section'><h2>💾 Storage Information</h2><p class='fail'>❌ ERROR: Could not retrieve storage information</p></div>"
}

Write-Host "3. Generating Battery Report..." -ForegroundColor Cyan

# Battery Information
try {
    $BatteryPath = "$env:TEMP\battery-temp.html"
    $BatteryCmd = "powercfg /batteryreport /output `"$BatteryPath`""
    Invoke-Expression $BatteryCmd

    if (Test-Path $BatteryPath) {
        $BatteryContent = Get-Content $BatteryPath -Raw
        # Extract key battery info using regex
        if ($BatteryContent -match 'DESIGN CAPACITY</th><td>(\d+,?\d*) mWh') {
            $DesignCapacity = $matches[1]
        }
        if ($BatteryContent -match 'FULL CHARGE CAPACITY</th><td>(\d+,?\d*) mWh') {
            $CurrentCapacity = $matches[1]
        }

        $HTML += @"
        <div class="section">
            <h2>🔋 Battery Health</h2>
            <table>
                <tr><th>Parameter</th><th>Value</th><th>Status</th></tr>
                <tr><td>Design Capacity</td><td>$DesignCapacity mWh</td><td class="pass">✓ INFO</td></tr>
                <tr><td>Current Capacity</td><td>$CurrentCapacity mWh</td><td class="pass">✓ INFO</td></tr>
            </table>
            <p><strong>📊 Detailed Battery Report:</strong> <a href="file:///$BatteryPath">Open Full Battery Report</a></p>
        </div>
"@
    }
} catch {
    $HTML += "<div class='section'><h2>🔋 Battery Health</h2><p class='warning'>⚠ WARNING: Could not generate battery report</p></div>"
}

Write-Host "4. Checking Network Adapters..." -ForegroundColor Cyan

# Network Information
try {
    $NetworkAdapters = Get-NetAdapter | Where-Object {$_.Status -eq "Up" -or $_.Name -like "*Wi-Fi*" -or $_.Name -like "*Ethernet*"} | Select-Object Name, InterfaceDescription, LinkSpeed, Status

    $HTML += "<div class='section'><h2>🌐 Network Adapters</h2><table><tr><th>Adapter</th><th>Description</th><th>Speed</th><th>Status</th></tr>"

    foreach ($adapter in $NetworkAdapters) {
        $StatusClass = if($adapter.Status -eq "Up") {"pass"} else {"warning"}
        $StatusText = if($adapter.Status -eq "Up") {"✓ CONNECTED"} else {"⚠ DISCONNECTED"}
        $HTML += "<tr><td>$($adapter.Name)</td><td>$($adapter.InterfaceDescription)</td><td>$($adapter.LinkSpeed)</td><td class='$StatusClass'>$StatusText</td></tr>"
    }
    $HTML += "</table></div>"

} catch {
    $HTML += "<div class='section'><h2>🌐 Network Adapters</h2><p class='fail'>❌ ERROR: Could not retrieve network information</p></div>"
}

Write-Host "5. Running System File Check..." -ForegroundColor Cyan

# System File Check (Quick verification)
try {
    $SFCResult = & sfc /verifyonly 2>&1 | Out-String
    $SFCStatus = if($SFCResult -like "*did not find any integrity violations*") {"pass"} else {"warning"}
    $SFCText = if($SFCResult -like "*did not find any integrity violations*") {"✓ PASS - No Issues Found"} else {"⚠ WARNING - Issues Detected"}

    $HTML += @"
    <div class="section">
        <h2>🛠️ System File Integrity</h2>
        <table>
            <tr><th>Check</th><th>Result</th><th>Status</th></tr>
            <tr><td>System File Check</td><td>Verification Complete</td><td class="$SFCStatus">$SFCText</td></tr>
        </table>
        <div class="cmd-output">$SFCResult</div>
    </div>
"@
} catch {
    $HTML += "<div class='section'><h2>🛠️ System File Integrity</h2><p class='warning'>⚠ WARNING: Could not run system file check</p></div>"
}

Write-Host "6. Checking Hardware Devices..." -ForegroundColor Cyan

# Hardware Device Status
try {
    $ProblemDevices = Get-PnpDevice | Where-Object {$_.Status -ne "OK"}

    if ($ProblemDevices.Count -eq 0) {
        $HTML += @"
        <div class="section">
            <h2>🔧 Hardware Devices</h2>
            <p class="pass">✓ PASS - All hardware devices are functioning properly</p>
        </div>
"@
    } else {
        $HTML += "<div class='section'><h2>🔧 Hardware Devices</h2><table><tr><th>Device</th><th>Status</th><th>Problem Code</th></tr>"
        foreach ($device in $ProblemDevices) {
            $HTML += "<tr><td>$($device.FriendlyName)</td><td class='fail'>❌ $($device.Status)</td><td>$($device.Problem)</td></tr>"
        }
        $HTML += "</table></div>"
    }
} catch {
    $HTML += "<div class='section'><h2>🔧 Hardware Devices</h2><p class='fail'>❌ ERROR: Could not check hardware devices</p></div>"
}

Write-Host "7. Running Performance Assessment..." -ForegroundColor Cyan

# Basic Performance Check
try {
    $BootTime = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
    $Uptime = (Get-Date) - $BootTime
    $UptimeText = "$([int]$Uptime.TotalHours) hours, $([int]$Uptime.Minutes) minutes"

    # CPU Load
    $CPULoad = Get-WmiObject win32_processor | Measure-Object -property LoadPercentage -Average | Select-Object Average

    $HTML += @"
    <div class="section">
        <h2>📊 Performance Metrics</h2>
        <table>
            <tr><th>Metric</th><th>Value</th><th>Status</th></tr>
            <tr><td>Last Boot Time</td><td>$BootTime</td><td class="pass">✓ INFO</td></tr>
            <tr><td>System Uptime</td><td>$UptimeText</td><td class="pass">✓ INFO</td></tr>
            <tr><td>Current CPU Load</td><td>$($CPULoad.Average)%</td><td class="pass">✓ INFO</td></tr>
        </table>
    </div>
"@
} catch {
    $HTML += "<div class='section'><h2>📊 Performance Metrics</h2><p class='warning'>⚠ WARNING: Could not retrieve performance data</p></div>"
}

# Final Summary
$HTML += @"
    <div class="section">
        <h2>📋 Inspection Summary</h2>
        <h3>✅ What to Look For:</h3>
        <ul>
            <li><strong>RAM:</strong> Should show approximately 16GB</li>
            <li><strong>CPU:</strong> Should be AMD Ryzen 5 3500U with 4 cores/8 threads</li>
            <li><strong>Storage:</strong> Should show SSD (~250GB) and HDD (~500GB) both with OK status</li>
            <li><strong>Battery:</strong> Check detailed report - avoid if capacity <50% of design</li>
            <li><strong>Hardware:</strong> No devices should show error status</li>
            <li><strong>Network:</strong> WiFi and Ethernet adapters should be detected</li>
        </ul>

        <h3>🚩 Red Flags:</h3>
        <ul>
            <li>Any storage device showing status other than "OK"</li>
            <li>RAM showing less than 15GB total</li>
            <li>Multiple hardware devices with errors</li>
            <li>System file integrity violations</li>
            <li>Missing network adapters</li>
        </ul>

        <p><strong>💡 Tip:</strong> Save this report and compare with seller's claims. Run this test in front of the seller before purchase!</p>
    </div>

    <div class="section">
        <h2>🔧 Additional Manual Tests Required</h2>
        <ul>
            <li><strong>Physical:</strong> Test all USB ports, HDMI output, keyboard keys, trackpad</li>
            <li><strong>Screen:</strong> Check for dead pixels using solid color backgrounds</li>
            <li><strong>Battery:</strong> Test charging and unplug to verify battery life</li>
            <li><strong>Audio:</strong> Test speakers and headphone jack</li>
            <li><strong>Camera:</strong> Open Camera app to verify webcam functionality</li>
        </ul>
    </div>

    <footer style="text-align: center; margin-top: 30px; padding: 20px; background-color: #34495e; color: white; border-radius: 5px;">
        <p>Laptop Inspection Report - Generated automatically for refurbished laptop testing</p>
        <p>Always verify results manually and test all physical components before purchase</p>
    </footer>
</body>
</html>
"@

# Save HTML Report
try {
    $HTML | Out-File -FilePath $OutputPath -Encoding UTF8
    Write-Host "`n✅ INSPECTION COMPLETE!" -ForegroundColor Green
    Write-Host "📄 Report saved to: $OutputPath" -ForegroundColor Yellow
    Write-Host "📂 Opening report..." -ForegroundColor Cyan
    Start-Process $OutputPath
} catch {
    Write-Host "❌ ERROR: Could not save report to $OutputPath" -ForegroundColor Red
    Write-Host "📄 Report content:" -ForegroundColor Yellow
    Write-Host $HTML
}

Write-Host "`n🔍 MANUAL TESTS STILL REQUIRED:" -ForegroundColor Magenta
Write-Host "• Test all USB ports with a device"
Write-Host "• Check HDMI output with external monitor"
Write-Host "• Test every keyboard key and trackpad"
Write-Host "• Look for dead pixels on screen"
Write-Host "• Verify battery charging and life"
Write-Host "• Test audio output (speakers/headphones)"
Write-Host "`nPress any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

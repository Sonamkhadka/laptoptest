# Laptop Inspection Script - Fixed Version
# Run as Administrator for best results
# Compatible with strict IT environments

param(
    [string]$OutputPath = "$env:USERPROFILE\Desktop\Laptop_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
)

# Set execution policy for current session
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

Write-Host "=== LAPTOP INSPECTION TOOL ===" -ForegroundColor Green
Write-Host "Starting diagnostics..." -ForegroundColor Yellow
Write-Host "Report will save to: $OutputPath" -ForegroundColor Cyan

# Initialize report with error handling
$ReportData = @{
    SystemInfo = @{}
    StorageInfo = @()
    NetworkInfo = @()
    BatteryInfo = @{}
    Errors = @()
    Warnings = @()
    TestResults = @{}
}

# Function to add test result
function Add-TestResult {
    param($TestName, $Result, $Status, $Details = "")
    $ReportData.TestResults[$TestName] = @{
        Result = $Result
        Status = $Status
        Details = $Details
        Timestamp = Get-Date -Format "HH:mm:ss"
    }
    Write-Host "[$Status] $TestName`: $Result" -ForegroundColor $(if($Status -eq "PASS"){"Green"}elseif($Status -eq "WARN"){"Yellow"}else{"Red"})
}

# Test 1: Basic System Information
Write-Host "`nTest 1: System Information..." -ForegroundColor Cyan
try {
    $OS = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    $CS = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
    $CPU = Get-CimInstance Win32_Processor -ErrorAction Stop | Select-Object -First 1

    $ReportData.SystemInfo = @{
        OSName = $OS.Caption
        OSVersion = $OS.Version
        ComputerName = $CS.Name
        TotalRAM_GB = [math]::Round($CS.TotalPhysicalMemory / 1GB, 2)
        CPUName = $CPU.Name
        CPUCores = $CPU.NumberOfCores
        CPUThreads = $CPU.NumberOfLogicalProcessors
    }

    Add-TestResult "System Info" "OS: $($OS.Caption), RAM: $($ReportData.SystemInfo.TotalRAM_GB) GB" "PASS"

    # Check if RAM matches expected (should be ~16GB)
    if ($ReportData.SystemInfo.TotalRAM_GB -lt 15) {
        Add-TestResult "RAM Check" "$($ReportData.SystemInfo.TotalRAM_GB) GB (Expected ~16GB)" "WARN" "Lower than expected RAM"
    } else {
        Add-TestResult "RAM Check" "$($ReportData.SystemInfo.TotalRAM_GB) GB" "PASS"
    }
} catch {
    Add-TestResult "System Info" "Failed to retrieve" "FAIL" $_.Exception.Message
    $ReportData.Errors += "System Info: $($_.Exception.Message)"
}

# Test 2: CPU Information
Write-Host "`nTest 2: CPU Details..." -ForegroundColor Cyan
try {
    if ($ReportData.SystemInfo.CPUName -like "*Ryzen 5 3500U*") {
        Add-TestResult "CPU Model" $ReportData.SystemInfo.CPUName "PASS"
    } else {
        Add-TestResult "CPU Model" $ReportData.SystemInfo.CPUName "WARN" "Expected AMD Ryzen 5 3500U"
    }

    Add-TestResult "CPU Cores" "$($ReportData.SystemInfo.CPUCores) cores, $($ReportData.SystemInfo.CPUThreads) threads" "PASS"
} catch {
    Add-TestResult "CPU Check" "Failed" "FAIL" $_.Exception.Message
}

# Test 3: Storage Drives
Write-Host "`nTest 3: Storage Drives..." -ForegroundColor Cyan
try {
    $PhysicalDisks = Get-CimInstance Win32_DiskDrive -ErrorAction Stop
    $LogicalDisks = Get-CimInstance Win32_LogicalDisk -ErrorAction Stop | Where-Object {$_.DriveType -eq 3}

    foreach ($disk in $PhysicalDisks) {
        $SizeGB = [math]::Round($disk.Size / 1GB, 2)
        $DriveInfo = @{
            Model = $disk.Model
            Size_GB = $SizeGB
            Status = if($disk.Status -eq "OK") {"OK"} else {"ERROR"}
            MediaType = $disk.MediaType
        }
        $ReportData.StorageInfo += $DriveInfo

        $StatusResult = if($disk.Status -eq "OK") {"PASS"} else {"FAIL"}
        Add-TestResult "Drive: $($disk.Model)" "$SizeGB GB - Status: $($disk.Status)" $StatusResult
    }

    # Check for expected SSD + HDD configuration
    $TotalDrives = $PhysicalDisks.Count
    Add-TestResult "Drive Count" "$TotalDrives drives detected" $(if($TotalDrives -ge 2){"PASS"}else{"WARN"})

} catch {
    Add-TestResult "Storage Check" "Failed" "FAIL" $_.Exception.Message
    $ReportData.Errors += "Storage: $($_.Exception.Message)"
}

# Test 4: Network Adapters
Write-Host "`nTest 4: Network Adapters..." -ForegroundColor Cyan
try {
    $NetAdapters = Get-CimInstance Win32_NetworkAdapter -ErrorAction Stop | Where-Object {$_.NetConnectionStatus -ne $null -or $_.Name -like "*Wi-Fi*" -or $_.Name -like "*Ethernet*" -or $_.Name -like "*Wireless*"}

    $WiFiFound = $false
    $EthernetFound = $false

    foreach ($adapter in $NetAdapters) {
        $AdapterInfo = @{
            Name = $adapter.Name
            Description = $adapter.Description
            Status = if($adapter.NetConnectionStatus -eq 2) {"Connected"} elseif($adapter.NetConnectionStatus -eq 7) {"Media Disconnected"} else {"Other"}
        }
        $ReportData.NetworkInfo += $AdapterInfo

        if ($adapter.Name -like "*Wi-Fi*" -or $adapter.Name -like "*Wireless*") {
            $WiFiFound = $true
        }
        if ($adapter.Name -like "*Ethernet*") {
            $EthernetFound = $true
        }

        Add-TestResult "Network: $($adapter.Name)" $AdapterInfo.Status "PASS"
    }

    Add-TestResult "WiFi Adapter" $(if($WiFiFound){"Detected"}else{"Not Found"}) $(if($WiFiFound){"PASS"}else{"FAIL"})
    Add-TestResult "Ethernet Adapter" $(if($EthernetFound){"Detected"}else{"Not Found"}) $(if($EthernetFound){"PASS"}else{"FAIL"})

} catch {
    Add-TestResult "Network Check" "Failed" "FAIL" $_.Exception.Message
    $ReportData.Errors += "Network: $($_.Exception.Message)"
}

# Test 5: Battery Information (with better error handling)
Write-Host "`nTest 5: Battery Health..." -ForegroundColor Cyan
try {
    $Battery = Get-CimInstance Win32_Battery -ErrorAction Stop | Select-Object -First 1
    if ($Battery) {
        $ReportData.BatteryInfo = @{
            Name = $Battery.Name
            Status = $Battery.Status
            PowerManagementSupported = $Battery.PowerManagementSupported
            Chemistry = $Battery.Chemistry
        }
        Add-TestResult "Battery Detection" "Battery found: $($Battery.Name)" "PASS"

        # Try to generate battery report (may fail on some systems)
        try {
            $BatteryReportPath = "$env:TEMP\battery_report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
            $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
            $ProcessInfo.FileName = "powercfg.exe"
            $ProcessInfo.Arguments = "/batteryreport /output `"$BatteryReportPath`""
            $ProcessInfo.UseShellExecute = $false
            $ProcessInfo.CreateNoWindow = $true
            $Process = [System.Diagnostics.Process]::Start($ProcessInfo)
            $Process.WaitForExit(5000) # 5 second timeout

            if (Test-Path $BatteryReportPath) {
                Add-TestResult "Battery Report" "Generated successfully" "PASS" "Report: $BatteryReportPath"
                $ReportData.BatteryInfo.ReportPath = $BatteryReportPath
            } else {
                Add-TestResult "Battery Report" "Could not generate" "WARN"
            }
        } catch {
            Add-TestResult "Battery Report" "Failed to generate" "WARN" $_.Exception.Message
        }
    } else {
        Add-TestResult "Battery Detection" "No battery found" "WARN" "May be desktop or battery not detected"
    }
} catch {
    Add-TestResult "Battery Check" "Failed" "WARN" $_.Exception.Message
}

# Test 6: System File Integrity (Quick check)
Write-Host "`nTest 6: System Health..." -ForegroundColor Cyan
try {
    # Quick system health indicators
    $Services = Get-Service -ErrorAction Stop | Where-Object {$_.Status -eq "Running"}
    $ServiceCount = $Services.Count
    Add-TestResult "Windows Services" "$ServiceCount services running" "PASS"

    # Check for obvious system issues
    $EventLogErrors = Get-WinEvent -FilterHashtable @{LogName='System'; Level=1,2} -MaxEvents 5 -ErrorAction SilentlyContinue
    if ($EventLogErrors) {
        Add-TestResult "Recent System Errors" "$($EventLogErrors.Count) critical errors in system log" "WARN" "Check Event Viewer"
    } else {
        Add-TestResult "Recent System Errors" "No critical errors found" "PASS"
    }
} catch {
    Add-TestResult "System Health" "Could not check" "WARN" $_.Exception.Message
}

# Test 7: Hardware Devices
Write-Host "`nTest 7: Hardware Devices..." -ForegroundColor Cyan
try {
    $PnpDevices = Get-CimInstance Win32_PnPEntity -ErrorAction Stop | Where-Object {$_.Status -ne "OK" -and $_.Status -ne $null}

    if ($PnpDevices.Count -eq 0) {
        Add-TestResult "Hardware Devices" "All devices OK" "PASS"
    } else {
        $ProblemDevices = $PnpDevices | Select-Object Name, Status -First 5
        Add-TestResult "Hardware Devices" "$($PnpDevices.Count) devices with issues" "WARN" ($ProblemDevices | Out-String)
    }
} catch {
    Add-TestResult "Hardware Check" "Could not check devices" "WARN" $_.Exception.Message
}

# Generate HTML Report
Write-Host "`nGenerating HTML report..." -ForegroundColor Cyan

$CurrentTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$PassCount = ($ReportData.TestResults.Values | Where-Object {$_.Status -eq "PASS"}).Count
$WarnCount = ($ReportData.TestResults.Values | Where-Object {$_.Status -eq "WARN"}).Count
$FailCount = ($ReportData.TestResults.Values | Where-Object {$_.Status -eq "FAIL"}).Count
$TotalTests = $ReportData.TestResults.Count

$HTML = @"
<!DOCTYPE html>
<html>
<head>
    <title>Laptop Inspection Report</title>
    <meta charset="UTF-8">
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 20px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); }
        .container { max-width: 1200px; margin: 0 auto; background: white; border-radius: 10px; overflow: hidden; box-shadow: 0 10px 30px rgba(0,0,0,0.3); }
        .header { background: linear-gradient(135deg, #2c3e50 0%, #34495e 100%); color: white; padding: 30px; text-align: center; }
        .header h1 { margin: 0; font-size: 2.5em; }
        .summary { padding: 20px; background: #f8f9fa; display: flex; justify-content: space-around; }
        .summary-item { text-align: center; }
        .summary-number { font-size: 2em; font-weight: bold; }
        .pass { color: #27ae60; }
        .warn { color: #f39c12; }
        .fail { color: #e74c3c; }
        .section { margin: 20px; padding: 20px; background: white; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .section h2 { color: #2c3e50; margin-top: 0; padding-bottom: 10px; border-bottom: 3px solid #3498db; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background: #3498db; color: white; font-weight: 600; }
        tr:hover { background: #f5f5f5; }
        .status-badge { padding: 4px 12px; border-radius: 20px; font-weight: bold; font-size: 0.9em; }
        .pass-badge { background: #d4edda; color: #155724; }
        .warn-badge { background: #fff3cd; color: #856404; }
        .fail-badge { background: #f8d7da; color: #721c24; }
        .recommendation { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin: 15px 0; }
        .alert { background: #ffebee; border-left: 4px solid #f44336; padding: 15px; margin: 15px 0; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🔍 Laptop Inspection Report</h1>
            <p>AMD Ryzen 5 3500U Refurbished Laptop Analysis</p>
            <p>Generated: $CurrentTime</p>
        </div>

        <div class="summary">
            <div class="summary-item">
                <div class="summary-number pass">$PassCount</div>
                <div>Tests Passed</div>
            </div>
            <div class="summary-item">
                <div class="summary-number warn">$WarnCount</div>
                <div>Warnings</div>
            </div>
            <div class="summary-item">
                <div class="summary-number fail">$FailCount</div>
                <div>Failed Tests</div>
            </div>
            <div class="summary-item">
                <div class="summary-number">$TotalTests</div>
                <div>Total Tests</div>
            </div>
        </div>

        <div class="section">
            <h2>📊 Test Results Overview</h2>
            <table>
                <tr><th>Test Name</th><th>Result</th><th>Status</th><th>Time</th></tr>
"@

foreach ($test in $ReportData.TestResults.GetEnumerator()) {
    $BadgeClass = switch ($test.Value.Status) {
        "PASS" { "pass-badge" }
        "WARN" { "warn-badge" }  
        "FAIL" { "fail-badge" }
    }

    $HTML += "<tr><td>$($test.Key)</td><td>$($test.Value.Result)</td><td><span class='status-badge $BadgeClass'>$($test.Value.Status)</span></td><td>$($test.Value.Timestamp)</td></tr>"
}

$HTML += "</table></div>"

# System Information Section
if ($ReportData.SystemInfo.Count -gt 0) {
    $HTML += @"
    <div class="section">
        <h2>💻 System Information</h2>
        <table>
            <tr><th>Component</th><th>Details</th></tr>
            <tr><td>Operating System</td><td>$($ReportData.SystemInfo.OSName)</td></tr>
            <tr><td>OS Version</td><td>$($ReportData.SystemInfo.OSVersion)</td></tr>
            <tr><td>Computer Name</td><td>$($ReportData.SystemInfo.ComputerName)</td></tr>
            <tr><td>CPU</td><td>$($ReportData.SystemInfo.CPUName)</td></tr>
            <tr><td>CPU Cores/Threads</td><td>$($ReportData.SystemInfo.CPUCores) cores / $($ReportData.SystemInfo.CPUThreads) threads</td></tr>
            <tr><td>Total RAM</td><td>$($ReportData.SystemInfo.TotalRAM_GB) GB</td></tr>
        </table>
    </div>
"@
}

# Storage Information
if ($ReportData.StorageInfo.Count -gt 0) {
    $HTML += "<div class='section'><h2>💾 Storage Drives</h2><table><tr><th>Drive Model</th><th>Size (GB)</th><th>Status</th></tr>"
    foreach ($drive in $ReportData.StorageInfo) {
        $StatusClass = if($drive.Status -eq "OK") {"pass"} else {"fail"}
        $HTML += "<tr><td>$($drive.Model)</td><td>$($drive.Size_GB)</td><td class='$StatusClass'>$($drive.Status)</td></tr>"
    }
    $HTML += "</table></div>"
}

# Network Information
if ($ReportData.NetworkInfo.Count -gt 0) {
    $HTML += "<div class='section'><h2>🌐 Network Adapters</h2><table><tr><th>Adapter Name</th><th>Status</th></tr>"
    foreach ($adapter in $ReportData.NetworkInfo) {
        $HTML += "<tr><td>$($adapter.Name)</td><td>$($adapter.Status)</td></tr>"
    }
    $HTML += "</table></div>"
}

# Recommendations Section
$HTML += @"
<div class="section">
    <h2>💡 Purchase Recommendations</h2>

    <div class="recommendation">
        <h3>✅ Good Signs (Proceed with Purchase):</h3>
        <ul>
            <li>RAM shows 15.8-16GB total</li>
            <li>CPU detected as AMD Ryzen 5 3500U</li>
            <li>All storage drives show "OK" status</li>
            <li>Both WiFi and Ethernet adapters detected</li>
            <li>No critical hardware failures</li>
        </ul>
    </div>

    <div class="alert">
        <h3>🚩 Red Flags (Avoid Purchase):</h3>
        <ul>
            <li>RAM significantly less than 15GB</li>
            <li>Storage drives showing error status</li>
            <li>Missing WiFi or Ethernet adapters</li>
            <li>Multiple hardware device failures</li>
            <li>Wrong CPU model (not Ryzen 5 3500U)</li>
        </ul>
    </div>
</div>

<div class="section">
    <h2>🔧 Manual Tests Still Required</h2>
    <p><strong>This script cannot test everything. You must also verify:</strong></p>
    <ul>
        <li><strong>Screen:</strong> Look for dead pixels, bright spots, color accuracy</li>
        <li><strong>Keyboard:</strong> Test every key, including function keys</li>
        <li><strong>Trackpad:</strong> Test clicking, scrolling, multi-touch</li>
        <li><strong>Ports:</strong> Plug devices into all USB ports, test HDMI output</li>
        <li><strong>Audio:</strong> Test speakers and headphone jack</li>
        <li><strong>Camera:</strong> Open Camera app to test webcam</li>
        <li><strong>Battery:</strong> Test charging and unplug to verify battery life</li>
        <li><strong>Physical condition:</strong> Check for cracks, loose hinges, wear</li>
    </ul>
</div>

<div class="section" style="text-align: center; background: #2c3e50; color: white;">
    <h2>📋 Final Verdict</h2>
    <p style="font-size: 1.2em;">
        Based on automated tests: 
        <strong style="color: $(if($FailCount -eq 0 -and $WarnCount -le 2){'#2ecc71'}elseif($FailCount -gt 0){'#e74c3c'}else{'#f39c12'});">
            $(if($FailCount -eq 0 -and $WarnCount -le 2){'RECOMMENDED FOR PURCHASE'}elseif($FailCount -gt 0){'NOT RECOMMENDED - ISSUES FOUND'}else{'PROCEED WITH CAUTION'})
        </strong>
    </p>
    <p>Remember: Always test manually and negotiate based on any issues found!</p>
</div>

</div>
</body>
</html>
"@

# Save the report
try {
    $HTML | Out-File -FilePath $OutputPath -Encoding UTF8
    Write-Host "`n✅ SUCCESS: Report saved to $OutputPath" -ForegroundColor Green

    # Try to open the report
    try {
        Start-Process $OutputPath
        Write-Host "📄 Report opened in browser" -ForegroundColor Green
    } catch {
        Write-Host "📄 Report saved but could not auto-open. Navigate to: $OutputPath" -ForegroundColor Yellow
    }
} catch {
    Write-Host "❌ ERROR: Could not save report to $OutputPath" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}

# Final summary
Write-Host "`n" + "="*50
Write-Host "INSPECTION COMPLETE" -ForegroundColor Green
Write-Host "="*50
Write-Host "✅ Passed: $PassCount tests" -ForegroundColor Green
Write-Host "⚠️  Warnings: $WarnCount tests" -ForegroundColor Yellow  
Write-Host "❌ Failed: $FailCount tests" -ForegroundColor Red
Write-Host "`nOverall recommendation: $(if($FailCount -eq 0 -and $WarnCount -le 2){'GOOD TO BUY'}elseif($FailCount -gt 0){'DO NOT BUY'}else{'BUYER BEWARE'})" -ForegroundColor $(if($FailCount -eq 0 -and $WarnCount -le 2){'Green'}elseif($FailCount -gt 0){'Red'}else{'Yellow'})

Write-Host "`n🔍 Don't forget to test manually:" -ForegroundColor Cyan
Write-Host "• Screen, keyboard, trackpad, ports, audio, camera" -ForegroundColor White

Write-Host "`nPress Enter to exit..." -ForegroundColor Gray
Read-Host

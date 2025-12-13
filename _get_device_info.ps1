$Host.UI.RawUI.WindowTitle = "Asset Inventory Report"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Get current date and time
$DateTime = Get-Date
$DateStr = $DateTime.ToString("yyyy/MM/dd")
$TimeStr = $DateTime.ToString("HH:mm:ss.ff")
$FileDate = $DateTime.ToString("yyyy-MM-dd")

# Find SANDISK Flash Drive automatically
$OutputPath = $null
$Drives = Get-Volume | Where-Object { $_.FileSystemLabel -like "*SANDISK*" -and $_.DriveLetter }

if ($Drives) {
    $OutputPath = "$($Drives[0].DriveLetter):\"
    Write-Host "Found SANDISK drive at $OutputPath"
} else {
    $OutputPath = "$env:USERPROFILE\Desktop\"
    Write-Host "Warning: SANDISK drive not found. Saving to Desktop instead."
    Start-Sleep -Seconds 3
}

$FileName = "$env:COMPUTERNAME`_$FileDate.txt"
$FullPath = Join-Path $OutputPath $FileName

# Collect system information
$ComputerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
$ComputerSystemProduct = Get-CimInstance -ClassName Win32_ComputerSystemProduct
$OS = Get-CimInstance -ClassName Win32_OperatingSystem
$BIOS = Get-CimInstance -ClassName Win32_BIOS
$CPU = Get-CimInstance -ClassName Win32_Processor
$DiskDrives = Get-CimInstance -ClassName Win32_DiskDrive | Where-Object { $_.InterfaceType -notlike "*USB*" }

# Get Windows Product ID
$ProductID = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name ProductId -ErrorAction SilentlyContinue).ProductId

# Calculate RAM in GB
$RAMInGB = [math]::Round($ComputerSystem.TotalPhysicalMemory / 1GB)

# Create report header
$Report = @"
====================================================
--- IT DEVICE INVENTORY DATA ---
Date/Time: $DateStr  $TimeStr
====================================================

[1] Brand: $($ComputerSystemProduct.Vendor)

[2] Model: $($ComputerSystemProduct.Name)

[3] Spec:
    CPU: $($CPU.Name)
    RAM: $RAMInGB GB
"@

# Add HDD/SSD information
foreach ($Disk in $DiskDrives) {
    $DiskSizeGB = [math]::Round($Disk.Size / 1GB, 2)
    $Report += "`n    HDD/SSD: $($Disk.Caption) ($DiskSizeGB GB) $($Disk.InterfaceType)"
}

$Report += @"

[4] Operating System: $($OS.Caption) $($OS.OSArchitecture)

[5] Device ID: $($ComputerSystemProduct.UUID)

[6] Product ID: $ProductID

[7] Serial No: $($BIOS.SerialNumber)

[8] Device Name: $($ComputerSystem.Name)

[9] User Email: $env:USERNAME

[10] Microsoft Office:
"@

# Check Microsoft Office installation
$OfficeFound = $false

# Check Office version from Registry (Office 2016, 2019, 2021, 365)
$OfficeVersions = @("16.0", "15.0", "14.0")
foreach ($Version in $OfficeVersions) {
    $RegPath = "HKLM:\SOFTWARE\Microsoft\Office\$Version\Registration"
    if (Test-Path $RegPath) {
        $SubKeys = Get-ChildItem -Path $RegPath -ErrorAction SilentlyContinue
        foreach ($SubKey in $SubKeys) {
            $ProductName = (Get-ItemProperty -Path $SubKey.PSPath -Name ProductName -ErrorAction SilentlyContinue).ProductName
            if ($ProductName) {
                $Report += "`n    Version: $ProductName"
                $OfficeFound = $true
                break
            }
        }
        if ($OfficeFound) { break }
    }
}

# Get Product ID
foreach ($Version in $OfficeVersions) {
    $RegPath = "HKLM:\SOFTWARE\Microsoft\Office\$Version\Registration"
    if (Test-Path $RegPath) {
        $SubKeys = Get-ChildItem -Path $RegPath -ErrorAction SilentlyContinue
        foreach ($SubKey in $SubKeys) {
            $ProductID = (Get-ItemProperty -Path $SubKey.PSPath -Name ProductID -ErrorAction SilentlyContinue).ProductID
            if ($ProductID) {
                $Report += "`n    Product ID: $ProductID"
                break
            }
        }
        if ($ProductID) { break }
    }
}

# Check License Status
try {
    $OfficeLicense = Get-CimInstance -Query "SELECT * FROM SoftwareLicensingProduct WHERE ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND LicenseStatus=1" -ErrorAction Stop | 
                     Where-Object { $_.Name -like "*Office*" } | 
                     Select-Object -First 1
    if ($OfficeLicense) {
        $Report += "`n    License Status: Licensed"
    }
} catch {
    # License check failed
}

# Check Click-to-Run installation (Office 365)
$C2RPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
if (Test-Path $C2RPath) {
    $Edition = (Get-ItemProperty -Path $C2RPath -Name ProductReleaseIds -ErrorAction SilentlyContinue).ProductReleaseIds
    if ($Edition) {
        $Report += "`n    Edition: $Edition"
        $OfficeFound = $true
    }
}

if (-not $OfficeFound) {
    $Report += "`n    Status: Not installed or not detected"
}

$Report += "`n"

# Save report to file
$Report | Out-File -FilePath $FullPath -Encoding UTF8

# Display success message
Write-Host ""
Write-Host "====================================================" -ForegroundColor Green
Write-Host "SUCCESS! Report saved to:" -ForegroundColor Green
Write-Host $FullPath -ForegroundColor Cyan
Write-Host "====================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Press any key to close this window..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

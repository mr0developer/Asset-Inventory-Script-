# Get the USB drive letter (location of script)
$usbDrive = Split-Path -Path $PSScriptRoot -Qualifier
$outputFile = Join-Path $usbDrive "asset_inventory.txt"

$info = @()
$info += "ASSET INVENTORY REPORT"
$info += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$info += "----------------------------------------"

# Windows Version
$os = Get-CimInstance Win32_OperatingSystem
$info += "Windows Version: $($os.Caption) ($($os.Version))"

# Windows Activation Status
$winLic = Get-CimInstance -Query "SELECT * FROM SoftwareLicensingProduct WHERE PartialProductKey IS NOT NULL AND Name LIKE 'Windows%'" -ErrorAction SilentlyContinue
if ($winLic) {
    switch ($winLic.LicenseStatus) {
        0 { $status = "Unlicensed" }
        1 { $status = "Licensed" }
        2 { $status = "Out-of-Box Grace Period" }
        3 { $status = "Out-of-Tolerance Grace Period" }
        4 { $status = "Non-Genuine Grace Period" }
        Default { $status = "Unknown" }
    }
    $info += "Windows Activation Status: $status"
} else {
    $info += "Windows Activation Status: Not Found"
}

# Office Detection & Version
$officeFound = $false
$officePaths = @(
    "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",  # Office 365/2019
    "HKLM:\SOFTWARE\Microsoft\Office",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office"
)
foreach ($path in $officePaths) {
    if (Test-Path $path) {
        $officeFound = $true
        try {
            $props = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
            if ($props.VersionToReport) { $info += "Microsoft Office Version: $($props.VersionToReport)" }
            elseif ($props.Version) { $info += "Microsoft Office Version: $($props.Version)" }
        } catch {}
    }
}

if (-not $officeFound) {
    # Try modern UWP Office apps
    $officeApps = Get-AppxPackage -Name "*Office*" -ErrorAction SilentlyContinue
    if ($officeApps) {
        foreach ($app in $officeApps) {
            $info += "Microsoft Office App: $($app.Name) $($app.Version)"
        }
        $officeFound = $true
    }
}

if (-not $officeFound) {
    $info += "Microsoft Office: Not Installed"
}

# Office Activation
$officeLic = Get-CimInstance -Query "SELECT * FROM SoftwareLicensingProduct WHERE Name LIKE 'Office%'" -ErrorAction SilentlyContinue
if ($officeLic) {
    foreach ($lic in $officeLic) {
        switch ($lic.LicenseStatus) {
            0 { $status = "Unlicensed" }
            1 { $status = "Licensed" }
            2 { $status = "Out-of-Box Grace Period" }
            3 { $status = "Out-of-Tolerance Grace Period" }
            4 { $status = "Non-Genuine Grace Period" }
            Default { $status = "Unknown" }
        }
        $info += "Office Product: $($lic.Name)"
        $info += "Office Activation Status: $status"
    }
} else {
    $info += "Office Activation: Not Found"
}

# Machine Serial Number
$serial = (Get-CimInstance Win32_BIOS).SerialNumber
$info += "Machine Serial Number: $serial"

# OneDrive Presence
$oneDriveFound = $false
$oneDrivePaths = @(
    "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe",
    "${env:ProgramFiles}\Microsoft OneDrive\OneDrive.exe",
    "${env:ProgramFiles(x86)}\Microsoft OneDrive\OneDrive.exe"
)
foreach ($odPath in $oneDrivePaths) {
    if (Test-Path $odPath) {
        $oneDriveFound = $true
        break
    }
}
if ($oneDriveFound) {
    $info += "OneDrive: Installed"
} else {
    $info += "OneDrive: Not Installed"
}

# Antivirus & Install Date (Improved)
$antivirus = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct -ErrorAction SilentlyContinue
if ($antivirus) {
    foreach ($av in $antivirus) {
        $info += "Antivirus: $($av.displayName)"
        $dateFound = $false

        # Check registry for 3rd-party AV install date
        $regPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        )
        foreach ($regPath in $regPaths) {
            if (Test-Path $regPath) {
                $avKeys = Get-ChildItem $regPath -ErrorAction SilentlyContinue
                foreach ($key in $avKeys) {
                    $props = Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue
                    if ($props.DisplayName -and $props.DisplayName -like "*$($av.displayName)*") {
                        if ($props.InstallDate) {
                            try {
                                $dateStr = $props.InstallDate.ToString()
                                if ($dateStr.Length -eq 8) {
                                    $installDate = [datetime]::ParseExact($dateStr, "yyyyMMdd", $null).ToString("yyyy-MM-dd")
                                    $info += "Antivirus Install Date: $installDate"
                                    $dateFound = $true
                                    break
                                }
                            } catch {}
                        }
                    }
                }
            }
            if ($dateFound) { break }
        }

        # Fallback for Defender
        if (-not $dateFound -and $av.displayName -like "*Defender*") {
            try {
                $firstEvent = Get-WinEvent -LogName "Microsoft-Windows-Windows Defender/Operational" -MaxEvents 2000 -ErrorAction SilentlyContinue |
                    Sort-Object TimeCreated |
                    Select-Object -First 1
                if ($firstEvent) {
                    $info += "Antivirus Install Date: $($firstEvent.TimeCreated.ToString('yyyy-MM-dd'))"
                    $dateFound = $true
                }
            } catch {}
        }

        if (-not $dateFound) { $info += "Antivirus Install Date: Unknown" }
    }
} else {
    $info += "Antivirus: Not Found"
}

# RAM Info
$ramGB = [math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
$info += "Installed RAM: $ramGB GB"

# Save output
$info | Out-File -FilePath $outputFile -Encoding UTF8
Write-Host "Asset inventory saved to $outputFile"

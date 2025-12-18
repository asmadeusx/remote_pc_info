# RemotePCInfo_ENG.ps1 - Optimized script for collecting remote PC information

# Determine the script folder path
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$computersFilePath = Join-Path $scriptPath "computers.txt"
$outputFilePath = Join-Path $scriptPath "CollectedPCInfo.txt"

# Search for PsExec (preference is given to the 64-bit version)
$psexecPath = $null
$psexec64Path = Join-Path $scriptPath "PsExec64.exe"
$psexec32Path = Join-Path $scriptPath "PsExec.exe"

if (Test-Path $psexec64Path) {
    $psexecPath = $psexec64Path
}
elseif (Test-Path $psexec32Path) {
    $psexecPath = $psexec32Path
}
else {
    Write-Host "ERROR: PsExec not found!" -ForegroundColor Red
    Write-Host "Place one of the files in the script folder ($scriptPath):" -ForegroundColor Yellow
    Write-Host "  - PsExec64.exe (preferred)" -ForegroundColor Yellow
    Write-Host "  - PsExec.exe" -ForegroundColor Yellow
    exit 1
}

# Function to check computer availability via ping
function Test-ComputerOnline {
    param([string]$ComputerIP)
    
    try {
        $pingResult = Test-Connection -ComputerName $ComputerIP -Count 1 -Quiet -ErrorAction Stop
        return $pingResult
    }
    catch {
        return $false
    }
}

# Function to write to both file and console simultaneously
function Write-OutputBoth {
    param([string]$Message, [string]$FilePath, [switch]$NoNewLine)
    
    # Write to console
    if ($NoNewLine) {
        Write-Host $Message -NoNewline
    }
    else {
        Write-Host $Message
    }
    
    # Write to file
    if ($NoNewLine) {
        $Message | Out-File -FilePath $FilePath -Encoding UTF8 -Append -NoNewline
    }
    else {
        $Message | Out-File -FilePath $FilePath -Encoding UTF8 -Append
    }
}

# Function for capturing PsExec output and writing to file
function Invoke-RemoteCommand {
    param([string]$RemotePCIP, [string]$Command)
    
    $output = & $psexecPath \\$RemotePCIP -accepteula -nobanner powershell $Command 2>&1
    
    # Filter PsExec error messages
    $cleanOutput = $output | Where-Object { $_ -is [string] } | ForEach-Object {
        if ($_ -notmatch '^PsExec' -and $_ -notmatch '^Copyright' -and $_ -notmatch '^Sysinternals') {
            $_
        }
    }
    
    return $cleanOutput -join "`n"
}

# Main information collection function
function Get-RemotePCInfo {
    param([string]$RemotePCIP, [string]$OutputFile)

    Write-OutputBoth "`n=== COLLECTING INFORMATION ABOUT REMOTE PC ($RemotePCIP) ===" -FilePath $OutputFile
    Write-OutputBoth "" -FilePath $OutputFile

    # 1. General system information
    Write-OutputBoth "1. GENERAL SYSTEM INFORMATION:" -FilePath $OutputFile
    $systemInfo = Invoke-RemoteCommand -RemotePCIP $RemotePCIP -Command @'
$system = Get-WmiObject -Class Win32_ComputerSystem
Write-Host "Computer Name: $($system.Name)"
Write-Host "   Domain: $($system.Domain)"
'@
    Write-OutputBoth $systemInfo -FilePath $OutputFile
    Write-OutputBoth "" -FilePath $OutputFile

    # 2. Operating system age information
    Write-OutputBoth "2. OPERATING SYSTEM AGE INFORMATION:" -FilePath $OutputFile
    $osInfo = Invoke-RemoteCommand -RemotePCIP $RemotePCIP -Command @'
$os = Get-WmiObject -Class Win32_OperatingSystem
$installDate = [Management.ManagementDateTimeConverter]::ToDateTime($os.InstallDate)
$lastBootTime = [Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime)
$localDateTime = [Management.ManagementDateTimeConverter]::ToDateTime($os.LocalDateTime)

$currentDate = $localDateTime
$daysSinceInstall = [math]::Round(($currentDate - $installDate).TotalDays, 0)

Write-Host "Operating System: $($os.Caption)"
Write-Host "   Version: $($os.Version)"
Write-Host "   Installation Date: $($installDate.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "   Last Boot Time: $($lastBootTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "   Current Time: $($localDateTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "   Days Since Installation: $daysSinceInstall"
'@
    Write-OutputBoth $osInfo -FilePath $OutputFile
    Write-OutputBoth "" -FilePath $OutputFile

    # 3. Motherboard information
    Write-OutputBoth "3. MOTHERBOARD INFORMATION:" -FilePath $OutputFile
    $motherboard = Invoke-RemoteCommand -RemotePCIP $RemotePCIP -Command @'
$board = Get-WmiObject -Class Win32_BaseBoard
Write-Host "Manufacturer: $($board.Manufacturer)"
Write-Host "   Model: $($board.Product)"
Write-Host "   Serial Number: $($board.SerialNumber)"
'@
    Write-OutputBoth $motherboard -FilePath $OutputFile
    Write-OutputBoth "" -FilePath $OutputFile

    # 4. Processor information
    Write-OutputBoth "4. PROCESSOR INFORMATION:" -FilePath $OutputFile
    $processor = Invoke-RemoteCommand -RemotePCIP $RemotePCIP -Command @'
$cpu = Get-WmiObject -Class Win32_Processor
Write-Host "Model: $($cpu.Name)"
Write-Host "   Number of Cores: $($cpu.NumberOfCores)"
Write-Host "   Maximum Clock Speed: $($cpu.MaxClockSpeed) MHz"
'@
    Write-OutputBoth $processor -FilePath $OutputFile
    Write-OutputBoth "" -FilePath $OutputFile

    # 5. RAM information
    Write-OutputBoth "5. RAM INFORMATION:" -FilePath $OutputFile
    $memory = Invoke-RemoteCommand -RemotePCIP $RemotePCIP -Command @'
$memoryModules = Get-WmiObject -Class Win32_PhysicalMemory
$totalGB = 0

foreach ($module in $memoryModules) {
    $sizeGB = [math]::Round($module.Capacity / 1GB, 2)
    $totalGB += $sizeGB
    Write-Host "Module: $($module.DeviceLocator)"
    Write-Host "   Size: $sizeGB GB"
    Write-Host "   Manufacturer: $($module.Manufacturer)"
    Write-Host "   Serial Number: $($module.SerialNumber)"
    Write-Host "   Part Number: $($module.PartNumber)"
    Write-Host ""
}

Write-Host "Total Capacity: $totalGB GB"
'@
    Write-OutputBoth $memory -FilePath $OutputFile
    Write-OutputBoth "" -FilePath $OutputFile

    # 6. Disk information
    Write-OutputBoth "6. DISK INFORMATION:" -FilePath $OutputFile
    $disks = Invoke-RemoteCommand -RemotePCIP $RemotePCIP -Command @'
$disksInfo = Get-WmiObject -Class Win32_DiskDrive

foreach ($disk in $disksInfo) {
    $sizeGB = [math]::Round($disk.Size / 1GB, 2)
    Write-Host "Model: $($disk.Model)"
    Write-Host "   Size: $sizeGB GB"
    Write-Host "   Interface: $($disk.InterfaceType)"
    Write-Host "   Media Type: $($disk.MediaType)"
    Write-Host "   Serial Number: $($disk.SerialNumber)"
    Write-Host ""
}
'@
    Write-OutputBoth $disks -FilePath $OutputFile
    Write-OutputBoth "" -FilePath $OutputFile

    # 7. Video adapter information
    Write-OutputBoth "7. VIDEO ADAPTER INFORMATION:" -FilePath $OutputFile
    $gpu = Invoke-RemoteCommand -RemotePCIP $RemotePCIP -Command @'
$videoCards = Get-WmiObject -Class Win32_VideoController

foreach ($card in $videoCards) {
    $vramGB = [math]::Round($card.AdapterRAM / 1GB, 2)
    Write-Host "Video Card: $($card.Name)"
    Write-Host "   Video Memory: $vramGB GB"
    Write-Host "   Driver Version: $($card.DriverVersion)"
    Write-Host "   Video Processor: $($card.VideoProcessor)"
    Write-Host ""
}
'@
    Write-OutputBoth $gpu -FilePath $OutputFile
    Write-OutputBoth "" -FilePath $OutputFile

    # 8. BIOS/UEFI information
    Write-OutputBoth "8. BIOS (UEFI) INFORMATION:" -FilePath $OutputFile
    $biosInfo = Invoke-RemoteCommand -RemotePCIP $RemotePCIP -Command @'
$bios = Get-WmiObject -Class Win32_BIOS
$releaseDate = [Management.ManagementDateTimeConverter]::ToDateTime($bios.ReleaseDate)
$currentDate = Get-Date
$daysSinceRelease = [math]::Round(($currentDate - $releaseDate).TotalDays, 0)

Write-Host "BIOS Manufacturer: $($bios.Manufacturer)"
Write-Host "   Version: $($bios.SMBIOSBIOSVersion)"
Write-Host "   Serial Number: $($bios.SerialNumber)"
Write-Host "   Firmware Version: $($bios.Version)"
Write-Host "   Release Date: $($releaseDate.ToString('yyyy-MM-dd'))"
Write-Host "   Days Since Release: $daysSinceRelease"
'@
    Write-OutputBoth $biosInfo -FilePath $OutputFile
    Write-OutputBoth "" -FilePath $OutputFile

    Write-OutputBoth "=== INFORMATION COLLECTION COMPLETED ===" -FilePath $OutputFile
}

# --- MAIN SCRIPT CODE ---

Write-Host "=== RemotePCInfo Script ===" -ForegroundColor Cyan
Write-Host ""

# Display information about the PsExec version used AFTER the header
if (Test-Path $psexec64Path) {
    Write-Host "Using: PsExec64.exe" -ForegroundColor Green
}
elseif (Test-Path $psexec32Path) {
    Write-Host "Using: PsExec.exe" -ForegroundColor Green
}
Write-Host "Script path: $scriptPath" -ForegroundColor Gray
Write-Host ""

# Clear output file before starting work
if (Test-Path $outputFilePath) {
    Write-Host "Clearing previous output file: $outputFilePath" -ForegroundColor Yellow
}
else {
    Write-Host "Creating output file: $outputFilePath" -ForegroundColor Green
}

# Create/clear file and write header
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
"=== RemotePCInfo - Report from $timestamp ===" | Out-File -FilePath $outputFilePath -Encoding UTF8
"=== Generated by script: $($MyInvocation.MyCommand.Name)" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
"" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append

# Check for computers.txt file
if (-not (Test-Path $computersFilePath)) {
    Write-Host "computers.txt file not found in the script folder." -ForegroundColor Yellow
    Write-Host "Creating computers.txt file with example..." -ForegroundColor Yellow
    
    @"
# File with list of computers for scanning
# Each IP address on a new line
# Lines starting with # are ignored

# Examples:
192.168.1.100
192.168.1.101
# 192.168.1.102 - this computer is commented out
"@ | Out-File -FilePath $computersFilePath -Encoding UTF8
    
    Write-Host "File created: $computersFilePath" -ForegroundColor Green
    Write-Host "Add IP addresses to the file and run the script again." -ForegroundColor Yellow
    exit 0
}

# Read computers.txt file
Write-Host "Reading computer list from: $computersFilePath" -ForegroundColor Cyan
$computers = Get-Content $computersFilePath | Where-Object {
    $_ -notmatch '^\s*#' -and $_ -match '\S' -and $_ -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$'
}

if ($computers.Count -eq 0) {
    Write-Host "ERROR: No valid IP addresses found in computers.txt file!" -ForegroundColor Red
    Write-Host "Format: each IP address on a new line (e.g.: 192.168.1.100)" -ForegroundColor Yellow
    exit 1
}

Write-Host "Found computers to check: $($computers.Count)" -ForegroundColor Green
Write-Host "Results will be saved to: $outputFilePath" -ForegroundColor Green
Write-Host ""

# Write scanning information to file
"=== Scan Start ===" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
"Start Time: $timestamp" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
"Number of computers to check: $($computers.Count)" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
"" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append

# Process each computer
$successCount = 0
$skippedCount = 0

foreach ($computer in $computers) {
    $computer = $computer.Trim()
    
    Write-Host "Checking computer: $computer" -ForegroundColor Cyan
    "`n--- Checking computer: $computer ---" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
    
    # Check availability via ping
    if (Test-ComputerOnline -ComputerIP $computer) {
        Write-Host "Computer is available, starting information collection..." -ForegroundColor Green
        "Status: available, collecting information..." | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
        Write-Host ""
        
        try {
            Get-RemotePCInfo -RemotePCIP $computer -OutputFile $outputFilePath
            $successCount++
            "Status: successfully processed" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
        }
        catch {
            $errorMsg = "ERROR while collecting information from $computer : $_"
            Write-Host $errorMsg -ForegroundColor Red
            $errorMsg | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
            Write-Host ""
        }
    }
    else {
        $warningMsg = "WARNING: Computer $computer is unavailable (ping), skipping..."
        Write-Host $warningMsg -ForegroundColor Yellow
        $warningMsg | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
        Write-Host ""
        $skippedCount++
    }
    
    # Pause between computers (optional)
    if ($computer -ne $computers[-1]) {
        Start-Sleep -Seconds 1
    }
}

# Final statistics
Write-Host "=== FINAL STATISTICS ===" -ForegroundColor Cyan
Write-Host "Successfully processed: $successCount computers" -ForegroundColor Green
Write-Host "Skipped (unavailable): $skippedCount computers" -ForegroundColor Yellow
Write-Host "Total in list: $($computers.Count) computers" -ForegroundColor White
Write-Host ""
Write-Host "Results saved to file: $outputFilePath" -ForegroundColor Green

# Write final statistics to file
"`n=== FINAL STATISTICS ===" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
"Successfully processed: $successCount computers" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
"Skipped (unavailable): $skippedCount computers" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
"Total in list: $($computers.Count) computers" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
"`n=== Scanning Completed ===" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
$endTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
"End Time: $endTime" | Out-File -FilePath $outputFilePath -Encoding UTF8 -Append
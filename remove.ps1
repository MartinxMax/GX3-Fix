[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = New-Object -TypeName System.Text.UTF8Encoding
$env:LANG = "en_US.UTF-8"
Write-Host "Maptnh @　https://github.com/MartinxMax/" -ForegroundColor Cyan
$softwareList = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" | 
    Get-ItemProperty | 
    Where-Object { $_.Publisher -like "*MITSUBISHI ELECTRIC CORPORATION*" } | 
    Select-Object DisplayName, UninstallString

if ($softwareList.Count -eq 0) {
    Write-Host "No Mitsubishi Electric software found" -ForegroundColor Gray
    return
}

Write-Host "Found the following Mitsubishi Electric software:" -ForegroundColor Cyan
$softwareList | ForEach-Object {
    Write-Host "→ $($_.DisplayName)" -ForegroundColor Yellow
}

$choice = Read-Host "Do you want to uninstall? (Y/N)"
if ($choice -notmatch "Y|y") {
    Write-Host "Operation cancelled" -ForegroundColor Gray
    return
}

$softwareList | ForEach-Object {
    $uninstallCmd = $_.UninstallString -replace '"', ''  # Remove quotes
    Write-Host "Uninstalling: $($_.DisplayName)" -ForegroundColor Green

    Start-Process -FilePath "cmd.exe" -ArgumentList "/c $uninstallCmd" -Wait
    Write-Host "Uninstallation complete: $($_.DisplayName)" -ForegroundColor Green
}

Write-Host "Starting to check and remove Mitsubishi Electric related services..." -ForegroundColor Cyan
$mitsubishiServices = Get-Service | Where-Object { 
    $_.DisplayName -like "*MITSUBISHI*" -or $_.Name -like "*MELSOFT*" 
}

if ($mitsubishiServices.Count -gt 0) {
    $mitsubishiServices | ForEach-Object {
        $serviceName = $_.Name
        $serviceDisplay = $_.DisplayName
        
        try {
            if ($_.Status -eq "Running") {
                Stop-Service -Name $serviceName -Force -ErrorAction Stop
                Write-Host "Service stopped: $serviceDisplay ($serviceName)" -ForegroundColor Green
            }

            $regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\$serviceName"
            if (Test-Path $regPath) {
                Remove-Item -Path $regPath -Recurse -Force -ErrorAction Stop
                Write-Host "Service registry removed: $serviceDisplay ($serviceName)" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "Failed to remove service: $serviceDisplay ($serviceName) → $_" -ForegroundColor Red
        }
    }
} else {
    Write-Host "No Mitsubishi Electric related services detected" -ForegroundColor Gray
}

Write-Host "Starting to clean up residual files..." -ForegroundColor Cyan
Get-ChildItem -Path C:\ -Recurse -Include *MELSOFT*, *MITSUBISHI* -ErrorAction SilentlyContinue | 
    Where-Object { 
        $_.VersionInfo.CompanyName -like "*MITSUBISHI ELECTRIC CORPORATION*" -or 
        $_.FullName -match "MELSOFT|MITSUBISHI" 
    } | 
    ForEach-Object {
        try {
            Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction Stop
            Write-Host "Residual file removed: $($_.FullName)" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to remove file: $($_.FullName) → $_" -ForegroundColor Red
        }
    }

$logoutChoice = Read-Host "Do you want to log off now to apply the changes? (Y/N)"
if ($logoutChoice -match "Y|y") {
    Write-Host "Logging off the system..." -ForegroundColor Green
    Start-Sleep -Seconds 3   
    shutdown /l   
} else {
    Write-Host "Please log off or restart the computer manually to complete the cleanup" -ForegroundColor Gray
}

Write-Host "All operations completed!" -ForegroundColor Cyan

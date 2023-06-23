$filePath = Get-ChildItem -Path "C:\Program Files (x86)" -Recurse -Filter "NinjaRMMAgent.exe" | Select-Object -ExpandProperty FullName

if ($filePath.Count -eq 1) {
    $directory = Split-Path -Path $filePath -Parent
    Set-Location -Path $directory
    
    # Run ninjarmmagent.exe /collectlogs
    Start-Process -FilePath "ninjarmmagent.exe" -ArgumentList "/collectlogs" -Wait

    # Ensure that C:\Temp folder exists
    $logPath = "C:\Temp"
    if (-not (Test-Path -Path $logPath)) {
        New-Item -ItemType Directory -Path $logPath | Out-Null
    }
    
    # Collect logs from Event Viewer applications and system
    $eventLogs = @("Application", "System")
    
    foreach ($eventLog in $eventLogs) {
        $logFile = Join-Path -Path $logPath -ChildPath "$eventLog.log"
        Get-EventLog -LogName $eventLog | Out-File -FilePath $logFile
    }
    
    # Copy ninjalogs.cab to C:\Temp
    $sourceFile = "C:\Windows\temp\ninjalogs.cab"
    $destinationFile = Join-Path -Path $logPath -ChildPath "ninjalogs.cab"
    Copy-Item -Path $sourceFile -Destination $destinationFile -Force
}
else {
    Write-Host "Multiple occurrences of NinjaRMMAgent.exe found."
}

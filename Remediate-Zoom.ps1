# PowerShell script to uninstall Zoom using CleanZoom and reinstall latest version

# Set up logging
$logFile = "$env:TEMP\Reinstall-Zoom.txt"

function Write-Log {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $message" | Out-File -FilePath $logFile -Append
    Write-Host "$timestamp - $message"
}

Write-Log "Starting Zoom Removal and Reinstallation"

# Create temp directory for downloads
$tempDir = Join-Path $env:TEMP "ZoomReinstall_$(Get-Random)"
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
Write-Log "Created temporary directory: $tempDir"

try {
    # Download CleanZoom.zip
    $cleanZoomUrl = "https://assets.zoom.us/docs/msi-templates/CleanZoom.zip"
    $cleanZoomZip = Join-Path $tempDir "CleanZoom.zip"
    Write-Log "Downloading CleanZoom from $cleanZoomUrl"
    
    $webClient = New-Object System.Net.WebClient
    $webClient.DownloadFile($cleanZoomUrl, $cleanZoomZip)
    Write-Log "Downloaded CleanZoom to $cleanZoomZip"
    
    # Extract the zip file
    $cleanZoomExtractPath = Join-Path $tempDir "CleanZoom"
    Write-Log "Extracting CleanZoom to $cleanZoomExtractPath"
    
    Expand-Archive -Path $cleanZoomZip -DestinationPath $cleanZoomExtractPath -Force
    Write-Log "Extraction completed"
    
    # Find CleanZoom executable
    $cleanZoomExe = Get-ChildItem -Path $cleanZoomExtractPath -Recurse -Filter "CleanZoom.exe" | Select-Object -First 1
    
    if ($null -eq $cleanZoomExe) {
        throw "CleanZoom.exe not found in the extracted files"
    }
    
    $cleanZoomExePath = $cleanZoomExe.FullName
    Write-Log "Found CleanZoom executable at: $cleanZoomExePath"
    
    # Run CleanZoom to uninstall Zoom
    Write-Log "Running CleanZoom with /silent parameter to uninstall Zoom"
    Start-Process -FilePath $cleanZoomExePath -ArgumentList "/silent /keep_outlook_plugin /keep_notes_plugin" -Wait
    Write-Log "CleanZoom uninstallation process completed"
    
    # Brief pause to ensure uninstallation completes
    Start-Sleep -Seconds 5
    
    # Download latest Zoom installer
    $zoomInstallerUrl = "https://zoom.us/client/latest/ZoomInstallerFull.msi"
    $zoomInstallerPath = Join-Path $tempDir "ZoomInstallerFull.msi"
    Write-Log "Downloading latest Zoom installer from $zoomInstallerUrl"
    
    $webClient.DownloadFile($zoomInstallerUrl, $zoomInstallerPath)
    Write-Log "Downloaded Zoom installer to $zoomInstallerPath"
    
    # Install Zoom silently
    Write-Log "Installing Zoom silently"
    Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$zoomInstallerPath`" /quiet /qn /norestart MSIRestartManagerControl=Disable /log zoominstall.log" -Wait
    Write-Log "Zoom installation completed"
    
    Write-Log "Zoom reinstallation process completed successfully"
}
catch {
    $errorMessage = $_.Exception.Message
    Write-Log "ERROR: $errorMessage"
    Write-Log "Stack Trace: $($_.Exception.StackTrace)"
    throw $_
}
finally {
    # Clean up temp files
    if (Test-Path $tempDir) {
        Write-Log "Cleaning up temporary files"
        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Log "Script execution completed"

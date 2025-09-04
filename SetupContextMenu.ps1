# Self-elevate the script if required
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process PowerShell -Verb RunAs "-NoProfile -ExecutionPolicy Bypass -Command `"cd '$pwd'; & '$PSCommandPath';`"";
    exit;
}

# First, create the VBS wrapper script (simplified version)
$vbsContent = @'
CreateObject("WScript.Shell").Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & WScript.Arguments(0) & """ -InputFile """ & WScript.Arguments(1) & """ -OutputFormat " & WScript.Arguments(2), 0, False
'@

$vbsPath = "$PSScriptRoot\silent_wrapper.vbs"
$vbsContent | Out-File -FilePath $vbsPath -Encoding ASCII -Force

$extensionActions = @{
    # Image File Types
    ".png" = @{
        ".jpg" = "jpg";
        ".jpeg" = "jpeg";
        ".gif" = "gif";
        ".bmp" = "bmp";
        ".tiff" = "tiff";
        ".ico" = "ico";
        ".webp" = "webp";
        ".jfif" = "jfif";
    };
    ".jpg" = @{
        ".png" = "png";
        ".gif" = "gif";
        ".bmp" = "bmp";
        ".tiff" = "tiff";
        ".ico" = "ico";
        ".webp" = "webp";
        ".jfif" = "jfif";
    };
    ".jpeg" = @{
        ".png" = "png";
        ".gif" = "gif";
        ".bmp" = "bmp";
        ".tiff" = "tiff";
        ".ico" = "ico";
        ".webp" = "webp";
        ".jfif" = "jfif";
    };
    ".gif" = @{
        ".png" = "png";
        ".jpg" = "jpg";
        ".jpeg" = "jpeg";
        ".bmp" = "bmp";
        ".tiff" = "tiff";
        ".ico" = "ico";
        ".webp" = "webp";
        ".jfif" = "jfif";
    };
    ".bmp" = @{
        ".png" = "png";
        ".jpg" = "jpg";
        ".jpeg" = "jpeg";
        ".gif" = "gif";
        ".tiff" = "tiff";
        ".ico" = "ico";
        ".webp" = "webp";
        ".jfif" = "jfif";
    };
    ".tiff" = @{
        ".png" = "png";
        ".jpg" = "jpg";
        ".jpeg" = "jpeg";
        ".gif" = "gif";
        ".bmp" = "bmp";
        ".ico" = "ico";
        ".webp" = "webp";
        ".jfif" = "jfif";
    };
    ".ico" = @{
        ".png" = "png";
        ".jpg" = "jpg";
        ".jpeg" = "jpeg";
        ".gif" = "gif";
        ".bmp" = "bmp";
        ".tiff" = "tiff";
        ".webp" = "webp";
        ".jfif" = "jfif";
    };
    ".webp" = @{
        ".png" = "png";
        ".jpg" = "jpg";
        ".jpeg" = "jpeg";
        ".gif" = "gif";
        ".bmp" = "bmp";
        ".tiff" = "tiff";
        ".ico" = "ico";
        ".jfif" = "jfif";
    };
    ".jfif" = @{
        ".png" = "png";
        ".jpg" = "jpg";
        ".jpeg" = "jpeg";
        ".gif" = "gif";
        ".bmp" = "bmp";
        ".tiff" = "tiff";
        ".ico" = "ico";
        ".webp" = "webp";
    };

    # Video File Types
    ".mp4" = @{
        ".mkv" = "mkv";
        ".avi" = "avi";
        ".mov" = "mov";
        ".wmv" = "wmv";
        ".flv" = "flv";
        ".webm" = "webm";
    };
    ".mkv" = @{
        ".mp4" = "mp4";
        ".avi" = "avi";
        ".mov" = "mov";
        ".wmv" = "wmv";
        ".flv" = "flv";
        ".webm" = "webm";
    };
    ".avi" = @{
        ".mp4" = "mp4";
        ".mkv" = "mkv";
        ".mov" = "mov";
        ".wmv" = "wmv";
        ".flv" = "flv";
        ".webm" = "webm";
    };
    ".mov" = @{
        ".mp4" = "mp4";
        ".mkv" = "mkv";
        ".avi" = "avi";
        ".wmv" = "wmv";
        ".flv" = "flv";
        ".webm" = "webm";
    };
    ".wmv" = @{
        ".mp4" = "mp4";
        ".mkv" = "mkv";
        ".avi" = "avi";
        ".mov" = "mov";
        ".flv" = "flv";
        ".webm" = "webm";
    };
    ".flv" = @{
        ".mp4" = "mp4";
        ".mkv" = "mkv";
        ".avi" = "avi";
        ".mov" = "mov";
        ".wmv" = "wmv";
        ".webm" = "webm";
    };
    ".webm" = @{
        ".mp4" = "mp4";
        ".mkv" = "mkv";
        ".avi" = "avi";
        ".mov" = "mov";
        ".wmv" = "wmv";
        ".flv" = "flv";
    };

    # Audio File Types
    ".mp3" = @{
        ".wav" = "wav";
        ".aac" = "aac";
        ".flac" = "flac";
        ".ogg" = "ogg";
        ".m4a" = "m4a";
        ".aif" = "aif";
        ".aiff" = "aiff";
        ".opus" = "opus";
    };
    ".wav" = @{
        ".mp3" = "mp3";
        ".aac" = "aac";
        ".flac" = "flac";
        ".ogg" = "ogg";
        ".m4a" = "m4a";
        ".aif" = "aif";
        ".aiff" = "aiff";
        ".opus" = "opus";
    };
    ".aac" = @{
        ".mp3" = "mp3";
        ".wav" = "wav";
        ".flac" = "flac";
        ".ogg" = "ogg";
        ".m4a" = "m4a";
        ".aif" = "aif";
        ".aiff" = "aiff";
        ".opus" = "opus";
    };
    ".flac" = @{
        ".mp3" = "mp3";
        ".wav" = "wav";
        ".aac" = "aac";
        ".ogg" = "ogg";
        ".m4a" = "m4a";
        ".aif" = "aif";
        ".aiff" = "aiff";
        ".opus" = "opus";
    };
    ".ogg" = @{
        ".mp3" = "mp3";
        ".wav" = "wav";
        ".aac" = "aac";
        ".flac" = "flac";
        ".m4a" = "m4a";
        ".aif" = "aif";
        ".aiff" = "aiff";
        ".opus" = "opus";
    };
    ".m4a" = @{
        ".mp3" = "mp3";
        ".wav" = "wav";
        ".aac" = "aac";
        ".flac" = "flac";
        ".ogg" = "ogg";
        ".aif" = "aif";
        ".aiff" = "aiff";
        ".opus" = "opus";
    };
    ".aif" = @{
        ".mp3" = "mp3";
        ".wav" = "wav";
        ".aac" = "aac";
        ".flac" = "flac";
        ".ogg" = "ogg";
        ".m4a" = "m4a";
        ".aiff" = "aiff";
        ".opus" = "opus";
    };
    ".aiff" = @{
        ".mp3" = "mp3";
        ".wav" = "wav";
        ".aac" = "aac";
        ".flac" = "flac";
        ".ogg" = "ogg";
        ".m4a" = "m4a";
        ".aif" = "aif";
        ".opus" = "opus";
    };
    ".opus" = @{
        ".mp3" = "mp3";
        ".wav" = "wav";
        ".aac" = "aac";
        ".flac" = "flac";
        ".ogg" = "ogg";
        ".m4a" = "m4a";
        ".aif" = "aif";
        ".aiff" = "aiff";
    };
}

foreach ($ext in $extensionActions.Keys) {
    $shellKey = "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\$ext\shell"
    
    $convertToKey = "$shellKey\Convert To"
    New-Item -Path $convertToKey -Force | Out-Null
    Set-ItemProperty -Path $convertToKey -Name "(Default)" -Value "Convert To"
    Set-ItemProperty -Path $convertToKey -Name "Icon" -Value "shell32.dll,16"
    Set-ItemProperty -Path $convertToKey -Name "ExtendedSubCommandsKey" -Value "SystemFileAssociations\$ext\shell\Convert To\conversion"

    $subMenuKey = "$convertToKey\conversion"
    New-Item -Path $subMenuKey -Force | Out-Null

    foreach ($target in $extensionActions[$ext].Keys) {
        $targetKey = "$subMenuKey\shell\to$($extensionActions[$ext][$target])"
        New-Item -Path $targetKey -Force | Out-Null
        Set-ItemProperty -Path $targetKey -Name "(Default)" -Value "Convert to $($target.ToUpper())"
        
        $commandKey = "$targetKey\command"
        New-Item -Path $commandKey -Force | Out-Null
        
        # Use WScript to run completely silently
        $command = "wscript.exe `"$vbsPath`" `"$PSScriptRoot\convert.ps1`" `"%1`" `"$($extensionActions[$ext][$target])`""
        Set-ItemProperty -Path $commandKey -Name "(Default)" -Value $command
    }
}

Write-Host "Context menu entries have been created successfully."
Write-Host "Press any key to continue..."
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
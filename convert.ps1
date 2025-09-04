param(
    [string]$InputFile,
    [string]$OutputFormat
)

function Show-Notification {
    param(
        [string]$Title,
        [string]$Message
    )
    
    Add-Type -AssemblyName System.Windows.Forms
    $global:balloon = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balloon.Visible = $true
    $balloon.ShowBalloonTip(5000, $Title, $Message, [System.Windows.Forms.ToolTipIcon]::Info)
}

function Convert-Media {
    param(
        [string]$InputPath,
        [string]$Format
    )

    $extensionMap = @{
        'Image' = @{
            'png' = '.png'
            'jpg' = '.jpg'
            'jpeg' = '.jpeg'
            'gif' = '.gif'
            'bmp' = '.bmp'
            'tiff' = '.tiff'
            'ico' = '.ico'
            'webp' = '.webp'
            'jfif' = '.jfif'
        }
        'Video' = @{
            'mp4' = '.mp4'
            'mkv' = '.mkv'
            'avi' = '.avi'
            'mov' = '.mov'
            'wmv' = '.wmv'
            'flv' = '.flv'
            'webm' = '.webm'
        }
        'Audio' = @{
            'mp3' = '.mp3'
            'wav' = '.wav'
            'aac' = '.aac'
            'flac' = '.flac'
            'ogg' = '.ogg'
            'm4a' = '.m4a'
            'aif' = '.aif'
            'aiff' = '.aiff'
            'opus' = '.opus'
        }
    }

    $fileType = Switch -Regex ($InputPath) {
        '\.(gif|png|jpe?g|bmp|tiff?|ico|webp|jfif)$' { 'Image' }
        '\.(mp4|mkv|avi|mov|wmv|flv|webm)$' { 'Video' }
        '\.(wav|mp3|aac|flac|ogg|m4a|aif|aiff|opus)$' { 'Audio' }
    }

    $outputExtension = $extensionMap[$fileType][$Format]
    if ($outputExtension -ne $null) {
        $outputPath = [System.IO.Path]::ChangeExtension($InputPath, $outputExtension)
        try {
            & ffmpeg -i $InputPath $outputPath -y
            # Show-Notification "Conversion Complete" "Successfully converted $(Split-Path $InputPath -Leaf) to $Format"
        }
        catch {
            Show-Notification "Conversion Error" "Failed to convert file: $_"
        }
    }
    else {
        Show-Notification "Conversion Error" "Invalid format for the type of file."
    }
}

Convert-Media -InputPath $InputFile -Format $OutputFormat
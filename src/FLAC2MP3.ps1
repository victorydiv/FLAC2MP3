# Registry Key for Settings
$registryKey = "HKCU:\Software\FLAC2MP3"
Add-Type -AssemblyName System.Windows.Forms
# Ensure the Registry Key Exists
if (-not (Test-Path $registryKey)) {
    New-Item -Path $registryKey -Force | Out-Null
}

# Function to Prompt for Missing Registry Values
function Initialize-Settings {
    Write-Host "Checking for required settings in the registry..." -ForegroundColor Cyan

    # Check FLAC Folder
    $flacFolder = (Get-ItemProperty -Path $registryKey -Name FlacFolder -ErrorAction SilentlyContinue).FlacFolder
    if (-not $flacFolder) {
        Write-Host "FLAC folder location is not set."
        do {
            $flacFolder = Read-Host "Enter the FLAC folder location"
            if (-not (Test-Path -Path $flacFolder)) {
                Write-Warning "The specified folder does not exist. Please try again."
                $flacFolder = $null
            }
        } while (-not $flacFolder)
        Set-ItemProperty -Path $registryKey -Name FlacFolder -Value $flacFolder
        Write-Host "FLAC folder location saved: $flacFolder" -ForegroundColor Green
    }

    # Check MP3 Folder
    $mp3Folder = (Get-ItemProperty -Path $registryKey -Name Mp3Folder -ErrorAction SilentlyContinue).Mp3Folder
    if (-not $mp3Folder) {
        Write-Host "MP3 folder location is not set."
        do {
            $mp3Folder = Read-Host "Enter the MP3 folder location"
            if (-not (Test-Path -Path $mp3Folder)) {
                Write-Warning "The specified folder does not exist. Please try again."
                $mp3Folder = $null
            }
        } while (-not $mp3Folder)
        Set-ItemProperty -Path $registryKey -Name Mp3Folder -Value $mp3Folder
        Write-Host "MP3 folder location saved: $mp3Folder" -ForegroundColor Green
    }

    # Check ffmpeg Path
    $ffmpegPath = (Get-ItemProperty -Path $registryKey -Name FfmpegPath -ErrorAction SilentlyContinue).FfmpegPath
    if (-not $ffmpegPath) {
        Write-Host "ffmpeg path is not set."
        do {
            $ffmpegPath = Read-Host "Enter the ffmpeg executable path"
            if (-not (Test-Path -Path $ffmpegPath)) {
                Write-Warning "The specified path does not exist. Please try again."
                $ffmpegPath = $null
            }
        } while (-not $ffmpegPath)
        Set-ItemProperty -Path $registryKey -Name FfmpegPath -Value $ffmpegPath
        Write-Host "ffmpeg path saved: $ffmpegPath" -ForegroundColor Green
    }

    # Check ffprobe Path
    $ffprobePath = (Get-ItemProperty -Path $registryKey -Name FfprobePath -ErrorAction SilentlyContinue).FfprobePath
    if (-not $ffprobePath) {
        Write-Host "ffprobe path is not set."
        do {
            $ffprobePath = Read-Host "Enter the ffprobe executable path"
            if (-not (Test-Path -Path $ffprobePath)) {
                Write-Warning "The specified path does not exist. Please try again."
                $ffprobePath = $null
            }
        } while (-not $ffprobePath)
        Set-ItemProperty -Path $registryKey -Name FfprobePath -Value $ffprobePath
        Write-Host "ffprobe path saved: $ffprobePath" -ForegroundColor Green
    }

    # Check Plex Token
    $plexToken = (Get-ItemProperty -Path $registryKey -Name PlexToken -ErrorAction SilentlyContinue).PlexToken
    if (-not $plexToken) {
        Write-Host "Plex token is not set."
        do {
            $plexToken = Read-Host "Enter the Plex token"
            if (-not $plexToken) {
                Write-Warning "The Plex token cannot be empty. Please try again."
            }
        } while (-not $plexToken)
        Set-ItemProperty -Path $registryKey -Name PlexToken -Value $plexToken
        Write-Host "Plex token saved: $plexToken" -ForegroundColor Green
    }

    # Check Minimum Bitrate
    $minBitrate = (Get-ItemProperty -Path $registryKey -Name MinBitrate -ErrorAction SilentlyContinue).MinBitrate
    if (-not $minBitrate) {
        $minBitrate = 320 # Default value
        Set-ItemProperty -Path $registryKey -Name MinBitrate -Value $minBitrate
        Write-Host "Minimum bitrate set to default: $minBitrate kbps" -ForegroundColor Green
    }

    # Check Advanced Renamer Path
    $advancedRenamerPath = (Get-ItemProperty -Path $registryKey -Name AdvancedRenamerPath -ErrorAction SilentlyContinue).AdvancedRenamerPath
    if (-not $advancedRenamerPath) {
        Write-Host "Advanced Renamer path is not set. Please set it in the Options Menu."
    }

    # Check MP3Tag Path
    $mp3TagPath = (Get-ItemProperty -Path $registryKey -Name Mp3TagPath -ErrorAction SilentlyContinue).Mp3TagPath
    if (-not $mp3TagPath) {
        Write-Host "MP3Tag path is not set. Please set it in the Options Menu."
    }

    # Check MP3 File Count
    $mp3FileCount = (Get-ItemProperty -Path $registryKey -Name Mp3FileCount -ErrorAction SilentlyContinue).Mp3FileCount
    if (-not $mp3FileCount) {
        $mp3FileCount = 0 # Default value
        Set-ItemProperty -Path $registryKey -Name Mp3FileCount -Value $mp3FileCount
        Write-Host "MP3 file count initialized to: $mp3FileCount" -ForegroundColor Green
    }
}

# Initialize Settings
Initialize-Settings

# Load Folder Settings from Registry
$flacFolder = (Get-ItemProperty -Path $registryKey -Name FlacFolder -ErrorAction SilentlyContinue).FlacFolder
if (-not $flacFolder) {
    $flacFolder = "C:\Users\sean\Downloads\put.io\FLAC" # Default value
    Set-ItemProperty -Path $registryKey -Name FlacFolder -Value $flacFolder
}

$mp3Folder = (Get-ItemProperty -Path $registryKey -Name Mp3Folder -ErrorAction SilentlyContinue).Mp3Folder
if (-not $mp3Folder) {
    $mp3Folder = "Z:\Multimedia\Music" # Default value
    Set-ItemProperty -Path $registryKey -Name Mp3Folder -Value $mp3Folder
}

$ffmpegPath = (Get-ItemProperty -Path $registryKey -Name FfmpegPath -ErrorAction SilentlyContinue).FfmpegPath
if (-not $ffmpegPath) {
    $ffmpegPath = "C:\Progra~2\Lame\ffmpeg.exe" # Default value
    Set-ItemProperty -Path $registryKey -Name FfmpegPath -Value $ffmpegPath
}

$ffprobePath = (Get-ItemProperty -Path $registryKey -Name FfprobePath -ErrorAction SilentlyContinue).FfprobePath
if (-not $ffprobePath) {
    $ffprobePath = "C:\Progra~2\Lame\ffprobe.exe" # Default value
    Set-ItemProperty -Path $registryKey -Name FfprobePath -Value $ffprobePath
}

# Load Plex Token from Registry
$plexToken = (Get-ItemProperty -Path $registryKey -Name PlexToken -ErrorAction SilentlyContinue).PlexToken
if (-not $plexToken) {
    $plexToken = "ug75pCt7VLXxHxDHbKQz" # Default value
    Set-ItemProperty -Path $registryKey -Name PlexToken -Value $plexToken
}

# Check Minimum Bitrate
$minBitrate = (Get-ItemProperty -Path $registryKey -Name MinBitrate -ErrorAction SilentlyContinue).MinBitrate
if (-not $minBitrate) {
    $minBitrate = 320 # Default value
    Set-ItemProperty -Path $registryKey -Name MinBitrate -Value $minBitrate
    Write-Host "Minimum bitrate set to default: $minBitrate kbps" -ForegroundColor Green
}

# Check Advanced Renamer Path
$advancedRenamerPath = (Get-ItemProperty -Path $registryKey -Name AdvancedRenamerPath -ErrorAction SilentlyContinue).AdvancedRenamerPath
if (-not $advancedRenamerPath) {
    Write-Host "Advanced Renamer path is not set. Please set it in the Options Menu."
}

# Check MP3Tag Path
$mp3TagPath = (Get-ItemProperty -Path $registryKey -Name Mp3TagPath -ErrorAction SilentlyContinue).Mp3TagPath
if (-not $mp3TagPath) {
    Write-Host "MP3Tag path is not set. Please set it in the Options Menu."
}

# Check MP3 File Count
$mp3FileCount = (Get-ItemProperty -Path $registryKey -Name Mp3FileCount -ErrorAction SilentlyContinue).Mp3FileCount
if (-not $mp3FileCount) {
    $mp3FileCount = 0 # Default value
    Set-ItemProperty -Path $registryKey -Name Mp3FileCount -Value $mp3FileCount
    Write-Host "MP3 file count initialized to: $mp3FileCount" -ForegroundColor Green
}

# Menu Function
function Show-Menu {
    while ($true) {
        Clear-Host
        Write-Host "FLAC2MP3 Utility"
        Write-Host "================"
        Write-Host "1.  Start Monitoring FLAC Folder"
        Write-Host "2.  Copy and Organize MP3 Files"
        Write-Host "3.  Delete Empty Folders in FLAC Folder"
        Write-Host "4.  Display Folders in FLAC Directory"
        Write-Host "5.  Options Menu"
        Write-Host "6.  Refresh Plex Library"
        Write-Host "7.  Scan for Low-Bitrate Files"
        Write-Host "8.  Scan and Delete Low-Bitrate Files"
        Write-Host "9.  Open Other Applications(BROKEN)"
        Write-Host "10. Scan Folder and Generate HTML Report"
        Write-Host "11. Scan and Fix Invalid Characters in ID3 Tags"
        Write-Host "12. Split FLAC by CUE"
        Write-Host "13. Manually Convert FLAC Files to MP3"
        Write-Host "99. Exit"
        Write-Host ""
        $option = Read-Host "Select an option"
        switch ($option) {
            '1'  { Start-Monitoring }
            '2'  { Copy-And-Organize-MP3 }
            '3'  { Delete-EmptyFolders }
            '4'  { Display-Folders }
            '5'  { Show-OptionsMenu }
            '6'  { Refresh-PlexLibrary }
            '7'  { Scan-Bitrate }
            '8'  { Scan-And-Delete-LowBitrate }
            '9'  { Show-ApplicationsMenu }
            '10' { Scan-And-Generate-HTMLReport }
            '11' { Scan-And-Fix-ID3Tags }
            '12' { Split-FLAC-By-CUE }
            '13' { Manual-FLAC-Process }
            '99' { Write-Host "Exiting..."; return }
            default {
                Write-Host "Invalid option. Please try again."
                Start-Sleep -Seconds 1
            }
        }
    }
}

# Options Menu Function
function Show-OptionsMenu {
    do {
        Clear-Host
        Write-Host "=========================================" -ForegroundColor Cyan
        Write-Host "               Options Menu              " -ForegroundColor Cyan
        Write-Host "=========================================" -ForegroundColor Cyan
        Write-Host "1. Set FLAC Folder Location"
        Write-Host "2. Set MP3 Folder Location"
        Write-Host "3. Set ffmpeg Location"
        Write-Host "4. Set ffprobe Location"
        Write-Host "5. Set Plex Token"
        Write-Host "6. Set Minimum Bitrate"
        Write-Host "7. Set Advanced Renamer Path"
        Write-Host "8. Set MP3Tag Path"
        Write-Host "9. Update MP3 File Count"
        Write-Host "10. Back to Main Menu"
        Write-Host "========================================="

        $optionChoice = Read-Host "Enter your choice (1-10)"
        switch ($optionChoice) {
            "1" { Set-FLACFolder }
            "2" { Set-MP3Folder }
            "3" { Set-FFmpegPath }
            "4" { Set-FFprobePath }
            "5" { Set-PlexToken }
            "6" { Set-MinBitrate }
            "7" { Set-AdvancedRenamerPath }
            "8" { Set-MP3TagPath }
            "9" { Update-MP3FileCount }
            "10" { break }
            default { Write-Host "Invalid choice. Please select a valid option (1-10)." }
        }
    } while ($optionChoice -ne "10")
}

# Function to Copy and Organize MP3 Files (Integrated mover.ps1 Logic)
function Copy-And-Organize-MP3 {
    Write-Host "Starting MP3 organization..."
    $sourceFolder = $flacFolder
    $destinationBase = $mp3Folder

    # Ensure the destination base folder exists
    if (-not (Test-Path -Path $destinationBase)) {
        New-Item -ItemType Directory -Path $destinationBase -Force | Out-Null
    }

    # Define invalid characters for Windows filenames, including Unicode punctuation, brackets, and parentheses
    $invalidCharsPattern = '[\\/:*?"<>|\[\]\(\)\u2018\u2019\u201C\u201D\u2013\u2014]'

    # Get all .mp3 files in the source folder and its subfolders
    $mp3Files = Get-ChildItem -Path $sourceFolder -Recurse -Filter *.mp3

    foreach ($mp3File in $mp3Files) {
        try {
            $originalName = $mp3File.Name
            $fixedName = $originalName -replace $invalidCharsPattern, ''
            $fixedName = $fixedName -replace '\s+', ' ' # Replace multiple spaces with a single space
            $fixedName = $fixedName.Trim() # Remove leading/trailing spaces
            if ($originalName -ne $fixedName) {
                $fixedPath = Join-Path -Path $mp3File.DirectoryName -ChildPath $fixedName
                if (-not (Test-Path $fixedPath)) {
                    Rename-Item -Path $mp3File.FullName -NewName $fixedName
                    Write-Host "Renamed file: $originalName -> $fixedName" -ForegroundColor Yellow
                    $mp3File = Get-Item $fixedPath
                } else {
                    Write-Warning "Cannot rename $originalName to $fixedName because the target file already exists."
                }
            }

            $quotedPath = '"' + $mp3File.FullName + '"'
            # Use the path directly, do not add extra quotes
$artist = (& $ffprobePath -v error -show_entries format_tags=artist -of default=noprint_wrappers=1:nokey=1 $mp3File.FullName).Trim()
$album  = (& $ffprobePath -v error -show_entries format_tags=album  -of default=noprint_wrappers=1:nokey=1 $mp3File.FullName).Trim()
$bitrate = & $ffprobePath -v error -select_streams a:0 -show_entries stream=bit_rate -of default=noprint_wrappers=1:nokey=1 $mp3File.FullName            
if ($bitrate -and $bitrate -match '^\d+$') {
                $bitrateKbps = [math]::Round($bitrate / 1000)
            } else {
                $bitrateKbps = 0
            }

            # Check if bitrate is less than the minimum bitrate
            if ($bitrateKbps -lt $minBitrate) {
                Write-Warning "The MP3 file '$($mp3File.FullName)' has a bitrate of $bitrateKbps kbps, which is less than $minBitrate kbps."
                $confirmMove = Read-Host "Do you want to move this file anyway? (y/n)"
                if ($confirmMove -ne 'y') {
                    Write-Host "Skipping move for file: $($mp3File.FullName)"
                    continue
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($artist) -and -not [string]::IsNullOrWhiteSpace($album)) {
                # Create the destination folder structure (Artist\Album)
                $destinationFolder = [System.IO.Path]::Combine($destinationBase, $artist, $album)
                if (-not (Test-Path -Path $destinationFolder)) {
                    New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
                }

                # Move the .mp3 file to the destination folder
                $destinationPath = [System.IO.Path]::Combine($destinationFolder, $mp3File.Name)
                Move-Item -Path $mp3File.FullName -Destination $destinationPath -Force
                Write-Host "Moved: $($mp3File.FullName) -> $destinationPath"
            } else {
                Write-Host "Metadata missing for: $($mp3File.FullName). Skipping..."
            }
        } catch {
            Write-Host "Error processing file: $($mp3File.FullName). Error: $_"
        }
    }
    Write-Host "MP3 organization complete."
    Read-Host "Press Enter to return to the menu"
}

# Add-Type -AssemblyName System.Windows.Forms

# Function to Set FLAC Folder Location
function Set-FLACFolder {
    Write-Host "Current FLAC folder location: $flacFolder"

    # Open a folder browser dialog
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select the FLAC folder location"
    $folderBrowser.ShowNewFolderButton = $true

    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $newFlacFolder = $folderBrowser.SelectedPath
        if (Test-Path -Path $newFlacFolder) {
            $global:flacFolder = $newFlacFolder
            Set-ItemProperty -Path $registryKey -Name FlacFolder -Value $flacFolder
            Write-Host "FLAC folder location updated to: $flacFolder" -ForegroundColor Green
        } else {
            Write-Warning "The selected folder does not exist. Please try again."
        }
    } else {
        Write-Host "No folder selected. Operation canceled." -ForegroundColor Yellow
    }

    Read-Host "Press Enter to return to the menu"
}

# Function to Set MP3 Folder Location
function Set-MP3Folder {
    Write-Host "Current MP3 folder location: $mp3Folder"

    # Open a folder browser dialog
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select the MP3 folder location"
    $folderBrowser.ShowNewFolderButton = $true

    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $newMp3Folder = $folderBrowser.SelectedPath
        if (Test-Path -Path $newMp3Folder) {
            $global:mp3Folder = $newMp3Folder
            Set-ItemProperty -Path $registryKey -Name Mp3Folder -Value $mp3Folder
            Write-Host "MP3 folder location updated to: $mp3Folder" -ForegroundColor Green
        } else {
            Write-Warning "The selected folder does not exist. Please try again."
        }
    } else {
        Write-Host "No folder selected. Operation canceled." -ForegroundColor Yellow
    }

    Read-Host "Press Enter to return to the menu"
}

# Function to Set ffmpeg Path
function Set-FFmpegPath {
    Write-Host "Current ffmpeg location: $ffmpegPath"

    # Open a file dialog to select the ffmpeg executable
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Title = "Select the ffmpeg executable"
    $fileDialog.Filter = "Executable Files (*.exe)|*.exe"
    $fileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($ffmpegPath)

    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $newFfmpegPath = $fileDialog.FileName
        if (Test-Path -Path $newFfmpegPath) {
            $global:ffmpegPath = $newFfmpegPath
            Set-ItemProperty -Path $registryKey -Name FfmpegPath -Value $ffmpegPath
            Write-Host "ffmpeg location updated to: $ffmpegPath" -ForegroundColor Green
        } else {
            Write-Warning "The selected file does not exist. Please try again."
        }
    } else {
        Write-Host "No file selected. Operation canceled." -ForegroundColor Yellow
    }

    Read-Host "Press Enter to return to the menu"
}

# Function to Set ffprobe Path
function Set-FFprobePath {
    Write-Host "Current ffprobe location: $ffprobePath"

    # Open a file dialog to select the ffprobe executable
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Title = "Select the ffprobe executable"
    $fileDialog.Filter = "Executable Files (*.exe)|*.exe"
    $fileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($ffprobePath)

    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $newFfprobePath = $fileDialog.FileName
        if (Test-Path -Path $newFfprobePath) {
            $global:ffprobePath = $newFfprobePath
            Set-ItemProperty -Path $registryKey -Name FfprobePath -Value $ffprobePath
            Write-Host "ffprobe location updated to: $ffprobePath" -ForegroundColor Green
        } else {
            Write-Warning "The selected file does not exist. Please try again."
        }
    } else {
        Write-Host "No file selected. Operation canceled." -ForegroundColor Yellow
    }

    Read-Host "Press Enter to return to the menu"
}

# Function to Set Plex Token
function Set-PlexToken {
    Write-Host "Current Plex Token: $plexToken"
    do {
        $newPlexToken = Read-Host "Enter the new Plex Token (or press Enter to cancel)"
        if (-not $newPlexToken) {
            Write-Host "No input provided. Operation canceled." -ForegroundColor Yellow
            return
        }
        $global:plexToken = $newPlexToken
        Set-ItemProperty -Path $registryKey -Name PlexToken -Value $plexToken
        Write-Host "Plex Token updated to: $plexToken" -ForegroundColor Green
        break
    } while ($true)
    Read-Host "Press Enter to return to the menu"
}

# Function to Set Minimum Bitrate
function Set-MinBitrate {
    Write-Host "Current minimum bitrate: $minBitrate kbps"
    do {
        $newMinBitrate = Read-Host "Enter the new minimum bitrate in kbps (e.g., 320) (or press Enter to cancel)"
        if (-not $newMinBitrate) {
            Write-Host "No input provided. Operation canceled." -ForegroundColor Yellow
            return
        }
        if ($newMinBitrate -match '^\d+$' -and [int]$newMinBitrate -gt 0) {
            $global:minBitrate = [int]$newMinBitrate
            Set-ItemProperty -Path $registryKey -Name MinBitrate -Value $minBitrate
            Write-Host "Minimum bitrate updated to: $minBitrate kbps" -ForegroundColor Green
            break
        } else {
            Write-Warning "Invalid input. Please enter a positive integer."
        }
    } while ($true)
    Read-Host "Press Enter to return to the menu"
}

# Function to Set Advanced Renamer Path
function Set-AdvancedRenamerPath {
    Write-Host "Current Advanced Renamer path: $advancedRenamerPath"

    # Open a file dialog to select the Advanced Renamer executable
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Title = "Select the Advanced Renamer executable"
    $fileDialog.Filter = "Executable Files (*.exe)|*.exe"

    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $newAdvancedRenamerPath = $fileDialog.FileName
        if (Test-Path -Path $newAdvancedRenamerPath) {
            $global:advancedRenamerPath = $newAdvancedRenamerPath
            Set-ItemProperty -Path $registryKey -Name AdvancedRenamerPath -Value $advancedRenamerPath
            Write-Host "Advanced Renamer path updated to: $advancedRenamerPath" -ForegroundColor Green
        } else {
            Write-Warning "The selected file does not exist. Please try again."
        }
    } else {
        Write-Host "No file selected. Operation canceled." -ForegroundColor Yellow
    }

    Read-Host "Press Enter to return to the menu"
}

# Function to Set MP3Tag Path
function Set-MP3TagPath {
    Write-Host "Current MP3Tag path: $mp3TagPath"

    # Open a file dialog to select the MP3Tag executable
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Title = "Select the MP3Tag executable"
    $fileDialog.Filter = "Executable Files (*.exe)|*.exe"

    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $newMp3TagPath = $fileDialog.FileName
        if (Test-Path -Path $newMp3TagPath) {
            $global:mp3TagPath = $newMp3TagPath
            Set-ItemProperty -Path $registryKey -Name Mp3TagPath -Value $mp3TagPath
            Write-Host "MP3Tag path updated to: $mp3TagPath" -ForegroundColor Green
        } else {
            Write-Warning "The selected file does not exist. Please try again."
        }
    } else {
        Write-Host "No file selected. Operation canceled." -ForegroundColor Yellow
    }

    Read-Host "Press Enter to return to the menu"
}

# Function to Display Folders in the FLAC Directory
function Display-Folders {
    Write-Host "Displaying all folders in the FLAC directory..."
    $folders = Get-ChildItem -Path $flacFolder -Directory -Recurse

    if ($folders.Count -eq 0) {
        Write-Host "No folders found in the FLAC directory."
    } else {
        foreach ($folder in $folders) {
            Write-Host $folder.FullName
        }
    }
    Write-Host "Folder display complete."
    Read-Host "Press Enter to return to the menu"
}

# Function to Delete Empty Folders
function Delete-EmptyFolders {
    Write-Host "Scanning for empty folders in the FLAC folder..."
    $folders = Get-ChildItem -Path $flacFolder -Directory -Recurse

    foreach ($folder in $folders) {
        $hasFiles = Get-ChildItem -Path $folder.FullName -Recurse -Include *.mp3, *.flac -ErrorAction SilentlyContinue
        if (-not $hasFiles) {
            try {
                Remove-Item -Path $folder.FullName -Recurse -Force
                Write-Host "Deleted empty folder: $($folder.FullName)"
            } catch {
                Write-Warning "Failed to delete folder: $($folder.FullName). Error: $_"
            }
        }
    }
    Write-Host "Empty folder cleanup complete."
}

# Monitoring Function
function Start-Monitoring {
    Write-Host "Monitoring $flacFolder and subfolders for new FLAC and WMA files. Press 'x' to exit monitoring." -ForegroundColor Cyan

    $fsw = New-Object System.IO.FileSystemWatcher
    $fsw.Path = $flacFolder
    $fsw.IncludeSubdirectories = $true
    $fsw.Filter = "*.*"
    $fsw.EnableRaisingEvents = $true

    $onCreated = Register-ObjectEvent $fsw Created -Action {
        $regexPattern = "[\[\]\\/:*?""<>|{}\(\)']"

        # Registry settings
        $registryKey = "HKCU:\Software\FLAC2MP3"
        $minBitrate = (Get-ItemProperty -Path $registryKey -Name MinBitrate -ErrorAction SilentlyContinue).MinBitrate
        if (-not $minBitrate) { $minBitrate = 192 }
        $ffmpegPath = (Get-ItemProperty -Path $registryKey -Name FfmpegPath -ErrorAction SilentlyContinue).FfmpegPath
        if (-not $ffmpegPath) { $ffmpegPath = "ffmpeg" }
        $ffprobePath = (Get-ItemProperty -Path $registryKey -Name FfprobePath -ErrorAction SilentlyContinue).FfprobePath
        if (-not $ffprobePath) { $ffprobePath = "ffprobe" }
        $mp3Folder = (Get-ItemProperty -Path $registryKey -Name Mp3Folder -ErrorAction SilentlyContinue).Mp3Folder
        if (-not $mp3Folder) { $mp3Folder = $env:USERPROFILE }

        $fullPath = $Event.SourceEventArgs.FullPath
        $ext = [System.IO.Path]::GetExtension($fullPath)
        if ($ext -ieq ".flac" -or $ext -ieq ".wma") {
            $fileName = [System.IO.Path]::GetFileName($fullPath)
            $dirName = [System.IO.Path]::GetDirectoryName($fullPath)
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
            $extension = [System.IO.Path]::GetExtension($fileName)

            # Rename if needed
            if ($baseName -match $regexPattern) {
                $safeBaseName = $baseName -replace $regexPattern, "_"
                $safeName = "${safeBaseName}${extension}"
                $safePath = Join-Path $dirName $safeName
                try {
                    Rename-Item -Path $fullPath -NewName $safeName -ErrorAction Stop
                    Write-Host "Renamed file: ${fileName} -> ${safeName}" -ForegroundColor Yellow
                    Write-Host "New file detected: $safePath" -ForegroundColor Green
                    $fullPath = $safePath
                    $fileName = $safeName
                    $baseName = $safeBaseName
                } catch {
                    Write-Host "Failed to rename ${fileName}: ${_}" -ForegroundColor Red
                    return
                }
            } else {
                Write-Host "New file detected: $fullPath" -ForegroundColor Green
            }

            # Convert to MP3
            $mp3Path = [System.IO.Path]::Combine($dirName, "$baseName.mp3")
            $ffmpegArgs = @(
                "-y"
                "-i"; "$fullPath"
                "-b:a"; "${minBitrate}k"
                "$mp3Path"
            )
            Write-Host "Converting to MP3: $mp3Path (min bitrate: $minBitrate kbps)" -ForegroundColor Cyan
            $ffmpegResult = & $ffmpegPath @ffmpegArgs 2>&1

            if (Test-Path $mp3Path) {
                Write-Host "Conversion successful. Deleting original file: $fullPath" -ForegroundColor Green
                Remove-Item $fullPath -Force

                # Extract artist and album using ffprobe (ID3v2: TPE1=artist, TALB=album)
                $ffprobeArgs = @(
                    "-v", "error"
                    "-show_entries", "format_tags=TPE1,TALB"
                    "-of", "default=noprint_wrappers=1:nokey=1"
                    "$mp3Path"
                )
                $ffprobeOutput = & $ffprobePath @ffprobeArgs 2>&1

                # Debug: Show all ffprobe tag output
                Write-Host "ffprobe tag output for ${mp3Path}:" -ForegroundColor Magenta
                $ffprobeOutput | ForEach-Object { Write-Host $_ -ForegroundColor Magenta }

                $artist = ""
                $album = ""
                if ($ffprobeOutput.Count -ge 1) { $artist = $ffprobeOutput[0].Trim() }
                if ($ffprobeOutput.Count -ge 2) { $album = $ffprobeOutput[1].Trim() }

                # If not found, try artist and album
                if (-not $artist -or $artist -eq "") {
                    $ffprobeArgs = @(
                        "-v", "error"
                        "-show_entries", "format_tags=artist"
                        "-of", "default=noprint_wrappers=1:nokey=1"
                        "$mp3Path"
                    )
                    $artist = (& $ffprobePath @ffprobeArgs 2>&1 | Select-Object -First 1).Trim()
                }
                if (-not $album -or $album -eq "") {
                    $ffprobeArgs = @(
                        "-v", "error"
                        "-show_entries", "format_tags=album"
                        "-of", "default=noprint_wrappers=1:nokey=1"
                        "$mp3Path"
                    )
                    $album = (& $ffprobePath @ffprobeArgs 2>&1 | Select-Object -First 1).Trim()
                }

                if (-not $artist) { $artist = "Unknown Artist" }
                if (-not $album) { $album = "Unknown Album" }

                # Clean up artist/album folder names
                $artistSafe = $artist -replace $regexPattern, "_"
                $albumSafe = $album -replace $regexPattern, "_"

                $destDir = Join-Path $mp3Folder $artistSafe
                $destDir = Join-Path $destDir $albumSafe
                if (-not (Test-Path $destDir)) { New-Item -Path $destDir -ItemType Directory | Out-Null }

                $destPath = Join-Path $destDir ([System.IO.Path]::GetFileName($mp3Path))
                try {
                    Move-Item -Path $mp3Path -Destination $destPath -Force
                    Write-Host "Moved MP3 to: $destPath" -ForegroundColor Green
                } catch {
                    Write-Host "Failed to move MP3: $_" -ForegroundColor Red
                }
            } else {
                Write-Host "Conversion failed for $fullPath" -ForegroundColor Red
                Write-Host $ffmpegResult
            }
        }
    }

    while ($true) {
        if ([System.Console]::KeyAvailable) {
            $key = [System.Console]::ReadKey($true)
            if ($key.KeyChar -eq 'x' -or $key.KeyChar -eq 'X') {
                break
            }
        }
        Start-Sleep -Milliseconds 200
    }

    Unregister-Event -SourceIdentifier $onCreated.Name
    $fsw.Dispose()
    Write-Host "Stopped monitoring. Returning to main menu." -ForegroundColor Yellow
}

# Function to Refresh Plex Library
function Refresh-PlexLibrary {
    $plexUrl = "http://192.168.1.17:32400/library/sections/6/refresh?X-Plex-Token=$plexToken"
    Write-Host "Refreshing Plex Library..."
    try {
        Invoke-WebRequest -Uri $plexUrl -Method Get -UseBasicParsing | Out-Null
        Write-Host "Plex Library refresh request sent successfully." -ForegroundColor Green
    } catch {
        Write-Warning "Failed to refresh Plex Library. Error: $_"
    }
    Read-Host "Press Enter to return to the menu"
}

# Function to Scan a Directory for Low-Bitrate Files
function Scan-Bitrate {
    Write-Host "Select a folder to scan for low-bitrate files..."

    # Open a folder browser dialog
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select the folder to scan for low-bitrate files"
    $folderBrowser.ShowNewFolderButton = $false

    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedFolder = $folderBrowser.SelectedPath
        Write-Host "Scanning folder: $selectedFolder" -ForegroundColor Cyan

        # Get all MP3 files in the selected folder and its subfolders
        $mp3Files = Get-ChildItem -Path $selectedFolder -Recurse -Filter *.mp3
        if ($mp3Files.Count -eq 0) {
            Write-Host "No MP3 files found in the selected folder." -ForegroundColor Yellow
            return
        }

        # Initialize a list for low-bitrate files
        $lowBitrateFiles = @()

        foreach ($mp3File in $mp3Files) {
            try {
                # Extract bitrate using ffprobe
                $bitrate = & $ffprobePath -v error -select_streams a:0 -show_entries stream=bit_rate -of default=noprint_wrappers=1:nokey=1 "`"$($mp3File.FullName)`""
                $bitrateKbps = [math]::Round($bitrate / 1000)

                # Check if bitrate is less than the minimum bitrate
                if ($bitrateKbps -lt $minBitrate) {
                    $lowBitrateFiles += $mp3File.FullName
                }
            } catch {
                Write-Warning "Error processing file: $($mp3File.FullName). Error: $_"
            }
        }

        # Display results
        if ($lowBitrateFiles.Count -gt 0) {
            Write-Host "The following files have a bitrate less than $minBitrate kbps:" -ForegroundColor Yellow
            $lowBitrateFiles | ForEach-Object { Write-Host $_ }
        } else {
            Write-Host "All files meet the minimum bitrate standard of $minBitrate kbps." -ForegroundColor Green
        }
    } else {
        Write-Host "No folder selected. Operation canceled." -ForegroundColor Yellow
    }

    Read-Host "Press Enter to return to the menu"
}

# Function to Scan and Delete Low-Bitrate Files
function Scan-And-Delete-LowBitrate {
    Write-Host "Select a folder to scan for low-bitrate files..."

    # Open a folder browser dialog
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select the folder to scan for low-bitrate files"
    $folderBrowser.ShowNewFolderButton = $false

    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedFolder = $folderBrowser.SelectedPath
        Write-Host "Scanning folder: $selectedFolder" -ForegroundColor Cyan

        # Get all MP3 files in the selected folder and its subfolders
        $mp3Files = Get-ChildItem -Path $selectedFolder -Recurse -Filter *.mp3
        if ($mp3Files.Count -eq 0) {
            Write-Host "No MP3 files found in the selected folder." -ForegroundColor Yellow
            return
        }

        # Initialize a list for low-bitrate files
        $lowBitrateFiles = @()

        foreach ($mp3File in $mp3Files) {
            try {
                # Extract bitrate using ffprobe
                $bitrate = & $ffprobePath -v error -select_streams a:0 -show_entries stream=bit_rate -of default=noprint_wrappers=1:nokey=1 "`"$($mp3File.FullName)`""
                $bitrateKbps = [math]::Round($bitrate / 1000)

                # Check if bitrate is less than the minimum bitrate
                if ($bitrateKbps -lt $minBitrate) {
                    $lowBitrateFiles += $mp3File.FullName
                }
            } catch {
                Write-Warning "Error processing file: $($mp3File.FullName). Error: $_"
            }
        }

        # Display results
        if ($lowBitrateFiles.Count -gt 0) {
            Write-Host "The following files have a bitrate less than $minBitrate kbps:" -ForegroundColor Yellow
            $lowBitrateFiles | ForEach-Object { Write-Host $_ }

            # Ask the user if they want to delete the files
            $confirmDelete = Read-Host "Do you want to delete these files? (y/n)"
            if ($confirmDelete -eq 'y') {
                foreach ($file in $lowBitrateFiles) {
                    try {
                        Remove-Item -Path $file -Force
                        Write-Host "Deleted: $file" -ForegroundColor Green
                    } catch {
                        Write-Warning "Failed to delete file: $file. Error: $_"
                    }
                }
                Write-Host "Low-bitrate files deleted successfully." -ForegroundColor Green
            } else {
                Write-Host "No files were deleted." -ForegroundColor Yellow
            }
        } else {
            Write-Host "All files meet the minimum bitrate standard of $minBitrate kbps." -ForegroundColor Green
        }
    } else {
        Write-Host "No folder selected. Operation canceled." -ForegroundColor Yellow
    }

    Read-Host "Press Enter to return to the menu"
}

# Function to Scan a Folder and Generate an HTML Report
function Scan-And-Generate-HTMLReport {
    Write-Host "Select a folder to scan for metadata..."

    # Open a folder browser dialog
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select the folder to scan for metadata"
    $folderBrowser.ShowNewFolderButton = $false

    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedFolder = $folderBrowser.SelectedPath
        Write-Host "Scanning folder: $selectedFolder" -ForegroundColor Cyan

        # Get the folder name to use as the HTML file name
        $folderName = [System.IO.Path]::GetFileName($selectedFolder)

        # Get all MP3 files in the selected folder and its subfolders
        $mp3Files = Get-ChildItem -Path $selectedFolder -Recurse -Filter *.mp3
        if ($mp3Files.Count -eq 0) {
            Write-Host "No MP3 files found in the selected folder." -ForegroundColor Yellow
            return
        }

        # Initialize a list for metadata
        $metadataList = @()

        foreach ($mp3File in $mp3Files) {
            try {
                # Extract Artist, Album, and Bitrate metadata using ffprobe
                $quotedPath = '"' + $mp3File.FullName + '"'
                $artist = & $ffprobePath -v error -show_entries format_tags=artist -of default=noprint_wrappers=1:nokey=1 $mp3File.FullName
                $album  = & $ffprobePath -v error -show_entries format_tags=album  -of default=noprint_wrappers=1:nokey=1 $mp3File.FullName
                $bitrate = & $ffprobePath -v error -select_streams a:0 -show_entries stream=bit_rate -of default=noprint_wrappers=1:nokey=1 $mp3File.FullName
                if ($bitrate -and $bitrate -match '^\d+$') {
                    $bitrateKbps = [math]::Round($bitrate / 1000)
                } else {
                    $bitrateKbps = 0
                }

                # Add metadata to the list
                $metadataList += [PSCustomObject]@{
                    Filename = $mp3File.Name
                    Artist   = $artist
                    Album    = $album
                    Bitrate  = "$bitrateKbps kbps"
                }
            } catch {
                Write-Warning "Error processing file: $($mp3File.FullName). Error: $_"
            }
        }

        # Generate the HTML file name based on the folder name
        $outputFile = Join-Path -Path $PSScriptRoot -ChildPath "$folderName-MetadataReport.html"

        # Generate HTML report using PSWriteHTML
        New-HTML -TitleText 'Metadata Report' -Online:$false -FilePath $outputFile -ShowHTML {
            New-HTMLSection -HeaderText "MP3 Metadata Report for '$folderName'" {
                New-HTMLTable -DataTable $metadataList -HideFooter -ScrollCollapse
            }
        }

        Write-Host "HTML report generated successfully: $outputFile" -ForegroundColor Green
    } else {
        Write-Host "No folder selected. Operation canceled." -ForegroundColor Yellow
    }

    Read-Host "Press Enter to return to the menu"
}

# Submenu to Open Other Applications
function Show-ApplicationsMenu {
    do {
        Clear-Host
        Write-Host "=========================================" -ForegroundColor Cyan
        Write-Host "           Open Other Applications       " -ForegroundColor Cyan
        Write-Host "=========================================" -ForegroundColor Cyan
        Write-Host "1. Open Advanced Renamer"
        Write-Host "2. Open MP3Tag"
        Write-Host "3. Back to Main Menu"
        Write-Host "========================================="

        $appChoice = Read-Host "Enter your choice (1-3)"
        switch ($appChoice) {
            "1" {
                if (Test-Path -Path $advancedRenamerPath) {
                    Start-Process -FilePath $advancedRenamerPath
                } else {
                    Write-Warning "Advanced Renamer path is not set. Please set it in the Options Menu."
                }
            }
            "2" {
                if (Test-Path -Path $mp3TagPath) {
                    Start-Process -FilePath $mp3TagPath
                } else {
                    Write-Warning "MP3Tag path is not set. Please set it in the Options Menu."
                }
            }
            "3" { break }
            default { Write-Host "Invalid choice. Please select a valid option (1-3)." }
        }
    } while ($appChoice -ne "3")
}

# Function to Update MP3 File Count
function Update-MP3FileCount {
    Write-Host "Updating MP3 file count..."
    if (Test-Path -Path $mp3Folder) {
        $mp3FileCount = (Get-ChildItem -Path $mp3Folder -Recurse -Filter *.mp3 -ErrorAction SilentlyContinue).Count
        Set-ItemProperty -Path $registryKey -Name Mp3FileCount -Value $mp3FileCount
        Write-Host "MP3 file count updated to: $mp3FileCount" -ForegroundColor Green
    } else {
        Write-Warning "MP3 folder does not exist. Please check the folder path."
    }
    Read-Host "Press Enter to return to the menu"
}

# Add this function to manually convert FLAC files in a selected folder to MP3
function Manual-FLAC-Process {
    Write-Host "Select a folder to manually convert FLAC files to MP3..." -ForegroundColor Cyan

    # Open a folder browser dialog
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select the folder containing FLAC files"
    $folderBrowser.ShowNewFolderButton = $false

    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedFolder = $folderBrowser.SelectedPath
        Write-Host "Converting FLAC files in: $selectedFolder" -ForegroundColor Cyan

        $flacFiles = Get-ChildItem -Path $selectedFolder -Filter *.flac
        if ($flacFiles.Count -eq 0) {
            Write-Host "No FLAC files found in the selected folder." -ForegroundColor Yellow
            return
        }

        foreach ($flacFile in $flacFiles) {
            try {
                $mp3File = [System.IO.Path]::ChangeExtension($flacFile.FullName, ".mp3")
                $lameArgs = "-i `"$($flacFile.FullName)`" -ab ${minBitrate}k -map_metadata 0 -id3v2_version 3 `"$mp3File`""
                Write-Host "Converting: $($flacFile.Name) -> $([System.IO.Path]::GetFileName($mp3File))"
                $process = Start-Process -FilePath $ffmpegPath -ArgumentList $lameArgs -NoNewWindow -Wait -PassThru
                if ($process.ExitCode -eq 0) {
                    Write-Host "Converted: $($flacFile.Name) -> $([System.IO.Path]::GetFileName($mp3File))" -ForegroundColor Green
                } else {
                    Write-Warning "Failed to convert: $($flacFile.Name)"
                }
            } catch {
                Write-Warning "Error converting file: $($flacFile.FullName). Error: $_"
            }
        }
        Write-Host "Manual FLAC to MP3 conversion complete." -ForegroundColor Green
    } else {
        Write-Host "No folder selected. Operation canceled." -ForegroundColor Yellow
    }
    Read-Host "Press Enter to return to the menu"
}

# Start the MP3 Folder Monitor Script in the Background
$monitorScriptPath = "C:\Users\sean\OneDrive\Documents\MonitorMP3Folder.ps1"
if (Test-Path -Path $monitorScriptPath) {
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$monitorScriptPath`"" -WindowStyle Hidden
    Write-Host "Started MP3 folder monitor in the background."
} else {
    Write-Warning "Monitor script not found at: $monitorScriptPath"
}

# Function to Scan and Fix Invalid Characters in ID3 Tags
function Scan-And-Fix-ID3Tags {
    Write-Host "Select a folder to scan and fix ID3 tags..." -ForegroundColor Cyan

    # Open a folder browser dialog
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select the folder to scan and fix ID3 tags"
    $folderBrowser.ShowNewFolderButton = $false

    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedFolder = $folderBrowser.SelectedPath
        Write-Host "Scanning folder: $selectedFolder" -ForegroundColor Cyan

        # Get all MP3 and FLAC files in the selected folder and its subfolders
        $audioFiles = Get-ChildItem -Path $selectedFolder -Recurse -Include *.mp3,*.flac
        if ($audioFiles.Count -eq 0) {
            Write-Host "No audio files (MP3 or FLAC) found in the selected folder." -ForegroundColor Yellow
            return
        }

        # Define invalid characters for Windows filenames (excluding spaces)
        #$invalidChars = [regex]::Escape('\ / : * ? " < > |')
        $invalidChars = '\\/:*?"<>|'
        # Function to replace invalid characters with underscores
        function Fix-InvalidCharacters {
            param (
                [string]$value
            )
            return ($value -replace "[$invalidChars]", "_")
        }

        foreach ($audioFile in $audioFiles) {
            try {
                # Extract ID3 tags using ffprobe
                $ffprobeOutput = & $ffprobePath -v quiet -show_entries format_tags=artist,album,title -of json -i "`"$($audioFile.FullName)`""
                if (-not $ffprobeOutput) {
                    Write-Host "Error: Failed to retrieve metadata for $($audioFile.FullName)" -ForegroundColor Yellow
                    continue
                }

                # Parse the JSON output
                $tags = $ffprobeOutput | ConvertFrom-Json
                if (-not $tags.format.tags) {
                    Write-Host "No ID3 tags found for $($audioFile.FullName)" -ForegroundColor Yellow
                    continue
                }

                # Initialize a flag to track if changes are needed
                $changesNeeded = $false
                $newTags = @{}

                # Check and fix each tag
                foreach ($tag in @("artist", "album", "title")) {
                    if ($tags.format.tags.$tag) {
                        $originalValue = $tags.format.tags.$tag
                        $fixedValue = Fix-InvalidCharacters -value $originalValue

                        if ($originalValue -ne $fixedValue) {
                            Write-Host "Fixing invalid characters in ${tag}: $($originalValue) -> $($fixedValue)" -ForegroundColor Yellow
                            $newTags[$tag] = $fixedValue
                            $changesNeeded = $true
                        } else {
                            Write-Host "No invalid characters found in ${tag}: $($originalValue)" -ForegroundColor Green
                        }
                    }
                }

                # Apply changes if needed
                if ($changesNeeded) {
                    # Create a temporary file for the updated FLAC or MP3 file
                    $tempFile = [System.IO.Path]::GetTempFileName()
                    $tempFile = [System.IO.Path]::ChangeExtension($tempFile, [System.IO.Path]::GetExtension($audioFile.FullName))

                    # Build ffmpeg arguments to update the tags
                    $ffmpegArgs = @("-i", "`"$($audioFile.FullName)`"", "-y")
                    foreach ($key in $newTags.Keys) {
                        $ffmpegArgs += @("-metadata", "$key=`"$($newTags[$key])`"")
                    }
                    $ffmpegArgs += @("-codec", "copy", "`"$tempFile`"")

                    # Run ffmpeg to update the tags
                    Write-Host "Updating tags for $($audioFile.FullName)..." -ForegroundColor Cyan
                    & $ffmpegPath @ffmpegArgs

                    if ($LASTEXITCODE -eq 0) {
                        # Replace the original file with the updated file
                        Move-Item -Path $tempFile -Destination $audioFile.FullName -Force
                        Write-Host "Tags updated successfully for $($audioFile.FullName)" -ForegroundColor Green
                    } else {
                        Write-Host "Error: Failed to update tags for $($audioFile.FullName)" -ForegroundColor Red
                        # Clean up the temporary file if the update failed
                        Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
                    }
                } else {
                    Write-Host "No changes needed for $($audioFile.FullName)" -ForegroundColor Green
                }
            } catch {
                Write-Host "Error processing file: $($audioFile.FullName). Error: $_" -ForegroundColor Red
            }
        }

        Write-Host "ID3 tag scan and fix complete." -ForegroundColor Cyan
    } else {
        Write-Host "No folder selected. Operation canceled." -ForegroundColor Yellow
    }

    Read-Host "Press Enter to return to the menu"
}

# Function to Split FLAC Files by CUE
function Split-FLAC-By-CUE {
    Write-Host "Scanning for CUE files in $flacFolder..." -ForegroundColor Cyan
    $cueFiles = Get-ChildItem -Path $flacFolder -Recurse -Filter *.cue

    if ($cueFiles.Count -eq 0) {
        Write-Host "No CUE files found." -ForegroundColor Yellow
        return
    }

    foreach ($cueFile in $cueFiles) {
        try {
            $cueDir = $cueFile.DirectoryName
            $cueContent = Get-Content -LiteralPath $cueFile.FullName -Raw

            # Find the referenced FLAC file
            if ($cueContent -match 'FILE\s+"(.+?\.flac)"') {
                $flacName = $Matches[1].Trim()
                $flacPath = Join-Path $cueDir $flacName
                if (-not (Test-Path -LiteralPath $flacPath)) {
                    Write-Warning "Referenced FLAC file not found for CUE: $($cueFile.FullName)"
                    continue
                }
            } else {
                Write-Warning "No FLAC file reference found in CUE: $($cueFile.FullName)"
                continue
            }

            # Parse TRACK info
            $tracks = @()
            $cueLines = $cueContent -split "`r?`n"
            $currentTrack = $null
            foreach ($line in $cueLines) {
                if ($line -match '^\s*TRACK\s+(\d+)\s+AUDIO') {
                    if ($currentTrack) { $tracks += $currentTrack }
                    $currentTrack = @{ Number = $Matches[1]; Title = ""; Index = "" }
                }
                if ($line -match '^\s*TITLE\s+"(.+)"' -and $currentTrack) {
                    $currentTrack.Title = $Matches[1]
                }
                if ($line -match '^\s*INDEX 01\s+(\d{2}):(\d{2}):(\d{2})' -and $currentTrack) {
                    $cueTime = "$($Matches[1]):$($Matches[2]):$($Matches[3])"
                    $currentTrack.Index = Convert-CueTimeToSeconds $cueTime
                }
            }
            if ($currentTrack) { $tracks += $currentTrack }

            if ($tracks.Count -lt 1) {
                Write-Warning "No tracks found in CUE: $($cueFile.FullName)"
                continue
            }

            # Calculate start and duration for each track
            for ($i = 0; $i -lt $tracks.Count; $i++) {
                $start = $tracks[$i].Index
                if ($i -lt $tracks.Count - 1) {
                    $nextStart = $tracks[$i+1].Index
                    # Convert to TimeSpan for subtraction
                    $startTS = [TimeSpan]::ParseExact($start, "hh\:mm\:ss\.fff", $null)
                    $nextTS = [TimeSpan]::ParseExact($nextStart, "hh\:mm\:ss\.fff", $null)
                    $durationTS = $nextTS - $startTS
                    $duration = "{0:00}:{1:00}:{2:00}.{3:000}" -f $durationTS.Hours, $durationTS.Minutes, $durationTS.Seconds, $durationTS.Milliseconds
                } else {
                    $duration = $null
                }
                $title = $tracks[$i].Title
                $trackNum = [int]$tracks[$i].Number
                $outputFile = "{0:D2} - {1}.flac" -f $trackNum, ($title -replace '[\\/:*?"<>|]', '_')
                $outputPath = Join-Path $cueDir $outputFile

                $ffmpegArgs = @("-hide_banner", "-y", "-i", $flacPath, "-ss", $start)
                if ($duration) { $ffmpegArgs += @("-t", $duration) }
                $ffmpegArgs += @("-c", "copy", $outputPath)

                Write-Host "Extracting track ${trackNum}: ${title}"
                & $ffmpegPath @ffmpegArgs
            }

            # Delete original FLAC and CUE if at least one split file exists
            $splitFiles = Get-ChildItem -Path $cueDir -Filter "*.flac" | Where-Object { $_.FullName -ne $flacPath }
            if ($splitFiles.Count -gt 0) {
                Remove-Item -Path $flacPath -Force
                Remove-Item -Path $cueFile.FullName -Force
                Write-Host "Deleted original FLAC and CUE: $flacName, $($cueFile.Name)" -ForegroundColor Yellow
            } else {
                Write-Warning "No split files found. Original FLAC and CUE not deleted."
            }
        } catch {
            Write-Warning "Error processing CUE file: $($cueFile.FullName). Error: $_"
        }
    }
    Write-Host "CUE splitting complete."
    Read-Host "Press Enter to return to the menu"
}

function Convert-CueTimeToSeconds {
    param($cueTime)
    # cueTime is mm:ss:ff
    if ($cueTime -match '^(\d{2}):(\d{2}):(\d{2})$') {
        $mm = [int]$Matches[1]
        $ss = [int]$Matches[2]
        $ff = [int]$Matches[3]
        $totalSeconds = ($mm * 60) + $ss + ($ff / 75)
        # Return as hh:mm:ss.fff for ffmpeg
        $ts = [TimeSpan]::FromSeconds($totalSeconds)
        return "{0:00}:{1:00}:{2:00}.{3:000}" -f $ts.Hours, $ts.Minutes, $ts.Seconds, $ts.Milliseconds
    }
    return $cueTime
}

# Read config from registry
$flacFolder = (Get-ItemProperty -Path "HKCU:\Software\FLAC2MP3" -Name "FLACFolder").FLACFolder
$mp3Folder  = (Get-ItemProperty -Path "HKCU:\Software\FLAC2MP3" -Name "MP3Folder").MP3Folder
$minBitrate = (Get-ItemProperty -Path "HKCU:\Software\FLAC2MP3" -Name "MinBitRate").MinBitRate

function Convert-And-MoveFile($inputPath) {
    $ext = [System.IO.Path]::GetExtension($inputPath).ToLower()
    if ($ext -eq ".flac" -or $ext -eq ".wma") {
        $tempMp3 = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.mp3'
        Write-Host "Converting $inputPath to $tempMp3 at $minBitrate kbps"
        ffmpeg -y -i "$inputPath" -ab ${minBitrate}k "$tempMp3"

        if (Test-Path $tempMp3) {
            # Read ID3 tags from the new MP3
            try {
                $tag = [TagLib.File]::Create($tempMp3)
                $artist = ($tag.Tag.Performers -join ', ').Trim()
                $album  = $tag.Tag.Album.Trim()
                $title  = $tag.Tag.Title.Trim()
                if (-not $artist) { $artist = "Unknown Artist" }
                if (-not $album)  { $album  = "Unknown Album" }
                if (-not $title)  { $title  = [System.IO.Path]::GetFileNameWithoutExtension($inputPath) }
            } catch {
                $artist = "Unknown Artist"
                $album  = "Unknown Album"
                $title  = [System.IO.Path]::GetFileNameWithoutExtension($inputPath)
            }

            # Build destination path
            $destDir = Join-Path $mp3Folder $artist
            $destDir = Join-Path $destDir $album
            if (-not (Test-Path $destDir)) { New-Item -Path $destDir -ItemType Directory | Out-Null }
            $destMp3 = Join-Path $destDir ($title + ".mp3")

            # Move and clean up
            Move-Item $tempMp3 $destMp3 -Force
            Remove-Item $inputPath
            Write-Host "Moved $destMp3"
        } else {
            Write-Host "Failed to convert $inputPath"
        }
    }
}

# Load TagLib# for ID3 reading (requires TagLibSharp.dll in script folder or system)
Add-Type -Path "$PSScriptRoot\TagLibSharp.dll"

# Read FLAC folder path from registry
$flacFolder = (Get-ItemProperty -Path "HKCU:\Software\FLAC2MP3" -Name "FLACFolder").FLACFolder

if ($option -eq 1) {
    Write-Host "Monitoring $flacFolder and subfolders for new FLAC and WMA files..."
    Write-Host "Press 'x' then Enter to stop monitoring."

    # Set up FileSystemWatcher for new files
    $watcher = New-Object System.IO.FileSystemWatcher
    $watcher.Path = $flacFolder
    $watcher.IncludeSubdirectories = $true
    $watcher.Filter = "*.*"
    $watcher.EnableRaisingEvents = $true

    $event = Register-ObjectEvent $watcher Created -Action {
        $fullPath = $Event.SourceEventArgs.FullPath
        $ext = [System.IO.Path]::GetExtension($fullPath).ToLower()
        if ($ext -eq ".flac" -or $ext -eq ".wma") {
            Write-Host $fullPath
        }
    }

    # Wait for 'x' key to exit
    while ($true) {
        $input = Read-Host
        if ($input -eq 'x') {
            Unregister-Event -SourceIdentifier $event.Name
            $watcher.EnableRaisingEvents = $false
            $watcher.Dispose()
            Write-Host "Monitoring stopped."
            break
        }
    }
}

Show-Menu

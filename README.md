
# FLAC2MP3 Utility

## Overview

**FLAC2MP3.ps1** is a comprehensive PowerShell utility designed to help users manage, convert, and organize their music library, specifically focusing on converting FLAC files to MP3, organizing MP3 files, cleaning up directories, and maintaining metadata standards. It also integrates with Plex and supports additional tools like Advanced Renamer and MP3Tag.

---

## Features

- **Initial Setup & Configuration**
  - Stores settings (folders, tool paths, Plex token, etc.) in the Windows Registry.
  - Prompts for missing configuration values on first run or via the Options Menu.

- **Main Menu Functions**
  1. **Start Monitoring FLAC Folder** *(BROKEN)*  
     Monitors the FLAC folder for new files and processes them automatically.
  2. **Copy and Organize MP3 Files**  
     Renames MP3 files to remove invalid characters, extracts metadata, and moves them into an organized folder structure (`Artist\Album`).
  3. **Delete Empty Folders in FLAC Folder**  
     Scans and deletes empty directories in the FLAC folder.
  4. **Display Folders in FLAC Directory**  
     Lists all folders within the FLAC directory.
  5. **Options Menu**  
     Allows updating of all configuration settings.
  6. **Refresh Plex Library**  
     Sends a refresh request to a Plex Media Server library section.
  7. **Scan for Low-Bitrate Files**  
     Scans a selected folder for MP3 files below a user-defined minimum bitrate.
  8. **Scan and Delete Low-Bitrate Files**  
     Scans and optionally deletes MP3 files below the minimum bitrate.
  9. **Open Other Applications** *(BROKEN)*  
     Launches Advanced Renamer or MP3Tag if configured.
  10. **Scan Folder and Generate HTML Report**  
      Scans a folder and generates an HTML report of MP3 metadata using PSWriteHTML.
  11. **Scan and Fix Invalid Characters in ID3 Tags**  
      Scans audio files for invalid characters in metadata and fixes them.
  12. **Split FLAC by CUE**  
      Splits FLAC files into tracks using CUE sheets.
  13. **Manually Convert FLAC Files to MP3**  
      Batch converts FLAC files in a selected folder to MP3.
  99. **Exit**  
      Exits the script.

---

## Requirements

- **PowerShell 5.1+** (Windows)
- **ffmpeg** and **ffprobe** executables
- **TagLibSharp.dll** (for advanced ID3 tag reading)
- **PSWriteHTML** PowerShell module (for HTML reports)
- Optional: **Advanced Renamer** and **MP3Tag** applications

---

## Setup

1. Place FLAC2MP3.ps1 and `TagLibSharp.dll` in the same directory.
2. Run the script in PowerShell:
   ```
   powershell -ExecutionPolicy Bypass -File .\FLAC2MP3.ps1
   ```
3. On first run, follow prompts to configure folder locations and tool paths.
4. Use the menu to perform desired operations.

---

## Notes

- All settings are stored in the Windows Registry under `HKCU:\Software\FLAC2MP3`.
- The script is designed for interactive use and will prompt for user input as needed.
- Some menu options are marked as **BROKEN** and may not function as intended.
- The script can be extended or customized as needed for your workflow.

---

## Disclaimer

This script is provided as-is. Use at your own risk. Always back up your music library before performing batch operations.
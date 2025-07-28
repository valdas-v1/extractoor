#Requires -Version 5.1

<#
.SYNOPSIS
    Extractoor - Smart Media Backup Tool for Windows
    
.DESCRIPTION
    This script helps backup media files (images, videos, audio) from old drives/directories
    to a new location with deduplication based on file hashes and organized folder structure.
    
    Key Features:
    - Preserves original file timestamps (creation, modification, access dates)
    - Maintains file attributes and metadata
    - Deduplicates files using MD5 hashing
    - Organizes files by type and date
    - Provides preview mode for safe testing
    
.AUTHOR
    Generated for media backup and deduplication
    
.VERSION
    1.0
#>

param(
    [string]$SourcePath = "",
    [string]$DestinationPath = "",
    [switch]$Preview = $false,
    [switch]$Verbose = $false,
    [int]$MinFileSizeKB = 10
)

# Media file extensions to backup
$MediaExtensions = @(
    # Images
    "*.jpg", "*.jpeg", "*.png", "*.gif", "*.bmp", "*.tiff", "*.tif", 
    "*.webp", "*.raw", "*.cr2", "*.nef", "*.arw", "*.dng", "*.heic", "*.heif",
    # Videos
    "*.mp4", "*.avi", "*.mkv", "*.mov", "*.wmv", "*.flv", "*.webm", 
    "*.m4v", "*.3gp", "*.mpg", "*.mpeg", "*.ts", "*.mts", "*.m2ts",
    # Audio
    "*.mp3", "*.wav", "*.flac", "*.aac", "*.ogg", "*.wma", "*.m4a"
)

# Global variables
$Global:ProcessedFiles = @{}
$Global:DuplicateCount = 0
$Global:BackedUpCount = 0
$Global:ErrorCount = 0
$Global:ErrorMessages = @()
$Global:SkippedSmallFiles = 0
$Global:TimestampFixCount = 0
$Global:TotalSize = 0
$Global:SavedSpace = 0

# Minimum file size threshold (to exclude icons and thumbnails)
$Global:MinFileSizeBytes = $MinFileSizeKB * 1024  # Convert KB to bytes

# Function to write colored output
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

# Function to display banner
function Show-Banner {
    Clear-Host
    Write-ColorOutput "=================================================================" "Cyan"
    Write-ColorOutput "                      EXTRACTOOR                                " "Cyan"
    Write-ColorOutput "                  Media Backup Tool                             " "Cyan"
    Write-ColorOutput "              Backup & Deduplicate Media Files                  " "Cyan"
    Write-ColorOutput "=================================================================" "Cyan"
    Write-Host ""
}

# Function to get available drives
function Get-AvailableDrives {
    $drives = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 -or $_.DriveType -eq 2 }
    return $drives | Select-Object DeviceID, VolumeName, Size, FreeSpace, @{Name='SizeGB';Expression={[math]::Round($_.Size/1GB,2)}}, @{Name='FreeGB';Expression={[math]::Round($_.FreeSpace/1GB,2)}}
}

# Function to prompt for source path
function Get-SourcePath {
    Write-ColorOutput "`nAvailable Drives:" "Yellow"
    Write-ColorOutput "=================================================================" "Yellow"
    
    $drives = Get-AvailableDrives
    $drives | Format-Table -AutoSize DeviceID, VolumeName, SizeGB, FreeGB | Out-Host
    
    do {
        Write-Host ""
        $sourcePath = Read-Host "Enter source path (drive letter like 'D:\' or specific folder path)"
        
        if ([string]::IsNullOrWhiteSpace($sourcePath)) {
            Write-ColorOutput "ERROR: Source path cannot be empty!" "Red"
            continue
        }
        
        if (-not (Test-Path $sourcePath)) {
            Write-ColorOutput "ERROR: Path '$sourcePath' does not exist!" "Red"
            continue
        }
        
        # Confirm if it's a drive root
        if ($sourcePath -match "^[A-Za-z]:\\?$") {
            $confirm = Read-Host "WARNING: You selected a drive root ($sourcePath). This will scan the entire drive. Continue? (y/n)"
            if ($confirm -notmatch "^[Yy]") {
                continue
            }
        }
        
        return $sourcePath
        
    } while ($true)
}

# Function to prompt for destination path
function Get-DestinationPath {
    do {
        Write-Host ""
        $destPath = Read-Host "Enter destination path for backup (will be created if it doesn't exist)"
        
        if ([string]::IsNullOrWhiteSpace($destPath)) {
            Write-ColorOutput "ERROR: Destination path cannot be empty!" "Red"
            continue
        }
        
        # Check if parent directory exists
        $parentDir = Split-Path $destPath -Parent
        if (-not (Test-Path $parentDir)) {
            Write-ColorOutput "ERROR: Parent directory '$parentDir' does not exist!" "Red"
            continue
        }
        
        # Create destination if it doesn't exist
        if (-not (Test-Path $destPath)) {
            try {
                New-Item -ItemType Directory -Path $destPath -Force | Out-Null
                Write-ColorOutput "SUCCESS: Created destination directory: $destPath" "Green"
            }
            catch {
                Write-ColorOutput "ERROR: Failed to create destination directory: $($_.Exception.Message)" "Red"
                continue
            }
        }
        
        return $destPath
        
    } while ($true)
}

# Function to preserve file attributes and metadata
function Set-PreservedFileAttributes {
    param(
        [System.IO.FileInfo]$SourceFile,
        [string]$DestinationPath
    )
    
    try {
        $destFile = Get-Item -LiteralPath $DestinationPath
        
        # Get source timestamps
        $sourceCreation = $SourceFile.CreationTime
        $sourceLastWrite = $SourceFile.LastWriteTime
        $sourceLastAccess = $SourceFile.LastAccessTime
        
        # Try to get the most accurate date from EXIF/metadata first
        $mediaDateTaken = Get-MediaDateTaken -File $SourceFile
        $bestCreationDate = $sourceCreation
        $bestModificationDate = $sourceLastWrite
        
        if ($mediaDateTaken) {
            # Use media metadata date as the creation date since it's the most accurate
            $bestCreationDate = $mediaDateTaken
            
            # For modification date, use metadata date if it's newer than the current modification date,
            # otherwise keep the file system modification date (which might be when it was edited)
            if ($mediaDateTaken -gt $sourceLastWrite) {
                $bestModificationDate = $mediaDateTaken
            }
            
            if ($Verbose) {
                $mediaType = if ($SourceFile.Extension -match "\.(jpg|jpeg|tiff|tif)$") { "EXIF" } else { "Media metadata" }
                Write-ColorOutput "$mediaType DATE APPLIED: $($SourceFile.Name) - Using metadata date: $($mediaDateTaken.ToString('yyyy-MM-dd HH:mm:ss'))" "Cyan"
            }
            $Global:TimestampFixCount++
        } else {
            # No EXIF data available, apply standard timestamp fix logic
            # Fix timestamp logic: Creation should never be later than modification
            if ($sourceCreation -gt $sourceLastWrite) {
                # If creation date is later than modification date, use the modification date for creation
                $bestCreationDate = $sourceLastWrite
                
                $Global:TimestampFixCount++
                if ($Verbose) {
                    Write-ColorOutput "TIMESTAMP FIX: $($SourceFile.Name) - Creation date was later than modification date" "Yellow"
                }
            }
        }
        
        # Apply the determined timestamps
        $destFile.CreationTime = $bestCreationDate
        $destFile.CreationTimeUtc = $bestCreationDate.ToUniversalTime()
        $destFile.LastWriteTime = $bestModificationDate
        $destFile.LastWriteTimeUtc = $bestModificationDate.ToUniversalTime()
        $destFile.LastAccessTime = $sourceLastAccess
        $destFile.LastAccessTimeUtc = $sourceLastAccess.ToUniversalTime()
        
        # Preserve file attributes (Hidden, ReadOnly, etc.)
        $destFile.Attributes = $SourceFile.Attributes
        
        return $true
    }
    catch {
        Write-ColorOutput "WARNING: Could not preserve all file attributes for: $($SourceFile.Name)" "Yellow"
        return $false
    }
}

# Function to calculate file hash
function Get-FileHashMD5 {
    param([string]$FilePath)
    
    try {
        # Use -LiteralPath for better handling of special characters and long paths
        $hash = Get-FileHash -LiteralPath $FilePath -Algorithm MD5
        return $hash.Hash
    }
    catch {
        # Handle long path names and other file access issues gracefully
        $fileName = Split-Path $FilePath -Leaf
        if ($FilePath.Length -gt 250) {
            Write-ColorOutput "WARNING: Path too long, skipping file: $fileName" "Yellow"
        } else {
            Write-ColorOutput "WARNING: Failed to calculate hash for: $fileName" "Yellow"
        }
        return $null
    }
}

# Function to create clean filename from original file
function Get-CleanFileName {
    param(
        [System.IO.FileInfo]$File,
        [string]$FileHash,
        [DateTime]$FileDate
    )
    
    try {
        # Get hash prefix (8 characters)
        $hashPrefix = $FileHash.Substring(0, 8)
        
        # Format date as YYYY-MM-DD
        $dateStr = $FileDate.ToString("yyyy-MM-dd")
        
        # Get clean extension
        $extension = $File.Extension.ToLower()
        
        # Create clean filename: YYYY-MM-DD_HASH.ext (no time component for cleaner names)
        $cleanFileName = "$dateStr`_$hashPrefix$extension"
        
        return $cleanFileName
    }
    catch {
        # Fallback to simple hash + extension if anything fails
        $hashPrefix = $FileHash.Substring(0, 8)
        $extension = $File.Extension.ToLower()
        return "$hashPrefix$extension"
    }
}

# Function to extract metadata date from image and video files
function Get-MediaDateTaken {
    param([System.IO.FileInfo]$File)
    
    try {
        $shell = New-Object -ComObject Shell.Application
        $folder = $shell.Namespace($File.DirectoryName)
        $item = $folder.ParseName($File.Name)
        $dateStr = $null
        
        # Get the appropriate metadata field based on file type
        if ($File.Extension -match "\.(jpg|jpeg|tiff|tif)$") {
            # For images, use "Date taken" (index 12)
            $dateStr = $folder.GetDetailsOf($item, 12)
        }
        elseif ($File.Extension -match "\.(mp4|mov|avi|mkv|wmv|m4v|flv|webm)$") {
            # For videos, use "Media created" (index 208)
            $dateStr = $folder.GetDetailsOf($item, 208)
        }
        
        if ($dateStr -and $dateStr -ne "") {
            try {
                # Clean the date string - remove invisible Unicode characters and extra whitespace
                $cleanDateStr = $dateStr -replace '[\u200E\u200F\u202A\u202B\u202C\u202D\u202E\u2066\u2067\u2068\u2069]', '' # Remove Unicode formatting characters
                $cleanDateStr = $cleanDateStr -replace '\s+', ' ' # Replace multiple spaces with single space
                $cleanDateStr = $cleanDateStr.Trim()
                
                # Try multiple date formats
                $dateFormats = @(
                    'yyyy-MM-dd HH:mm',
                    'yyyy-MM-dd H:mm',
                    'M/d/yyyy H:mm:ss tt',
                    'M/d/yyyy HH:mm:ss',
                    'yyyy:MM:dd HH:mm:ss',
                    'dd/MM/yyyy HH:mm:ss',
                    'MM/dd/yyyy HH:mm:ss'
                )
                
                foreach ($format in $dateFormats) {
                    try {
                        $mediaDate = [DateTime]::ParseExact($cleanDateStr, $format, $null)
                        return $mediaDate
                    }
                    catch {
                        # Try next format
                    }
                }
                
                # If exact parsing fails, try general parsing as last resort
                $mediaDate = [DateTime]::Parse($cleanDateStr)
                return $mediaDate
            }
            catch {
                # Parse failed, return null
                return $null
            }
        }
        
        return $null
    }
    catch {
        return $null
    }
}

# Function to get file creation date for folder organization
function Get-FileDate {
    param([System.IO.FileInfo]$File)
    
    try {
        # Try to get date from metadata first (for images and videos)
        $mediaDate = Get-MediaDateTaken -File $File
        if ($mediaDate) {
            return $mediaDate
        }
        
        # Get file timestamps as fallback
        $creationTime = $File.CreationTime
        $lastWriteTime = $File.LastWriteTime
        
        # Choose the most logical date for organization
        # If creation date is later than modification date, it's likely incorrect
        if ($creationTime -gt $lastWriteTime) {
            # Use the modification date as it's more reliable in this case
            return $lastWriteTime
        } else {
            # Use the earlier of creation or modification time
            return if ($creationTime -lt $lastWriteTime) { $creationTime } else { $lastWriteTime }
        }
    }
    catch {
        return $File.CreationTime
    }
}

# Function to find media files
function Find-MediaFiles {
    param([string]$Path)
    
    Write-ColorOutput "`nScanning for media files..." "Yellow"
    Write-ColorOutput "Source: $Path" "Gray"
    Write-ColorOutput "Filter: Files >= ${MinFileSizeKB}KB (excludes icons/thumbnails)" "Gray"
    
    $mediaFiles = @()
    
    foreach ($extension in $MediaExtensions) {
        try {
            # Use -LiteralPath for better handling of paths with special characters
            $files = Get-ChildItem -LiteralPath $Path -Filter $extension -Recurse -File -ErrorAction SilentlyContinue
            $mediaFiles += $files
        }
        catch {
            # Continue with other extensions if one fails
        }
    }
    
    return $mediaFiles
}

# Function to backup single file
function Backup-MediaFile {
    param(
        [System.IO.FileInfo]$File,
        [string]$DestinationRoot,
        [bool]$PreviewMode = $false
    )
    
    try {
        # Check for extremely long file paths that could cause issues
        if ($File.FullName.Length -gt 250) {
            $errorMsg = "Skipping file with very long path (>250 chars): $($File.Name)"
            Write-ColorOutput $errorMsg "Yellow"
            $Global:ErrorCount++
            $Global:ErrorMessages += $errorMsg
            return $false
        }
        
        # Check file size first - skip small files (likely icons, thumbnails)
        if ($File.Length -lt $Global:MinFileSizeBytes) {
            $Global:SkippedSmallFiles++
            
            if ($Verbose) {
                $sizeKB = [math]::Round($File.Length / 1024, 1)
                $minKB = [math]::Round($Global:MinFileSizeBytes / 1024, 0)
                Write-ColorOutput "Skipped small file: $($File.Name) (${sizeKB}KB < ${minKB}KB)" "Gray"
            }
            return $true
        }
        
        # Calculate file hash
        $fileHash = Get-FileHashMD5 -FilePath $File.FullName
        if (-not $fileHash) {
            $errorMsg = "Failed to calculate hash for: $($File.FullName)"
            $Global:ErrorCount++
            $Global:ErrorMessages += $errorMsg
            return $false
        }
        
        # Check if we've already processed this hash
        if ($Global:ProcessedFiles.ContainsKey($fileHash)) {
            $Global:DuplicateCount++
            $Global:SavedSpace += $File.Length
            
            if ($Verbose) {
                Write-ColorOutput "Duplicate found: $($File.Name) (same as $($Global:ProcessedFiles[$fileHash]))" "Yellow"
            }
            return $true
        }
        
        # Get file date for organization
        $fileDate = Get-FileDate -File $File
        $yearMonth = $fileDate.ToString("yyyy-MM")
        
        # Determine file type folder
        $fileType = switch ($File.Extension.ToLower()) {
            {$_ -in @(".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".tif", ".webp", ".raw", ".cr2", ".nef", ".arw", ".dng", ".heic", ".heif")} { "Images" }
            {$_ -in @(".mp4", ".avi", ".mkv", ".mov", ".wmv", ".flv", ".webm", ".m4v", ".3gp", ".mpg", ".mpeg", ".ts", ".mts", ".m2ts")} { "Videos" }
            {$_ -in @(".mp3", ".wav", ".flac", ".aac", ".ogg", ".wma", ".m4a")} { "Audio" }
            default { "Other" }
        }
        
        # Create destination path: DestinationRoot\FileType\YYYY-MM\
        if (-not $PreviewMode) {
            $destFolder = Join-Path $DestinationRoot $fileType
            $destFolder = Join-Path $destFolder $yearMonth
            
            # Create clean filename
            $newFileName = Get-CleanFileName -File $File -FileHash $fileHash -FileDate $fileDate
            $destFilePath = Join-Path $destFolder $newFileName
        } else {
            # In preview mode, just create the relative path for display
            $destFolder = "$fileType\$yearMonth"
            $newFileName = Get-CleanFileName -File $File -FileHash $fileHash -FileDate $fileDate
            $destFilePath = "$destFolder\$newFileName"
        }
        
        if (-not $PreviewMode) {
            # Create destination folder if it doesn't exist
            if (-not (Test-Path -LiteralPath $destFolder)) {
                New-Item -ItemType Directory -Path $destFolder -Force | Out-Null
            }
            
            # Copy the file using -LiteralPath for better path handling
            Copy-Item -LiteralPath $File.FullName -Destination $destFilePath -Force
            
            # Preserve original file timestamps and attributes
            Set-PreservedFileAttributes -SourceFile $File -DestinationPath $destFilePath
        }
        
        # Record this hash as processed
        $Global:ProcessedFiles[$fileHash] = $newFileName
        $Global:BackedUpCount++
        $Global:TotalSize += $File.Length
        
        if ($Verbose) {
            $action = if ($PreviewMode) { "Would backup" } else { "Backed up" }
            Write-ColorOutput "$action`: $($File.Name) -> $fileType\$yearMonth\$newFileName" "Green"
        }
        
        return $true
    }
    catch {
        $errorMsg = "Error backing up $($File.FullName): $($_.Exception.Message)"
        Write-ColorOutput $errorMsg "Red"
        $Global:ErrorCount++
        $Global:ErrorMessages += $errorMsg
        return $false
    }
}

# Function to display progress
function Show-Progress {
    param(
        [int]$Current,
        [int]$Total,
        [string]$CurrentFile
    )
    
    $percent = [math]::Round(($Current / $Total) * 100, 1)
    $completed = [math]::Round($percent / 4)  # 25 character width
    $remaining = 25 - $completed
    
    # Use simple ASCII characters for better compatibility
    $progressBar = "[" + ("=" * $completed) + ("." * $remaining) + "]"
    
    # Truncate filename if too long to prevent line wrapping
    $displayFile = $CurrentFile
    if ($displayFile.Length -gt 30) {
        $displayFile = $displayFile.Substring(0, 27) + "..."
    }
    
    # Clear the line first, then write progress
    Write-Host "`r" + (" " * 80) + "`r" -NoNewline
    Write-Host "$progressBar $percent% ($Current/$Total) - $displayFile" -NoNewline
}

# Function to format file size
function Format-FileSize {
    param([long]$Bytes)
    
    if ($Bytes -gt 1GB) {
        return "{0:N2} GB" -f ($Bytes / 1GB)
    }
    elseif ($Bytes -gt 1MB) {
        return "{0:N2} MB" -f ($Bytes / 1MB)
    }
    elseif ($Bytes -gt 1KB) {
        return "{0:N2} KB" -f ($Bytes / 1KB)
    }
    else {
        return "$Bytes bytes"
    }
}

# Function to display summary
function Show-Summary {
    param([bool]$PreviewMode = $false)
    
    $action = if ($PreviewMode) { "Would be backed up" } else { "Backed up" }
    $saved = if ($PreviewMode) { "Would save" } else { "Saved" }
    
    Write-Host "`n"
    Write-ColorOutput "=================================================================" "Cyan"
    Write-ColorOutput "                        SUMMARY                                 " "Cyan"
    Write-ColorOutput "=================================================================" "Cyan"
    Write-ColorOutput "Files $action`: $Global:BackedUpCount" "Green"
    Write-ColorOutput "Small files skipped: $Global:SkippedSmallFiles (< ${MinFileSizeKB}KB)" "Gray"
    Write-ColorOutput "Timestamp fixes applied: $Global:TimestampFixCount" "Cyan"
    Write-ColorOutput "Duplicates found: $Global:DuplicateCount" "Yellow"
    
    if ($Global:ErrorCount -gt 0) {
        Write-ColorOutput "Errors encountered: $Global:ErrorCount" "Red"
        foreach ($errorMsg in $Global:ErrorMessages) {
            Write-ColorOutput "  - $errorMsg" "Red"
        }
    } else {
        Write-ColorOutput "Errors encountered: 0" "Green"
    }
    
    Write-ColorOutput "Total size $action`: $(Format-FileSize $Global:TotalSize)" "Cyan"
    Write-ColorOutput "Space $saved by deduplication: $(Format-FileSize $Global:SavedSpace)" "Magenta"
    
    if (-not $PreviewMode) {
        Write-ColorOutput "Original timestamps and metadata preserved" "Cyan"
    }
    
    if (-not $PreviewMode -and $Global:BackedUpCount -gt 0) {
        Write-ColorOutput "`nBACKUP COMPLETED SUCCESSFULLY!" "Green"
        Write-ColorOutput "All files maintain their original creation dates for proper photo library import!" "Green"
    }
}

# Main execution function
function Start-MediaBackup {
    param(
        [string]$Source,
        [string]$Destination,
        [bool]$PreviewMode = $false
    )
    
    # Reset counters
    $Global:ProcessedFiles = @{}
    $Global:DuplicateCount = 0
    $Global:BackedUpCount = 0
    $Global:ErrorCount = 0
    $Global:ErrorMessages = @()
    $Global:SkippedSmallFiles = 0
    $Global:TimestampFixCount = 0
    $Global:TotalSize = 0
    $Global:SavedSpace = 0
    
    # Get source path if not provided
    if ([string]::IsNullOrWhiteSpace($Source)) {
        $Source = Get-SourcePath
    }
    
    # Get destination path if not provided (and not in preview mode)
    if ([string]::IsNullOrWhiteSpace($Destination) -and -not $PreviewMode) {
        $Destination = Get-DestinationPath
    }
    
    # Find media files
    $mediaFiles = Find-MediaFiles -Path $Source
    
    if ($mediaFiles.Count -eq 0) {
        Write-ColorOutput "WARNING: No media files found in the specified path!" "Yellow"
        return
    }
    
    Write-ColorOutput "Found $($mediaFiles.Count) media files" "Green"
    
    if (-not $PreviewMode) {
        Write-ColorOutput "Backup destination: $Destination" "Gray"
    }
    
    if ($PreviewMode) {
        Write-ColorOutput "`nPREVIEW MODE - No files will be copied" "Magenta"
    }
    
    # Confirm before proceeding
    if (-not $PreviewMode) {
        Write-Host ""
        $confirm = Read-Host "Continue with backup? (y/n)"
        if ($confirm -notmatch "^[Yy]") {
            Write-ColorOutput "Backup cancelled by user" "Red"
            return
        }
    }
    
    Write-ColorOutput "`nStarting backup process..." "Green"
    Write-Host ""
    
    # Process each file
    for ($i = 0; $i -lt $mediaFiles.Count; $i++) {
        $file = $mediaFiles[$i]
        
        # Show progress
        if (-not $Verbose) {
            Show-Progress -Current ($i + 1) -Total $mediaFiles.Count -CurrentFile $file.Name
        }
        
        # Backup the file
        $result = Backup-MediaFile -File $file -DestinationRoot $Destination -PreviewMode $PreviewMode
    }
    
    Write-Host "`n"
    Show-Summary -PreviewMode $PreviewMode
}

# Main script execution
try {
    Show-Banner
    
    # Check if running with parameters
    if ($SourcePath -and ($DestinationPath -or $Preview)) {
        Start-MediaBackup -Source $SourcePath -Destination $DestinationPath -PreviewMode $Preview
    }
    else {
        # Interactive mode
        Write-ColorOutput "Welcome to Extractoor!" "Green"
        Write-ColorOutput "This tool will help you backup and deduplicate media files from old drives." "Gray"
        
        do {
            Write-Host ""
            Write-ColorOutput "Options:" "Yellow"
            Write-ColorOutput "1. Preview backup (scan only, no files copied)" "Cyan"
            Write-ColorOutput "2. Start backup" "Green"
            Write-ColorOutput "3. Exit" "Red"
            
            $choice = Read-Host "`nSelect an option (1-3)"
            
            switch ($choice) {
                "1" {
                    $source = Get-SourcePath
                    Start-MediaBackup -Source $source -Destination "" -PreviewMode $true
                }
                "2" {
                    Start-MediaBackup
                }
                "3" {
                    Write-ColorOutput "Goodbye!" "Green"
                    exit 0
                }
                default {
                    Write-ColorOutput "Invalid option. Please select 1, 2, or 3." "Red"
                }
            }
            
            if ($choice -in @("1", "2")) {
                Write-Host ""
                $continue = Read-Host "Press Enter to return to main menu or 'q' to quit"
                if ($continue -eq 'q') {
                    break
                }
            }
            
        } while ($true)
    }
}
catch {
    Write-ColorOutput "An unexpected error occurred: $($_.Exception.Message)" "Red"
    Write-ColorOutput "Stack trace: $($_.ScriptStackTrace)" "Gray"
}
finally {
    Write-ColorOutput "`nThank you for using Extractoor!" "Green"
}

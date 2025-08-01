# Extractoor

*Smart media backup with authentic timestamps*

Extract and organize your valuable media files from old drives, phones, and backups. Automatically deduplicates, preserves authentic capture dates, and creates a clean chronological archive of your photos, videos, and audio.

## What Makes It Different

**Metadata Intelligence**: Reads actual capture dates from EXIF (photos) and media metadata (videos), not filesystem timestamps that lie.

**Zero Duplicates**: MD5 deduplication across your entire backup.

**True Chronology**: Files organized by real capture date, perfect for photo libraries.

**Unicode Safe**: Handles special characters and international filenames.

**Safe by Design**: Read-only operation with preview mode.

## Supported Formats

**Images**: JPG, PNG, TIFF, RAW, CR2, NEF, ARW, DNG, HEIC, WebP  
**Videos**: MP4, AVI, MKV, MOV, WMV, M4V, 3GP, MPEG  
**Audio**: MP3, WAV, FLAC, AAC, OGG, M4A

## Quick Start

```powershell
# Interactive mode
.\Extractoor.ps1

# Preview what would be backed up
.\Extractoor.ps1 -SourcePath "D:\Old Phone" -Preview

# Direct backup with progress
.\Extractoor.ps1 -SourcePath "D:\Photos" -DestinationPath "E:\Backup" -Verbose
```

## How It Works

1. **Metadata Extraction**: Reads EXIF "Date Taken" (photos) and "Media Created" (videos)
2. **Intelligent Fallback**: Uses filesystem dates only when metadata unavailable  
3. **Deduplication**: MD5 hashing prevents duplicate storage
4. **Smart Organization**: Groups by type and real capture date
5. **Clean Naming**: `2024-07-15_14.32.45_A1B2C3D4.jpg` format with full timestamp
6. **Incremental Safety**: Skips identical files, warns on conflicts

## Output Structure

```
Backup/
├── Images/2024-07/2024-07-15_14.32.45_A1B2C3D4.jpg    # Full timestamp for perfect sorting
├── Videos/2024-10/2024-10-08_09.15.22_9I0J1K2L.mp4    # Precise chronological order  
└── Audio/2024-11/2024-11-20_18.45.12_M3N4O5P6.mp3
```

Files named with precise capture time for perfect chronological sorting. Re-running backups only copies new files.

## Technical Details

**Date Sources**: EXIF index 12 (images), Media Created index 208 (videos)  
**Unicode Handling**: Strips invisible formatting characters from metadata  
**Path Safety**: Handles 250+ character paths and special characters  
**Performance**: Single directory scan with extension filtering (not per-extension recursion)  
**Incremental Backups**: Hash comparison skips identical files, shows new vs existing counts  
**Error Recovery**: Continues processing with detailed error reporting  
**Size Filtering**: Excludes files <10KB (configurable) to skip thumbnails

## Advanced Usage

```powershell
# Handle international characters
.\Extractoor.ps1 -SourcePath "D:\Fotos\Año 2024\Niños y niñas"

# Custom size threshold
.\Extractoor.ps1 -SourcePath "D:\Media" -DestinationPath "E:\Backup" -MinFileSizeKB 50

# Preview with detailed output
.\Extractoor.ps1 -SourcePath "F:\" -Preview -Verbose
```

## Why It Matters

Most backup tools copy files with current timestamps, destroying chronological order. Extractoor preserves the moment you actually captured that photo or recorded that video, making your backup truly useful for photo management software.

**Before**: `IMG_20241008_190245.jpg` (Creation: Dec 23, 2024)  
**After**: `2024-10-08_19.02.45_9A45C7B2.jpg` (Creation: Oct 8, 2024)

The second file will appear in your photo library on the correct date and time. Incremental backups are fast since identical files are automatically skipped.

## Installation

Download `Extractoor.ps1` and run. No installation required.

## Requirements

- Windows PowerShell 5.1+
- Read access to source
- Write access to destination

## License

**Copyright (c) 2025 Valdas Paulavičius**  
**LinkedIn**: https://www.linkedin.com/in/valdas-paulavicius/

All rights reserved.

### Personal & Educational Use
Permission is granted to use, copy, and distribute this software for **personal and educational purposes only**, provided that:
- The above copyright notice and this permission notice appear in all copies
- Proper attribution to Valdas Paulavičius is maintained
- No commercial use is made of the software

### Commercial Use
Commercial use, modification, and distribution require **explicit written permission** from the copyright holder. Contact via LinkedIn for commercial licensing inquiries.

### Attribution Requirement
Attribution must be maintained in all copies, derivative works, and distributions. Suggested attribution format:
```
Extractoor by Valdas Paulavičius (https://www.linkedin.com/in/valdas-paulavicius/)
```

### Disclaimer
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.

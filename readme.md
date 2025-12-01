# Add-TracklistChapters

Automatically add chapter markers to video files using tracklists from [1001Tracklists.com](https://www.1001tracklists.com).

Navigate DJ sets, live recordings, and music mixes track by track.

## Features

- 🔍 **Smart Search** - Search 1001Tracklists.com directly from the command line
- 📁 **Filename Detection** - Automatically derives search query from video filename
- 🎯 **Intelligent Relevance Scoring** - Results ranked by duration match, keywords, abbreviations, event patterns, year, and recency
- 🔤 **Abbreviation Matching** - Recognizes event abbreviations (e.g., `AMF` matches "Amsterdam Music Festival")
- 📅 **Event Pattern Detection** - Understands multi-day event notation (`WE1`, `WE2`, `Day1`, `D2`, etc.)
- 🌍 **Accent-Insensitive Matching** - `Chateau` matches `Château`, `Ibanez` matches `Ibañez`
- 🎬 **YouTube ID Stripping** - Automatically removes yt-dlp video IDs from filenames (e.g., `[dQw4w9WgXcQ]`)
- ⚡ **Session Caching** - Login cookies cached for ~100 days for faster consecutive runs
- 🔄 **Duplicate Detection** - Skips files that already have identical chapters
- 📦 **Pipeline Support** - Batch process multiple files via PowerShell pipeline
- ⏭️ **Auto-Select Mode** - Fully automated chapter embedding for batch processing

## Requirements

- PowerShell 5.1+ (Windows) or PowerShell Core 7+ (cross-platform)
- [MKVToolNix](https://mkvtoolnix.download/) installed (provides `mkvmerge` and `mkvextract`)
- 1001Tracklists.com account (free, required for search functionality)

## Installation

1. Download `Add-TracklistChapters.ps1`
2. Create a configuration file with your credentials:
   ```powershell
   .\Add-TracklistChapters.ps1 -CreateConfig
   ```
3. Edit `config.json` in the script directory:
   ```json
   {
     "Email": "your-email@example.com",
     "Password": "your-password",
     "ChapterLanguage": "eng",
     "MkvMergePath": "",
     "ReplaceOriginal": false,
     "NoDurationFilter": false
   }
   ```

## Usage

### Basic Usage (Search from Filename)

Simply provide a video file - the script will search using the filename:

```powershell
.\Add-TracklistChapters.ps1 -InputFile "2025 - AMF - Sub Zero Project.webm"
```

This searches for "2025 AMF Sub Zero Project" and presents matching tracklists for selection.

### Automated Mode

Use `-AutoSelect` to automatically pick the best match:

```powershell
.\Add-TracklistChapters.ps1 -InputFile "video.mkv" -AutoSelect
```

### Custom Search Query

```powershell
.\Add-TracklistChapters.ps1 -InputFile "video.mkv" -Tracklist "Hardwell Ultra Miami 2024"
```

### Direct URL or ID

```powershell
# Using URL
.\Add-TracklistChapters.ps1 -InputFile "video.mkv" -Tracklist "https://www.1001tracklists.com/tracklist/1g6g22ut/..."

# Using tracklist ID
.\Add-TracklistChapters.ps1 -InputFile "video.mkv" -Tracklist "1g6g22ut"
```

### Local Tracklist File

```powershell
.\Add-TracklistChapters.ps1 -InputFile "video.mkv" -TrackListFile "chapters.txt"
```

Tracklist file format:
```
[00:00] Artist - Track Title
[03:45] Another Artist - Another Track
[1:07:30] Third Artist - Third Track
```

### From Clipboard

Copy a tracklist from your browser, then:

```powershell
.\Add-TracklistChapters.ps1 -InputFile "video.mkv" -FromClipboard
```

### Preview Mode

See what chapters would be added without modifying the file:

```powershell
.\Add-TracklistChapters.ps1 -InputFile "video.mkv" -Preview
```

### Replace Original File

```powershell
.\Add-TracklistChapters.ps1 -InputFile "video.mkv" -ReplaceOriginal
```

## Batch Processing

The script supports pipeline input for processing multiple files:

```powershell
# Process all MKV files with auto-select
Get-ChildItem "D:\DJ Sets\*.mkv" | .\Add-TracklistChapters.ps1 -AutoSelect -ReplaceOriginal

# Process all WEBM files with a specific search query
Get-ChildItem "*.webm" | .\Add-TracklistChapters.ps1 -Tracklist "Festival Name 2025" -AutoSelect

# Preview chapters for all files without modifying
Get-ChildItem "*.webm" | .\Add-TracklistChapters.ps1 -Preview
```

A summary is displayed after batch processing:

```
--------------------------------------------------
Summary: 5 succeeded, 1 failed
```

## Search Algorithm

The script uses intelligent relevance scoring to find the best matching tracklist:

### Scoring Factors (in order of importance)

| Factor | Points | Description |
|--------|--------|-------------|
| **Duration Match** | -20 to +100 | Exact duration match scores highest; large mismatches are penalized |
| **Keywords** | 0 to 75 | Proportional to matched keywords, +15 bonus if all keywords match |
| **Event Patterns** | -30 to +40 | Matches `WE1`/`Weekend 1`, `D2`/`Day 2`, etc. Wrong pattern is penalized |
| **Abbreviations** | +35 each | Matches `AMF` to "Amsterdam Music Festival", `EDC` to "Electric Daisy Carnival" |
| **Year** | +25 | Matches year in query (e.g., `2024`) to tracklist date |
| **Recency** | 0 to 10 | Newer tracklists score slightly higher |

### Smart Query Parsing

The script automatically handles common filename patterns:

- **YouTube IDs**: `Video Title [dQw4w9WgXcQ]` → strips the ID
- **Accented characters**: `Château` and `Chateau` both match
- **Event patterns**: `WE2`, `W2`, `Weekend2` all recognized as "Weekend 2"
- **Abbreviations**: Uppercase words like `AMF`, `EDC`, `ASOT` are matched against full event names

### Example Verbose Output

```
VERBOSE: Query parts: Year=2024, Keywords=[belgium, otto, knows], Abbreviations=[], EventPatterns=[Weekend2]
VERBOSE:   Score 215.4 for 'Otto Knows @ Tomorrowland Weekend 2...' [Dur:100(0m diff) | Kw:45(3/3) | Year:25 | Pat:40 | Rec:5.4]
```

## Search Results Display

```
Search Results (video: 54m):
--------------------------------------------------------------------------------------------------------------
  1. Sub Zero Project @ Amsterdam Music Festival, Netherlands [📅 2025-10-25 | ⏱️ 54m ✓]
  2. Sub Zero Project @ Defqon.1, Netherlands [📅 2025-06-28 | ⏱️ 55m ≈]
  3. Sub Zero Project @ Reverze, Belgium [📅 2025-02-14 | ⏱️ 62m ~]
--------------------------------------------------------------------------------------------------------------
```

**Duration indicators:**
- ✓ (green) - Exact match (±1 minute)
- ≈ (yellow) - Close match (±5 minutes)
- ~ (dim) - Moderate difference (±15 minutes)
- No marker - Larger difference

## Parameters

| Parameter | Description |
|-----------|-------------|
| `-InputFile` | Video file to add chapters to (MKV or WEBM). Accepts pipeline input. |
| `-Tracklist` | Search query, URL, or tracklist ID |
| `-TrackListFile` | Local text file with timestamps |
| `-FromClipboard` | Read tracklist from clipboard |
| `-AutoSelect` | Automatically select the top result |
| `-NoDurationFilter` | Disable duration-based result filtering |
| `-Preview` | Show chapters without modifying file |
| `-ReplaceOriginal` | Replace original file instead of creating new |
| `-OutputFile` | Custom output filename |
| `-ChapterLanguage` | Chapter language code (default: `eng`) |
| `-MkvMergePath` | Custom path to mkvmerge.exe |
| `-Email` | 1001Tracklists.com email (overrides config) |
| `-Password` | 1001Tracklists.com password (overrides config) |
| `-Credential` | PSCredential object (alternative to Email/Password) |
| `-CreateConfig` | Generate default config.json file |
| `-Verbose` | Show detailed scoring breakdown and debug info |

## Configuration

The `config.json` file supports the following options:

```json
{
  "Email": "",
  "Password": "",
  "ChapterLanguage": "eng",
  "MkvMergePath": "",
  "ReplaceOriginal": false,
  "NoDurationFilter": false
}
```

Command-line parameters always override config file values.

## Behavior Notes

### Session Caching
Login cookies are cached in `.1001tl-cookies.json` in the script directory. This speeds up consecutive runs by avoiding repeated logins. Cache is valid until cookies expire (~100 days).

### Duplicate Chapter Detection
Before muxing, the script extracts existing chapters using `mkvextract` and compares them with the new chapters. If they're identical, the file is skipped with a message:
```
Chapters already exist and are identical - skipping.
```

### Tracklists Without Timestamps
If a selected tracklist has no timestamps (common for recently added tracklists), you'll be prompted to select a different one in interactive mode. In `-AutoSelect` mode, an error is thrown.

### Filename to Query Conversion
When using filename-based search, the script:
1. Removes the file extension
2. Replaces `-`, `_`, `.` with spaces
3. Strips YouTube video IDs (11-character codes in brackets at the end)
4. Normalizes multiple spaces

## Troubleshooting

**"Search requires a 1001Tracklists.com account"**
- Ensure your email and password are correct in config.json
- Verify your account works on the website
- Delete `.1001tl-cookies.json` to force a fresh login

**"No tracklists found"**
- Try a different search query
- Use `-NoDurationFilter` to see all results regardless of duration

**"This tracklist has no timestamps yet"**
- Some recent tracklists don't have timestamps added by users yet
- Select a different tracklist or wait for the community to add timestamps

**"Chapters already exist and are identical - skipping"**
- The file already has the same chapters - no action needed
- Use a different tracklist if you want different chapters

**mkvmerge not found**
- Install [MKVToolNix](https://mkvtoolnix.download/)
- Or specify the path: `-MkvMergePath "C:\Program Files\MKVToolNix\mkvmerge.exe"`

**Wrong tracklist selected in AutoSelect mode**
- Use `-Verbose` to see scoring breakdown
- Try adding more specific keywords to the filename or `-Tracklist` parameter
- Include the year, event abbreviation, or weekend/day number for better matching
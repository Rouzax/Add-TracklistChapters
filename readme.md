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
- ⚡ **Fast In-Place Editing** - Uses `mkvpropedit` for near-instant chapter embedding (no remuxing)
- 🍪 **Session Caching** - Login cookies cached for ~100 days for faster consecutive runs
- 🔄 **Duplicate Detection** - Skips files that already have identical chapters
- 🔗 **URL Storage** - Stores tracklist URL in file for instant reuse on subsequent runs
- 📦 **Pipeline Support** - Batch process multiple files via PowerShell pipeline
- ⏭️ **Auto-Select Mode** - Fully automated chapter embedding for batch processing
- 🛡️ **Rate Limit Protection** - Configurable delay between requests to avoid 1001Tracklists rate limits

## Requirements

- PowerShell 5.1+ (Windows) or PowerShell Core 7+ (cross-platform)
- [MKVToolNix](https://mkvtoolnix.download/) installed (provides `mkvmerge`, `mkvextract`, and `mkvpropedit`)
- 1001Tracklists.com account (free, required for search functionality)

## Installation

1. Download `Add-TracklistChapters.ps1`
2. Create configuration files:
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
     "NoDurationFilter": false,
     "AutoSelect": false,
     "DelaySeconds": 5
   }
   ```
4. Optionally edit `aliases.json` to add custom event abbreviations (see [Event Aliases](#event-aliases)).

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

# Process all WEBM files recursively
Get-ChildItem "D:\Videos\*.webm" -Recurse | .\Add-TracklistChapters.ps1 -AutoSelect -ReplaceOriginal

# Preview chapters for all files without modifying
Get-ChildItem "*.webm" | .\Add-TracklistChapters.ps1 -Preview
```

A delay is automatically added between files to avoid rate limiting (default: 5 seconds). You can adjust this:

```powershell
# Faster processing (may trigger rate limits)
Get-ChildItem "*.mkv" | .\Add-TracklistChapters.ps1 -AutoSelect -DelaySeconds 2

# More conservative delay
Get-ChildItem "*.mkv" | .\Add-TracklistChapters.ps1 -AutoSelect -DelaySeconds 10
```

A summary is displayed after batch processing:

```
--------------------------------------------------
Summary: 10 files processed
  7 chapters added
  2 already up-to-date
  1 skipped
```

## Search Algorithm

The script uses intelligent relevance scoring to find the best matching tracklist. Use `-Verbose` to see the scoring breakdown for each result.

### Scoring Factors

| Factor | Points | Description |
|--------|--------|-------------|
| **Duration** | -20 to +100 | ±1 min = 100, ±5 min = 80, ±15 min = 40, ±30 min = 10, >30 min = -20 |
| **Keywords** | 0 to 75 | Proportional to matched keywords, +15 bonus if all match |
| **Abbreviations** | +35 each | `AMF` → "Amsterdam Music Festival", `EDC` → "Electric Daisy Carnival" |
| **Aliases** | +35 each | Matches from aliases.json (case-insensitive), e.g., `tml` → "Tomorrowland" |
| **Event Patterns** | -30 to +40 | Correct pattern = +40, wrong pattern = -30 |
| **Year** | +25 | Matches year in query to tracklist date |
| **Recency** | 0 to 10 | Minor tiebreaker favoring newer tracklists |

### Result Filtering

Results with 0 keyword matches are always filtered out (these are pure noise from duration/year matching).

Additionally, when an event is specified (via alias or abbreviation), results with only 1 keyword match and no event match are filtered out. This removes false positives like "House Party 133" matching "Swedish House Mafia".

### Smart Query Parsing

The script automatically handles common filename patterns:

- **YouTube IDs**: `Video Title [dQw4w9WgXcQ]` → stripped automatically
- **Accented characters**: `Tiësto` and `Tiesto` both match, `Château` matches `Chateau`
- **Event patterns**: `WE2`, `W2`, `Weekend2` all recognized as "Weekend 2" and count as matched keywords
- **Abbreviations**: Uppercase words like `AMF`, `EDC`, `ASOT` matched against full event names
- **Aliases**: Any word matching a key in aliases.json (case-insensitive) boosts results containing the target event name

### Examples

**Event pattern matching (WE1 → Weekend 1):**
```
Query: "2025 - TML Belgium - Hardwell WE1"
QueryParts: Year=2025, Keywords=[tml, belgium, hardwell, we1], Aliases=[TML->Tomorrowland], EventPatterns=[Weekend1]

Score 290.1 for 'Hardwell @ Mainstage, Tomorrowland Weekend 1, Belgium...'
      [Dur:100(1m diff) | Alias:35(TML->Tomorrowland) | Kw:75(4/4) | Year:25 | Pat:40 | Rec:5.1]

Score 124.9 for 'Hardwell @ Mainstage, Tomorrowland Weekend 2, Belgium...'
      [Dur:100(1m diff) | Alias:35(TML->Tomorrowland) | Kw:45(3/4) | Year:25 | Pat:-30 | Rec:5.9]  ← Wrong weekend, penalized
```

**Abbreviation matching (AMF, EDC):**
```
Query: "2025 AMF KI/KI B2B Armin van Buuren"
QueryParts: Year=2025, Keywords=[amf, ki/ki, b2b, armin, van, buuren], Abbreviations=[AMF]

Score 184.2 for 'Armin van Buuren & KI/KI @ Two Is One, Amsterdam Music Festival...'
      [Dur:80(4m diff) | Abbr:35(AMF=AMF) | Kw:40(4/6) | Year:25 | Rec:4.2]  ← AMF matched

Score 141.8 for 'Armin van Buuren - Piano...'
      [Dur:80(5m diff) | Abbr:0 | Kw:30(3/6) | Year:25 | Rec:6.8]            ← No abbreviation match
```

**YouTube ID stripping:**
```
Query: "Marlon Hoffstadt Live at EDC Las Vegas 2025 [JX2STP6HL5k]"
       → YouTube ID stripped, EDC detected as abbreviation

Score 178.2 for 'Marlon Hoffstadt @ kineticFIELD, EDC Las Vegas...'
      [Dur:100(0m diff) | Abbr:35(EDC(direct)) | Kw:38.2(7/11) | Rec:5]
```

**Accent-insensitive matching:**
```
Query: "Tiësto - Dreamstate 2025 (Full Set)"
       → Tiësto normalized for matching

Score 164.9 for 'Tiësto @ The Dream Stage, Dreamstate SoCal...'
      [Dur:100(0m diff) | Kw:30(2/4) | Year:25 | Rec:9.9]
```

**Alias matching (lowercase/unofficial abbreviations):**
```
Query: "armin van buuren - asot 2025 - utrecht"
       → asot matched via aliases.json to "A State of Trance"

Score 184.5 for 'Armin van Buuren @ A State of Trance 1000, Utrecht...'
      [Dur:100(0m diff) | Alias:35(asot->A State of Trance) | Kw:45(4/6) | Year:25 | Rec:4.5]
```

### Score Interpretation

| Score Range | Meaning |
|-------------|---------|
| 200+ | Excellent match - likely the exact recording |
| 150-200 | Good match - probably correct |
| 100-150 | Partial match - review manually |
| <100 | Poor match - likely wrong tracklist |

## Search Results Display

```
Search Results (video: 47m):
--------------------------------------------------------------------------------------------------------------
  1. Armin van Buuren & KI/KI @ Two Is One, Amsterdam Music Festival, Netherlands [📅 2025-10-25 | ⏱️ 51m ≈]
  2. Armin van Buuren - Piano [📅 2025-11-10 | ⏱️ 52m ≈]
  3. Armin van Buuren @ SLAM! (Amsterdam Dance Event, Netherlands) [📅 2025-10-22 | ⏱️ 51m ≈]
  4. Armin van Buuren @ Mainstage, Tomorrowland Brasil [📅 2025-10-10 | ⏱️ 58m ~]
  5. Armin van Buuren @ Amsterdam Music Festival, Netherlands [📅 2025-10-25 | ⏱️ 84m ✗]
--------------------------------------------------------------------------------------------------------------
```

**Duration indicators:**
- ✓ (green) - Exact match (±1 minute)
- ≈ (yellow) - Close match (±5 minutes)
- ~ (dim) - Moderate difference (±15 minutes)
- ✗ (red) - Poor match (>30 minutes difference)
- No marker - Significant difference (±30 minutes)

## Parameters

| Parameter | Description |
|-----------|-------------|
| `-InputFile` | Video file to add chapters to (MKV or WEBM). Accepts pipeline input. |
| `-Tracklist` | Search query, URL, or tracklist ID |
| `-TrackListFile` | Local text file with timestamps |
| `-FromClipboard` | Read tracklist from clipboard |
| `-AutoSelect` | Automatically select the top result |
| `-DelaySeconds` | Delay between files in pipeline mode (default: 5, range: 0-60) |
| `-NoDurationFilter` | Disable duration-based result filtering |
| `-Preview` | Show chapters without modifying file |
| `-IgnoreStoredUrl` | Force fresh search, ignoring any stored tracklist URL |
| `-ReplaceOriginal` | Replace original file instead of creating new |
| `-OutputFile` | Custom output filename |
| `-ChapterLanguage` | Chapter language code (default: `eng`) |
| `-MkvMergePath` | Custom path to MKVToolNix directory |
| `-Email` | 1001Tracklists.com email (overrides config) |
| `-Password` | 1001Tracklists.com password (overrides config) |
| `-Credential` | PSCredential object (alternative to Email/Password) |
| `-CreateConfig` | Generate or update config.json and aliases.json files |
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
  "NoDurationFilter": false,
  "AutoSelect": false,
  "DelaySeconds": 5
}
```

Command-line parameters always override config file values.

### Updating Configuration

Running `-CreateConfig` is non-destructive - it merges new default settings with your existing configuration:

```powershell
.\Add-TracklistChapters.ps1 -CreateConfig
```

Output:
```
config.json: Updated (added: AutoSelect, DelaySeconds)
aliases.json: Unchanged (already up to date)
```

Your existing credentials and settings are preserved while new options are added.

## Event Aliases

The `aliases.json` file maps event abbreviations to their full names for improved search matching. This helps when:
- Using lowercase abbreviations in filenames (e.g., `asot` instead of `ASOT`)
- Using unofficial abbreviations that can't be dynamically detected (e.g., `TML` for Tomorrowland)

```json
{
  "AMF": "Amsterdam Music Festival",
  "ADE": "Amsterdam Dance Event",
  "EDC": "Electric Daisy Carnival",
  "UMF": "Ultra Music Festival",
  "ASOT": "A State of Trance",
  "ABGT": "Above & Beyond Group Therapy",
  "WAO138": "Who's Afraid of 138",
  "FSOE": "Future Sound of Egypt",
  "GDJB": "Global DJ Broadcast",
  "SW4": "South West Four",
  "TML": "Tomorrowland",
  "TL": "Tomorrowland",
  "DWP": "Djakarta Warehouse Project",
  "MMW": "Miami Music Week"
}
```

Alias matching is case-insensitive. When a filename contains an alias key (e.g., `asot`), results containing the target name (e.g., "A State of Trance") receive a +35 score boost.

Add your own abbreviations for regional events or personal naming conventions.

## Behavior Notes

### Fast In-Place Editing

The script uses `mkvpropedit` to modify chapters and tags directly in the file without remuxing. This makes chapter embedding nearly instantaneous regardless of file size:

| File Size | Time |
|-----------|------|
| 1 GB | < 1 second |
| 10 GB | < 1 second |
| 50 GB | < 1 second |

When not using `-ReplaceOriginal`, the file is first copied to the output location, then modified in place.

### Session Caching

Login cookies are cached in `.1001tl-cookies.json` in the script directory. This speeds up consecutive runs by avoiding repeated logins. Cache is valid until cookies expire (~100 days).

### Rate Limit Protection

1001Tracklists.com has rate limits to prevent abuse. The script includes several protections:

- **Automatic delay**: A configurable delay (default: 5 seconds) is added between files in pipeline mode
- **Detection**: Rate limit responses are detected and reported clearly
- **Auto-recovery**: When rate limited, the cookie cache is automatically deleted so the next run starts fresh after you solve the captcha

If you encounter a rate limit:
1. Visit 1001Tracklists.com in your browser
2. Solve the captcha
3. Re-run the script (it will create a fresh login session)

### Duplicate Chapter Detection

Before modifying, the script extracts existing chapters using `mkvextract` and compares them with the new chapters. If they're identical **and** the tracklist URL is already stored, the file is skipped:
```
Chapters already exist and are identical - skipping.
```

If chapters are identical but no URL is stored yet (e.g., files processed before this feature), the file will be updated to add the URL:
```
Chapters identical, but adding tracklist URL to file...
```

### Tracklist URL Storage

When chapters are added from 1001Tracklists.com, the tracklist URL and title are stored as MKV global tags in the output file. URLs are stored in short format (e.g., `https://www.1001tracklists.com/tracklist/2ft171s9/`).

On subsequent runs:

- **Interactive mode**: Prompts with the stored tracklist info:
  ```
  Found stored tracklist:
    Martin Garrix @ Mainstage, Tomorrowland Weekend 2, Belgium 2025
    https://www.1001tracklists.com/tracklist/2ft171s9/

    Y = Use this tracklist  |  S = Skip file  |  R = Retry search
  Use this tracklist? (Y/s/r)
  ```
  Press Enter or `Y` to reuse, `S` to skip the file, or `R` to start a new search.

- **AutoSelect mode**: Uses the stored URL directly without searching, significantly speeding up batch re-processing.

Use `-IgnoreStoredUrl` to force a fresh search even when a stored URL exists.

### Tracklists Without Timestamps

If a selected tracklist has no timestamps (common for recently added tracklists), you'll see:
```
VERBOSE: Tracklist analysis: 0 timestamped lines, 22 numbered tracks
No timestamps available - skipping.
```

The tracklist exists but users haven't added timestamps yet on 1001Tracklists.com. In interactive mode, you'll be prompted to select a different tracklist. In `-AutoSelect` mode, the file is skipped.

### Filename to Query Conversion

When using filename-based search, the script:
1. Removes the file extension
2. Removes YouTube video IDs (11-character codes in brackets at the end)
3. Replaces `-`, `_`, `.` with spaces
4. Normalizes multiple spaces

## Troubleshooting

**"No tracklists found"**
- Try a different search query
- Use `-NoDurationFilter` to see all results regardless of duration
- Check if the tracklist exists on 1001Tracklists.com

**"Rate limited by 1001Tracklists"**
- Visit 1001Tracklists.com in your browser and solve the captcha
- The script automatically deletes the cookie cache when rate limited
- Re-run the script to create a fresh session
- Consider increasing `-DelaySeconds` for batch processing

**"This tracklist has no timestamps yet"**
- Some recent tracklists don't have timestamps added by users yet
- Select a different tracklist or wait for the community to add timestamps

**"Chapters already exist and are identical - skipping"**
- The file already has the same chapters - no action needed
- Use a different tracklist if you want different chapters

**mkvmerge/mkvpropedit not found**
- Install [MKVToolNix](https://mkvtoolnix.download/)
- Or specify the path: `-MkvMergePath "C:\Program Files\MKVToolNix\mkvmerge.exe"`

**Wrong tracklist selected in AutoSelect mode**
- Use `-Verbose` to see scoring breakdown
- Try adding more specific keywords to the filename or `-Tracklist` parameter
- Include the year, event abbreviation, or weekend/day number for better matching

**Cookie/session errors**
- Delete `.1001tl-cookies.json` in the script directory to force a fresh login
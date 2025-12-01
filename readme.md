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
     "NoDurationFilter": false
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
- **Event patterns**: `WE2`, `W2`, `Weekend2` all recognized as "Weekend 2"
- **Abbreviations**: Uppercase words like `AMF`, `EDC`, `ASOT` matched against full event names
- **Aliases**: Any word matching a key in aliases.json (case-insensitive) boosts results containing the target event name

### Examples

**Event pattern matching (WE2 → Weekend 2):**
```
Query: "2025 - Tomorrowland Belgium - Martin Garrix WE2"
QueryParts: Year=2025, Keywords=[tomorrowland, belgium, martin, garrix, we2], EventPatterns=[Weekend2]

Score 219.2 for 'Martin Garrix @ Mainstage, Tomorrowland Weekend 2...'
      [Dur:100(1m diff) | Kw:48(4/5) | Year:25 | Pat:40 | Rec:6.2]   ← Correct weekend, high score

Score 124.9 for 'Cence Brothers @ House Of Fortune Stage, Tomorrowland Weekend 1...'
      [Dur:100(1m diff) | Kw:24(2/5) | Year:25 | Pat:-30 | Rec:5.9]  ← Wrong weekend, penalized
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
| `-NoDurationFilter` | Disable duration-based result filtering |
| `-Preview` | Show chapters without modifying file |
| `-ReplaceOriginal` | Replace original file instead of creating new |
| `-OutputFile` | Custom output filename |
| `-ChapterLanguage` | Chapter language code (default: `eng`) |
| `-MkvMergePath` | Custom path to mkvmerge.exe |
| `-Email` | 1001Tracklists.com email (overrides config) |
| `-Password` | 1001Tracklists.com password (overrides config) |
| `-Credential` | PSCredential object (alternative to Email/Password) |
| `-CreateConfig` | Generate default config.json and aliases.json files |
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
<#
.SYNOPSIS
    Adds or replaces chapters in an MKV or WEBM file using mkvmerge.

.DESCRIPTION
    Extracts chapter information from a provided tracklist file, clipboard,
    or directly from 1001Tracklists.com, converts it to Matroska XML format, and uses
    mkvmerge to embed chapters in the media file.
    
    When no tracklist source is specified, the input filename is used as a search query
    on 1001Tracklists.com.
    
    Configuration is loaded from config.json in the script directory if present.
    Event aliases for search boosting are loaded from aliases.json if present.
    Command-line parameters override config values.
    
    Login cookies are cached in .1001tl-cookies.json to speed up consecutive runs.
    Cache is valid until the cookies expire (typically ~100 days).

.PARAMETER InputFile
    The MKV or WEBM file to which chapters will be added. Accepts pipeline input.
    If no tracklist source is specified, the filename is used as the search query
    on 1001Tracklists.com.

.PARAMETER Tracklist
    Tracklist source from 1001Tracklists.com. Auto-detects the type:
    - URL: If it contains "1001tracklists.com", fetches directly from that URL
    - ID: If it matches a short alphanumeric pattern (e.g., "1g6g22ut"), fetches by ID
    - Search query: Otherwise, searches 1001Tracklists.com for matching tracklists

.PARAMETER TrackListFile
    Text file containing chapter timestamps and titles.

.PARAMETER FromClipboard
    Read tracklist directly from the clipboard. Copy tracklist from browser, then run with this flag.

.PARAMETER Email
    Email address for 1001Tracklists.com account. Required for search and export API.

.PARAMETER Password
    Password for 1001Tracklists.com account. Required for search and export API.

.PARAMETER Credential
    PSCredential object containing login credentials. Alternative to Email/Password.

.PARAMETER NoDurationFilter
    Disable automatic filtering of search results by video duration. By default, search results
    are filtered to show only tracklists with a similar duration to the input video.

.PARAMETER AutoSelect
    Automatically select the top search result without prompting. Useful for automation and batch processing.

.PARAMETER OutputFile
    Custom output filename. Defaults to original filename with '-new' suffix.

.PARAMETER MkvMergePath
    Path to mkvmerge.exe. Auto-detected from PATH or default installation if not specified.

.PARAMETER ChapterLanguage
    Language code for chapters. Defaults to 'eng'.

.PARAMETER ReplaceOriginal
    Replace the original file instead of creating a new one.

.PARAMETER Preview
    Parse and display chapters without processing.

.PARAMETER CreateConfig
    Generate default config.json and aliases.json files in the script directory and exit.

.EXAMPLE
    .\Add-TracklistChapters.ps1 -InputFile "2025 - AMF - Sub Zero Project.webm"
    
    Searches 1001Tracklists.com using the filename as query (credentials from config.json).

.EXAMPLE
    .\Add-TracklistChapters.ps1 -InputFile "video.mkv" -Tracklist "Sub Zero Project AMF 2025"
    
    Searches 1001Tracklists.com with the specified query.

.EXAMPLE
    .\Add-TracklistChapters.ps1 -InputFile "video.mkv" -Tracklist "https://www.1001tracklists.com/tracklist/1g6g22ut/sub-zero-project.html"
    
    Fetches tracklist directly from the URL.

.EXAMPLE
    .\Add-TracklistChapters.ps1 -InputFile "video.mkv" -Tracklist "1g6g22ut"
    
    Fetches tracklist by ID.

.EXAMPLE
    .\Add-TracklistChapters.ps1 -InputFile "video.mkv" -TrackListFile "chapters.txt"

.EXAMPLE
    .\Add-TracklistChapters.ps1 -InputFile "video.mkv" -FromClipboard

.EXAMPLE
    .\Add-TracklistChapters.ps1 -CreateConfig
    
    Creates a default config.json file for storing credentials and preferences.

.EXAMPLE
    .\Add-TracklistChapters.ps1 -InputFile "video.mkv" -Preview

.EXAMPLE
    Get-ChildItem *.mkv | .\Add-TracklistChapters.ps1 -AutoSelect -ReplaceOriginal
    
    Batch process all MKV files, auto-selecting tracklists and replacing originals.

.EXAMPLE
    Get-ChildItem *.webm | .\Add-TracklistChapters.ps1 -Tracklist "Festival Name 2025"
    
    Batch process all WEBM files using the same search query.
#>

[CmdletBinding(DefaultParameterSetName = 'Default')]
param (
    [Parameter(Mandatory, ParameterSetName = 'Default', Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Parameter(Mandatory, ParameterSetName = 'Tracklist', Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Parameter(Mandatory, ParameterSetName = 'File', Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Parameter(Mandatory, ParameterSetName = 'Clipboard', Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [Alias('FullName')]
    [string]$InputFile,

    [Parameter(Mandatory, ParameterSetName = 'Tracklist', Position = 1)]
    [string]$Tracklist,

    [Parameter(Mandatory, ParameterSetName = 'File')]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$TrackListFile,

    [Parameter(Mandatory, ParameterSetName = 'Clipboard')]
    [switch]$FromClipboard,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'Tracklist')]
    [string]$Email,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'Tracklist')]
    [string]$Password,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'Tracklist')]
    [PSCredential]$Credential,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'Tracklist')]
    [switch]$NoDurationFilter,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'Tracklist')]
    [switch]$AutoSelect,

    [string]$OutputFile,

    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$MkvMergePath,

    [ValidatePattern('^[a-z]{2,3}$')]
    [string]$ChapterLanguage,

    [switch]$ReplaceOriginal,

    [switch]$Preview,

    [Parameter(Mandatory, ParameterSetName = 'CreateConfig')]
    [switch]$CreateConfig
)

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    # Determine script directory for config and aliases files
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
    $configPath = Join-Path $scriptDir 'config.json'
    $aliasesPath = Join-Path $scriptDir 'aliases.json'

    #region Config Functions

    function New-DefaultConfig {
        <#
        .SYNOPSIS
            Creates a default configuration object.
        #>
        return [ordered]@{
            Email            = ''
            Password         = ''
            ChapterLanguage  = 'eng'
            MkvMergePath     = ''
            ReplaceOriginal  = $false
            NoDurationFilter = $false
        }
    }

    function Get-Config {
        <#
        .SYNOPSIS
            Loads configuration from config.json if it exists.
        #>
        param(
            [Parameter(Mandatory)]
            [string]$Path
        )

        if (Test-Path $Path) {
            try {
                $jsonObj = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
                # Convert PSCustomObject to hashtable (PS5 compatibility)
                $config = @{}
                $jsonObj.PSObject.Properties | ForEach-Object { $config[$_.Name] = $_.Value }
                Write-Verbose "Loaded configuration from: $Path"
                return $config
            }
            catch {
                Write-Warning "Failed to load config from $Path`: $_"
            }
        }

        return @{}
    }

    function Save-Config {
        <#
        .SYNOPSIS
            Saves configuration to config.json.
        #>
        param(
            [Parameter(Mandatory)]
            [string]$Path,

            [Parameter(Mandatory)]
            [hashtable]$Config
        )

        $Config | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $Path -Encoding UTF8
        Write-Host "Configuration saved to: $Path" -ForegroundColor Green
    }

    function New-DefaultAliases {
        <#
        .SYNOPSIS
            Creates default event aliases for abbreviation matching.
        .DESCRIPTION
            Returns a hashtable mapping common event abbreviations to their full names.
            Used as fallback when dynamic abbreviation detection fails (e.g., lowercase
            abbreviations in filenames, or unofficial abbreviations like TML).
        #>
        return [ordered]@{
            AMF     = 'Amsterdam Music Festival'
            ADE     = 'Amsterdam Dance Event'
            EDC     = 'Electric Daisy Carnival'
            UMF     = 'Ultra Music Festival'
            ASOT    = 'A State of Trance'
            ABGT    = 'Above & Beyond Group Therapy'
            WAO138  = "Who's Afraid of 138"
            FSOE    = 'Future Sound of Egypt'
            GDJB    = 'Global DJ Broadcast'
            SW4     = 'South West Four'
            TML     = 'Tomorrowland'
            TL      = 'Tomorrowland'
            DWP     = 'Djakarta Warehouse Project'
            MMW     = 'Miami Music Week'
        }
    }

    function Get-Aliases {
        <#
        .SYNOPSIS
            Loads event aliases from aliases.json if it exists.
        .DESCRIPTION
            Returns a hashtable with lowercase keys for case-insensitive lookup.
        #>
        param(
            [Parameter(Mandatory)]
            [string]$Path
        )

        if (Test-Path $Path) {
            try {
                $jsonObj = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
                # Convert to hashtable with lowercase keys for case-insensitive lookup
                $aliases = @{}
                $jsonObj.PSObject.Properties | ForEach-Object { 
                    $aliases[$_.Name.ToLower()] = $_.Value 
                }
                Write-Verbose "Loaded $($aliases.Count) aliases from: $Path"
                return $aliases
            }
            catch {
                Write-Warning "Failed to load aliases from $Path`: $_"
            }
        }

        return @{}
    }

    function ConvertTo-SearchQuery {
        <#
        .SYNOPSIS
            Converts a filename to a search query by removing extension and special characters.
        #>
        param(
            [Parameter(Mandatory)]
            [string]$FileName
        )

        # Remove extension
        $name = [System.IO.Path]::GetFileNameWithoutExtension($FileName)

        # Replace common separators with spaces
        $name = $name -replace '[-_.]', ' '

        # Remove multiple spaces
        $name = $name -replace '\s+', ' '

        return $name.Trim()
    }

    function Get-TracklistType {
        <#
        .SYNOPSIS
            Detects the type of tracklist source from the input string.
        .OUTPUTS
            Returns a hashtable with Type ('Url', 'Id', or 'Search') and Value.
        #>
        param(
            [Parameter(Mandatory)]
            [string]$TracklistInput
        )

        # Check if it's a URL
        if ($TracklistInput -match '1001tracklists\.com') {
            return @{ Type = 'Url'; Value = $TracklistInput }
        }

        # Check if it's an ID (short alphanumeric, typically 6-10 characters, no spaces)
        if ($TracklistInput -match '^[a-z0-9]{6,12}$') {
            return @{ Type = 'Id'; Value = $TracklistInput }
        }

        # Otherwise treat as search query
        return @{ Type = 'Search'; Value = $TracklistInput }
    }

    #endregion Config Functions

    # Handle CreateConfig parameter set
    if ($PSCmdlet.ParameterSetName -eq 'CreateConfig') {
        $defaultConfig = New-DefaultConfig
        $defaultAliases = New-DefaultAliases

        # Handle config.json
        $writeConfig = $true
        if (Test-Path $configPath) {
            Write-Warning "config.json already exists: $configPath"
            $overwrite = Read-Host "Overwrite? (y/N)"
            $writeConfig = $overwrite -match '^[Yy]'
        }

        # Handle aliases.json
        $writeAliases = $true
        if (Test-Path $aliasesPath) {
            Write-Warning "aliases.json already exists: $aliasesPath"
            $overwrite = Read-Host "Overwrite? (y/N)"
            $writeAliases = $overwrite -match '^[Yy]'
        }

        # Write files as needed
        if ($writeConfig) {
            Save-Config -Path $configPath -Config $defaultConfig
        }
        else {
            Write-Host "Skipped config.json" -ForegroundColor Yellow
        }

        if ($writeAliases) {
            $defaultAliases | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $aliasesPath -Encoding UTF8
            Write-Host "Aliases saved to: $aliasesPath" -ForegroundColor Green
        }
        else {
            Write-Host "Skipped aliases.json" -ForegroundColor Yellow
        }

        # Show summary
        if ($writeConfig) {
            Write-Host "`nEdit config.json to add your 1001Tracklists.com credentials:" -ForegroundColor Cyan
            Write-Host $configPath
        }
        if ($writeAliases) {
            Write-Host "`nEdit aliases.json to add custom event abbreviations:" -ForegroundColor Cyan
            Write-Host $aliasesPath
        }
        return
    }

    # Load config and apply defaults (command-line parameters override config)
    $config = Get-Config -Path $configPath

    if (-not $PSBoundParameters.ContainsKey('Email') -and $config.Email) {
        $Email = $config.Email
    }
    if (-not $PSBoundParameters.ContainsKey('Password') -and $config.Password) {
        $Password = $config.Password
    }
    if (-not $PSBoundParameters.ContainsKey('ChapterLanguage')) {
        $ChapterLanguage = if ($config.ChapterLanguage) { $config.ChapterLanguage } else { 'eng' }
    }
    if (-not $PSBoundParameters.ContainsKey('MkvMergePath') -and $config.MkvMergePath) {
        $MkvMergePath = $config.MkvMergePath
    }
    if (-not $PSBoundParameters.ContainsKey('ReplaceOriginal') -and $config.ReplaceOriginal) {
        $ReplaceOriginal = [bool]$config.ReplaceOriginal
    }
    if (-not $PSBoundParameters.ContainsKey('NoDurationFilter') -and $config.NoDurationFilter) {
        $NoDurationFilter = [bool]$config.NoDurationFilter
    }

    # Load event aliases for abbreviation matching
    $script:EventAliases = Get-Aliases -Path $aliasesPath
    if ($script:EventAliases.Count -eq 0) {
        # Use defaults if no aliases file exists
        $defaultAliases = New-DefaultAliases
        $defaultAliases.Keys | ForEach-Object { $script:EventAliases[$_.ToLower()] = $defaultAliases[$_] }
        Write-Verbose "Using default aliases (no aliases.json found)"
    }

    #region 1001Tracklists Configuration
    $script:BaseUrl = 'https://www.1001tracklists.com'
    $script:UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:145.0) Gecko/20100101 Firefox/145.0'
    $script:Session = $null
    $script:CookieCachePath = Join-Path $scriptDir '.1001tl-cookies.json'
    #endregion

    #region 1001Tracklists Functions

    function Save-CookieCache {
        <#
        .SYNOPSIS
            Saves session cookies to cache file for reuse.
        #>
        param(
            [Parameter(Mandatory)]
            [string]$Email
        )

        if (-not $script:Session) { return }

        $cookies = $script:Session.Cookies.GetCookies($script:BaseUrl)
        $cookieData = @{
            Email     = $Email
            Timestamp = (Get-Date).ToString('o')
            Cookies   = @($cookies | ForEach-Object {
                @{
                    Name    = $_.Name
                    Value   = $_.Value
                    Domain  = $_.Domain
                    Path    = $_.Path
                    Expires = if ($_.Expires -ne [DateTime]::MinValue) { $_.Expires.ToString('o') } else { $null }
                }
            })
        }

        try {
            $cookieData | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $script:CookieCachePath -Encoding UTF8
            Write-Verbose "Saved cookies to cache."
        }
        catch {
            Write-Verbose "Failed to save cookie cache: $_"
        }
    }

    function Restore-CookieCache {
        <#
        .SYNOPSIS
            Restores session cookies from cache file.
        .OUTPUTS
            Returns $true if cookies were restored and are valid, $false otherwise.
        #>
        param(
            [Parameter(Mandatory)]
            [string]$Email
        )

        if (-not (Test-Path $script:CookieCachePath)) {
            Write-Verbose "No cookie cache found."
            return $false
        }

        try {
            $cacheContent = Get-Content -LiteralPath $script:CookieCachePath -Raw | ConvertFrom-Json

            # Check if cache is for the same email
            if ($cacheContent.Email -ne $Email) {
                Write-Verbose "Cookie cache is for different user, ignoring."
                return $false
            }

            # Check if cookies are expired (use earliest expiration from cookies)
            $now = Get-Date
            $expired = $false
            foreach ($c in $cacheContent.Cookies) {
                if ($c.Expires) {
                    $expiry = [DateTime]::Parse($c.Expires)
                    if ($expiry -lt $now) {
                        $expired = $true
                        break
                    }
                }
            }
            if ($expired) {
                Write-Verbose "Cookie cache expired."
                return $false
            }

            # Create session and restore cookies
            $script:Session = [Microsoft.PowerShell.Commands.WebRequestSession]::new()
            $script:Session.UserAgent = $script:UserAgent

            foreach ($c in $cacheContent.Cookies) {
                $cookie = [System.Net.Cookie]::new($c.Name, $c.Value, $c.Path, $c.Domain)
                if ($c.Expires) {
                    $cookie.Expires = [DateTime]::Parse($c.Expires)
                }
                $script:Session.Cookies.Add($cookie)
            }

            # Validate cookies by making a test request to profile page
            Write-Verbose "Validating cached cookies..."
            $testResponse = Invoke-WebRequest -Uri "$script:BaseUrl/my/" -WebSession $script:Session -UseBasicParsing -MaximumRedirection 20

            # Check if response indicates we're logged in (page contains logout link or profile content)
            # If redirected to login or content indicates not logged in, cookies are invalid
            $isLoggedIn = $testResponse.Content -match 'logout|/my/|signout' -and $testResponse.Content -notmatch 'login-form|Please log in'

            if (-not $isLoggedIn) {
                Write-Verbose "Cached cookies are invalid."
                $script:Session = $null
                return $false
            }

            # Double-check we have the expected cookies
            $cookies = $script:Session.Cookies.GetCookies($script:BaseUrl)
            $hasSid = $cookies | Where-Object { $_.Name -eq 'sid' }
            $hasUid = $cookies | Where-Object { $_.Name -eq 'uid' }

            if (-not $hasSid -or -not $hasUid) {
                Write-Verbose "Cached cookies missing required values."
                $script:Session = $null
                return $false
            }

            Write-Verbose "Using cached session."
            return $true
        }
        catch {
            Write-Verbose "Failed to restore cookie cache: $_"
            $script:Session = $null
            return $false
        }
    }

    function Initialize-1001Session {
        [CmdletBinding()]
        param(
            [string]$UserEmail,
            [string]$UserPassword
        )

        # Try to restore from cache first
        if ($UserEmail -and $UserPassword) {
            if (Restore-CookieCache -Email $UserEmail) {
                return
            }
        }

        # Fresh session
        $script:Session = [Microsoft.PowerShell.Commands.WebRequestSession]::new()
        $script:Session.UserAgent = $script:UserAgent

        Write-Verbose "Initializing 1001Tracklists session..."
        $null = Invoke-WebRequest -Uri $script:BaseUrl -WebSession $script:Session -UseBasicParsing -MaximumRedirection 20

        if ($UserEmail -and $UserPassword) {
            Write-Verbose "Logging in as $UserEmail..."
            $loginBody = @{
                email    = $UserEmail
                password = $UserPassword
                referer  = "$script:BaseUrl/"
            }

            $loginParams = @{
                Uri                = "$script:BaseUrl/action/login.html"
                Method             = 'POST'
                Body               = $loginBody
                WebSession         = $script:Session
                UseBasicParsing    = $true
                MaximumRedirection = 5
            }

            # Login may redirect on success - let it complete, then verify via cookies
            $null = Invoke-WebRequest @loginParams

            $cookies = $script:Session.Cookies.GetCookies($script:BaseUrl)
            $hasSid = $cookies | Where-Object { $_.Name -eq 'sid' }
            $hasUid = $cookies | Where-Object { $_.Name -eq 'uid' }

            if (-not $hasSid -or -not $hasUid) {
                throw "Login failed. Please check your credentials."
            }

            Write-Verbose "Login successful."

            # Save cookies for next run
            Save-CookieCache -Email $UserEmail
        }
    }

    function Invoke-1001Request {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Uri,

            [string]$Method = 'GET',
            [hashtable]$Body,
            [hashtable]$Headers,
            [string]$ContentType,
            [int]$MaximumRedirection = 20
        )

        if (-not $script:Session) {
            throw "Session not initialized. Call Initialize-1001Session first."
        }

        $params = @{
            Uri                = $Uri
            Method             = $Method
            WebSession         = $script:Session
            UseBasicParsing    = $true
            MaximumRedirection = $MaximumRedirection
        }

        if ($Body) { $params.Body = $Body }
        if ($Headers) { $params.Headers = $Headers }
        if ($ContentType) { $params.ContentType = $ContentType }

        Invoke-WebRequest @params
    }

    function Search-1001Tracklists {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Query,

            [int]$DurationMinutes = 0,

            [string]$Year
        )

        Write-Verbose "Searching for: $Query"

        $searchBody = @{
            main_search      = $Query
            search_selection = '9'
            filterObject     = '9'
            orderby          = 'added'
        }

        # Add duration filter if specified (server-side filtering)
        # Use duration - 3 to catch more results since tracklist duration may differ slightly
        if ($DurationMinutes -gt 0) {
            $serverDuration = [Math]::Max(1, $DurationMinutes - 3)
            $searchBody.duration = $serverDuration.ToString()
        }

        # Add year filter if specified (server-side filtering)
        if ($Year) {
            $searchBody.startDate = "$Year-01-01"
            $searchBody.endDate = "$Year-12-31"
            Write-Verbose "Filtering by year: $Year"
        }

        $response = Invoke-1001Request -Uri "$script:BaseUrl/search/result.php" -Method 'POST' -Body $searchBody

        if (-not $response -or -not $response.Content) {
            Write-Warning "Search returned no content. Authentication may be required."
            return @()
        }

        $results = [System.Collections.Generic.List[PSCustomObject]]::new()

        # Split HTML by bItm class to process each result item
        $items = $response.Content -split 'class="bItm(?:\s|")'

        $tracklistPattern = '<a href="(/tracklist/([^/]+)/[^"]+)"[^>]*>([^<]+)</a>'
        # Updated pattern: matches "1h 15m", "58m", or just "1h"
        $durationPattern = 'title="play time"[^>]*>.*?</i>((?:\d+h\s*)?(?:\d+m)?)\s*</div>'
        $datePattern = 'title="tracklist date"[^>]*>.*?</i>([^<]+)</div>'

        $seen = @{}
        foreach ($item in $items[1..($items.Count - 1)]) {
            $tlMatch = [regex]::Match($item, $tracklistPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            
            if (-not $tlMatch.Success) {
                continue
            }

            $url = $tlMatch.Groups[1].Value
            $id = $tlMatch.Groups[2].Value
            $title = [System.Web.HttpUtility]::HtmlDecode($tlMatch.Groups[3].Value.Trim())

            if ($seen.ContainsKey($id) -or $title -match '^(Previous|Next|First|Last)$' -or [string]::IsNullOrWhiteSpace($title)) {
                continue
            }
            $seen[$id] = $true

            # Extract duration from the item
            $durMatch = [regex]::Match($item, $durationPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline)
            $duration = if ($durMatch.Success -and $durMatch.Groups[1].Value.Trim()) { 
                $durMatch.Groups[1].Value.Trim() 
            } else { 
                $null 
            }

            # Parse duration to minutes for sorting
            $durationMins = ConvertTo-Minutes -Duration $duration

            # Extract date from the item
            $dateMatch = [regex]::Match($item, $datePattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline)
            $date = if ($dateMatch.Success) { $dateMatch.Groups[1].Value.Trim() } else { $null }

            $results.Add([PSCustomObject]@{
                Index               = 0      # Will be set after sorting
                Id                  = $id
                Title               = $title
                Url                 = "$script:BaseUrl$url"
                Duration            = $duration
                DurationMins        = $durationMins
                Date                = $date
                Score               = 0      # Will be calculated
                MatchedKeywordCount = 0      # Will be calculated
                HasEventMatch       = $false # Will be calculated
            })
        }

        # Calculate relevance scores and sort
        if ($results.Count -gt 0) {
            # Parse query for keywords and year
            $queryParts = Get-QueryParts -Query $Query
            
            $aliasesInfo = if ($queryParts.ResolvedAliases.Count -gt 0) {
                $queryParts.ResolvedAliases | ForEach-Object { "$($_.Alias)->$($_.Target)" }
            } else { @() }
            Write-Verbose "Query parts: Year=$($queryParts.Year), Keywords=[$($queryParts.Keywords -join ', ')], Abbreviations=[$($queryParts.Abbreviations -join ', ')], Aliases=[$($aliasesInfo -join ', ')], EventPatterns=[$($queryParts.EventPatterns | ForEach-Object { "$($_.Type)$($_.Number)" })]"
            
            # Find date range for recency scoring
            $dates = $results | Where-Object { $_.Date } | ForEach-Object { 
                try { [datetime]::Parse($_.Date) } catch { $null } 
            } | Where-Object { $_ }
            
            $minDate = if ($dates) { ($dates | Measure-Object -Minimum).Minimum } else { $null }
            $maxDate = if ($dates) { ($dates | Measure-Object -Maximum).Maximum } else { $null }
            $dateRange = if ($minDate -and $maxDate -and $maxDate -ne $minDate) { 
                ($maxDate - $minDate).TotalDays 
            } else { 
                1 
            }

            foreach ($result in $results) {
                $scoreInfo = Get-RelevanceScore -Result $result -QueryParts $queryParts -VideoDurationMinutes $DurationMinutes -MinDate $minDate -DateRange $dateRange
                $result.Score = $scoreInfo.Score
                $result.MatchedKeywordCount = $scoreInfo.MatchedKeywordCount
                $result.HasEventMatch = $scoreInfo.HasEventMatch
            }

            # Always filter out results with 0 keyword matches (pure noise from duration/year)
            $beforeCount = $results.Count
            $results = [System.Collections.Generic.List[PSCustomObject]]@(
                $results | Where-Object { $_.MatchedKeywordCount -gt 0 }
            )
            $zeroKwFiltered = $beforeCount - $results.Count
            if ($zeroKwFiltered -gt 0) {
                Write-Verbose "Filtered out $zeroKwFiltered results with 0 keyword matches"
            }

            # Filter out low-quality results when an event is specified
            # If query has aliases or abbreviations, remove results with ≤1 keyword match and no event match
            $hasEventInQuery = ($queryParts.ResolvedAliases.Count -gt 0) -or ($queryParts.Abbreviations.Count -gt 0)
            if ($hasEventInQuery) {
                $beforeCount = $results.Count
                $results = [System.Collections.Generic.List[PSCustomObject]]@(
                    $results | Where-Object { $_.HasEventMatch -or $_.MatchedKeywordCount -gt 1 }
                )
                $filteredCount = $beforeCount - $results.Count
                if ($filteredCount -gt 0) {
                    Write-Verbose "Filtered out $filteredCount low-relevance results (no event match, ≤1 keyword)"
                }
            }

            # Sort by score descending
            $sortedResults = $results | Sort-Object -Property Score -Descending
            $results = [System.Collections.Generic.List[PSCustomObject]]::new()
            foreach ($r in $sortedResults) { $results.Add($r) }
        }

        # Re-index after sorting
        for ($i = 0; $i -lt $results.Count; $i++) {
            $results[$i].Index = $i + 1
        }

        # Return as array to ensure .Count works even with 0 or 1 result
        return @($results)
    }

    function ConvertTo-Minutes {
        <#
        .SYNOPSIS
            Converts a duration string like "1h 15m", "58m", or "1h" to total minutes.
        #>
        param(
            [string]$Duration
        )

        if (-not $Duration) { return $null }

        $totalMinutes = 0

        if ($Duration -match '(\d+)h') {
            $totalMinutes += [int]$Matches[1] * 60
        }

        if ($Duration -match '(\d+)m') {
            $totalMinutes += [int]$Matches[1]
        }

        if ($totalMinutes -eq 0) { return $null }

        return $totalMinutes
    }

    function Get-Abbreviation {
        <#
        .SYNOPSIS
            Extracts abbreviation from a multi-word string by taking first letter of each capitalized word.
        .DESCRIPTION
            Matches words starting with uppercase letters and combines their initials.
            Requires at least 2 matching words to return an abbreviation.
        .EXAMPLE
            Get-Abbreviation "Amsterdam Music Festival"  # Returns "AMF"
            Get-Abbreviation "A State Of Trance"         # Returns "ASOT"
            Get-Abbreviation "Electric Daisy Carnival"   # Returns "EDC"
        #>
        param(
            [Parameter(Mandatory)]
            [string]$Text
        )

        # Match words starting with uppercase letter (captures the first letter)
        $words = [regex]::Matches($Text, '\b([A-Z])[a-z]*')
        if ($words.Count -ge 2) {
            return ($words | ForEach-Object { $_.Groups[1].Value }) -join ''
        }
        return $null
    }

    function Remove-Diacritics {
        <#
        .SYNOPSIS
            Removes diacritical marks (accents) from text.
        .DESCRIPTION
            Normalizes text to decomposed form and removes combining marks,
            converting accented characters to their base form.
        .EXAMPLE
            Remove-Diacritics "Château"  # Returns "Chateau"
            Remove-Diacritics "Ibañez"   # Returns "Ibanez"
        #>
        param(
            [Parameter(Mandatory)]
            [string]$Text
        )

        $normalized = $Text.Normalize([System.Text.NormalizationForm]::FormD)
        return ($normalized -replace '\p{M}', '')
    }

    function Get-QueryParts {
        <#
        .SYNOPSIS
            Parses a search query into year, keywords, potential abbreviations, and event patterns.
        .DESCRIPTION
            Detects special patterns commonly used in filenames:
            - WE1, WE2, W1, W2 -> Weekend 1, Weekend 2
            - D1, D2 -> Day 1, Day 2
            Also resolves event aliases from aliases.json (case-insensitive).
        #>
        param(
            [Parameter(Mandatory)]
            [string]$Query
        )

        # Remove YouTube video ID at end of query (yt-dlp format: [xXxXxXxXxXx])
        $Query = $Query -replace '\s*\[[A-Za-z0-9_-]{11}\]\s*$', ''

        $words = $Query -split '\s+' | Where-Object { $_.Length -gt 0 }
        $year = $null
        $keywords = @()
        $abbreviations = @()
        $eventPatterns = @()
        $resolvedAliases = @()

        foreach ($word in $words) {
            if ($word -match '^(19|20)\d{2}$') {
                $year = $word
            }
            else {
                # Detect weekend patterns: WE1, WE2, W1, W2, Weekend1, etc. (case-insensitive)
                if ($word -match '(?i)^(?:WE|W|Weekend)(\d+)$') {
                    $eventPatterns += @{ Type = 'Weekend'; Number = $Matches[1] }
                }
                # Detect day patterns: D1, D2, Day1, etc. (case-insensitive)
                elseif ($word -match '(?i)^(?:D|Day)(\d+)$') {
                    $eventPatterns += @{ Type = 'Day'; Number = $Matches[1] }
                }
                # Check if this word looks like an abbreviation (2+ uppercase letters)
                elseif ($word -cmatch '^[A-Z]{2,}$') {
                    $abbreviations += $word
                }
                
                # Check if word matches an alias (case-insensitive)
                $wordLower = $word.ToLower()
                if ($script:EventAliases.ContainsKey($wordLower)) {
                    $resolvedAliases += @{ 
                        Alias  = $word
                        Target = $script:EventAliases[$wordLower] 
                    }
                }
                
                # Filter out very short words (at, the, in, etc.) but keep pattern words as keywords too
                if ($word.Length -gt 2) {
                    $keywords += $wordLower
                }
            }
        }

        return @{
            Year            = $year
            Keywords        = $keywords
            Abbreviations   = $abbreviations
            EventPatterns   = $eventPatterns
            ResolvedAliases = $resolvedAliases
        }
    }

    function Get-RelevanceScore {
        <#
        .SYNOPSIS
            Calculates a relevance score for a search result.
        .DESCRIPTION
            Scoring is designed to find the exact same recording:
            - Duration match is the strongest signal (same recording = same length)
            - Event patterns (WE1/Weekend 1) distinguish multi-day events
            - Keyword/abbreviation matching identifies the correct event
            - Year matching ensures correct edition of recurring events
            - Recency is a minor tiebreaker

            Score ranges:
            - Perfect match: ~270 points
            - Good match: 180-240 points
            - Partial match: 100-180 points
            - Poor match: <100 points (may include negative scores for bad duration/pattern mismatch)
        #>
        param(
            [Parameter(Mandatory)]
            [PSCustomObject]$Result,

            [Parameter(Mandatory)]
            [hashtable]$QueryParts,

            [int]$VideoDurationMinutes = 0,

            [datetime]$MinDate,

            [double]$DateRange
        )

        $score = 0
        $breakdown = @()
        $matchedCount = 0

        # Duration score (most important - indicates same recording)
        # Exact match is strong positive, large mismatch is penalized
        $durationScore = 0
        if ($VideoDurationMinutes -gt 0 -and $null -ne $Result.DurationMins) {
            $diff = [Math]::Abs($Result.DurationMins - $VideoDurationMinutes)
            if ($diff -le 1) {
                $durationScore = 100  # Exact match - almost certainly the same recording
            }
            elseif ($diff -le 5) {
                $durationScore = 80   # Close match - likely same recording, minor variance
            }
            elseif ($diff -le 15) {
                $durationScore = 40   # Moderate difference - possibly edited/partial recording
            }
            elseif ($diff -le 30) {
                $durationScore = 10   # Significant difference - unlikely same recording
            }
            else {
                $durationScore = -20   # Very different duration - penalize
            }
            $score += $durationScore
            $breakdown += "Dur:$durationScore(${diff}m diff)"
        }
        elseif ($VideoDurationMinutes -gt 0 -and $null -eq $Result.DurationMins) {
            $breakdown += "Dur:0(no data)"
        }

        # Abbreviation detection (do this first so we can credit keywords)
        $abbrScore = 0
        $matchedAbbreviations = @()  # Track which abbreviations matched
        if ($QueryParts.Abbreviations.Count -gt 0) {
            $segments = $Result.Title -split '[@,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
            
            foreach ($abbrev in $QueryParts.Abbreviations) {
                $matched = $false
                $matchInfo = $null
                
                # Check if abbreviation appears directly in title
                if ($Result.Title -match "\b$abbrev\b") {
                    $matched = $true
                    $matchInfo = "$abbrev(direct)"
                }
                else {
                    # Try to match abbreviation against each title segment
                    foreach ($segment in $segments) {
                        $segmentAbbrev = Get-Abbreviation -Text $segment
                        if ($segmentAbbrev -and $segmentAbbrev -eq $abbrev) {
                            $matched = $true
                            $matchInfo = "$abbrev=$segmentAbbrev"
                            break
                        }
                    }
                }
                
                if ($matched) {
                    $abbrScore += 35
                    $matchedAbbreviations += @{ Abbrev = $abbrev; Info = $matchInfo }
                }
            }
            $score += $abbrScore
            if ($matchedAbbreviations.Count -gt 0) {
                $breakdown += "Abbr:$abbrScore($($matchedAbbreviations.Info -join ','))"
            } else {
                $breakdown += "Abbr:0"
            }
        }

        # Alias matching (for lowercase abbreviations or unofficial abbreviations like TML)
        $aliasScore = 0
        $matchedAliases = @()
        if ($QueryParts.ResolvedAliases.Count -gt 0) {
            $titleNormalized = (Remove-Diacritics $Result.Title).ToLower()
            
            foreach ($alias in $QueryParts.ResolvedAliases) {
                $targetNormalized = (Remove-Diacritics $alias.Target).ToLower()
                
                # Check if target event name appears in title
                if ($titleNormalized -match [regex]::Escape($targetNormalized)) {
                    $aliasScore += 35
                    $matchedAliases += @{ Alias = $alias.Alias; Target = $alias.Target }
                }
            }
            $score += $aliasScore
            if ($matchedAliases.Count -gt 0) {
                $aliasInfo = $matchedAliases | ForEach-Object { "$($_.Alias)->$($_.Target)" }
                $breakdown += "Alias:$aliasScore($($aliasInfo -join ','))"
            }
        }

        # Keyword score (important - identifies the correct event/artist)
        # Matched abbreviations and aliases count as matched keywords
        $kwScore = 0
        if ($QueryParts.Keywords.Count -gt 0) {
            $titleNormalized = (Remove-Diacritics $Result.Title).ToLower()
            $matchedCount = 0
            $matchedKws = @()
            
            foreach ($keyword in $QueryParts.Keywords) {
                # Normalize keyword for accent-insensitive matching
                $keywordNormalized = (Remove-Diacritics $keyword).ToLower()
                
                # Check direct keyword match
                if ($titleNormalized -match [regex]::Escape($keywordNormalized)) {
                    $matchedCount++
                    $matchedKws += $keyword
                }
                # Check if keyword matches an abbreviation that was found
                elseif ($matchedAbbreviations | Where-Object { $_.Abbrev -eq $keyword.ToUpper() }) {
                    $matchedCount++
                    $matchedKws += "$keyword(abbr)"
                }
                # Check if keyword matches an alias that was resolved
                elseif ($matchedAliases | Where-Object { $_.Alias.ToLower() -eq $keywordNormalized }) {
                    $matchedCount++
                    $matchedKws += "$keyword(alias)"
                }
            }
            
            # Base score: proportional to match ratio (max 60)
            $kwScore = [Math]::Round(($matchedCount / $QueryParts.Keywords.Count) * 60, 1)
            $score += $kwScore
            
            # Bonus for exact match (all keywords found)
            if ($matchedCount -eq $QueryParts.Keywords.Count) {
                $score += 15
                $kwScore += 15
            }
            $breakdown += "Kw:$kwScore($matchedCount/$($QueryParts.Keywords.Count))"
        }

        # Year bonus
        $yearScore = 0
        if ($QueryParts.Year -and $Result.Date) {
            if ($Result.Date -match $QueryParts.Year) {
                $yearScore = 25
                $score += $yearScore
            }
            $breakdown += "Year:$yearScore"
        }

        # Event pattern score
        $patternScore = 0
        if ($QueryParts.EventPatterns.Count -gt 0) {
            $titleLower = $Result.Title.ToLower()
            foreach ($pattern in $QueryParts.EventPatterns) {
                $num = $pattern.Number
                $type = $pattern.Type
                
                if ($type -eq 'Weekend') {
                    $matchPattern = "(?:weekend\s*$num|w$num|we$num)"
                    $wrongPattern = "(?:weekend\s*[0-9]|w[0-9]|we[0-9])"
                }
                else {
                    $matchPattern = "(?:day\s*$num|d$num)"
                    $wrongPattern = "(?:day\s*[0-9]|d[0-9])"
                }
                
                if ($titleLower -match $matchPattern) {
                    $patternScore = 40
                }
                elseif ($titleLower -match $wrongPattern) {
                    $patternScore = -30
                }
            }
            $score += $patternScore
            $breakdown += "Pat:$patternScore"
        }

        # Recency bonus
        $recencyScore = 0
        if ($Result.Date -and $MinDate -and $DateRange -gt 0) {
            try {
                $resultDate = [datetime]::Parse($Result.Date)
                $daysFromMin = ($resultDate - $MinDate).TotalDays
                $recencyScore = [Math]::Round(($daysFromMin / $DateRange) * 10, 1)
                $score += $recencyScore
            }
            catch { }
        }
        $breakdown += "Rec:$recencyScore"

        $finalScore = [Math]::Round($score, 2)
        Write-Verbose "  Score $finalScore for '$($Result.Title.Substring(0, [Math]::Min(50, $Result.Title.Length)))...' [$($breakdown -join ' | ')]"
        
        return @{
            Score               = $finalScore
            MatchedKeywordCount = $matchedCount
            HasEventMatch       = ($matchedAbbreviations.Count -gt 0) -or ($matchedAliases.Count -gt 0)
        }
    }

    function Get-1001TracklistExport {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Id,

            [string]$FullUrl
        )

        Write-Verbose "Fetching tracklist: $Id"

        $initialUrl = if ($FullUrl) { $FullUrl } else { "$script:BaseUrl/tracklist/$Id/" }
        $pageResponse = Invoke-1001Request -Uri $initialUrl
        
        # Get actual URL after redirects (PS5/PS7 compatible)
        # Use safe property checking to avoid errors
        $actualUrl = $initialUrl  # Default fallback
        
        if ($pageResponse.BaseResponse) {
            if ($pageResponse.BaseResponse.PSObject.Properties['ResponseUri'] -and $pageResponse.BaseResponse.ResponseUri) {
                # PS5: ResponseUri property
                $actualUrl = $pageResponse.BaseResponse.ResponseUri.AbsoluteUri
            }
            elseif ($pageResponse.BaseResponse.PSObject.Properties['RequestMessage'] -and $pageResponse.BaseResponse.RequestMessage) {
                # PS7: RequestMessage.RequestUri property
                $actualUrl = $pageResponse.BaseResponse.RequestMessage.RequestUri.AbsoluteUri
            }
        }

        Write-Verbose "Tracklist URL: $actualUrl"

        $exportBody = @{
            object = 'tracklist'
            idTL   = $Id
        }

        $headers = @{
            'X-Requested-With' = 'XMLHttpRequest'
            'Referer'          = $actualUrl
            'Accept'           = 'application/json, text/javascript, */*; q=0.01'
            'Origin'           = $script:BaseUrl
        }

        $response = Invoke-1001Request -Uri "$script:BaseUrl/ajax/export_data.php" -Method 'POST' -Body $exportBody -Headers $headers -ContentType 'application/x-www-form-urlencoded; charset=UTF-8'

        $json = $response.Content | ConvertFrom-Json

        if (-not $json.success) {
            throw "Export failed: $($json.message)"
        }

        return $json.data
    }

    function Get-1001TracklistFromHtml {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Id,

            [string]$FullUrl
        )

        Write-Verbose "Parsing tracklist from HTML: $Id"

        $initialUrl = if ($FullUrl) { $FullUrl } else { "$script:BaseUrl/tracklist/$Id/" }
        $response = Invoke-1001Request -Uri $initialUrl
        $html = $response.Content

        $titleMatch = [regex]::Match($html, '<title>([^<]+)</title>')
        $title = if ($titleMatch.Success) {
            [System.Web.HttpUtility]::HtmlDecode($titleMatch.Groups[1].Value.Trim())
        }
        else {
            "Unknown Tracklist"
        }

        $tracks = [System.Collections.Generic.List[PSCustomObject]]::new()
        
        # Parse the cueValueData JavaScript to get timestamps in seconds
        # Format: cueValuesEntry.seconds = 12; cueValuesEntry.number = '1';
        $cueData = @{}
        $cueMatches = [regex]::Matches($html, "cueValuesEntry\.seconds\s*=\s*(\d+);\s*cueValuesEntry\.number\s*=\s*'(\d+)'")
        foreach ($match in $cueMatches) {
            $seconds = [int]$match.Groups[1].Value
            $number = $match.Groups[2].Value
            $cueData[$number] = $seconds
        }
        
        # Parse track items - look for tlpItem divs with track number and meta itemprop="name"
        # Pattern: div with class containing "tlpItem" and data-trno, then find meta itemprop="name"
        $trackPattern = '<div[^>]+class="[^"]*tlpItem[^"]*"[^>]+data-trno="(\d+)"[^>]*>[\s\S]*?<meta\s+itemprop="name"\s+content="([^"]+)"'
        $trackMatches = [regex]::Matches($html, $trackPattern)
        
        foreach ($match in $trackMatches) {
            $trNo = $match.Groups[1].Value
            $trackName = [System.Web.HttpUtility]::HtmlDecode($match.Groups[2].Value.Trim())
            
            # Skip "w/" (played with) entries - they share timestamp with previous track
            # Look for the tracknumber_value span to check if it's a "w/" entry
            $itemSection = $match.Value
            if ($itemSection -match 'title="played together with previous track"') {
                continue
            }
            
            # Get timestamp from cue data
            $timestamp = ''
            if ($cueData.ContainsKey(($tracks.Count + 1).ToString())) {
                $seconds = $cueData[($tracks.Count + 1).ToString()]
                $hours = [math]::Floor($seconds / 3600)
                $mins = [math]::Floor(($seconds % 3600) / 60)
                $secs = $seconds % 60
                if ($hours -gt 0) {
                    $timestamp = "{0}:{1:D2}:{2:D2}" -f $hours, $mins, $secs
                } else {
                    $timestamp = "{0}:{1:D2}" -f $mins, $secs
                }
            }
            
            if (-not [string]::IsNullOrWhiteSpace($trackName)) {
                $tracks.Add([PSCustomObject]@{
                    Position  = $tracks.Count + 1
                    Timestamp = $timestamp
                    Track     = $trackName
                })
            }
        }

        # Fallback: If the above didn't work, try a simpler approach using cue divs
        if ($tracks.Count -eq 0) {
            Write-Verbose "Primary parsing failed, trying fallback method..."
            
            # Find all meta itemprop="name" tags paired with cue timestamps
            $sections = $html -split '<div[^>]+class="[^"]*tlpItem'
            $trackNumber = 0
            
            foreach ($section in $sections[1..($sections.Count - 1)]) {
                # Skip mashup sub-positions (linked tracks)
                if ($section -match 'data-mashpos="true"') {
                    continue
                }
                
                # Skip "w/" entries
                if ($section -match 'title="played together with previous track"') {
                    continue
                }
                
                # Get track name from meta itemprop="name"
                $nameMatch = [regex]::Match($section, '<meta\s+itemprop="name"\s+content="([^"]+)"')
                if (-not $nameMatch.Success) { continue }
                
                $trackName = [System.Web.HttpUtility]::HtmlDecode($nameMatch.Groups[1].Value.Trim())
                
                # Get timestamp from cue div (format: >00:12< or >05:22<)
                $cueMatch = [regex]::Match($section, '<div[^>]+class="cue[^"]*"[^>]*>(\d{1,2}:\d{2}(?::\d{2})?)</div>')
                $timestamp = if ($cueMatch.Success) { $cueMatch.Groups[1].Value } else { '' }
                
                if (-not [string]::IsNullOrWhiteSpace($trackName) -and -not [string]::IsNullOrWhiteSpace($timestamp)) {
                    $trackNumber++
                    $tracks.Add([PSCustomObject]@{
                        Position  = $trackNumber
                        Timestamp = $timestamp
                        Track     = $trackName
                    })
                }
            }
        }

        return [PSCustomObject]@{
            Title  = $title
            Tracks = $tracks
        }
    }

    function Select-1001SearchResult {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [AllowEmptyCollection()]
            [PSCustomObject[]]$Results,

            [int]$VideoDurationMinutes = 0,

            [switch]$AutoSelect
        )

        if (-not $Results -or $Results.Count -eq 0) {
            Write-Host "No results found." -ForegroundColor Yellow
            return $null
        }

        # Auto-select mode: pick the first result (highest score)
        if ($AutoSelect) {
            $selected = $Results[0]
            
            # Build the display string with duration in minutes
            $meta = @()
            if ($selected.Date) { $meta += "📅 $($selected.Date)" }
            if ($null -ne $selected.DurationMins) { 
                $durationDisplay = "⏱️ $($selected.DurationMins)m"
                if ($VideoDurationMinutes -gt 0) {
                    $diff = [Math]::Abs($selected.DurationMins - $VideoDurationMinutes)
                    if ($diff -le 1) { $durationDisplay += " ✓" }
                    elseif ($diff -le 5) { $durationDisplay += " ≈" }
                    elseif ($diff -le 15) { $durationDisplay += " ~" }
                    elseif ($diff -gt 30) { $durationDisplay += " ✗" }
                }
                $meta += $durationDisplay
            }
            $metaStr = if ($meta.Count -gt 0) { " [{0}]" -f ($meta -join ' | ') } else { "" }
            
            Write-Host ""
            if ($VideoDurationMinutes -gt 0) {
                Write-Host "Auto-selected (video: ${VideoDurationMinutes}m): " -ForegroundColor Cyan -NoNewline
            } else {
                Write-Host "Auto-selected: " -ForegroundColor Cyan -NoNewline
            }
            Write-Host "$($selected.Title)$metaStr" -ForegroundColor Green
            Write-Host ""
            
            return $selected
        }

        # Interactive mode: show list and prompt
        # Limit results to 15
        $displayResults = $Results | Select-Object -First 15
        $maxIndex = $displayResults.Count
        
        Write-Host ""
        $headerText = "Search Results"
        if ($VideoDurationMinutes -gt 0) {
            $headerText += " (video: ${VideoDurationMinutes}m)"
        }
        Write-Host "${headerText}:" -ForegroundColor Cyan
        Write-Host ("-" * 110) -ForegroundColor DarkGray

        foreach ($result in $displayResults) {
            # Determine match type based on duration difference (matches scoring system)
            $matchType = 'none'
            $diff = $null
            if ($VideoDurationMinutes -gt 0 -and $null -ne $result.DurationMins) {
                $diff = [Math]::Abs($result.DurationMins - $VideoDurationMinutes)
                if ($diff -le 1) { $matchType = 'exact' }      # +100 score
                elseif ($diff -le 5) { $matchType = 'close' }  # +80 score
                elseif ($diff -le 15) { $matchType = 'moderate' } # +40 score
                elseif ($diff -le 30) { $matchType = 'far' }   # +10 score
                else { $matchType = 'bad' }                     # -20 score
            }

            # Index
            Write-Host ("{0,3}. " -f $result.Index) -ForegroundColor Cyan -NoNewline

            # Title
            Write-Host $result.Title -NoNewline

            # Metadata bracket
            if ($result.Date -or $result.DurationMins) {
                Write-Host " [" -ForegroundColor DarkGray -NoNewline

                # Date with emoji
                if ($result.Date) {
                    Write-Host "📅 " -NoNewline
                    Write-Host $result.Date -ForegroundColor Yellow -NoNewline
                }

                # Separator
                if ($result.Date -and $result.DurationMins) {
                    Write-Host " | " -ForegroundColor DarkGray -NoNewline
                }

                # Duration in minutes with match indicator
                if ($null -ne $result.DurationMins) {
                    Write-Host "⏱️ " -NoNewline
                    $durationStr = "$($result.DurationMins)m"
                    
                    switch ($matchType) {
                        'exact' {
                            Write-Host $durationStr -ForegroundColor Green -NoNewline
                            Write-Host " ✓" -ForegroundColor Green -NoNewline
                        }
                        'close' {
                            Write-Host $durationStr -ForegroundColor Green -NoNewline
                            Write-Host " ≈" -ForegroundColor Green -NoNewline
                        }
                        'moderate' {
                            Write-Host $durationStr -ForegroundColor DarkYellow -NoNewline
                            Write-Host " ~" -ForegroundColor DarkYellow -NoNewline
                        }
                        'far' {
                            Write-Host $durationStr -ForegroundColor Gray -NoNewline
                        }
                        'bad' {
                            Write-Host $durationStr -ForegroundColor DarkGray -NoNewline
                            Write-Host " ✗" -ForegroundColor DarkRed -NoNewline
                        }
                        default {
                            Write-Host $durationStr -ForegroundColor White -NoNewline
                        }
                    }
                }

                Write-Host "]" -ForegroundColor DarkGray
            }
            else {
                Write-Host ""
            }
        }

        Write-Host ("-" * 110) -ForegroundColor DarkGray
        Write-Host "  0. " -ForegroundColor DarkGray -NoNewline
        Write-Host "Cancel" -ForegroundColor DarkGray
        Write-Host ""

        do {
            $selection = Read-Host "Select a tracklist (1-$maxIndex, or 0 to cancel)"
            $num = 0
            $valid = [int]::TryParse($selection, [ref]$num) -and $num -ge 0 -and $num -le $maxIndex
            if (-not $valid) {
                Write-Host "Invalid selection. Please enter a number between 0 and $maxIndex." -ForegroundColor Yellow
            }
        } while (-not $valid)

        if ($num -eq 0) {
            return $null
        }

        return $Results | Where-Object { $_.Index -eq $num }
    }

    function Get-1001TracklistIdFromUrl {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Url
        )

        if ($Url -match '/tracklist/([^/]+)') {
            return $Matches[1]
        }

        throw "Could not extract tracklist ID from URL: $Url"
    }

    function Get-1001TracklistLines {
        <#
        .SYNOPSIS
            Fetches tracklist from 1001Tracklists.com and returns lines in standard format.
        #>
        [CmdletBinding()]
        param(
            [string]$SearchQuery,
            [string]$Url,
            [string]$Id,
            [string]$UserEmail,
            [string]$UserPassword,
            [int]$VideoDurationMinutes = 0,
            [switch]$AutoSelect
        )

        # Initialize session
        Initialize-1001Session -UserEmail $UserEmail -UserPassword $UserPassword

        # Determine tracklist ID
        $targetId = $null
        $targetUrl = $null

        if ($SearchQuery) {
            # Extract year from query for server-side filtering
            $queryYear = $null
            if ($SearchQuery -match '\b(19|20)\d{2}\b') {
                $queryYear = $Matches[0]
            }
            
            $results = Search-1001Tracklists -Query $SearchQuery -DurationMinutes $VideoDurationMinutes -Year $queryYear

            if (-not $results -or $results.Count -eq 0) {
                throw "No tracklists found for query: $SearchQuery`nNote: Search requires a 1001Tracklists.com account. Use -Email and -Password to authenticate."
            }

            $selected = Select-1001SearchResult -Results $results -VideoDurationMinutes $VideoDurationMinutes -AutoSelect:$AutoSelect
            if (-not $selected) {
                return $null  # User cancelled
            }

            $targetId = $selected.Id
            $targetUrl = $selected.Url
        }
        elseif ($Url) {
            $targetId = Get-1001TracklistIdFromUrl -Url $Url
            $targetUrl = $Url
        }
        elseif ($Id) {
            $targetId = $Id
        }

        # Fetch the tracklist
        Write-Host "Fetching tracklist..." -ForegroundColor Cyan

        $tracklistData = $null
        $useExport = $UserEmail -and $UserPassword

        if ($useExport) {
            try {
                $tracklistData = Get-1001TracklistExport -Id $targetId -FullUrl $targetUrl
            }
            catch {
                Write-Warning "Export API failed: $_"
                Write-Warning "Falling back to HTML parsing..."
                $useExport = $false
            }
        }

        if (-not $useExport) {
            $parsed = Get-1001TracklistFromHtml -Id $targetId -FullUrl $targetUrl

            $lines = [System.Collections.Generic.List[string]]::new()
            $lines.Add($parsed.Title)
            $lines.Add("")

            foreach ($track in $parsed.Tracks) {
                $line = if ($track.Timestamp) {
                    "[$($track.Timestamp)] $($track.Track)"
                }
                else {
                    $track.Track
                }
                $lines.Add($line)
            }

            $lines.Add("")
            $lines.Add("https://1001.tl/$targetId")

            $tracklistData = $lines -join "`r`n"
        }

        # Return as lines array
        return $tracklistData -split "`r?`n"
    }

    #endregion 1001Tracklists Functions

    #region MKV Functions

    function Find-MkvMerge {
        <#
        .SYNOPSIS
            Locates mkvmerge.exe in PATH or default installation directory.
        #>
        $executable = 'mkvmerge.exe'
        $command = Get-Command $executable -ErrorAction SilentlyContinue

        if ($command) {
            return $command.Source
        }

        $defaultPath = Join-Path $env:ProgramFiles 'MKVToolNix\mkvmerge.exe'
        if (Test-Path $defaultPath) {
            return $defaultPath
        }

        throw "mkvmerge.exe not found in PATH or $env:ProgramFiles\MKVToolNix. Specify path via -MkvMergePath."
    }

    function Get-VideoDurationMinutes {
        <#
        .SYNOPSIS
            Gets the duration of a video file in minutes using mkvmerge.
        #>
        param(
            [Parameter(Mandatory)]
            [string]$Path,

            [Parameter(Mandatory)]
            [string]$MkvMergePath
        )

        $json = & $MkvMergePath --identify --identification-format json $Path 2>$null | ConvertFrom-Json

        if ($json.container.properties.duration) {
            # Duration is in nanoseconds
            $durationNs = [long]$json.container.properties.duration
            $durationMinutes = [math]::Round($durationNs / 1000000000 / 60)
            return [int]$durationMinutes
        }

        return $null
    }

    function Get-ExistingChapters {
        <#
        .SYNOPSIS
            Extracts existing chapters from an MKV/WEBM file using mkvextract.
        .DESCRIPTION
            Returns an array of chapter objects with Timestamp and Title properties,
            or $null if no chapters exist or mkvextract is not available.
        #>
        param(
            [Parameter(Mandatory)]
            [string]$Path,

            [Parameter(Mandatory)]
            [string]$MkvMergePath
        )

        # mkvextract should be in the same directory as mkvmerge
        $mkvExtractPath = Join-Path (Split-Path $MkvMergePath -Parent) 'mkvextract.exe'
        
        if (-not (Test-Path $mkvExtractPath)) {
            Write-Verbose "mkvextract not found, skipping chapter comparison"
            return $null
        }

        # Extract chapters to a temp file
        $tempChapterFile = Join-Path ([System.IO.Path]::GetTempPath()) "chapters_extract_$([guid]::NewGuid().ToString('N')).xml"
        
        try {
            $output = & $mkvExtractPath $Path chapters $tempChapterFile 2>&1
            
            if (-not (Test-Path $tempChapterFile) -or (Get-Item $tempChapterFile).Length -eq 0) {
                Write-Verbose "No existing chapters found in file"
                return $null
            }

            # Parse the XML
            [xml]$xml = Get-Content -LiteralPath $tempChapterFile -Raw
            
            if (-not $xml.Chapters -or -not $xml.Chapters.EditionEntry -or -not $xml.Chapters.EditionEntry.ChapterAtom) {
                return $null
            }

            $chapters = @()
            foreach ($atom in $xml.Chapters.EditionEntry.ChapterAtom) {
                $timestamp = $atom.ChapterTimeStart
                $title = if ($atom.ChapterDisplay -and $atom.ChapterDisplay.ChapterString) {
                    $atom.ChapterDisplay.ChapterString
                } else {
                    ''
                }
                
                $chapters += [PSCustomObject]@{
                    Timestamp = $timestamp
                    Title     = $title
                }
            }

            return $chapters
        }
        catch {
            Write-Verbose "Failed to extract chapters: $_"
            return $null
        }
        finally {
            if (Test-Path $tempChapterFile) {
                Remove-Item -LiteralPath $tempChapterFile -Force -ErrorAction SilentlyContinue
            }
        }
    }

    function Compare-Chapters {
        <#
        .SYNOPSIS
            Compares two sets of chapters to determine if they are equivalent.
        .DESCRIPTION
            Compares chapter count, timestamps, and titles.
            Returns $true if chapters are the same, $false otherwise.
        #>
        param(
            [Parameter(Mandatory)]
            [AllowNull()]
            [PSCustomObject[]]$ExistingChapters,

            [Parameter(Mandatory)]
            [PSCustomObject[]]$NewChapters
        )

        # If no existing chapters, they're different
        if (-not $ExistingChapters -or $ExistingChapters.Count -eq 0) {
            return $false
        }

        # Different count = different chapters
        if ($ExistingChapters.Count -ne $NewChapters.Count) {
            return $false
        }

        # Compare each chapter
        for ($i = 0; $i -lt $NewChapters.Count; $i++) {
            $existing = $ExistingChapters[$i]
            $new = $NewChapters[$i]

            # Compare titles (exact match)
            if ($existing.Title -ne $new.Title) {
                return $false
            }

            # Compare timestamps (normalize both to compare)
            # Existing might be "00:00:00.000000000" format, new is "00:00:00.000"
            $existingTime = $existing.Timestamp -replace '(\.\d{3})\d*', '$1'
            $newTime = $new.Timestamp

            if ($existingTime -ne $newTime) {
                return $false
            }
        }

        return $true
    }

    function ConvertTo-NormalizedTimestamp {
        <#
        .SYNOPSIS
            Normalizes timestamps to hh:mm:ss.mmm format.
        .PARAMETER Time
            Timestamp in mm:ss, hh:mm:ss, or hh:mm:ss.mmm format.
        #>
        param (
            [Parameter(Mandatory)]
            [string]$Time
        )

        $milliseconds = '000'
        if ($Time -match '\.(\d+)$') {
            $milliseconds = $Matches[1].PadRight(3, '0').Substring(0, 3)
            $Time = $Time -replace '\.\d+$', ''
        }

        $parts = $Time -split ':'

        switch ($parts.Count) {
            2 { $normalized = "00:$Time" }
            3 { $normalized = $Time }
            default { throw "Invalid timestamp format: $Time" }
        }

        if ($normalized -notmatch '^(\d{1,2}):([0-5]\d):([0-5]\d)$') {
            throw "Invalid timestamp: $normalized"
        }

        $components = $normalized -split ':'
        $normalized = '{0:D2}:{1}:{2}' -f [int]$components[0], $components[1], $components[2]

        return "$normalized.$milliseconds"
    }

    function Get-TracklistLines {
        <#
        .SYNOPSIS
            Retrieves tracklist lines from file or clipboard.
        #>
        param (
            [string]$FilePath,
            [switch]$UseClipboard
        )

        $pattern = '^\s*\[([0-9:.]+)\]\s*(.+?)\s*$'

        if ($FilePath) {
            $lines = Get-Content -LiteralPath $FilePath
        }
        elseif ($UseClipboard) {
            $clipboardText = Get-Clipboard -Raw
            if ([string]::IsNullOrWhiteSpace($clipboardText)) {
                throw 'Clipboard is empty. Copy a tracklist first.'
            }
            $lines = $clipboardText -split '\r?\n'
            Write-Host "Read $($lines.Count) lines from clipboard." -ForegroundColor Cyan
        }
        else {
            throw "No tracklist source specified."
        }

        $matchedLines = @($lines | Where-Object { $_ -match $pattern })

        if ($matchedLines.Count -eq 0) {
            # Check if we have numbered tracks without timestamps
            $numberedPattern = '^\s*\d{1,3}\.\s+.+'
            $hasNumberedTracks = $lines | Where-Object { $_ -match $numberedPattern }

            if ($hasNumberedTracks) {
                throw "This tracklist has no timestamps yet. Timestamps are added by users on 1001Tracklists.com and may not be available for recent sets. Please wait until timestamps are added or use a different tracklist."
            }

            throw 'No valid tracklist entries found. Expected format: [mm:ss] Title or [hh:mm:ss] Title'
        }

        return $matchedLines
    }

    function ConvertTo-Chapter {
        <#
        .SYNOPSIS
            Parses a tracklist line into a chapter object.
        #>
        param (
            [Parameter(Mandatory)]
            [string]$Line,

            [Parameter(Mandatory)]
            [string]$Language
        )

        if ($Line -notmatch '^\s*\[([0-9:.]+)\]\s*(.+?)\s*$') {
            throw "Invalid tracklist line: $Line"
        }

        $timestamp = ConvertTo-NormalizedTimestamp -Time $Matches[1]
        $title = $Matches[2].Trim()

        return [PSCustomObject]@{
            Timestamp = $timestamp
            Title     = $title
            Language  = $Language
        }
    }

    function ConvertTo-ChapterXml {
        <#
        .SYNOPSIS
            Converts chapter objects to Matroska XML format.
        #>
        param (
            [Parameter(Mandatory)]
            [PSCustomObject[]]$Chapters
        )

        $xmlDoc = [System.Xml.XmlDocument]::new()
        $declaration = $xmlDoc.CreateXmlDeclaration('1.0', 'UTF-8', $null)
        [void]$xmlDoc.AppendChild($declaration)

        $chaptersElement = $xmlDoc.CreateElement('Chapters')
        [void]$xmlDoc.AppendChild($chaptersElement)

        $editionEntry = $xmlDoc.CreateElement('EditionEntry')
        [void]$chaptersElement.AppendChild($editionEntry)

        foreach ($chapter in $Chapters) {
            $chapterAtom = $xmlDoc.CreateElement('ChapterAtom')

            $uidElement = $xmlDoc.CreateElement('ChapterUID')
            $uidElement.InnerText = [BitConverter]::ToUInt64([guid]::NewGuid().ToByteArray(), 0).ToString()
            [void]$chapterAtom.AppendChild($uidElement)

            $timeStart = $xmlDoc.CreateElement('ChapterTimeStart')
            $timeStart.InnerText = $chapter.Timestamp
            [void]$chapterAtom.AppendChild($timeStart)

            $display = $xmlDoc.CreateElement('ChapterDisplay')

            $chapterString = $xmlDoc.CreateElement('ChapterString')
            $chapterString.InnerText = $chapter.Title
            [void]$display.AppendChild($chapterString)

            $language = $xmlDoc.CreateElement('ChapterLanguage')
            $language.InnerText = $chapter.Language
            [void]$display.AppendChild($language)

            [void]$chapterAtom.AppendChild($display)
            [void]$editionEntry.AppendChild($chapterAtom)
        }

        return $xmlDoc
    }

    function Invoke-MkvMerge {
        <#
        .SYNOPSIS
            Executes mkvmerge to embed chapters.
        #>
        param (
            [Parameter(Mandatory)]
            [string]$MkvMergePath,

            [Parameter(Mandatory)]
            [string]$InputFile,

            [Parameter(Mandatory)]
            [string]$OutputFile,

            [Parameter(Mandatory)]
            [string]$ChapterXmlPath
        )

        $arguments = @(
            '-o', $OutputFile
            '--no-chapters'
            '--chapters', $ChapterXmlPath
            $InputFile
        )

        # Capture output and only display with -Verbose
        $output = & $MkvMergePath @arguments 2>&1

        foreach ($line in $output) {
            Write-Verbose $line
        }

        if ($LASTEXITCODE -ne 0 -and $LASTEXITCODE -ne 1) {
            # On error, show output regardless of verbose setting
            foreach ($line in $output) {
                Write-Host $line -ForegroundColor Red
            }
            throw "mkvmerge failed with exit code $LASTEXITCODE"
        }
    }

    #endregion MKV Functions

    # Resolve items that only need to happen once
    if (-not $MkvMergePath) {
        $MkvMergePath = Find-MkvMerge
        Write-Verbose "Using mkvmerge: $MkvMergePath"
    }

    # Resolve credentials for 1001Tracklists
    $resolvedEmail = $Email
    $resolvedPassword = $Password

    if ($Credential) {
        $resolvedEmail = $Credential.UserName
        $resolvedPassword = $Credential.GetNetworkCredential().Password
    }

    # For File/Clipboard modes, read tracklist once in begin block
    $sharedTrackLines = $null
    if ($PSCmdlet.ParameterSetName -eq 'File') {
        $sharedTrackLines = Get-TracklistLines -FilePath $TrackListFile
    }
    elseif ($PSCmdlet.ParameterSetName -eq 'Clipboard') {
        $sharedTrackLines = Get-TracklistLines -UseClipboard
    }

    # Track processed files for summary
    $script:processedCount = 0
    $script:errorCount = 0
}

process {
    # Skip if CreateConfig was handled
    if ($PSCmdlet.ParameterSetName -eq 'CreateConfig') {
        return
    }

    try {
        $currentFile = $InputFile
        Write-Host "`n Processing: $(Split-Path $currentFile -Leaf)" -ForegroundColor Cyan

        # Determine tracklist source and type
        $tracklistSource = $null

        switch ($PSCmdlet.ParameterSetName) {
            'Default' {
                # Search from filename
                $searchQuery = ConvertTo-SearchQuery -FileName (Split-Path $currentFile -Leaf)
                $tracklistSource = @{ Type = 'Search'; Value = $searchQuery }
                Write-Host "Searching for: $searchQuery" -ForegroundColor Cyan
            }

            'Tracklist' {
                # Auto-detect type from -Tracklist parameter
                $tracklistSource = Get-TracklistType -TracklistInput $Tracklist
                
                if ($tracklistSource.Type -eq 'Search') {
                    Write-Host "Searching for: $($tracklistSource.Value)" -ForegroundColor Cyan
                }
                elseif ($tracklistSource.Type -eq 'Url') {
                    Write-Verbose "Fetching from URL: $($tracklistSource.Value)"
                }
                else {
                    Write-Verbose "Fetching by ID: $($tracklistSource.Value)"
                }
            }
        }

        # Get video duration for search filtering
        $videoDurationMinutes = 0
        if ($tracklistSource.Type -eq 'Search' -and -not $NoDurationFilter) {
            $videoDurationMinutes = Get-VideoDurationMinutes -Path $currentFile -MkvMergePath $MkvMergePath
            if ($videoDurationMinutes) {
                Write-Verbose "Video duration: $videoDurationMinutes minutes"
            }
        }

        # Get tracklist lines based on parameter set
        # Loop to allow re-selection if tracklist has no timestamps (interactive mode only)
        $trackLines = $null
        $retrySelection = $true
        
        while ($retrySelection) {
            $retrySelection = $false
            
            $trackLines = switch ($PSCmdlet.ParameterSetName) {
                'Default' {
                    $lines = Get-1001TracklistLines -SearchQuery $tracklistSource.Value -UserEmail $resolvedEmail -UserPassword $resolvedPassword -VideoDurationMinutes $videoDurationMinutes -AutoSelect:$AutoSelect
                    if (-not $lines) {
                        Write-Host "Cancelled." -ForegroundColor Yellow
                        return
                    }
                    $lines
                }

                'Tracklist' {
                    switch ($tracklistSource.Type) {
                        'Search' {
                            $lines = Get-1001TracklistLines -SearchQuery $tracklistSource.Value -UserEmail $resolvedEmail -UserPassword $resolvedPassword -VideoDurationMinutes $videoDurationMinutes -AutoSelect:$AutoSelect
                            if (-not $lines) {
                                Write-Host "Cancelled." -ForegroundColor Yellow
                                return
                            }
                            $lines
                        }
                        'Url' {
                            Get-1001TracklistLines -Url $tracklistSource.Value -UserEmail $resolvedEmail -UserPassword $resolvedPassword
                        }
                        'Id' {
                            Get-1001TracklistLines -Id $tracklistSource.Value -UserEmail $resolvedEmail -UserPassword $resolvedPassword
                        }
                    }
                }

                'File' {
                    $sharedTrackLines
                }

                'Clipboard' {
                    $sharedTrackLines
                }
            }

            # Filter to only lines with timestamps
            $timestampPattern = '^\s*\[([0-9:.]+)\]\s*(.+?)\s*$'
            $filteredLines = @($trackLines | Where-Object { $_ -match $timestampPattern })

            if ($filteredLines.Count -eq 0) {
                # Check if we have numbered tracks without timestamps (common for incomplete tracklists)
                $numberedPattern = '^\s*\d{1,3}\.\s+.+'
                $hasNumberedTracks = $trackLines | Where-Object { $_ -match $numberedPattern }

                if ($hasNumberedTracks) {
                    $noTimestampMsg = "This tracklist has no timestamps yet. Timestamps are added by users on 1001Tracklists.com."
                    
                    # In interactive search mode, allow re-selection
                    $isInteractiveSearch = -not $AutoSelect -and ($PSCmdlet.ParameterSetName -eq 'Default' -or ($PSCmdlet.ParameterSetName -eq 'Tracklist' -and $tracklistSource.Type -eq 'Search'))
                    
                    if ($isInteractiveSearch) {
                        Write-Host "$noTimestampMsg" -ForegroundColor Yellow
                        Write-Host "Please select a different tracklist.`n" -ForegroundColor Yellow
                        $retrySelection = $true
                        continue
                    }
                    else {
                        throw "$noTimestampMsg Please wait until timestamps are added or use a different tracklist."
                    }
                }

                throw 'No valid tracklist entries found. Expected format: [mm:ss] Title or [hh:mm:ss] Title'
            }
            
            $trackLines = $filteredLines
        }

        $chapters = foreach ($line in $trackLines) {
            ConvertTo-Chapter -Line $line -Language $ChapterLanguage
        }

        Write-Verbose "Parsed $($chapters.Count) chapters"

        # Preview mode
        if ($Preview) {
            Write-Host "`nParsed Chapters:" -ForegroundColor Green
            $script:i = 0
            $chapters | Format-Table -Property @(
                @{ Label = '#'; Expression = { $script:i++; $script:i } }
                'Timestamp'
                'Title'
            ) -AutoSize
            $script:processedCount++
            return
        }

        # Determine output file
        $resolvedInputFile = Resolve-Path -LiteralPath $currentFile | Select-Object -ExpandProperty Path

        # Check if existing chapters are the same as new chapters
        $existingChapters = Get-ExistingChapters -Path $resolvedInputFile -MkvMergePath $MkvMergePath
        if (Compare-Chapters -ExistingChapters $existingChapters -NewChapters $chapters) {
            Write-Host "Chapters already exist and are identical - skipping." -ForegroundColor Yellow
            $script:processedCount++
            return
        }

        if ($ReplaceOriginal) {
            $finalOutputFile = $resolvedInputFile
            $tempOutputFile = [System.IO.Path]::ChangeExtension($resolvedInputFile, '.tmp' + [System.IO.Path]::GetExtension($resolvedInputFile))
        }
        elseif ($OutputFile) {
            $finalOutputFile = $OutputFile
            $tempOutputFile = $OutputFile
        }
        else {
            $extension = [System.IO.Path]::GetExtension($resolvedInputFile)
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedInputFile)
            $directory = [System.IO.Path]::GetDirectoryName($resolvedInputFile)
            $finalOutputFile = Join-Path $directory "$baseName-new$extension"
            $tempOutputFile = $finalOutputFile
        }

        # Generate chapter XML
        $xmlDoc = ConvertTo-ChapterXml -Chapters $chapters
        $xmlFile = Join-Path ([System.IO.Path]::GetTempPath()) "mkv_chapters_$([guid]::NewGuid().ToString('N')).xml"

        try {
            $xmlDoc.Save($xmlFile)
            Write-Verbose "Chapter XML saved to: $xmlFile"

            # Run mkvmerge
            Write-Host "Muxing chapters..." -ForegroundColor Cyan
            Invoke-MkvMerge -MkvMergePath $MkvMergePath -InputFile $resolvedInputFile -OutputFile $tempOutputFile -ChapterXmlPath $xmlFile

            # Handle replace original
            if ($ReplaceOriginal) {
                Remove-Item -LiteralPath $resolvedInputFile -Force
                Move-Item -LiteralPath $tempOutputFile -Destination $resolvedInputFile -Force
            }

            Write-Host "Chapters added successfully: $finalOutputFile" -ForegroundColor Green
            $script:processedCount++
        }
        finally {
            if (Test-Path $xmlFile) {
                Remove-Item -LiteralPath $xmlFile -Force
            }
        }
    }
    catch {
        Write-Error "Error processing $currentFile`: $_"
        $script:errorCount++
    }
}

end {
    # Skip if CreateConfig was handled or no files processed
    if ($PSCmdlet.ParameterSetName -eq 'CreateConfig') {
        return
    }

    # Show summary if multiple files were processed
    if ($script:processedCount + $script:errorCount -gt 1) {
        Write-Host "`n" -NoNewline
        Write-Host ("-" * 50) -ForegroundColor DarkGray
        Write-Host "Summary: " -ForegroundColor Cyan -NoNewline
        Write-Host "$($script:processedCount) succeeded" -ForegroundColor Green -NoNewline
        if ($script:errorCount -gt 0) {
            Write-Host ", $($script:errorCount) failed" -ForegroundColor Red
        }
        else {
            Write-Host ""
        }
    }
}
<#
.SYNOPSIS
    Adds or replaces chapters in an MKV or WEBM file using mkvmerge.

.DESCRIPTION
    Extracts chapter information from a provided tracklist file, clipboard,
    or directly from 1001Tracklists.com, converts it to Matroska XML format, and uses
    mkvmerge to embed chapters in the media file.
    
    When no tracklist source is specified, the input filename is used as a search query
    on 1001Tracklists.com.
    
    The selected tracklist URL is stored in the output file's MKV tags. On subsequent runs,
    this URL is detected and can be reused automatically (in AutoSelect mode) or with
    confirmation (in interactive mode). Use -IgnoreStoredUrl to force a fresh search.
    
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

.PARAMETER DelaySeconds
    Delay in seconds between processing files in a pipeline to avoid rate limiting.
    Defaults to 5 seconds. Set to 0 to disable delay.

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

.PARAMETER IgnoreStoredUrl
    Ignore any stored 1001Tracklists URL in the file and perform a fresh search.
    By default, if a file has a previously stored tracklist URL, it will be used
    (with confirmation in interactive mode, automatically in AutoSelect mode).

.PARAMETER CreateConfig
    Generate or update config.json and aliases.json files in the script directory.
    Merges with existing files - your settings (credentials, etc.) are preserved,
    and any new default properties are added.

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
    
    Creates or updates config.json, adding any new settings while preserving existing values.

.EXAMPLE
    .\Add-TracklistChapters.ps1 -InputFile "video.mkv" -Preview

.EXAMPLE
    Get-ChildItem *.mkv | .\Add-TracklistChapters.ps1 -AutoSelect -ReplaceOriginal
    
    Batch process all MKV files, auto-selecting tracklists and replacing originals.

.EXAMPLE
    Get-ChildItem *.webm | .\Add-TracklistChapters.ps1 -Tracklist "Festival Name 2025"
    
    Batch process all WEBM files using the same search query.

.EXAMPLE
    .\Add-TracklistChapters.ps1 -InputFile "video.mkv" -IgnoreStoredUrl
    
    Force a new search even if the file has a stored tracklist URL from a previous run.

.EXAMPLE
    .\Add-TracklistChapters.ps1 -InputFile "video.mkv" -AutoSelect -ReplaceOriginal
    
    If the file has a stored URL, uses it directly without searching.
    Otherwise, searches using filename and auto-selects the best match.
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

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'Tracklist')]
    [ValidateRange(0, 60)]
    [int]$DelaySeconds = 5,

    [string]$OutputFile,

    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$MkvMergePath,

    [ValidatePattern('^[a-z]{2,3}$')]
    [string]$ChapterLanguage,

    [switch]$ReplaceOriginal,

    [switch]$Preview,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'Tracklist')]
    [switch]$IgnoreStoredUrl,

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
            AutoSelect       = $false
            DelaySeconds     = 5
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

        # Remove YouTube video ID at end (yt-dlp format: [xXxXxXxXxXx])
        $name = $name -replace '\s*\[[A-Za-z0-9_-]{11}\]\s*$', ''

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

        # Handle config.json - merge with existing if present
        $finalConfig = [ordered]@{}
        foreach ($key in $defaultConfig.Keys) { $finalConfig[$key] = $defaultConfig[$key] }
        $configAction = "Created"
        
        if (Test-Path $configPath) {
            $existingConfig = Get-Config -Path $configPath
            if ($existingConfig.Count -gt 0) {
                # Merge: existing values take precedence over defaults
                foreach ($key in $existingConfig.Keys) {
                    $finalConfig[$key] = $existingConfig[$key]
                }
                
                # Check if any new keys were added
                $newKeys = $defaultConfig.Keys | Where-Object { -not $existingConfig.ContainsKey($_) }
                if ($newKeys) {
                    $configAction = "Updated (added: $($newKeys -join ', '))"
                }
                else {
                    $configAction = "Unchanged (already up to date)"
                }
            }
        }

        # Handle aliases.json - merge with existing if present
        $finalAliases = @{}
        foreach ($key in $defaultAliases.Keys) { $finalAliases[$key] = $defaultAliases[$key] }
        $aliasesAction = "Created"
        
        if (Test-Path $aliasesPath) {
            $existingAliases = Get-Aliases -Path $aliasesPath
            if ($existingAliases.Count -gt 0) {
                # Merge: existing values take precedence over defaults
                foreach ($key in $existingAliases.Keys) {
                    $finalAliases[$key] = $existingAliases[$key]
                }
                
                # Check if any new keys were added
                $newKeys = $defaultAliases.Keys | Where-Object { -not $existingAliases.ContainsKey($_) }
                if ($newKeys) {
                    $aliasesAction = "Updated (added: $($newKeys -join ', '))"
                }
                else {
                    $aliasesAction = "Unchanged (already up to date)"
                }
            }
        }

        # Write config.json
        Save-Config -Path $configPath -Config $finalConfig
        Write-Host "config.json: $configAction" -ForegroundColor $(if ($configAction -eq "Created") { "Green" } elseif ($configAction -match "^Updated") { "Cyan" } else { "DarkGray" })

        # Write aliases.json
        $finalAliases | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $aliasesPath -Encoding UTF8
        Write-Host "aliases.json: $aliasesAction" -ForegroundColor $(if ($aliasesAction -eq "Created") { "Green" } elseif ($aliasesAction -match "^Updated") { "Cyan" } else { "DarkGray" })

        # Show summary
        Write-Host "`nConfig location: $configPath" -ForegroundColor DarkGray
        Write-Host "Aliases location: $aliasesPath" -ForegroundColor DarkGray
        return
    }

    # Load config and apply defaults (command-line parameters override config)
    $config = Get-Config -Path $configPath

    if (-not $PSBoundParameters.ContainsKey('Email') -and $config.ContainsKey('Email') -and $config.Email) {
        $Email = $config.Email
    }
    if (-not $PSBoundParameters.ContainsKey('Password') -and $config.ContainsKey('Password') -and $config.Password) {
        $Password = $config.Password
    }
    if (-not $PSBoundParameters.ContainsKey('ChapterLanguage')) {
        $ChapterLanguage = if ($config.ContainsKey('ChapterLanguage') -and $config.ChapterLanguage) { $config.ChapterLanguage } else { 'eng' }
    }
    if (-not $PSBoundParameters.ContainsKey('MkvMergePath') -and $config.ContainsKey('MkvMergePath') -and $config.MkvMergePath) {
        $MkvMergePath = $config.MkvMergePath
    }
    if (-not $PSBoundParameters.ContainsKey('ReplaceOriginal') -and $config.ContainsKey('ReplaceOriginal') -and $config.ReplaceOriginal) {
        $ReplaceOriginal = [bool]$config.ReplaceOriginal
    }
    if (-not $PSBoundParameters.ContainsKey('NoDurationFilter') -and $config.ContainsKey('NoDurationFilter') -and $config.NoDurationFilter) {
        $NoDurationFilter = [bool]$config.NoDurationFilter
    }
    if (-not $PSBoundParameters.ContainsKey('AutoSelect') -and $config.ContainsKey('AutoSelect') -and $config.AutoSelect) {
        $AutoSelect = [bool]$config.AutoSelect
    }
    if (-not $PSBoundParameters.ContainsKey('DelaySeconds') -and $config.ContainsKey('DelaySeconds')) {
        $DelaySeconds = [int]$config.DelaySeconds
    }

    # Load event aliases for abbreviation matching
    $script:EventAliases = Get-Aliases -Path $aliasesPath
    if ($script:EventAliases.Count -eq 0) {
        # Use defaults if no aliases file exists
        $defaultAliases = New-DefaultAliases
        $defaultAliases.Keys | ForEach-Object { $script:EventAliases[$_.ToLower()] = $defaultAliases[$_] }
        Write-Verbose "Using default aliases (no aliases.json found)"
    }

    # Initialize skip flag
    $script:SkipAllProcessing = $false

    # Early validation: check credentials if we're going to search 1001Tracklists
    # This prevents users from going through the entire search flow only to fail at the end
    $requiresOnlineSearch = -not $TrackListFile -and -not $FromClipboard
    if ($requiresOnlineSearch -and (-not $Email -or -not $Password)) {
        Write-Host "Error: " -ForegroundColor Red -NoNewline
        Write-Host "1001Tracklists.com credentials required."
        Write-Host ""
        Write-Host "To set up credentials:" -ForegroundColor Yellow
        Write-Host "  1. Run: .\Add-TracklistChapters.ps1 -CreateConfig" -ForegroundColor Gray
        Write-Host "  2. Edit config.json and add your 1001Tracklists.com email and password" -ForegroundColor Gray
        Write-Host ""
        Write-Host "Alternatively, use -TrackListFile or -FromClipboard for offline chapter sources." -ForegroundColor Gray
        $script:SkipAllProcessing = $true
        return
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
                    # Handle both DateTime objects (auto-converted by ConvertFrom-Json) and strings
                    $expiry = if ($c.Expires -is [DateTime]) { 
                        $c.Expires 
                    } else { 
                        [DateTime]::Parse($c.Expires, [System.Globalization.CultureInfo]::InvariantCulture) 
                    }
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
                    # Handle both DateTime objects and strings
                    $cookie.Expires = if ($c.Expires -is [DateTime]) { 
                        $c.Expires 
                    } else { 
                        [DateTime]::Parse($c.Expires, [System.Globalization.CultureInfo]::InvariantCulture) 
                    }
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

        # If session already exists and has valid cookies, skip re-initialization
        if ($script:Session) {
            $cookies = $script:Session.Cookies.GetCookies($script:BaseUrl)
            $hasSid = $cookies | Where-Object { $_.Name -eq 'sid' }
            $hasUid = $cookies | Where-Object { $_.Name -eq 'uid' }
            if ($hasSid -and $hasUid) {
                return
            }
        }

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

        # Check for rate limiting - use specific text from the rate limit page
        # The page says "sent too many requests" and "Fill out the captcha to unblock"
        if ($response.Content -match 'sent too many requests|captcha to unblock') {
            Write-Warning "Rate limited by 1001Tracklists. Please solve the captcha in your browser, then retry."
            # Delete cookie cache since the session is now tainted
            if (Test-Path $script:CookieCachePath) {
                Remove-Item -LiteralPath $script:CookieCachePath -Force
                Write-Verbose "Deleted cookie cache due to rate limit."
            }
            $script:Session = $null
            return @()
        }

        $results = [System.Collections.Generic.List[PSCustomObject]]::new()

        # Split HTML by bItm class to process each result item
        $items = $response.Content -split 'class="bItm(?:\s|")'

        # Skip if no results (only pre-content exists)
        if ($items.Count -lt 2) {
            Write-Verbose "No search results found in response"
            return @()
        }

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

    function Write-HighlightedTitle {
        <#
        .SYNOPSIS
            Writes a title with search term matches highlighted.
        .DESCRIPTION
            Highlights matching keywords, abbreviation expansions, alias targets,
            and event pattern matches (e.g., WE1 -> "Weekend 1") in the title.
        #>
        param(
            [Parameter(Mandatory)]
            [string]$Title,

            [Parameter(Mandatory)]
            [hashtable]$QueryParts,

            [string]$HighlightColor = 'White',
            [string]$NormalColor = 'Gray'
        )

        # Build list of terms to highlight (case-insensitive)
        $highlightTerms = [System.Collections.Generic.List[string]]::new()

        # Add keywords (already lowercase)
        foreach ($kw in $QueryParts.Keywords) {
            if ($kw.Length -gt 2) {
                $highlightTerms.Add($kw)
            }
        }

        # Add year
        if ($QueryParts.Year) {
            $highlightTerms.Add($QueryParts.Year)
        }

        # Add abbreviations directly (e.g., "AMF" appearing in title)
        foreach ($abbrev in $QueryParts.Abbreviations) {
            $highlightTerms.Add($abbrev.ToLower())
        }

        # Add alias targets (e.g., TML -> Tomorrowland)
        foreach ($alias in $QueryParts.ResolvedAliases) {
            $highlightTerms.Add($alias.Target.ToLower())
        }

        # Add event pattern expansions (e.g., WE1 -> "Weekend 1")
        foreach ($pattern in $QueryParts.EventPatterns) {
            if ($pattern.Type -eq 'Weekend') {
                $highlightTerms.Add("weekend $($pattern.Number)")
                $highlightTerms.Add("weekend$($pattern.Number)")
                $highlightTerms.Add("w$($pattern.Number)")
                $highlightTerms.Add("we$($pattern.Number)")
            }
            elseif ($pattern.Type -eq 'Day') {
                $highlightTerms.Add("day $($pattern.Number)")
                $highlightTerms.Add("day$($pattern.Number)")
                $highlightTerms.Add("d$($pattern.Number)")
            }
        }

        # Remove duplicates and empty entries
        $highlightTerms = $highlightTerms | Where-Object { $_ } | Select-Object -Unique

        if ($highlightTerms.Count -eq 0) {
            Write-Host $Title -ForegroundColor $NormalColor -NoNewline
            return
        }

        # Build regex pattern - escape special chars and join with |
        # Sort by length descending so longer matches are preferred
        $sortedTerms = $highlightTerms | Sort-Object { $_.Length } -Descending
        $escapedTerms = $sortedTerms | ForEach-Object { [regex]::Escape($_) }
        $pattern = '(' + ($escapedTerms -join '|') + ')'

        # Split title by matches, keeping the delimiters
        $segments = [regex]::Split($Title, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        foreach ($segment in $segments) {
            if ([string]::IsNullOrEmpty($segment)) { continue }

            # Check if this segment matches any highlight term
            $isMatch = $false
            foreach ($term in $highlightTerms) {
                if ($segment -ieq $term) {
                    $isMatch = $true
                    break
                }
            }

            if ($isMatch) {
                Write-Host $segment -ForegroundColor $HighlightColor -NoNewline
            }
            else {
                Write-Host $segment -ForegroundColor $NormalColor -NoNewline
            }
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

        # Pre-calculate event pattern matches (needed for keyword scoring)
        $matchedEventPatterns = @()
        if ($QueryParts.EventPatterns.Count -gt 0) {
            $titleLowerForPatterns = $Result.Title.ToLower()
            foreach ($pattern in $QueryParts.EventPatterns) {
                $num = $pattern.Number
                $type = $pattern.Type
                
                if ($type -eq 'Weekend') {
                    $matchPattern = "(?:weekend\s*$num|w$num|we$num)"
                }
                else {
                    $matchPattern = "(?:day\s*$num|d$num)"
                }
                
                if ($titleLowerForPatterns -match $matchPattern) {
                    $matchedEventPatterns += $pattern
                }
            }
        }

        # Keyword score (important - identifies the correct event/artist)
        # Matched abbreviations, aliases, and event patterns count as matched keywords
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
                # Check if keyword matches an event pattern that was found in the title
                elseif ($matchedEventPatterns | Where-Object {
                    ($_.Type -eq 'Weekend' -and $keywordNormalized -match "^(?:we|w|weekend)$($_.Number)$") -or
                    ($_.Type -eq 'Day' -and $keywordNormalized -match "^(?:d|day)$($_.Number)$")
                }) {
                    $matchedCount++
                    $matchedKws += "$keyword(pat)"
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
            HasEventMatch       = ($matchedAbbreviations.Count -gt 0) -or ($matchedAliases.Count -gt 0) -or ($matchedEventPatterns.Count -gt 0)
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

    function Select-1001SearchResult {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [AllowEmptyCollection()]
            [PSCustomObject[]]$Results,

            [int]$VideoDurationMinutes = 0,

            [string]$SearchQuery,

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
        
        # Parse query for highlighting
        $queryParts = if ($SearchQuery) { Get-QueryParts -Query $SearchQuery } else { $null }
        
        Write-Host ""
        if ($SearchQuery) {
            # Clean YouTube ID for display (same regex as in Get-QueryParts)
            $displayQuery = $SearchQuery -replace '\s*\[[A-Za-z0-9_-]{11}\]\s*$', ''
            Write-Host "Searched for: " -ForegroundColor DarkGray -NoNewline
            Write-Host $displayQuery -ForegroundColor White
        }
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

            # Title with keyword highlighting
            if ($queryParts) {
                Write-HighlightedTitle -Title $result.Title -QueryParts $queryParts
            }
            else {
                Write-Host $result.Title -ForegroundColor Gray -NoNewline
            }

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
                throw "No tracklists found for query: $SearchQuery"
            }

            $selected = Select-1001SearchResult -Results $results -VideoDurationMinutes $VideoDurationMinutes -SearchQuery $SearchQuery -AutoSelect:$AutoSelect
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
            $targetUrl = "$script:BaseUrl/tracklist/$Id/"
        }

        # Fetch the tracklist
        Write-Host "Fetching tracklist..." -ForegroundColor Cyan

        # Export API requires authentication
        if (-not ($UserEmail -and $UserPassword)) {
            throw "Export API requires authentication. Use -Email and -Password or configure credentials in config.json."
        }

        $tracklistData = Get-1001TracklistExport -Id $targetId -FullUrl $targetUrl
        # Extract title from first non-empty line
        $dataLines = $tracklistData -split "`r?`n"
        $tracklistTitle = ($dataLines | Where-Object { $_.Trim() } | Select-Object -First 1)

        # Normalize URL to short format (just ID, no slug)
        $shortUrl = "$script:BaseUrl/tracklist/$targetId/"

        # Return as object with Lines, Url, and Title
        return [PSCustomObject]@{
            Lines = $tracklistData -split "`r?`n"
            Url   = $shortUrl
            Title = $tracklistTitle
        }
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

    function Get-StoredTracklistInfo {
        <#
        .SYNOPSIS
            Extracts stored 1001Tracklists URL and title from MKV/WEBM file tags.
        .DESCRIPTION
            Uses mkvextract to extract tags and looks for 1001TRACKLISTS_URL and
            1001TRACKLISTS_TITLE tags at the global level.
        .OUTPUTS
            Returns a hashtable with Url and Title properties, or $null if not found.
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
            Write-Verbose "mkvextract not found, cannot check for stored tracklist URL"
            return $null
        }

        # Extract tags to a temp file
        $tempTagsFile = Join-Path ([System.IO.Path]::GetTempPath()) "tags_extract_$([guid]::NewGuid().ToString('N')).xml"
        
        try {
            $null = & $mkvExtractPath $Path tags $tempTagsFile 2>&1
            
            if (-not (Test-Path $tempTagsFile)) {
                return $null
            }
            
            $fileSize = (Get-Item $tempTagsFile).Length
            if ($fileSize -eq 0) {
                return $null
            }

            # Parse the XML
            $xmlContent = Get-Content -LiteralPath $tempTagsFile -Raw
            [xml]$xml = $xmlContent
            
            if (-not $xml.Tags) {
                return $null
            }

            # Get Tag elements - handle both single and multiple
            $tagElements = @($xml.Tags.SelectNodes('Tag'))
            if ($tagElements.Count -eq 0) {
                return $null
            }

            # Look for global tags (TargetTypeValue = 70 or no Targets element means global)
            $url = $null
            $title = $null

            foreach ($tag in $tagElements) {
                # Check if this is a global tag (no Targets or TargetTypeValue >= 70)
                $isGlobal = $true
                $targetsNode = $tag.SelectSingleNode('Targets')
                if ($targetsNode) {
                    $targetTypeNode = $targetsNode.SelectSingleNode('TargetTypeValue')
                    if ($targetTypeNode -and $targetTypeNode.InnerText) {
                        $targetType = [int]$targetTypeNode.InnerText
                        $isGlobal = $targetType -ge 70
                    }
                }

                if (-not $isGlobal) { continue }

                # Look for our custom tags in the Simple elements
                $simpleNodes = @($tag.SelectNodes('Simple'))
                foreach ($simple in $simpleNodes) {
                    $nameNode = $simple.SelectSingleNode('Name')
                    $stringNode = $simple.SelectSingleNode('String')
                    
                    if (-not $nameNode -or -not $stringNode) { continue }
                    
                    $tagName = $nameNode.InnerText
                    $tagValue = $stringNode.InnerText
                    
                    if ($tagName -eq '1001TRACKLISTS_URL') {
                        $url = $tagValue
                    }
                    elseif ($tagName -eq '1001TRACKLISTS_TITLE') {
                        $title = $tagValue
                    }
                }
            }

            if ($url) {
                Write-Verbose "Found stored tracklist: $title"
                return @{
                    Url   = $url
                    Title = $title
                }
            }

            return $null
        }
        catch {
            Write-Verbose "Failed to read stored tracklist info: $_"
            return $null
        }
        finally {
            if (Test-Path $tempTagsFile) {
                Remove-Item -LiteralPath $tempTagsFile -Force -ErrorAction SilentlyContinue
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
            Timestamp in mm:ss, m:ss, hh:mm:ss, or hh:mm:ss.mmm format.
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
            2 {
                # mm:ss or m:ss format - pad each part and prepend hours
                $mins = $parts[0].PadLeft(2, '0')
                $secs = $parts[1].PadLeft(2, '0')
                $normalized = "00:$mins`:$secs"
            }
            3 {
                # hh:mm:ss format - pad each part
                $hours = $parts[0].PadLeft(2, '0')
                $mins = $parts[1].PadLeft(2, '0')
                $secs = $parts[2].PadLeft(2, '0')
                $normalized = "$hours`:$mins`:$secs"
            }
            default { throw "Invalid timestamp format: $Time" }
        }

        if ($normalized -notmatch '^(\d{2}):([0-5]\d):([0-5]\d)$') {
            throw "Invalid timestamp: $normalized"
        }

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

    function ConvertTo-TagsXml {
        <#
        .SYNOPSIS
            Creates a Matroska tags XML document with 1001Tracklists metadata.
        .DESCRIPTION
            Creates a global tag containing the tracklist URL and title for embedding
            in the MKV file using mkvmerge --global-tags.
        #>
        param (
            [Parameter(Mandatory)]
            [string]$TracklistUrl,

            [string]$TracklistTitle
        )

        $xmlDoc = [System.Xml.XmlDocument]::new()
        $declaration = $xmlDoc.CreateXmlDeclaration('1.0', 'UTF-8', $null)
        [void]$xmlDoc.AppendChild($declaration)

        $tagsElement = $xmlDoc.CreateElement('Tags')
        [void]$xmlDoc.AppendChild($tagsElement)

        $tagElement = $xmlDoc.CreateElement('Tag')
        [void]$tagsElement.AppendChild($tagElement)

        # Create Targets element for global scope (TargetTypeValue 70 = COLLECTION)
        $targetsElement = $xmlDoc.CreateElement('Targets')
        $targetTypeValue = $xmlDoc.CreateElement('TargetTypeValue')
        $targetTypeValue.InnerText = '70'
        [void]$targetsElement.AppendChild($targetTypeValue)
        [void]$tagElement.AppendChild($targetsElement)

        # Add URL tag
        $simpleUrl = $xmlDoc.CreateElement('Simple')
        $nameUrl = $xmlDoc.CreateElement('Name')
        $nameUrl.InnerText = '1001TRACKLISTS_URL'
        [void]$simpleUrl.AppendChild($nameUrl)
        $stringUrl = $xmlDoc.CreateElement('String')
        $stringUrl.InnerText = $TracklistUrl
        [void]$simpleUrl.AppendChild($stringUrl)
        [void]$tagElement.AppendChild($simpleUrl)

        # Add Title tag if provided
        if ($TracklistTitle) {
            $simpleTitle = $xmlDoc.CreateElement('Simple')
            $nameTitle = $xmlDoc.CreateElement('Name')
            $nameTitle.InnerText = '1001TRACKLISTS_TITLE'
            [void]$simpleTitle.AppendChild($nameTitle)
            $stringTitle = $xmlDoc.CreateElement('String')
            $stringTitle.InnerText = $TracklistTitle
            [void]$simpleTitle.AppendChild($stringTitle)
            [void]$tagElement.AppendChild($simpleTitle)
        }

        return $xmlDoc
    }

    function Invoke-MkvPropedit {
        <#
        .SYNOPSIS
            Executes mkvpropedit to embed chapters and optionally tags in-place.
        .DESCRIPTION
            Uses mkvpropedit for fast in-place metadata modification instead of
            full remuxing with mkvmerge. This is nearly instantaneous regardless
            of file size since only metadata sections are modified.
        #>
        param (
            [Parameter(Mandatory)]
            [string]$MkvToolNixPath,

            [Parameter(Mandatory)]
            [string]$TargetFile,

            [Parameter(Mandatory)]
            [string]$ChapterXmlPath,

            [string]$TagsXmlPath
        )

        $mkvPropeditPath = Join-Path (Split-Path $MkvToolNixPath -Parent) 'mkvpropedit.exe'
        
        if (-not (Test-Path $mkvPropeditPath)) {
            throw "mkvpropedit.exe not found at: $mkvPropeditPath"
        }

        $arguments = @(
            $TargetFile
            '--chapters', $ChapterXmlPath
        )

        # Add tags if provided
        if ($TagsXmlPath) {
            $arguments += '--tags', "global:$TagsXmlPath"
        }

        # Capture output and only display with -Verbose
        $output = & $mkvPropeditPath @arguments 2>&1

        foreach ($line in $output) {
            Write-Verbose $line
        }

        if ($LASTEXITCODE -ne 0) {
            # On error, show output regardless of verbose setting
            foreach ($line in $output) {
                Write-Host $line -ForegroundColor Red
            }
            throw "mkvpropedit failed with exit code $LASTEXITCODE"
        }
    }

    #endregion MKV Functions

    # Resolve items that only need to happen once
    if (-not $MkvMergePath) {
        $MkvMergePath = Find-MkvMerge
        Write-Verbose "Using MKVToolNix: $(Split-Path $MkvMergePath -Parent)"
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
    $script:addedCount = 0
    $script:upToDateCount = 0
    $script:skippedCount = 0
    $script:errorCount = 0
}

process {
    # Skip if early validation failed (e.g., missing credentials)
    if ($script:SkipAllProcessing) {
        return
    }

    # Skip if CreateConfig was handled
    if ($PSCmdlet.ParameterSetName -eq 'CreateConfig') {
        return
    }

    # Add delay between files to avoid rate limiting (skip first file)
    if ($script:processedCount -gt 0 -and $DelaySeconds -gt 0) {
        Write-Host "Waiting $DelaySeconds seconds..." -ForegroundColor DarkGray
        Start-Sleep -Seconds $DelaySeconds
    }
    $script:processedCount++

    try {
        $currentFile = $InputFile
        Write-Host ""
        Write-Host ("=" * 80) -ForegroundColor DarkGray
        Write-Host " Processing: $(Split-Path $currentFile -Leaf)" -ForegroundColor Cyan

        # Track tracklist metadata for embedding in output file
        $script:tracklistUrl = $null
        $script:tracklistTitle = $null

        # Check for stored tracklist URL (only for Default and Tracklist parameter sets)
        $storedInfo = $null
        $useStoredUrl = $false
        
        if (-not $IgnoreStoredUrl -and ($PSCmdlet.ParameterSetName -eq 'Default' -or $PSCmdlet.ParameterSetName -eq 'Tracklist')) {
            $resolvedPath = Resolve-Path -LiteralPath $currentFile | Select-Object -ExpandProperty Path
            $storedInfo = Get-StoredTracklistInfo -Path $resolvedPath -MkvMergePath $MkvMergePath
            
            if ($storedInfo) {
                $displayTitle = if ($storedInfo.Title) { $storedInfo.Title } else { $storedInfo.Url }
                
                if ($AutoSelect) {
                    # In AutoSelect mode, use stored URL directly
                    Write-Host "Using stored tracklist: $displayTitle" -ForegroundColor Green
                    $useStoredUrl = $true
                }
                else {
                    # In interactive mode, prompt user
                    Write-Host "`nFound stored tracklist:" -ForegroundColor Cyan
                    Write-Host "  $displayTitle" -ForegroundColor Yellow
                    Write-Host "  $($storedInfo.Url)" -ForegroundColor DarkGray
                    Write-Host ""
                    Write-Host "  Y = Use this tracklist  |  S = Skip file  |  R = Retry search" -ForegroundColor DarkGray
                    
                    $response = Read-Host "Use this tracklist? (Y/s/r)"
                    if ($response -match '^[Ss]') {
                        Write-Host "Skipping file." -ForegroundColor Yellow
                        $script:skippedCount++
                        return
                    }
                    elseif ($response -match '^[Rr]') {
                        Write-Host "Starting new search..." -ForegroundColor Cyan
                        # Reset storedInfo since user wants fresh search
                        $storedInfo = $null
                        # Continue with normal search flow
                    }
                    else {
                        # Default to Yes
                        $useStoredUrl = $true
                    }
                }
            }
        }

        # Determine tracklist source and type
        $tracklistSource = $null

        # If using stored URL, set tracklistSource to use it directly
        if ($useStoredUrl) {
            $tracklistSource = @{ Type = 'Url'; Value = $storedInfo.Url }
            $script:tracklistUrl = $storedInfo.Url
            $script:tracklistTitle = $storedInfo.Title
        }
        else {
            switch ($PSCmdlet.ParameterSetName) {
                'Default' {
                    # Search from filename
                    $searchQuery = ConvertTo-SearchQuery -FileName (Split-Path $currentFile -Leaf)
                    $tracklistSource = @{ Type = 'Search'; Value = $searchQuery }
                }

                'Tracklist' {
                    # Auto-detect type from -Tracklist parameter
                    $tracklistSource = Get-TracklistType -TracklistInput $Tracklist
                    
                    if ($tracklistSource.Type -eq 'Url') {
                        Write-Verbose "Fetching from URL: $($tracklistSource.Value)"
                    }
                    elseif ($tracklistSource.Type -eq 'Id') {
                        Write-Verbose "Fetching by ID: $($tracklistSource.Value)"
                    }
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
            
            # Helper to process 1001Tracklists result and extract lines/metadata
            $tracklistResult = $null
            
            switch ($PSCmdlet.ParameterSetName) {
                'Default' {
                    if ($tracklistSource.Type -eq 'Url') {
                        # Using stored URL
                        $tracklistResult = Get-1001TracklistLines -Url $tracklistSource.Value -UserEmail $resolvedEmail -UserPassword $resolvedPassword
                    }
                    else {
                        # Search from filename
                        $tracklistResult = Get-1001TracklistLines -SearchQuery $tracklistSource.Value -UserEmail $resolvedEmail -UserPassword $resolvedPassword -VideoDurationMinutes $videoDurationMinutes -AutoSelect:$AutoSelect
                    }
                    
                    if (-not $tracklistResult) {
                        Write-Host "Cancelled." -ForegroundColor Yellow
                        $script:skippedCount++
                        return
                    }
                    
                    $trackLines = $tracklistResult.Lines
                    if (-not $script:tracklistUrl -and $tracklistResult.Url) {
                        $script:tracklistUrl = $tracklistResult.Url
                        $script:tracklistTitle = $tracklistResult.Title
                    }
                }

                'Tracklist' {
                    switch ($tracklistSource.Type) {
                        'Search' {
                            $tracklistResult = Get-1001TracklistLines -SearchQuery $tracklistSource.Value -UserEmail $resolvedEmail -UserPassword $resolvedPassword -VideoDurationMinutes $videoDurationMinutes -AutoSelect:$AutoSelect
                            if (-not $tracklistResult) {
                                Write-Host "Cancelled." -ForegroundColor Yellow
                                $script:skippedCount++
                                return
                            }
                            $trackLines = $tracklistResult.Lines
                            $script:tracklistUrl = $tracklistResult.Url
                            $script:tracklistTitle = $tracklistResult.Title
                        }
                        'Url' {
                            $tracklistResult = Get-1001TracklistLines -Url $tracklistSource.Value -UserEmail $resolvedEmail -UserPassword $resolvedPassword
                            $trackLines = $tracklistResult.Lines
                            if (-not $script:tracklistUrl) {
                                $script:tracklistUrl = $tracklistResult.Url
                                $script:tracklistTitle = $tracklistResult.Title
                            }
                        }
                        'Id' {
                            $tracklistResult = Get-1001TracklistLines -Id $tracklistSource.Value -UserEmail $resolvedEmail -UserPassword $resolvedPassword
                            $trackLines = $tracklistResult.Lines
                            $script:tracklistUrl = $tracklistResult.Url
                            $script:tracklistTitle = $tracklistResult.Title
                        }
                    }
                }

                'File' {
                    $trackLines = $sharedTrackLines
                }

                'Clipboard' {
                    $trackLines = $sharedTrackLines
                }
            }

            # Filter to only lines with timestamps
            $timestampPattern = '^\s*\[([0-9:.]+)\]\s*(.+?)\s*$'
            $filteredLines = @($trackLines | Where-Object { $_ -match $timestampPattern })

            # Check for incomplete tracklists
            $numberedPattern = '^\s*\d{1,3}\.\s+.+'
            $numberedTracks = @($trackLines | Where-Object { $_ -match $numberedPattern })
            
            Write-Verbose "Tracklist analysis: $($filteredLines.Count) timestamped lines, $($numberedTracks.Count) numbered tracks"

            if ($filteredLines.Count -eq 0) {
                if ($numberedTracks.Count -gt 0) {
                    $noTimestampMsg = "This tracklist has no timestamps yet. Timestamps are added by users on 1001Tracklists.com."
                    
                    if ($AutoSelect) {
                        # AutoSelect mode: skip gracefully and continue with next file
                        Write-Host "No timestamps available - skipping." -ForegroundColor Yellow
                        $script:skippedCount++
                        return
                    }
                    else {
                        # Interactive mode: offer skip or new search
                        Write-Host "$noTimestampMsg" -ForegroundColor Yellow
                        Write-Host ""
                        Write-Host "  R = Retry search  |  S = Skip file" -ForegroundColor DarkGray
                        $retryResponse = Read-Host "Retry or skip? (R/s)"
                        
                        if ($retryResponse -match '^[Ss]') {
                            Write-Host "Skipping file." -ForegroundColor Yellow
                            $script:skippedCount++
                            return
                        }
                        
                        # Switch to search mode for retry - build search query from filename
                        $searchQuery = ConvertTo-SearchQuery -FileName (Split-Path $currentFile -Leaf)
                        $tracklistSource = @{ Type = 'Search'; Value = $searchQuery }
                        $script:tracklistUrl = $null
                        $script:tracklistTitle = $null
                        
                        # Calculate video duration if not already done (needed for search filtering)
                        if ($videoDurationMinutes -eq 0 -and -not $NoDurationFilter) {
                            $videoDurationMinutes = Get-VideoDurationMinutes -Path $currentFile -MkvMergePath $MkvMergePath
                            if ($videoDurationMinutes) {
                                Write-Verbose "Video duration: $videoDurationMinutes minutes"
                            }
                        }
                        
                        $retrySelection = $true
                        continue
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
            return
        }

        # Determine output file
        $resolvedInputFile = Resolve-Path -LiteralPath $currentFile | Select-Object -ExpandProperty Path

        # Check if existing chapters are the same as new chapters
        $existingChapters = Get-ExistingChapters -Path $resolvedInputFile -MkvMergePath $MkvMergePath
        $chaptersIdentical = Compare-Chapters -ExistingChapters $existingChapters -NewChapters $chapters
        
        # Skip only if chapters are identical AND URL is already stored
        if ($chaptersIdentical -and $storedInfo) {
            Write-Host "Chapters already exist and are identical - skipping." -ForegroundColor Yellow
            $script:upToDateCount++
            return
        }
        
        # If chapters are identical but URL not stored, we still need to update to add the URL
        if ($chaptersIdentical -and $script:tracklistUrl) {
            Write-Host "Chapters identical, but adding tracklist URL to file..." -ForegroundColor Cyan
        }

        # Determine the target file for mkvpropedit
        # mkvpropedit modifies files in-place, so for non-replace mode we copy first
        if ($ReplaceOriginal) {
            $targetFile = $resolvedInputFile
        }
        elseif ($OutputFile) {
            $targetFile = $OutputFile
            Write-Host "Copying file..." -ForegroundColor Cyan
            Copy-Item -LiteralPath $resolvedInputFile -Destination $targetFile -Force
        }
        else {
            $extension = [System.IO.Path]::GetExtension($resolvedInputFile)
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedInputFile)
            $directory = [System.IO.Path]::GetDirectoryName($resolvedInputFile)
            $targetFile = Join-Path $directory "$baseName-new$extension"
            Write-Host "Copying file..." -ForegroundColor Cyan
            Copy-Item -LiteralPath $resolvedInputFile -Destination $targetFile -Force
        }

        # Generate chapter XML
        $xmlDoc = ConvertTo-ChapterXml -Chapters $chapters
        $xmlFile = Join-Path ([System.IO.Path]::GetTempPath()) "mkv_chapters_$([guid]::NewGuid().ToString('N')).xml"

        # Generate tags XML if we have a tracklist URL
        $tagsXmlFile = $null
        if ($script:tracklistUrl) {
            $tagsXmlDoc = ConvertTo-TagsXml -TracklistUrl $script:tracklistUrl -TracklistTitle $script:tracklistTitle
            $tagsXmlFile = Join-Path ([System.IO.Path]::GetTempPath()) "mkv_tags_$([guid]::NewGuid().ToString('N')).xml"
            $tagsXmlDoc.Save($tagsXmlFile)
            Write-Verbose "Tags XML saved to: $tagsXmlFile"
        }

        try {
            $xmlDoc.Save($xmlFile)
            Write-Verbose "Chapter XML saved to: $xmlFile"

            # Run mkvpropedit (fast in-place modification)
            Write-Host "Adding chapters..." -ForegroundColor Cyan
            $editParams = @{
                MkvToolNixPath = $MkvMergePath
                TargetFile     = $targetFile
                ChapterXmlPath = $xmlFile
            }
            if ($tagsXmlFile) {
                $editParams.TagsXmlPath = $tagsXmlFile
            }
            Invoke-MkvPropedit @editParams

            Write-Host "Chapters added successfully: $targetFile" -ForegroundColor Green
            $script:addedCount++
        }
        catch {
            # If we copied the file and failed, clean up the copy
            if (-not $ReplaceOriginal -and (Test-Path $targetFile)) {
                Remove-Item -LiteralPath $targetFile -Force -ErrorAction SilentlyContinue
            }
            throw
        }
        finally {
            if (Test-Path $xmlFile) {
                Remove-Item -LiteralPath $xmlFile -Force
            }
            if ($tagsXmlFile -and (Test-Path $tagsXmlFile)) {
                Remove-Item -LiteralPath $tagsXmlFile -Force
            }
        }
    }
    catch {
        Write-Error "Error processing $currentFile`: $_"
        $script:errorCount++
    }
}

end {
    # Skip if early validation failed or CreateConfig was handled
    if ($script:SkipAllProcessing -or $PSCmdlet.ParameterSetName -eq 'CreateConfig') {
        return
    }

    # Show summary if multiple files were processed
    $totalFiles = $script:addedCount + $script:upToDateCount + $script:skippedCount + $script:errorCount
    if ($totalFiles -gt 1) {
        Write-Host "`n" -NoNewline
        Write-Host ("-" * 50) -ForegroundColor DarkGray
        Write-Host "Summary: $totalFiles files processed" -ForegroundColor Cyan
        
        if ($script:addedCount -gt 0) {
            Write-Host "  $($script:addedCount) chapters added" -ForegroundColor Green
        }
        if ($script:upToDateCount -gt 0) {
            Write-Host "  $($script:upToDateCount) already up-to-date" -ForegroundColor DarkGray
        }
        if ($script:skippedCount -gt 0) {
            Write-Host "  $($script:skippedCount) skipped" -ForegroundColor Yellow
        }
        if ($script:errorCount -gt 0) {
            Write-Host "  $($script:errorCount) failed" -ForegroundColor Red
        }
    }
}
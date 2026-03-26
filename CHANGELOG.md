# Changelog

All notable changes to Add-TracklistChapters will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [2.0.0] - 2026-03-26

### Changed
- **BREAKING: Scoring algorithm rewritten to multiplicative approach** — Duration now acts as a multiplier (0.8x–2.0x) on content relevance instead of an independent additive score. Prevents duration-only matches from outranking results with strong keyword matches (e.g., partial recordings that don't match the full set length on 1001Tracklists). AutoSelect results may differ from previous versions.
- Keyword score range increased from 0–75 to 0–120 for better differentiation between partial and full matches
- Exponential backoff increased from 3 retries (max 8s) to 5 retries (max 30s) with jitter

### Added
- **Folder support**: `-InputFile` now accepts folders — all .mkv/.webm files are processed automatically
- `-Recurse` switch to include subfolders when processing a folder
- Multiple paths (files and folders mixed) can be passed in a single command
- Version tracking (`$script:Version` and `.NOTES` in help)
- HTTP 429 status code detection for rate limiting (previously only detected via page content)
- Retry-after-wait on rate limit: waits 30 seconds and retries once before giving up
- Warning when search response contains content but no results could be parsed (detects site format changes)
- `-WhatIf` and `-Confirm` support via `SupportsShouldProcess` (`-Preview` still works as before)
- `Write-Progress` for batch pipeline processing
- `CHANGELOG.md`

### Fixed
- `Get-VideoDurationMinutes` no longer crashes on corrupt files or invalid mkvmerge output
- `Select-1001SearchResult` now always returns a single object instead of potentially an array
- Duration of 0 minutes (very short clips) no longer silently skips the duration filter
- Rate limiting no longer deletes the cookie cache — a rate limit doesn't invalidate the session

## [1.0.0] - 2025-01-01

Initial tracked version. All prior changes are in git history.

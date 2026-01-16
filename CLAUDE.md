# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

RDG Stuttgart Song Wish Processor - A Python tool for validating and processing song wish requests for Random Dance Games (RDG) events.

## Build & Run Commands

```bash
# Setup virtual environment (Python 3.10+ required)
python3.11 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# Run the processor
source .venv/bin/activate
python songwish_processor.py
```

## Architecture

The project consists of a single main script `songwish_processor.py` with the following components:

### Core Functions

- `normalize_text()` - Normalizes text by removing special chars, spaces, accents and converting to lowercase for fuzzy matching
- `clean_youtube_url()` - Removes playlist parameters (`&list`, `&index`) from YouTube URLs
- `parse_timestamp()` - Parses timestamps in MM:SS:00 or MM:SS format to seconds
- `get_youtube_info()` - Fetches video metadata using yt-dlp
- `validate_song()` - Main validation function that runs all checks

### Validation Checks

1. `check_artist_title_match()` - Verifies artist/title appear in YouTube video title
2. `check_is_lyric_video()` - Detects if video is a lyric video (vs official MV or dance practice)
3. `check_duration()` - Ensures song section is ≤90 seconds
4. `check_age_restriction()` - Rejects 18+ content
5. `check_blocked_song()` - Checks against manual blocklist

### Message Generation

- `create_contact_url()` - Generates Instagram or WhatsApp contact URLs
- `create_message()` - Generates bilingual (DE/EN) messages based on validation results, includes artist + title for verification
- `get_greeting_name()` - Extracts Instagram name for personalized greetings

### Output Logic

- **Requester/Dancer**: Uses Instagram name if available, otherwise falls back to email
- **Description**: Populated from "Teil des Liedes" field, "Other (Please use...)" placeholder is removed, defaults to "Chorus" if empty

## File Structure

- `songwish.xlsx` - Input: Song wish requests from Google Forms
- `request.xlsx` - Reference: Template format for Songlist worksheet
- `output.xlsx` - Output: Messages + Songlist worksheets
- `blocked_songs.xlsx` - Config: Manual song blocklist (Artist, Title)

## Key Configuration

In `songwish_processor.py`:
- `FORM_URL` - URL to the correction form (placeholder by default)
- `MAX_SONG_DURATION_SECONDS` - Maximum allowed song section length (90s)
- `FIRST_GUARANTEED_COUNT` - Number of guaranteed first song requests (50)

## Dependencies

- pandas, openpyxl - Excel file handling
- yt-dlp - YouTube video metadata extraction
- requests - HTTP requests

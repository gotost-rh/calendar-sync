# Calendar Sync

One-way Google Calendar synchronization utility for Apps Script.

## Purpose

Syncs events from a source calendar to a destination calendar, excluding working location events (WFH, WFO, hybrid, etc.).

## Features

- **One-way sync**: Copies events from source to destination
- **Configurable sync window**: Default 90 days future, 0 days past
- **Keyword filtering**: Automatically excludes working location events
- **Duplicate prevention**: Checks existing events before creating
- **Calendar validation**: Verifies access to both calendars on startup

## Setup

### 1. Configuration

Copy the example config and fill in your calendar IDs:

```bash
cp _Config.example.js _Config.js
```

Edit `_Config.js` and set:
```javascript
const USER_CONFIG = {
  SOURCE_CALENDAR_ID: 'your-source-calendar@example.com',
  DESTINATION_CALENDAR_ID: 'your-destination-calendar@example.com'
};
```

**Finding Calendar IDs:**
1. Open Google Calendar
2. Click the three dots next to the calendar name → Settings
3. Scroll to "Integrate calendar" section
4. Copy the Calendar ID

### 2. Deploy

```bash
clasp push
```

### 3. Run

Open the Apps Script editor and run the `syncCalendars()` function:

```bash
clasp open
```

Or set up a time-driven trigger to run automatically.

## Configuration Options

### Sync Window

In `Code.js`, adjust these values:

```javascript
SYNC_DAYS_PAST: 0,     // Days in the past to sync
SYNC_DAYS_FUTURE: 90,  // Days in the future to sync
```

### Exclude Keywords

Events containing these keywords (case-insensitive) will be skipped:

- home
- office  
- working from home / working from office
- work from home / work from office
- flexible
- hybrid
- out of office

Modify the `EXCLUDE_KEYWORDS` array in `Code.js` to customize.

## Usage

### Manual Sync

Run `syncCalendars()` from the Apps Script editor or via:

```bash
clasp run syncCalendars
```

### Automated Sync

Set up a time-driven trigger in the Apps Script editor:
1. Click "Triggers" (clock icon)
2. Add Trigger
3. Choose `syncCalendars` function
4. Select time-based trigger (e.g., every hour, daily at 9am)

## How It Works

1. Validates access to both source and destination calendars
2. Calculates sync date range based on configuration
3. Fetches events from source calendar
4. Filters out excluded keywords and user's own RSVP responses
5. Checks for existing events in destination (by title and date)
6. Creates missing events in destination calendar

## Security

- `_Config.js` is excluded from git (contains calendar IDs)
- Use `_Config.example.js` as a template
- Never commit your actual calendar IDs

## Troubleshooting

**Error: "Cannot find source/destination calendar"**
- Verify calendar IDs in `_Config.js`
- Ensure you have access to both calendars
- Check for typos in calendar IDs

**Events not syncing**
- Check if event titles contain exclude keywords
- Verify sync window covers the event dates
- Check Apps Script execution logs for errors

**Duplicate events**
- The script checks for duplicates by title and date
- If events are still duplicating, check if titles are slightly different

## Files

- `Code.js` - Main sync logic
- `_Config.example.js` - Configuration template
- `_Config.js` - Your actual configuration (not committed)
- `appsscript.json` - Apps Script manifest

## License

For internal Red Hat use.

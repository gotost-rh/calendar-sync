// === CONFIGURATION ===
const CONFIG = {
  SOURCE_CALENDAR_ID: "shgoto@redhat.com",
  DESTINATION_CALENDAR_ID: "bc739c9a917559627e4badf6894211393916409fa7c2b0a046da73effc3549f5@group.calendar.google.com",
  SCRIPT_SUFFIX: "[Synced]",
  SYNC_DAYS_PAST: 0,
  SYNC_DAYS_FUTURE: 90,
  EXCLUDE_KEYWORDS: [
    "home", "office", "working from home", "working from office",
    "work from home", "work from office", "flexible", "hybrid", "out of office"
  ]
};

// === UTILITY FUNCTIONS ===

// --- Helper Function: Get Current User's Email (Cached) ---
let cachedUserEmail = null;
function getCurrentUserEmail() {
  if (cachedUserEmail === null) {
    cachedUserEmail = Session.getActiveUser().getEmail().toLowerCase();
  }
  return cachedUserEmail;
}

/**
 * Calculate sync date range based on configuration
 * @returns {Object} Object with startTime and endTime properties
 */
function getSyncDateRange() {
  const startTime = new Date();
  startTime.setDate(startTime.getDate() - CONFIG.SYNC_DAYS_PAST);
  startTime.setHours(0, 0, 0, 0);
  
  const endTime = new Date();
  endTime.setDate(endTime.getDate() + CONFIG.SYNC_DAYS_FUTURE);
  endTime.setHours(23, 59, 59, 999);
  
  return { startTime, endTime };
}

/**
 * Get and validate calendar instances
 * @returns {Object} Object with sourceCalendar and destinationCalendar properties
 * @throws {Error} If calendars cannot be found
 */
function getCalendars() {
  const sourceCalendar = CalendarApp.getCalendarById(CONFIG.SOURCE_CALENDAR_ID);
  const destinationCalendar = CalendarApp.getCalendarById(CONFIG.DESTINATION_CALENDAR_ID);

  if (!sourceCalendar) {
    throw new Error(`Source calendar not found: ${CONFIG.SOURCE_CALENDAR_ID}`);
  }
  if (!destinationCalendar) {
    throw new Error(`Destination calendar not found: ${CONFIG.DESTINATION_CALENDAR_ID}`);
  }

  return { sourceCalendar, destinationCalendar };
}

// === EVENT FILTERING FUNCTIONS ===

/**
 * Check if event should be excluded based on attendee status
 * @param {CalendarEvent} event - The calendar event to check
 * @returns {boolean} True if event should be excluded
 */
function shouldExcludeByAttendeeStatus(event) {
  const myStatusRaw = event.getMyStatus();
  const myStatus = String(myStatusRaw || '').trim().toLowerCase();
  
  if (myStatus === "declined" || myStatus === "no") {
    Logger.log(`    -> Event EXCLUDED because attendee status is '${myStatus}'`);
    return true;
  }
  return false;
}

/**
 * Check if event should be excluded based on overall event status
 * @param {CalendarEvent} event - The calendar event to check
 * @returns {boolean} True if event should be excluded
 */
function shouldExcludeByEventStatus(event) {
  let eventOverallStatusRaw = null;
  if (typeof event.getStatus === 'function') {
    eventOverallStatusRaw = event.getStatus();
  }
  const overallStatus = String(eventOverallStatusRaw || '').trim().toLowerCase();
  
  if (overallStatus === "canceled") {
    Logger.log("    -> Event EXCLUDED because overall event status is 'canceled'");
    return true;
  }
  return false;
}

/**
 * Check if event should be excluded based on keywords (for all-day events)
 * @param {CalendarEvent} event - The calendar event to check
 * @returns {boolean} True if event should be excluded
 */
function shouldExcludeByKeywords(event) {
  const title = event.getTitle();
  
  if (title && event.isAllDayEvent()) {
    const lowerCaseTitle = title.toLowerCase();
    for (const keyword of CONFIG.EXCLUDE_KEYWORDS) {
      if (lowerCaseTitle.includes(keyword)) {
        Logger.log("    -> Event EXCLUDED due to keywords/all-day status");
        return true;
      }
    }
  }
  return false;
}

/**
 * Determine if an event should be excluded from sync
 * @param {CalendarEvent} event - The calendar event to check
 * @returns {boolean} True if event should be excluded
 */
function shouldExcludeEvent(event) {
  return shouldExcludeByAttendeeStatus(event) || 
         shouldExcludeByEventStatus(event) || 
         shouldExcludeByKeywords(event);
}

// === EVENT MANAGEMENT FUNCTIONS ===

/**
 * Clean up existing synced events from destination calendar
 * @param {Calendar} destinationCalendar - The destination calendar
 * @param {Date} startTime - Start time for cleanup range
 * @param {Date} endTime - End time for cleanup range
 * @returns {number} Number of events deleted
 */
function cleanupExistingSyncedEvents(destinationCalendar, startTime, endTime) {
  Logger.log("Step 1: Deleting all previously synced events from destination calendar within the range...");
  
  const existingDestinationEvents = destinationCalendar.getEvents(startTime, endTime);
  let deletedCount = 0;

  for (const event of existingDestinationEvents) {
    if (event.getTitle().includes(CONFIG.SCRIPT_SUFFIX)) {
      try {
        event.deleteEvent();
        deletedCount++;
      } catch (e) {
        Logger.log(`  ERROR deleting old synced event '${event.getTitle()}': ${e.toString()}`);
      }
    }
  }
  
  Logger.log(`Deleted ${deletedCount} previously synced events from destination calendar`);
  return deletedCount;
}

/**
 * Create a new event in the destination calendar
 * @param {Calendar} destinationCalendar - The destination calendar
 * @param {CalendarEvent} sourceEvent - The source event to copy
 * @returns {boolean} True if event was created successfully
 */
function createDestinationEvent(destinationCalendar, sourceEvent) {
  try {
    const newTitle = sourceEvent.getTitle() + " " + CONFIG.SCRIPT_SUFFIX;
    const newDescription = sourceEvent.getDescription();
    const newStartTime = sourceEvent.getStartTime();
    const newEndTime = sourceEvent.getEndTime();
    const newIsAllDay = sourceEvent.isAllDayEvent();

    const options = {
      description: newDescription
    };

    if (newIsAllDay) {
      destinationCalendar.createAllDayEvent(newTitle, newStartTime, newEndTime, options);
    } else {
      destinationCalendar.createEvent(newTitle, newStartTime, newEndTime, options);
    }
    
    Logger.log(`    -> CREATED EVENT: '${newTitle}'`);
    return true;
  } catch (e) {
    Logger.log(`    -> ERROR creating event '${sourceEvent.getTitle()}': ${e.toString()}`);
    return false;
  }
}

/**
 * Process and sync events from source to destination calendar
 * @param {Calendar} sourceCalendar - The source calendar
 * @param {Calendar} destinationCalendar - The destination calendar
 * @param {Date} startTime - Start time for sync range
 * @param {Date} endTime - End time for sync range
 * @returns {number} Number of events created
 */
function processAndSyncEvents(sourceCalendar, destinationCalendar, startTime, endTime) {
  const sourceEvents = sourceCalendar.getEvents(startTime, endTime);
  Logger.log(`Step 2: Processing ${sourceEvents.length} raw events from source calendar within the date range`);
  Logger.log("Step 3: Creating/Re-creating events in destination calendar...");

  let createdCount = 0;

  for (const event of sourceEvents) {
    const title = event.getTitle();
    const eventId = event.getId();
    const myStatusRaw = event.getMyStatus();
    const myStatus = String(myStatusRaw || '').trim().toLowerCase();

    let eventOverallStatusRaw = null;
    if (typeof event.getStatus === 'function') {
      eventOverallStatusRaw = event.getStatus();
    }
    const overallStatus = String(eventOverallStatusRaw || '').trim().toLowerCase();

    Logger.log(`  Source Event Found: Title: '${title}', ID: '${eventId}', AllDay: ${event.isAllDayEvent()}, Start: ${event.getStartTime()}, MyStatus (Normalized): ${myStatus}, Overall Status (Normalized): ${overallStatus}`);

    if (!shouldExcludeEvent(event)) {
      if (createDestinationEvent(destinationCalendar, event)) {
        createdCount++;
      }
    }
  }

  return createdCount;
}

// === MAIN SYNCHRONIZATION FUNCTION ===

/**
 * Main function to synchronize calendars
 * Implements a clear & re-create strategy for reliability
 */
function syncCalendars() {
  try {
    // Get date range and calendars
    const { startTime, endTime } = getSyncDateRange();
    const { sourceCalendar, destinationCalendar } = getCalendars();

    Logger.log(`Starting calendar sync (Clear & Re-create strategy) from '${sourceCalendar.getName()}' to '${destinationCalendar.getName()}'`);
    Logger.log(`Syncing events from ${startTime.toDateString()} to ${endTime.toDateString()}`);

    // Clean up existing synced events
    const deletedCount = cleanupExistingSyncedEvents(destinationCalendar, startTime, endTime);

    // Process and sync events
    const createdCount = processAndSyncEvents(sourceCalendar, destinationCalendar, startTime, endTime);

    // Log summary
    Logger.log("Sync complete!");
    Logger.log(`Summary: Created: ${createdCount} events, Deleted Old Synced: ${deletedCount} events`);

  } catch (error) {
    Logger.log(`FATAL ERROR during sync: ${error.toString()}`);
    throw error;
  }
}
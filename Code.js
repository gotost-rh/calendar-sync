// --- Helper Function: Get Current User's Email (Cached) ---
let cachedUserEmail = null;
function getCurrentUserEmail() {
  if (cachedUserEmail === null) {
    cachedUserEmail = Session.getActiveUser().getEmail().toLowerCase();
  }
  return cachedUserEmail;
}


// --- Main Synchronization Function ---
function syncCalendars() {
  // === CONFIGURATION ===
  const SOURCE_CALENDAR_ID = "shgoto@redhat.com";
  const DESTINATION_CALENDAR_ID = "bc739c9a917559627e4badf6894211393916409fa7c2b0a046da73effc3549f5@group.calendar.google.com";

  const SCRIPT_SUFFIX = "[Synced]";

  const SYNC_DAYS_PAST = 0;
  const SYNC_DAYS_FUTURE = 90;

  const startTime = new Date();
  startTime.setHours(0, 0, 0, 0);
  const endTime = new Date();
  endTime.setDate(endTime.getDate() + SYNC_DAYS_FUTURE);
  endTime.setHours(23, 59, 59, 999);

  const EXCLUDE_KEYWORDS = [
    "home", "office", "working from home", "working from office",
    "work from home", "work from office", "flexible", "hybrid", "out of office"
  ];
  // === END CONFIGURATION ===

  const sourceCalendar = CalendarApp.getCalendarById(SOURCE_CALENDAR_ID);
  const destinationCalendar = CalendarApp.getCalendarById(DESTINATION_CALENDAR_ID);

  if (!sourceCalendar) {
    Logger.log("Error: Source calendar not found: " + SOURCE_CALENDAR_ID);
    return;
  }
  if (!destinationCalendar) {
    Logger.log("Error: Destination calendar not found: " + DESTINATION_CALENDAR_ID);
    return;
  }

  Logger.log("Starting calendar sync (Clear & Re-create strategy) from '" + sourceCalendar.getName() + "' to '" + destinationCalendar.getName() + "'.");
  Logger.log("Syncing events from " + startTime.toDateString() + " to " + endTime.toDateString());


  // --- Step 1: Clean up existing synced events (unchanged) ---
  Logger.log("Step 1: Deleting all previously synced events from destination calendar within the range...");
  const existingDestinationEvents = destinationCalendar.getEvents(startTime, endTime);
  let deletedCount = 0;

  for (let i = 0; i < existingDestinationEvents.length; i++) {
    const event = existingDestinationEvents[i];
    if (event.getTitle().includes(SCRIPT_SUFFIX)) {
      try {
        event.deleteEvent();
        deletedCount++;
      } catch (e) {
        Logger.log("  ERROR deleting old synced event '" + event.getTitle() + "': " + e.toString());
      }
    }
  }
  Logger.log("Deleted " + deletedCount + " previously synced events from destination calendar.");


  // --- Step 2: Get all relevant event instances from the source calendar and re-create ---
  const sourceEvents = sourceCalendar.getEvents(startTime, endTime);
  Logger.log("Step 2: Processing " + sourceEvents.length + " raw events from source calendar within the date range.");

  let createdCount = 0;
  Logger.log("Step 3: Creating/Re-creating events in destination calendar...");

  const currentUserEmail = getCurrentUserEmail(); // Still useful for reference

  for (let i = 0; i < sourceEvents.length; i++) {
    const event = sourceEvents[i];
    const title = event.getTitle();
    const eventId = event.getId();

    const myStatusRaw = event.getMyStatus(); // e.g., "OWNER", "YES", "MAYBE", "NO", "INVITED"
    const myStatus = String(myStatusRaw || '').trim().toLowerCase();

    let eventOverallStatusRaw = null;
    if (typeof event.getStatus === 'function') {
        eventOverallStatusRaw = event.getStatus();
    }
    const overallStatus = String(eventOverallStatusRaw || '').trim().toLowerCase();

    Logger.log("  Source Event Found: Title: '" + title + "', ID: '" + eventId + "', AllDay: " + event.isAllDayEvent() + ", Start: " + event.getStartTime() + ", MyStatus (Normalized): " + myStatus + ", Overall Status (Normalized): " + overallStatus);

    let shouldExclude = false;

    // EXCLUSION 1: Exclude if current user's attendee status is "declined" or "no"
    // This will work for events where you are *not* the owner and you declined/said no.
    if (myStatus === "declined" || myStatus === "no") {
        shouldExclude = true;
        Logger.log("    -> Event EXCLUDED because attendee status is 'declined' or 'no'.");
    }

    // EXCLUSION 2: Exclude based on overall event status (e.g., CANCELED)
    // This will work for events that have been formally cancelled.
    if (!shouldExclude && overallStatus === "canceled") {
        shouldExclude = true;
        Logger.log("    -> Event EXCLUDED because overall event status is 'canceled'.");
    }

    // EXCLUSION 3: Existing exclusion for keywords (only if not already excluded)
    if (!shouldExclude && title && event.isAllDayEvent()) {
      const lowerCaseTitle = title.toLowerCase();
      for (const keyword of EXCLUDE_KEYWORDS) {
        if (lowerCaseTitle.includes(keyword)) {
          shouldExclude = true;
          Logger.log("    -> Event EXCLUDED due to keywords/all-day status.");
          break;
        }
      }
    }

    if (!shouldExclude) {
      // Create the event in the destination calendar
      try {
        const newTitle = title + " " + SCRIPT_SUFFIX;
        const newDescription = event.getDescription();
        const newStartTime = event.getStartTime();
        const newEndTime = event.getEndTime();
        const newIsAllDay = event.isAllDayEvent();

        const options = {
            description: newDescription
        };

        if (newIsAllDay) {
          destinationCalendar.createAllDayEvent(newTitle, newStartTime, newEndTime, options);
        } else {
          destinationCalendar.createEvent(newTitle, newStartTime, newEndTime, options);
        }
        createdCount++;
        Logger.log("    -> CREATED EVENT: '" + newTitle + "'");
      } catch (e) {
        Logger.log("    -> ERROR creating event '" + title + "': " + e.toString());
      }
    } else {
      // Log for excluded events already handled in the if blocks above.
    }
  }

  Logger.log("Sync complete!");
  Logger.log("Summary: Created: " + createdCount + " events, Deleted Old Synced: " + deletedCount + " events.");
}
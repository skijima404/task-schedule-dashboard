// === DebugInspector.gs ===

function listCalendarsWithEventsOnMay1() {
  const timezone = Session.getScriptTimeZone();
  const calendars = CalendarApp.getAllCalendars();

  const date = new Date('2025-05-01T00:00:00');
  const start = new Date(date.setHours(0, 0, 0, 0));
  const end = new Date(date.setHours(23, 59, 59, 999));

  calendars.forEach(cal => {
    const events = cal.getEvents(start, end);
    if (events.length > 0) {
      Logger.log(`📅 Name: ${cal.getName()}, ID: ${cal.getId()}, Events: ${events.length}`);
      events.forEach(e => {
        Logger.log(`  └▶ ${e.getTitle()} (${e.getStartTime()} - ${e.getEndTime()})`);
        Logger.log(`     Status: ${e.getMyStatus?.()}`);
        Logger.log(`     IsOwnedByMe: ${e.isOwnedByMe?.()}`);
      });
    }
  });
}

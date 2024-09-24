function postCalendars()
// gets calendar details from Google calendar and post it on "Calendars" sheet
// Visibility of all calendars is reset to "Visible"
{
  let calendars = []
  calendars.push(CalendarApp.getCalendarsByName("Sample Calendar 1")[0])
  calendars.push(CalendarApp.getCalendarsByName("Sample Calendar 2")[0])

  let sheet_calendars = SpreadsheetApp.openById("{sheet_id}").getSheetByName("Calendars")
  let sheet_calendars_data = []

  sheet_calendars_data.push(["Calendar ID", "Calendar Name", "Calendar Description", "Calendar Visibility"]) //headings
  calendars.forEach(calendar => sheet_calendars_data.push([calendar.getId(), calendar.getName(), calendar.getDescription(), "Visible"]))

  console.warn("clearing data from Calendars sheet")
  sheet_calendars.clearContents()
  console.warn("putting calendar data on sheet")

  sheet_calendars.getRange(1, 1, sheet_calendars_data.length, sheet_calendars_data[0].length).setValues(sheet_calendars_data)

  return
}

function postEvents()
// gets today's events from Google calendar and posts it on "Events" sheet
{
  let calendars = []
  calendars.push(CalendarApp.getCalendarsByName("Sample Calendar 1")[0])
  calendars.push(CalendarApp.getCalendarsByName("Sample Calendar 2")[0])
  let events = []
  calendars.forEach(calendar => events.push(...calendar.getEventsForDay(new Date())))

  let sheet_events = SpreadsheetApp.openById("{sheet_id}").getSheetByName("Events")
  let sheet_events_data = []

  sheet_events_data.push(["Calendar ID", "Event ID", "Event Name", "Event Description", "Event End DateTime", "Event Color"])
  events.forEach(event => sheet_events_data.push([event.getOriginalCalendarId(), event.getId(), event.getTitle(), event.getDescription(), event.getEndTime(), event.getColor() == "" ? 0 : event.getColor()]))

  console.warn("Clearing data on Events sheet")
  sheet_events.clearContents()
  console.warn("Posting events data on sheet")
  sheet_events.getRange(1, 1, sheet_events_data.length, sheet_events_data[0].length).setValues(sheet_events_data)

  return
}

function postFilterEvents(start_datetime, end_datetime) 
// gets events between start_datetime and end_datetime from Google Calendar and post it on "Filter_Data" sheet
{
  start_datetime = new Date(start_datetime)
  end_datetime = new Date(end_datetime)

  let calendars = []
  calendars.push(CalendarApp.getCalendarsByName("Sample Calendar 1")[0])
  calendars.push(CalendarApp.getCalendarsByName("Sample Calendar 2")[0])
  let events = []
  calendars.forEach(calendar => events.push(...calendar.getEvents(start_datetime, end_datetime)))

  let filter_history_sheet = SpreadsheetApp.openById("{sheet_id}").getSheetByName("Filter_Data")
  let filter_history_data = []

  filter_history_data.push(["Calendar ID", "Event ID", "Event Name", "Event Description", "Event Color", "Event Start DateTime", "Event End DateTime"])
  events.forEach(event => filter_history_data.push([event.getOriginalCalendarId(), event.getId(), event.getTitle(), event.getDescription(), event.getColor() == "" ? 0 : event.getColor(), event.getStartTime(), event.getEndTime()]))

  console.warn("clearing data from Filter_History sheet")
  filter_history_sheet.clearContents()
  console.warn("Putting events for range: \n" + start_datetime + " to " + end_datetime + "\non Filter_History sheet")
  filter_history_sheet.getRange(1, 1, filter_history_data.length, filter_history_data[0].length).setValues(filter_history_data)

  return
}

function putEventCompletion(calendar_id, event_id, event_color) 
// Chages the highlight color of event based on data changes from the App
{

  let calendar = CalendarApp.getCalendarById(calendar_id)
  let event = calendar.getEventsForDay(new Date())
  event = event.filter(e => e.getId() == event_id)[0]

  console.warn("setting color code \"" + event_color + "\" of event \"" + event.getTitle() + "\"")
  event.setColor(event_color)

  return
}

function syncCalendarsEvents() 
// Syncs all data (calendars, events, filter) from Google Calendar to spreasheets
{
  postCalendars()
  postEvents()

  let filter_fields_sheet = SpreadsheetApp.openById("{sheet_id}").getSheetByName("Filter_Fields")
  let start_datetime = filter_fields_sheet.getRange(2, 4).getValue()
  let end_datetime = filter_fields_sheet.getRange(2, 5).getValue()
  postFilterEvents(start_datetime, end_datetime)
  
  return
}

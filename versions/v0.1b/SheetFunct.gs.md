# Sheet Functions



## Config

```javascript
const CONFIG = {
  CALENDAR_MODE: 'id', // 'default' or 'id'
  // NOTE: This ID is to the LIVE calendar Tutors Bookings - LIVE.
  CALENDAR_ID: 'fb9defee74020745b9fe88a22a9f95429d6499037d116a29f151616c912a6683@group.calendar.google.com', // only used if CALENDAR_MODE = 'id'

  SHEET_NAME: 'WebForm_Submissions',
  ROOM_VALUES: ['Studio', 'Gallery'],

  STATUS_VALUES: {
    PENDING: 'Pending',
    APPROVED: 'Approved',
    REJECTED: 'Rejected',
    CANCELLED: 'Cancelled',
    CONFLICT: 'Conflict'
  },

  HEADERS: {
    TIMESTAMP: 'Timestamp',
    FULL_NAME: 'Full Name',
    EMAIL: 'Email',
    COURSE_NAME: 'Course Name',
    ROOM: 'Room',
    EVENT_DATE: 'Event Date',
    START_TIME: 'Start Time',
    END_TIME: 'End Time',
    RECURRING: 'Recurring',
    FREQUENCY: 'Frequency',
    REPEAT_UNTIL: 'Repeat Until',
    STATUS: 'Status',
    CALENDAR_EVENT_ID: 'Calendar Event ID',
    PROCESSING_NOTE: 'Processing Note',
    ASSIGNED_ROOM: 'Assigned Room',
  }
};
```





## onApprovalEdit(e)

```javascript
function onApprovalEdit(e) {
// Writes action to System_Log
  logAction_('INFO', 'onApprovalEdit fired', {
    sheet: e && e.range ? e.range.getSheet().getName() : '',
    row: e && e.range ? e.range.getRow() : '',
    col: e && e.range ? e.range.getColumn() : '',
    value: e && e.value ? e.value : '',
    oldValue: e && e.oldValue ? e.oldValue : ''
  });


  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.SHEET_NAME) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 2) return;

  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const editedHeader = String(headerRow[col - 1] || '').trim();

  const headers = getHeaders_(sheet);

  if (editedHeader === CONFIG.HEADERS.ASSIGNED_ROOM) {
    handleAssignedRoomEdit_(e, sheet, row, headers);
    return;
  }

  const statusCol = headers[CONFIG.HEADERS.STATUS];
  if (!statusCol) {
    throw new Error(`Missing header: ${CONFIG.HEADERS.STATUS}`);
  }

  if (col === statusCol) {
    handleStatusEdit_(sheet, row, headers, e);
  }
}




```javascript
function handleStatusEdit_(sheet, row, headers, e) {
  const newStatus = trim_(e.range.getValue());
  const oldStatus = trim_(e.oldValue);
  const noteCol = headers[CONFIG.HEADERS.PROCESSING_NOTE];

  if (!noteCol) throw new Error(`Missing header: ${CONFIG.HEADERS.PROCESSING_NOTE}`);

  // Ignore an empty status cell
  if (!newStatus) return;

  // Added to trap for someone deleting the dropdown and typing their own status
  const allowedStatuses = Object.values(CONFIG.STATUS_VALUES);
  if (!allowedStatuses.includes(newStatus)) {
    sheet.getRange(row, noteCol).setValue(`Unknown status: ${newStatus}`);
    return;
  }

  if (newStatus === CONFIG.STATUS_VALUES.APPROVED) {
    processApprovalRow_(sheet, row, headers);
    return;
  }

  if (newStatus === CONFIG.STATUS_VALUES.CANCELLED) {
    processCancellationRow_(sheet, row, headers, oldStatus);
    return;
  }

  if (newStatus === CONFIG.STATUS_VALUES.REJECTED) {
    if (oldStatus !== CONFIG.STATUS_VALUES.APPROVED) {
      sheet.getRange(row, noteCol).setValue('Booking rejected by manager.');
      return;
    }
    // fall through to safety net below
  } else if (newStatus === CONFIG.STATUS_VALUES.CONFLICT) {
    if (oldStatus !== CONFIG.STATUS_VALUES.APPROVED) {
      sheet.getRange(row, noteCol).setValue(
        'Booking marked as conflict. Review room/date/time before re-approval.'
      );
      return;
    }
    // fall through to safety net below
  }

  // Safety net: any move away from Approved should remove the live event.
  if (
    oldStatus === CONFIG.STATUS_VALUES.APPROVED &&
    newStatus !== CONFIG.STATUS_VALUES.APPROVED
  ) {
    removeLiveCalendarEvent_(
      sheet,
      row,
      headers,
      `Status changed from Approved to ${newStatus}.`,
      newStatus
    );
  }
}
```




## handleAssignedRoomEdit

```javascript
function handleAssignedRoomEdit_(e, sheet, row, headers) {
  const oldValue = String(e.oldValue || '').trim();
  const newValue = String(e.value || '').trim();

  // If nothing materially changed, do nothing
  if (oldValue === newValue) return;

  const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const status = trim_(valueByHeader_(rowValues, headers, CONFIG.HEADERS.STATUS));
  const requestedRoom = trim_(valueByHeader_(rowValues, headers, CONFIG.HEADERS.ROOM));
  const noteCol = headers[CONFIG.HEADERS.PROCESSING_NOTE];
  if (!noteCol) throw new Error(`Missing header: ${CONFIG.HEADERS.PROCESSING_NOTE}`);

  const oldEffectiveRoom = normalizeRoom_(oldValue || requestedRoom);
  const newEffectiveRoom = normalizeRoom_(newValue || requestedRoom);

  const oldEffectiveRoomDisplay = oldEffectiveRoom || oldValue || requestedRoom || '(blank)';
  const newEffectiveRoomDisplay = newEffectiveRoom || newValue || requestedRoom || '(blank)';

  // If effective room did not actually change, do nothing
  if (oldEffectiveRoom === newEffectiveRoom) return;

  // If the booking is live, changing room invalidates approval and removes the event
  if (status === CONFIG.STATUS_VALUES.APPROVED) {
    removeLiveCalendarEvent_(
      sheet,
      row,
      headers,
      `Assigned Room changed from ${oldEffectiveRoomDisplay} to ${newEffectiveRoomDisplay}; approval invalidated and returned to Pending.`,
      CONFIG.STATUS_VALUES.PENDING
    );
    return;
  }

  // If the booking is currently in conflict, keep Conflict but update the note
  if (status === CONFIG.STATUS_VALUES.CONFLICT) {
    sheet.getRange(row, noteCol).setValue(
      `Assigned Room changed from ${oldEffectiveRoomDisplay} to ${newEffectiveRoomDisplay}. Re-approve to test availability for the new room.`
    );
    return;
  }

  // For Pending / Rejected / Cancelled, just note the room change
  sheet.getRange(row, noteCol).setValue(
    `Assigned Room updated from ${oldEffectiveRoomDisplay} to ${newEffectiveRoomDisplay}.`
  );
}
```




## processApprovalRow

```javascript
function processApprovalRow_(sheet, row, headers) {
  const calendar = getTargetCalendar_();
  if (!calendar) {
    throw new Error('Calendar not found. Check CALENDAR_MODE / CALENDAR_ID.');
  }

  const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const fullName = valueByHeader_(rowValues, headers, CONFIG.HEADERS.FULL_NAME);
  const fnEmail = valueByHeader_(rowValues, headers, CONFIG.HEADERS.EMAIL);
  const courseName = valueByHeader_(rowValues, headers, CONFIG.HEADERS.COURSE_NAME);
  const requestedRoom = valueByHeader_(rowValues, headers, CONFIG.HEADERS.ROOM);
  const assignedRoom = valueByHeader_(rowValues, headers, CONFIG.HEADERS.ASSIGNED_ROOM);
  const effectiveRoomRaw = trim_(assignedRoom) || trim_(requestedRoom);
  const room = normalizeRoom_(effectiveRoomRaw);
  const bookingDate = valueByHeader_(rowValues, headers, CONFIG.HEADERS.EVENT_DATE);
  const startTime = valueByHeader_(rowValues, headers, CONFIG.HEADERS.START_TIME);
  const endTime = valueByHeader_(rowValues, headers, CONFIG.HEADERS.END_TIME);
  const recurring = trim_(valueByHeader_(rowValues, headers, CONFIG.HEADERS.RECURRING));
  const frequency = trim_(valueByHeader_(rowValues, headers, CONFIG.HEADERS.FREQUENCY));
  const repeatUntil = valueByHeader_(rowValues, headers, CONFIG.HEADERS.REPEAT_UNTIL);
  const existingEventId = trim_(valueByHeader_(rowValues, headers, CONFIG.HEADERS.CALENDAR_EVENT_ID));

  const statusCol = headers[CONFIG.HEADERS.STATUS];
  const eventIdCol = headers[CONFIG.HEADERS.CALENDAR_EVENT_ID];
  const procNoteCol = headers[CONFIG.HEADERS.PROCESSING_NOTE];

  if (!statusCol) throw new Error(`Missing header: ${CONFIG.HEADERS.STATUS}`);
  if (!eventIdCol) throw new Error(`Missing header: ${CONFIG.HEADERS.CALENDAR_EVENT_ID}`);
  if (!procNoteCol) throw new Error(`Missing header: ${CONFIG.HEADERS.PROCESSING_NOTE}`);

  if (existingEventId) {
    sheet.getRange(row, procNoteCol).setValue('Already created; skipped duplicate approval.');
    return;
  }

  if (!room) {
    sheet.getRange(row, statusCol).setValue(CONFIG.STATUS_VALUES.CONFLICT);
    sheet.getRange(row, procNoteCol).setValue(`Unknown room: ${effectiveRoomRaw}`);
    return;
  }

  if (!bookingDate || !startTime || !endTime) {
    sheet.getRange(row, statusCol).setValue(CONFIG.STATUS_VALUES.CONFLICT);
    sheet.getRange(row, procNoteCol).setValue('Missing date/time fields.');
    return;
  }

  const start = combineDateAndTime_(bookingDate, startTime);
  const end = combineDateAndTime_(bookingDate, endTime);

  if (
    !(start instanceof Date) || isNaN(start.getTime()) ||
    !(end instanceof Date) || isNaN(end.getTime())
  ) {
    sheet.getRange(row, statusCol).setValue(CONFIG.STATUS_VALUES.CONFLICT);
    sheet.getRange(row, procNoteCol).setValue('Invalid date/time values.');
    return;
  }

  const now = new Date();
  if (start < now) {
    sheet.getRange(row, statusCol).setValue(CONFIG.STATUS_VALUES.CONFLICT);
    sheet.getRange(row, procNoteCol).setValue('Booking start time is in the past.');
    return;
  }

  if (end <= start) {
    sheet.getRange(row, statusCol).setValue(CONFIG.STATUS_VALUES.CONFLICT);
    sheet.getRange(row, procNoteCol).setValue('End time must be after start time.');
    return;
  }

  if (recurring === 'Yes') {
    if (!frequency || !repeatUntil) {
      sheet.getRange(row, statusCol).setValue(CONFIG.STATUS_VALUES.CONFLICT);
      sheet.getRange(row, procNoteCol).setValue(
        'Recurring booking is missing Frequency or Repeat Until.'
      );
      return;
    }

    if (repeatUntil < bookingDate) {
      sheet.getRange(row, statusCol).setValue(CONFIG.STATUS_VALUES.CONFLICT);
      sheet.getRange(row, procNoteCol).setValue(
        'Repeat Until must be on or after the Event Date.'
      );
      return;
    }
  }

  const sameRoomConflict = hasRoomConflict_(calendar, start, end, room);
  if (sameRoomConflict) {
    sheet.getRange(row, statusCol).setValue(CONFIG.STATUS_VALUES.CONFLICT);
    sheet.getRange(row, procNoteCol).setValue(
      `Conflict detected at approval time: ${room} is already booked.`
    );
    return;
  }

  const title = `${fullName || 'Tutor'} (${room})`;

  const descriptionLines = [
    `Full Name: ${fullName || ''}`,
    `Email: ${fnEmail || ''}`,
    `Room: ${room}`,
    `Course Name: ${courseName || ''}`,
    `Recurring: ${recurring || 'No'}`,
    `Frequency: ${frequency || ''}`,
    `Repeat Until: ${repeatUntil || ''}`
  ];

let eventId;
let noteMessage;

if (recurring === 'Yes') {

  const recurrence = buildRecurrence_(frequency, repeatUntil, bookingDate);

  const series = calendar.createEventSeries(title, start, end, recurrence, {
    location: room,
    description: descriptionLines.join('\n')
  });

  eventId = series.getId();
  noteMessage = `Recurring event series created successfully in ${room} until ${repeatUntil}.`;

} else {

  const event = calendar.createEvent(title, start, end, {
    location: room,
    description: descriptionLines.join('\n')
  });

  eventId = event.getId();
  noteMessage = `Calendar event created successfully in ${room}.`;
}

sheet.getRange(row, eventIdCol).setValue(eventId);
sheet.getRange(row, procNoteCol).setValue(noteMessage);
}
```





## processCancellationRow

```javascript
function processCancellationRow_(sheet, row, headers, oldStatus) {
  const prefix =
    oldStatus === CONFIG.STATUS_VALUES.APPROVED
      ? 'Booking cancelled.'
      : 'Booking marked as cancelled.';

  removeLiveCalendarEvent_(
    sheet,
    row,
    headers,
    prefix,
    CONFIG.STATUS_VALUES.CANCELLED
  );
}
```





## Helper Functions

```javascript
function hasRoomConflict_(calendar, start, end, room) {
  const overlapping = calendar.getEvents(start, end);

  return overlapping.some(event => {
    const eventRoom = normalizeRoom_(event.getLocation());
    return eventRoom === room;
  });
}







function normalizeRoom_(value) {
  const room = String(value || '').trim();

  if (!room) return '';

  const normalized = room.toLowerCase();

  if (normalized === 'studio') return 'Studio';
  if (normalized === 'gallery') return 'Gallery';

  return '';
}



// * removeLiveCalendarEvent

function removeLiveCalendarEvent_(sheet, row, headers, baseNote, newStatus) {
  const eventIdCol = headers[CONFIG.HEADERS.CALENDAR_EVENT_ID];
  const statusCol = headers[CONFIG.HEADERS.STATUS];
  const noteCol = headers[CONFIG.HEADERS.PROCESSING_NOTE];

  if (!eventIdCol) throw new Error(`Missing header: ${CONFIG.HEADERS.CALENDAR_EVENT_ID}`);
  if (!noteCol) throw new Error(`Missing header: ${CONFIG.HEADERS.PROCESSING_NOTE}`);

  const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const existingEventId = trim_(
    valueByHeader_(rowValues, headers, CONFIG.HEADERS.CALENDAR_EVENT_ID)
  );

  let outcome = 'No linked calendar event to remove.';

  if (existingEventId) {
    const calendar = getTargetCalendar_();
    if (!calendar) {
      throw new Error('Calendar not found. Check CALENDAR_MODE / CALENDAR_ID.');
    }

    const event = calendar.getEventById(existingEventId);

    if (event) {
      event.deleteEvent();
      outcome = 'Linked calendar event removed.';
    } else {
      outcome = 'Linked calendar event not found; ID cleared anyway.';
    }
  }

  sheet.getRange(row, eventIdCol).clearContent();

  if (newStatus && statusCol) {
    const currentStatus = trim_(sheet.getRange(row, statusCol).getValue());
    if (currentStatus !== newStatus) {
      sheet.getRange(row, statusCol).setValue(newStatus);
    }
  }

  sheet.getRange(row, noteCol).setValue(`${baseNote} ${outcome}`.trim());
}



function getHeaders_(sheet) {
  const headerValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headers = {};

  headerValues.forEach((name, i) => {
    const key = String(name).trim();
    if (!key) return;

    if (headers[key]) {
      throw new Error(`Duplicate header found: ${key}`);
    }

    headers[key] = i + 1;
  });

  return headers;
}



function valueByHeader_(rowValues, headers, headerName) {
  const col = headers[headerName];
  if (!col) return '';
  return col - 1 < rowValues.length ? rowValues[col - 1] : '';
}



function trim_(value) {
  return value == null ? '' : String(value).trim();
}



// * getTargetCalendar

function getTargetCalendar_() {
  if (CONFIG.CALENDAR_MODE === 'default') {
    return CalendarApp.getDefaultCalendar();
  }

  if (CONFIG.CALENDAR_MODE === 'id') {
    if (!CONFIG.CALENDAR_ID) {
      throw new Error('CONFIG.CALENDAR_ID is missing.');
    }
    return CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  }

  throw new Error(`Invalid CALENDAR_MODE: ${CONFIG.CALENDAR_MODE}`);
}



// * setPendingOnFormSubmit(e)

function setPendingOnFormSubmit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.SHEET_NAME) return;

  const row = e.range.getRow();
  if (row < 2) return;

  SpreadsheetApp.flush();

  const headers = getHeaders_(sheet);
  const statusCol = headers[CONFIG.HEADERS.STATUS];
  if (!statusCol) {
    throw new Error(`Header '${CONFIG.HEADERS.STATUS}' not found.`);
  }

  const statusCell = sheet.getRange(row, statusCol);
  if (statusCell.isBlank()) {
    statusCell.setValue(CONFIG.STATUS_VALUES.PENDING);
  }
}



function logAction_(level, message, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('System_Log');

  if (!logSheet) {
    logSheet = ss.insertSheet('System_Log');
    logSheet.appendRow(['Timestamp', 'Level', 'Message', 'Details']);
    // logSheet.hideSheet();   // leave this off while testing
  }

  logSheet.appendRow([
    new Date(),
    level || 'INFO',
    message || '',
    details ? JSON.stringify(details) : ''
  ]);
}




// * combineDateAndTime

function combineDateAndTime_(dateValue, timeValue) {
  if (!dateValue || !timeValue) return null;

  const date = new Date(dateValue);
  if (isNaN(date.getTime())) return null;

  let hours;
  let minutes;

  if (timeValue instanceof Date) {
    hours = timeValue.getHours();
    minutes = timeValue.getMinutes();
  } else {
    const timeText = String(timeValue).trim();
    const match = timeText.match(/^(\d{1,2}):(\d{2})$/);
    if (!match) return null;

    hours = Number(match[1]);
    minutes = Number(match[2]);
  }

  const combined = new Date(date);
  combined.setHours(hours, minutes, 0, 0);
  return combined;
}



// buildRecurrence

function buildRecurrence_(frequency, repeatUntil, startDate) {
  const recurrence = CalendarApp.newRecurrence();

  const untilDate = new Date(repeatUntil);

  if (frequency === 'Weekly') {
    recurrence.addWeeklyRule().until(untilDate);
  } else if (frequency === 'Fortnightly') {
    recurrence.addWeeklyRule().interval(2).until(untilDate);
  } else {
    throw new Error(`Unsupported frequency: ${frequency}`);
  }

  return recurrence;
}
```



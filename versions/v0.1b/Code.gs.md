# Form Code


## Main 

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Booking Request - Test App');
}



function submitBooking(formData) {
  const sheetName = CONFIG.SHEET_NAME;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found.`);
  }

  const requiredFields = [
    'fullName',
    'email',
    'courseName',
    'room',
    'eventDate',
    'startTime',
    'endTime',
    'recurring'
  ];

  for (const field of requiredFields) {
    if (!formData[field] || String(formData[field]).trim() === '') {
      throw new Error(`Missing required field: ${field}`);
    }
  }

  if (formData.recurring === 'Yes') {
    if (!formData.frequency || !formData.repeatUntil) {
      throw new Error('Recurring bookings require Frequency and Repeat Until.');
    }
  }

  const start = parseTimeToMinutes(formData.startTime);
  const end = parseTimeToMinutes(formData.endTime);

  if (start % 30 !== 0 || end % 30 !== 0) {
    throw new Error('Times must be entered in 30-minute increments.');
  }

  if (end <= start) {
    throw new Error('End time must be later than start time.');
  }

  const lock = LockService.getScriptLock();
    lock.waitLock(10000);

  try {
    const row = sheet.getLastRow() + 1;

    const rowData = [
      new Date(),
      formData.fullName,
      formData.email,
      formData.courseName,
      formData.room,
      formData.eventDate,
      formData.startTime,
      formData.endTime,
      formData.recurring,
      formData.frequency || '',
      formData.repeatUntil || '',
      CONFIG.STATUS_VALUES.PENDING,
    // Billing columns
      '', // Billable Hours
      '', // Hourly Rate
      '', // Total Fee

    // any later admin columns, if any
      ''
    ];

    sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);

    // Assuming:
    // G = Start Time
    // H = End Time
    // L = Status
    // P = Billable Hours
    // Q = Hourly Rate
    // R = Total Fee

    sheet.getRange(row, 16).setFormula(
      `=IF(OR($L${row}<>"Approved",$G${row}="",$H${row}=""),"",($H${row}-$G${row})*24)`
    );

    sheet.getRange(row, 18).setFormula(
      `=IF(OR($P${row}="",$Q${row}=""),"",$P${row}*$Q${row})`
    );


  logAction_('INFO', 'Booking submitted', {
    fullName: formData.fullName,
    email: formData.email,
    courseName: formData.courseName,
    room: formData.room,
    eventDate: formData.eventDate,
    startTime: formData.startTime,
    endTime: formData.endTime,
    recurring: formData.recurring
  });

  return {
    success: true,
    message: 'Booking request submitted successfully.'
  };

  } finally {
    try {
      lock.releaseLock();
    } catch (err) {
      // Ignore release errors; lock may not have been acquired.
    }
  }
}
```


## Helpers

```javascript
function parseTimeToMinutes(timeStr) {
  const parts = String(timeStr).split(':');
  if (parts.length !== 2) {
    throw new Error(`Invalid time format: ${timeStr}`);
  }

  const hours = Number(parts[0]);
  const minutes = Number(parts[1]);

  if (
    Number.isNaN(hours) ||
    Number.isNaN(minutes) ||
    hours < 0 ||
    hours > 23 ||
    minutes < 0 ||
    minutes > 59
  ) {
    throw new Error(`Invalid time value: ${timeStr}`);
  }

  return hours * 60 + minutes;
}
```

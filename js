const SHEET_NAME = 'form_resp'; #sheet name
const RECIPIENT_EMAIL = 'yourmail@gmail.com'; 

function sendDueEmailReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found');

  const data = sheet.getDataRange().getValues();
  const now = new Date();
  
  for (let i = 1; i < data.length; i++) {
    const subject  = data[i][1]; // Col B
    const message  = data[i][2]; // Col C
    const rawDate  = data[i][3]; // Col D (Expected format: Date)
    const rawTime  = data[i][4]; // Col E (Expected format: Time)
    const sentAt   = data[i][5]; // Col F (Status)

    if (!message || sentAt) continue;

    // --- FIX: Robust Date/Time Extraction ---
    const datePart = new Date(rawDate);
    const timePart = new Date(rawTime);

    if (isNaN(datePart.getTime()) || isNaN(timePart.getTime())) {
      Logger.log(`Row ${i + 1} skipped: Invalid format.`);
      continue;
    }

    // Construct the actual intended time
    const scheduledAt = new Date(
      datePart.getFullYear(),
      datePart.getMonth(),
      datePart.getDate(),
      timePart.getHours(),
      timePart.getMinutes(),
      0
    );

    // --- FIX: Buffer for execution ---
    // If current time is earlier than scheduled time, skip it.
    if (scheduledAt > now) {
      Logger.log(`Row ${i + 1} is scheduled for the future (${scheduledAt}). Skipping.`);
      continue;
    }

    // 1. Send Email
    MailApp.sendEmail({
      to: RECIPIENT_EMAIL,
      subject: subject || 'Reminder',
      body: message
    });

    // 2. Create Calendar Event
    const endTime = new Date(scheduledAt.getTime() + (30 * 60000));
    CalendarApp.getDefaultCalendar().createEvent(
      subject || 'Reminder', 
      scheduledAt, 
      endTime, 
      { description: message }
    );

    // 3. Mark as sent (Column F)
    sheet.getRange(i + 1, 6).setValue(new Date()); 
    Logger.log(`Row ${i + 1} processed successfully.`);
  }
}

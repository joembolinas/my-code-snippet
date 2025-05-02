
## Step 1: Create a Google Form
1. Go to Google Forms (forms.google.com).
2. Create a new blank form.
3. Give your form a title (e.g., "Attendance Logger").
4. Add the following questions (customize as needed):
    Timestamp: (Automatically added)
    Name: (Short answer, Required)
    Activity/Event: (Dropdown or Multiple choice, Required)
    Date: (Date, Optional)
    Notes: (Paragraph, Optional)
5. Go to the "Responses" tab and click the Google Sheets icon to create a linked Google Sheet (e.g., "Attendance Log Data").

## Step 2: Google Apps Script Code
1. Open the linked Google Sheet ("Attendance Log Data").
2. Go to "Extensions" > "Apps Script".
3. Name the script (e.g., "AttendanceLoggerScript").
4. Replace the code with the following:

```
/**
 * Triggers when a form response is submitted.
 * Processes the attendance data and stores it in the "Attendance Records" sheet.
 * @param {Object} e The event object from the form submission.
 */
function onFormSubmit(e) {
  try {
    // Get form response data
    const responses = e.namedValues;
    const timestamp = responses['Timestamp'][0];
    const name = responses['Name'][0];
    const activity = responses['Activity/Event'][0];
    const date = responses['Date'] ? responses['Date'][0] : new Date(timestamp).toLocaleDateString(); // Use provided date or extract from timestamp
    const notes = responses['Notes'] ? responses['Notes'][0] : '';

    // Get the spreadsheet and the attendance sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let attendanceSheet = ss.getSheetByName('Attendance Records');

    // Create the "Attendance Records" sheet if it doesn't exist
    if (!attendanceSheet) {
      attendanceSheet = ss.insertSheet('Attendance Records');
      attendanceSheet.appendRow(['Timestamp', 'Name', 'Activity/Event', 'Date', 'Notes']);
    }

    // Append the attendance record to the sheet
    attendanceSheet.appendRow([timestamp, name, activity, date, notes]);
  } catch (error) {
    // Log any errors to the console
    console.error('Error in onFormSubmit:', error);
    // Optionally, notify the user (e.g., send an email)
    //  MailApp.sendEmail({
    //    to: 'your-email@example.com',
    //    subject: 'Error in Attendance Logger',
    //    body: 'An error occurred while processing a form submission: ' + error.message
    //  });
  }
}

/**
 * Generates a summary report of attendance data.
 * Creates a new sheet ("Attendance Summary") with the report.
 */
function generateSummaryReport() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const attendanceSheet = ss.getSheetByName('Attendance Records');
    const summarySheetName = 'Attendance Summary';

    // Check if the "Attendance Records" sheet exists
    if (!attendanceSheet) {
      SpreadsheetApp.getUi().alert('Error: "Attendance Records" sheet not found.');
      return;
    }

    // Get all data from the "Attendance Records" sheet
    const data = attendanceSheet.getDataRange().getValues();
    if (data.length <= 1) {
      SpreadsheetApp.getUi().alert('No attendance data to summarize.');
      return;
    }
    const headers = data[0]; // Get the headers
    const attendanceRecords = data.slice(1); // Get the data without headers

    // Create or clear the "Attendance Summary" sheet
    let summarySheet = ss.getSheetByName(summarySheetName);
    if (summarySheet) {
      summarySheet.clearContents(); // Clear previous summary data
    } else {
      summarySheet = ss.insertSheet(summarySheetName);
    }

    // Prepare summary data structures
    const activityCounts = {};
    const studentAttendance = {};

    // Process each attendance record
    attendanceRecords.forEach(record => {
      const activity = record[headers.indexOf('Activity/Event')]; // Use header index
      const name = record[headers.indexOf('Name')];         // Use header index

      // Count attendance for each activity
      activityCounts[activity] = (activityCounts[activity] || 0) + 1;

      // Count attendance for each student
      studentAttendance[name] = (studentAttendance[name] || 0) + 1;
    });

    // Write summary data to the "Attendance Summary" sheet
    summarySheet.appendRow(['Activity/Event', 'Total Attendees']);
    for (const activity in activityCounts) {
      summarySheet.appendRow([activity, activityCounts[activity]]);
    }

    summarySheet.appendRow([]); // Add a blank row for separation
    summarySheet.appendRow(['Student Name', 'Total Attendance Count']);
    for (const student in studentAttendance) {
      summarySheet.appendRow([student, studentAttendance[student]]);
    }

    // Notify the user that the report has been generated.
    SpreadsheetApp.getUi().alert('Attendance summary report generated on the "' + summarySheetName + '" sheet.');
  } catch (error) {
    console.error("Error in generateSummaryReport", error);
    SpreadsheetApp.getUi().alert('An error occurred while generating the report.  See the log for details.');
  }
}

/**
 * Creates a custom menu in the spreadsheet.
 * Adds an option to generate the attendance summary report.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Attendance Tools')
    .addItem('Generate Summary Report', 'generateSummaryReport')
    .addToUi();
}
```

## Step 3: Set up the Form Submission Trigger
1. In the Script Editor, click the clock icon (Triggers).
2. Click "Add Trigger".
3. Configure the trigger:
     - Choose which function to run: ```onFormSubmit```
     - Choose which deployment should run: ```Head```
     - Select event source: ```From spreadsheet```
     - Select event type: ```On form submit```
     - Click "Save" and authorize the script.
     
 ## How to Use:
 1. Share the Google Form for attendance logging.
 2. Form submissions are automatically recorded in the "Attendance Records" sheet.
 3. In the Google Sheet, go to "Attendance Tools" > "Generate Summary Report" to create a summary in the "Attendance Summary" sheet.

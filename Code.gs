function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("College Attendance System")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Include HTML files
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Check advisor authorization
function checkAdvisor() {
  const email = Session.getActiveUser().getEmail();
  const sheet = SpreadsheetApp.getActive().getSheetByName("Advisors");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      return { allowed: true, email: email };
    }
  }
  return { allowed: false };
}

// Ensure date columns exist (7 hours)
function ensureDateColumns(dateStr) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Daily_Attendance");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  if (!headers.includes(dateStr + "_H1")) {
    let newCols = [];
    for (let h = 1; h <= 7; h++) {
      newCols.push(dateStr + "_H" + h);
    }
    sheet.getRange(1, headers.length + 1, 1, 7).setValues([newCols]);
  }
}

// Main attendance function
function markAttendance(data) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Daily_Attendance");

  const dateStr = data.date;
  const hour = data.hour;
  const absent = data.absent;
  const od = data.od;
  const others = data.others;

  ensureDateColumns(dateStr);

  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const colIndex = headers.indexOf(dateStr + "_H" + hour) + 1;

  const lastRow = sheet.getLastRow();
  const students = sheet.getRange(2, 1, lastRow-1, 2).getValues();

  // Prevent duplicate marking
  const existing = sheet.getRange(2, colIndex, lastRow-1, 1).getValues().flat();
  if (existing.some(v => v !== "")) {
    return "Attendance already marked for this hour";
  }

  for (let i = 0; i < students.length; i++) {
    const regNo = students[i][0];
    const last3 = regNo.slice(-3);
    let status = "P";

    if (absent.includes(last3)) status = "A";
    else if (od.includes(last3)) status = "OD";
    else if (others.includes(last3)) status = "OTH";

    sheet.getRange(i+2, colIndex).setValue(status);
  }

  return "Attendance marked successfully";
}

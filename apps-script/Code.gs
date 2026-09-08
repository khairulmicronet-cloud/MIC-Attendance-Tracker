const SHEET_NAME = "Log";

function doPost(e) {
  try {
    const sheet = getOrCreateSheet_();
    const data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      new Date(),
      data.name || "",
      data.type || "",
      data.date || "",
      data.time || "",
      data.address || "",
      data.lat || "",
      data.lon || ""
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput("MIC Attendance Logger is running.");
}

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "Server Timestamp",
      "Staff Name",
      "Record Type",
      "Client Date",
      "Client Time",
      "Location Address",
      "Latitude",
      "Longitude"
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function setupSheet() {
  getOrCreateSheet_();
}

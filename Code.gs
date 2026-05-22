// ============================================================
// Field Force Tracker — Google Apps Script Backend
// Deploy as Web App: Execute as "Me", Access "Anyone"
// ============================================================

const SHEET_NAME_VISITS = "Visits";
const SHEET_NAME_SUMMARY = "Summary";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    if (action === "checkin") return handleCheckin(data);
    if (action === "checkout") return handleCheckout(data);
    if (action === "summary") return handleSummary(data);
    
    return jsonResponse({ success: false, error: "Unknown action" });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

function doGet(e) {
  const action = e.parameter.action;
  if (action === "summary") {
    return handleSummary(e.parameter);
  }
  return jsonResponse({ success: false, error: "GET not supported for this action" });
}

function handleCheckin(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME_VISITS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME_VISITS);
    sheet.appendRow([
      "Visit ID", "User ID", "User Name", "Party Name", "Area",
      "Check In Time", "Check In Lat", "Check In Lng", "Check In Address",
      "Geo Verified", "Check Out Time", "Check Out Lat", "Check Out Lng",
      "Duration (mins)", "Order Value", "Collection Value",
      "Payment Remark", "Notes", "Date"
    ]);
    sheet.getRange(1, 1, 1, 19).setFontWeight("bold").setBackground("#1D9E75").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  
  const visitId = "V" + new Date().getTime();
  const now = new Date();
  
  sheet.appendRow([
    visitId,
    data.userId,
    data.userName,
    data.partyName,
    data.area || "",
    now.toLocaleString("en-IN"),
    data.lat,
    data.lng,
    data.address || "",
    data.geoVerified ? "YES" : "NO",
    "", "", "",          // checkout fields — filled later
    "", "", "", "", "",  // order/collection/remark — filled on checkout
    Utilities.formatDate(now, "Asia/Kolkata", "yyyy-MM-dd")
  ]);
  
  return jsonResponse({ success: true, visitId: visitId });
}

function handleCheckout(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_VISITS);
  if (!sheet) return jsonResponse({ success: false, error: "No visits sheet" });
  
  const values = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === data.visitId) { rowIndex = i + 1; break; }
  }
  if (rowIndex === -1) return jsonResponse({ success: false, error: "Visit not found" });
  
  const now = new Date();
  const checkInTime = new Date(values[rowIndex - 1][5]);
  const durationMins = Math.round((now - checkInTime) / 60000);
  
  sheet.getRange(rowIndex, 11).setValue(now.toLocaleString("en-IN"));   // checkout time
  sheet.getRange(rowIndex, 12).setValue(data.lat);
  sheet.getRange(rowIndex, 13).setValue(data.lng);
  sheet.getRange(rowIndex, 14).setValue(durationMins);
  sheet.getRange(rowIndex, 15).setValue(Number(data.orderValue) || 0);
  sheet.getRange(rowIndex, 16).setValue(Number(data.collectionValue) || 0);
  sheet.getRange(rowIndex, 17).setValue(data.paymentRemark || "");
  sheet.getRange(rowIndex, 18).setValue(data.notes || "");
  
  return jsonResponse({ success: true, duration: durationMins });
}

function handleSummary(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_VISITS);
  if (!sheet) return jsonResponse({ success: true, today: [], month: { visits: 0, order: 0, collection: 0 } });
  
  const values = sheet.getDataRange().getValues();
  const userId = data.userId;
  const today = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd");
  const thisMonth = today.substring(0, 7); // yyyy-MM
  
  const todayVisits = [];
  let monthOrder = 0, monthCollection = 0, monthVisits = 0;
  
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row[1] !== userId) continue; // filter by user
    const rowDate = String(row[18]);
    
    if (rowDate === today && row[10] !== "") { // checked out today
      todayVisits.push({
        party: row[3],
        checkIn: row[5],
        checkOut: row[10],
        duration: row[13],
        order: row[14],
        collection: row[15],
        remark: row[16]
      });
    }
    if (rowDate.startsWith(thisMonth) && row[10] !== "") {
      monthOrder += Number(row[14]) || 0;
      monthCollection += Number(row[15]) || 0;
      monthVisits++;
    }
  }
  
  return jsonResponse({
    success: true,
    today: todayVisits,
    month: { visits: monthVisits, order: monthOrder, collection: monthCollection }
  });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

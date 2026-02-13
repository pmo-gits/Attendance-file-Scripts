/************************************************
 * 04_Leave_GetBalances.gs
 * FINALIZED: if snapshot missing => pull FULL LEAVE_MASTER rows
 ************************************************/
function getLeaveBalances_Button() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attendanceFileName = ss.getName().trim();

  if (isAttendanceLocked_()) {
    ui.alert("This attendance file is LOCKED.\nNo changes allowed.");
    return;
  }

  const confirm = ui.alert(
    "Lock confirmation",
    "Shall I lock the file and fetch leave balances?\n\n(This action is ONE TIME only.)",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) {
    ui.alert("Cancelled. No changes made.");
    return;
  }

  const idsInAttendance = getAttendanceEmployeeIds_();
  if (idsInAttendance.size === 0) {
    ui.alert("No employee IDs found in Attendance entry sheet.");
    return;
  }

  const rows = computeLeaveBalancesRows_(attendanceFileName);
  if (rows.length === 0) {
    ui.alert("No leave balances found.");
    return;
  }

  writeLeaveBalancesToAttendance_(rows);

  setAttendanceLocked_();

  ui.alert(`LOCKED ✅\nLeave balances updated successfully.\n\nRows: ${rows.length}`);
}

function computeLeaveBalancesRows_(attendanceFileName) {
  const cachedAllForMonth = getSharedSummaryForAttendance_(attendanceFileName);
  if (cachedAllForMonth.length > 0) {
    return cachedAllForMonth;
  }

  const allRows = getAllBalancesFromLeaveMaster_();
  if (allRows.length > 0) {
    appendToSharedSummary_(attendanceFileName, allRows);
  }
  return allRows;
}

function getAttendanceEmployeeIds_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ATTENDANCE_SHEET_NAME);
  const out = new Set();
  if (!sh) return out;

  const last = sh.getLastRow();
  if (last < 2) return out;

  const vals = sh.getRange(2, 2, last - 1, 1).getValues();
  vals.forEach(r => {
    const id = String(r[0] || "").trim();
    if (id) out.add(id);
  });
  return out;
}

function writeLeaveBalancesToAttendance_(rows) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LEAVE_BALANCES_SHEET_NAME);
  if (!sh) throw new Error(`Sheet "${LEAVE_BALANCES_SHEET_NAME}" not found.`);

  sh.clearContents();
  sh.getRange(1, 1, 1, 7).setValues([[
    "ID.NO","NAME","CATEGORY","D.O.J","EL BALANCE","CL BALANCE","SL BALANCE"
  ]]);
  sh.getRange(2, 1, rows.length, 7).setValues(rows);
}


/************************************************
 * 03_Attendance_SyncNew.gs
 ************************************************/
function syncNewEmployeesAppendOnly() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ATTENDANCE_SHEET_NAME);
  if (!sheet) {
    ui.alert(`Sheet "${ATTENDANCE_SHEET_NAME}" not found. Please update ATTENDANCE_SHEET_NAME in script.`);
    return;
  }

  // Block if locked
  if (isAttendanceLocked_()) {
    ui.alert("This attendance file is LOCKED.\nSync New Employees is not allowed.");
    return;
  }

  // Detect month/year from header I1
  const firstHeader = sheet.getRange(1, DATE_START_COL).getDisplayValue();
  const parsedHeader = parseHeaderDate_(firstHeader);
  if (!parsedHeader) {
    ui.alert('Cannot detect month/year from header I1. Please run "Refresh the attendance month - full sheet" once first.');
    return;
  }

  const { year, monthIndex0 } = parsedHeader;
  const daysInMonth = new Date(year, monthIndex0 + 1, 0).getDate();

  // Fetch ACTIVE employees
  const empRows = fetchActiveEmployees_();
  if (empRows.length === 0) {
    ui.alert("No ACTIVE employees found in Employee Master.");
    return;
  }

  // Existing employee codes (Column B)
  const lastRow = sheet.getLastRow();
  const existing = new Set();
  if (lastRow >= 2) {
    const ids = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    ids.forEach(r => {
      const v = String(r[0] || "").trim();
      if (v) existing.add(v);
    });
  }

  // New employees not present in Attendance entry
  const newRows = empRows.filter(r => {
    const empCode = String(r[1] || "").trim();
    return empCode && !existing.has(empCode);
  });

  if (newRows.length === 0) {
    ui.alert("No new ACTIVE employees to append.");
    return;
  }

  const appendStartRow = lastRow + 1;

  // Ensure enough rows
  const needed = appendStartRow + newRows.length - 1;
  if (sheet.getMaxRows() < needed) {
    sheet.insertRowsAfter(sheet.getMaxRows(), needed - sheet.getMaxRows());
  }

  // Append A:H in Attendance entry
  sheet.getRange(appendStartRow, 1, newRows.length, STATIC_COLS_COUNT).setValues(newRows);

  // Apply dropdowns for appended rows (I:AM)
  applyAttendanceDropdownValidation_(sheet, appendStartRow, newRows.length);

  // Fill WO in Sundays for appended rows only
  fillWOSundaysForRows_(sheet, year, monthIndex0, daysInMonth, appendStartRow, newRows.length);

  // Rebuild protections to include new rows
  const totalEmpCount = sheet.getLastRow() - 1;
  applySundayWOFormattingAndProtection_(sheet, year, monthIndex0, daysInMonth, totalEmpCount);

  // ✅ Append same employees into Late Entry
  appendNewEmployeesToLateEntry_(newRows, year, monthIndex0, daysInMonth);
  
  ui.alert(`Appended new employees: ${newRows.length}`);
}

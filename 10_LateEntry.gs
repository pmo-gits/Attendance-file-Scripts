/************************************************
 * 10_LateEntry.gs
 * Late Entry sheet automation
 *
 * Behavior:
 * - On Refresh Month:
 *    - Rewrite A:H roster from empRows (ACTIVE employees)
 *    - Rebuild date headers (I:AM) + hide unused day columns
 *    - Apply Sundays: red header + WO + protection (dynamic to emp count)
 * - On Sync New Employees:
 *    - Append only new employees (A:H)
 *    - Fill WO on Sundays for appended rows
 *    - Rebuild Sunday protections to include new rows
 ************************************************/

function refreshLateEntrySheet_(year, monthIndex0, empRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(LATE_ENTRY_SHEET_NAME);
  if (!sh) {
    throw new Error(`Sheet "${LATE_ENTRY_SHEET_NAME}" not found.`);
  }

  const daysInMonth = new Date(year, monthIndex0 + 1, 0).getDate();

  // Ensure enough rows
  const neededRows = empRows.length + 1; // header + employees
  if (sh.getMaxRows() < neededRows) {
    sh.insertRowsAfter(sh.getMaxRows(), neededRows - sh.getMaxRows());
  }

  // Write roster A:H
  sh.getRange(2, 1, empRows.length, STATIC_COLS_COUNT).setValues(empRows);

  // Build date headers I:AM and hide unused columns
  updateDateHeadersAndVisibility_(sh, year, monthIndex0, daysInMonth);

  // Sundays: red header + WO fill + protect (only till emp rows)
  applySundayWOFormattingAndProtection_(sh, year, monthIndex0, daysInMonth, empRows.length);
}

function appendNewEmployeesToLateEntry_(newEmpRows, year, monthIndex0, daysInMonth) {
  if (!newEmpRows || newEmpRows.length === 0) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(LATE_ENTRY_SHEET_NAME);
  if (!sh) throw new Error(`Sheet "${LATE_ENTRY_SHEET_NAME}" not found.`);

  // Existing employee codes (Column B)
  const lastRow = sh.getLastRow();
  const existing = new Set();
  if (lastRow >= 2) {
    const ids = sh.getRange(2, 2, lastRow - 1, 1).getValues();
    ids.forEach(r => {
      const v = String(r[0] || "").trim();
      if (v) existing.add(v);
    });
  }

  // New employees not present in Late Entry
  const rowsToAppend = newEmpRows.filter(r => {
    const empCode = String(r[1] || "").trim(); // A:H index 1 = Employee Code
    return empCode && !existing.has(empCode);
  });

  if (rowsToAppend.length === 0) return;

  const appendStartRow = sh.getLastRow() + 1;

  // Ensure enough rows
  const needed = appendStartRow + rowsToAppend.length - 1;
  if (sh.getMaxRows() < needed) {
    sh.insertRowsAfter(sh.getMaxRows(), needed - sh.getMaxRows());
  }

  // Append A:H only
  sh.getRange(appendStartRow, 1, rowsToAppend.length, STATIC_COLS_COUNT).setValues(rowsToAppend);

  // Fill WO in Sundays for appended rows only (I:AM)
  fillWOSundaysForRows_(sh, year, monthIndex0, daysInMonth, appendStartRow, rowsToAppend.length);

  // Rebuild Sunday protections to include new rows
  const totalEmpCount = sh.getLastRow() - 1;
  applySundayWOFormattingAndProtection_(sh, year, monthIndex0, daysInMonth, totalEmpCount);
}

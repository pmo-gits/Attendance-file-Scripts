/************************************************
 * 14_RecoveredLedgerPush.gs  (Attendance File)
 *
 * Auto-push snapshot to Salary Advance Master -> Recovered Amount Ledger
 * when Payroll Status is set to LOCKED in "Salary Advance deductions".
 *
 * ✅ Behavior:
 * - Installable onEdit trigger fires when Payroll Status edited to "LOCKED..."
 * - Snapshot includes rows where:
 *     (1) EMI Reference Number is not blank, AND
 *     (2) Payroll Status starts with "LOCKED"
 * - Skip ONLY if master ledger already contains this Attendance File Name in column A
 * - Uses getDisplayValues() so formula results like Recovered Amount = 0 are captured
 *
 * ✅ Option A Support:
 * - bulkLockPayrollStatus_SalaryAdvanceDeductions_(lockText)
 *   -> fills Payroll Status for ALL EMI rows in one write
 *   -> then calls _pushRecoveredLedgerSnapshot_() once
 *
 * 🔁 Change (minimal):
 * - Instead of reading "PLANNED EMI", now reads "RECOVERED AMOUNT"
 * - Writes to Master ledger column "Recovered Amount" (same position as earlier last column)
 ************************************************/

// =======================
// Master constants (kept here to avoid config redeclare issues)
// =======================
const _MASTER_SALARY_ADVANCE_SPREADSHEET_ID = "172uLkW5v1Dr_fYO8rZ3GaAwYEyYyRc_n0PaVYjHDwUg";
const _MASTER_RECOVERED_LEDGER_SHEET_NAME = "Recovered Amount Ledger";

// =======================
// Attendance constants (kept here to avoid config collisions)
// =======================
const _ATT_SALARY_ADVANCE_DEDUCTIONS_SHEET_NAME = "Salary Advance deductions";

/**
 * Run once manually to create installable onEdit trigger.
 */
function setupRecoveredLedgerAutoTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "onEdit_RecoveredLedgerAuto_") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("onEdit_RecoveredLedgerAuto_")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

/**
 * Installable onEdit trigger handler.
 * Fires ONLY when Payroll Status column in Salary Advance deductions is edited to LOCKED.
 */
function onEdit_RecoveredLedgerAuto_(e) {
  try {
    if (!e || !e.range) return;

    const sh = e.range.getSheet();
    if (!sh || sh.getName() !== _ATT_SALARY_ADVANCE_DEDUCTIONS_SHEET_NAME) return;

    const headerMap = _getHeaderMapUpper_(sh);
    const payrollIdx0 = headerMap["PAYROLL STATUS"];
    if (payrollIdx0 == null) return;

    // Only when Payroll Status column edited
    if (e.range.getColumn() !== payrollIdx0 + 1) return;

    const newVal = String(e.value || "").trim().toUpperCase();
    if (!newVal.startsWith("LOCKED")) return;

    _pushRecoveredLedgerSnapshot_();

  } catch (err) {
    // silent for trigger safety
  }
}

/**
 * ✅ OPTION A: Use this from your payroll script.
 *
 * Bulk updates Payroll Status for ALL rows that have EMI Reference Number (not blank),
 * using ONE setValues() call, then pushes the snapshot ONCE.
 *
 * @param {string} lockText Example: "LOCKED - 2026-01-23 17:29"
 */
function bulkLockPayrollStatus_SalaryAdvanceDeductions_(lockText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(_ATT_SALARY_ADVANCE_DEDUCTIONS_SHEET_NAME);
  if (!sh) throw new Error(`Sheet not found: ${_ATT_SALARY_ADVANCE_DEDUCTIONS_SHEET_NAME}`);

  const map = _getHeaderMapUpper_(sh);
  const payrollIdx0 = _reqHeader_(map, "PAYROLL STATUS");
  const refIdx0 = _reqHeader_(map, "EMI REFERENCE NUMBER");

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const numRows = lastRow - 1;

  const lock = String(lockText || "").trim();
  if (!lock.toUpperCase().startsWith("LOCKED")) {
    throw new Error(`lockText must start with "LOCKED". Got: ${lockText}`);
  }

  // Read EMI Reference to decide which rows to lock
  const refVals = sh.getRange(2, refIdx0 + 1, numRows, 1).getValues();

  // Build output for Payroll Status (single write)
  const out = new Array(numRows);
  for (let i = 0; i < numRows; i++) {
    const hasRef = String(refVals[i][0] || "").trim() !== "";
    out[i] = [hasRef ? lock : ""];
  }

  // ✅ single write to Payroll Status column
  sh.getRange(2, payrollIdx0 + 1, numRows, 1).setValues(out);

  // ✅ push once (independent of trigger timing)
  _pushRecoveredLedgerSnapshot_();
}

/**
 * Core push:
 * - Snap rows where EMI Reference exists AND Payroll Status starts with LOCKED
 * - Append to Master -> Recovered Amount Ledger
 * - Skip ONLY if master already contains this attendance file name in col A
 */
function _pushRecoveredLedgerSnapshot_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attendanceFileName = ss.getName(); // e.g., Attendance_January_2026

  const sh = ss.getSheetByName(_ATT_SALARY_ADVANCE_DEDUCTIONS_SHEET_NAME);
  if (!sh) throw new Error(`Sheet not found: ${_ATT_SALARY_ADVANCE_DEDUCTIONS_SHEET_NAME}`);

  const map = _getHeaderMapUpper_(sh);

  const A = {
    month: _reqHeader_(map, "MONTH"),
    emp: _reqHeader_(map, "EMPLOYEE CODE"),
    name: _reqHeader_(map, "NAME"),
    dept: _reqHeader_(map, "DEPARTMENT"),
    ref: _reqHeader_(map, "EMI REFERENCE NUMBER"),
    emiAmt: _reqHeader_(map, "EMI AMOUNT"),
    curBal: _reqHeader_(map, "CURRENT BALANCE"),
    hrDecision: _reqHeader_(map, "HR DECISION"),
    recoveredAmt: _reqHeader_(map, "RECOVERED AMOUNT"), // ✅ CHANGED (was PLANNED EMI)
    payroll: _reqHeader_(map, "PAYROLL STATUS"),
  };

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const lastCol = sh.getLastColumn();

  // ✅ display values capture formula results like Recovered Amount = 0
  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();

  const out = [];
  for (let i = 0; i < values.length; i++) {
    const r = values[i];

    const ref = String(r[A.ref] || "").trim();
    if (!ref) continue;

    const payroll = String(r[A.payroll] || "").trim().toUpperCase();
    if (!payroll.startsWith("LOCKED")) continue;

    out.push([
      attendanceFileName,
      r[A.month] || "",
      r[A.emp] || "",
      r[A.name] || "",
      r[A.dept] || "",
      ref,
      r[A.emiAmt] || "",
      r[A.curBal] || "",
      r[A.hrDecision] || "",
      r[A.recoveredAmt] || "", // ✅ CHANGED (writes Recovered Amount)
    ]);
  }

  if (out.length === 0) return;

  const masterSS = SpreadsheetApp.openById(_MASTER_SALARY_ADVANCE_SPREADSHEET_ID);
  const led = masterSS.getSheetByName(_MASTER_RECOVERED_LEDGER_SHEET_NAME);
  if (!led) throw new Error(`Master sheet not found: ${_MASTER_RECOVERED_LEDGER_SHEET_NAME}`);

  // ✅ skip only if file name already exists in col A
  if (_attendanceFileNameExistsInLedger_(led, attendanceFileName)) return;

  const startRow = _getNextEmptyRowByColA_(led);
  const needLast = startRow + out.length - 1;

  if (led.getMaxRows() < needLast) {
    led.insertRowsAfter(led.getMaxRows(), needLast - led.getMaxRows());
  }

  led.getRange(startRow, 1, out.length, out[0].length).setValues(out);
}

/* =========================
 * Helpers (local, self-contained)
 * ========================= */

function _getHeaderMapUpper_(sh) {
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0]
    .map(h => String(h || "").trim().toUpperCase());
  const map = {};
  headers.forEach((h, i) => { if (h) map[h] = i; });
  return map;
}

function _reqHeader_(map, name) {
  const key = String(name || "").trim().toUpperCase();
  if (!(key in map)) throw new Error(`Missing required header: ${name}`);
  return map[key];
}

function _getNextEmptyRowByColA_(sh) {
  const maxRows = sh.getMaxRows();
  const colA = sh.getRange(1, 1, maxRows, 1).getDisplayValues();
  for (let i = colA.length - 1; i >= 0; i--) {
    if (String(colA[i][0] || "").trim() !== "") return i + 2;
  }
  return 2;
}

function _attendanceFileNameExistsInLedger_(ledgerSheet, attendanceFileName) {
  const lastRow = ledgerSheet.getLastRow();
  if (lastRow < 2) return false;

  const colA = ledgerSheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  const target = String(attendanceFileName || "").trim();

  for (let i = 0; i < colA.length; i++) {
    if (String(colA[i][0] || "").trim() === target) return true;
  }
  return false;
}

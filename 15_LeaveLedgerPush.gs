/************************************************
 * 15_LeaveLedgerPush.gs  (Attendance File)
 *
 * Trigger:
 * - Installable onEdit trigger
 * - Fires when "Salary Advance deductions" -> "Payroll Status" edited
 *   and contains "LOCKED" (case-insensitive)
 *
 * Actions (ONLY STAFF):
 * 1) Snapshot Attendance entry -> Leave Master -> APPROVED_LEAVE_LEDGER
 *    - YEAR + MONTH from Attendance file name (Attendance_<Month>_<Year>)
 *    - Mistake-proof: if same YEAR+MONTH already exists in ledger (bottom scan), SKIP ALL
 * 2) Update Leave Master balances (overwrite + policy):
 *    - Match by key:
 *        Attendance entry: "EMPLOYEE CODE"
 *        Leave Master + Ledger: "ID.NO"
 *    - NEW EL/CL/SL BALANCE -> EL/CL/SL BALANCE (overwrite)
 *
 * Policy additions:
 * - DEC event (first time YEAR+DEC received):
 *    CL & SL reset to 0
 *    EL_CF_REMAINING = Dec closing EL BALANCE (after overwrite)
 *    EL_CF_YEAR = YEAR
 *    EL_JAN_RESET_DONE_YEAR = YEAR + 1
 * - JAN/FEB/MAR:
 *    Reduce EL_CF_REMAINING by CF_USED = min(CF_REMAINING_before, APPROVED EL DAYS)
 *    Record EL_CF_USED + EL_CY_USED in ledger
 * - MAR event (first time YEAR+MAR received):
 *    After overwrite + CF reduction:
 *      EL BALANCE = EL BALANCE - EL_CF_REMAINING
 *      EL_CF_REMAINING = 0
 *      EL_MAR_EXPIRY_DONE_YEAR = YEAR
 *
 * Notes:
 * - Uses global constants from 00_Config.gs (no redeclare)
 * - No function names starting with "_" so they appear in dropdown/trigger UI
 ************************************************/

/**
 * Run once manually to create installable onEdit trigger.
 */
function setupLeaveLedgerAutoTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Remove existing triggers for this handler (safe reset)
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === "onEditLeaveLedgerAuto") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("onEditLeaveLedgerAuto")
    .forSpreadsheet(ss)
    .onEdit()
    .create();
}

/**
 * Installable onEdit trigger handler
 */
function onEditLeaveLedgerAuto(e) {
  try {
    if (!e || !e.range) return;

    const sh = e.range.getSheet();
    if (!sh || sh.getName() !== SALARY_ADVANCE_DEDUCTIONS_SHEET_NAME) return;

    const map = getHeaderMapUpperLeaveLocal(sh);
    const payrollIdx0 = map["PAYROLL STATUS"];
    if (payrollIdx0 == null) return;

    // Only when Payroll Status column edited
    if (e.range.getColumn() !== payrollIdx0 + 1) return;

    const newVal = String(e.value || "").trim().toUpperCase();
    if (!newVal.includes("LOCKED")) return;

    pushLeaveLedgerSnapshotAndUpdateBalances();

  } catch (err) {
    // silent for trigger safety
  }
}

/**
 * Manual + trigger entry:
 * - Derive YEAR+MONTH from Attendance file name
 * - If ledger already has same YEAR+MONTH -> skip all
 * - Read Attendance entry rows (display values)
 * - Filter ONLY STAFF based on Leave Master CATEGORY (key match)
 * - Append snapshot to APPROVED_LEAVE_LEDGER (including EL_CF_USED, EL_CY_USED)
 * - Update balances in Leave Master (overwrite + policy)
 */
function pushLeaveLedgerSnapshotAndUpdateBalances() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attendanceFileName = String(ss.getName() || "").trim();

  const ym = yearMonthFromAttendanceFileNameLeaveLocal(attendanceFileName); // {year:"2026", month:"JAN"}
  if (!ym) return; // fail-safe

  const year = ym.year;
  const month = ym.month;

  // Source sheet: Attendance entry
  const shEntry = ss.getSheetByName(ATTENDANCE_SHEET_NAME);
  if (!shEntry) throw new Error(`Sheet not found: ${ATTENDANCE_SHEET_NAME}`);

  const entryLastRow = shEntry.getLastRow();
  const entryLastCol = shEntry.getLastColumn();
  if (entryLastRow < 2 || entryLastCol < 1) return;

  const entryMap = getHeaderMapUpperLeaveLocal(shEntry);

  // Attendance entry uses "EMPLOYEE CODE" (not ID.NO)
  const A = {
    empCode: reqLeaveLocal(entryMap, "EMPLOYEE CODE"),
    name: reqLeaveLocal(entryMap, "NAME"),
    elEarn: reqLeaveLocal(entryMap, "EL EARNED"),
    clEarn: reqLeaveLocal(entryMap, "CL EARNED"),
    slEarn: reqLeaveLocal(entryMap, "SL EARNED"),
    elBal: reqLeaveLocal(entryMap, "EL BALANCE"),
    clBal: reqLeaveLocal(entryMap, "CL BALANCE"),
    slBal: reqLeaveLocal(entryMap, "SL BALANCE"),
    apEl: reqLeaveLocal(entryMap, "APPROVED EL DAYS"),
    apCl: reqLeaveLocal(entryMap, "APPROVED CL DAYS"),
    apSl: reqLeaveLocal(entryMap, "APPROVED SL DAYS"),
    newEl: reqLeaveLocal(entryMap, "NEW EL BALANCE"),
    newCl: reqLeaveLocal(entryMap, "NEW CL BALANCE"),
    newSl: reqLeaveLocal(entryMap, "NEW SL BALANCE"),
  };

  // Use display values (safe for formula results)
  const entryVals = shEntry.getRange(2, 1, entryLastRow - 1, entryLastCol).getDisplayValues();

  // Open Leave Master
  const leaveSS = SpreadsheetApp.openById(LEAVE_MASTER_SPREADSHEET_ID);

  const shLedger = leaveSS.getSheetById(APPROVED_LEAVE_LEDGER_SHEET_ID);
  if (!shLedger) throw new Error(`Leave Master sheet not found by ID: ${APPROVED_LEAVE_LEDGER_SHEET_ID}`);

  const shMaster = leaveSS.getSheetById(LEAVE_MASTER_SHEET_ID);
  if (!shMaster) throw new Error(`Leave Master sheet not found by ID: ${LEAVE_MASTER_SHEET_ID}`);

  // Mistake-proofing: if same YEAR+MONTH already exists -> skip everything
  if (ledgerHasYearMonthLeaveLocal(shLedger, year, month)) return;

  // Build Leave Master lookup for CATEGORY and row index
  const mLastRow = shMaster.getLastRow();
  const mLastCol = shMaster.getLastColumn();
  if (mLastRow < 2 || mLastCol < 1) return;

  const mMap = getHeaderMapUpperLeaveLocal(shMaster);
  const M = {
    id: reqLeaveLocal(mMap, "ID.NO"),
    category: reqLeaveLocal(mMap, "CATEGORY"),

    elBal: reqLeaveLocal(mMap, "EL BALANCE"),
    clBal: reqLeaveLocal(mMap, "CL BALANCE"),
    slBal: reqLeaveLocal(mMap, "SL BALANCE"),

    // New helper columns (must exist as per your updated structure)
    cfRem: reqLeaveLocal(mMap, "EL_CF_REMAINING"),
    cfYear: reqLeaveLocal(mMap, "EL_CF_YEAR"),
    janDone: reqLeaveLocal(mMap, "EL_JAN_RESET_DONE_YEAR"),
    marDone: reqLeaveLocal(mMap, "EL_MAR_EXPIRY_DONE_YEAR"),
  };

  const masterVals = shMaster.getRange(2, 1, mLastRow - 1, mLastCol).getDisplayValues();

  const masterById = new Map(); // ID.NO -> {idx0, category}
  for (let i = 0; i < masterVals.length; i++) {
    const id = String(masterVals[i][M.id] || "").trim();
    if (!id) continue;
    const cat = String(masterVals[i][M.category] || "").trim().toUpperCase();
    masterById.set(id, { idx0: i, category: cat });
  }

  // Ensure ledger has required headers (append if missing)
  let ledMap = getHeaderMapUpperLeaveLocal(shLedger);

  // Final ledger structure required (your confirmed structure)
  const requiredHeaders = [
    "ID.NO","NAME",
    "EL EARNED","CL EARNED","SL EARNED",
    "EL BALANCE","CL BALANCE","SL BALANCE",
    "APPROVED EL DAYS","APPROVED CL DAYS","APPROVED SL DAYS",
    "NEW EL BALANCE","NEW CL BALANCE","NEW SL BALANCE",
    "YEAR","MONTH",
    "EL_CF_USED","EL_CY_USED"
  ];

  let mutated = false;
  requiredHeaders.forEach(h => {
    if (ledMap[h] == null) {
      shLedger.getRange(1, shLedger.getLastColumn() + 1).setValue(h);
      mutated = true;
    }
  });
  if (mutated) ledMap = getHeaderMapUpperLeaveLocal(shLedger);

  const isDec = String(month || "").toUpperCase() === "DEC";
  const isMar = String(month || "").toUpperCase() === "MAR";
  const monthIdx = monthIndexFromAbbrLeaveLocal(month); // 0=JAN ... 11=DEC
  const yNum = toIntLeaveLocal(year);
  const nextYear = String((isNaN(yNum) ? "" : (yNum + 1)));
  const prevYear = String((isNaN(yNum) ? "" : (yNum - 1)));

  // Prepare snapshot rows (ONLY STAFF) + updates
  const outAligned = [];
  const updates = []; // for master overwrite

  for (let i = 0; i < entryVals.length; i++) {
    const r = entryVals[i];

    const idKey = String(r[A.empCode] || "").trim(); // EMPLOYEE CODE used as ID.NO key
    if (!idKey) continue;

    const info = masterById.get(idKey);
    if (!info || info.category !== "STAFF") continue;

    const idx0 = info.idx0;
    const mRow = masterVals[idx0];

    // Attendance values (strings)
    const apElStr = r[A.apEl] || "";
    const newElStr = r[A.newEl] || "";
    const newClStr = r[A.newCl] || "";
    const newSlStr = r[A.newSl] || "";

    // Master helper context
    const cfRemBefore = toNumberLeaveLocal(mRow[M.cfRem]);
    const cfYearVal = String(mRow[M.cfYear] || "").trim();
    const apElNum = toNumberLeaveLocal(apElStr);

    // Determine whether CF applies for this month:
    // - CF is from previous year (cfYear == year-1) and month is Jan/Feb/Mar (0/1/2)
    const cfApplies = (cfYearVal && cfYearVal === prevYear && monthIdx >= 0 && monthIdx <= 2);

    // Compute CF usage audit
    const cfUsed = cfApplies ? Math.min(Math.max(cfRemBefore, 0), Math.max(apElNum, 0)) : 0;
    const cyUsed = Math.max(apElNum, 0) - cfUsed;

    // Build one ledger row by header mapping
    const row = new Array(shLedger.getLastColumn()).fill("");
    setIfColLeaveLocal(row, ledMap, "ID.NO", idKey);
    setIfColLeaveLocal(row, ledMap, "NAME", r[A.name] || "");
    setIfColLeaveLocal(row, ledMap, "EL EARNED", r[A.elEarn] || "");
    setIfColLeaveLocal(row, ledMap, "CL EARNED", r[A.clEarn] || "");
    setIfColLeaveLocal(row, ledMap, "SL EARNED", r[A.slEarn] || "");
    setIfColLeaveLocal(row, ledMap, "EL BALANCE", r[A.elBal] || "");
    setIfColLeaveLocal(row, ledMap, "CL BALANCE", r[A.clBal] || "");
    setIfColLeaveLocal(row, ledMap, "SL BALANCE", r[A.slBal] || "");
    setIfColLeaveLocal(row, ledMap, "APPROVED EL DAYS", apElStr);
    setIfColLeaveLocal(row, ledMap, "APPROVED CL DAYS", r[A.apCl] || "");
    setIfColLeaveLocal(row, ledMap, "APPROVED SL DAYS", r[A.apSl] || "");
    setIfColLeaveLocal(row, ledMap, "NEW EL BALANCE", newElStr);
    setIfColLeaveLocal(row, ledMap, "NEW CL BALANCE", newClStr);
    setIfColLeaveLocal(row, ledMap, "NEW SL BALANCE", newSlStr);
    setIfColLeaveLocal(row, ledMap, "YEAR", year);
    setIfColLeaveLocal(row, ledMap, "MONTH", month);
    setIfColLeaveLocal(row, ledMap, "EL_CF_USED", numberToCellLeaveLocal(cfUsed));
    setIfColLeaveLocal(row, ledMap, "EL_CY_USED", numberToCellLeaveLocal(cyUsed));

    outAligned.push(row);

    updates.push({
      idx0: idx0,
      idKey: idKey,

      // Attendance new balances (truth overwrite)
      newEl: newElStr,
      newCl: newClStr,
      newSl: newSlStr,

      // Approved EL for CF reduction
      apElNum: apElNum,

      // Master helper context
      cfRemBefore: cfRemBefore,
      cfYearVal: cfYearVal,
      cfApplies: cfApplies,

      // Flags
      janDoneYear: String(mRow[M.janDone] || "").trim(),
      marDoneYear: String(mRow[M.marDone] || "").trim(),

      // Audit (for convenience)
      cfUsed: cfUsed,
      cyUsed: cyUsed,
    });
  }

  if (!outAligned.length) return;

  // Append snapshot
  const startRow = nextEmptyRowByColALeaveLocal(shLedger);
  ensureRowsLeaveLocal(shLedger, startRow + outAligned.length - 1);
  shLedger.getRange(startRow, 1, outAligned.length, outAligned[0].length).setValues(outAligned);

  // Update balances in Leave Master (overwrite + policy)
  const elRange = shMaster.getRange(2, M.elBal + 1, masterVals.length, 1);
  const clRange = shMaster.getRange(2, M.clBal + 1, masterVals.length, 1);
  const slRange = shMaster.getRange(2, M.slBal + 1, masterVals.length, 1);

  const cfRange = shMaster.getRange(2, M.cfRem + 1, masterVals.length, 1);
  const cfYearRange = shMaster.getRange(2, M.cfYear + 1, masterVals.length, 1);
  const janDoneRange = shMaster.getRange(2, M.janDone + 1, masterVals.length, 1);
  const marDoneRange = shMaster.getRange(2, M.marDone + 1, masterVals.length, 1);

  const elCol = elRange.getDisplayValues();
  const clCol = clRange.getDisplayValues();
  const slCol = slRange.getDisplayValues();

  const cfCol = cfRange.getDisplayValues();
  const cfYearCol = cfYearRange.getDisplayValues();
  const janDoneCol = janDoneRange.getDisplayValues();
  const marDoneCol = marDoneRange.getDisplayValues();

  let changed = false;

  for (const u of updates) {
    const j = u.idx0;
    if (j == null || j < 0 || j >= masterVals.length) continue;

    // 1) Always overwrite EL/CL/SL with Attendance NEW values first (truth)
    // (But DEC reset may override CL/SL after)
    if (String(elCol[j][0] || "") !== String(u.newEl || "")) { elCol[j][0] = u.newEl; changed = true; }
    if (String(clCol[j][0] || "") !== String(u.newCl || "")) { clCol[j][0] = u.newCl; changed = true; }
    if (String(slCol[j][0] || "") !== String(u.newSl || "")) { slCol[j][0] = u.newSl; changed = true; }

    // 2) JAN/FEB/MAR CF reduction (only if CF applies)
    // CF_REMAINING = CF_REMAINING - CF_USED
    if (u.cfApplies) {
      const cfAfter = Math.max(0, (u.cfRemBefore || 0) - (u.cfUsed || 0));
      if (String(cfCol[j][0] || "") !== String(numberToCellLeaveLocal(cfAfter))) {
        cfCol[j][0] = numberToCellLeaveLocal(cfAfter);
        changed = true;
      }
    }

    // 3) DEC event: reset CL/SL + set CF for next year (only once per employee per cycle)
    // Trigger condition: current month is DEC and jan-reset for nextYear not yet done.
    // After Dec overwrite, set CF_REMAINING = Dec closing EL BALANCE (current elCol value)
    if (isDec && nextYear && String(janDoneCol[j][0] || "").trim() !== nextYear) {
      // Reset CL/SL to 0
      if (String(clCol[j][0] || "") !== "0") { clCol[j][0] = "0"; changed = true; }
      if (String(slCol[j][0] || "") !== "0") { slCol[j][0] = "0"; changed = true; }

      // Set carry-forward for next year cycle from Dec closing EL (after overwrite)
      const decClosingEl = toNumberLeaveLocal(elCol[j][0]);
      if (String(cfCol[j][0] || "") !== String(numberToCellLeaveLocal(decClosingEl))) {
        cfCol[j][0] = numberToCellLeaveLocal(decClosingEl);
        changed = true;
      }
      if (String(cfYearCol[j][0] || "").trim() !== String(year)) {
        cfYearCol[j][0] = String(year);
        changed = true;
      }

      // Mark jan reset done for next year
      janDoneCol[j][0] = nextYear;
      changed = true;
    }

    // 4) MAR expiry: after overwrite + CF reduction, expire remaining CF (only once per YEAR)
    if (isMar && String(marDoneCol[j][0] || "").trim() !== String(year)) {
      // Only expire if CF belongs to previous year and we're in the correct cycle
      const cfYearVal = String(cfYearCol[j][0] || "").trim();
      const cfNow = toNumberLeaveLocal(cfCol[j][0]);
      const elNow = toNumberLeaveLocal(elCol[j][0]);

      // Expiry should apply to prior-year CF when in Jan-Mar window of current year
      // i.e. cfYear == year-1
      const shouldExpire = (cfYearVal && cfYearVal === prevYear);

      if (shouldExpire && cfNow > 0) {
        const elAfterExpiry = elNow - cfNow;
        elCol[j][0] = numberToCellLeaveLocal(elAfterExpiry);
        cfCol[j][0] = "0";
        changed = true;
      }

      // Mark expiry done for this year (even if cf was 0; prevents repeats)
      marDoneCol[j][0] = String(year);
      changed = true;
    }
  }

  if (changed) {
    elRange.setValues(elCol);
    clRange.setValues(clCol);
    slRange.setValues(slCol);

    cfRange.setValues(cfCol);
    cfYearRange.setValues(cfYearCol);
    janDoneRange.setValues(janDoneCol);
    marDoneRange.setValues(marDoneCol);
  }
}

/* =========================
 * Mistake-proofing: check YEAR+MONTH exists in ledger (bottom scan)
 * ========================= */
function ledgerHasYearMonthLeaveLocal(shLedger, year, month) {
  const map = getHeaderMapUpperLeaveLocal(shLedger);
  const yIdx0 = map["YEAR"];
  const mIdx0 = map["MONTH"];
  if (yIdx0 == null || mIdx0 == null) return false;

  const lastRow = shLedger.getLastRow();
  if (lastRow < 2) return false;

  const yVals = shLedger.getRange(2, yIdx0 + 1, lastRow - 1, 1).getDisplayValues();
  const mVals = shLedger.getRange(2, mIdx0 + 1, lastRow - 1, 1).getDisplayValues();

  const yT = String(year || "").trim();
  const mT = String(month || "").trim().toUpperCase();

  for (let i = yVals.length - 1; i >= 0; i--) {
    const y = String(yVals[i][0] || "").trim();
    const m = String(mVals[i][0] || "").trim().toUpperCase();
    if (!y && !m) continue;
    if (y === yT && m === mT) return true;
  }
  return false;
}

/* =========================
 * YEAR+MONTH from file name: Attendance_January_2026 -> {year:"2026", month:"JAN"}
 * ========================= */
function yearMonthFromAttendanceFileNameLeaveLocal(fileName) {
  const m = String(fileName || "").trim().match(/^Attendance[_\-\s]+([A-Za-z]+)[_\-\s]+(\d{4})$/i);
  if (!m) return null;

  const mon = String(m[1] || "").trim().toLowerCase();
  const year = String(m[2] || "").trim();

  const idx = monthIndexFromNameLeaveLocal(mon);
  if (idx < 0) return null;

  const abbr = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"][idx];
  return { year, month: abbr };
}

function monthIndexFromNameLeaveLocal(monLower) {
  const months = [
    "january","february","march","april","may","june",
    "july","august","september","october","november","december"
  ];
  return months.indexOf(String(monLower || "").trim());
}

function monthIndexFromAbbrLeaveLocal(monAbbr) {
  const abbrs = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
  return abbrs.indexOf(String(monAbbr || "").trim().toUpperCase());
}

/* =========================
 * Local helpers (unique names to avoid collisions)
 * ========================= */
function getHeaderMapUpperLeaveLocal(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return {};
  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0]
    .map(h => String(h || "").trim().toUpperCase());
  const map = {};
  headers.forEach((h, i) => { if (h) map[h] = i; });
  return map;
}

function reqLeaveLocal(map, header) {
  const key = String(header || "").trim().toUpperCase();
  if (!(key in map)) throw new Error(`Missing required header: ${header}`);
  return map[key];
}

function setIfColLeaveLocal(rowArr, map, header, value) {
  const idx0 = map[String(header || "").trim().toUpperCase()];
  if (idx0 != null) rowArr[idx0] = value;
}

function nextEmptyRowByColALeaveLocal(sh) {
  const maxRows = sh.getMaxRows();
  if (maxRows < 2) return 2;

  const colA = sh.getRange(1, 1, maxRows, 1).getDisplayValues();
  for (let i = colA.length - 1; i >= 0; i--) {
    if (String(colA[i][0] || "").trim() !== "") return i + 2;
  }
  return 2;
}

function ensureRowsLeaveLocal(sh, neededLastRow) {
  const cur = sh.getMaxRows();
  if (cur < neededLastRow) sh.insertRowsAfter(cur, neededLastRow - cur);
}

function toNumberLeaveLocal(v) {
  const s = String(v == null ? "" : v).replace(/,/g, "").trim();
  if (!s) return 0;
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function toIntLeaveLocal(v) {
  const s = String(v == null ? "" : v).trim();
  const n = parseInt(s, 10);
  return isNaN(n) ? NaN : n;
}

function numberToCellLeaveLocal(n) {
  // Keep as string to preserve display-values consistency
  if (n == null || isNaN(n)) return "0";
  // Avoid trailing .0 noise where possible, but keep decimals if present
  const x = Math.round(n * 1000000) / 1000000;
  return String(x);
}

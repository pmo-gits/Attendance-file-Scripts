/************************************************
 * 11_LateEntry_PenaltyCalc.gs
 * Calculates LEAVE PENALTY COUNT for Late Entry sheet
 * FINAL RULES IMPLEMENTED
 * - Can run ONLY from "Late Entry" sheet
 ************************************************/

function recalcLateEntryLeavePenaltyCount() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ✅ NEW: Allow run only from Late Entry sheet
  const activeSheet = ss.getActiveSheet();
  if (!activeSheet || activeSheet.getName() !== "Late Entry") {
    ui.alert('Please run this function from the "Late Entry" sheet only.');
    return;
  }

  const sh = ss.getSheetByName(LATE_ENTRY_SHEET_NAME);
  if (!sh) {
    ui.alert(`Sheet "${LATE_ENTRY_SHEET_NAME}" not found.`);
    return;
  }

  const lastCol = sh.getLastColumn();
  if (lastCol < 1) {
    ui.alert("Late Entry sheet is empty.");
    return;
  }

  const header = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0]
    .map(h => String(h || "").trim());

  const idx = (name) => header.findIndex(h => h.toUpperCase() === String(name).toUpperCase());

  const iEmpCode = idx("EMPLOYEE CODE");
  const iCategory = idx("CATEGORY");
  const iPenalty = idx("LATE PENALTY COUNT");
  const iDOJ = idx("DOJ");
  const iTotalLate = idx("TOTAL LATE DURATION");

  if (iEmpCode === -1) throw new Error('Late Entry: "Employee Code" not found.');
  if (iCategory === -1) throw new Error('Late Entry: "Category" not found.');
  if (iPenalty === -1) throw new Error('Late Entry: "Late Penalty Count" not found.');
  if (iDOJ === -1) throw new Error('Late Entry: "DOJ" not found.');
  if (iTotalLate === -1) throw new Error('Late Entry: "Total Late Duration" not found.');

  const dateStartCol1 = iDOJ + 2;
  const dateEndCol1 = iTotalLate;
  if (dateEndCol1 < dateStartCol1) {
    throw new Error("Late Entry: Date columns not detected.");
  }

  const maxRows = sh.getLastRow();
  if (maxRows < 2) {
    ui.alert("No employee rows found.");
    return;
  }

  const empCodes = sh.getRange(2, iEmpCode + 1, maxRows - 1, 1).getValues();
  let lastEmpIndex0 = -1;
  for (let i = empCodes.length - 1; i >= 0; i--) {
    if (String(empCodes[i][0] || "").trim() !== "") {
      lastEmpIndex0 = i;
      break;
    }
  }

  if (lastEmpIndex0 === -1) {
    ui.alert("No Employee Code values found.");
    return;
  }

  const rowsToProcess = lastEmpIndex0 + 1;
  const data = sh.getRange(2, 1, rowsToProcess, lastCol).getValues();
  const outPenalty = new Array(rowsToProcess).fill(0).map(() => [""]);

  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    const empCode = String(row[iEmpCode] || "").trim();
    if (!empCode) continue;

    const cat = String(row[iCategory] || "").trim().toUpperCase();
    const isStaff = cat === "STAFF";
    const isWorker = cat === "WORKER";

    let permSlots = isStaff ? 2 : (isWorker ? 1 : 0);
    let graceUsed = 0;
    let penalty = 0;

    for (let c = dateStartCol1 - 1; c <= dateEndCol1 - 1; c++) {
      const v = row[c];
      if (v === "" || v === null) continue;

      const t = String(v).trim().toUpperCase();
      if (t === "WO") continue;

      if (t === "P60") {
        if (permSlots >= 1) {
          permSlots -= 1;
          continue;
        }
        applyLate_(60);
        continue;
      }

      if (t === "P120") {
        if (isStaff && permSlots >= 2) {
          permSlots -= 2;
          continue;
        }
        applyLate_(120);
        continue;
      }

      const num = parseFloat(String(v).replace(/,/g, "").trim());
      if (!isFinite(num)) continue;

      applyLate_(num);
    }

    outPenalty[r] = [penalty];

    function applyLate_(minutes) {
      if (minutes < 5) return;

      // 5–10 min: grace only
      if (minutes <= 10) {
        if (graceUsed < 3) {
          graceUsed += 1;
          return;
        }
        // grace exhausted → penalty applies
      }

      // >10 min OR grace exhausted
      penalty += (minutes > 270) ? 1 : 0.5;
    }
  }

  sh.getRange(2, iPenalty + 1, outPenalty.length, 1).setValues(outPenalty);
  ui.alert(`Late penalty calculated ✅\nEmployees processed: ${outPenalty.length}`);
}

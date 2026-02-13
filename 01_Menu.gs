/************************************************
 * 01_Menu.gs
 ************************************************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("Attendance")
    .addItem("Refresh the attendance month - full sheet", "refreshAttendanceMonth_FullSheet")
    .addSeparator()
    .addItem("Sync New Employees - add new rows", "syncNewEmployeesAppendOnly")
    .addToUi();

  ui.createMenu("Leave")
    .addItem("Get Leave Balances", "getLeaveBalances_Button")
    .addSeparator()
    .addItem("Calculate late penalty", "recalcLateEntryLeavePenaltyCount")
    .addToUi();

  // ✅ Salary Advance
  ui.createMenu("Salary Advance")
    .addItem("Get Salary Advance EMI", "getSalaryAdvanceEMI_Button")
    .addSeparator()
    .addItem("Refresh Salary Advance EMI", "refreshSalaryAdvanceEMI_Button")
    .addToUi();
}

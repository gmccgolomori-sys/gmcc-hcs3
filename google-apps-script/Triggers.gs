/**
 * Triggers.gs - Scheduled Daily Automation
 */

function installDailyTriggers() {
  ScriptApp.newTrigger('checkContractExpiriesDaily')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .inTimezone('Asia/Makassar')
    .create();
}

function checkContractExpiriesDaily() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Employment_Contracts');
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  
  const thresholds = [90, 60, 30, 14, 7, 0];

  for (let i = 1; i < data.length; i++) {
    const contract = data[i];
    const endDateStr = contract[8];
    const status = contract[9];

    if (endDateStr && status === 'ACTIVE') {
      const endDate = new Date(endDateStr);
      const diffTime = endDate.getTime() - today.getTime();
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

      if (thresholds.includes(diffDays)) {
        NotificationService.sendContractExpiryAlert(contract[1], contract[3], diffDays);
        if (diffDays <= 30) {
          sheet.getRange(i + 1, 10).setValue('EXPIRING_SOON');
        }
      }
    }
  }
}
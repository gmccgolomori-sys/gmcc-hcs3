/**
 * LeaveApprovalService.gs
 * Sequential Leave Workflow: Staff -> Dept Head -> HC -> General Manager -> Final Deduction
 */

const LeaveApprovalService = {
  processApproval: function(applicationId, action, comment) {
    const user = AuthService.getCurrentUser();
    const lock = LockService.getScriptLock();
    lock.waitLock(15000); // Prevent concurrent duplicate deductions

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const appSheet = ss.getSheetByName('Leave_Applications');
      const data = appSheet.getDataRange().getValues();

      let targetRow = -1;
      let app = null;

      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === applicationId) {
          targetRow = i + 1;
          app = {
            id: data[i][0],
            employeeId: data[i][1],
            leaveType: data[i][4],
            startDate: data[i][5],
            endDate: data[i][6],
            totalDays: Number(data[i][7]),
            status: data[i][9]
          };
          break;
        }
      }

      if (!app) throw new Error('Application ' + applicationId + ' not found.');

      // Check stage permissions
      if (app.status === 'PENDING_DEPT_HEAD') {
        if (user.role !== 'DEPT_HEAD' && user.role !== 'ADMIN') {
          throw new Error('Only Department Head can approve at Stage 1.');
        }
        const nextStatus = (action === 'APPROVE') ? 'PENDING_HC' : (action === 'REJECT' ? 'REJECTED_BY_DEPT_HEAD' : 'RETURNED_TO_EMPLOYEE');
        appSheet.getRange(targetRow, 10).setValue(nextStatus);
        appSheet.getRange(targetRow, 11).setValue(action);
        appSheet.getRange(targetRow, 12).setValue(new Date());
        appSheet.getRange(targetRow, 13).setValue(comment);
      }
      else if (app.status === 'PENDING_HC') {
        if (user.role !== 'HC' && user.role !== 'ADMIN') {
          throw new Error('Only Human Capital can verify at Stage 2.');
        }
        const nextStatus = (action === 'APPROVE') ? 'PENDING_GM' : (action === 'REJECT' ? 'REJECTED_BY_HC' : 'RETURNED_TO_EMPLOYEE');
        appSheet.getRange(targetRow, 10).setValue(nextStatus);
        appSheet.getRange(targetRow, 14).setValue(action);
        appSheet.getRange(targetRow, 15).setValue(new Date());
        appSheet.getRange(targetRow, 16).setValue(comment);
      }
      else if (app.status === 'PENDING_GM') {
        if (user.role !== 'GM' && user.role !== 'ADMIN') {
          throw new Error('Only General Manager can grant final sanction.');
        }

        if (action === 'APPROVE') {
          // CRITICAL: Execute official ledger deduction ONLY now!
          const txnId = 'TXN-' + new Date().getTime();
          const ledgerSheet = ss.getSheetByName('Leave_Ledger');
          ledgerSheet.appendRow([
            txnId,
            app.employeeId,
            app.leaveType,
            'APPROVED_LEAVE_DEDUCTION',
            -app.totalDays, // Negative debit
            app.startDate,
            '',
            app.id,
            user.employeeId || user.email,
            user.fullName,
            app.id,
            'Approved by General Manager: ' + comment,
            new Date()
          ]);

          appSheet.getRange(targetRow, 10).setValue('APPROVED');
          appSheet.getRange(targetRow, 17).setValue('APPROVE');
          appSheet.getRange(targetRow, 18).setValue(new Date());
          appSheet.getRange(targetRow, 19).setValue(comment);
          appSheet.getRange(targetRow, 20).setValue(txnId); // Link deduction
        } else {
          const nextStatus = (action === 'REJECT') ? 'REJECTED_BY_GM' : 'RETURNED_TO_EMPLOYEE';
          appSheet.getRange(targetRow, 10).setValue(nextStatus);
          appSheet.getRange(targetRow, 17).setValue(action);
          appSheet.getRange(targetRow, 18).setValue(new Date());
          appSheet.getRange(targetRow, 19).setValue(comment);
        }
      }

      AuditService.log('LEAVE', 'LEAVE_' + action, 'Leave ' + applicationId + ' processed by ' + user.fullName);
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  }
};
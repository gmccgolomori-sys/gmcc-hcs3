/**
 * AuthService.gs - Server-Side Identity Verification & Access Control
 */

const AuthService = {
  getCurrentUser: function() {
    const email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Users_Roles');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toLowerCase() === email.toLowerCase() && data[i][5] === true) {
        return {
          email: data[i][0],
          fullName: data[i][1],
          role: data[i][2], // 'STAFF' | 'DEPT_HEAD' | 'HC' | 'GM' | 'ADMIN' | 'AUDITOR'
          employeeId: data[i][3],
          departmentId: data[i][4]
        };
      }
    }

    // Default fallback if unmapped
    return {
      email: email,
      fullName: 'GMCC Staff',
      role: 'STAFF',
      employeeId: '',
      departmentId: ''
    };
  },

  assertRole: function(allowedRoles) {
    const user = this.getCurrentUser();
    if (!allowedRoles.includes(user.role)) {
      throw new Error('Access Denied. Role ' + user.role + ' is unauthorized.');
    }
    return user;
  }
};
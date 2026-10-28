/**
 * Code.gs - Entry point for Google Apps Script Web App & REST API Sync
 */

function doGet(e) {
  // 1. Tarik seluruh data dari Google Sheets ke aplikasi GMCC (PULL)
  if (e && e.parameter && (e.parameter.action === 'GET_DATA' || e.parameter.action === 'PULL_ALL')) {
    const sheetData = getAllSheetData();
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      data: sheetData,
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // 2. Standar respons status online
  const responseData = {
    status: 'ONLINE',
    system: 'GMCC Human Capital Management System - Google Apps Script Service',
    message: 'Backend Google Apps Script siap menerima sinkronisasi data dan Google Drive terhubung!',
    timestamp: new Date().toISOString()
  };

  return ContentService.createTextOutput(JSON.stringify(responseData))
    .setMimeType(ContentService.MimeType.JSON);
}

function getAllSheetData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  function getRows(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    const vals = sheet.getDataRange().getValues();
    if (vals.length <= 1) return [];
    const headers = vals[0];
    const rows = [];
    for (let r = 1; r < vals.length; r++) {
      const obj = {};
      for (let c = 0; c < headers.length; c++) {
        let cellVal = vals[r][c];
        if (cellVal instanceof Date) {
          cellVal = Utilities.formatDate(cellVal, 'Asia/Makassar', 'yyyy-MM-dd');
        }
        obj[headers[c]] = cellVal;
      }
      rows.push(obj);
    }
    return rows;
  }

  return {
    employees: getRows('Employees'),
    contracts: getRows('Employment_Contracts'),
    trainingPlans: getRows('Training_Plans'),
    leaveApplications: getRows('Leave_Applications'),
    documents: getRows('Employee_Documents')
  };
}

function doPost(e) {
  try {
    let postData = {};
    if (e && e.postData && e.postData.contents) {
      try {
        postData = JSON.parse(e.postData.contents);
      } catch (err) {
        postData = {};
      }
    } else if (e && e.parameter && e.parameter.data) {
      try {
        postData = JSON.parse(e.parameter.data);
      } catch (err) {
        postData = {};
      }
    }

    const action = postData.action || (e && e.parameter && e.parameter.action) || (postData.data || postData.employees ? 'SYNC_ALL' : 'PING');

    if (action === 'PING') {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: 'Koneksi GMCC Google Apps Script Berhasil!',
        timestamp: new Date().toISOString()
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'UPLOAD_DOCUMENT') {
      const doc = postData.document || postData;
      let driveResult = null;
      if (doc.fileData && typeof PersonalFileService !== 'undefined' && PersonalFileService.uploadDocumentToDrive) {
        try {
          driveResult = PersonalFileService.uploadDocumentToDrive(
            doc.employeeId,
            doc.employeeName || 'Staff',
            doc.category,
            doc.fileName,
            doc.fileData,
            doc.mimeType
          );
        } catch (err) {
          Logger.log('Drive file upload error: ' + err);
        }
      }

      // Catat record juga ke lembar kerja Employee_Documents di Google Sheet
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const docSheet = getOrCreateSheet('Employee_Documents', [
        'DocumentID', 'EmployeeID', 'EmployeeName', 'Category', 'Title', 'FileName',
        'DriveFileID', 'DriveFileUrl', 'DrivePath', 'UploadedBy', 'UploadedAt'
      ]);
      docSheet.appendRow([
        doc.id || ('DOC-' + Date.now()),
        doc.employeeId,
        doc.employeeName || '',
        doc.category,
        doc.title || doc.fileName,
        doc.fileName,
        driveResult ? driveResult.fileId : (doc.driveFileId || ''),
        driveResult ? driveResult.fileUrl : '',
        driveResult ? driveResult.path : (doc.drivePath || ''),
        doc.uploadedByName || doc.uploadedBy || 'Admin',
        new Date().toISOString()
      ]);

      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: 'File image/dokumen berhasil disimpan ke Google Drive!',
        data: driveResult
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'UPLOAD_TRAINING_ATTENDANCE') {
      const att = postData.attendance || postData;
      let drivePhotoResult = null;
      if (att.photoProofUrl && (att.photoProofUrl.indexOf('data:') === 0 || att.photoProofUrl.indexOf('http') === -1)) {
        try {
          const rootName = 'GMCC HUMAN CAPITAL';
          const roots = DriveApp.getFoldersByName(rootName);
          const root = roots.hasNext() ? roots.next() : DriveApp.createFolder(rootName);
          
          const tdFolders = root.getFoldersByName('03 TRAINING & DEVELOPMENT');
          const tdFolder = tdFolders.hasNext() ? tdFolders.next() : root.createFolder('03 TRAINING & DEVELOPMENT');
          
          const presensiFolders = tdFolder.getFoldersByName('PRESENSI & BUKTI FOTO');
          const presensiFolder = presensiFolders.hasNext() ? presensiFolders.next() : tdFolder.createFolder('PRESENSI & BUKTI FOTO');
          
          const cleanPhotoName = 'PRESENSI_' + (att.employeeId || 'EMP') + '_' + (att.sessionId || 'TR') + '_' + new Date().getTime() + '.jpg';
          let cleanBase64 = att.photoProofUrl;
          let mime = 'image/jpeg';
          if (att.photoProofUrl.indexOf(',') > -1) {
            const parts = att.photoProofUrl.split(',');
            cleanBase64 = parts[1];
            const mimeMatch = parts[0].match(/data:([^;]+);/);
            if (mimeMatch) mime = mimeMatch[1];
          }
          const blob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), mime, cleanPhotoName);
          const photoFile = presensiFolder.createFile(blob);
          try { photoFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
          drivePhotoResult = {
            fileId: photoFile.getId(),
            fileUrl: photoFile.getUrl()
          };
        } catch (err) {
          Logger.log('Error saving attendance photo to Drive: ' + err);
        }
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const attSheet = getOrCreateSheet('Training_Attendance', [
        'AttendanceID', 'SessionID', 'SessionTitle', 'TrainingScope', 'EmployeeID', 
        'EmployeeName', 'DepartmentID', 'DepartmentName', 'AttendanceDate', 'CheckInTime', 
        'TrainingHours', 'Location', 'PhotoDriveFileID', 'PhotoDriveUrl', 'Status', 
        'Notes', 'VerifiedBy', 'VerifiedAt', 'CreatedAt'
      ]);

      attSheet.appendRow([
        att.id || ('ATT-' + Date.now()),
        att.sessionId || '',
        att.sessionTitle || '',
        att.trainingScope || '',
        att.employeeId || '',
        att.employeeName || '',
        att.departmentId || '',
        att.departmentName || '',
        att.attendanceDate || Utilities.formatDate(new Date(), 'Asia/Makassar', 'yyyy-MM-dd'),
        att.checkInTime || Utilities.formatDate(new Date(), 'Asia/Makassar', 'HH:mm'),
        att.trainingHours || 4,
        att.location || 'GMCC Training Hall',
        drivePhotoResult ? drivePhotoResult.fileId : '',
        drivePhotoResult ? drivePhotoResult.fileUrl : (att.photoProofUrl || ''),
        att.status || 'VERIFIED',
        att.notes || '',
        att.verifiedBy || 'System Verified',
        att.verifiedAt || new Date().toISOString(),
        new Date().toISOString()
      ]);

      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: 'Presensi training & bukti foto berhasil disimpan ke Google Drive dan Google Sheets!',
        drivePhoto: drivePhotoResult
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'SUBMIT_LEAVE_APPLICATION') {
      const app = postData.application || postData;
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const leaveSheet = getOrCreateSheet('Leave_Applications', [
        'ApplicationID', 'EmployeeID', 'EmployeeName', 'DepartmentID', 'LeaveType', 
        'StartDate', 'EndDate', 'TotalDays', 'Reason', 'Status', 
        'DeptHeadDecision', 'DeptHeadDate', 'DeptHeadNotes', 
        'HCDecision', 'HCDate', 'HCNotes', 
        'GMDecision', 'GMDate', 'GMNotes', 
        'DeductionTxnID', 'SubmittedAt', 'UpdatedAt'
      ]);

      leaveSheet.appendRow([
        app.id,
        app.employeeId,
        app.employeeName,
        app.departmentId,
        app.leaveType,
        app.startDate,
        app.endDate,
        app.totalDays,
        app.reason || '',
        app.status || 'PENDING_DEPT_HEAD',
        '', '', '',
        '', '', '',
        '', '', '',
        app.deductionTransactionId || '',
        app.submittedAt || new Date().toISOString(),
        app.updatedAt || new Date().toISOString()
      ]);

      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: 'Pengajuan cuti ' + app.id + ' berhasil dicatat ke Google Sheets!'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'UPDATE_LEAVE_APPLICATION') {
      const app = postData.application || postData;
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const leaveSheet = getOrCreateSheet('Leave_Applications', [
        'ApplicationID', 'EmployeeID', 'EmployeeName', 'DepartmentID', 'LeaveType', 
        'StartDate', 'EndDate', 'TotalDays', 'Reason', 'Status', 
        'DeptHeadDecision', 'DeptHeadDate', 'DeptHeadNotes', 
        'HCDecision', 'HCDate', 'HCNotes', 
        'GMDecision', 'GMDate', 'GMNotes', 
        'DeductionTxnID', 'SubmittedAt', 'UpdatedAt'
      ]);

      const vals = leaveSheet.getDataRange().getValues();
      let rowIdx = -1;
      for (let i = 1; i < vals.length; i++) {
        if (vals[i][0] === app.id) {
          rowIdx = i + 1;
          break;
        }
      }

      if (rowIdx > 0) {
        leaveSheet.getRange(rowIdx, 10).setValue(app.status);
        leaveSheet.getRange(rowIdx, 20).setValue(app.deductionTransactionId || postData.deductionTxnId || '');
        leaveSheet.getRange(rowIdx, 22).setValue(new Date().toISOString());
      } else {
        leaveSheet.appendRow([
          app.id, app.employeeId, app.employeeName, app.departmentId, app.leaveType,
          app.startDate, app.endDate, app.totalDays, app.reason || '', app.status,
          '', '', '', '', '', '', '', '', '',
          app.deductionTransactionId || postData.deductionTxnId || '',
          app.submittedAt || new Date().toISOString(), new Date().toISOString()
        ]);
      }

      // If approved by GM and deduction exists, also record into Leave_Ledger
      if (app.status === 'APPROVED' && postData.deductionTxnId) {
        const ledgerSheet = getOrCreateSheet('Leave_Ledger', [
          'TransactionID', 'EmployeeID', 'LeaveType', 'TransactionType', 'Days', 
          'EffectiveDate', 'ExpiryDate', 'RelatedApplicationID', 'AuthorizedBy', 
          'AuthorizedByName', 'ReferenceDoc', 'Reason', 'CreatedAt'
        ]);
        ledgerSheet.appendRow([
          postData.deductionTxnId,
          app.employeeId,
          app.leaveType,
          'APPROVED_LEAVE_DEDUCTION',
          -app.totalDays,
          app.startDate,
          '',
          app.id,
          'GM',
          'General Manager GMCC',
          app.id,
          'Persetujuan Final GM Pengajuan ' + app.id + ' (' + app.leaveType + ')',
          new Date().toISOString()
        ]);
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: 'Status pengajuan cuti ' + app.id + ' (' + app.status + ') berhasil diupdate ke Google Sheets!'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'SYNC_ALL' || postData.data || postData.employees) {
      const payloadData = postData.data || postData;
      const summary = syncDatabaseAndFolders(payloadData);
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: 'Database Google Sheets & Folder Drive GMCC berhasil disinkronkan!',
        summary: summary,
        timestamp: new Date().toISOString()
      })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'Unknown action: ' + action
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Sinkronisasi Master Spreadsheet & Pembuatan Folder Google Drive
 */
/**
 * Sinkronisasi Master Spreadsheet & Pembuatan Folder Google Drive
 */
function syncDatabaseAndFolders(data) {
  if (!data) return { synced: 0 };
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  function getOrCreateSheet(sheetName, headers) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      if (headers && headers.length > 0) {
        sheet.appendRow(headers);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e2e8f0");
      }
    }
    return sheet;
  }

  // 1. Sinkronkan Lembar Kerja Employees
  if (data.employees && data.employees.length > 0) {
    const empHeaders = [
      'EmployeeID', 'EmployeeNumber', 'FullName', 'PreferredName', 'Email', 
      'CompanyEmail', 'Phone', 'DepartmentID', 'JobTitle', 'EmploymentType', 
      'EmploymentStatus', 'JoinDate', 'ProbationEndDate', 'SupervisorID', 
      'Location', 'WorkSchedule', 'DriveFolderID', 'CreatedAt', 'UpdatedAt'
    ];
    const empSheet = getOrCreateSheet('Employees', empHeaders);

    const values = empSheet.getDataRange().getValues();
    const existingIds = {};
    for (let i = 1; i < values.length; i++) {
      existingIds[values[i][0]] = i + 1; // row number
    }

    data.employees.forEach(emp => {
      const rowData = [
        emp.id,
        emp.employeeNumber || '',
        emp.fullName || '',
        emp.preferredName || '',
        emp.email || '',
        emp.companyEmail || '',
        emp.phone || '',
        emp.departmentId || '',
        emp.jobTitle || '',
        emp.employmentType || '',
        emp.employmentStatus || '',
        emp.joinDate || '',
        emp.probationEndDate || '',
        emp.directSupervisorId || '',
        emp.location || '',
        emp.workSchedule || '',
        emp.driveFolderId || '',
        emp.createdAt || new Date().toISOString(),
        new Date().toISOString()
      ];

      if (existingIds[emp.id]) {
        empSheet.getRange(existingIds[emp.id], 1, 1, rowData.length).setValues([rowData]);
      } else {
        empSheet.appendRow(rowData);
      }

      // Buat folder Google Drive jika status aktif
      if (emp.employmentStatus === 'ACTIVE' && typeof PersonalFileService !== 'undefined') {
        try {
          PersonalFileService.createEmployeeFolders(emp.id, emp.fullName);
        } catch (e) {
          Logger.log('Folder creation notice: ' + e);
        }
      }
    });
  }

  // 2. Sinkronkan Lembar Kerja Contracts
  if (data.contracts && data.contracts.length > 0) {
    const ctrHeaders = [
      'ContractID', 'ContractNumber', 'EmployeeID', 'EmployeeName', 'DepartmentID', 
      'Position', 'EmploymentType', 'StartDate', 'EndDate', 'Status', 
      'Version', 'DriveFileID', 'ResponsibleHC', 'Notes', 'CreatedAt', 'UpdatedAt'
    ];
    const ctrSheet = getOrCreateSheet('Employment_Contracts', ctrHeaders);
    const cVals = ctrSheet.getDataRange().getValues();
    const existingCIds = {};
    for (let i = 1; i < cVals.length; i++) {
      existingCIds[cVals[i][0]] = i + 1;
    }

    data.contracts.forEach(c => {
      const cRow = [
        c.id, c.contractNumber, c.employeeId, c.employeeName, c.departmentId,
        c.position, c.employmentType, c.startDate, c.endDate || '', c.status,
        c.version, c.driveFileId || '', c.responsibleHC || '', c.notes || '',
        c.createdAt, new Date().toISOString()
      ];
      if (existingCIds[c.id]) {
        ctrSheet.getRange(existingCIds[c.id], 1, 1, cRow.length).setValues([cRow]);
      } else {
        ctrSheet.appendRow(cRow);
      }
    });
  }

  // 3. Sinkronkan Training Plans
  if (data.trainingPlans && data.trainingPlans.length > 0) {
    const tpHeaders = ['PlanID', 'Title', 'DepartmentID', 'DepartmentName', 'TrainingHours', 'Venue', 'TrainerName', 'Status', 'UpdatedAt'];
    const tpSheet = getOrCreateSheet('Training_Plans', tpHeaders);
    const tpVals = tpSheet.getDataRange().getValues();
    const existingTpIds = {};
    for (let i = 1; i < tpVals.length; i++) {
      existingTpIds[tpVals[i][0]] = i + 1;
    }
    data.trainingPlans.forEach(tp => {
      const tpRow = [tp.id, tp.title, tp.departmentId, tp.departmentName, tp.actualTrainingHours || tp.trainingHours || 4, tp.actualVenue || tp.venue || 'GMCC Training Hall', tp.trainerName || '', tp.status, new Date().toISOString()];
      if (existingTpIds[tp.id]) {
        tpSheet.getRange(existingTpIds[tp.id], 1, 1, tpRow.length).setValues([tpRow]);
      } else {
        tpSheet.appendRow(tpRow);
      }
    });
  }

  // 4. Sinkronkan Lembar Kerja Employee_Documents & Simpan Image/Dokumen ke Google Drive
  let documentsSynced = 0;
  if (data.documents && data.documents.length > 0) {
    const docHeaders = [
      'DocumentID', 'EmployeeID', 'EmployeeName', 'Category', 'Title', 'FileName', 
      'DriveFileID', 'DriveFileUrl', 'DrivePath', 'UploadedBy', 'UploadedAt'
    ];
    const docSheet = getOrCreateSheet('Employee_Documents', docHeaders);
    const dVals = docSheet.getDataRange().getValues();
    const existingDocIds = {};
    for (let i = 1; i < dVals.length; i++) {
      existingDocIds[dVals[i][0]] = i + 1;
    }

    data.documents.forEach(doc => {
      let fileId = doc.driveFileId || '';
      let fileUrl = '';
      let drivePath = doc.drivePath || '';

      // Jika ada fileData (gambar / dokumen) dan belum tersimpan di Drive
      if (doc.fileData && (!fileId || fileId.startsWith('DRV-FILE-')) && typeof PersonalFileService !== 'undefined' && PersonalFileService.uploadDocumentToDrive) {
        try {
          const res = PersonalFileService.uploadDocumentToDrive(
            doc.employeeId,
            doc.employeeName || 'Staff',
            doc.category,
            doc.fileName,
            doc.fileData,
            doc.mimeType
          );
          if (res && res.fileId) {
            fileId = res.fileId;
            fileUrl = res.fileUrl;
            drivePath = res.path;
          }
        } catch (e) {
          Logger.log('Auto-upload document to Drive error: ' + e);
        }
      }

      const docRow = [
        doc.id,
        doc.employeeId,
        doc.employeeName || '',
        doc.category,
        doc.title || doc.fileName,
        doc.fileName,
        fileId,
        fileUrl,
        drivePath,
        doc.uploadedByName || doc.uploadedBy || 'Admin',
        doc.uploadedAt || new Date().toISOString()
      ];

      if (existingDocIds[doc.id]) {
        docSheet.getRange(existingDocIds[doc.id], 1, 1, docRow.length).setValues([docRow]);
      } else {
        docSheet.appendRow(docRow);
      }
      documentsSynced++;
    });
  }

  // 5. Sinkronkan Training Attendance (Presensi & Bukti Foto)
  let attendanceSynced = 0;
  if (data.trainingAttendance && data.trainingAttendance.length > 0) {
    const attHeaders = [
      'AttendanceID', 'SessionID', 'SessionTitle', 'TrainingScope', 'EmployeeID', 
      'EmployeeName', 'DepartmentID', 'DepartmentName', 'AttendanceDate', 'CheckInTime', 
      'TrainingHours', 'Location', 'PhotoDriveFileID', 'PhotoDriveUrl', 'Status', 
      'Notes', 'VerifiedBy', 'VerifiedAt', 'CreatedAt'
    ];
    const attSheet = getOrCreateSheet('Training_Attendance', attHeaders);
    const aVals = attSheet.getDataRange().getValues();
    const existingAttIds = {};
    for (let i = 1; i < aVals.length; i++) {
      existingAttIds[aVals[i][0]] = i + 1;
    }

    data.trainingAttendance.forEach(att => {
      const attRow = [
        att.id,
        att.sessionId || '',
        att.sessionTitle || '',
        att.trainingScope || '',
        att.employeeId || '',
        att.employeeName || '',
        att.departmentId || '',
        att.departmentName || '',
        att.attendanceDate || '',
        att.checkInTime || '',
        att.trainingHours || 4,
        att.location || 'GMCC Training Hall',
        att.photoDriveFileId || '',
        att.photoProofUrl || '',
        att.status || 'VERIFIED',
        att.notes || '',
        att.verifiedBy || '',
        att.verifiedAt || '',
        new Date().toISOString()
      ];
      if (existingAttIds[att.id]) {
        attSheet.getRange(existingAttIds[att.id], 1, 1, attRow.length).setValues([attRow]);
      } else {
        attSheet.appendRow(attRow);
      }
      attendanceSynced++;
    });
  }

  // 6. Sinkronkan Leave Applications (Pengajuan Cuti)
  let leaveAppsSynced = 0;
  if (data.leaveApplications && data.leaveApplications.length > 0) {
    const leaveHeaders = [
      'ApplicationID', 'EmployeeID', 'EmployeeName', 'DepartmentID', 'LeaveType', 
      'StartDate', 'EndDate', 'TotalDays', 'Reason', 'Status', 
      'DeptHeadDecision', 'DeptHeadDate', 'DeptHeadNotes', 
      'HCDecision', 'HCDate', 'HCNotes', 
      'GMDecision', 'GMDate', 'GMNotes', 
      'DeductionTxnID', 'SubmittedAt', 'UpdatedAt'
    ];
    const leaveSheet = getOrCreateSheet('Leave_Applications', leaveHeaders);
    const lVals = leaveSheet.getDataRange().getValues();
    const existingLeaveIds = {};
    for (let i = 1; i < lVals.length; i++) {
      existingLeaveIds[lVals[i][0]] = i + 1;
    }

    data.leaveApplications.forEach(app => {
      const appRow = [
        app.id,
        app.employeeId,
        app.employeeName,
        app.departmentId,
        app.leaveType,
        app.startDate,
        app.endDate,
        app.totalDays,
        app.reason || '',
        app.status,
        '', '', '', '', '', '', '', '', '',
        app.deductionTransactionId || '',
        app.submittedAt || '',
        app.updatedAt || new Date().toISOString()
      ];
      if (existingLeaveIds[app.id]) {
        leaveSheet.getRange(existingLeaveIds[app.id], 1, 1, appRow.length).setValues([appRow]);
      } else {
        leaveSheet.appendRow(appRow);
      }
      leaveAppsSynced++;
    });
  }

  // 7. Sinkronkan Leave Ledger
  let ledgerSynced = 0;
  if (data.leaveLedger && data.leaveLedger.length > 0) {
    const ledgerHeaders = [
      'TransactionID', 'EmployeeID', 'LeaveType', 'TransactionType', 'Days', 
      'EffectiveDate', 'ExpiryDate', 'RelatedApplicationID', 'AuthorizedBy', 
      'AuthorizedByName', 'ReferenceDoc', 'Reason', 'CreatedAt'
    ];
    const ledgerSheet = getOrCreateSheet('Leave_Ledger', ledgerHeaders);
    const ledVals = ledgerSheet.getDataRange().getValues();
    const existingLedIds = {};
    for (let i = 1; i < ledVals.length; i++) {
      existingLedIds[ledVals[i][0]] = i + 1;
    }

    data.leaveLedger.forEach(entry => {
      const entryRow = [
        entry.id,
        entry.employeeId,
        entry.leaveType,
        entry.transactionType,
        entry.days,
        entry.effectiveDate,
        entry.expiryDate || '',
        entry.relatedApplicationId || '',
        entry.authorizedBy || '',
        entry.authorizedByName || '',
        entry.referenceDocNumber || '',
        entry.reason || '',
        entry.createdAt || new Date().toISOString()
      ];
      if (existingLedIds[entry.id]) {
        ledgerSheet.getRange(existingLedIds[entry.id], 1, 1, entryRow.length).setValues([entryRow]);
      } else {
        ledgerSheet.appendRow(entryRow);
      }
      ledgerSynced++;
    });
  }

  return {
    employeesSynced: data.employees ? data.employees.length : 0,
    contractsSynced: data.contracts ? data.contracts.length : 0,
    trainingPlansSynced: data.trainingPlans ? data.trainingPlans.length : 0,
    trainingAttendanceSynced: attendanceSynced,
    leaveApplicationsSynced: leaveAppsSynced,
    leaveLedgerSynced: ledgerSynced,
    documentsSynced: documentsSynced
  };
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
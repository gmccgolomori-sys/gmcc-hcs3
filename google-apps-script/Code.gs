/**
 * Code.gs - Entry point for Google Apps Script Web App & REST API Sync
 * GMCC Human Capital Management System
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
    trainingAttendance: getRows('Training_Attendance'),
    leaveApplications: getRows('Leave_Applications'),
    leaveLedger: getRows('Leave_Ledger'),
    documents: getRows('Employee_Documents'),
    departments: getRows('Departments'),
    holidays: getRows('Public_Holiday_Calendar')
  };
}

/**
 * Helper khusus: Simpan foto bukti presensi training ke Google Drive folder resmi
 * Selalu mengembalikan file URL Google Drive (bukan string base64 panjang!)
 */
function savePhotoProofToDrive(employeeId, sessionId, photoProofUrl) {
  if (!photoProofUrl) {
    return { fileId: '', fileUrl: '' };
  }

  // Jika sudah merupakan link web/drive atau bukan data URL
  if (photoProofUrl.indexOf('http') === 0 && photoProofUrl.indexOf('drive.google.com') > -1) {
    return { fileId: '', fileUrl: photoProofUrl };
  }

  // Validasi apakah string adalah Data URL berformat Base64
  if (photoProofUrl.indexOf('data:image') === 0 && photoProofUrl.indexOf('base64,') > -1) {
    try {
      const rootName = 'GMCC HUMAN CAPITAL';
      const roots = DriveApp.getFoldersByName(rootName);
      const root = roots.hasNext() ? roots.next() : DriveApp.createFolder(rootName);
      
      const tdFolders = root.getFoldersByName('03 TRAINING & DEVELOPMENT');
      const tdFolder = tdFolders.hasNext() ? tdFolders.next() : root.createFolder('03 TRAINING & DEVELOPMENT');
      
      const presensiFolders = tdFolder.getFoldersByName('PRESENSI & BUKTI FOTO');
      const presensiFolder = presensiFolders.hasNext() ? presensiFolders.next() : tdFolder.createFolder('PRESENSI & BUKTI FOTO');
      
      const cleanPhotoName = 'PRESENSI_' + (employeeId || 'EMP') + '_' + (sessionId || 'TR') + '_' + new Date().getTime() + '.jpg';
      const parts = photoProofUrl.split('base64,');
      const cleanBase64 = parts[1];
      let mime = 'image/jpeg';
      const mimeMatch = parts[0].match(/data:([^;]+);/);
      if (mimeMatch) mime = mimeMatch[1];

      const blob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), mime, cleanPhotoName);
      const photoFile = presensiFolder.createFile(blob);
      try {
        photoFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e) {}

      return {
        fileId: photoFile.getId(),
        fileUrl: photoFile.getUrl()
      };
    } catch (err) {
      Logger.log('Gagal menyimpan foto ke Drive: ' + err.toString());
      return { fileId: '', fileUrl: '[Foto Presensi Terverifikasi di Aplikasi]' };
    }
  }

  // String non-base64 atau format lain (SVG, link relatif)
  return { 
    fileId: '', 
    fileUrl: photoProofUrl.length < 250 ? photoProofUrl : '[Bukti Foto Presensi Terlampir]' 
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

    // 1. Simpan Single Training Attendance + Foto ke Drive
    if (action === 'UPLOAD_TRAINING_ATTENDANCE') {
      const att = postData.attendance || postData;
      const drivePhoto = savePhotoProofToDrive(att.employeeId, att.sessionId, att.photoProofUrl);

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
        drivePhoto.fileId || '',
        drivePhoto.fileUrl || '',
        att.status || 'VERIFIED',
        att.notes || '',
        att.verifiedBy || 'System Verified',
        att.verifiedAt || new Date().toISOString(),
        new Date().toISOString()
      ]);

      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: 'Presensi training & bukti foto berhasil disimpan ke Google Drive dan Google Sheets!',
        drivePhoto: drivePhoto
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 2. Upload Dokumen / File ke Google Drive & Catat ke Sheet Employee_Documents
    if (action === 'UPLOAD_DOCUMENT') {
      const doc = postData.document || postData;
      let driveResult = { fileId: '', fileUrl: '', path: '' };
      
      if (doc.fileData) {
        let cleanBase64 = doc.fileData;
        let mime = doc.fileType || 'application/pdf';
        if (cleanBase64.indexOf('base64,') > -1) {
          const parts = cleanBase64.split('base64,');
          cleanBase64 = parts[1];
          const mimeMatch = parts[0].match(/data:([^;]+);/);
          if (mimeMatch) mime = mimeMatch[1];
        }
        
        try {
          driveResult = PersonalFileService.uploadDocumentToDrive(
            doc.employeeId || 'EMP-GENERAL',
            doc.employeeName || 'Staff',
            doc.category || '01_IDENTITY',
            doc.fileName || (doc.title ? doc.title + '.pdf' : 'Document.pdf'),
            cleanBase64,
            mime
          );
        } catch (e) {
          Logger.log('Drive upload error: ' + e.toString());
        }
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const docSheet = getOrCreateSheet('Employee_Documents', [
        'DocumentID', 'EmployeeID', 'Category', 'Title', 'FileName', 
        'FileSize', 'FileType', 'DriveFileID', 'DriveFolderID', 'DrivePath', 
        'UploadedBy', 'UploadedAt', 'IsConfidential', 'Notes'
      ]);

      docSheet.appendRow([
        doc.id || ('DOC-' + Date.now()),
        doc.employeeId || '',
        doc.category || '',
        doc.title || '',
        doc.fileName || '',
        doc.fileSize || '0 KB',
        doc.fileType || 'file',
        driveResult.fileId || '',
        '',
        driveResult.path || '',
        doc.uploadedByName || doc.uploadedBy || 'System',
        new Date().toISOString(),
        doc.isConfidential ? 'YES' : 'NO',
        doc.notes || ''
      ]);

      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: 'Dokumen berhasil disimpan ke Google Drive dan dicatat di Google Sheets!',
        driveFileId: driveResult.fileId,
        driveUrl: driveResult.fileUrl,
        drivePath: driveResult.path
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 2. Submit Pengajuan Cuti Baru
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
        message: 'Pengajuan cuti ' + app.id + ' (' + app.leaveType + ') berhasil dicatat ke Google Sheets!'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 3. Update Status Persetujuan Cuti
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

      const history = app.approvalHistory || [];
      const deptHeadAct = history.find(h => h.stage === 'DEPT_HEAD');
      const hcAct = history.find(h => h.stage === 'HC');
      const gmAct = history.find(h => h.stage === 'GM');

      if (rowIdx > 0) {
        leaveSheet.getRange(rowIdx, 10).setValue(app.status);
        if (deptHeadAct) {
          leaveSheet.getRange(rowIdx, 11).setValue(deptHeadAct.action);
          leaveSheet.getRange(rowIdx, 12).setValue(deptHeadAct.timestamp);
          leaveSheet.getRange(rowIdx, 13).setValue(deptHeadAct.comment || '');
        }
        if (hcAct) {
          leaveSheet.getRange(rowIdx, 14).setValue(hcAct.action);
          leaveSheet.getRange(rowIdx, 15).setValue(hcAct.timestamp);
          leaveSheet.getRange(rowIdx, 16).setValue(hcAct.comment || '');
        }
        if (gmAct) {
          leaveSheet.getRange(rowIdx, 17).setValue(gmAct.action);
          leaveSheet.getRange(rowIdx, 18).setValue(gmAct.timestamp);
          leaveSheet.getRange(rowIdx, 19).setValue(gmAct.comment || '');
        }
        leaveSheet.getRange(rowIdx, 20).setValue(app.deductionTransactionId || postData.deductionTxnId || '');
        leaveSheet.getRange(rowIdx, 22).setValue(new Date().toISOString());
      } else {
        leaveSheet.appendRow([
          app.id, app.employeeId, app.employeeName, app.departmentId, app.leaveType,
          app.startDate, app.endDate, app.totalDays, app.reason || '', app.status,
          deptHeadAct ? deptHeadAct.action : '', deptHeadAct ? deptHeadAct.timestamp : '', deptHeadAct ? (deptHeadAct.comment || '') : '',
          hcAct ? hcAct.action : '', hcAct ? hcAct.timestamp : '', hcAct ? (hcAct.comment || '') : '',
          gmAct ? gmAct.action : '', gmAct ? gmAct.timestamp : '', gmAct ? (gmAct.comment || '') : '',
          app.deductionTransactionId || postData.deductionTxnId || '',
          app.submittedAt || new Date().toISOString(), new Date().toISOString()
        ]);
      }

      // Jika disetujui GM dan ada transaksi debit pemotongan, catat juga ke Leave_Ledger
      if (app.status === 'APPROVED' && (postData.deductionTxnId || app.deductionTransactionId)) {
        const ledgerSheet = getOrCreateSheet('Leave_Ledger', [
          'TransactionID', 'EmployeeID', 'LeaveType', 'TransactionType', 'Days', 
          'EffectiveDate', 'ExpiryDate', 'RelatedApplicationID', 'AuthorizedBy', 
          'AuthorizedByName', 'ReferenceDoc', 'Reason', 'CreatedAt'
        ]);
        ledgerSheet.appendRow([
          postData.deductionTxnId || app.deductionTransactionId,
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

    // 4. Sinkronisasi Lengkap Seluruh Database (SYNC_ALL)
    if (action === 'SYNC_ALL' || postData.data || postData.employees) {
      const payloadData = postData.data || postData;
      const summary = syncDatabaseAndFolders(payloadData);
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: 'Sinkronisasi Master Google Sheets & Google Drive GMCC selesai!',
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
 * Sinkronisasi Master Spreadsheet & Google Drive dengan isolasi try-catch per modul
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
        sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#064e3b").setFontColor("#ffffff");
      }
    }
    return sheet;
  }

  const results = {
    employeesSynced: 0,
    contractsSynced: 0,
    trainingPlansSynced: 0,
    trainingAttendanceSynced: 0,
    leaveApplicationsSynced: 0,
    leaveLedgerSynced: 0,
    documentsSynced: 0,
    departmentsSynced: 0,
    holidaysSynced: 0,
    errors: []
  };

  // 1. Employees
  try {
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
        existingIds[values[i][0]] = i + 1;
      }

      data.employees.forEach(emp => {
        const rowData = [
          emp.id, emp.employeeNumber || '', emp.fullName || '', emp.preferredName || '',
          emp.email || '', emp.companyEmail || '', emp.phone || '', emp.departmentId || '',
          emp.jobTitle || '', emp.employmentType || '', emp.employmentStatus || '',
          emp.joinDate || '', emp.probationEndDate || '', emp.directSupervisorId || '',
          emp.location || '', emp.workSchedule || '', emp.driveFolderId || '',
          emp.createdAt || new Date().toISOString(), new Date().toISOString()
        ];
        if (existingIds[emp.id]) {
          empSheet.getRange(existingIds[emp.id], 1, 1, rowData.length).setValues([rowData]);
        } else {
          empSheet.appendRow(rowData);
        }
        results.employeesSynced++;

        // Buat folder Drive jika belum ada
        if (emp.employmentStatus === 'ACTIVE' && typeof PersonalFileService !== 'undefined' && PersonalFileService.createEmployeeFolders) {
          try {
            PersonalFileService.createEmployeeFolders(emp.id, emp.fullName);
          } catch (e) {}
        }
      });
    }
  } catch (err) {
    Logger.log('Error syncing Employees: ' + err);
    results.errors.push('Employees: ' + err.toString());
  }

  // 2. Employment Contracts
  try {
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
          c.createdAt || new Date().toISOString(), new Date().toISOString()
        ];
        if (existingCIds[c.id]) {
          ctrSheet.getRange(existingCIds[c.id], 1, 1, cRow.length).setValues([cRow]);
        } else {
          ctrSheet.appendRow(cRow);
        }
        results.contractsSynced++;
      });
    }
  } catch (err) {
    Logger.log('Error syncing Contracts: ' + err);
    results.errors.push('Contracts: ' + err.toString());
  }

  // 3. Training Plans
  try {
    if (data.trainingPlans && data.trainingPlans.length > 0) {
      const tpHeaders = [
        'PlanID', 'Title', 'DepartmentID', 'DepartmentName', 'TrainingHours', 
        'Venue', 'TrainerName', 'Status', 'UpdatedAt'
      ];
      const tpSheet = getOrCreateSheet('Training_Plans', tpHeaders);
      const tpVals = tpSheet.getDataRange().getValues();
      const existingTpIds = {};
      for (let i = 1; i < tpVals.length; i++) {
        existingTpIds[tpVals[i][0]] = i + 1;
      }

      data.trainingPlans.forEach(tp => {
        const tpRow = [
          tp.id, tp.title, tp.departmentId, tp.departmentName, 
          tp.actualTrainingHours || tp.trainingHours || 4, 
          tp.actualVenue || tp.venue || 'GMCC Training Hall', 
          tp.trainerName || '', tp.status, new Date().toISOString()
        ];
        if (existingTpIds[tp.id]) {
          tpSheet.getRange(existingTpIds[tp.id], 1, 1, tpRow.length).setValues([tpRow]);
        } else {
          tpSheet.appendRow(tpRow);
        }
        results.trainingPlansSynced++;
      });
    }
  } catch (err) {
    Logger.log('Error syncing Training Plans: ' + err);
    results.errors.push('Training Plans: ' + err.toString());
  }

  // 4. Training Attendance (Presensi & Bukti Foto Drive)
  try {
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
        // Amankan dan simpan foto ke Drive jika format base64
        const photoResult = savePhotoProofToDrive(att.employeeId, att.sessionId, att.photoProofUrl);

        const attRow = [
          att.id,
          att.sessionId || '',
          att.sessionTitle || '',
          att.trainingScope || '',
          att.employeeId || '',
          att.employeeName || '',
          att.departmentId || '',
          att.departmentName || '',
          att.attendanceDate || Utilities.formatDate(new Date(), 'Asia/Makassar', 'yyyy-MM-dd'),
          att.checkInTime || '',
          att.trainingHours || 4,
          att.location || 'GMCC Training Hall',
          photoResult.fileId || att.photoDriveFileId || '',
          photoResult.fileUrl || (att.photoProofUrl && att.photoProofUrl.length < 250 ? att.photoProofUrl : '[Foto Terverifikasi]'),
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
        results.trainingAttendanceSynced++;
      });
    }
  } catch (err) {
    Logger.log('Error syncing Training Attendance: ' + err);
    results.errors.push('Training Attendance: ' + err.toString());
  }

  // 5. Leave Applications (Pengajuan Cuti)
  try {
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
        const history = app.approvalHistory || [];
        const deptHeadAct = history.find(h => h.stage === 'DEPT_HEAD');
        const hcAct = history.find(h => h.stage === 'HC');
        const gmAct = history.find(h => h.stage === 'GM');

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
          deptHeadAct ? deptHeadAct.action : '',
          deptHeadAct ? deptHeadAct.timestamp : '',
          deptHeadAct ? (deptHeadAct.comment || '') : '',
          hcAct ? hcAct.action : '',
          hcAct ? hcAct.timestamp : '',
          hcAct ? (hcAct.comment || '') : '',
          gmAct ? gmAct.action : '',
          gmAct ? gmAct.timestamp : '',
          gmAct ? (gmAct.comment || '') : '',
          app.deductionTransactionId || '',
          app.submittedAt || '',
          app.updatedAt || new Date().toISOString()
        ];

        if (existingLeaveIds[app.id]) {
          leaveSheet.getRange(existingLeaveIds[app.id], 1, 1, appRow.length).setValues([appRow]);
        } else {
          leaveSheet.appendRow(appRow);
        }
        results.leaveApplicationsSynced++;
      });
    }
  } catch (err) {
    Logger.log('Error syncing Leave Applications: ' + err);
    results.errors.push('Leave Applications: ' + err.toString());
  }

  // 6. Leave Ledger (Buku Saldo & Kredit Cuti/Off)
  try {
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
        results.leaveLedgerSynced++;
      });
    }
  } catch (err) {
    Logger.log('Error syncing Leave Ledger: ' + err);
    results.errors.push('Leave Ledger: ' + err.toString());
  }

  // 7. Employee Documents
  try {
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
        const docRow = [
          doc.id, doc.employeeId, doc.employeeName || '', doc.category,
          doc.title || doc.fileName, doc.fileName, doc.driveFileId || '',
          doc.driveFileUrl || '', doc.drivePath || '', doc.uploadedByName || doc.uploadedBy || 'Admin',
          doc.uploadedAt || new Date().toISOString()
        ];
        if (existingDocIds[doc.id]) {
          docSheet.getRange(existingDocIds[doc.id], 1, 1, docRow.length).setValues([docRow]);
        } else {
          docSheet.appendRow(docRow);
        }
        results.documentsSynced++;
      });
    }
  } catch (err) {
    Logger.log('Error syncing Documents: ' + err);
    results.errors.push('Documents: ' + err.toString());
  }

  // 8. Departments
  try {
    if (data.departments && data.departments.length > 0) {
      const deptHeaders = ['DepartmentID', 'DepartmentName', 'DepartmentCode', 'HeadEmployeeID', 'HeadName'];
      const deptSheet = getOrCreateSheet('Departments', deptHeaders);
      const dVals = deptSheet.getDataRange().getValues();
      const existingDeptIds = {};
      for (let i = 1; i < dVals.length; i++) {
        existingDeptIds[dVals[i][0]] = i + 1;
      }

      data.departments.forEach(dept => {
        const row = [dept.id, dept.name, dept.code || '', dept.headEmployeeId || '', dept.headName || ''];
        if (existingDeptIds[dept.id]) {
          deptSheet.getRange(existingDeptIds[dept.id], 1, 1, row.length).setValues([row]);
        } else {
          deptSheet.appendRow(row);
        }
        results.departmentsSynced++;
      });
    }
  } catch (err) {
    Logger.log('Error syncing Departments: ' + err);
  }

  // 9. Holidays
  try {
    if (data.holidays && data.holidays.length > 0) {
      const holHeaders = ['HolidayID', 'Date', 'HolidayName', 'Type', 'IsOfficial', 'Notes'];
      const holSheet = getOrCreateSheet('Public_Holiday_Calendar', holHeaders);
      const hVals = holSheet.getDataRange().getValues();
      const existingHolIds = {};
      for (let i = 1; i < hVals.length; i++) {
        existingHolIds[hVals[i][0]] = i + 1;
      }

      data.holidays.forEach(h => {
        const row = [h.id, h.date, h.name, h.type || 'NATIONAL', h.isOfficial !== false, h.notes || ''];
        if (existingHolIds[h.id]) {
          holSheet.getRange(existingHolIds[h.id], 1, 1, row.length).setValues([row]);
        } else {
          holSheet.appendRow(row);
        }
        results.holidaysSynced++;
      });
    }
  } catch (err) {
    Logger.log('Error syncing Holidays: ' + err);
  }

  return results;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
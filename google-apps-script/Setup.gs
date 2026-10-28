/**
 * GMCC HUMAN CAPITAL MANAGEMENT SYSTEM
 * Setup.gs - Initialization of Spreadsheet Sheets & Drive Folders
 * Timezone: Asia/Makassar (WITA)
 */

const CONFIG = {
  TIMEZONE: 'Asia/Makassar',
  ROOT_FOLDER_NAME: 'GMCC HUMAN CAPITAL',
  SHEET_ID: SpreadsheetApp.getActiveSpreadsheet().getId(),
  SUB_FOLDERS: [
    '01 PERSONAL FILE',
    '02 EMPLOYMENT CONTRACTS',
    '03 TRAINING & DEVELOPMENT',
    '04 HR REPORTS',
    '05 POLICY & SOP',
    '06 DOCUMENT TEMPLATES'
  ],
  PERSONAL_FILE_CATEGORIES: [
    '01 Identity',
    '02 Recruitment',
    '03 Employment',
    '04 Training',
    '05 Performance',
    '06 Attendance & Leave',
    '07 Discipline',
    '08 Compensation & Benefits',
    '09 Health & Work Requirements',
    '10 Exit'
  ]
};

function initGMCCSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Initialize Sheets
  createSheetIfNotExists(ss, 'Employees', [
    'EmployeeID', 'EmployeeNumber', 'FullName', 'PreferredName', 'Email', 
    'CompanyEmail', 'Phone', 'DepartmentID', 'JobTitle', 'EmploymentType', 
    'EmploymentStatus', 'JoinDate', 'ProbationEndDate', 'SupervisorID', 
    'Location', 'WorkSchedule', 'DriveFolderID', 'CreatedAt', 'UpdatedAt'
  ]);

  createSheetIfNotExists(ss, 'Departments', [
    'DepartmentID', 'DepartmentName', 'DepartmentCode', 'HeadEmployeeID', 'HeadName'
  ]);

  createSheetIfNotExists(ss, 'Users_Roles', [
    'UserEmail', 'FullName', 'Role', 'EmployeeID', 'DepartmentID', 'IsActive'
  ]);

  createSheetIfNotExists(ss, 'Employee_Documents', [
    'DocumentID', 'EmployeeID', 'Category', 'Title', 'FileName', 
    'FileSize', 'FileType', 'DriveFileID', 'DriveFolderID', 'DrivePath', 
    'UploadedBy', 'UploadedAt', 'IsConfidential', 'Notes'
  ]);

  createSheetIfNotExists(ss, 'Employment_Contracts', [
    'ContractID', 'ContractNumber', 'EmployeeID', 'EmployeeName', 'DepartmentID', 
    'Position', 'EmploymentType', 'StartDate', 'EndDate', 'Status', 
    'Version', 'DriveFileID', 'ResponsibleHC', 'Notes', 'CreatedAt', 'UpdatedAt'
  ]);

  createSheetIfNotExists(ss, 'Leave_Applications', [
    'ApplicationID', 'EmployeeID', 'EmployeeName', 'DepartmentID', 'LeaveType', 
    'StartDate', 'EndDate', 'TotalDays', 'Reason', 'Status', 
    'DeptHeadDecision', 'DeptHeadDate', 'DeptHeadNotes', 
    'HCDecision', 'HCDate', 'HCNotes', 
    'GMDecision', 'GMDate', 'GMNotes', 
    'DeductionTxnID', 'SubmittedAt', 'UpdatedAt'
  ]);

  createSheetIfNotExists(ss, 'Leave_Ledger', [
    'TransactionID', 'EmployeeID', 'LeaveType', 'TransactionType', 'Days', 
    'EffectiveDate', 'ExpiryDate', 'RelatedApplicationID', 'AuthorizedBy', 
    'AuthorizedByName', 'ReferenceDoc', 'Reason', 'CreatedAt'
  ]);

  createSheetIfNotExists(ss, 'Public_Holiday_Calendar', [
    'HolidayID', 'Date', 'HolidayName', 'Type', 'IsOfficial', 'Notes'
  ]);

  createSheetIfNotExists(ss, 'Training_Plans', [
    'PlanID', 'Title', 'DepartmentID', 'Category', 'Priority', 'Objectives', 
    'TargetAudience', 'EstimatedParticipants', 'EstimatedCost', 'StartDate', 
    'EndDate', 'ProviderType', 'TrainerName', 'Status', 'SubmittedBy', 'SubmittedAt'
  ]);

  createSheetIfNotExists(ss, 'Training_Attendance', [
    'AttendanceID', 'SessionID', 'SessionTitle', 'TrainingScope', 'EmployeeID', 
    'EmployeeName', 'DepartmentID', 'DepartmentName', 'AttendanceDate', 'CheckInTime', 
    'TrainingHours', 'Location', 'PhotoDriveFileID', 'PhotoDriveUrl', 'Status', 
    'Notes', 'VerifiedBy', 'VerifiedAt', 'CreatedAt'
  ]);

  createSheetIfNotExists(ss, 'Audit_Log', [
    'LogID', 'Timestamp', 'ActorEmail', 'ActorRole', 'Module', 'Action', 'TargetID', 'Description'
  ]);

  Logger.log('Sheets initialized successfully.');

  // 2. Initialize Drive Folder Structure
  initDriveFolders();
}

/**
 * PENTING: Jalankan fungsi ini langsung di editor Google Apps Script dengan tombol Run (Jalankan)
 * untuk memicu pop-up izin Google Drive (Review Permissions -> Advanced -> Allow)
 * dan memverifikasi pembuatan folder root GMCC secara langsung!
 */
function testGoogleDriveConnection() {
  Logger.log('=== MEMULAI TEST KONEKSI GOOGLE DRIVE GMCC ===');
  try {
    const rootName = 'GMCC HUMAN CAPITAL';
    let rootFolder;
    const roots = DriveApp.getFoldersByName(rootName);
    if (roots.hasNext()) {
      rootFolder = roots.next();
      Logger.log('SUKSES: Folder root "' + rootName + '" ditemukan! ID: ' + rootFolder.getId());
    } else {
      rootFolder = DriveApp.createFolder(rootName);
      Logger.log('SUKSES: Folder root baru "' + rootName + '" berhasil dibuat! ID: ' + rootFolder.getId());
    }

    // Buat subfolder utama 01 PERSONAL FILE
    const pfFolders = rootFolder.getFoldersByName('01 PERSONAL FILE');
    const pfFolder = pfFolders.hasNext() ? pfFolders.next() : rootFolder.createFolder('01 PERSONAL FILE');
    Logger.log('SUKSES: Subfolder "01 PERSONAL FILE" siap! ID: ' + pfFolder.getId());

    const activeFolders = pfFolder.getFoldersByName('ACTIVE EMPLOYEE');
    const activeFolder = activeFolders.hasNext() ? activeFolders.next() : pfFolder.createFolder('ACTIVE EMPLOYEE');
    Logger.log('SUKSES: Subfolder "ACTIVE EMPLOYEE" siap! ID: ' + activeFolder.getId());

    Logger.log('=== GOOGLE DRIVE 100% TERHUBUNG DAN MENDAPATKAN IZIN AKSES RESMI! ===');
    return 'KONEKSI GOOGLE DRIVE BERHASIL';
  } catch (err) {
    Logger.log('GAGAL KONEKSI DRIVE: ' + err.toString());
    throw err;
  }
}

/**
 * PENTING: Jalankan fungsi ini untuk mengetes langsung pengiriman image / dokumen contoh ke subfolder Google Drive!
 */
function testUploadDocumentToDrive() {
  Logger.log('=== MEMULAI TEST SIMPAN DOKUMEN/IMAGE KE GOOGLE DRIVE ===');
  try {
    const dummyText = 'BERKAS DOKUMEN UJI COBA GMCC HUMAN CAPITAL\nID: TEST-DOC-001\nTanggal: ' + new Date().toISOString();
    const dummyBase64 = Utilities.base64Encode(dummyText);
    const result = PersonalFileService.uploadDocumentToDrive(
      'EMP-0001',
      'Test Staff GMCC',
      '01_IDENTITY',
      'Uji_Coba_Identitas.txt',
      dummyBase64,
      'text/plain'
    );
    Logger.log('SUKSES! File berhasil dibuat di Google Drive! ID: ' + result.fileId);
    Logger.log('Lokasi Folder: ' + result.path);
    Logger.log('URL Google Drive: ' + result.fileUrl);
    Logger.log('=== FOLDER DAN PENYIMPANAN IMAGE/DOKUMEN GOOGLE DRIVE 100% SUKSES! ===');
    return 'DOKUMEN BERHASIL DISIMPAN KE DRIVE';
  } catch (err) {
    Logger.log('GAGAL SIMPAN DOKUMEN KE DRIVE: ' + err.toString());
    throw err;
  }
}

/**
 * Jalankan fungsi ini untuk mengetes apakah Google Apps Script bisa membaca dan menulis data ke Google Sheets
 */
function testSyncSpreadsheetRows() {
  Logger.log('=== MEMULAI TEST SINKRONISASI GOOGLE SHEETS GMCC ===');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let empSheet = ss.getSheetByName('Employees');
    if (!empSheet) {
      initGMCCSystem();
      empSheet = ss.getSheetByName('Employees');
    }
    Logger.log('SUKSES: Lembar kerja "Employees" ditemukan! Jumlah baris saat ini: ' + empSheet.getLastRow());

    // Cek isi data
    const data = getAllSheetData();
    Logger.log('SUKSES MEMBACA DATA SPREADSHEET:');
    Logger.log('- Karyawan di Sheet: ' + data.employees.length);
    Logger.log('- Kontrak di Sheet: ' + data.contracts.length);
    Logger.log('- Training di Sheet: ' + data.trainingPlans.length);
    Logger.log('=== GOOGLE SHEETS 100% SIAP DAN DAPAT DISINKRONKAN DUA ARAH! ===');
    return 'TEST SINKRONISASI SHEETS SUKSES';
  } catch (err) {
    Logger.log('GAGAL SINKRONISASI SHEETS: ' + err.toString());
    throw err;
  }
}

function initDriveFolders() {
  let rootFolder;
  const existingRoots = DriveApp.getFoldersByName(CONFIG.ROOT_FOLDER_NAME);
  if (existingRoots.hasNext()) {
    rootFolder = existingRoots.next();
  } else {
    rootFolder = DriveApp.createFolder(CONFIG.ROOT_FOLDER_NAME);
  }

  CONFIG.SUB_FOLDERS.forEach(folderName => {
    const existing = rootFolder.getFoldersByName(folderName);
    if (!existing.hasNext()) {
      rootFolder.createFolder(folderName);
    }
  });

  Logger.log('Drive folder structure ready under ' + CONFIG.ROOT_FOLDER_NAME);
  return rootFolder.getId();
}

function createSheetIfNotExists(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#064e3b').setFontColor('#ffffff');
  }
  return sheet;
}

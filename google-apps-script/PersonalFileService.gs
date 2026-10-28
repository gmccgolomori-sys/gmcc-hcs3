/**
 * PersonalFileService.gs - Employee Personal File Subfolders in Google Drive
 */

const PersonalFileService = {
  createEmployeeFolders: function(employeeId, employeeName) {
    const rootName = 'GMCC HUMAN CAPITAL';
    const roots = DriveApp.getFoldersByName(rootName);
    const root = roots.hasNext() ? roots.next() : DriveApp.createFolder(rootName);

    const pfFolders = root.getFoldersByName('01 PERSONAL FILE');
    const pfFolder = pfFolders.hasNext() ? pfFolders.next() : root.createFolder('01 PERSONAL FILE');
    
    const activeFolders = pfFolder.getFoldersByName('ACTIVE EMPLOYEE');
    const activeFolder = activeFolders.hasNext() ? activeFolders.next() : pfFolder.createFolder('ACTIVE EMPLOYEE');

    const empFolderName = employeeId + ' - ' + employeeName;
    let empFolder;
    const existingEmp = activeFolder.getFoldersByName(empFolderName);
    if (existingEmp.hasNext()) {
      empFolder = existingEmp.next();
    } else {
      empFolder = activeFolder.createFolder(empFolderName);
    }

    // Create 10 standardized subfolders
    const categories = [
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
    ];

    const categoryMap = {};
    categories.forEach(cat => {
      const existingCat = empFolder.getFoldersByName(cat);
      const sub = existingCat.hasNext() ? existingCat.next() : empFolder.createFolder(cat);
      categoryMap[cat] = sub.getId();
    });

    return {
      employeeFolderId: empFolder.getId(),
      subfolders: categoryMap
    };
  },

  /**
   * Simpan file/gambar (Base64) langsung ke dalam subfolder Google Drive yang sesuai
   */
  uploadDocumentToDrive: function(employeeId, employeeName, categoryCode, fileName, fileBase64, mimeType) {
    if (!fileBase64) throw new Error('File data tidak boleh kosong.');

    const rootName = 'GMCC HUMAN CAPITAL';
    const roots = DriveApp.getFoldersByName(rootName);
    const root = roots.hasNext() ? roots.next() : DriveApp.createFolder(rootName);

    const pfFolders = root.getFoldersByName('01 PERSONAL FILE');
    const pfFolder = pfFolders.hasNext() ? pfFolders.next() : root.createFolder('01 PERSONAL FILE');
    
    const activeFolders = pfFolder.getFoldersByName('ACTIVE EMPLOYEE');
    const activeFolder = activeFolders.hasNext() ? activeFolders.next() : pfFolder.createFolder('ACTIVE EMPLOYEE');

    const empFolderName = employeeId + ' - ' + employeeName;
    let empFolder;
    const existingEmp = activeFolder.getFoldersByName(empFolderName);
    if (existingEmp.hasNext()) {
      empFolder = existingEmp.next();
    } else {
      empFolder = activeFolder.createFolder(empFolderName);
    }

    // Map kategori
    const categoryFolders = {
      '01_IDENTITY': '01 Identity',
      '02_RECRUITMENT': '02 Recruitment',
      '03_EMPLOYMENT': '03 Employment',
      '04_TRAINING': '04 Training',
      '05_PERFORMANCE': '05 Performance',
      '06_ATTENDANCE_LEAVE': '06 Attendance & Leave',
      '07_DISCIPLINE': '07 Discipline',
      '08_COMPENSATION_BENEFITS': '08 Compensation & Benefits',
      '09_HEALTH_WORK_REQUIREMENTS': '09 Health & Work Requirements',
      '10_EXIT': '10 Exit'
    };

    const targetSubName = categoryFolders[categoryCode] || categoryCode || '01 Identity';
    let targetFolder;
    const existingSubs = empFolder.getFoldersByName(targetSubName);
    if (existingSubs.hasNext()) {
      targetFolder = existingSubs.next();
    } else {
      targetFolder = empFolder.createFolder(targetSubName);
    }

    // Ekstrak base64 & MIME type
    let cleanBase64 = fileBase64;
    let detectedMime = mimeType || 'application/octet-stream';
    if (fileBase64.indexOf(',') > -1) {
      const parts = fileBase64.split(',');
      const meta = parts[0];
      cleanBase64 = parts[1];
      const mimeMatch = meta.match(/data:([^;]+);/);
      if (mimeMatch) detectedMime = mimeMatch[1];
    }

    const decodedBlob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), detectedMime, fileName);
    const driveFile = targetFolder.createFile(decodedBlob);
    
    try {
      driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {}

    return {
      success: true,
      fileId: driveFile.getId(),
      fileUrl: driveFile.getUrl(),
      downloadUrl: driveFile.getDownloadUrl(),
      fileName: driveFile.getName(),
      folderName: targetSubName,
      path: 'GMCC HUMAN CAPITAL/01 PERSONAL FILE/ACTIVE EMPLOYEE/' + empFolderName + '/' + targetSubName + '/' + fileName
    };
  }
};
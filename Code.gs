/**
 * SERVER CODE (Code.gs)
 * Google Apps Script backend for the UNY Vocational Faculty (IT Network Department) Internship Portal.
 * 
 * Instructions:
 * 1. Create a Google Spreadsheet.
 * 2. In Google Spreadsheet, go to Extensions > Apps Script.
 * 3. Replace the Code.gs content with this code.
 * 4. Create an Index.html file in the Apps Script project and paste the Index.html content.
 * 5. Configure the SCRIPT PROPERTIES in Apps Script Settings (or use the Setup button in the app to initialize them automatically).
 *    - DRIVE_FOLDER_ID: ID of the folder where member logbook file uploads are stored.
 *    - DOC_TEMPLATE_ID: ID of the Google Docs Portfolio template.
 *    - SLIDE_TEMPLATE_ID: ID of the Google Slides Certificate template.
 * 6. Deploy as a Web App (Execute as: "Me", Who has access: "Anyone").
 */

// Global Configurations / Property Keys
const PROP_DRIVE_FOLDER_ID = 'DRIVE_FOLDER_ID';
const PROP_DOC_TEMPLATE_ID = 'DOC_TEMPLATE_ID';
const PROP_SLIDE_TEMPLATE_ID = 'SLIDE_TEMPLATE_ID';

/**
 * Serves the HTML file when the Web App URL is loaded
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('UNY IT Network Internship Portal')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, shrink-to-fit=no');
}

/**
 * Access helper for the Active Spreadsheet
 */
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Initialize Tables, Sheets, and default data structures.
 * Can be run from the Spreadsheet or triggered via Web UI for auto-setup.
 */
function setupDatabase() {
  const ss = getSpreadsheet();
  
  // 1. Create Sheets with columns if they do not exist
  // Users: kolom diperluas dengan Periode, TanggalMulai, TanggalSelesai, NomorSurat
  const requiredSheets = {
    'Users': ['NIM', 'Name', 'Email', 'Password', 'Role', 'Periode', 'TanggalMulai', 'TanggalSelesai', 'NomorSurat'],
    'Jobdesks': ['RoleName', 'Jobdesks'],
    'Tasks': ['TaskId', 'TaskName', 'Category', 'Description', 'AssignedNIM', 'Status', 'CreatedBy'],
    'Logbooks': ['LogbookId', 'NIM', 'TaskId', 'TaskName', 'Category', 'Timestamp', 'WorkDescription', 'FileUrl', 'FileName', 'Grade', 'Notes']
  };
  
  for (let sheetName in requiredSheets) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(requiredSheets[sheetName]);
      sheet.getRange(1, 1, 1, requiredSheets[sheetName].length)
           .setFontWeight('bold')
           .setBackground('#003366')
           .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }
  }
  
  // 2. Pre-populate default Jobdesks if empty
  const jobdeskSheet = ss.getSheetByName('Jobdesks');
  if (jobdeskSheet.getLastRow() <= 1) {
    const defaultJobdesks = [
      ['Anggota', 'Mengikuti instruksi dari Ketua dan PIC\nMenyelesaikan tugas jaringan / website / admin yang ditugaskan\nMembuat laporan logbook harian dilengkapi berkas bukti dukung\nMenjaga kebersihan dan ketertiban ruang lab praktik'],
      ['Ketua', 'Mengoordinasi semua anggota magang\nMendelegasikan tugas-tugas harian (jaringan, website, admin)\nMemantau progress pengerjaan logbook anggota\nMenjadi jembatan komunikasi antara anggota dengan PIC Dosen'],
      ['Sekretaris', 'Mengelola administrasi surat menyurat magang\nMengarsipkan berkas-berkas digital magang\nMembantu penyusunan timeline program magang'],
      ['Koordinator Ruangan', 'Bertanggung jawab atas kerapian dan keamanan laboratorium\nMelakukan inventarisasi perangkat PC dan switch di lab\nMelaporkan kerusakan perangkat keras kepada PIC'],
      ['Koordinator Alat', 'Mengelola peminjaman router, kabel tester, dan toolkit\nMemastikan semua alat yang dipinjam kembali dengan selamat\nMelakukan QC berkala terhadap perangkat alat praktik'],
      ['Tim Support', 'Membantu instalasi OS dan software pendukung perkuliahan\nMelakukan penarikan kabel LAN dan troubleshooting jaringan harian\nMembantu dosen PIC dalam konfigurasi jaringan lokal']
    ];
    defaultJobdesks.forEach(function(row) { jobdeskSheet.appendRow(row); });
  }
  
  // 3. Populate default PIC user if empty
  const userSheet = ss.getSheetByName('Users');
  if (userSheet.getLastRow() <= 1) {
    userSheet.appendRow(['19600101', 'PIC Dosen Koordinator', 'pic.itnetwork@uny.ac.id', 'pic123', 'PIC', 12, '2026-01-01', '2026-12-31', '']);
  }
  
  return { success: true, message: 'Database Spreadsheet UNY berhasil di-setup!' };
}

/**
 * Helper: Format tanggal YYYY-MM-DD ke format Indonesia (1 Maret 2026)
 */
function formatTanggalID(dateStr) {
  if (!dateStr) return '-';
  const bulan = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return dateStr;
  return d.getDate() + ' ' + bulan[d.getMonth()] + ' ' + d.getFullYear();
}

/**
 * Helper: Hitung predikat dari nilai numerik
 */
function getPredicate(grade) {
  const g = parseFloat(grade);
  if (g >= 90) return 'SANGAT BAIK';
  if (g >= 80) return 'BAIK';
  if (g >= 70) return 'CUKUP';
  if (g >= 60) return 'KURANG';
  return 'TIDAK LULUS';
}

/**
 * Register a new user
 * Default role is 'Anggota'
 * Parameter opsional: periode, tanggalMulai, tanggalSelesai, nomorSurat
 */
function registerUser(nim, name, email, password, periode, tanggalMulai, tanggalSelesai, nomorSurat) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === nim.toString()) {
      return { success: false, message: 'NIM ' + nim + ' sudah terdaftar!' };
    }
  }
  
  sheet.appendRow([
    nim.toString(), name, email, password, 'Anggota',
    periode || 3,
    tanggalMulai || '',
    tanggalSelesai || '',
    nomorSurat || ''
  ]);
  return { success: true, message: 'Registrasi Berhasil! Silakan Login.' };
}

/**
 * Validate credentials and login
 */
function loginUser(nim, password) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === nim.toString() && data[i][3].toString() === password.toString()) {
      return {
        success: true,
        user: {
          nim: data[i][0],
          name: data[i][1],
          email: data[i][2],
          role: data[i][4],
          periode: data[i][5] || 3,
          tanggalMulai: data[i][6] || '',
          tanggalSelesai: data[i][7] || '',
          nomorSurat: data[i][8] || ''
        }
      };
    }
  }
  return { success: false, message: 'NIM atau Password salah!' };
}

/**
 * Get all users registered in the system
 */
function getUsersList() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  const users = [];
  
  for (let i = 1; i < data.length; i++) {
    users.push({
      nim: data[i][0],
      name: data[i][1],
      email: data[i][2],
      role: data[i][4]
    });
  }
  return users;
}

/**
 * Update a specific user's role (PIC Action)
 */
function updateUserRole(nim, newRole) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === nim.toString()) {
      sheet.getRange(i + 1, 5).setValue(newRole);
      return { success: true, message: 'Role NIM ' + nim + ' berhasil diperbarui ke ' + newRole };
    }
  }
  return { success: false, message: 'NIM tidak ditemukan!' };
}

/**
 * Get standard jobdesks for all roles
 */
function getJobdesks() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Jobdesks');
  const data = sheet.getDataRange().getValues();
  const jobdesks = {};
  
  for (let i = 1; i < data.length; i++) {
    jobdesks[data[i][0]] = data[i][1];
  }
  return jobdesks;
}

/**
 * Save / Update a list of jobdesks for a specific role (PIC Action)
 */
function saveJobdesk(roleName, jobdeskText) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Jobdesks');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === roleName) {
      sheet.getRange(i + 1, 2).setValue(jobdeskText);
      return { success: true, message: 'Jobdesk untuk ' + roleName + ' berhasil diperbarui!' };
    }
  }
  // If role doesn't exist, append new
  sheet.appendRow([roleName, jobdeskText]);
  return { success: true, message: 'Role baru ' + roleName + ' dan Jobdesk berhasil ditambahkan!' };
}

/**
 * Get core responsibilities for a specific active member's role
 */
function getMemberJobdesk(role) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Jobdesks');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === role) {
      return data[i][1];
    }
  }
  return "Jobdesk belum diinput oleh PIC.";
}

/**
 * Create and delegate new tasks (Ketua Action)
 */
function createTask(taskName, category, description, assignedNim, ketuaNim) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Tasks');
  const taskId = 'TSK-' + new Date().getTime();
  
  sheet.appendRow([taskId, taskName, category, description, assignedNim.toString(), 'Pending', ketuaNim.toString()]);
  return { success: true, message: 'Tugas berhasil didelegasikan ke NIM ' + assignedNim };
}

/**
 * Get assigned tasks for a specific member
 */
function getTasksForMember(nim) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Tasks');
  const data = sheet.getDataRange().getValues();
  const tasks = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][4].toString() === nim.toString()) {
      tasks.push({
        taskId: data[i][0],
        taskName: data[i][1],
        category: data[i][2],
        description: data[i][3],
        status: data[i][5],
        createdBy: data[i][6]
      });
    }
  }
  return tasks;
}

/**
 * Get tasks created/delegated by the Ketua
 */
function getTasksByKetua(ketuaNim) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Tasks');
  const data = sheet.getDataRange().getValues();
  const tasks = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][6].toString() === ketuaNim.toString()) {
      tasks.push({
        taskId: data[i][0],
        taskName: data[i][1],
        category: data[i][2],
        description: data[i][3],
        assignedNim: data[i][4],
        status: data[i][5]
      });
    }
  }
  return tasks;
}

/**
 * Complete a task, upload file to Google Drive and save to Logbooks sheet.
 * fileData must be an object with { base64: '...', mimeType: '...', name: '...' }
 */
function completeTaskAndUpload(nim, taskId, workDesc, fileObj) {
  try {
    const ss = getSpreadsheet();
    
    // 1. Upload File/Photo to Google Drive Folder
    let fileUrl = '';
    let fileName = '';
    
    const driveFolderId = PropertiesService.getScriptProperties().getProperty(PROP_DRIVE_FOLDER_ID);
    
    if (fileObj && fileObj.base64 && fileObj.name) {
      let folder;
      if (driveFolderId) {
        folder = DriveApp.getFolderById(driveFolderId);
      } else {
        // Fallback: Create folder at My Drive root
        const folders = DriveApp.getFoldersByName('UNY_Internship_Logbooks');
        if (folders.hasNext()) {
          folder = folders.next();
        } else {
          folder = DriveApp.createFolder('UNY_Internship_Logbooks');
        }
      }
      
      const fileBytes = Utilities.base64Decode(fileObj.base64);
      const blob = Utilities.newBlob(fileBytes, fileObj.mimeType, fileObj.name);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
      fileName = fileObj.name;
    }
    
    // 2. Fetch specific Task Name and Category
    const taskSheet = ss.getSheetByName('Tasks');
    const taskData = taskSheet.getDataRange().getValues();
    let taskName = 'Tugas Mandiri/Lain-lain';
    let category = 'Admin';
    let taskRowIndex = -1;
    
    for (let i = 1; i < taskData.length; i++) {
      if (taskData[i][0].toString() === taskId.toString()) {
        taskName = taskData[i][1];
        category = taskData[i][2];
        taskRowIndex = i + 1;
        break;
      }
    }
    
    // 3. Mark Task as Completed in Tasks Sheet
    if (taskRowIndex !== -1) {
      taskSheet.getRange(taskRowIndex, 6).setValue('Completed');
    }
    
    // 4. Save to Logbooks Sheet
    const logbookSheet = ss.getSheetByName('Logbooks');
    const logbookId = 'LOG-' + new Date().getTime();
    const timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
    
    logbookSheet.appendRow([
      logbookId,
      nim.toString(),
      taskId,
      taskName,
      category,
      timestamp,
      workDesc,
      fileUrl,
      fileName,
      '', // Nilai Akhir (graded by PIC later)
      ''  // Notes
    ]);
    
    return { success: true, message: 'Logbook pekerjaan berhasil diunggah!' };
  } catch (error) {
    return { success: false, message: 'Error Server: ' + error.toString() };
  }
}

/**
 * Get all completed logbooks for grading review (PIC View)
 */
function getCompletedLogbooks() {
  const ss = getSpreadsheet();
  const logbookSheet = ss.getSheetByName('Logbooks');
  const logbookData = logbookSheet.getDataRange().getValues();
  const userSheet = ss.getSheetByName('Users');
  const userData = userSheet.getDataRange().getValues();
  
  // Make a User map for quick Lookup of Student Name and Email
  const userMap = {};
  for (let i = 1; i < userData.length; i++) {
    userMap[userData[i][0].toString()] = {
      name: userData[i][1],
      email: userData[i][2]
    };
  }
  
  const logbooks = [];
  for (let i = 1; i < logbookData.length; i++) {
    const nimStr = logbookData[i][1].toString();
    const sName = userMap[nimStr] ? userMap[nimStr].name : 'Student';
    const sEmail = userMap[nimStr] ? userMap[nimStr].email : '';
    
    logbooks.push({
      logbookId: logbookData[i][0],
      nim: nimStr,
      studentName: sName,
      studentEmail: sEmail,
      taskId: logbookData[i][2],
      taskName: logbookData[i][3],
      category: logbookData[i][4],
      timestamp: logbookData[i][5],
      workDescription: logbookData[i][6],
      fileUrl: logbookData[i][7],
      fileName: logbookData[i][8],
      grade: logbookData[i][9],
      notes: logbookData[i][10]
    });
  }
  return logbooks;
}

/**
 * Grade a completed logbook log (PIC Action)
 */
function gradeLogbook(logbookId, grade, notes) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Logbooks');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === logbookId.toString()) {
      sheet.getRange(i + 1, 10).setValue(grade); // Column index 10 (J) is Grade
      sheet.getRange(i + 1, 11).setValue(notes); // Column index 11 (K) is Notes
      return { success: true, message: 'Logbook ' + logbookId + ' berhasil dinilai!' };
    }
  }
  return { success: false, message: 'Logbook tidak ditemukan!' };
}

/**
 * Core Automation Process (PIC Action):
 * 1. Mengisi semua {{TAG}} / <<TAG>> di Google Slides template → Export PDF Sertifikat
 * 2. Mengisi semua {{TAG}} / <<TAG>> di Google Docs template → Export PDF Portofolio  
 * 3. Menyimpan PDF ke Google Drive
 * 4. Mengirim PDF ke email mahasiswa secara otomatis
 *
 * TAG yang didukung di template Slides:
 *   {{NAMA}} / <<Nama>>      → Nama Mahasiswa
 *   {{NIM}}  / <<NIM>>       → Nomor Induk Mahasiswa
 *   {{NOMOR}}/ <<Nomor>>     → Nomor Surat Sertifikat
 *   {{PERIODE}}/<<Periode>>  → Lama magang (bulan)
 *   {{NILAI}} / <<Nilai>>    → Nilai numerik
 *   {{PREDIKAT}}/<<Predikat>>→ Predikat (BAIK, SANGAT BAIK, dll.)
 *   {{MULAI}} / <<Mulai>>    → Tanggal mulai (format Indonesia)
 *   {{SELESAI}}/<<Selesai>>  → Tanggal selesai (format Indonesia)
 *
 * TAG yang didukung di template Docs:
 *   {{NAMA}}, {{NIM}}, {{NOMOR}}, {{PERIODE}}, {{NILAI}}, {{PREDIKAT}},
 *   {{MULAI}}, {{SELESAI}}, {{LogbookTable}} / <<LogbookTable>>
 */
function generatePortfolioAndCertificate(studentNim, overallGrade) {
  try {
    const props = PropertiesService.getScriptProperties();
    const driveFolderId = props.getProperty(PROP_DRIVE_FOLDER_ID);
    const docTemplateId = props.getProperty(PROP_DOC_TEMPLATE_ID);
    const slideTemplateId = props.getProperty(PROP_SLIDE_TEMPLATE_ID);
    
    if (!driveFolderId || !docTemplateId || !slideTemplateId) {
      return { 
        success: false, 
        message: 'Lengkapi Konfigurasi API Template (Drive Folder ID, Doc Template ID, Slide Template ID) di menu Pengaturan PIC terlebih dahulu sebelum mencetak berkas.' 
      };
    }
    
    const ss = getSpreadsheet();
    
    // ─── Ambil Data Lengkap Mahasiswa dari Sheet Users ───────────────────────
    // Kolom: NIM(0) Name(1) Email(2) Password(3) Role(4) Periode(5) TanggalMulai(6) TanggalSelesai(7) NomorSurat(8)
    const userSheet = ss.getSheetByName('Users');
    const userData = userSheet.getDataRange().getValues();
    let studentName = '';
    let studentEmail = '';
    let studentPeriode = 3;
    let studentTanggalMulai = '';
    let studentTanggalSelesai = '';
    let studentNomorSurat = '';
    
    for (let i = 1; i < userData.length; i++) {
      if (userData[i][0].toString() === studentNim.toString()) {
        studentName        = userData[i][1] || '';
        studentEmail       = userData[i][2] || '';
        studentPeriode     = userData[i][5] || 3;
        studentTanggalMulai   = userData[i][6] ? userData[i][6].toString() : '';
        studentTanggalSelesai = userData[i][7] ? userData[i][7].toString() : '';
        studentNomorSurat  = userData[i][8] ? userData[i][8].toString() : '';
        break;
      }
    }
    
    if (!studentName) {
      return { success: false, message: 'Mahasiswa dengan NIM ' + studentNim + ' tidak ditemukan!' };
    }
    
    // ─── Siapkan Nilai Substitusi ────────────────────────────────────────────
    const predikat = getPredicate(overallGrade);
    const tanggalMulaiID   = formatTanggalID(studentTanggalMulai)   || '-';
    const tanggalSelesaiID = formatTanggalID(studentTanggalSelesai) || '-';
    
    // Peta semua pasangan TAG → Nilai (mendukung format {{TAG}} dan <<Tag>>)
    const replacements = [
      ['{{NAMA}}',    studentName],
      ['<<Nama>>',    studentName],
      ['{{NIM}}',     studentNim.toString()],
      ['<<NIM>>',     studentNim.toString()],
      ['{{NOMOR}}',   studentNomorSurat || '-'],
      ['<<Nomor>>',   studentNomorSurat || '-'],
      ['{{PERIODE}}', studentPeriode.toString()],
      ['<<Periode>>', studentPeriode.toString()],
      ['{{NILAI}}',   overallGrade.toString()],
      ['<<Nilai>>',   overallGrade.toString()],
      ['{{PREDIKAT}}', predikat],
      ['<<Predikat>>', predikat],
      ['{{MULAI}}',   tanggalMulaiID],
      ['<<Mulai>>',   tanggalMulaiID],
      ['{{SELESAI}}', tanggalSelesaiID],
      ['<<Selesai>>', tanggalSelesaiID],
    ];
    
    // ─── Ambil Logbook Mahasiswa ──────────────────────────────────────────────
    const logbookSheet = ss.getSheetByName('Logbooks');
    const logbookData = logbookSheet.getDataRange().getValues();
    const studentLogs = [];
    
    for (let i = 1; i < logbookData.length; i++) {
      if (logbookData[i][1].toString() === studentNim.toString()) {
        studentLogs.push({
          taskName: logbookData[i][3],
          category: logbookData[i][4],
          timestamp: logbookData[i][5],
          description: logbookData[i][6],
          fileUrl: logbookData[i][7]
        });
      }
    }
    
    const outputFolder = DriveApp.getFolderById(driveFolderId);
    
    // ════════════════════════════════════════════════════════════════════════
    // PART A: GENERATE CERTIFICATE — Google Slides template
    // Google Apps Script mengganti SEMUA {{TAG}} di semua slide, lalu export PDF
    // ════════════════════════════════════════════════════════════════════════
    const slideCopy   = DriveApp.getFileById(slideTemplateId).makeCopy('Sertifikat_' + studentName, outputFolder);
    const slideCopyId = slideCopy.getId();
    const presentation = SlidesApp.openById(slideCopyId);
    
    // replaceAllText() di level presentasi mengganti teks di SEMUA slide sekaligus
    replacements.forEach(function(pair) {
      presentation.replaceAllText(pair[0], pair[1]);
    });
    
    presentation.saveAndClose();
    
    // Export Slides → PDF (re-fetch untuk mendapat versi tersimpan terbaru)
    const certPdfBlob = DriveApp.getFileById(slideCopyId)
      .getAs('application/pdf')
      .setName('Sertifikat_Magang_UNY_' + studentName + '.pdf');
    
    // ════════════════════════════════════════════════════════════════════════
    // PART B: GENERATE PORTFOLIO — Google Docs template
    // ════════════════════════════════════════════════════════════════════════
    const docCopy   = DriveApp.getFileById(docTemplateId).makeCopy('Portofolio_' + studentName, outputFolder);
    const docCopyId = docCopy.getId();
    // Variable renamed from 'document' to 'portfolioDoc' to avoid shadowing GAS built-in
    const portfolioDoc = DocumentApp.openById(docCopyId);
    const body = portfolioDoc.getBody();
    
    // Ganti semua {{TAG}} dan <<Tag>> di body Docs
    replacements.forEach(function(pair) {
      body.replaceText(pair[0], pair[1]);
    });
    
    // ─── Sisipkan Tabel Logbook pada marker {{LogbookTable}} / <<LogbookTable>> ─
    let tableIndex = -1;
    for (let i = 0; i < body.getNumChildren(); i++) {
      let child = body.getChild(i);
      if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const txt = child.asParagraph().getText();
        if (txt.indexOf('{{LogbookTable}}') !== -1 || txt.indexOf('<<LogbookTable>>') !== -1) {
          tableIndex = i;
          child.asParagraph().setText(''); // Hapus teks placeholder
          break;
        }
      }
    }
    
    // Prepare table headers
    const tableData = [
      ['No', 'Nama Tugas', 'Kategori', 'Tanggal', 'Deskripsi Hasil Pekerjaan', 'Tautan Bukti Dukung']
    ];
    
    studentLogs.forEach(function(item, idx) {
      tableData.push([
        (idx + 1).toString(),
        item.taskName,
        item.category,
        item.timestamp ? item.timestamp.toString().substring(0, 10) : '-',
        item.description,
        item.fileUrl ? item.fileUrl : 'Tidak Ada File'
      ]);
    });
    
    let table;
    if (tableIndex !== -1) {
      table = body.insertTable(tableIndex + 1, tableData);
    } else {
      // Append at end of document if no marker found
      body.appendParagraph('\nREKAPITULASI LOGBOOK MAGANG:');
      table = body.appendTable(tableData);
    }
    
    // Style the generated table beautifully
    const headerStyle = {};
    headerStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#003366';
    headerStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#FFFFFF';
    headerStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
    headerStyle[DocumentApp.Attribute.FONT_SIZE] = 10;
    headerStyle[DocumentApp.Attribute.BOLD] = true;
    
    const rowStyle = {};
    rowStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
    rowStyle[DocumentApp.Attribute.FONT_SIZE] = 9;
    rowStyle[DocumentApp.Attribute.BOLD] = false;
    
    // Apply styling to table
    for (let r = 0; r < tableData.length; r++) {
      let row = table.getRow(r);
      for (let c = 0; c < tableData[r].length; c++) {
        let cell = row.getCell(c);
        if (r === 0) {
          cell.setAttributes(headerStyle);
        } else {
          cell.setAttributes(rowStyle);
          // Set cell text padding for a clean neat grid look
          cell.setPaddingTop(4);
          cell.setPaddingBottom(4);
        }
      }
    }
    
    portfolioDoc.saveAndClose();
    
    // Export Portfolio to PDF — re-fetch file to get latest saved version
    const portfolioPdfBlob = DriveApp.getFileById(docCopyId).getAs('application/pdf').setName('Portofolio_Magang_UNY_' + studentName + '.pdf');
    
    // --- PART C: SEND EMAIL WITH PDF ATTACHMENTS ---
    if (studentEmail) {
      const emailSubject = '[E-Sertifikat & Portofolio Magang] Fakultas Vokasi UNY IT Network';
      const emailBodyHTML = `
        <div style="font-family: Arial, sans-serif; max-width: 600px; border: 1px solid #ddd; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
          <div style="background-color: #003366; padding: 25px; text-align: center;">
            <h1 style="color: #FFD700; margin: 0; font-size: 24px; font-weight: bold; text-shadow: 1px 1px 2px rgba(0,0,0,0.5);">FAKULTAS VOKASI UNY</h1>
            <p style="color: #ffffff; margin: 5px 0 0 0; font-size: 14px; letter-spacing: 1px;">DEPARTEMEN IT NETWORK & INFRASTRUKTUR</p>
          </div>
          <div style="padding: 30px; line-height: 1.6; color: #333333; background-color: #ffffff;">
            <p style="margin-top: 0; font-size: 16px;">Yth. <b>${studentName}</b> (NIM. ${studentNim}),</p>
            <p style="text-align: justify;">Selamat! Anda telah resmi menyelesaikan seluruh rangkaian program magang industri di Departemen IT Network Fakultas Vokasi Universitas Negeri Yogyakarta dengah hasil yang membanggakan.</p>
            <p style="text-align: justify;">Dosen PIC Pembimbing Lapangan telah memverifikasi portofolio logbook harian Anda dan memberikan rekapitulasi penilaian akhir sebagai berikut:</p>
            
            <div style="background-color: #f7f9fc; border-left: 4px solid #003366; padding: 15px; margin: 20px 0; border-radius: 4px;">
              <table style="width: 100%; font-size: 15px;">
                <tr><td style="width: 40%; color: #666;">Nama Lengkap:</td><td><b>${studentName}</b></td></tr>
                <tr><td style="color: #666;">Nomor Induk Mahasiswa:</td><td><b>${studentNim}</b></td></tr>
                <tr><td style="color: #666;">Predikat Penilaian:</td><td><b style="color: #003366; font-size: 18px;">${overallGrade} (Sangat Baik / Lulus)</b></td></tr>
              </table>
            </div>
            
            <p style="text-align: justify;">Terlampir kami kirimkan dokumen e-sertifikat kelulusan magang resmi dan berkas portofolio magang harian Anda dalam format PDF pendukung karir Anda untuk diunduh.</p>
            <p style="margin-bottom: 0;">Terima kasih atas dedikasi luar biasa Anda demi kemajuan infrastruktur IT Fakultas Vokasi UNY.</p>
          </div>
          <div style="background-color: #f1f3f5; padding: 15px 30px; font-color: #888888; font-size: 11px; text-align: center; border-top: 1px solid #e9ecef;">
            Email ini dibuat secara otomatis oleh UNY Apps Script Integration System.<br>
            © 2026 Fakultas Vokasi Universitas Negeri Yogyakarta. All Rights Reserved.
          </div>
        </div>
      `;
      
      MailApp.sendEmail({
        to: studentEmail,
        subject: emailSubject,
        htmlBody: emailBodyHTML,
        attachments: [certPdfBlob, portfolioPdfBlob]
      });
    }
    
    // Clean up temporary workspace copies to keep drive tidy (Optional but highly recommended)
    // SlideCopy.setTrashed(true);
    // DocCopy.setTrashed(true);
    
    return { 
      success: true, 
      message: 'Sertifikat dan Portofolio PDF ' + studentName + ' berhasil dibuat, distor ke Drive, dan dikirimkan ke Email ' + studentEmail + '!' 
    };
  } catch (error) {
    return { success: false, message: 'Gagal mencetak & mengirim berkas: ' + error.toString() };
  }
}

/**
 * Save configuration properties for Templates
 */
function saveConfiguration(driveId, docId, slideId) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty(PROP_DRIVE_FOLDER_ID, driveId);
    props.setProperty(PROP_DOC_TEMPLATE_ID, docId);
    props.setProperty(PROP_SLIDE_TEMPLATE_ID, slideId);
    return { success: true, message: 'Pengaturan Template Google API Berhasil Disimpan!' };
  } catch (error) {
    return { success: false, message: 'Gagal menyimpan: ' + error.toString() };
  }
}

/**
 * Load configuration properties
 */
function getConfiguration() {
  const props = PropertiesService.getScriptProperties();
  return {
    driveFolderId: props.getProperty(PROP_DRIVE_FOLDER_ID) || '',
    docTemplateId: props.getProperty(PROP_DOC_TEMPLATE_ID) || '',
    slideTemplateId: props.getProperty(PROP_SLIDE_TEMPLATE_ID) || ''
  };
}

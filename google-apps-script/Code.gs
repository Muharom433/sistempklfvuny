/**
 * SERVER CODE (Code.gs) — VERSI GABUNGAN FINAL
 * Google Apps Script backend for the UNY Vocational Faculty (IT Network Department) Internship Portal.
 * 
 * Instructions:
 * 1. Create a Google Spreadsheet.
 * 2. In Google Spreadsheet, go to Extensions > Apps Script.
 * 3. Replace the Code.gs content with this code.
 * 4. Create an Index.html file in the Apps Script project and paste the Index.html content.
 * 5. (Opsional) Isi FALLBACK_SPREADSHEET_ID jika script berjalan standalone (bukan dari dalam Spreadsheet).
 * 6. Configure the SCRIPT PROPERTIES di menu Pengaturan PIC di web app:
 *    - DRIVE_FOLDER_ID     : ID folder Google Drive tempat file upload disimpan.
 *    - DOC_TEMPLATE_ID     : ID template Google Docs LOGBOOK HARIAN.
 *    - PORTFOLIO_TEMPLATE_ID: ID template Google Docs PORTOFOLIO AKHIR (BERBEDA dari logbook!).
 *    - SLIDE_TEMPLATE_ID   : ID template Google Slides SERTIFIKAT KELULUSAN.
 * 7. Deploy as a Web App (Execute as: "Me", Who has access: "Anyone").
 */

// Global Configurations / Property Keys
const PROP_DRIVE_FOLDER_ID      = 'DRIVE_FOLDER_ID';
const PROP_DOC_TEMPLATE_ID      = 'DOC_TEMPLATE_ID';       // Template Logbook Harian
const PROP_PORTFOLIO_TEMPLATE_ID = 'PORTFOLIO_TEMPLATE_ID'; // Template Portofolio Akhir
const PROP_SLIDE_TEMPLATE_ID    = 'SLIDE_TEMPLATE_ID';

// JIKA MENGALAMI ERROR "Cannot read properties of null (reading 'getSheetByName')",
// PASTE ID SPREADSHEET GOOGLE ANDA DI DALAM TANDA KUTIP DI BAWAH INI:
const FALLBACK_SPREADSHEET_ID = '';

/**
 * Serves the HTML file when the Web App URL is loaded,
 * OR handles GET API requests if action parameter is provided.
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'export_certificate') {
    try {
      const payloadStr = e.parameter.payload;
      if (!payloadStr) throw new Error("Payload data kosong");
      const payload = JSON.parse(payloadStr);
      const result = generatePortfolioAndCertificate(
        payload.studentData,
        payload.logbooks,
        payload.driveId,
        payload.portfolioId,
        payload.slideId
      );
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (error) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Fallback: Serve UI HTML
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('UNY IT Network Internship Portal')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, shrink-to-fit=no');
}

/**
 * Handles incoming POST requests — API router utama dari React frontend.
 * Mendukung action:
 *   - 'export_certificate'   : Generate PDF Sertifikat + Portofolio (Base64)
 *   - 'generate_logbook_doc' : Salin template Logbook Harian (lama — dipertahankan untuk kompatibilitas)
 *   - 'generateTaskLogbook'  : Salin template Logbook Harian (versi baru, lebih lengkap)
 */
function doPost(e) {
  try {
    if (!e || !e.postData) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'No data received' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    const payload = params.payload || {};
    let result;
    
    switch (action) {
      case 'export_certificate':
        result = generatePortfolioAndCertificate(
          payload.studentData,
          payload.logbooks,
          payload.driveId,
          payload.portfolioId,  // ← Sekarang menggunakan portfolioId, bukan docId
          payload.slideId
        );
        break;
      
      case 'generateTaskLogbook':
        result = generateTaskLogbook(
          payload.studentNim,
          payload.taskId,
          payload.taskName,
          payload.category,
          payload.description,
          payload.docReportTitle,
          payload.docReportOverview,
          payload.docReportSteps,
          payload.docReportChallenges,
          payload.docReportConclusion,
          payload.timelineLogs || []
        );
        break;
      
      case 'generate_logbook_doc':
        // Versi lama — tetap berfungsi
        result = generateLogbookDoc(
          payload.studentData,
          payload.taskData,
          payload.driveId,
          payload.docId
        );
        break;
      
      default:
        result = { success: false, message: 'Action tidak dikenal: ' + action };
    }
    
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Server Error: ' + error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * Access helper for the Active Spreadsheet
 * — mendukung FALLBACK_SPREADSHEET_ID jika script standalone
 */
function getSpreadsheet() {
  const activeSs = SpreadsheetApp.getActiveSpreadsheet();
  if (activeSs) return activeSs;
  
  if (FALLBACK_SPREADSHEET_ID && FALLBACK_SPREADSHEET_ID.trim() !== '') {
    return SpreadsheetApp.openById(FALLBACK_SPREADSHEET_ID);
  }
  
  throw new Error("Spreadsheet tidak terdeteksi. Isi FALLBACK_SPREADSHEET_ID di Code.gs dengan ID Spreadsheet Anda.");
}

/**
 * Initialize Tables, Sheets, and default data structures.
 */
function setupDatabase() {
  const ss = getSpreadsheet();
  
  const requiredSheets = {
    'Users':    ['NIM', 'Name', 'Email', 'Password', 'Role', 'Periode', 'TanggalMulai', 'TanggalSelesai', 'NomorSurat'],
    'Jobdesks': ['RoleName', 'Jobdesks'],
    'Tasks':    ['TaskId', 'TaskName', 'Category', 'Description', 'AssignedNIM', 'Status', 'CreatedBy'],
    'Logbooks': ['LogbookId', 'NIM', 'TaskId', 'TaskName', 'Category', 'Timestamp', 'WorkDescription', 'FileUrl', 'FileName', 'Grade', 'Notes']
  };
  
  for (let sheetName in requiredSheets) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(requiredSheets[sheetName]);
      sheet.getRange(1, 1, 1, requiredSheets[sheetName].length)
           .setFontWeight('bold').setBackground('#003366').setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }
  }
  
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

function registerUser(nim, name, email, password, periode, tanggalMulai, tanggalSelesai, nomorSurat) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === nim.toString()) {
      return { success: false, message: 'NIM ' + nim + ' sudah terdaftar!' };
    }
  }
  sheet.appendRow([nim.toString(), name, email, password, 'Anggota', periode || 3, tanggalMulai || '', tanggalSelesai || '', nomorSurat || '']);
  return { success: true, message: 'Registrasi Berhasil! Silakan Login.' };
}

function loginUser(nim, password) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === nim.toString() && data[i][3].toString() === password.toString()) {
      return { success: true, user: { nim: data[i][0], name: data[i][1], email: data[i][2], role: data[i][4], periode: data[i][5] || 3, tanggalMulai: data[i][6] || '', tanggalSelesai: data[i][7] || '', nomorSurat: data[i][8] || '' } };
    }
  }
  return { success: false, message: 'NIM atau Password salah!' };
}

function getUsersList() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  const users = [];
  for (let i = 1; i < data.length; i++) {
    users.push({ nim: data[i][0], name: data[i][1], email: data[i][2], role: data[i][4] });
  }
  return users;
}

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

function getJobdesks() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Jobdesks');
  const data = sheet.getDataRange().getValues();
  const jobdesks = {};
  for (let i = 1; i < data.length; i++) { jobdesks[data[i][0]] = data[i][1]; }
  return jobdesks;
}

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
  sheet.appendRow([roleName, jobdeskText]);
  return { success: true, message: 'Role baru ' + roleName + ' dan Jobdesk berhasil ditambahkan!' };
}

function getMemberJobdesk(role) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Jobdesks');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === role) return data[i][1];
  }
  return "Jobdesk belum diinput oleh PIC.";
}

function createTask(taskName, category, description, assignedNim, ketuaNim) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Tasks');
  const taskId = 'TSK-' + new Date().getTime();
  sheet.appendRow([taskId, taskName, category, description, assignedNim.toString(), 'Pending', ketuaNim.toString()]);
  return { success: true, message: 'Tugas berhasil didelegasikan ke NIM ' + assignedNim };
}

function getTasksForMember(nim) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Tasks');
  const data = sheet.getDataRange().getValues();
  const tasks = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][4].toString() === nim.toString()) {
      tasks.push({ taskId: data[i][0], taskName: data[i][1], category: data[i][2], description: data[i][3], status: data[i][5], createdBy: data[i][6] });
    }
  }
  return tasks;
}

function getTasksByKetua(ketuaNim) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Tasks');
  const data = sheet.getDataRange().getValues();
  const tasks = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][6].toString() === ketuaNim.toString()) {
      tasks.push({ taskId: data[i][0], taskName: data[i][1], category: data[i][2], description: data[i][3], assignedNim: data[i][4], status: data[i][5] });
    }
  }
  return tasks;
}

function completeTaskAndUpload(nim, taskId, workDesc, fileObj) {
  try {
    const ss = getSpreadsheet();
    let fileUrl = '', fileName = '';
    const driveFolderId = PropertiesService.getScriptProperties().getProperty(PROP_DRIVE_FOLDER_ID);
    if (fileObj && fileObj.base64 && fileObj.name) {
      let folder;
      if (driveFolderId) { folder = DriveApp.getFolderById(driveFolderId); }
      else {
        const folders = DriveApp.getFoldersByName('UNY_Internship_Logbooks');
        folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('UNY_Internship_Logbooks');
      }
      const fileBytes = Utilities.base64Decode(fileObj.base64);
      const blob = Utilities.newBlob(fileBytes, fileObj.mimeType, fileObj.name);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl(); fileName = fileObj.name;
    }
    const taskSheet = ss.getSheetByName('Tasks');
    const taskData = taskSheet.getDataRange().getValues();
    let taskName = 'Tugas Mandiri/Lain-lain', category = 'Admin', taskRowIndex = -1;
    for (let i = 1; i < taskData.length; i++) {
      if (taskData[i][0].toString() === taskId.toString()) { taskName = taskData[i][1]; category = taskData[i][2]; taskRowIndex = i + 1; break; }
    }
    if (taskRowIndex !== -1) taskSheet.getRange(taskRowIndex, 6).setValue('Completed');
    const logbookSheet = ss.getSheetByName('Logbooks');
    const logbookId = 'LOG-' + new Date().getTime();
    const timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
    logbookSheet.appendRow([logbookId, nim.toString(), taskId, taskName, category, timestamp, workDesc, fileUrl, fileName, '', '']);
    return { success: true, message: 'Logbook pekerjaan berhasil diunggah!' };
  } catch (error) { return { success: false, message: 'Error Server: ' + error.toString() }; }
}

function getCompletedLogbooks() {
  const ss = getSpreadsheet();
  const logbookData = ss.getSheetByName('Logbooks').getDataRange().getValues();
  const userData = ss.getSheetByName('Users').getDataRange().getValues();
  const userMap = {};
  for (let i = 1; i < userData.length; i++) { userMap[userData[i][0].toString()] = { name: userData[i][1], email: userData[i][2] }; }
  const logbooks = [];
  for (let i = 1; i < logbookData.length; i++) {
    const nimStr = logbookData[i][1].toString();
    const u = userMap[nimStr] || { name: 'Student', email: '' };
    logbooks.push({ logbookId: logbookData[i][0], nim: nimStr, studentName: u.name, studentEmail: u.email, taskId: logbookData[i][2], taskName: logbookData[i][3], category: logbookData[i][4], timestamp: logbookData[i][5], workDescription: logbookData[i][6], fileUrl: logbookData[i][7], fileName: logbookData[i][8], grade: logbookData[i][9], notes: logbookData[i][10] });
  }
  return logbooks;
}

function gradeLogbook(logbookId, grade, notes) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Logbooks');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === logbookId.toString()) {
      sheet.getRange(i + 1, 10).setValue(grade);
      sheet.getRange(i + 1, 11).setValue(notes);
      return { success: true, message: 'Logbook ' + logbookId + ' berhasil dinilai!' };
    }
  }
  return { success: false, message: 'Logbook tidak ditemukan!' };
}

/**
 * ════════════════════════════════════════════════════════════════════════
 * GENERATE SERTIFIKAT + PORTOFOLIO AKHIR (PDF Base64)
 * — Menggunakan PORTFOLIO_TEMPLATE_ID (bukan logbook template!)
 * — Data mahasiswa dikirim langsung dari React (tidak ambil dari Sheets)
 * — Output: PDF Base64 bisa langsung di-download oleh browser
 *
 * TAG di template Slides Sertifikat:
 *   {{NAMA}}, {{NIM}}, {{NOMOR}}, {{PERIODE}}, {{NILAI}}, {{PREDIKAT}}, {{MULAI}}, {{SELESAI}}, {{PERAN}}
 *
 * TAG di template Docs Portofolio:
 *   {{NAMA}}, {{NIM}}, {{NOMOR}}, {{PERIODE}}, {{NILAI}}, {{PREDIKAT}}, {{MULAI}}, {{SELESAI}},
 *   {{LogbookTable}} / <<LogbookTable>> / {{LOGBOOK}}
 * ════════════════════════════════════════════════════════════════════════
 */
function generatePortfolioAndCertificate(studentData, logbooks, paramDriveId, paramPortfolioId, paramSlideId) {
  try {
    const props = PropertiesService.getScriptProperties();
    const slideTemplateId     = paramSlideId     || props.getProperty(PROP_SLIDE_TEMPLATE_ID);
    const portfolioTemplateId = paramPortfolioId || props.getProperty(PROP_PORTFOLIO_TEMPLATE_ID);
    
    if (!slideTemplateId) {
      return { success: false, message: 'Slide Template ID (Sertifikat) belum diisi di Pengaturan PIC!' };
    }
    
    const studentName          = studentData.name           || '';
    const studentNim           = studentData.nim            || '';
    const studentEmail         = studentData.email          || '';
    const studentRole          = studentData.role           || 'Anggota';
    const studentPeriode       = studentData.periode        || 3;
    const studentTanggalMulai  = studentData.tanggalMulai  || '';
    const studentTanggalSelesai= studentData.tanggalSelesai|| '';
    const studentNomorSurat    = studentData.nomorSurat     || '-';
    const overallGrade         = studentData.overallGrade  || 0;
    
    if (!studentName) return { success: false, message: 'Data mahasiswa kosong atau tidak valid!' };
    
    const predikat         = getPredicate(overallGrade);
    const tanggalMulaiID   = formatTanggalID(studentTanggalMulai)    || '-';
    const tanggalSelesaiID = formatTanggalID(studentTanggalSelesai)  || '-';
    
    const replacements = [
      ['{{NAMA}}',    studentName],      ['<<Nama>>',    studentName],
      ['{{NIM}}',     studentNim],       ['<<NIM>>',     studentNim],
      ['{{NIS}}',     studentNim],       ['<<NIS>>',     studentNim],
      ['{{NOMOR}}',   studentNomorSurat],['<<Nomor>>',   studentNomorSurat],
      ['{{PERIODE}}', studentPeriode.toString()],['<<Periode>>', studentPeriode.toString()],
      ['{{NILAI}}',   overallGrade.toString()],  ['<<Nilai>>',   overallGrade.toString()],
      ['{{RATARATA}}',overallGrade.toString()],  ['<<RataRata>>',overallGrade.toString()],
      ['{{PREDIKAT}}',predikat],         ['<<Predikat>>',predikat],
      ['{{MULAI}}',   tanggalMulaiID],   ['<<Mulai>>',   tanggalMulaiID],
      ['{{SELESAI}}', tanggalSelesaiID], ['<<Selesai>>', tanggalSelesaiID],
      ['{{PERAN}}',   studentRole],      ['<<Peran>>',   studentRole],
    ];
    
    const studentLogs  = logbooks || [];
    const outputFolder = DriveApp.getRootFolder(); // file sementara, lalu dihapus
    
    // ── PART A: SERTIFIKAT — Google Slides ────────────────────────────────
    const slideCopy = DriveApp.getFileById(slideTemplateId).makeCopy('Sertifikat_' + studentName, outputFolder);
    const presentation = SlidesApp.openById(slideCopy.getId());
    replacements.forEach(function(pair) { try { presentation.replaceAllText(pair[0], pair[1]); } catch(e) {} });
    presentation.saveAndClose();
    const certPdfBlob = slideCopy.getAs('application/pdf');
    const certBase64  = Utilities.base64Encode(certPdfBlob.getBytes());
    slideCopy.setTrashed(true);
    
    // ── PART B: PORTOFOLIO AKHIR — Google Docs ───────────────────────────
    let portBase64 = null;
    if (portfolioTemplateId) {
      try {
        const docCopy = DriveApp.getFileById(portfolioTemplateId).makeCopy('Portofolio_' + studentName, outputFolder);
        const doc  = DocumentApp.openById(docCopy.getId());
        const body = doc.getBody();
        
        replacements.forEach(function(pair) { try { body.replaceText(pair[0], pair[1] || ' '); } catch(e) {} });
        try { body.replaceText('Nama: Afif', 'Nama: ' + studentName); } catch(e) {}
        try { body.replaceText('NIS: 1',     'NIM: '  + studentNim);  } catch(e) {}
        
        // --- 1. GENERATE PIE CHART ---
        if (studentLogs && studentLogs.length > 0) {
          const categoryCounts = {};
          studentLogs.forEach(function(log) {
            const cat = log.category || 'Lain-lain';
            categoryCounts[cat] = (categoryCounts[cat] || 0) + 1;
          });
          
          let dataBuilder = Charts.newDataTable()
            .addColumn(Charts.ColumnType.STRING, 'Kategori')
            .addColumn(Charts.ColumnType.NUMBER, 'Jumlah');
          
          for (const cat in categoryCounts) {
            dataBuilder.addRow([cat, categoryCounts[cat]]);
          }
          
          const chart = Charts.newPieChart()
            .setDataTable(dataBuilder.build())
            .setDimensions(500, 300)
            .set3D()
            .build();
            
          const chartBlob = chart.getAs('image/png');
          
          const piechartElement = body.findText('{{PIECHART}}');
          if (piechartElement) {
            const el = piechartElement.getElement();
            el.getParent().asParagraph().insertInlineImage(0, chartBlob);
            el.asText().setText('');
          }
        } else {
          try { body.replaceText('{{PIECHART}}', '(Belum ada data)'); } catch(e) {}
        }
        
        // --- 2. GENERATE TABLE & QR CODES ---
        const tables = body.getTables();
        let targetTable = null;
        let templateRowIndex = -1;
        
        // Cari tabel yang punya marker {{JUDUL}}
        for (let i = 0; i < tables.length; i++) {
          const t = tables[i];
          for (let r = 0; r < t.getNumRows(); r++) {
            const row = t.getRow(r);
            if (row.getText().indexOf('{{JUDUL}}') !== -1) {
              targetTable = t;
              templateRowIndex = r;
              break;
            }
          }
          if (targetTable) break;
        }
        
        if (targetTable && templateRowIndex !== -1) {
          const templateRow = targetTable.getRow(templateRowIndex);
          
          if (studentLogs.length === 0) {
            templateRow.getCell(0).setText('(Belum ada riwayat pekerjaan)');
            templateRow.getCell(1).setText('-');
            templateRow.getCell(2).setText('-');
          } else {
            studentLogs.forEach(function(log) {
              const newRow = targetTable.appendTableRow(templateRow.copy());
              try { newRow.replaceText('{{JUDUL}}', log.name || log.taskName || ' '); } catch(e) {}
              try { newRow.replaceText('{{NILAI}}', (log.grade || '-').toString()); } catch(e) {}
              
              const qrElement = newRow.findText('{{QR}}');
              if (qrElement) {
                const linkStr = log.googleDocUrl || log.file || '';
                const par = qrElement.getElement().getParent();
                if (linkStr) {
                  try {
                    const qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=100x100&data=' + encodeURIComponent(linkStr);
                    const qrBlob = UrlFetchApp.fetch(qrUrl).getBlob();
                    qrElement.getElement().asText().setText('');
                    if (par.getType() === DocumentApp.ElementType.PARAGRAPH) {
                      par.asParagraph().insertInlineImage(0, qrBlob);
                    }
                  } catch(e) {
                    qrElement.getElement().asText().setText(linkStr);
                  }
                } else {
                  qrElement.getElement().asText().setText('-');
                }
              }
            });
            targetTable.removeRow(templateRowIndex);
          }
        } else {
          // Fallback jika tidak menemukan tabel dengan {{JUDUL}}, buat tabel baru (Backward Compatibility)
          let tableIndex = -1;
          for (let i = 0; i < body.getNumChildren(); i++) {
            const child = body.getChild(i);
            if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
              const txt = child.asParagraph().getText();
              if (txt.indexOf('{{LOGBOOK}}') !== -1 || txt.indexOf('<<LogbookTable>>') !== -1 || txt.indexOf('{{LogbookTable}}') !== -1) {
                tableIndex = i; child.asParagraph().setText(' '); break;
              }
            }
          }
          
          const tableData = [['Judul Tugas', 'Nilai', 'Link G-Doc (QR)']];
          if (studentLogs.length === 0) {
            tableData.push(['(Belum ada riwayat pekerjaan)', '-', '-']);
          } else {
            studentLogs.forEach(function(log) { tableData.push([log.name || log.taskName || ' ', (log.grade || '-').toString(), log.googleDocUrl || log.file || '-']); });
          }
          
          let table;
          if (tableIndex !== -1) { table = body.insertTable(tableIndex + 1, tableData); }
          else { body.appendParagraph('\nDAFTAR RIWAYAT PEKERJAAN:'); table = body.appendTable(tableData); }
          
          const hStyle = {}; hStyle[DocumentApp.Attribute.BACKGROUND_COLOR]='#003366'; hStyle[DocumentApp.Attribute.FOREGROUND_COLOR]='#FFFFFF'; hStyle[DocumentApp.Attribute.BOLD]=true;
          const rStyle = {}; rStyle[DocumentApp.Attribute.BOLD]=false;
          for (let r = 0; r < table.getNumRows(); r++) {
            const row = table.getRow(r);
            for (let c = 0; c < row.getNumCells(); c++) {
              const cell = row.getCell(c);
              if (r === 0) { cell.setAttributes(hStyle); } else { cell.setAttributes(rStyle); cell.setPaddingTop(5); cell.setPaddingBottom(5); }
            }
          }
        }
        
        doc.saveAndClose();
        portBase64 = Utilities.base64Encode(docCopy.getAs('application/pdf').getBytes());
        docCopy.setTrashed(true);
      } catch (docError) {
        Logger.log('Portfolio generation error: ' + docError.toString());
      }
    }
    
    return { success: true, message: 'Sertifikat dan Portofolio PDF ' + studentName + ' berhasil dibuat!', certBase64: certBase64, portBase64: portBase64 };
    
  } catch (error) {
    return { success: false, message: 'Gagal mencetak berkas: ' + error.toString() };
  }
}

/**
 * ════════════════════════════════════════════════════════════════════════
 * GENERATE TASK LOGBOOK — Versi Baru (dari modal "Gandakan Laporan")
 * — Menggunakan DOC_TEMPLATE_ID (Template Logbook Harian)
 * — Mengisi: {{NAMA}}, {{NIM}}, {{JUDUL}}, {{KATEGORI}}, {{DESKRIPSI}},
 *            {{PEMBAHASAN}}, {{LANGKAH}}, {{KENDALA}}, {{KESIMPULAN}}
 * — Menyisipkan tabel timeline kerja harian jika ada marker {{TimelineTable}}
 * — Otomatis simpan ke sheet Logbooks & tandai tugas Completed
 * ════════════════════════════════════════════════════════════════════════
 */
function generateTaskLogbook(studentNim, taskId, taskName, category, description, docReportTitle, docReportOverview, docReportSteps, docReportChallenges, docReportConclusion, timelineLogs) {
  try {
    const props = PropertiesService.getScriptProperties();
    const driveFolderId     = props.getProperty(PROP_DRIVE_FOLDER_ID);
    const logbookTemplateId = props.getProperty(PROP_DOC_TEMPLATE_ID);
    
    if (!driveFolderId || !logbookTemplateId) {
      return { success: false, message: 'Lengkapi Drive Folder ID dan Logbook Template ID di Pengaturan PIC.' };
    }
    
    const ss = getSpreadsheet();
    const userData = ss.getSheetByName('Users').getDataRange().getValues();
    let studentName = 'Mahasiswa';
    for (let i = 1; i < userData.length; i++) {
      if (userData[i][0].toString() === studentNim.toString()) { studentName = userData[i][1] || 'Mahasiswa'; break; }
    }
    
    const outputFolder = DriveApp.getFolderById(driveFolderId);
    const docCopy = DriveApp.getFileById(logbookTemplateId).makeCopy('Logbook_' + taskName + '_' + studentName, outputFolder);
    const doc  = DocumentApp.openById(docCopy.getId());
    const body = doc.getBody();
    
    const replacements = [
      ['{{NAMA}}',      studentName],       ['<<Nama>>',      studentName],
      ['{{NIM}}',       studentNim.toString()],['<<NIM>>',    studentNim.toString()],
      ['{{JUDUL}}',     docReportTitle || taskName],['<<Judul>>',docReportTitle || taskName],
      ['{{KATEGORI}}',  category || '-'],   ['<<Kategori>>', category || '-'],
      ['{{DESKRIPSI}}', description || '-'],['<<Deskripsi>>',description || '-'],
      ['{{PEMBAHASAN}}',docReportOverview || '-'],['<<Pembahasan>>',docReportOverview || '-'],
      ['{{LANGKAH}}',   docReportSteps || '-'],   ['<<Langkah>>',   docReportSteps || '-'],
      ['{{KENDALA}}',   docReportChallenges || '-'],['<<Kendala>>',  docReportChallenges || '-'],
      ['{{KESIMPULAN}}',docReportConclusion || '-'],['<<Kesimpulan>>',docReportConclusion || '-'],
      // Tangani hardcode di template lama
      ['Nama: Afif',   'Nama: ' + studentName],
      ['NIS: 1',       'NIM: '  + studentNim.toString()],
    ];
    
    replacements.forEach(function(pair) { try { body.replaceText(pair[0], pair[1] || ' '); } catch(e) {} });
    
    // Sisipkan tabel timeline jika ada marker
    if (timelineLogs && timelineLogs.length > 0) {
      let tableIndex = -1;
      for (let i = 0; i < body.getNumChildren(); i++) {
        const child = body.getChild(i);
        if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
          const txt = child.asParagraph().getText();
          if (txt.indexOf('{{TimelineTable}}') !== -1 || txt.indexOf('<<TimelineTable>>') !== -1 ||
              txt.indexOf('{{LOGBOOK}}')       !== -1 || txt.indexOf('<<LogbookTable>>')  !== -1) {
            tableIndex = i;
            child.asParagraph().setText(' ');
            break;
          }
        }
      }
      const tableData = [['Tanggal', 'Aktivitas Riil', 'Durasi']];
      timelineLogs.forEach(function(log) { tableData.push([log.date || ' ', log.description || ' ', (log.hours || '') + ' Jam']); });
      let table;
      if (tableIndex !== -1) { table = body.insertTable(tableIndex + 1, tableData); }
      else { body.appendParagraph('\nLOGBOOK KEGIATAN:'); table = body.appendTable(tableData); }
      const hStyle = {}; hStyle[DocumentApp.Attribute.BACKGROUND_COLOR]='#003366'; hStyle[DocumentApp.Attribute.FOREGROUND_COLOR]='#FFFFFF'; hStyle[DocumentApp.Attribute.BOLD]=true;
      for (let c = 0; c < tableData[0].length; c++) { table.getRow(0).getCell(c).setAttributes(hStyle); }
    }
    
    doc.saveAndClose();
    docCopy.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    
    // Simpan ke sheet Logbooks
    const logbookSheet = ss.getSheetByName('Logbooks');
    const newLogbookId = 'LOG-' + new Date().getTime();
    const timestamp    = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
    logbookSheet.appendRow([newLogbookId, studentNim.toString(), taskId, taskName, category, timestamp, docReportOverview || description, docCopy.getUrl(), 'Logbook_' + taskName + '_' + studentName, '', '']);
    
    // Tandai tugas selesai
    const taskData = ss.getSheetByName('Tasks').getDataRange().getValues();
    const taskSheet = ss.getSheetByName('Tasks');
    for (let i = 1; i < taskData.length; i++) {
      if (taskData[i][0].toString() === taskId.toString()) { taskSheet.getRange(i + 1, 6).setValue('Completed'); break; }
    }
    
    return { success: true, message: 'GDoc Logbook berhasil digandakan!', fileUrl: docCopy.getUrl(), docId: docCopy.getId() };
    
  } catch (error) {
    return { success: false, message: 'Gagal menggandakan GDoc Logbook: ' + error.toString() };
  }
}

/**
 * ════════════════════════════════════════════════════════════════════════
 * GENERATE LOGBOOK DOC — Versi Lama (dari tombol "Gandakan Template" di tabel tugas)
 * Dipertahankan untuk kompatibilitas, menggunakan struktur data lama (studentData, taskData)
 * ════════════════════════════════════════════════════════════════════════
 */
function generateLogbookDoc(studentData, taskData, paramDriveId, paramDocId) {
  try {
    const props = PropertiesService.getScriptProperties();
    const driveFolderId  = paramDriveId || props.getProperty(PROP_DRIVE_FOLDER_ID);
    const logbookTemplId = paramDocId   || props.getProperty(PROP_DOC_TEMPLATE_ID);
    
    if (!driveFolderId || !logbookTemplId) {
      return { success: false, message: 'Lengkapi Drive Folder ID dan Logbook Template ID di Pengaturan PIC.' };
    }
    
    const studentName = studentData.name || ' ';
    const studentNim  = studentData.nim  || ' ';
    const studentRole = studentData.role || 'Anggota';
    const taskName     = taskData.taskName  || ' ';
    const taskCategory = taskData.category  || ' ';
    const checklists   = taskData.points    || [];
    const checkDates   = taskData.checkDates || [];
    
    const outputFolder = DriveApp.getFolderById(driveFolderId);
    const docCopy = DriveApp.getFileById(logbookTemplId).makeCopy('Draf Laporan Hasil ' + taskName + ' - ' + studentName, outputFolder);
    const doc  = DocumentApp.openById(docCopy.getId());
    const body = doc.getBody();
    docCopy.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    
    const replacements = [
      ['{{JUDUL}}',    taskName],       ['{{KATEGORI}}', taskCategory],
      ['{{NAMA}}',     studentName],    ['Nama: Afif',   'Nama: ' + studentName],
      ['{{NIM}}',      studentNim],     ['NIS: 1',       'NIM: '  + studentNim],
      ['{{PERAN}}',    studentRole],    ['<<Peran>>',    studentRole],
    ];
    replacements.forEach(function(pair) { try { body.replaceText(pair[0], pair[1] || ' '); } catch(e) {} });
    
    let tableIndex = -1;
    for (let i = 0; i < body.getNumChildren(); i++) {
      const child = body.getChild(i);
      if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const txt = child.asParagraph().getText();
        if (txt.indexOf('{{LOGBOOK}}') !== -1 || txt.indexOf('<<LOGBOOK>>') !== -1 || txt.indexOf('{{LogbookTable}}') !== -1 || txt.indexOf('<<LogbookTable>>') !== -1) {
          tableIndex = i; child.asParagraph().setText(' '); break;
        }
      }
    }
    
    const tableData = [['Tanggal', 'Deskripsi Tugas', 'Catatan']];
    if (checklists.length === 0) { tableData.push([' ', '(Belum ada checklist capaian)', ' ']); }
    else { checklists.forEach(function(point, idx) { tableData.push([(checkDates && checkDates[idx]) ? checkDates[idx].toString() : ' ', point.toString() || ' ', ' ']); }); }
    
    let table;
    if (tableIndex !== -1) { table = body.insertTable(tableIndex + 1, tableData); }
    else { body.appendParagraph('\nLOGBOOK KEGIATAN:'); table = body.appendTable(tableData); }
    
    const hStyle = {}; hStyle[DocumentApp.Attribute.BACKGROUND_COLOR]='#003366'; hStyle[DocumentApp.Attribute.FOREGROUND_COLOR]='#FFFFFF'; hStyle[DocumentApp.Attribute.BOLD]=true; hStyle[DocumentApp.Attribute.FONT_SIZE]=11;
    const rStyle = {}; rStyle[DocumentApp.Attribute.BOLD]=false; rStyle[DocumentApp.Attribute.FONT_SIZE]=10;
    for (let r = 0; r < tableData.length; r++) {
      const row = table.getRow(r);
      for (let c = 0; c < tableData[r].length; c++) {
        const cell = row.getCell(c);
        if (r === 0) { cell.setAttributes(hStyle); } else { cell.setAttributes(rStyle); cell.setPaddingTop(5); cell.setPaddingBottom(5); }
      }
    }
    
    doc.saveAndClose();
    return { success: true, docId: docCopy.getId(), docUrl: 'https://docs.google.com/document/d/' + docCopy.getId() + '/edit', message: 'Logbook berhasil di-generate!' };
  } catch (error) { return { success: false, message: 'Gagal generate Laporan Doc: ' + error.toString() }; }
}

/**
 * Save configuration properties for Templates
 */
function saveConfiguration(driveId, logbookId, portfolioId, slideId) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty(PROP_DRIVE_FOLDER_ID,      driveId);
    props.setProperty(PROP_DOC_TEMPLATE_ID,      logbookId);
    props.setProperty(PROP_PORTFOLIO_TEMPLATE_ID, portfolioId);
    props.setProperty(PROP_SLIDE_TEMPLATE_ID,    slideId);
    return { success: true, message: 'Pengaturan Template Google API Berhasil Disimpan!' };
  } catch (error) { return { success: false, message: 'Gagal menyimpan: ' + error.toString() }; }
}

/**
 * Load configuration properties
 */
function getConfiguration() {
  const props = PropertiesService.getScriptProperties();
  return {
    driveFolderId:      props.getProperty(PROP_DRIVE_FOLDER_ID)       || '',
    docTemplateId:      props.getProperty(PROP_DOC_TEMPLATE_ID)       || '',
    portfolioTemplateId: props.getProperty(PROP_PORTFOLIO_TEMPLATE_ID) || '',
    slideTemplateId:    props.getProperty(PROP_SLIDE_TEMPLATE_ID)     || ''
  };
}

import React, { useState, useEffect } from 'react';
import { 
  BookOpen, 
  FileCode, 
  CheckCircle, 
  Award, 
  FileText, 
  Mail, 
  Users, 
  Check, 
  Copy, 
  ExternalLink, 
  Folder, 
  Download, 
  Database, 
  Sparkles, 
  RefreshCw, 
  Laptop, 
  ChevronRight, 
  Settings, 
  Plus, 
  UserCheck, 
  AlertCircle,
  FileCode2,
  Trash2,
  Info,
  Pencil,
  Cloud,
  AlertTriangle,
  Send
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import Select from 'react-select';
import Swal from 'sweetalert2';
import { initAuth, googleSignIn, logout, getAccessToken } from './lib/firebase';
import { checkAndCreateSheets } from './lib/sync';
import { db } from './lib/supabase';


// Ready-to-copy code templates pre-loaded into the React Client
const CODE_GS_CONTENT = `/**
 * SERVER CODE (Code.gs)
 * Google Apps Script backend for the UNY Vocational Faculty (IT Network Department) Internship Portal.
 * 
 * Instructions:
 * 1. Create a Google Spreadsheet.
 * 2. In Google Spreadsheet, go to Extensions > Apps Script.
 * 3. Replace the Code.gs content with this code.
 * 4. Create an Index.html file in the Apps Script project and paste the Index.html content.
 * 5. Run setupDatabase() and configure the API Properties via the UI frontend.
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
 * Config Getter: Pulls parameter from the "Properties" spreadsheet tab
 */
function getConfigProperty(key) {
  const sheet = getSpreadsheet().getSheetByName('Properties');
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return data[i][1];
    }
  }
  return null;
}

/**
 * Config Setter: Saves parameter to the "Properties" spreadsheet tab
 */
function setConfigProperty(key, value) {
  const sheet = getSpreadsheet().getSheetByName('Properties');
  if (!sheet) return { success: false, message: 'Sheet Properties tidak ditemukan.' };
  
  const data = sheet.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      sheet.getRange(i + 1, 3).setValue(new Date().toISOString());
      found = true;
      break;
    }
  }
  
  if (!found) {
    sheet.appendRow([key, value, new Date().toISOString()]);
  }
  return { success: true, message: 'Berhasil menyimpan properti ' + key };
}

/**
 * Save multiple config properties
 */
function saveAllConfigProperties(driveId, docId, slideId) {
  setConfigProperty(PROP_DRIVE_FOLDER_ID, driveId);
  setConfigProperty(PROP_DOC_TEMPLATE_ID, docId);
  setConfigProperty(PROP_SLIDE_TEMPLATE_ID, slideId);
  return { success: true, message: 'Parameters saved to Spreadsheet' };
}

/**
 * Initialize Tables, Sheets, and data structures.
 */
function setupDatabase() {
  const ss = getSpreadsheet();
  const requiredSheets = {
    'Users': ['NIM', 'Name', 'Email', 'Password', 'Role'],
    'Jobdesks': ['RoleName', 'Jobdesks'],
    'Tasks': ['TaskId', 'TaskName', 'Category', 'Description', 'AssignedNIM', 'Status', 'CreatedBy'],
    'Logbooks': ['LogbookId', 'NIM', 'TaskId', 'TaskName', 'Category', 'Timestamp', 'WorkDescription', 'FileUrl', 'FileName', 'Grade', 'Notes'],
    'Properties': ['PropertyKey', 'PropertyValue', 'LastUpdated']
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
  
  const jobdeskSheet = ss.getSheetByName('Jobdesks');
  if (jobdeskSheet.getLastRow() <= 1) {
    const defaultJobdesks = [
      ['Anggota', 'Mengikuti instruksi dari Ketua dan PIC\\nMenyelesaikan tugas jaringan / website / admin yang ditugaskan\\nMembuat laporan logbook harian dilengkapi berkas bukti dukung\\nMenjaga kebersihan dan ketertiban ruang lab praktik'],
      ['Ketua', 'Mengoordinasi semua anggota magang\\nMendelegasikan tugas-tugas harian (jaringan, website, admin)\\nMemantau progress pengerjaan logbook anggota\\nMenjadi jembatan komunikasi antara anggota dengan PIC Dosen'],
      ['Sekretaris', 'Mengelola administrasi surat menyurat magang\\nMengarsipkan berkas-berkas digital magang\\nMembantu penyusunan timeline program magang'],
      ['Koordinator Ruangan', 'Bertanggung jawab atas kerapian dan keamanan laboratorium\\nMelakukan inventarisasi perangkat PC dan switch di lab\\nMelaporkan kerusakan perangkat keras kepada PIC'],
      ['Koordinator Alat', 'Mengelola peminjaman router, kabel tester, dan toolkit\\nMemastikan semua alat yang dipinjam kembali dengan selamat\\nMelakukan QC berkala terhadap perangkat alat praktik'],
      ['Tim Support', 'Membantu instalasi OS dan software pendukung perkuliahan\\nMelakukan penarikan kabel LAN dan troubleshooting jaringan harian\\nMembantu dosen PIC dalam konfigurasi jaringan lokal']
    ];
    defaultJobdesks.forEach(function(row) {
      jobdeskSheet.appendRow(row);
    });
  }
  return { success: true, message: 'Database Spreadsheet UNY berhasil di-setup!' };
}`;

const INDEX_HTML_CONTENT = `<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <title>Portal Magang IT Network - Fakultas Vokasi UNY</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
  <style>
    :root { --uny-blue: #003366; --uny-gold: #FFD700; }
    body { background-color: #F3F5F7; font-family: sans-serif; }
    .navbar { background-color: var(--uny-blue) !important; border-bottom: 4px solid var(--uny-gold); }
    .btn-primary { background-color: var(--uny-blue); border-color: var(--uny-blue); }
    .hero-banner { background: linear-gradient(135deg, var(--uny-blue) 0%, #002244 100%); color: white; border-bottom: 6px solid var(--uny-gold); }
  </style>
</head>
<body>
  <!-- See full code in sidebar and copy ready-to-run Bootstrap 5 file from editor! -->
</body>
</html>`;

interface User {
  nim: string;
  name: string;
  email: string;
  role: string;
  password?: string;
  periode?: number;
  tanggalMulai?: string;
  tanggalSelesai?: string;
  nomorSurat?: string;
}

interface MasterTask {
  id: string;
  title: string;
  category: string;
  description: string;
  workType: 'Individu' | 'Kelompok';
  targetRole?: string; // New: target role for the master task
  points?: string[]; // Checklist items defined by the PIC
}

interface Task {
  taskId: string;
  masterTaskId?: string; // Reference to picnic masterTask ID
  taskName: string;
  category: string;
  description: string;
  assignedNim: string;
  status: 'Pending' | 'Completed';
  createdBy: string;
  workType: 'Individu' | 'Kelompok';
  googleDocId?: string;
  googleDocUrl?: string;
  googleDocTitle?: string;
  points?: string[]; // Checklist copy from MasterTask
  pointsChecked?: boolean[]; // Checked state for each point harian
  checkDates?: string[]; // Date each point was checked
  timelineLogs?: { date: string; description: string; hours: number }[]; // daily timeline logs
  docContent?: {
    judul: string;
    instansi: string;
    pembahasan: string;
    tantangan: string;
    statusAkhir: string;
  };
}

interface Logbook {
  logbookId: string;
  masterTaskId?: string; // Reference to picnic masterTask ID
  nim: string;
  studentName: string;
  studentEmail: string;
  taskId: string;
  taskName: string;
  category: string;
  timestamp: string;
  workDescription: string;
  fileUrl: string;
  fileName: string;
  grade: string;
  notes: string;
  hoursSpent?: number | string;
  workType: 'Individu' | 'Kelompok';
  googleDocId?: string;
  googleDocUrl?: string;
  googleDocTitle?: string;
  docContent?: {
    judul: string;
    instansi: string;
    pembahasan: string;
    tantangan: string;
    statusAkhir: string;
  };
}

interface AppProperty {
  propKey: string;
  propValue: string;
  lastUpdated: string;
}

const SUPABASE_SQL_CODE = `-- Supabase SQL Schema for FV-UNY Logbook App
-- Jalankan kode SQL ini di SQL Editor Supabase Anda

-- 1. Table users (Data Pengguna)
CREATE TABLE users (
  nim VARCHAR PRIMARY KEY,
  name VARCHAR NOT NULL,
  email VARCHAR NOT NULL,
  role VARCHAR NOT NULL,
  password VARCHAR NOT NULL,
  periode INTEGER,
  tanggalMulai VARCHAR,
  tanggalSelesai VARCHAR,
  nomorSurat VARCHAR
);

-- 2. Table jobdesks (Peran & Deskripsi Kerja / Bank Peran)
CREATE TABLE jobdesks (
  roleName VARCHAR PRIMARY KEY,
  description TEXT NOT NULL
);

-- 3. Table master_tasks (Bank Tugas Pokok)
CREATE TABLE master_tasks (
  id VARCHAR PRIMARY KEY,
  title VARCHAR NOT NULL,
  category VARCHAR,
  description TEXT,
  points JSONB, 
  workType VARCHAR NOT NULL,
  targetRole VARCHAR NOT NULL
);

-- 4. Table tasks (Penugasan ke Mahasiswa)
CREATE TABLE tasks (
  id VARCHAR PRIMARY KEY,
  masterId VARCHAR REFERENCES master_tasks(id) ON DELETE CASCADE,
  assignedNim VARCHAR REFERENCES users(nim) ON DELETE CASCADE,
  taskName VARCHAR NOT NULL,
  category VARCHAR,
  dateAssigned VARCHAR NOT NULL,
  status VARCHAR NOT NULL,
  progress INTEGER DEFAULT 0,
  completedDesc TEXT,
  completedDate VARCHAR,
  googleDocUrl VARCHAR,
  points JSONB
);

-- 5. Table logbooks (Catatan Harian Pekerjaan)
CREATE TABLE logbooks (
  logbookId VARCHAR PRIMARY KEY,
  taskId VARCHAR REFERENCES tasks(id) ON DELETE CASCADE,
  nim VARCHAR REFERENCES users(nim) ON DELETE CASCADE,
  taskName VARCHAR NOT NULL,
  category VARCHAR,
  date VARCHAR NOT NULL,
  workDescription TEXT NOT NULL,
  hoursSpent INTEGER NOT NULL,
  grade VARCHAR,
  gradeNote TEXT,
  googleDocUrl VARCHAR
);

-- 6. Table categories (Kategori Kegiatan Tambahan - Opsional)
CREATE TABLE categories (
  name VARCHAR PRIMARY KEY
);

-- 7. Table app_state (Sinkronisasi Cepat JSON)
CREATE TABLE app_state (
  id VARCHAR PRIMARY KEY,
  data JSONB NOT NULL
);

ALTER TABLE users ENABLE ROW LEVEL SECURITY;
ALTER TABLE jobdesks ENABLE ROW LEVEL SECURITY;
ALTER TABLE master_tasks ENABLE ROW LEVEL SECURITY;
ALTER TABLE tasks ENABLE ROW LEVEL SECURITY;
ALTER TABLE logbooks ENABLE ROW LEVEL SECURITY;
ALTER TABLE categories ENABLE ROW LEVEL SECURITY;
ALTER TABLE app_state ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Enable all for anon (users)" ON users FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Enable all for anon (jobdesks)" ON jobdesks FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Enable all for anon (master_tasks)" ON master_tasks FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Enable all for anon (tasks)" ON tasks FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Enable all for anon (logbooks)" ON logbooks FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Enable all for anon (categories)" ON categories FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Enable all for anon (app_state)" ON app_state FOR ALL USING (true) WITH CHECK (true);
`;

export default function App() {
  // Navigation Tabs states
  const [activeTab, setActiveTab] = useState<'simulator' | 'code_gs' | 'index_html' | 'db_schema' | 'deploy_guide'>('simulator');
  const [isProductionViewMode, setIsProductionViewMode] = useState<boolean>(true);
  
  // Modals state
  const [appModal, setAppModal] = useState<{ title: string, content: string, type: 'error' | 'success' | 'info' | 'warning' } | null>(null);
  const [confirmDialog, setConfirmDialog] = useState<{ message: string, onConfirm: () => void } | null>(null);

  const customAlert = (msg: string, type: 'error' | 'success' | 'info' | 'warning' = 'info', forceTitle?: string) => {
    let title = "Pemberitahuan";
    if (type === 'error') title = "Kesalahan Dideteksi";
    if (type === 'success') title = "Sukses";
    if (type === 'warning') title = "Peringatan";
    if (forceTitle) title = forceTitle;
    setAppModal({ title, content: msg, type });
  };

  const customConfirm = (message: string, onConfirm: () => void) => {
    setConfirmDialog({ message, onConfirm });
  };

  // Copy feedback triggers
  const [copiedGs, setCopiedGs] = useState(false);
  const [copiedHtml, setCopiedHtml] = useState(false);

  // --- SIMULATOR DATABASE STATE (with LocalStorage fallbacks when Supabase not configured) ---
  const hasSupabase = !!import.meta.env.VITE_SUPABASE_URL;
  
  const [users, setUsers] = useState<User[]>(() => {
    if (hasSupabase) return []; // Will be loaded from Supabase
    const saved = localStorage.getItem('uny_sim_users');
    let initialUsers: User[] = [
      { nim: '19600101', name: 'Dr. Eng. Ir. Rian, M.T. (Dosen PIC)', email: 'pic.itnetwork@uny.ac.id', role: 'PIC', password: 'password', periode: 12, tanggalMulai: '2026-01-01', tanggalSelesai: '2026-12-31' },
      { nim: '2153214002', name: 'Andi Setyawan', email: 'andi.setyawan@student.uny.ac.id', role: 'Ketua', password: 'password', periode: 3, tanggalMulai: '2026-03-01', tanggalSelesai: '2026-06-01' },
      { nim: '2153214003', name: 'Budi Hartono', email: 'budi.hartono@student.uny.ac.id', role: 'Anggota', password: 'password', periode: 3, tanggalMulai: '2026-03-01', tanggalSelesai: '2026-06-01' },
      { nim: '2153214004', name: 'Chandra Wijaya', email: 'chandra.w@student.uny.ac.id', role: 'Anggota', password: 'password', periode: 3, tanggalMulai: '2026-03-01', tanggalSelesai: '2026-06-01' }
    ];

    if (saved) {
      // Migrate old saved users if they don't have periode
      const parsed = JSON.parse(saved);
      initialUsers = parsed.map((u: any) => ({
        ...u,
        periode: u.periode || 3,
        tanggalMulai: u.tanggalMulai || '2026-03-01',
        tanggalSelesai: u.tanggalSelesai || '2026-06-01'
      }));
    }

    // Force user 23090620052 to exist and be PIC
    const existingSpecialUserIndex = initialUsers.findIndex(u => u.nim === '23090620052');
    if (existingSpecialUserIndex >= 0) {
      initialUsers[existingSpecialUserIndex] = {
        ...initialUsers[existingSpecialUserIndex],
        password: 'Muh4120m',
        role: 'PIC'
      };
    } else {
      initialUsers.push({
        nim: '23090620052',
        name: 'Dosen PIC Tambahan',
        email: 'pic2@uny.ac.id',
        role: 'PIC',
        password: 'Muh4120m',
        periode: 12,
        tanggalMulai: '2026-01-01',
        tanggalSelesai: '2026-12-31'
      });
    }

    return initialUsers;
  });

  const [masterTasks, setMasterTasks] = useState<MasterTask[]>(() => {
    if (hasSupabase) return []; // Will be loaded from Supabase
    const saved = localStorage.getItem('uny_sim_master_tasks');
    if (saved) return JSON.parse(saved);
    return [
      {
        id: 'MT-1',
        title: 'Segmentasi Switch & VLAN Vokasi Lantai 2',
        category: 'Jaringan',
        description: 'Lakukan instalasi trunking port dan mapping range VLAN ID 10 (dosen) dan 20 (lab komputer) pada switch Cisco.',
        workType: 'Individu',
        targetRole: 'Tim Support',
        points: [
          'Lakukan survei topologi switch lantai 2',
          'Hubungkan trunk port port 1 s.d. 4',
          'Setup VLAN ID 10 (Dosen) & 20 (Lab)',
          'Lakukan pengetesan ping & routing ketersambungan'
        ]
      },
      {
        id: 'MT-2',
        title: 'Perbaikan Total Jalur Kabel Backbone LAN Gedung',
        category: 'Jaringan',
        description: 'Lakukan sensor kabel rumpang, reorganisasi patch panel lab, labeling, crimping ulang, dan tes redaman kabel.',
        workType: 'Kelompok',
        targetRole: 'Tim Support',
        points: [
          'Identifikasi kabel rumpang di patch panel',
          'Lakukan labeling jalur backbone gedung baru',
          'Crimping ulang konektor RJ45 yang bermasalah',
          'Ukur redaman kabel dengan LAN Tester / OTDR'
        ]
      },
      {
        id: 'MT-3',
        title: 'Pembaruan Modul Pendaftaran Mahasiswa Baru',
        category: 'Website',
        description: 'Optimalisasi database query form pendaftaran vokasi dan sinkronisasi endpoint data secara instan.',
        workType: 'Individu',
        targetRole: 'Anggota',
        points: [
          'Analisis slow query pada database pendaftaran',
          'Buat indexing pada kolom NIM & email pendaftar',
          'Sinkronisasi endpoint API pendaftaran harian',
          'Lakukan stress test load concurrency pendaftaran'
        ]
      },
      {
        id: 'MT-4',
        title: 'Instalasi & Deployment Web Server Virtual Lab',
        category: 'Website',
        description: 'Setup sistem operasi Linux Ubuntu Server, Nginx reverse proxy, keamanan SSL Let\'s Encrypt, dan deploy Git.',
        workType: 'Kelompok',
        targetRole: 'Anggota',
        points: [
          'Instalasi & konfigurasi OS Ubuntu Server LTS',
          'Setup web server Nginx & reverse proxy upstream',
          'Konfigurasi enkripsi SSL Let\'s Encrypt gratis',
          'Setting auto-deploy Git webhook triggers'
        ]
      },
      {
        id: 'MT-5',
        title: 'Sensus & Inventarisasi Perangkat Router Lab Vokasi',
        category: 'Admin',
        description: 'Lakukan pemindaian serial number, mac address, status keaktifan router Mikrotik, switch Cisco, dan toolkit praktikum.',
        workType: 'Kelompok',
        targetRole: 'Koordinator Alat',
        points: [
          'Catat serial number router Mikrotik laboratorium',
          'Catat MAC address Cisco Switch lantai 1 s.d. 3',
          'Uji fungsional & identifikasi status perangkat',
          'Kompilasi rekap spreadsheet inventarisasi final'
        ]
      },
      {
        id: 'MT-6',
        title: 'Arsip Surat Masuk & Pengisian Agenda Surat Dinas',
        category: 'Admin',
        description: 'Menata dokumen pendaftaran magang dari kemitraan industri, mencatat nomor surat menyurat, dan input lembar disposisi.',
        workType: 'Individu',
        targetRole: 'Sekretaris',
        points: [
          'Pengumpulan formulir cetak magang mitra industri',
          'Pencatatan nomor urut surat masuk internal',
          'Pengarsipan digital & penyimpanan link folder Drive',
          'Penyusunan draf lembar disposisi pimpinan vokasi'
        ]
      }
    ];
  });

  const [tasks, setTasks] = useState<Task[]>(() => {
    if (hasSupabase) return []; // Will be loaded from Supabase
    const saved = localStorage.getItem('uny_sim_tasks');
    if (saved) return JSON.parse(saved);
    return [
      {
        taskId: 'TSK-1',
        taskName: 'Segmentasi Switch & VLAN Vokasi Lantai 2',
        category: 'Jaringan',
        description: 'Lakukan instalasi trunking port dan mapping range VLAN ID 10 (dosen) dan 20 (lab komputer) pada switch Cisco.',
        assignedNim: '2153214003',
        status: 'Pending',
        createdBy: '19600101',
        workType: 'Individu',
        points: [
          'Lakukan survei topologi switch lantai 2',
          'Hubungkan trunk port port 1 s.d. 4',
          'Setup VLAN ID 10 (Dosen) & 20 (Lab)',
          'Lakukan pengetesan ping & routing ketersambungan'
        ],
        pointsChecked: [true, true, false, false],
        timelineLogs: [
          { date: '24 Juni 2026', description: 'Survei dan cek kelayakan port switch Cisco lantai 2', hours: 3 },
          { date: '25 Juni 2026', description: 'Melakukan trunking port 1 s.d. 4 pada Rack Switch', hours: 4 }
        ],
        docContent: {
          judul: 'Laporan Segmentasi Switch & VLAN Lantai 2',
          instansi: 'Fakultas Vokasi Universitas Negeri Yogyakarta',
          pembahasan: 'Mengonfigurasi switch Cisco Catalyst agar port port ethernet dikelompokkan ke dalam VLAN ID 10 untuk dosen dan VLAN ID 20 untuk praktikum lab komputer mahasiswa.',
          tantangan: 'Kabel patch cord lama agak longgar, diganti dengan yang baru.',
          statusAkhir: 'Selesai Sebagian (Progres 50%)'
        }
      },
      {
        taskId: 'TSK-2',
        taskName: 'Pembaruan Modul Pendaftaran Mahasiswa Baru',
        category: 'Website',
        description: 'Optimalisasi database query form pendaftaran vokasi dan sinkronisasi endpoint data secara instan.',
        assignedNim: '2153214004',
        status: 'Completed',
        createdBy: '19600101',
        workType: 'Individu',
        googleDocId: '1Doc_report_TSK_2_2153214004',
        googleDocUrl: 'https://docs.google.com/document/d/1Doc_report_TSK_2_2153214004/edit',
        googleDocTitle: 'Laporan Tugas Pembaruan Pendaftaran Baru v1.0',
        points: [
          'Analisis slow query pada database pendaftaran',
          'Buat indexing pada kolom NIM & email pendaftar',
          'Sinkronisasi endpoint API pendaftaran harian',
          'Lakukan stress test load concurrency pendaftaran'
        ],
        pointsChecked: [true, true, true, true],
        timelineLogs: [
          { date: '22 Juni 2026', description: 'Profiling database query, dideteksi lambat pada join tabel user', hours: 4 },
          { date: '23 Juni 2026', description: 'Membuat compound index pada tabel pendaftaran vokasi', hours: 3 },
          { date: '24 Juni 2026', description: 'Pengujian endpoint asinkronous terintegrasi harian', hours: 5 }
        ],
        docContent: {
          judul: 'Laporan Pembaruan Landing Page & Modul Form Pendaftaran Mahasiswa Baru',
          instansi: 'Fakultas Vokasi Universitas Negeri Yogyakarta',
          pembahasan: 'Mengganti Layout Grid CSS default dengan framework Bootstrap 5. Konfigurasi endpoint REST API agar database Google Sheets dapat dipopulasi secara asynchronous menggunakan google.script.run.',
          tantangan: 'Browser sering lagging saat rendering canvas. Diselesaikan dengan debouncing event resize.',
          statusAkhir: 'Selesai dan Siap Deploy'
        }
      }
    ];
  });

  const [logbooks, setLogbooks] = useState<Logbook[]>(() => {
    if (hasSupabase) return []; // Will be loaded from Supabase
    const saved = localStorage.getItem('uny_sim_logbooks');
    if (saved) return JSON.parse(saved);
    return [
      {
        logbookId: 'LOG-1',
        nim: '2153214004',
        studentName: 'Chandra Wijaya',
        studentEmail: 'chandra.w@student.uny.ac.id',
        taskId: 'TSK-2',
        taskName: 'Pembaruan Modul Pendaftaran Mahasiswa Baru',
        category: 'Website',
        timestamp: '2026-06-20 15:42:11',
        workDescription: 'Telah merapikan grid UI website vokasi uny ke format responsif Bootstrap 5, memperbarui formulir masukan database, dan menguji load time backend.',
        fileUrl: 'https://images.unsplash.com/photo-1547082299-de196ea013d6?q=80&w=600',
        fileName: 'screenshot_form_pendaftaran.png',
        grade: '92',
        notes: 'Sangat mengagumkan, struktur grid responsif berjalan lancar di seluler.',
        workType: 'Individu',
        googleDocId: '1Doc_report_TSK_2_2153214004',
        googleDocUrl: 'https://docs.google.com/document/d/1Doc_report_TSK_2_2153214004/edit',
        googleDocTitle: 'Laporan Tugas Pembaruan Pendaftaran Baru v1.0',
        docContent: {
          judul: 'Laporan Pembaruan Landing Page & Modul Form Pendaftaran Mahasiswa Baru',
          instansi: 'Fakultas Vokasi Universitas Negeri Yogyakarta',
          pembahasan: 'Mengganti Layout Grid CSS default dengan framework Bootstrap 5. Konfigurasi endpoint REST API agar database Google Sheets dapat dipopulasi secara asynchronous menggunakan google.script.run.',
          tantangan: 'Browser sering lagging saat rendering canvas. Diselesaikan dengan debouncing event resize.',
          statusAkhir: 'Selesai dan Siap Deploy'
        }
      }
    ];
  });

  const [jobdesks, setJobdesks] = useState<Record<string, string>>(() => {
    const saved = localStorage.getItem('uny_sim_jobdesks');
    if (saved) return JSON.parse(saved);
    return {
      'Anggota': 'Mengikuti instruksi dari Ketua dan PIC\nMenyelesaikan tugas jaringan / website / admin yang ditugaskan\nMembuat laporan logbook harian dilengkapi berkas bukti dukung\nMenjaga kebersihan dan ketertiban ruang lab praktik',
      'Ketua': 'Mengoordinasi semua anggota magang\nMendelegasikan tugas-tugas harian (jaringan, website, admin)\nMemantau progress pengerjaan logbook anggota\nMenjadi jembatan komunikasi antara anggota dengan PIC Dosen',
      'Sekretaris': 'Mengelola administrasi surat menyurat magang\nMengarsipkan berkas-berkas digital magang\nMembantu penyusunan timeline program magang',
      'Koordinator Ruangan': 'Bertanggung jawab atas kerapian dan keamanan laboratorium\nMelakukan inventarisasi perangkat PC dan switch di lab\nMelaporkan kerusakan perangkat keras kepada PIC',
      'Koordinator Alat': 'Mengelola peminjaman router, kabel tester, dan toolkit\nMemastikan semua alat yang di pinjam kembali dengan selamat\nMelakukan QC berkala terhadap perangkat alat praktik',
      'Tim Support': 'Membantu instalasi OS dan software pendukung perkuliahan\nMelakukan penarikan kabel LAN dan troubleshooting jaringan harian\nMembantu dosen PIC dalam konfigurasi jaringan lokal'
    };
  });

  // Keep state sync with LocalStorage (only when Supabase is NOT configured)
  useEffect(() => {
    if (!hasSupabase) localStorage.setItem('uny_sim_users', JSON.stringify(users));
  }, [users]);
  useEffect(() => {
    if (!hasSupabase) localStorage.setItem('uny_sim_master_tasks', JSON.stringify(masterTasks));
  }, [masterTasks]);
  useEffect(() => {
    if (!hasSupabase) localStorage.setItem('uny_sim_tasks', JSON.stringify(tasks));
  }, [tasks]);
  useEffect(() => {
    if (!hasSupabase) localStorage.setItem('uny_sim_logbooks', JSON.stringify(logbooks));
  }, [logbooks]);
  useEffect(() => {
    if (!hasSupabase) localStorage.setItem('uny_sim_jobdesks', JSON.stringify(jobdesks));
  }, [jobdesks]);

  // Category management state
  const [categories, setCategories] = useState<string[]>(() => {
    if (hasSupabase) return []; // Will be loaded from Supabase
    return ['Jaringan', 'Website', 'Admin'];
  });
  const [newCategoryName, setNewCategoryName] = useState('');
  const [editingCategoryOld, setEditingCategoryOld] = useState<string | null>(null);
  const [editingCategoryNew, setEditingCategoryNew] = useState('');

  // --- SIMULATOR RUNTIME STATE ---
  const [activeUser, setActiveUser] = useState<User | null>(() => {
    const saved = localStorage.getItem('activeUser');
    if (saved) {
      try { return JSON.parse(saved); } catch(e) {}
    }
    return null;
  });

  useEffect(() => {
    if (activeUser) {
      localStorage.setItem('activeUser', JSON.stringify(activeUser));
    } else {
      localStorage.removeItem('activeUser');
    }
  }, [activeUser]);

  // --- SUPABASE AUTO SYNC (REAL RELATIONAL TABLES) ---
  const [isDataLoaded, setIsDataLoaded] = useState(false);
  const [isSyncingSession, setIsSyncingSession] = useState(false);

  // App API configuration states
  const [propertiesData, setPropertiesData] = useState<AppProperty[]>([
    { propKey: 'SUPABASE_URL', propValue: '', lastUpdated: new Date().toISOString() },
    { propKey: 'SUPABASE_ANON_KEY', propValue: '', lastUpdated: new Date().toISOString() },
    { propKey: 'DRIVE_FOLDER_ID', propValue: '1A2b3C_uny_drive_logbooks_root', lastUpdated: new Date().toISOString() },
    { propKey: 'TEMPLATE_M_DOCS_ID', propValue: '1A3Zp7vV-c2_35I3zYy7nJ_4qB1W-0y1YvT-k1wQ0XQc', lastUpdated: new Date().toISOString() },
    { propKey: 'TEMPLATE_M_SLIDES_ID', propValue: '1xwOSVbBKbH7M4RPT1QNRAYNM-yvuS8f9JG8LDJigZPI', lastUpdated: new Date().toISOString() }
  ]);
  const [driveIdInput, setDriveIdInput] = useState('1A2b3C_uny_drive_logbooks_root');
  const [docTemplateInput, setDocTemplateInput] = useState('1A3Zp7vV-c2_35I3zYy7nJ_4qB1W-0y1YvT-k1wQ0XQc');
  const [portfolioTemplateInput, setPortfolioTemplateInput] = useState('');
  const [slideTemplateInput, setSlideTemplateInput] = useState('1xwOSVbBKbH7M4RPT1QNRAYNM-yvuS8f9JG8LDJigZPI');

  useEffect(() => {
    // Clear any stale localStorage data if Supabase is configured (Supabase is source of truth)
    if (hasSupabase) {
      localStorage.removeItem('uny_sim_users');
      localStorage.removeItem('uny_sim_master_tasks');
      localStorage.removeItem('uny_sim_tasks');
      localStorage.removeItem('uny_sim_logbooks');
      localStorage.removeItem('uny_sim_jobdesks');
    }
    
    // Initial Pull from Relational Tables
    db.fetchAll().then(data => {
      if(data) {
        let updatedUsers = data.users as User[];
          if (data.nomorSuratData) {
            const nomData = typeof data.nomorSuratData === 'string' ? JSON.parse(data.nomorSuratData) : data.nomorSuratData;
            updatedUsers = updatedUsers.map(u => ({
              ...u,
              nomorSurat: nomData[u.nim] || u.nomorSurat || ''
            }));
          }
        setUsers(updatedUsers);
        setMasterTasks(data.masterTasks as MasterTask[]);
        
        if (data.tasks.length || data.masterTasks.length) {
          const mTasks = data.masterTasks as MasterTask[];
          const repairedTasks = (data.tasks as Task[]).map(t => {
            const master = mTasks.find(m => m.id === t.masterTaskId || (t.taskName || '').toLowerCase().trim() === (m.title || '').toLowerCase().trim());
            if (master && master.points) {
              const rawPoints = (t as any).points; // Supabase column 'points'
              let oldChecked: boolean[] = [];
              let oldDates: string[] = [];
              if (Array.isArray(rawPoints)) {
                oldChecked = rawPoints;
              } else if (rawPoints && typeof rawPoints === 'object') {
                oldChecked = rawPoints.checked || [];
                oldDates = rawPoints.dates || [];
              } else {
                oldChecked = t.pointsChecked || [];
                oldDates = t.checkDates || [];
              }
              const finalPoints = master.points;
              const newChecked = new Array(finalPoints.length).fill(false);
              const newDates = new Array(finalPoints.length).fill('');
              // Preserve existing checkmarks up to the new length
              for (let i = 0; i < Math.min(oldChecked.length, finalPoints.length); i++) {
                newChecked[i] = oldChecked[i];
                newDates[i] = oldDates[i] || '';
              }
              return {
                ...t,
                masterTaskId: master.id, // Fix legacy missing IDs
                taskName: master.title,
                category: master.category,
                description: master.description,
                workType: master.workType,
                points: finalPoints,
                pointsChecked: newChecked,
                checkDates: newDates
              };
            }
            return t;
          });
          setTasks(repairedTasks);
        } else {
          setTasks(data.tasks as Task[]);
        }
        
        setLogbooks(data.logbooks as Logbook[]);
        setJobdesks(data.jobdesks);
        // Baca kategori dari app_state (lebih reliable dari tabel categories)
        if (data.categoriesData && Array.isArray(data.categoriesData) && data.categoriesData.length > 0) {
          setCategories(data.categoriesData);
        } else if (data.categories && data.categories.length > 0) {
          // Fallback: dari tabel categories lama
          setCategories(data.categories);
        }
        
        if(data.properties) {
          const parsedProps = typeof data.properties === 'string' ? JSON.parse(data.properties) : data.properties;
          setPropertiesData(parsedProps);
          
          let fetchedDrive = parsedProps.find((p: any) => p.propKey === 'DRIVE_FOLDER_ID')?.propValue;
          let fetchedDoc = parsedProps.find((p: any) => p.propKey === 'TEMPLATE_M_DOCS_ID')?.propValue;
          let fetchedPortfolio = parsedProps.find((p: any) => p.propKey === 'TEMPLATE_PORTFOLIO_ID')?.propValue;
          let fetchedSlide = parsedProps.find((p: any) => p.propKey === 'TEMPLATE_M_SLIDES_ID')?.propValue;

          if (fetchedDoc === '1Doc_master_portfolio_template_uny') fetchedDoc = '1A3Zp7vV-c2_35I3zYy7nJ_4qB1W-0y1YvT-k1wQ0XQc';
          if (fetchedSlide === '1Slide_master_certificate_template_uny') fetchedSlide = '1xwOSVbBKbH7M4RPT1QNRAYNM-yvuS8f9JG8LDJigZPI';
          
          setDriveIdInput(fetchedDrive || '');
          setDocTemplateInput(fetchedDoc || '');
          setPortfolioTemplateInput(fetchedPortfolio || '');
          setSlideTemplateInput(fetchedSlide || '');
        }
      }
      setIsDataLoaded(true);
    }).catch(e => {
        setIsDataLoaded(true);
        console.error("Gagal menarik data awal dari Supabase:", e);
    });
  }, []);

  // Background Auto-Push to Supabase when state changes
  useEffect(() => {
    if (!isDataLoaded) return;
    const t = setTimeout(async () => {
      setIsSyncingSession(true);
      try {
        if (users.length) await db.runMutation('users', 'upsert', users.map(u => ({
          nim: u.nim, name: u.name, email: u.email, role: u.role, password: u.password, 
          periode: parseInt(u.periode as any) || 3, tanggalmulai: u.tanggalMulai, tanggalselesai: u.tanggalSelesai, nomorsurat: u.nomorSurat || ''
        })));
        if (Object.keys(jobdesks).length) await db.runMutation('jobdesks', 'upsert', Object.entries(jobdesks).map(([roleName, description]) => ({
          rolename: roleName, description
        })));
        if (masterTasks.length) await db.runMutation('master_tasks', 'upsert', masterTasks.map(m => ({
          id: m.id, title: m.title, category: m.category, description: m.description, points: m.points, worktype: m.workType, targetrole: m.targetRole
        })));
        if (tasks.length) await db.runMutation('tasks', 'upsert', tasks.map(t => {
          const pointsLen = t.points?.length || 0;
          const checkedArr = t.pointsChecked || [];
          const checkedCount = checkedArr.filter(Boolean).length;
          const progressPct = pointsLen > 0 ? Math.round((checkedCount / pointsLen) * 100) : 0;
          return {
            id: t.taskId, masterid: t.masterTaskId || null, assignednim: t.assignedNim, taskname: t.taskName, category: t.category,
            dateassigned: (t as any).dateAssigned || new Date().toISOString(), status: t.status, progress: progressPct, completeddesc: (t as any).completedDesc || '-', 
            completeddate: (t as any).completedDate || '-', googledocurl: t.googleDocUrl, points: { checked: t.pointsChecked || [], dates: t.checkDates || [] }
          };
        }));
        if (logbooks.length) await db.runMutation('logbooks', 'upsert', logbooks.map(l => ({
          logbookid: l.logbookId, taskid: l.taskId, nim: l.nim, taskname: l.taskName, category: l.category, date: l.timestamp, 
          workdescription: l.workDescription, hoursspent: l.hoursSpent || 0, grade: l.grade, gradenote: l.notes, googledocurl: l.googleDocUrl
        })));
        if (categories.length) {
          // Simpan kategori ke app_state (lebih reliable dari tabel categories)
          await db.runMutation('app_state', 'upsert', { id: 'categoriesData', data: categories });
          // Juga coba sync ke tabel categories (best effort)
          try { await db.runMutation('categories', 'upsert', categories.map(name => ({name}))); } catch(_) {}
        }
        await db.runMutation('app_state', 'upsert', { id: 'propertiesData', data: propertiesData });
        
        if (users.length) {
          const nData: Record<string, string> = {};
          users.forEach(u => { nData[u.nim] = u.nomorSurat || ''; });
          await db.runMutation('app_state', 'upsert', { id: 'nomorSuratData', data: nData });
        }
      } catch (e) {
        console.error("Auto Sync Error:", e);
      } finally {
        setIsSyncingSession(false);
      }
    }, 2000);
    return () => clearTimeout(t);
  }, [users, tasks, masterTasks, logbooks, jobdesks, categories, propertiesData, isDataLoaded]);

  const pullFromDatabase = async () => {
    // Deprecated via UI, tapi tetap digunakan kalau terpaksa
    customAlert("Aplikasi sekarang tersinkronisasi otomatis (Real-time) dengan tabel relasional Supabase. Tombol ini tidak diperlukan lagi.", "info");
  };

  const pushToDatabase = async () => {
    // Deprecated
    customAlert("Aplikasi menyinkronkan data secara otomatis ke Supabase di latar belakang setiap ada perubahan.", "success");
  };

  // Forms & Modal states
  const [authTab, setAuthTab] = useState<'login' | 'register'>('login');
  const [loginNIM, setLoginNIM] = useState('');
  const [loginPass, setLoginPass] = useState('');

  // Register inputs
  const [regNIM, setRegNIM] = useState('');
  const [regName, setRegName] = useState('');
  const [regEmail, setRegEmail] = useState('');
  const [regPass, setRegPass] = useState('');
  const [regRole, setRegRole] = useState<string>('Anggota');
  const [regPeriode, setRegPeriode] = useState('3');
  const [regTanggalMulai, setRegTanggalMulai] = useState('');
  const [regTanggalSelesai, setRegTanggalSelesai] = useState('');

  // PIC: Master Task Creation Inputs
  const [picTaskTitle, setPicTaskTitle] = useState('');
  const [picTaskCategory, setPicTaskCategory] = useState<string>('Jaringan');
  const [picTaskWorkType, setPicTaskWorkType] = useState<'Individu' | 'Kelompok'>('Individu');
  const [picTaskTargetRole, setPicTaskTargetRole] = useState<string>('Semua Peran');
  const [picTaskDesc, setPicTaskDesc] = useState('');
  const [picChecklistItems, setPicChecklistItems] = useState<string[]>([
    'Melakukan observasi & plotting port switch lab',
    'Pemasangan/instalasi kabel & RJ45 sesuai standar EIA/TIA',
    'Konfigurasi access list & routing di switch layer 3',
    'Melakukan ping test & draf google doc laporan'
  ]);
  const [newCheckItemText, setNewCheckItemText] = useState('');
  const [editingCheckItemIdx, setEditingCheckItemIdx] = useState<number | null>(null);
  const [editingCheckItemText, setEditingCheckItemText] = useState('');
  const [editingMasterTaskId, setEditingMasterTaskId] = useState<string | null>(null);

  // For pre-determined Google Doc Template editor modal
  const [selectedTaskToDuplicateDoc, setSelectedTaskToDuplicateDoc] = useState<Task | null>(null);
  const [newTimelineDate, setNewTimelineDate] = useState('');
  const [newTimelineDesc, setNewTimelineDesc] = useState('');
  const [newTimelineHours, setNewTimelineHours] = useState('');

  // Dynamic Role & Jobdesk Management Inputs
  const [newRoleName, setNewRoleName] = useState('');
  const [selectedRoleForJobdesk, setSelectedRoleForJobdesk] = useState('Anggota');


  const [certTagName, setCertTagName] = useState('{{NAMA}}');
  const [certTagNim, setCertTagNim] = useState('{{NIM}}');
  const [certTagGrade, setCertTagGrade] = useState('{{PREDIKAT}}');
  const [certTagCertNum, setCertTagCertNum] = useState('{{NOMOR}}');
  const [certTagRole, setCertTagRole] = useState('{{PERAN}}');
  const [certTagPeriode, setCertTagPeriode] = useState('{{PERIODE}}');
  const [certTagMulai, setCertTagMulai] = useState('{{MULAI}}');
  const [certTagSelesai, setCertTagSelesai] = useState('{{SELESAI}}');
  const [certDesignTheme, setCertDesignTheme] = useState<'amber_gold' | 'sapphire_blue' | 'emerald_classic' | 'royal_purple'>('sapphire_blue');
  const [certDesignTitle, setCertDesignTitle] = useState('SERTIFIKAT PENGHARGAAN');
  const [certDesignInstitution, setCertDesignInstitution] = useState('FAKULTAS VOKASI');
  const [certCustomSigneeName, setCertCustomSigneeName] = useState('Prof. Dr. Komarudin, S.Pd., M.A.');
  const [certCustomSigneeTitle, setCertCustomSigneeTitle] = useState('Dekan');
  const [certCustomLogoText, setCertCustomLogoText] = useState('UNIVERSITAS NEGERI YOGYAKARTA');

  // Ketua: createTask inputs (delegator selects from PIC defined masterTasks)
  const [taskFormMasterId, setTaskFormMasterId] = useState('');
  const [taskFormAssigned, setTaskFormAssigned] = useState('');
  const [ketuaActiveSection, setKetuaActiveSection] = useState<'delegasi' | 'penugasan'>('delegasi');

  // Member Action Modal (Completing Task with Simulated Google Doc Template)
  const [selectedTaskToComplete, setSelectedTaskToComplete] = useState<Task | null>(null);
  const [completeDesc, setCompleteDesc] = useState('');
  const [completeFileName, setCompleteFileName] = useState('');
  const [completeFileUrl, setCompleteFileUrl] = useState('');
  
  // Simulated Google Doc Editor inputs (saved inside logbooks/tasks)
  const [docReportTitle, setDocReportTitle] = useState('');
  const [docReportOverview, setDocReportOverview] = useState('');
  const [docReportSteps, setDocReportSteps] = useState('');
  const [docReportChallenges, setDocReportChallenges] = useState('');
  const [docReportConclusion, setDocReportConclusion] = useState('');

  // PIC Action Modal (Grading Logbook)
  const [selectedLogbookToGrade, setSelectedLogbookToGrade] = useState<Logbook | null>(null);
  const [gradeInput, setGradeInput] = useState('');
  const [gradeNotes, setGradeNotes] = useState('');

  // User Credential/Data Editing Modal
  const [editingUser, setEditingUser] = useState<User | null>(null);
  const [editUserNim, setEditUserNim] = useState('');
  const [editUserName, setEditUserName] = useState('');
  const [editUserEmail, setEditUserEmail] = useState('');
  const [editUserPassword, setEditUserPassword] = useState('');
  const [editUserPeriode, setEditUserPeriode] = useState('3');
  const [editUserTanggalMulai, setEditUserTanggalMulai] = useState('');
  const [editUserTanggalSelesai, setEditUserTanggalSelesai] = useState('');
  const [editUserNomorSurat, setEditUserNomorSurat] = useState('');

  // Generated docs visualization Modal
  const [printDocument, setPrintDocument] = useState<{
    studentName: string;
    studentNim: string;
    overallGrade: string;
    email: string;
    certId: string;
    portfolioId: string;
    periode?: number;
    tanggalMulai?: string;
    tanggalSelesai?: string;
    nomorSurat?: string;
    role?: string;
    logs: { 
      name: string; 
      category: string; 
      wordDesc: string; 
      file: string; 
      date: string;
      grade?: string;
      workType?: string;
      googleDocUrl?: string;
      googleDocTitle?: string;
    }[];
  } | null>(null);

  // Loading animation simulation state
  const [isSimulatingCall, setIsSimulatingCall] = useState(false);
  const [simulationMsg, setSimulationMsg] = useState('');

  const triggerCallSimulation = (msg: string, action: () => void) => {
    setIsSimulatingCall(true);
    setSimulationMsg(msg);
    setTimeout(() => {
      setIsSimulatingCall(false);
      action();
    }, 1200);
  };

  const getStudentGradesInfo = (studentNim: string) => {
    const studentLogs = logbooks.filter(l => l.nim === studentNim && l.grade !== '');
    if (studentLogs.length === 0) {
      return { average: 0, predicate: 'Belum Dinilai', numGraded: 0 };
    }
    const sum = studentLogs.reduce((acc, curr) => acc + parseFloat(curr.grade || '0'), 0);
    const average = Math.round((sum / studentLogs.length) * 10) / 10;
    
    let predicate = 'Perlu Ditingkatkan';
    if (average >= 95) {
      predicate = 'Sangat Baik';
    } else if (average >= 85) {
      predicate = 'Baik';
    } else if (average >= 80) {
      predicate = 'Cukup';
    }
    return { average, predicate, numGraded: studentLogs.length };
  };

  const formatIdDate = (dateStr?: string) => {
    if (!dateStr) return '20 Juni 2026';
    const parts = dateStr.split('-');
    if (parts.length !== 3) return dateStr;
    const months = [
      'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
      'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
    ];
    const day = parseInt(parts[2], 10);
    const monthIdx = parseInt(parts[1], 10) - 1;
    const year = parts[0];
    if (monthIdx >= 0 && monthIdx < 12) {
      return `${day} ${months[monthIdx]} ${year}`;
    }
    return dateStr;
  };

  const getCategoryPercentages = (studentNim: string) => {
    // Ambil persentase kategori berdasarkan jumlah Task yang berstatus Completed
    const studentTasks = tasks.filter(t => t.assignedNim === studentNim && t.status === 'Completed');
    const total = studentTasks.length;
    if (total === 0) return [];

    const counts: Record<string, number> = {};
    studentTasks.forEach(t => {
      const cat = t.category || 'Lainnya';
      counts[cat] = (counts[cat] || 0) + 1;
    });

    const list = Object.keys(counts).map(cat => ({
      category: cat,
      count: counts[cat],
      percentage: Math.round((counts[cat] / total) * 100)
    })).sort((a, b) => b.count - a.count);

    // Normalize percentages to sum to exactly 100%
    const sum = list.reduce((acc, curr) => acc + curr.percentage, 0);
    if (sum > 0 && sum !== 100 && list.length > 0) {
      list[0].percentage += (100 - sum);
    }
    return list;
  };

  const handleSimulatedLogin = (e: React.FormEvent) => {
    e.preventDefault();
    if (!loginNIM || !loginPass) {
      alert("Mohon masukkan NIS/NIP dan password Anda.");
      return;
    }
    
    triggerCallSimulation("Memvalidasi kredensial pengguna di spreadsheet...", () => {
      const found = users.find(u => u.nim === loginNIM);
      if (found) {
        if (found.password === loginPass) {
          setActiveUser(found);
          setLoginPass('');
        } else {
          customAlert(`Gagal Masuk: Password yang Anda masukkan untuk NIM [${loginNIM}] tidak cocok. Silakan coba lagi!`, 'error', 'Akses Ditolak');
        }
      } else {
        customAlert("NIM tidak terdaftar di database. Gunakan tab Daftar untuk mendaftarkan akun baru.", 'warning', 'Akses Ditolak');
      }
    });
  };

  const handleSimulatedRegister = (e: React.FormEvent) => {
    e.preventDefault();
    if (!regNIM || !regName || !regEmail || !regPass) {
      customAlert("Mohon lengkapi seluruh formulir registrasi dasar.", 'warning');
      return;
    }
    
    if (!regTanggalMulai || !regTanggalSelesai) {
      customAlert("Mohon lengkapi Tanggal Mulai dan Tanggal Selesai untuk Mahasiswa Magang.", 'warning');
      return;
    }

    const exists = users.some(u => u.nim === regNIM);
    if (exists) {
      customAlert("NIM/NIS sudah terdaftar didalam sistem!", 'error');
      return;
    }

    triggerCallSimulation("Menyimpan baris registrasi baru didalam database...", () => {
      const newUser: User = {
        nim: regNIM,
        name: regName,
        email: regEmail,
        role: 'Anggota',
        password: regPass,
        periode: regPeriode ? parseInt(regPeriode, 10) : 3,
        tanggalMulai: regTanggalMulai,
        tanggalSelesai: regTanggalSelesai
      };
      setUsers([...users, newUser]);
      customAlert(`Akun ${regName} berhasil terdaftar sebagai Peran 'Anggota'. Silakan login.`, 'success');
      setAuthTab('login');
      setLoginNIM(regNIM);
      // reset registration
      setRegNIM('');
      setRegName('');
      setRegEmail('');
      setRegPass('');
      setRegRole('Anggota');
      setRegPeriode('3');
      setRegTanggalMulai('');
      setRegTanggalSelesai('');
    });
  };

  const handleSimulatedRoleChange = (nim: string, newRole: string) => {
    triggerCallSimulation(`Memperbarui sel peran NIM ${nim} ke '${newRole}'...`, () => {
      setUsers(users.map(u => u.nim === nim ? { ...u, role: newRole } : u));
    });
  };

  const handleSaveEditedUser = (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingUser) return;

    const oldNim = editingUser.nim;
    const newNim = editUserNim.trim();
    const newName = editUserName.trim();
    const newEmail = editUserEmail.trim();
    const newPass = editUserPassword.trim();

    if (!newNim || !newName || !newEmail) {
      alert("Mohon isi NIM/NIP, Nama, dan Email dengan lengkap.");
      return;
    }

    // Check if new NIM/NIP already exists for another user
    if (newNim !== oldNim && users.some(u => u.nim === newNim)) {
      alert(`Galat: NIM/NIP/NIS ${newNim} sudah terdaftar untuk pengguna lain!`);
      return;
    }

    triggerCallSimulation(`Menyinkronkan data pengguna ${newName} ke Google Sheets...`, () => {
      if (oldNim !== newNim) {
        db.runMutation('users', 'delete', null, { column: 'nim', value: oldNim }).catch(() => {});
      }
      
      // 1. Update in the users list
      const updatedUsers = users.map(u => {
        if (u.nim === oldNim) {
          return {
            ...u,
            nim: newNim,
            name: newName,
            email: newEmail,
            password: newPass,
            periode: editUserPeriode ? parseInt(editUserPeriode, 10) : 3,
            tanggalMulai: editUserTanggalMulai,
            tanggalSelesai: editUserTanggalSelesai,
            nomorSurat: editUserNomorSurat.trim()
          };
        }
        return u;
      });
      setUsers(updatedUsers);

      // 2. Cascade changes to Tasks and Logbooks to prevent references being broken
      if (oldNim !== newNim) {
        setTasks(prev => prev.map(t => t.assignedNim === oldNim ? { ...t, assignedNim: newNim } : t));
        setLogbooks(prev => prev.map(l => l.nim === oldNim ? { 
          ...l, 
          nim: newNim, 
          studentName: newName, 
          studentEmail: newEmail 
        } : l));
      } else {
        // Just names & emails in logbooks if NIM remains same
        setLogbooks(prev => prev.map(l => l.nim === oldNim ? { 
          ...l, 
          studentName: newName, 
          studentEmail: newEmail 
        } : l));
      }

      // 3. Keep current session active in case they edited themselves
      if (activeUser?.nim === oldNim) {
        const matching = updatedUsers.find(u => u.nim === newNim);
        if (matching) {
          setActiveUser(matching);
        }
      }

      setEditingUser(null);
      alert(`Sukses: Akun "${newName}" telah diperbarui dan diselaraskan di database Google Sheets!`);
    });
  };

  const [userToDeleteConfirm, setUserToDeleteConfirm] = useState<string | null>(null);

  const handleDeleteUser = (nim: string) => {
    if (nim === activeUser?.nim) {
      Swal.fire({
        icon: 'error',
        title: 'Aksi Ditolak',
        text: 'Anda tidak bisa menghapus akun Anda sendiri!'
      });
      return;
    }
    const targetUser = users.find(u => u.nim === nim);
    if (!targetUser) return;

    const childTasks = tasks.filter(t => t.assignedNim === nim);
    const childTaskIds = childTasks.map(t => t.taskId);
    const childLogbooks = logbooks.filter(l => l.nim === nim || childTaskIds.includes(l.taskId));

    Swal.fire({
      title: 'Hapus Mahasiswa?',
      html: `Apakah Anda yakin ingin menghapus data Mahasiswa <b>"${targetUser.name}"</b> (NIM: ${nim})?<br/><br/>
             Sistem akan menghapus seluruh data terkait (ON DELETE CASCADE) di database Supabase:<br/>
             • <b>${childTasks.length}</b> Riwayat penugasan/pemberian instruksi kerja<br/>
             • <b>${childLogbooks.length}</b> Catatan logbook aktivitas harian`,
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
      cancelButtonColor: '#3085d6',
      confirmButtonText: 'Ya, Hapus Semua!',
      cancelButtonText: 'Batal'
    }).then((result) => {
      if (result.isConfirmed) {
        triggerCallSimulation(`[DATABASE CASCADE EXECUTION] Mengakses Supabase: Menghapus data login user, membersihkan ${childLogbooks.length} baris log di tabel 'logbooks', dan menghapus ${childTasks.length} delegasi penugasan di tabel 'tasks'...`, () => {
          // Update local state cascade
          setLogbooks(prev => prev.filter(l => l.nim !== nim && !childTaskIds.includes(l.taskId)));
          setTasks(prev => prev.filter(t => t.assignedNim !== nim));
          setUsers(users.filter(u => u.nim !== nim));
          setUserToDeleteConfirm(null);
          
          // Direct delete calls to Supabase (manually cascade-delete children first to satisfy constraints)
          const pLogbooks = db.runMutation('logbooks', 'delete', null, { column: 'nim', value: nim }).catch(err => console.warn(err));
          const pTasks = db.runMutation('tasks', 'delete', null, { column: 'assignednim', value: nim }).catch(err => console.warn(err));
          
          Promise.all([pLogbooks, pTasks]).then(() => {
            db.runMutation('users', 'delete', null, { column: 'nim', value: nim })
              .then(() => {
                Swal.fire({
                  icon: 'success',
                  title: 'Berhasil Dihapus',
                  text: `User "${targetUser.name}" beserta penugasan dan logbook berhasil dibersihkan total dari database Supabase!`
                });
              })
              .catch(err => {
                console.error("Gagal menghapus pengguna dari Supabase:", err);
                Swal.fire({
                  icon: 'error',
                  title: 'Ups, Ada Masalah',
                  text: `Gagal menghapus user dari database server: ${err.message || err}`
                });
              });
          });
        });
      }
    });
  };

  const handleDeleteTask = (taskId: string) => {
    const taskToDelete = tasks.find(t => t.taskId === taskId);
    if (!taskToDelete) return;

    const childLogbooks = logbooks.filter(l => l.taskId === taskId || ((l.taskName || '').toLowerCase().trim() === (taskToDelete.taskName || '').toLowerCase().trim() && l.nim === taskToDelete.assignedNim));
    const assigneeName = users.find(u => u.nim === taskToDelete.assignedNim)?.name || taskToDelete.assignedNim;

    Swal.fire({
      title: 'Hapus Penugasan?',
      html: `Apakah Anda yakin ingin menghapus/membatalkan penugasan <b>"${taskToDelete.taskName}"</b> untuk <b>"${assigneeName}"</b>?<br/><br/>
             Tindakan ini juga akan otomatis menghapus <b>${childLogbooks.length}</b> catatan logbook harian terkait secara permanen di database Supabase Anda.`,
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
      cancelButtonColor: '#3085d6',
      confirmButtonText: 'Ya, Hapus!',
      cancelButtonText: 'Batal'
    }).then((result) => {
      if (result.isConfirmed) {
        triggerCallSimulation(`[DATABASE CASCADE EXECUTION] Menghapus penugasan siswa ID ${taskId} dan membersihkan ${childLogbooks.length} logbook terkait di database Supabase...`, () => {
          setLogbooks(prev => prev.filter(l => l.taskId !== taskId && !((l.taskName || '').toLowerCase().trim() === (taskToDelete.taskName || '').toLowerCase().trim() && l.nim === taskToDelete.assignedNim)));
          setTasks(prev => prev.filter(t => t.taskId !== taskId));

          // Delete children first, then parent in Supabase
          db.runMutation('logbooks', 'delete', null, { column: 'taskid', value: taskId })
            .catch(err => console.warn(err))
            .then(() => {
              db.runMutation('tasks', 'delete', null, { column: 'id', value: taskId })
                .then(() => {
                  Swal.fire({
                    icon: 'success',
                    title: 'Penugasan Dihapus',
                    text: `Sukses: Penugasan siswa beserta ${childLogbooks.length} logbook terkait berhasil dihapus total dari database Supabase.`
                  });
                })
                .catch(err => {
                  console.error("Gagal menghapus penugasan dari Supabase:", err);
                  Swal.fire({
                    icon: 'error',
                    title: 'Gagal Menghapus',
                    text: `Gagal menghapus penugasan dari server: ${err.message || err}`
                  });
                });
            });
        });
      }
    });
  };

  const handleDeleteLogbook = (logbookId: string, studentName: string, taskName: string) => {
    Swal.fire({
      title: 'Hapus Logbook?',
      html: `Apakah Anda yakin ingin menghapus catatan logbook dari <b>"${studentName}"</b> untuk tugas <b>"${taskName}"</b>?<br/><br/>
             Tindakan ini akan menghapus baris logbook harian dari database Supabase agar bisa dilaporkan ulang oleh siswa, namun tidak akan menghapus berkas Google Docs/Drive asli hasil kerja demi keselamatan arsip.`,
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
      cancelButtonColor: '#3085d6',
      confirmButtonText: 'Ya, Hapus Catatan!',
      cancelButtonText: 'Batal'
    }).then((result) => {
      if (result.isConfirmed) {
        triggerCallSimulation(`Menghapus logbook ID: ${logbookId} dari tabel logbooks di database Supabase...`, () => {
          setLogbooks(prev => prev.filter(l => l.logbookId !== logbookId));

          // Direct delete call to Supabase
          db.runMutation('logbooks', 'delete', null, { column: 'logbookid', value: logbookId })
            .then(() => {
              Swal.fire({
                icon: 'success',
                title: 'Logbook Dihapus',
                text: 'Sukses: Catatan logbook berhasil dihapus dari database Supabase.'
              });
            })
            .catch(err => {
              console.error("Gagal menghapus logbook dari Supabase:", err);
              Swal.fire({
                icon: 'error',
                title: 'Gagal Menghapus',
                text: `Gagal menghapus logbook dari server: ${err.message || err}`
              });
            });
        });
      }
    });
  };

  const handleResetPicTaskForm = () => {
    setPicTaskTitle('');
    setPicTaskDesc('');
    setPicChecklistItems([
      'Melakukan observasi & plotting port switch lab',
      'Pemasangan/instalasi kabel & RJ45 sesuai standar EIA/TIA',
      'Konfigurasi access list & routing di switch layer 3',
      'Melakukan ping test & draf google doc laporan'
    ]);
    setPicTaskCategory('Jaringan');
    setPicTaskWorkType('Individu');
    setPicTaskTargetRole('Semua Peran');
    setEditingMasterTaskId(null);
  };

  const handleSelectMasterTaskToEdit = (m: MasterTask) => {
    setEditingMasterTaskId(m.id);
    setPicTaskTitle(m.title);
    setPicTaskCategory(m.category);
    setPicTaskWorkType(m.workType);
    setPicTaskTargetRole(m.targetRole || 'Semua Peran');
    setPicTaskDesc(m.description);
    setPicChecklistItems(m.points || []);
  };

  const handlePicDeleteMasterTask = (id: string, title: string) => {
    const master = masterTasks.find(m => m.id === id);
    if (!master) return;

    // Filter child tasks and logbooks to delete
    const childTasks = tasks.filter(t => (t.taskName || '').toLowerCase().trim() === (master.title || '').toLowerCase().trim() || (t.masterTaskId && t.masterTaskId === master.id));
    const childTaskIds = childTasks.map(t => t.taskId);
    const childLogbooks = logbooks.filter(l => childTaskIds.includes(l.taskId) || (l.taskName || '').toLowerCase().trim() === (master.title || '').toLowerCase().trim() || (l.masterTaskId && l.masterTaskId === master.id));

    Swal.fire({
      title: 'Hapus Tugas Pokok?',
      html: `Apakah Anda yakin ingin menghapus tugas pokok <b>"${master.title}"</b> dari Bank Tugas?<br/><br/>
             Sistem mendeteksi data anak (Cascade) yang harus dibersihkan untuk mencegah galat constraints (Foreign Key) di database Supabase:<br/>
             • <b>${childTasks.length}</b> Tugas Siswa aktif yang mendelegasikan materi ini<br/>
             • <b>${childLogbooks.length}</b> Catatan Logbook harian / hasil pengerjaan terkait`,
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
      cancelButtonColor: '#3085d6',
      confirmButtonText: 'Ya, Hapus Semua!',
      cancelButtonText: 'Batal'
    }).then((result) => {
      if (result.isConfirmed) {
        triggerCallSimulation(`[DATABASE CASCADE EXECUTION] Mengakses Supabase: Menghapus data Logbook, membersihkan tugas delegasi, dan menghapus baris induk ID '${id}' di tabel 'master_tasks' (Bank Tugas Pokok)...`, () => {
          // Cascade delete from states
          setLogbooks(prev => prev.filter(l => !childTaskIds.includes(l.taskId) && (l.taskName || '').toLowerCase().trim() !== (master.title || '').toLowerCase().trim() && (!l.masterTaskId || l.masterTaskId !== master.id)));
          setTasks(prev => prev.filter(t => (t.taskName || '').toLowerCase().trim() !== (master.title || '').toLowerCase().trim() && (!t.masterTaskId || t.masterTaskId !== master.id)));
          setMasterTasks(prev => prev.filter(m => m.id !== id));
          
          if (editingMasterTaskId === id) {
            handleResetPicTaskForm();
          }

          // Direct cascade deletes in Supabase manually
          const pLogbooks = Promise.all(childTaskIds.map(tid => db.runMutation('logbooks', 'delete', null, { column: 'taskid', value: tid }))).catch(e => console.warn(e));
          const pTasks = db.runMutation('tasks', 'delete', null, { column: 'masterid', value: id }).catch(e => console.warn(e));

          Promise.all([pLogbooks, pTasks]).then(() => {
            // Now safe to delete the parent master task
            db.runMutation('master_tasks', 'delete', null, { column: 'id', value: id })
              .then(() => {
                Swal.fire({
                  icon: 'success',
                  title: 'Tugas Pokok Dihapus',
                  text: `Sukses: Tugas pokok "${master.title}" beserta seluruh data anak terkait (${childTasks.length} penugasan, ${childLogbooks.length} logbook) berhasil dibersihkan total dari database Supabase!`
                });
              })
              .catch(err => {
                console.error("Gagal menghapus tugas pokok dari Supabase:", err);
                Swal.fire({
                  icon: 'error',
                  title: 'Gagal Menghapus',
                  text: `Gagal menghapus tugas pokok dari server: ${err.message || err}`
                });
              });
          });
        });
      }
    });
  };

  const handlePicCreateMasterTask = (e: React.FormEvent) => {
    e.preventDefault();
    if (!picTaskTitle || !picTaskDesc) {
      alert("Mohon isi judul dan deskripsi tugas standar.");
      return;
    }

    const actionText = editingMasterTaskId 
      ? `Mengubah / memperbarui tugas pokok "${picTaskTitle}" di Bank Tugas...` 
      : "Menambah standardisasi tugas pokok baru dari PIC...";

    triggerCallSimulation(actionText, () => {
      const finalPoints = picChecklistItems.length > 0 
        ? picChecklistItems.map(p => p.trim()).filter(Boolean)
        : [
            `Analisis kebutuhan dasar & topologi "${picTaskTitle}"`,
            `Melakukan konfigurasi / pengerjaan teknis spesifik`,
            `Melakukan pengetesan jalur & verifikasi ketersambungan`,
            `Pembuatan dokumentasi & draf Google Doc laporan`
          ];

      if (editingMasterTaskId) {
        const oldMaster = masterTasks.find(m => m.id === editingMasterTaskId);
        const oldTitle = oldMaster ? (oldMaster.title || '') : '';

        setMasterTasks(masterTasks.map(m => m.id === editingMasterTaskId ? {
          ...m,
          title: picTaskTitle,
          category: picTaskCategory,
          description: picTaskDesc,
          workType: picTaskWorkType,
          targetRole: picTaskTargetRole,
          points: finalPoints
        } : m));

        // [AUTO-SYNC] Automatically update all delegated tasks that originated from this Master Task
        setTasks(tasks.map(t => {
          const isMatch = t.masterTaskId === editingMasterTaskId || 
                          (t.taskName || '').toLowerCase().trim() === oldTitle.toLowerCase().trim();
          
          if (isMatch) {
            const oldChecked = t.pointsChecked || [];
            const newChecked = new Array(finalPoints.length).fill(false);
            // Preserve existing checkmarks up to the new length to avoid destroying student's work
            for (let i = 0; i < Math.min(oldChecked.length, finalPoints.length); i++) {
              newChecked[i] = oldChecked[i];
            }
            return {
              ...t,
              masterTaskId: editingMasterTaskId, // Ensure legacy tasks get the ID properly bound now
              taskName: picTaskTitle,
              category: picTaskCategory,
              description: picTaskDesc,
              workType: picTaskWorkType,
              points: finalPoints,
              pointsChecked: newChecked,
              checkDates: new Array(finalPoints.length).fill('')
            };
          }
          return t;
        }));

        alert(`Sukses memperbarui tugas pokok '${picTaskTitle}' dan menyinkronkannya ke semua tugas mahasiswa yang sedang berjalan!`);
        handleResetPicTaskForm();
      } else {
        const newMaster: MasterTask = {
          id: 'MT-' + Date.now(),
          title: picTaskTitle,
          category: picTaskCategory,
          description: picTaskDesc,
          workType: picTaskWorkType,
          targetRole: picTaskTargetRole,
          points: finalPoints
        };
        setMasterTasks([newMaster, ...masterTasks]);
        setPicTaskTitle('');
        setPicTaskDesc('');
        setPicChecklistItems([
          'Melakukan observasi & plotting port switch lab',
          'Pemasangan/instalasi kabel & RJ45 sesuai standar EIA/TIA',
          'Konfigurasi access list & routing di switch layer 3',
          'Melakukan ping test & draf google doc laporan'
        ]);
        setPicTaskTargetRole('Semua Peran');
        alert(`Sukses: Tugas pokok '${picTaskTitle}' (${picTaskWorkType}) terdefinisi khusus untuk: ${picTaskTargetRole} dengan ${finalPoints.length} butir checklist!`);
      }
    });
  };

  const handleSimulatedCreateTask = (e: React.FormEvent) => {
    e.preventDefault();
    if (!taskFormMasterId || !taskFormAssigned) {
      alert("Mohon pilih tugas pokok dari PIC dan mendelegasikan ke petugas magang.");
      return;
    }

    const master = masterTasks.find(m => m.id === taskFormMasterId);
    if (!master) {
      alert("Tugas pokok tidak valid.");
      return;
    }

    triggerCallSimulation("Ketua mendelegasikan materi tugas pokok dari PIC ke tim...", () => {
      const copiedPoints = master.points && master.points.length > 0
        ? master.points
        : [
            `Analisis kebutuhan dasar & topologi "${master.title}"`,
            `Melakukan konfigurasi / pengerjaan teknis spesifik`,
            `Melakukan pengetesan jalur & verifikasi ketersambungan`,
            `Pembuatan dokumentasi & draf Google Doc laporan`
          ];

      const newTask: Task = {
        taskId: 'TSK-' + Date.now(),
        masterTaskId: master.id,
        taskName: master.title,
        category: master.category,
        description: master.description,
        assignedNim: taskFormAssigned,
        status: 'Pending',
        createdBy: '19600101', // Selalu PIC yang membuat standardisasinya
        workType: master.workType,
        points: copiedPoints,
        pointsChecked: new Array(copiedPoints.length).fill(false),
        checkDates: new Array(copiedPoints.length).fill(''),
        timelineLogs: [],
        docContent: {
          judul: `Laporan Penugasan ${master.title}`,
          instansi: 'Fakultas Vokasi Universitas Negeri Yogyakarta',
          pembahasan: '',
          tantangan: '',
          statusAkhir: 'Progres Berlapis'
        }
      };
      setTasks([newTask, ...tasks]);
      setTaskFormMasterId('');
      setTaskFormAssigned('');
      alert(`Sukses: Tugas pokok '${master.title}' berhasil didelegasikan dan checklist progres bertahap sudah diaktifkan!`);
    });
  };

  const handleSimulatedCompleteTask = (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedTaskToComplete) return;
    if (!completeDesc || !docReportTitle || !docReportOverview) {
      alert("Mohon lengkapi deskripsi pengerjaan dan isi file Google Doc Laporan.");
      return;
    }

    if (!docTemplateInput) {
      alert("Template Doc ID belum diisi. Hubungi PIC atau isi di pengaturan.");
      return;
    }

    triggerCallSimulation(`Menghubungkan ke Google Apps Script: Menyalin template [ID Doc: ${docTemplateInput}] dari PIC, menulis isian laporan ke file Google Doc baru, menyimpan file di Folder Drive [ID Folder: ${driveIdInput}], dan menyinkronkan data...`, () => {
      const generatedDocId = `${docTemplateInput}_report_${selectedTaskToComplete.taskId}_${activeUser?.nim || 'student'}`;
      const generatedDocUrl = `https://docs.google.com/document/d/${generatedDocId}/edit`;
      
      // 1. Update task status
      setTasks(tasks.map(t => t.taskId === selectedTaskToComplete.taskId ? { 
        ...t, 
        status: 'Completed',
        googleDocId: generatedDocId,
        googleDocUrl: generatedDocUrl,
        googleDocTitle: docReportTitle
      } : t));
      
      // 2. Add logbook entry
      const newLog: Logbook = {
        logbookId: 'LOG-' + Date.now(),
        masterTaskId: selectedTaskToComplete.masterTaskId,
        nim: activeUser?.nim || '2153214003',
        studentName: activeUser?.name || 'Budi Hartono',
        studentEmail: activeUser?.email || 'budi.hartono@student.uny.ac.id',
        taskId: selectedTaskToComplete.taskId,
        taskName: selectedTaskToComplete.taskName,
        category: selectedTaskToComplete.category,
        timestamp: new Date().toISOString().replace('T', ' ').substring(0, 19),
        workDescription: completeDesc,
        fileName: completeFileName || 'bukti_dokumentasi_teknis.png',
        fileUrl: completeFileUrl || 'https://images.unsplash.com/photo-1544383835-bda2bc66a55d?q=80&w=600',
        grade: '',
        notes: '',
        workType: selectedTaskToComplete.workType || 'Individu',
        googleDocId: generatedDocId,
        googleDocUrl: generatedDocUrl,
        googleDocTitle: docReportTitle,
        docContent: {
          judul: docReportTitle,
          instansi: 'Fakultas Vokasi Universitas Negeri Yogyakarta',
          pembahasan: docReportOverview + '\n\nLangkah pengerjaan:\n' + docReportSteps,
          tantangan: docReportChallenges,
          statusAkhir: docReportConclusion || 'Selesai Berhasil'
        }
      };
      setLogbooks([newLog, ...logbooks]);
      
      // Reset Complete dialog states
      setSelectedTaskToComplete(null);
      setCompleteDesc('');
      setCompleteFileName('');
      setCompleteFileUrl('');
      
      // Reset Google Doc inputs
      setDocReportTitle('');
      setDocReportOverview('');
      setDocReportSteps('');
      setDocReportChallenges('');
      setDocReportConclusion('');

      alert(`Sukses: File laporan Google Doc khusus berhasil dibuat secara asinkron dari template PIC [ID Doc: ${docTemplateInput}] dan terekam di Logbooks Sheets!`);
    });
  };

  const handlePointToggle = (taskToUpdate: Task, pointIndex: number) => {
    const updatedChecked = [...(taskToUpdate.pointsChecked || [])];
    const updatedDates = [...(taskToUpdate.checkDates || [])];
    updatedChecked[pointIndex] = !updatedChecked[pointIndex];
    if (updatedChecked[pointIndex]) {
      updatedDates[pointIndex] = new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    } else {
      updatedDates[pointIndex] = '';
    }
    setTasks(tasks.map(t => t.taskId === taskToUpdate.taskId ? { ...t, pointsChecked: updatedChecked, checkDates: updatedDates } : t));
  };

  const handleAddTimelineRow = (taskId: string, date: string, desc: string, hrs: number) => {
    if (!date || !desc || !hrs) {
      alert("Mohon lengkapi Tanggal, Deskripsi Pekerjaan, dan Estimasi Jam.");
      return;
    }
    setTasks(tasks.map(t => {
      if (t.taskId === taskId) {
        const currentLogs = t.timelineLogs || [];
        return {
          ...t,
          timelineLogs: [...currentLogs, { date, description: desc, hours: Number(hrs) }]
        };
      }
      return t;
    }));
    setNewTimelineDate('');
    setNewTimelineDesc('');
    setNewTimelineHours('');
  };

  const handleRemoveTimelineRow = (taskId: string, indexToRemove: number) => {
    setTasks(tasks.map(t => {
      if (t.taskId === taskId) {
        const currentLogs = t.timelineLogs || [];
        return {
          ...t,
          timelineLogs: currentLogs.filter((_, idx) => idx !== indexToRemove)
        };
      }
      return t;
    }));
  };

  const handleDuplicateAndOpenDoc = async (taskToDup: Task) => {
    try {
      const gasUrl = import.meta.env.VITE_GAS_WEB_APP_URL;
      if (!gasUrl) {
        customAlert("Mohon maaf, API backend Google Apps Script belum terhubung! Silakan cek .env pada variabel VITE_GAS_WEB_APP_URL.", "error");
        return;
      }

      setIsSyncingSession(true);
      customAlert("Menduplikasi & Mengisi Template Laporan Google Docs (Memanggil Google Apps Script)...", 'info');

      const templateId = docTemplateInput || '1A3Zp7vV-c2_35I3zYy7nJ_4qB1W-0y1YvT-k1wQ0XQc'; // default placeholder if empty
      const destFolderId = driveIdInput || '';
      const docTitle = `Draf Laporan Hasil ${taskToDup.taskName} - ${activeUser?.name || 'Mahasiswa'}`;
      
      const payload = {
        studentData: {
          name: activeUser?.name || 'Mahasiswa',
          nim: activeUser?.nim || '',
          role: activeUser?.role || 'Anggota'
        },
        taskData: {
          taskName: taskToDup.taskName,
          category: taskToDup.category,
          workType: taskToDup.workType,
          points: (taskToDup.points || []).filter(p => typeof p === 'string' && p.trim() !== '').length > 0 
                  ? (taskToDup.points || []).filter(p => typeof p === 'string' && p.trim() !== '') 
                  : [
                      'Analisis kebutuhan dasar & topologi',
                      'Melakukan konfigurasi / pengerjaan teknis spesifik',
                      'Melakukan pengetesan jalur & ketersambungan',
                      'Pembuatan dokumentasi & draf Google Doc laporan'
                    ],
          checkDates: taskToDup.checkDates || []
        },
        driveId: destFolderId,
        docId: templateId
      };

      const res = await fetch(gasUrl, {
        method: 'POST',
        body: JSON.stringify({
          action: 'generate_logbook_doc',
          payload: payload
        })
      });

      const data = await res.json();

      if (!data.success) {
        throw new Error(data.message || 'Gagal menyalin dokumen melalui GAS.');
      }

      const generatedDocId = data.docId;
      const generatedDocUrl = data.docUrl;
      
      // Update local task with the new Google Doc link
      setTasks(tasks.map(t => t.taskId === taskToDup.taskId ? {
        ...t,
        googleDocId: generatedDocId,
        googleDocUrl: generatedDocUrl,
        googleDocTitle: docTitle
      } : t));

      customAlert(`Sukses Generate & Isi Logbook Master PIC!\n\nLink Dokumen siap dibuka pada tabel atau tombol akses dokumen.`, 'success');
      
    } catch (e: any) {
      console.error(e);
      customAlert(`Gagal menduplikasi Template API: ${e.message}\n\nMembuat tautan simulasi...`, 'warning');
      const generatedDocId = '1Doc_report_' + taskToDup.taskId + '_' + (activeUser?.nim || 'student') + '_' + Date.now().toString().slice(-4);
      const generatedDocUrl = `https://docs.google.com/document/d/${generatedDocId}/edit`;
      const docTitle = `Draf Laporan Hasil ${taskToDup.taskName} - ${activeUser?.name || 'Mahasiswa'}`;
      setTasks(tasks.map(t => t.taskId === taskToDup.taskId ? {
        ...t,
        googleDocId: generatedDocId,
        googleDocUrl: generatedDocUrl,
        googleDocTitle: docTitle
      } : t));
    } finally {
      setIsSyncingSession(false);
    }
  };

  const handleCompleteTaskDirect = (taskToComplete: Task) => {
    if (!taskToComplete.googleDocUrl) {
      alert("Silakan buat laporan Google Doc terlebih dahulu!");
      return;
    }
    const validPoints = (taskToComplete.points || []).filter(p => typeof p === 'string' && p.trim() !== '');
    const points = validPoints.length > 0 ? validPoints : [
      'Analisis kebutuhan dasar & topologi',
      'Melakukan konfigurasi / pengerjaan teknis spesifik',
      'Melakukan pengetesan jalur & verifikasi ketersambungan',
      'Pembuatan dokumentasi & draf Google Doc laporan'
    ];
    const checkedArr = taskToComplete.pointsChecked || new Array(points.length).fill(false);
    const checkedCount = checkedArr.filter(Boolean).length;
    const pct = Math.round((checkedCount / points.length) * 100);

    triggerCallSimulation(`Menyinkronkan rekap akhir tugas ke Sheets database Magang UNY dan memvalidasi URL Laporan Google Docs: ${taskToComplete.googleDocUrl}...`, () => {
      // Force 100% checked upon standard complete or leave checked as is
      const finalChecked = [...checkedArr];
      // Auto complete remaining check items of task just to be helpful, or keep as is. Let's force all checked to make it 100% complete
      const updatedChecked = new Array(points.length).fill(true);

      setTasks(tasks.map(t => t.taskId === taskToComplete.taskId ? {
        ...t,
        status: 'Completed',
        pointsChecked: updatedChecked
      } : t));

      // Append logbook entry so PIC can evaluate it in their dashboard
      const newLog: Logbook = {
        logbookId: 'LOG-' + Date.now(),
        masterTaskId: taskToComplete.masterTaskId,
        nim: activeUser?.nim || '2153214003',
        studentName: activeUser?.name || 'Mahasiswa Vokasi',
        studentEmail: activeUser?.email || 'mhs@student.uny.ac.id',
        taskId: taskToComplete.taskId,
        taskName: taskToComplete.taskName,
        category: taskToComplete.category,
        timestamp: new Date().toISOString().replace('T', ' ').substring(0, 19),
        workDescription: `Selesai Sempurna! Laporan resmi telah diisi langsung via Google Docs: ${taskToComplete.googleDocTitle || 'Dokumen Laporan'}. Progres 100% Beres.`,
        fileName: 'report_icon_drive.png',
        fileUrl: taskToComplete.googleDocUrl,
        grade: '',
        notes: '',
        workType: taskToComplete.workType || 'Individu',
        googleDocId: taskToComplete.googleDocId,
        googleDocUrl: taskToComplete.googleDocUrl,
        googleDocTitle: taskToComplete.googleDocTitle || 'Laporan Dokumen Magang',
        docContent: {
          judul: taskToComplete.googleDocTitle || 'Laporan Hasil Kerja',
          instansi: 'Fakultas Vokasi Universitas Negeri Yogyakarta',
          pembahasan: `Laporan diisikan lengkap langsung oleh siswa pada berkas Google Docs eksternal yang dihubungkan ke sistem. URL Berkas: ${taskToComplete.googleDocUrl}`,
          tantangan: 'Disisipkan langsung ke berkas Google DOCS di tab terpisah.',
          statusAkhir: 'Selesai 100%'
        }
      };
      setLogbooks([newLog, ...logbooks]);

      alert(`Sukses: Tugas '${taskToComplete.taskName}' resmi selesai disetor!\n\nTautan Google Doc Laporan & rekap logbook telah dikirimkan ke Dosen PIC di spreadsheet.`);
    });
  };

  const handleSimulatedGrading = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedLogbookToGrade) return;

    const updatedLogbook = { ...selectedLogbookToGrade, grade: gradeInput, notes: gradeNotes };
    
    // Update state lokal dulu
    setLogbooks(logbooks.map(l => l.logbookId === selectedLogbookToGrade.logbookId 
      ? updatedLogbook
      : l
    ));
    setSelectedLogbookToGrade(null);
    setGradeInput('');
    setGradeNotes('');

    // Langsung push nilai ke Supabase dengan upsert agar aman
    try {
      const l = updatedLogbook;
      await db.runMutation('logbooks', 'upsert', {
        logbookid: l.logbookId, taskid: l.taskId, nim: l.nim, taskname: l.taskName, category: l.category, date: l.timestamp, 
        workdescription: l.workDescription, hoursspent: l.hoursSpent || 0, grade: l.grade, gradenote: l.notes, googledocurl: l.googleDocUrl
      });
      Swal.fire({
        toast: true,
        position: 'top-end',
        icon: 'success',
        title: `Nilai ${gradeInput} berhasil disimpan ke database!`,
        showConfirmButton: false,
        timer: 3000
      });
    } catch (err: any) {
      console.error('[Supabase] Gagal simpan nilai:', err);
      Swal.fire({
        icon: 'error',
        title: 'Gagal Simpan Nilai',
        text: `Nilai sudah tersimpan di lokal tapi gagal sync ke server: ${err.message || err}`
      });
    }
  };

  const triggerPDFGenerationSimulation = (studentNim: string) => {
    const student = users.find(u => u.nim === studentNim);
    if (!student) return;

    triggerCallSimulation("Memproses integrasi Google Suite: Menyalin template Sertifikat Slides, menyusun tabel Portofolio Docs, dan mengirimkan lampiran PDF ke email...", () => {
      const studentLogsForPrint = logbooks.filter(l => l.nim === studentNim && l.grade !== '');
      
      let sum = 0;
      let count = 0;
      studentLogsForPrint.forEach(l => {
         const val = parseFloat(l.grade || '0');
         if (!isNaN(val)) { sum += val; count++; }
      });
      const calcGradeStr = count > 0 ? (sum / count).toFixed(1).replace(/\.0$/, '') : '0';

      const logs = studentLogsForPrint.map(l => {
        const relatedTask = tasks.find(t => t.taskId === l.taskId || t.taskName === l.taskName);
        return {
          name: l.taskName,
          category: relatedTask?.category || l.category || 'Lain-lain',
          wordDesc: l.workDescription,
          file: l.fileUrl,
          date: l.timestamp.split(' ')[0],
          grade: l.grade,
          workType: l.workType || 'Individu',
          googleDocUrl: l.googleDocUrl || `https://docs.google.com/document/d/1Doc_report_sim_${l.logbookId}/edit`,
          googleDocTitle: l.googleDocTitle || `Laporan Sim_Doc_${l.logbookId}`
        };
      });

      setPrintDocument({
        studentName: student.name,
        studentNim: student.nim,
        overallGrade: calcGradeStr,
        email: student.email,
        certId: 'cert_pdf_' + student.nim,
        portfolioId: 'portfolio_pdf_' + student.nim,
        logs: logs,
        periode: student.periode,
        tanggalMulai: student.tanggalMulai,
        tanggalSelesai: student.tanggalSelesai,
        nomorSurat: student.nomorSurat,
        role: student.role
      });
      setCertQuickEditNomorSurat(student.nomorSurat || '');
    });
  };

  const [certQuickEditNomorSurat, setCertQuickEditNomorSurat] = useState('');

  const handleSaveQuickNomorSurat = async () => {
    if (!printDocument) return;
    try {
      const uIndex = users.findIndex(u => u.nim === printDocument.studentNim);
      let updatedUsers = [...users];
      if (uIndex !== -1) {
        updatedUsers[uIndex] = { ...updatedUsers[uIndex], nomorSurat: certQuickEditNomorSurat };
        setUsers(updatedUsers);
        
        await db.runMutation('users', 'upsert', updatedUsers.map(u => ({
          nim: u.nim, name: u.name, email: u.email, role: u.role, password: u.password, 
          periode: parseInt(u.periode as any) || 3, tanggalmulai: u.tanggalMulai, tanggalselesai: u.tanggalSelesai, nomorsurat: u.nomorSurat || ''
        })));
        
        const nData: Record<string, string> = {};
        updatedUsers.forEach(u => { nData[u.nim] = u.nomorSurat || ''; });
        await db.runMutation('app_state', 'upsert', { id: 'nomorSuratData', data: nData });
      }
      
      setPrintDocument({...printDocument, nomorSurat: certQuickEditNomorSurat});
      Swal.fire({
        toast: true,
        position: 'top-end',
        icon: 'success',
        title: `Nomor Surat berhasil diperbarui!`,
        showConfirmButton: false,
        timer: 3000
      });
    } catch (e: any) {
      Swal.fire({
        toast: true,
        position: 'top-end',
        icon: 'error',
        title: `Gagal memperbarui nomor surat: ${e.message}`,
        showConfirmButton: false,
        timer: 4000
      });
    }
  };

  const copyToClipboard = (text: string, trigger: 'gs' | 'html') => {
    navigator.clipboard.writeText(text);
    if (trigger === 'gs') {
      setCopiedGs(true);
      setTimeout(() => setCopiedGs(false), 2000);
    } else {
      setCopiedHtml(true);
      setTimeout(() => setCopiedHtml(false), 2000);
    }
  };

  const resetAllSimulatorData = () => {
    if (window.confirm("Apakah Anda yakin ingin menyetel ulang data simulasi ke setelan awal?")) {
      localStorage.removeItem('uny_sim_users');
      localStorage.removeItem('uny_sim_tasks');
      localStorage.removeItem('uny_sim_logbooks');
      localStorage.removeItem('uny_sim_jobdesks');
      localStorage.removeItem('uny_sim_properties');
      window.location.reload();
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 flex flex-col font-sans antialiased selection:bg-amber-500 selection:text-slate-950">
      
      {/* 1. APP BENTO-STYLE HEADER HUB */}
      {!isProductionViewMode && (
        <>
          <header className="border-b border-slate-200 bg-white sticky top-0 z-40 px-6 py-4 flex flex-col md:flex-row md:items-center justify-between gap-4 shadow-[0_1px_3px_0_rgba(0,0,0,0.02)]">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-[#003a70] flex items-center justify-center text-white font-extrabold text-xl shadow-md shadow-blue-900/10">
                U
              </div>
              <div>
                <h1 className="text-lg font-bold text-slate-900 flex items-center gap-2">
                  UNY IT Internship Workspace <span className="text-[10px] bg-amber-100 text-amber-800 border border-amber-200 px-2 py-0.5 rounded-full font-mono font-bold">Apps Script 5.0</span>
                </h1>
                <p className="text-xs text-slate-500 font-medium">Departemen Jaringan Vokasi UNY • Developer Station & Simulator</p>
              </div>
            </div>
            
            {/* Core Bento Actions / Reset */}
            <div className="flex flex-wrap items-center gap-2">
              <button 
                onClick={resetAllSimulatorData}
                className="px-3.5 py-1.5 rounded-lg text-xs bg-rose-50 border border-rose-100 hover:bg-rose-100 text-rose-700 font-medium transition-all flex items-center gap-1.5 shadow-sm"
              >
                <RefreshCw className="w-3.5 h-3.5 animate-spin-hover" /> Reset Simulasi
              </button>
              <a
                href="/google-apps-script/Code.gs"
                download="Code.gs"
                className="px-3.5 py-1.5 rounded-lg text-xs bg-emerald-50 border border-emerald-100 hover:bg-emerald-100 text-emerald-700 font-medium transition-all flex items-center gap-1.5 shadow-sm"
              >
                <Download className="w-3.5 h-3.5" /> Code.gs
              </a>
              <a
                href="/google-apps-script/Index.html"
                download="Index.html"
                className="px-3.5 py-1.5 rounded-lg text-xs bg-blue-50 border border-blue-100 hover:bg-blue-100 text-blue-700 font-medium transition-all flex items-center gap-1.5 shadow-sm"
              >
                <Download className="w-3.5 h-3.5" /> Index.html
              </a>
            </div>
          </header>

          {/* 2. TAB CONTROLS NAVIGATION */}
          <nav className="flex items-center border-b border-slate-200 bg-white/70 backdrop-blur px-6 py-2.5 overflow-x-auto gap-2">
            <button
              onClick={() => setActiveTab('simulator')}
              className={`px-4 py-2 text-xs font-semibold rounded-lg transition-all flex items-center gap-2 shrink-0 ${activeTab === 'simulator' ? 'bg-[#003a70] text-white shadow-sm' : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'}`}
            >
              <Laptop className="w-4 h-4" /> 🔎 Simulator Aplikasi Live (Vokasi UI)
            </button>
            <button
              onClick={() => setActiveTab('code_gs')}
              className={`px-4 py-2 text-xs font-semibold rounded-lg transition-all flex items-center gap-2 shrink-0 ${activeTab === 'code_gs' ? 'bg-[#003a70] text-white shadow-sm' : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'}`}
            >
              <FileCode2 className="w-4 h-4" /> Code.gs (Backend)
            </button>
            <button
              onClick={() => setActiveTab('index_html')}
              className={`px-4 py-2 text-xs font-semibold rounded-lg transition-all flex items-center gap-2 shrink-0 ${activeTab === 'index_html' ? 'bg-[#003a70] text-white shadow-sm' : 'text-slate-600 hover:text-slate-900 hover:bg-slate-105'}`}
            >
              <FileCode className="w-4 h-4" /> Index.html (Frontend)
            </button>
            <button
              onClick={() => setActiveTab('db_schema')}
              className={`px-4 py-2 text-xs font-semibold rounded-lg transition-all flex items-center gap-2 shrink-0 ${activeTab === 'db_schema' ? 'bg-[#003a70] text-white shadow-sm' : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'}`}
            >
              <Database className="w-4 h-4" /> Supabase SQL Schema
            </button>
            <button
              onClick={() => setActiveTab('deploy_guide')}
              className={`px-4 py-2 text-xs font-semibold rounded-lg transition-all flex items-center gap-2 shrink-0 ${activeTab === 'deploy_guide' ? 'bg-[#003a70] text-white shadow-sm' : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'}`}
            >
              <BookOpen className="w-4 h-4" /> Panduan Deployment
            </button>
          </nav>
        </>
      )}

      {/* 3. CORE LAYOUT FRAME */}
      <main className="flex-1 w-full flex flex-col p-0 max-w-none">
        
        {/* ===================== TAB 1: INTERACTIVE APP SIMULATOR ===================== */}
        {activeTab === 'simulator' && (
          <div className={isProductionViewMode ? "flex-1 flex flex-col" : "space-y-6"}>
            {/* Top Workspace status bar */}
            {!isProductionViewMode && (
              <div className="bg-blue-50/70 border border-blue-100 rounded-2xl p-5 flex flex-col md:flex-row md:items-center justify-between gap-4 shadow-[0_4px_12px_rgba(0,0,0,0.01),0_1px_2px_rgba(0,0,0,0.01)]">
                <div className="flex items-start gap-3">
                  <Sparkles className="w-5 h-5 text-[#003a70] shrink-0 mt-0.5" />
                  <div>
                    <h3 className="font-bold text-[#003a70] text-sm">Simulator Interaktif Berfungsi Penuh</h3>
                    <p className="text-xs text-slate-600 leading-relaxed font-medium">
                      Jalankan skenario penuh program magang IT Network. Anda dapat berganti peran secara instan untuk mencoba integrasi Google Sheets, Drive, dan templating berkas kelulusan.
                    </p>
                  </div>
                </div>
                <div className="flex items-center gap-2.5 shrink-0 bg-white border border-slate-200/80 px-2 py-1 rounded-xl shadow-sm">
                  <span className="text-xs font-semibold text-slate-500">Akun Coba Instan:</span>
                  <div className="w-[200px]">
                    <Select 
                      options={[
                        { value: "", label: "-- Logout / Keluar Sesi --" },
                        ...users.map(u => ({ value: u.nim, label: `${u.name} (${u.role})` }))
                      ]}
                      value={
                        activeUser 
                        ? { value: activeUser.nim, label: `${activeUser.name} (${activeUser.role})` } 
                        : { value: "", label: "-- Logout / Keluar Sesi --" }
                      }
                      onChange={(option) => {
                        const val = option?.value || '';
                        if (val === '') {
                          setActiveUser(null);
                          setLoginNIM('');
                        } else {
                          const u = users.find(usr => usr.nim === val);
                          if (u) {
                            triggerCallSimulation(`Beralih peran secara cepat ke: ${u.name}...`, () => {
                               setActiveUser(u);
                            });
                          }
                        }
                      }}
                      styles={{
                        control: (base) => ({ ...base, minHeight: '30px', height: '30px', fontSize: '11px', border: 'none', boxShadow: 'none', background: 'transparent' }),
                        valueContainer: (base) => ({ ...base, padding: '0 4px' }),
                        input: (base) => ({ ...base, margin: '0 2px' }),
                        indicatorSeparator: () => ({ display: 'none' }),
                        dropdownIndicator: (base) => ({ ...base, padding: '2px' }),
                        menu: (base) => ({ ...base, fontSize: '11px', zIndex: 50 })
                      }}
                      isSearchable
                    />
                  </div>
                </div>
              </div>
            )}

            {/* Simulated browser wrapper / Real app wrapper */}
            <div className="flex-1 flex flex-col w-full bg-slate-50 border-y border-slate-200/80 shadow-sm min-h-[600px]">
              
              {/* Fake Chrome Address Bar */}
              {!isProductionViewMode && (
                <div className="bg-slate-200/50 border-b border-slate-300/40 px-4 py-3 flex items-center justify-between gap-4">
                  <div className="flex items-center gap-1.5 shrink-0">
                    <div className="w-3 h-3 rounded-full bg-rose-400" />
                    <div className="w-3 h-3 rounded-full bg-amber-400" />
                    <div className="w-3 h-3 rounded-full bg-emerald-400" />
                  </div>
                  
                  <div className="flex-1 max-w-2xl bg-white text-slate-500 text-xs py-1.5 px-4 rounded-xl font-mono flex items-center justify-between border border-slate-200/80">
                    <span className="truncate text-slate-400 selection:bg-blue-150">https://script.google.com/macros/s/AKfycbz_uny_vocational_it_network_app/exec</span>
                    <span className="text-emerald-600 text-[9px] bg-emerald-50 px-1.5 py-0.5 rounded font-extrabold tracking-wider border border-emerald-100">SECURE HTTPS</span>
                  </div>
                  
                  <div className="text-[11px] text-slate-600 font-bold bg-white px-2.5 py-1 rounded-lg border border-slate-200/80 shadow-sm shrink-0">
                    <span>UNY Dev Preview</span>
                  </div>
                </div>
              )}

              {/* SIMULATOR APPS CONTENT WINDOW */}
              <div className="min-h-[580px] bg-slate-50 text-slate-900 relative">
                
                {/* Simulated Loading Overlay */}
                <AnimatePresence>
                  {isSimulatingCall && (
                    <motion.div 
                      initial={{ opacity: 0 }}
                      animate={{ opacity: 1 }}
                      exit={{ opacity: 0 }}
                      className="fixed inset-0 bg-[#001529]/95 text-white flex flex-col items-center justify-center gap-4 z-[100] text-center px-6"
                    >
                      <div className="w-12 h-12 rounded-full border-4 border-amber-400 border-t-transparent animate-spin" />
                      <div className="text-sm font-mono mt-2 text-amber-300">{simulationMsg}</div>
                      <p className="text-[10px] text-slate-400">Mensimulasikan respons Google App Script dengan spreadsheet...</p>
                    </motion.div>
                  )}
                </AnimatePresence>

                {/* Simulated Web App Header Navbar (UNY Style) */}
                <header className="bg-gradient-to-r from-[#003a70] to-[#002244] text-white border-b-4 border-amber-400 px-6 py-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 shadow-md">
                  <div className="flex items-center gap-2">
                    <span className="bg-amber-400 text-slate-950 font-extrabold text-xs px-2.5 py-1 rounded tracking-wider">FV-UNY</span>
                    <h2 className="text-md font-bold tracking-tight">Portal Magang IT Network & Sistem Logbook</h2>
                  </div>
                  
                  {activeUser && (
                    <div className="flex items-center gap-3">
                      <div className="text-right d-none d-sm-block">
                        <div className="text-xs font-bold">{activeUser.name}</div>
                        <div className="text-[10px] text-slate-300">NIM/NIP: {activeUser.nim}</div>
                      </div>
                      <span className={`text-[10px] font-bold uppercase tracking-wider px-2 py-1 rounded-full ${
                        activeUser.role === 'PIC' ? 'bg-rose-600 text-white' :
                        activeUser.role === 'Ketua' ? 'bg-amber-400 text-slate-950' :
                        'bg-emerald-600 text-white'
                      }`}>
                        {activeUser.role}
                      </span>
                      <button 
                        onClick={() => {
                          triggerCallSimulation("Menutup sesi autentikasi SINKAD...", () => {
                            setActiveUser(null);
                          });
                        }} 
                        className="text-xs bg-[#001b33] hover:bg-[#002d54] text-red-400 hover:text-red-300 px-2.5 py-1.5 rounded-lg border border-red-900/20 font-mono transition-colors"
                      >
                        Logout
                      </button>
                    </div>
                  )}
                </header>

                {/* Interactive Portal Area */}
                <div className="p-6">
                  
                  {/* SIMULATED ROLE HERO BAR */}
                  {activeUser && (
                    <div className="bg-gradient-to-r from-[#003a70] to-[#002244] text-white p-5 rounded-2xl mb-6 shadow-sm border-b-4 border-amber-400">
                      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                        <div>
                          <h3 className="text-lg font-bold mb-0">Halo, {activeUser.name}!</h3>
                          <p className="text-xs text-sky-200 mb-0">Selamat berkontribusi di Departemen IT Network Vokasi UNY</p>
                        </div>
                        <div className="text-xs text-left sm:text-right bg-sky-950/50 p-2.5 rounded-xl border border-sky-800/30">
                          <div>NIM Aktif: <strong className="text-amber-300">{activeUser.nim}</strong></div>
                          <div>Sesi Aktif: <span className="text-emerald-400 font-mono text-[10px] font-bold">Google Sheets Sandbox</span></div>
                        </div>
                      </div>
                    </div>
                  )}

                  {/* ===================== SIMULATOR SUBVIEW: NOT LOGGED IN (AUTH) ===================== */}
                  {!activeUser && (
                    <div className="max-w-md mx-auto my-8 space-y-6">
                      <div className="bg-white p-8 rounded-2xl border border-slate-200/80 shadow-md">
                        <div className="text-center mb-6">
                          <span className="text-[#003a70] font-extrabold text-xl font-sans tracking-wide">FAKULTAS VOKASI UNY</span>
                          <div className="text-[10px] text-slate-400 uppercase tracking-widest mt-1 font-bold">Sistem Manajemen Kelulusan & Logbook</div>
                        </div>

                        {/* Tab toggles */}
                        <div className="flex border-b border-slate-200 mb-6">
                        <button
                          onClick={() => setAuthTab('login')}
                          className={`flex-1 pb-3 text-xs font-bold border-b-2 text-center transition-all ${authTab === 'login' ? 'border-[#003a70] text-[#003a70]' : 'border-transparent text-slate-400 hover:text-slate-600'}`}
                        >
                          Masuk Dengan Akun
                        </button>
                        <button
                          onClick={() => setAuthTab('register')}
                          className={`flex-1 pb-3 text-xs font-bold border-b-2 text-center transition-all ${authTab === 'register' ? 'border-[#003a70] text-[#003a70]' : 'border-transparent text-slate-400 hover:text-slate-600'}`}
                        >
                          Daftar Baru
                        </button>
                      </div>

                      {authTab === 'login' ? (
                        <form onSubmit={handleSimulatedLogin} className="space-y-4">
                          <div>
                            <label className="block text-xs font-bold text-slate-700 mb-1.5">Nomor Induk Siswa (NIS) / NIP</label>
                            <input 
                              type="text" 
                              className="w-full text-xs px-3.5 py-2 border border-slate-200 rounded-lg outline-none focus:border-[#003a70] focus:ring-1 focus:ring-[#003a70]/10 transition-all font-medium text-slate-800"
                              placeholder="Ketik NIS contoh: 1926"
                              value={loginNIM}
                              onChange={(e) => setLoginNIM(e.target.value)}
                            />
                          </div>
                          <div>
                            <label className="block text-xs font-bold text-slate-700 mb-1.5">Password</label>
                            <input 
                              type="password" 
                              className="w-full text-xs px-3.5 py-2 border border-slate-200 rounded-lg outline-none focus:border-[#003a70] focus:ring-1 focus:ring-[#003a70]/10 transition-all text-slate-800"
                              placeholder="Ketik password"
                              value={loginPass}
                              onChange={(e) => setLoginPass(e.target.value)}
                            />
                          </div>
                          <button 
                            type="submit"
                            className="w-full py-2.5 bg-[#003a70] hover:bg-[#002244] text-white rounded-lg text-xs font-bold tracking-wide shadow-md transition-all mt-2.5"
                          >
                            Masuk Portal
                          </button>
                        </form>
                      ) : (
                        <form onSubmit={handleSimulatedRegister} className="space-y-4">
                          <div>
                            <label className="block text-xs font-bold text-slate-700 mb-1.5">NIS Siswa</label>
                            <input 
                              type="text" 
                              className="w-full text-xs px-3.5 py-2 border border-slate-200 rounded-lg outline-none focus:border-[#003a70] font-mono font-medium text-slate-800"
                              placeholder="Masukkan NIS baru"
                              value={regNIM}
                              onChange={(e) => setRegNIM(e.target.value)}
                            />
                          </div>
                          <div>
                            <label className="block text-xs font-bold text-slate-700 mb-1.5">Nama Lengkap</label>
                            <input 
                              type="text" 
                              className="w-full text-xs px-3.5 py-2 border border-slate-200 rounded-lg outline-none focus:border-[#003a70] text-slate-800"
                              placeholder="Nama asli sesuai KTP"
                              value={regName}
                              onChange={(e) => setRegName(e.target.value)}
                            />
                          </div>
                          <div>
                            <label className="block text-xs font-bold text-slate-700 mb-1.5">Email Aktif</label>
                            <input 
                              type="email" 
                              className="w-full text-xs px-3.5 py-2 border border-slate-200 rounded-lg outline-none focus:border-[#003a70] text-slate-800"
                              placeholder="nama@student.uny.ac.id"
                              value={regEmail}
                              onChange={(e) => setRegEmail(e.target.value)}
                            />
                          </div>
                          <div>
                            <label className="block text-xs font-bold text-slate-700 mb-1.5">Password</label>
                            <input 
                              type="password" 
                              className="w-full text-xs px-3.5 py-2 border border-slate-200 rounded-lg outline-none focus:border-[#003a70] text-slate-800"
                              placeholder="Kata sandi minimal 6 karakter"
                              value={regPass}
                              onChange={(e) => setRegPass(e.target.value)}
                            />
                          </div>

                          <div>
                            <label className="block text-xs font-bold text-slate-700 mb-1.5">Periode Magang (Bulan)</label>
                            <input 
                              type="number" 
                              min="1"
                              max="12"
                              className="w-full text-xs px-3.5 py-2 border border-slate-200 rounded-lg outline-none focus:border-[#003a70] text-slate-800"
                              placeholder="Contoh: 3"
                              value={regPeriode}
                              onChange={(e) => setRegPeriode(e.target.value)}
                            />
                          </div>

                          <div className="grid grid-cols-2 gap-3">
                            <div>
                              <label className="block text-xs font-bold text-slate-700 mb-1.5">Tanggal Mulai</label>
                              <input 
                                type="date" 
                                className="w-full text-xs px-3.5 py-2 border border-[#cbd5e1] rounded-lg outline-none focus:border-[#003a70] text-slate-800 font-mono font-medium"
                                value={regTanggalMulai}
                                onChange={(e) => setRegTanggalMulai(e.target.value)}
                              />
                            </div>
                            <div>
                              <label className="block text-xs font-bold text-slate-700 mb-1.5">Tanggal Selesai</label>
                              <input 
                                type="date" 
                                className="w-full text-xs px-3.5 py-2 border border-[#cbd5e1] rounded-lg outline-none focus:border-[#003a70] text-slate-800 font-mono font-medium"
                                value={regTanggalSelesai}
                                onChange={(e) => setRegTanggalSelesai(e.target.value)}
                              />
                            </div>
                          </div>

                          <button 
                            type="submit"
                            className="w-full py-2.5 bg-[#003a70] hover:bg-[#002244] text-white rounded-lg text-xs font-bold tracking-wide shadow-md transition-all mt-2.5"
                          >
                            Daftarkan Sekarang
                          </button>
                        </form>
                      )}
                      
                      {/* Supabase Sync Config */}
                      <div className="border-t border-slate-150 mt-6 pt-5">
                        <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 text-center">
                          <h5 className="text-[11px] font-bold text-[#003a70] uppercase tracking-wider mb-2">
                            Sinkronisasi Supabase Aktif
                          </h5>
                          <div className="bg-emerald-50 border border-emerald-200 p-3 rounded text-emerald-800 text-[11px] text-center font-medium flex items-center justify-center gap-2">
                            <span className="relative flex h-3 w-3">
                              <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                              <span className="relative inline-flex rounded-full h-3 w-3 bg-emerald-500"></span>
                            </span>
                            Sistem secara otomatis melakukan sinkronisasi data secara real-time. Tombol Push/Pull tidak lagi diperlukan.
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                )}

                  {/* ===================== SIMULATOR SUBVIEW: ANGGOTA (MEMBER WORKSPACE) ===================== */}
                  {activeUser && activeUser.role !== 'PIC' && activeUser.role !== 'Ketua' && (
                    <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
                      
                      {/* Left Block: Core Jobdesks assigned by PIC */}
                      <div className="lg:col-span-4 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex flex-col">
                        <div className="pb-3 border-b border-slate-100 mb-4">
                          <h4 className="font-extrabold text-[#003a70] text-xs uppercase tracking-wider">Peran & Tanggung Jawab</h4>
                        </div>
                        <div className="bg-blue-50/50 border-l-4 border-[#003a70] p-4 rounded-r-xl mb-4">
                          <span className="text-[10px] font-bold bg-[#003a70] text-white px-2 py-0.5 rounded uppercase tracking-wider mb-2 inline-block">
                            {activeUser.role}
                          </span>
                          <p className="text-xs text-[#002d54] font-medium leading-relaxed whitespace-pre-line mt-1.5">
                            {jobdesks[activeUser.role] || "Tidak ada deskripsi spesifik."}
                          </p>
                        </div>
                        <div className="mt-auto pt-4 border-t border-slate-100 text-[11px] text-slate-500 space-y-2">
                          <div className="font-bold text-slate-700">Petunjuk Logbook Harian:</div>
                          <div className="text-slate-600">1. Klik tombol 'Selesaikan' pada tugas aktif.</div>
                          <div className="text-slate-600">2. Jabarkan penyelesaian teknis dengan detail.</div>
                          <div className="text-slate-600">3. Upload tangkapan layar perangkat / bukti.</div>
                        </div>
                      </div>

                      {/* Right Block: Tasks Assigned */}
                      <div className="lg:col-span-8 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
                        <div className="pb-3 border-b border-slate-100 mb-4 flex justify-between items-center">
                          <h4 className="font-extrabold text-[#003a70] text-xs uppercase tracking-wider">Penugasan Untuk Anda</h4>
                          <span className="text-xs bg-slate-100 text-slate-600 px-3 py-1 rounded-lg font-bold">
                            Total: {tasks.filter(t => t.assignedNim === activeUser.nim).length} Penugasan
                          </span>
                        </div>

                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                          {tasks.filter(t => t.assignedNim === activeUser.nim).length === 0 ? (
                            <div className="col-span-2 py-12 text-center text-slate-400 text-xs font-semibold">
                              Belum ada delegasi tugas dari Ketua / Koordinator Jaringan untuk Anda.
                            </div>
                          ) : (
                            tasks.filter(t => t.assignedNim === activeUser.nim).map(task => {
                              const isCompleted = task.status === 'Completed';
                              const validPoints = (task.points || []).filter(p => typeof p === 'string' && p.trim() !== '');
                              const points = validPoints.length > 0 ? validPoints : [
                                'Analisis kebutuhan dasar & topologi',
                                'Melakukan konfigurasi / pengerjaan teknis spesifik',
                                'Melakukan pengetesan jalur & ketersambungan',
                                'Pembuatan dokumentasi & draf Google Doc laporan'
                              ];
                              const checkedArr = task.pointsChecked || new Array(points.length).fill(false);
                              const checkedCount = checkedArr.filter(Boolean).length;
                              const pct = Math.round((checkedCount / points.length) * 100);

                              return (
                                <div key={task.taskId} className={`p-5 border rounded-3xl flex flex-col justify-between transition-all duration-300 ${isCompleted ? 'border-emerald-200 bg-emerald-50/25 shadow-sm' : 'border-blue-105 bg-[#f7fafe]/70 hover:bg-white shadow-sm'}`}>
                                  <div>
                                    <div className="flex items-center justify-between mb-2.5">
                                      <span className={`text-[9.5px] font-bold px-2 py-0.5 rounded ${
                                        task.category === 'Jaringan' ? 'bg-sky-100 text-sky-800' :
                                        task.category === 'Website' ? 'bg-emerald-100 text-emerald-800' :
                                        'bg-slate-100 text-slate-800'
                                      }`}>
                                        {task.category}
                                      </span>
                                      <span className={`text-[10px] font-bold ${isCompleted ? 'text-emerald-700' : 'text-amber-700'}`}>
                                        {task.status === 'Completed' ? 'Completed' : `In Progress / Pending (${pct}%)`}
                                      </span>
                                    </div>
                                    <h5 className="font-extrabold text-slate-850 text-sm mb-1.5 leading-snug">{task.taskName}</h5>
                                    <p className="text-slate-500 text-xs leading-relaxed mb-4 font-semibold">{task.description}</p>

                                    {/* Milestones Checkpoints Section */}
                                    <div className="mt-4 pt-3 border-t border-slate-100/90 space-y-2">
                                      <div className="flex items-center justify-between text-[10.5px] font-bold text-slate-700">
                                        <span className="text-slate-800">Capaian Progres Lapangan</span>
                                        <span className="text-[#003a70]">{checkedCount} / {points.length} ({pct}%)</span>
                                      </div>
                                      <div className="w-full bg-slate-100 h-2 rounded-full overflow-hidden shadow-inner">
                                        <div 
                                          className={`h-full transition-all duration-300 ${pct === 100 ? 'bg-emerald-600' : 'bg-[#003a70]'}`}
                                          style={{ width: `${pct}%` }}
                                        />
                                      </div>

                                      <div className="space-y-1.5 mt-2.5">
                                        {points.map((pt, index) => (
                                          <label 
                                            key={index} 
                                            className={`flex items-start gap-2.5 p-2 rounded-xl border text-[11px] leading-relaxed cursor-pointer transition-colors ${
                                              checkedArr[index] 
                                                ? 'bg-emerald-50/45 border-emerald-100 text-emerald-900 font-bold' 
                                                : 'bg-white hover:bg-slate-50 border-slate-100 text-slate-650'
                                            }`}
                                          >
                                            <input 
                                              type="checkbox"
                                              className="mt-0.5 rounded border-slate-300 text-[#003a70] focus:ring-[#003a70] disabled:opacity-50"
                                              checked={!!checkedArr[index]}
                                              disabled={isCompleted}
                                              onChange={() => handlePointToggle(task, index)}
                                            />
                                            <span className={checkedArr[index] ? 'line-through text-slate-400' : ''}>
                                              {pt}
                                            </span>
                                          </label>
                                        ))}
                                      </div>
                                    </div>

                                    {/* Display Registered Daily Timelines */}
                                    {task.timelineLogs && task.timelineLogs.length > 0 && (
                                      <div className="mt-3.5 bg-blue-50/30 border border-blue-100/50 p-3 rounded-2xl text-[10.5px] space-y-1.5 text-slate-655">
                                        <div className="font-bold text-[#003a70] flex items-center justify-between uppercase tracking-wider text-[9px] border-b pb-1 border-blue-100/40">
                                          <span>Timeline Kejadian Harian</span>
                                          <span className="font-mono bg-blue-100/60 px-1.5 py-0.5 rounded text-blue-900">{task.timelineLogs.length} Kegiatan</span>
                                        </div>
                                        <div className="max-h-[90px] overflow-y-auto space-y-1 pr-1">
                                          {task.timelineLogs.map((log, idx) => (
                                            <div key={idx} className="flex justify-between items-start border-b border-blue-50/40 pb-1.5 last:border-0 last:pb-0 gap-1.5">
                                              <span className="text-slate-800 font-extrabold shrink-0">{log.date}:</span>
                                              <span className="text-left w-full text-slate-500 font-medium leading-relaxed">{log.description}</span>
                                              <span className="font-mono text-slate-400 shrink-0">({log.hours}j)</span>
                                            </div>
                                          ))}
                                        </div>
                                      </div>
                                    )}

                                  </div>

                                  <div className="border-t border-slate-100 pt-3.5 flex flex-col gap-2 mt-4 text-[11px]">
                                    <div className="flex justify-between items-center text-[10px] text-slate-400 font-mono">
                                      <span>Dari NIM: {task.createdBy}</span>
                                      {task.workType && (
                                        <span className={`font-sans font-bold px-2 py-0.5 rounded text-[8.5px] ${task.workType === 'Kelompok' ? 'bg-purple-100 text-purple-800' : 'bg-emerald-100 text-emerald-800'}`}>
                                          {task.workType}
                                        </span>
                                      )}
                                    </div>

                                    {isCompleted ? (
                                      <div className="space-y-1.5 pt-1">
                                        <div className="text-emerald-800 font-bold flex items-center gap-1.5 bg-emerald-50 p-2 rounded-xl border border-emerald-100/70 text-[11px] justify-center">
                                          <span>✓ Selesai 100% & Terverifikasi</span>
                                        </div>
                                        {task.googleDocUrl && (
                                          <a 
                                            href={task.googleDocUrl} 
                                            target="_blank" 
                                            rel="noreferrer"
                                            className="block w-full py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-center text-xs font-bold transition-all shadow-sm"
                                          >
                                            📄 Buka Laporan Google Doc
                                          </a>
                                        )}
                                      </div>
                                    ) : (
                                      <div className="space-y-2 pt-1.5 w-full">
                                        {!task.googleDocUrl ? (
                                          <div className="flex gap-2 w-full">
                                            <button 
                                              type="button"
                                              onClick={async () => {
                                                try {
                                                  await db.runMutation('tasks', 'upsert', {
                                                    id: task.taskId,
                                                    masterid: task.masterTaskId || null,
                                                    assignednim: task.assignedNim,
                                                    taskname: task.taskName,
                                                    category: task.category,
                                                    dateassigned: (task as any).dateAssigned || new Date().toISOString(),
                                                    status: task.status,
                                                    progress: pct,
                                                    points: { checked: task.pointsChecked || [], dates: task.checkDates || [] },
                                                    googledocurl: task.googleDocUrl
                                                  });
                                                  Swal.fire({ toast: true, position: 'top-end', icon: 'success', title: `✅ Progres ${pct}% berhasil disimpan ke database!`, showConfirmButton: false, timer: 2500 });
                                                } catch(err: any) {
                                                  Swal.fire({ icon: 'error', title: 'Gagal Simpan', text: err.message || 'Tidak bisa terhubung ke Supabase' });
                                                }
                                              }}
                                              className="flex-1 py-2 border border-slate-200 hover:bg-slate-50 text-slate-700 rounded-xl transition-all font-bold text-center text-xs"
                                            >
                                              💾 Simpan Harian
                                            </button>
                                            <button 
                                              type="button"
                                              onClick={() => handleDuplicateAndOpenDoc(task)}
                                              className="flex-1 py-2 bg-[#003a70] hover:bg-[#002244] text-white rounded-xl transition-all font-bold text-center text-xs flex items-center justify-center gap-1"
                                            >
                                              📄 Buat Laporan Doc
                                            </button>
                                          </div>
                                        ) : (
                                          <div className="space-y-2">
                                            <a 
                                              href={task.googleDocUrl}
                                              target="_blank"
                                              rel="noreferrer"
                                              className="block w-full py-2 bg-[#f0f6ff] border border-blue-200 text-[#003a70] hover:bg-blue-100 text-center rounded-xl text-xs font-bold transition-all shadow-sm flex items-center justify-center gap-1.5"
                                            >
                                              📄 Buka Google Doc (Tab Baru) ↗
                                            </a>
                                            <div className="flex gap-2">
                                              <button 
                                                type="button"
                                                onClick={async () => {
                                                  try {
                                                    await db.runMutation('tasks', 'upsert', {
                                                      id: task.taskId,
                                                      masterid: task.masterTaskId || null,
                                                      assignednim: task.assignedNim,
                                                      taskname: task.taskName,
                                                      category: task.category,
                                                      dateassigned: (task as any).dateAssigned || new Date().toISOString(),
                                                      status: task.status,
                                                      progress: pct,
                                                      points: { checked: task.pointsChecked || [], dates: task.checkDates || [] },
                                                      googledocurl: task.googleDocUrl
                                                    });
                                                    Swal.fire({ toast: true, position: 'top-end', icon: 'success', title: `✅ Progres ${pct}% berhasil disimpan ke database!`, showConfirmButton: false, timer: 2500 });
                                                  } catch(err: any) {
                                                    Swal.fire({ icon: 'error', title: 'Gagal Simpan', text: err.message || 'Tidak bisa terhubung ke Supabase' });
                                                  }
                                                }}
                                                className="flex-1 py-2 border border-slate-200 hover:bg-slate-50 text-slate-700 rounded-xl transition-all font-bold text-center text-[10px]"
                                              >
                                                💾 Simpan Harian
                                              </button>
                                              <button 
                                                type="button"
                                                onClick={() => handleCompleteTaskDirect(task)}
                                                className="flex-1 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl transition-all font-bold text-center text-[10px] flex items-center justify-center gap-0.5"
                                              >
                                                ✓ Selesaikan & Setor
                                              </button>
                                            </div>
                                          </div>
                                        )}
                                      </div>
                                    )}
                                  </div>
                                </div>
                              );
                            })
                          )}
                        </div>

                        {/* Recent logbook status tracker */}
                        <div className="mt-8 border-t border-slate-100 pt-6">
                          <h5 className="font-bold text-xs text-[#003a70] uppercase tracking-wider mb-4">Aktivitas Logbook Terkini Anda</h5>
                          <div className="overflow-x-auto">
                            <table className="w-full text-left text-xs text-slate-600">
                              <thead>
                                <tr className="border-b border-slate-200 text-slate-400 text-[10px] uppercase font-bold tracking-wider">
                                  <th className="pb-2.5 font-bold">Tugas</th>
                                  <th className="pb-2.5 font-bold">Tanggal</th>
                                  <th className="pb-2.5 font-bold">Bukti</th>
                                  <th className="pb-2.5 font-bold text-right">Penilaian Dosen</th>
                                </tr>
                              </thead>
                              <tbody>
                                {logbooks.filter(l => l.nim === activeUser.nim).length === 0 ? (
                                  <tr>
                                    <td colSpan={4} className="py-4 text-center text-slate-400 italic">Belum ada logbook teknis yang tersimpan di Sheets.</td>
                                  </tr>
                                ) : (
                                  logbooks.filter(l => l.nim === activeUser.nim).map(log => (
                                    <tr key={log.logbookId} className="border-b border-slate-100 text-[11px] hover:bg-slate-50/50">
                                      <td className="py-3 font-semibold text-slate-800">{log.taskName}</td>
                                      <td className="py-3 text-slate-400">{log.timestamp.split(' ')[0]}</td>
                                      <td className="py-3">
                                        <a href={log.googleDocUrl || log.fileUrl || '#'} target="_blank" rel="noreferrer" className="text-[#003a70] underline flex items-center gap-1 hover:text-[#002244] font-bold">
                                          Tinjau Berkas
                                        </a>
                                      </td>
                                      <td className="py-3 text-right font-bold text-emerald-600">
                                        {log.grade ? `${log.grade} / 100` : <span className="text-slate-400 font-normal italic">Proses Audit</span>}
                                      </td>
                                    </tr>
                                  ))
                                )}
                              </tbody>
                            </table>
                          </div>
                        </div>

                      </div>
                    </div>
                  )}
                         {/* ===================== SIMULATOR SUBVIEW: KETUA DASHBOARD ===================== */}
                  {activeUser && activeUser.role === 'Ketua' && (
                    <div className="space-y-8 animate-fade-in">
                      
                     

                      {/* SECTION 1: Personal Tasks & Logbook (Penugasan Mandiri Saya) */}
                      <div className="space-y-4">
                        <div className="flex items-center justify-between border-b border-slate-200 pb-2">
                          <div className="flex items-center gap-2">
                            <span className="text-lg">📋</span>
                            <h4 className="font-extrabold text-[#003a70] text-xs uppercase tracking-wider mb-0">I. Penugasan & Logbook Mandiri Saya ({activeUser.name})</h4>
                          </div>
                          <span className="text-[10px] bg-amber-50 text-amber-800 font-extrabold px-2.5 py-1 rounded-lg border border-amber-200/50">
                            Status Akun: Aktif Berkolaborasi
                          </span>
                        </div>

                        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
                          
                          {/* Left Block: Core Jobdesks for Ketua */}
                          <div className="lg:col-span-4 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex flex-col">
                            <div className="pb-3 border-b border-slate-100 mb-4">
                              <h4 className="font-extrabold text-[#003a70] text-xs uppercase tracking-wider">Peran & Tanggung Jawab</h4>
                            </div>
                            <div className="bg-amber-50 border-l-4 border-amber-400 p-4 rounded-r-xl mb-4">
                              <span className="text-[10px] font-bold bg-[#003a70] text-white px-2 py-0.5 rounded uppercase tracking-wider mb-2 inline-block">
                                {activeUser.role} (Koordinator Utama)
                              </span>
                              <p className="text-xs text-[#002d54] font-medium leading-relaxed whitespace-pre-line mt-1.5">
                                {jobdesks[activeUser.role] || "Memimpin departemen magang, mendelegasikan instruksi harian, melacak progres pencapaian kompetensi tim, serta ikut melakukan penugasan dan pelaporan logbook teknis secara aktif."}
                              </p>
                            </div>
                            <div className="mt-auto pt-4 border-t border-slate-100 text-[11px] text-slate-500 space-y-2">
                              <div className="font-bold text-slate-700">Petunjuk Logbook Harian Ketua:</div>
                              <div className="text-slate-600">1. Klik tombol 'Selesaikan' pada tugas aktif Anda.</div>
                              <div className="text-slate-600">2. Jabarkan penyelesaian teknis dengan detail.</div>
                              <div className="text-slate-600">3. Upload tangkapan layar perangkat / bukti.</div>
                            </div>
                          </div>

                          {/* Right Block: Tasks Assigned to Ketua */}
                          <div className="lg:col-span-8 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
                            <div className="pb-3 border-b border-slate-100 mb-4 flex justify-between items-center">
                              <h4 className="font-extrabold text-[#003a70] text-xs uppercase tracking-wider">Tugas Mandiri yang Harus Anda Selesaikan</h4>
                              <span className="text-xs bg-slate-100 text-slate-600 px-3 py-1 rounded-lg font-bold">
                                Total: {tasks.filter(t => t.assignedNim === activeUser.nim).length} Penugasan
                              </span>
                            </div>

                            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                              {tasks.filter(t => t.assignedNim === activeUser.nim).length === 0 ? (
                                <div className="col-span-2 py-12 text-center text-slate-400 text-xs font-semibold">
                                  Belum ada penugasan terdelegasi untuk Anda. Silakan gulir ke panel di bawah untuk mendelegasikan tugas pokok ke diri sendiri pertama kali.
                                </div>
                              ) : (
                                tasks.filter(t => t.assignedNim === activeUser.nim).map(task => {
                                  const isCompleted = task.status === 'Completed';
                                  const validPoints = (task.points || []).filter(p => typeof p === 'string' && p.trim() !== '');
                                  const points = validPoints.length > 0 ? validPoints : [
                                    'Analisis kebutuhan dasar & topologi',
                                    'Melakukan konfigurasi / pengerjaan teknis spesifik',
                                    'Melakukan pengetesan jalur & ketersambungan',
                                    'Pembuatan dokumentasi & draf Google Doc laporan'
                                  ];
                                  const checkedArr = task.pointsChecked || new Array(points.length).fill(false);
                                  const checkedCount = checkedArr.filter(Boolean).length;
                                  const pct = Math.round((checkedCount / points.length) * 100);

                                  return (
                                    <div key={task.taskId} className={`p-5 border rounded-3xl flex flex-col justify-between transition-all duration-300 ${isCompleted ? 'border-emerald-200 bg-emerald-50/25 shadow-sm' : 'border-blue-105 bg-[#f7fafe]/70 hover:bg-white shadow-sm'}`}>
                                      <div>
                                        <div className="flex items-center justify-between mb-2.5">
                                          <span className={`text-[9.5px] font-bold px-2 py-0.5 rounded ${
                                            task.category === 'Jaringan' ? 'bg-sky-100 text-sky-800' :
                                            task.category === 'Website' ? 'bg-emerald-100 text-emerald-800' :
                                            'bg-slate-100 text-slate-800'
                                          }`}>
                                            {task.category}
                                          </span>
                                          <span className={`text-[10px] font-bold ${isCompleted ? 'text-emerald-700' : 'text-amber-700'}`}>
                                            {task.status === 'Completed' ? 'Completed' : `In Progress / Pending (${pct}%)`}
                                          </span>
                                        </div>
                                        <h5 className="font-extrabold text-slate-850 text-sm mb-1.5 leading-snug">{task.taskName}</h5>
                                        <p className="text-slate-500 text-xs leading-relaxed mb-4 font-semibold">{task.description}</p>

                                        {/* Milestones Checkpoints Section */}
                                        <div className="mt-4 pt-3 border-t border-slate-100/90 space-y-2">
                                          <div className="flex items-center justify-between text-[10.5px] font-bold text-slate-700">
                                            <span className="text-slate-800">Capaian Progres Lapangan</span>
                                            <span className="text-[#003a70]">{checkedCount} / {points.length} ({pct}%)</span>
                                          </div>
                                          <div className="w-full bg-slate-100 h-2 rounded-full overflow-hidden shadow-inner">
                                            <div 
                                              className={`h-full transition-all duration-300 ${pct === 100 ? 'bg-emerald-600' : 'bg-[#003a70]'}`}
                                              style={{ width: `${pct}%` }}
                                            />
                                          </div>

                                          <div className="space-y-1.5 mt-2.5">
                                            {points.map((pt, index) => (
                                              <label 
                                                key={index} 
                                                className={`flex items-start gap-2.5 p-2 rounded-xl border text-[11px] leading-relaxed cursor-pointer transition-colors ${
                                                  checkedArr[index] 
                                                    ? 'bg-emerald-50/45 border-emerald-100 text-emerald-900 font-bold' 
                                                    : 'bg-white hover:bg-slate-50 border-slate-100 text-slate-650'
                                                }`}
                                              >
                                                <input 
                                                  type="checkbox"
                                                  className="mt-0.5 rounded border-slate-300 text-[#003a70] focus:ring-[#003a70] disabled:opacity-50"
                                                  checked={!!checkedArr[index]}
                                                  disabled={isCompleted}
                                                  onChange={() => handlePointToggle(task, index)}
                                                />
                                                <span className={checkedArr[index] ? 'line-through text-slate-400' : ''}>
                                                  {pt}
                                                </span>
                                              </label>
                                            ))}
                                          </div>
                                        </div>

                                        {/* Display Registered Daily Timelines */}
                                        {task.timelineLogs && task.timelineLogs.length > 0 && (
                                          <div className="mt-3.5 bg-blue-50/30 border border-blue-100/50 p-3 rounded-2xl text-[10.5px] space-y-1.5 text-slate-655">
                                            <div className="font-bold text-[#003a70] flex items-center justify-between uppercase tracking-wider text-[9px] border-b pb-1 border-blue-100/40">
                                              <span>Timeline Kejadian Harian</span>
                                              <span className="font-mono bg-blue-100/60 px-1.5 py-0.5 rounded text-blue-900">{task.timelineLogs.length} Kegiatan</span>
                                            </div>
                                            <div className="max-h-[90px] overflow-y-auto space-y-1 pr-1">
                                              {task.timelineLogs.map((log, idx) => (
                                                <div key={idx} className="flex justify-between items-start border-b border-blue-50/40 pb-1.5 last:border-0 last:pb-0 gap-1.5">
                                                  <span className="text-slate-800 font-extrabold shrink-0">{log.date}:</span>
                                                  <span className="text-left w-full text-slate-500 font-medium leading-relaxed">{log.description}</span>
                                                  <span className="font-mono text-slate-400 shrink-0">({log.hours}j)</span>
                                                </div>
                                              ))}
                                            </div>
                                          </div>
                                        )}

                                      </div>

                                      <div className="border-t border-slate-100 pt-3.5 flex flex-col gap-2 mt-4 text-[11px]">
                                        <div className="flex justify-between items-center text-[10px] text-slate-400 font-mono">
                                          <span>Dari NIM: {task.createdBy}</span>
                                          {task.workType && (
                                            <span className={`font-sans font-bold px-2 py-0.5 rounded text-[8.5px] ${task.workType === 'Kelompok' ? 'bg-purple-100 text-purple-800' : 'bg-emerald-100 text-emerald-800'}`}>
                                              {task.workType}
                                            </span>
                                          )}
                                        </div>

                                        {isCompleted ? (
                                          <div className="space-y-1.5 pt-1">
                                            <div className="text-emerald-800 font-bold flex items-center gap-1.5 bg-emerald-50 p-2 rounded-xl border border-emerald-100/70 text-[11px] justify-center">
                                              <span>✓ Selesai 100% & Terverifikasi</span>
                                            </div>
                                            {task.googleDocUrl && (
                                              <a 
                                                href={task.googleDocUrl} 
                                                target="_blank" 
                                                rel="noreferrer"
                                                className="block w-full py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-center text-xs font-bold transition-all shadow-sm"
                                              >
                                                📄 Buka Laporan Google Doc
                                              </a>
                                            )}
                                          </div>
                                        ) : (
                                          <div className="space-y-2 pt-1.5 w-full">
                                            {!task.googleDocUrl ? (
                                              <div className="flex gap-2 w-full">
                                                <button 
                                                  type="button"
                                                  onClick={async () => {
                                                    try {
                                                      await db.runMutation('tasks', 'upsert', {
                                                        id: task.taskId,
                                                        masterid: task.masterTaskId || null,
                                                        assignednim: task.assignedNim,
                                                        taskname: task.taskName,
                                                        category: task.category,
                                                        dateassigned: (task as any).dateAssigned || new Date().toISOString(),
                                                        status: task.status,
                                                        progress: pct,
                                                        points: { checked: task.pointsChecked || [], dates: task.checkDates || [] },
                                                        googledocurl: task.googleDocUrl
                                                      });
                                                      Swal.fire({ toast: true, position: 'top-end', icon: 'success', title: `✅ Progres ${pct}% berhasil disimpan!`, showConfirmButton: false, timer: 2500 });
                                                    } catch(err: any) {
                                                      Swal.fire({ icon: 'error', title: 'Gagal Simpan', text: err.message || 'Tidak bisa terhubung ke Supabase' });
                                                    }
                                                  }}
                                                  className="flex-1 py-2 border border-slate-200 hover:bg-slate-50 text-slate-700 rounded-xl transition-all font-bold text-center text-xs"
                                                >
                                                  💾 Simpan Harian
                                                </button>
                                                <button 
                                                  type="button"
                                                  onClick={() => handleDuplicateAndOpenDoc(task)}
                                                  className="flex-1 py-2 bg-[#003a70] hover:bg-[#002244] text-white rounded-xl transition-all font-bold text-center text-xs flex items-center justify-center gap-1"
                                                >
                                                  📄 Buat Laporan Doc
                                                </button>
                                              </div>
                                            ) : (
                                              <div className="space-y-2">
                                                <a 
                                                  href={task.googleDocUrl}
                                                  target="_blank"
                                                  rel="noreferrer"
                                                  className="block w-full py-2 bg-[#f0f6ff] border border-blue-200 text-[#003a70] hover:bg-blue-100 text-center rounded-xl text-xs font-bold transition-all shadow-sm flex items-center justify-center gap-1.5"
                                                >
                                                  📄 Buka Google Doc (Tab Baru) ↗
                                                </a>
                                                <div className="flex gap-2">
                                                  <button 
                                                    type="button"
                                                    onClick={async () => {
                                                      try {
                                                        await db.runMutation('tasks', 'upsert', {
                                                          id: task.taskId,
                                                          masterid: task.masterTaskId || null,
                                                          assignednim: task.assignedNim,
                                                          taskname: task.taskName,
                                                          category: task.category,
                                                          dateassigned: (task as any).dateAssigned || new Date().toISOString(),
                                                          status: task.status,
                                                          progress: pct,
                                                          points: { checked: task.pointsChecked || [], dates: task.checkDates || [] },
                                                          googledocurl: task.googleDocUrl
                                                        });
                                                        Swal.fire({ toast: true, position: 'top-end', icon: 'success', title: `✅ Progres ${pct}% berhasil disimpan!`, showConfirmButton: false, timer: 2500 });
                                                      } catch(err: any) {
                                                        Swal.fire({ icon: 'error', title: 'Gagal Simpan', text: err.message || 'Tidak bisa terhubung ke Supabase' });
                                                      }
                                                    }}
                                                    className="flex-1 py-2 border border-slate-200 hover:bg-slate-50 text-slate-700 rounded-xl transition-all font-bold text-center text-[10px]"
                                                  >
                                                    💾 Simpan Harian
                                                  </button>
                                                  <button 
                                                    type="button"
                                                    onClick={() => handleCompleteTaskDirect(task)}
                                                    className="flex-1 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl transition-all font-bold text-center text-[10px] flex items-center justify-center gap-0.5"
                                                  >
                                                    ✓ Selesaikan & Setor
                                                  </button>
                                                </div>
                                              </div>
                                            )}
                                          </div>
                                        )}
                                      </div>
                                    </div>
                                  );
                                })
                              )}
                            </div>

                            {/* Recent logbook status tracker */}
                            <div className="mt-8 border-t border-slate-100 pt-6">
                              <h5 className="font-bold text-xs text-[#003a70] uppercase tracking-wider mb-4">Aktivitas Logbook Terkini Anda</h5>
                              <div className="overflow-x-auto">
                                <table className="w-full text-left text-xs text-slate-600">
                                  <thead>
                                    <tr className="border-b border-slate-200 text-slate-400 text-[10px] uppercase font-bold tracking-wider">
                                      <th className="pb-2.5 font-bold">Tugas</th>
                                      <th className="pb-2.5 font-bold">Tanggal</th>
                                      <th className="pb-2.5 font-bold">Bukti</th>
                                      <th className="pb-2.5 font-bold text-right">Penilaian Dosen</th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {logbooks.filter(l => l.nim === activeUser.nim).length === 0 ? (
                                      <tr>
                                        <td colSpan={4} className="py-4 text-center text-slate-400 italic">Belum ada logbook teknis yang tersimpan di Sheets.</td>
                                      </tr>
                                    ) : (
                                      logbooks.filter(l => l.nim === activeUser.nim).map(log => (
                                        <tr key={log.logbookId} className="border-b border-slate-100 text-[11px] hover:bg-slate-50/50">
                                          <td className="py-3 font-semibold text-slate-800">{log.taskName}</td>
                                          <td className="py-3 text-slate-400">{log.timestamp.split(' ')[0]}</td>
                                          <td className="py-3">
                                            <a href={log.googleDocUrl || log.fileUrl || '#'} target="_blank" rel="noreferrer" className="text-[#003a70] underline flex items-center gap-1 hover:text-[#002244] font-bold">
                                              Tinjau Berkas
                                            </a>
                                          </td>
                                          <td className="py-3 text-right font-bold text-emerald-600">
                                            {log.grade ? `${log.grade} / 100` : <span className="text-slate-400 font-normal italic">Proses Audit</span>}
                                          </td>
                                        </tr>
                                      ))
                                    )}
                                  </tbody>
                                </table>
                              </div>
                            </div>

                          </div>
                        </div>
                      </div>

                      {/* SECTION 2: Team task delegation (Delegasi & Distribusi Tugas) */}
                      <div className="space-y-4 pt-8 border-t border-slate-200/80">
                        <div className="flex items-center gap-2">
                          <span className="text-lg">⚡</span>
                          <h4 className="font-extrabold text-[#003a70] text-xs uppercase tracking-wider mb-0">II. Delegasi & Distribusi Tugas Pokok Tim</h4>
                        </div>

                        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
                          
                          {/* Left: Delegate Task Form */}
                          <form onSubmit={handleSimulatedCreateTask} className="lg:col-span-5 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm space-y-4">
                            <div className="pb-2 border-b border-slate-100">
                              <h4 className="font-extrabold text-[#003a70] text-xs uppercase tracking-wider mb-0">Delegasikan Penugasan Pokok (PIC)</h4>
                            </div>
                            <div>
                              <label className="block text-xs font-bold text-slate-700 mb-1.5">Pilih Tugas Pokok (Dibuat oleh PIC)</label>
                              <Select 
                                options={[
                                  { value: "", label: "-- Pilih Tugas Pokok Standar --" },
                                  ...masterTasks.map(m => ({
                                    value: m.id,
                                    label: `[${m.category}] ${m.title} (${m.workType})`
                                  }))
                                ]}
                                value={
                                  taskFormMasterId 
                                  ? { 
                                      value: taskFormMasterId, 
                                      label: `[${masterTasks.find(m => m.id === taskFormMasterId)?.category}] ${masterTasks.find(m => m.id === taskFormMasterId)?.title} (${masterTasks.find(m => m.id === taskFormMasterId)?.workType})` 
                                    } 
                                  : { value: "", label: "-- Pilih Tugas Pokok Standar --" }
                                }
                                onChange={(option) => {
                                  setTaskFormMasterId(option?.value || "");
                                }}
                                styles={{
                                  control: (base, state) => ({ ...base, fontSize: '12px', minHeight: '36px', borderColor: state.isFocused ? '#003a70' : '#e2e8f0', boxShadow: state.isFocused ? '0 0 0 1px #003a70' : 'none', borderRadius: '0.5rem' }),
                                  menu: (base) => ({ ...base, fontSize: '12px', zIndex: 50 })
                                }}
                                isSearchable
                                placeholder="Cari Tugas Pokok..."
                              />
                            </div>

                            {taskFormMasterId && (() => {
                              const selectedMaster = masterTasks.find(m => m.id === taskFormMasterId);
                              if (!selectedMaster) return null;
                              return (
                                <div className="p-3 bg-slate-50 border border-slate-200 rounded-xl space-y-1.5">
                                  <div className="flex items-center justify-between">
                                    <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wide">Panduan Deskripsi PIC</span>
                                    <span className={`text-[9px] font-extrabold px-2 py-0.5 rounded-full ${selectedMaster.workType === 'Kelompok' ? 'bg-purple-100 text-purple-800' : 'bg-emerald-100 text-emerald-850'}`}>
                                      {selectedMaster.workType === 'Kelompok' ? '👥 Boleh Berkelompok' : '👤 Kerjakan Sendiri'}
                                    </span>
                                  </div>
                                  <p className="text-xs text-slate-600 leading-relaxed font-semibold">
                                    {selectedMaster.description}
                                  </p>
                                  <div className="pt-1.5 border-t border-slate-100 flex items-center justify-between text-[11px] font-bold">
                                    <span className="text-slate-500">Target Standardisasi Peran:</span>
                                    <span className="text-[#003a70] bg-blue-50 px-2 py-0.5 rounded border border-blue-200">
                                      🎯 {selectedMaster.targetRole || 'Semua Peran'}
                                    </span>
                                  </div>
                                </div>
                              );
                            })()}

                            <div>
                              <label className="block text-xs font-bold text-slate-700 mb-1.5">Pilih Personel Pelaksana (Berasal dari Peran Terkait)</label>
                              <Select 
                                options={[
                                  { value: "", label: "-- Pilih Mahasiswa Magang --" },
                                  ...users.filter(u => u.role !== 'PIC').map(u => {
                                    const selectedMaster = masterTasks.find(m => m.id === taskFormMasterId);
                                    const isMatch = selectedMaster ? (!selectedMaster.targetRole || selectedMaster.targetRole === 'Semua Peran' || selectedMaster.targetRole === u.role) : true;
                                    return {
                                      value: u.nim,
                                      label: `${u.name} (NIM: ${u.nim} - Peran: ${u.role}) ${isMatch ? '✅ Cocok' : '⚠️ Berbeda'}`
                                    };
                                  })
                                ]}
                                value={
                                  taskFormAssigned 
                                  ? {
                                      value: taskFormAssigned,
                                      label: (() => {
                                        const u = users.find(usr => usr.nim === taskFormAssigned);
                                        const selectedMaster = masterTasks.find(m => m.id === taskFormMasterId);
                                        const isMatch = selectedMaster && u ? (!selectedMaster.targetRole || selectedMaster.targetRole === 'Semua Peran' || selectedMaster.targetRole === u.role) : true;
                                        return u ? `${u.name} (NIM: ${u.nim} - Peran: ${u.role}) ${isMatch ? '✅ Cocok' : '⚠️ Berbeda'}` : "-- Pilih Mahasiswa Magang --";
                                      })()
                                    }
                                  : { value: "", label: "-- Pilih Mahasiswa Magang --" }
                                }
                                onChange={(option) => {
                                  setTaskFormAssigned(option?.value || "");
                                }}
                                styles={{
                                  control: (base, state) => ({ ...base, fontSize: '12px', minHeight: '36px', borderColor: state.isFocused ? '#003a70' : '#e2e8f0', boxShadow: state.isFocused ? '0 0 0 1px #003a70' : 'none', borderRadius: '0.5rem' }),
                                  menu: (base) => ({ ...base, fontSize: '12px', zIndex: 50 })
                                }}
                                isSearchable
                                placeholder="Cari Mahasiswa..."
                              />
                            </div>

                            {taskFormMasterId && taskFormAssigned && (() => {
                              const selectedMaster = masterTasks.find(m => m.id === taskFormMasterId);
                              const assignedUser = users.find(u => u.nim === taskFormAssigned);
                              if (!selectedMaster || !assignedUser) return null;
                              const tgt = selectedMaster.targetRole || 'Semua Peran';
                              if (tgt !== 'Semua Peran' && tgt !== assignedUser.role) {
                                return (
                                  <div className="p-3 bg-amber-50 border border-amber-200 rounded-xl flex items-start gap-2 text-[11px] text-amber-855 font-semibold leading-relaxed">
                                    <AlertCircle className="w-4.5 h-4.5 text-amber-600 shrink-0 mt-0.5" />
                                    <div>
                                      Perhatian: Peran mahasiswa <strong>{assignedUser.name} ({assignedUser.role})</strong> berbeda dengan kompetensi target tugas <strong>({tgt})</strong>.
                                    </div>
                                  </div>
                                );
                              } else {
                                return (
                                  <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-xl flex items-start gap-2 text-[11px] text-emerald-800 font-semibold leading-relaxed">
                                    <CheckCircle className="w-4.5 h-4.5 text-emerald-600 shrink-0 mt-0.5" />
                                    <div>
                                      Kompatibilitas Sempurna: Peran mahasiswa <strong>{assignedUser.name} ({assignedUser.role})</strong> telah sesuai dengan target tugas pokok!
                                    </div>
                                  </div>
                                );
                              }
                            })()}
                            
                            <button 
                              type="submit"
                              className="w-full py-2.5 bg-[#003a70] hover:bg-[#002244] text-white rounded-lg text-xs font-bold transition-all flex items-center justify-center gap-2 shadow"
                            >
                              <Plus className="w-4 h-4" /> Delegasikan Instruksi Kerja
                            </button>
                          </form>

                          {/* Right: Distributed List progress */}
                          <div className="lg:col-span-7 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
                            <div className="pb-3 border-b border-slate-100 mb-4">
                              <h4 className="font-extrabold text-[#003a70] text-xs uppercase tracking-wider mb-0">Status Distribusi Penugasan</h4>
                            </div>
                            <div className="overflow-x-auto">
                              <table className="w-full text-left text-xs border-collapse">
                                <thead>
                                  <tr className="border-b border-slate-200 text-slate-400 text-[10px] uppercase font-bold tracking-wider">
                                    <th className="pb-2">Aktivitas Kerja</th>
                                    <th className="pb-2">Petugas Magang</th>
                                    <th className="pb-2 text-right">Status Penugasan</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {tasks.length === 0 ? (
                                    <tr>
                                      <td colSpan={3} className="py-8 text-center text-slate-400 italic">Belum ada delegasi tugas terdistribusi.</td>
                                    </tr>
                                  ) : (
                                   tasks.map(t => {
                                      const assignee = users.find(u => u.nim === t.assignedNim);
                                      const validPoints = (t.points || []).filter(p => typeof p === 'string' && p.trim() !== '');
                                      const points = validPoints.length > 0 ? validPoints : [
                                        'Analisis kebutuhan dasar & topologi',
                                        'Melakukan konfigurasi / pengerjaan teknis spesifik',
                                        'Melakukan pengetesan jalur & ketersambungan',
                                        'Pembuatan dokumentasi & draf Google Doc laporan'
                                      ];
                                      const checkedArr = t.pointsChecked || new Array(points.length).fill(false);
                                      const checkedCount = checkedArr.filter(Boolean).length;
                                      const pct = Math.round((checkedCount / points.length) * 100);

                                      return (
                                        <tr key={t.taskId} className="border-b border-slate-100 text-[11px] hover:bg-slate-50/50">
                                          <td className="py-3">
                                            <div className="font-bold text-slate-800">{t.taskName}</div>
                                            <div className="flex items-center gap-2 mt-1">
                                              <span className="text-[9px] text-[#003a70] bg-slate-100 font-mono px-1 rounded uppercase">{t.category}</span>
                                              <span className={`text-[9px] font-bold px-1 rounded ${t.workType === 'Kelompok' ? 'bg-purple-50 text-purple-700 border border-purple-150' : 'bg-emerald-50 text-emerald-700'}`}>{t.workType || 'Individu'}</span>
                                            </div>
                                          </td>
                                          <td className="py-3 text-slate-600 font-medium font-sans">
                                            {assignee ? assignee.name : t.assignedNim}
                                          </td>
                                          <td className="py-3 text-right">
                                            <div className="flex flex-col items-end gap-1">
                                              <span className={`px-2 py-0.5 rounded-full text-[9.5px] font-extrabold ${t.status === 'Completed' ? 'bg-emerald-100 text-emerald-800 border border-emerald-250' : 'bg-blue-50 text-[#003a70] border border-blue-150'}`}>
                                                {t.status === 'Completed' ? '✓ 100% Selesai' : `${pct}% Dituntaskan`}
                                              </span>
                                              <span className="text-[9.5px] text-slate-400 font-medium">
                                                {checkedCount} dari {points.length} Capaian
                                              </span>
                                              {/* Mini progress bar tracker */}
                                              <div className="w-16 bg-slate-100 h-1 rounded-full overflow-hidden mt-0.5 border border-slate-200">
                                                <div 
                                                  className={`h-full transition-all duration-300 ${pct === 100 ? 'bg-emerald-500' : 'bg-[#003a70]'}`}
                                                  style={{ width: `${pct}%` }}
                                                />
                                              </div>
                                            </div>
                                          </td>
                                        </tr>
                                      );
                                    })
                                  )}
                                </tbody>
                              </table>
                            </div>
                          </div>

                        </div>
                      </div>

                    </div>
                  )}

                  {/* ===================== SIMULATOR SUBVIEW: PIC WORKSPACE ===================== */}
                  {activeUser && activeUser.role === 'PIC' && (
                    <div className="space-y-6">
                      
                      {/* Sub-Tabs Grid or Layout for PIC panels */}
                      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
                        
                        {/* 1. Logbooks finished list & certificate trigger */}
                        <div className="lg:col-span-8 bg-white p-6 rounded-2xl border border-slate-200/80 shadow-sm space-y-4">
                          <div className="pb-3 border-b border-slate-100 flex items-center justify-between">
                            <h4 className="font-extrabold text-[#003a70] text-xs uppercase tracking-wider mb-0">Rekapitulasi Logbook Masuk</h4>
                            <span className="text-[10px] bg-[#003a70] text-white font-bold px-2 py-0.5 rounded-md">Sheets Logbooks</span>
                          </div>

                          <div className="overflow-x-auto">
                            <table className="w-full text-left text-xs border-collapse">
                              <thead>
                                <tr className="border-b border-slate-200 text-slate-400 text-[10px] uppercase font-bold tracking-wider">
                                  <th className="pb-3">Siswa (NIM)</th>
                                  <th className="pb-3">Tugas & Deskripsi</th>
                                  <th className="pb-3">File Bukti</th>
                                  <th className="pb-3">Audit Nilai</th>
                                  <th className="pb-3 text-right">Action SKEK</th>
                                </tr>
                              </thead>
                              <tbody>
                                {logbooks.length === 0 ? (
                                  <tr>
                                    <td colSpan={5} className="py-8 text-center text-slate-400 italic font-medium">Belum ada logbook masuk dari para anggota magang.</td>
                                  </tr>
                                ) : (
                                  logbooks.map(log => (
                                    <tr key={log.logbookId} className="border-b border-slate-100 text-[11px] hover:bg-slate-50/50">
                                      <td className="py-3">
                                        <div className="font-bold text-slate-800">{log.studentName}</div>
                                        <div className="text-[10px] text-slate-400 font-mono">NIM {log.nim}</div>
                                      </td>
                                      <td className="py-3">
                                        <div className="font-semibold text-slate-700">{log.taskName}</div>
                                        <p className="text-[10px] text-slate-400 max-w-[200px] truncate">{log.workDescription}</p>
                                      </td>
                                      <td className="py-3">
                                        <a href={log.googleDocUrl || log.fileUrl || '#'} target="_blank" rel="noreferrer" className="text-[#003a70] underline hover:text-[#002244] font-bold">
                                          Tinjau Berkas
                                        </a>
                                      </td>
                                      <td className="py-3">
                                        {log.grade ? (
                                          <span className="px-2 py-0.5 bg-emerald-100 text-emerald-800 rounded font-bold">{log.grade}</span>
                                        ) : (
                                          <button 
                                            onClick={() => {
                                              setSelectedLogbookToGrade(log);
                                              setGradeInput('');
                                              setGradeNotes('');
                                            }}
                                            className="px-2 py-1 bg-amber-500 hover:bg-amber-600 rounded-md text-slate-950 font-extrabold text-[10px]"
                                          >
                                            Beri Nilai
                                          </button>
                                        )}
                                      </td>
                                      <td className="py-3 text-right">
                                        <div className="flex items-center justify-end gap-1.5">
                                          <button 
                                            onClick={() => triggerPDFGenerationSimulation(log.nim)}
                                            className="px-2.5 py-1 bg-[#003a70] hover:bg-[#002244] rounded-md text-white font-bold text-[10px] tracking-wide transition-all"
                                            title="Cetak PDF Sertifikat/Rekap"
                                          >
                                            Cetak PDF
                                          </button>
                                          <button 
                                            onClick={() => handleDeleteLogbook(log.logbookId, log.studentName, log.taskName)}
                                            className="p-1.5 bg-red-50 hover:bg-red-150 text-red-650 rounded-md border border-red-200 transition-all cursor-pointer"
                                            title="Hapus Catatan Logbook"
                                          >
                                            <Trash2 className="w-3.5 h-3.5" />
                                          </button>
                                        </div>
                                      </td>
                                    </tr>
                                  ))
                                )}
                              </tbody>
                            </table>
                          </div>
                        </div>

                        {/* 2. PIC Users Profiling and Role Control */}
                        <div className="lg:col-span-4 bg-white p-6 rounded-2xl border border-slate-200/80 shadow-sm space-y-4">
                          <div className="pb-3 border-b border-slate-100 flex items-center justify-between">
                            <div>
                              <h4 className="font-extrabold text-[#003a70] text-xs uppercase tracking-wider mb-0">Manajemen User (Database Pengguna)</h4>
                              <p className="text-[9.5px] text-slate-400 mt-0.5 leading-none">Kelola semua akun, hak akses PIC, & hapus data</p>
                            </div>
                            <span className="text-[10px] bg-blue-105 text-[#003a70] font-mono font-bold px-2 py-0.5 rounded">
                              {users.length} Akun
                            </span>
                          </div>
                          
                          <div className="space-y-3 max-h-[350px] overflow-y-auto pr-2">
                            {users.map(u => {
                              const isSelf = u.nim === activeUser?.nim;
                              const isConfirming = userToDeleteConfirm === u.nim;
                              
                              return (
                                <div key={u.nim} className={`p-3 border rounded-xl flex flex-col gap-2 transition-all duration-200 ${isConfirming ? 'border-red-300 bg-red-50/20' : 'border-slate-100 bg-slate-50/50'}`}>
                                  <div className="flex items-start justify-between gap-1.5 w-full">
                                    <div>
                                      <div className="font-bold text-xs text-slate-800 flex items-center gap-1.5 flex-wrap">
                                        <span>{u.name}</span>
                                        {isSelf && (
                                          <span className="text-[8px] font-mono bg-[#003a70] text-white px-1 rounded">ANDA</span>
                                        )}
                                      </div>
                                      <div className="text-[10px] text-slate-400 font-mono">NIM {u.nim}</div>
                                      {u.role !== 'PIC' && (
                                        <div className="text-[9.5px] text-slate-500 mt-1 font-medium bg-slate-100 p-1 rounded border border-slate-200/50">
                                          ⏱️ Periode: <strong className="text-slate-700">{u.periode || 3} Bulan</strong> <br/>
                                          📅 Mulai: <span className="font-mono text-slate-650">{u.tanggalMulai || '-'}</span> s/d <span className="font-mono text-slate-650">{u.tanggalSelesai || '-'}</span>
                                        </div>
                                      )}
                                    </div>
                                    
                                    <div className="flex items-center gap-1 shrink-0">
                                      {/* Edit User Button (Universal) */}
                                      <button
                                        type="button"
                                        onClick={() => {
                                          setEditingUser(u);
                                          setEditUserNim(u.nim);
                                          setEditUserName(u.name);
                                          setEditUserEmail(u.email);
                                          setEditUserPassword(u.password || 'password');
                                          setEditUserPeriode(String(u.periode || 3));
                                          setEditUserTanggalMulai(u.tanggalMulai || '2026-03-01');
                                          setEditUserTanggalSelesai(u.tanggalSelesai || '2026-06-01');
                                          setEditUserNomorSurat(u.nomorSurat || '');
                                        }}
                                        title="Edit Profil & Kredensial (Ganti Password, Email, NIS/NIM)"
                                        className="p-1 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-colors border border-transparent hover:border-slate-200 bg-white"
                                      >
                                        <Pencil className="w-3.5 h-3.5" />
                                      </button>

                                      {/* Action Delete (Disallowed for Self) */}
                                      {!isSelf && (
                                        <div className="flex items-center gap-1">
                                          {isConfirming ? (
                                            <div className="flex items-center gap-1">
                                              <button
                                                type="button"
                                                onClick={() => handleDeleteUser(u.nim)}
                                                className="px-2 py-0.5 bg-red-600 text-white rounded text-[9.5px] font-bold shadow-sm hover:bg-red-700"
                                              >
                                                Yakin?
                                              </button>
                                              <button
                                                type="button"
                                                onClick={() => setUserToDeleteConfirm(null)}
                                                className="px-1.5 py-0.5 bg-slate-200 text-slate-700 rounded text-[9.5px] font-bold"
                                              >
                                                Batal
                                              </button>
                                            </div>
                                          ) : (
                                            <button
                                              type="button"
                                              onClick={() => setUserToDeleteConfirm(u.nim)}
                                              title="Hapus user ini"
                                              className="p-1 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded-lg transition-colors border border-transparent hover:border-slate-200 bg-white"
                                            >
                                              <Trash2 className="w-3.5 h-3.5" />
                                            </button>
                                          )}
                                        </div>
                                      )}
                                    </div>
                                  </div>

                                  <div className="flex items-center justify-between border-t border-dashed border-slate-250/50 pt-2 gap-2">
                                    <span className="text-[10px] text-slate-500 font-medium">Ubah Peran:</span>
                                    <div className="w-[140px]">
                                      <Select 
                                        options={[
                                          { value: "PIC", label: "PIC / Pelatih" },
                                          ...Object.keys(jobdesks).map(rName => ({ value: rName, label: rName }))
                                        ]}
                                        value={{ value: u.role, label: u.role === "PIC" ? "PIC / Pelatih" : u.role }}
                                        onChange={(option) => {
                                          if(option) handleSimulatedRoleChange(u.nim, option.value);
                                        }}
                                        isDisabled={isSelf}
                                        styles={{
                                          control: (base) => ({ ...base, minHeight: '26px', height: '26px', fontSize: '10px', borderRadius: '0.5rem' }),
                                          valueContainer: (base) => ({ ...base, padding: '0 6px' }),
                                          input: (base) => ({ ...base, margin: '0' }),
                                          dropdownIndicator: (base) => ({ ...base, padding: '2px' }),
                                          menu: (base) => ({ ...base, fontSize: '10px', zIndex: 50 })
                                        }}
                                        isSearchable
                                        menuPortalTarget={document.body}
                                      />
                                    </div>
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                          <div className="p-3 bg-blue-50 text-[#003a70] rounded-xl border border-blue-100 text-[10px] leading-relaxed font-semibold">
                            💡 <strong>Manajemen User:</strong> Multi-PIC didukung penuh! Anda dapat mempromosikan user lain menjadi PIC tambahan, atau menghapus user selain diri Anda sendiri.
                          </div>
                        </div>

                      </div>

                      {/* Sub-Row: PIC Student Assignments Control */}
                      <div className="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-sm space-y-4">
                        <div className="pb-3 border-b border-slate-100 flex items-center justify-between">
                          <div>
                            <h4 className="font-extrabold text-[#003a70] text-xs uppercase tracking-wider mb-0">Status Distribusi & Kontrol Penugasan Mahasiswa (PIC View)</h4>
                            <p className="text-[9.5px] text-slate-400 mt-0.5 leading-none">Kelola semua penugasan aktif, lihat progres checklist harian siswa, atau hapus penugasan mereka</p>
                          </div>
                          <span className="text-[10px] bg-[#003a70] text-white font-bold px-2 py-0.5 rounded-md">
                            {tasks.length} Penugasan Aktif
                          </span>
                        </div>

                        <div className="overflow-x-auto">
                          <table className="w-full text-left text-xs border-collapse">
                            <thead>
                              <tr className="border-b border-slate-200 text-slate-400 text-[10px] uppercase font-bold tracking-wider">
                                <th className="pb-3">Aktivitas Kerja / Tugas</th>
                                <th className="pb-3">Petugas Magang</th>
                                <th className="pb-3">Progres Checklist</th>
                                <th className="pb-3 text-right">Aksi</th>
                              </tr>
                            </thead>
                            <tbody>
                              {tasks.length === 0 ? (
                                <tr>
                                  <td colSpan={4} className="py-8 text-center text-slate-400 italic">Belum ada delegasi tugas terdistribusi ke mahasiswa. Ketua kelompok dapat membagikan modul tugas PIC.</td>
                                </tr>
                              ) : (
                                tasks.map(t => {
                                  const assignee = users.find(u => u.nim === t.assignedNim);
                                  const points = t.points || [
                                    'Analisis kebutuhan dasar & topologi',
                                    'Melakukan konfigurasi / pengerjaan teknis spesifik',
                                    'Melakukan pengetesan jalur & verifikasi ketersambungan',
                                    'Pembuatan dokumentasi & draf Google Doc laporan'
                                  ];
                                  const checkedArr = t.pointsChecked || new Array(points.length).fill(false);
                                  const checkedCount = checkedArr.filter(Boolean).length;
                                  const pct = Math.round((checkedCount / points.length) * 100);

                                  return (
                                    <tr key={t.taskId} className="border-b border-slate-100 text-[11px] hover:bg-slate-50/50">
                                      <td className="py-3">
                                        <div className="font-bold text-slate-850">{t.taskName}</div>
                                        <div className="flex items-center gap-2 mt-1">
                                          <span className="text-[9px] text-[#003a70] bg-slate-100 font-mono px-1 rounded uppercase">{t.category}</span>
                                          <span className={`text-[9px] font-bold px-1 rounded ${t.workType === 'Kelompok' ? 'bg-purple-50 text-purple-700 border border-purple-150' : 'bg-emerald-50 text-emerald-700'}`}>{t.workType || 'Individu'}</span>
                                        </div>
                                      </td>
                                      <td className="py-3 font-semibold text-slate-700">
                                        {assignee ? (
                                          <div>
                                            <div>{assignee.name}</div>
                                            <div className="text-[10px] text-slate-400 font-mono">NIM {assignee.nim}</div>
                                          </div>
                                        ) : (
                                          <span className="text-slate-405">NIM {t.assignedNim}</span>
                                        )}
                                      </td>
                                      <td className="py-3">
                                        <div className="flex flex-col gap-1 max-w-[120px]">
                                          <span className={`px-2 py-0.5 rounded text-[9px] font-extrabold w-max ${t.status === 'Completed' ? 'bg-emerald-100 text-emerald-800 border border-emerald-250' : 'bg-blue-50 text-[#003a70] border border-blue-150'}`}>
                                            {t.status === 'Completed' ? '✓ 100% Selesai' : `${pct}% Dituntaskan`}
                                          </span>
                                          <div className="w-24 bg-slate-100 h-1 rounded-full overflow-hidden border border-slate-205">
                                            <div 
                                              className={`h-full transition-all duration-300 ${t.status === 'Completed' ? 'bg-emerald-500' : 'bg-[#003a70]'}`}
                                              style={{ width: `${pct}%` }}
                                            />
                                          </div>
                                        </div>
                                      </td>
                                      <td className="py-3 text-right">
                                        <button 
                                          type="button"
                                          onClick={() => handleDeleteTask(t.taskId)}
                                          title="Hapus / Batalkan Penugasan ini"
                                          className="p-1.5 bg-red-50 hover:bg-red-100 text-red-650 rounded-md border border-red-200 transition-all cursor-pointer"
                                        >
                                          <Trash2 className="w-3.5 h-3.5" />
                                        </button>
                                      </td>
                                    </tr>
                                  );
                                })
                              )}
                            </tbody>
                          </table>
                        </div>
                      </div>

                      {/* Sub-Row: PIC Master Tasks Bank & Creator */}
                      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 border-t border-slate-200 pt-6">
                        
                        <div className="lg:col-span-5 flex flex-col gap-6">
                          <form onSubmit={handlePicCreateMasterTask} className="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-sm space-y-4 w-full">
                          <div className="pb-2 border-b border-slate-100 flex items-center justify-between">
                            <h4 className="font-extrabold text-[#003a70] text-xs uppercase tracking-wider mb-0">
                              {editingMasterTaskId ? "✍ Edit Tugas Pokok Standar" : "Definisikan Standar Tugas Pokok (PIC)"}
                            </h4>
                            <span className="text-[10px] bg-[#003a70] text-white font-bold px-2 py-0.5 rounded-md">Master Task</span>
                          </div>
                          
                          <div>
                            <label className="block text-xs font-bold text-slate-700 mb-1.5">Judul Tugas</label>
                            <input 
                              type="text" 
                              className="w-full text-xs px-3.5 py-2 border border-slate-200 rounded-lg outline-none focus:border-[#003a70] text-slate-800"
                              placeholder="Contoh: Instalasi Rackmount Switch Lab"
                              value={picTaskTitle}
                              onChange={(e) => setPicTaskTitle(e.target.value)}
                              required
                            />
                          </div>

                          <div className="grid grid-cols-2 gap-3">
                            <div>
                              <label className="block text-xs font-bold text-slate-700 mb-1.5">Kategori</label>
                              <Select 
                                options={categories.map(cat => ({ value: cat, label: cat }))}
                                value={{ value: picTaskCategory, label: picTaskCategory }}
                                onChange={(option) => setPicTaskCategory(option?.value || "")}
                                styles={{
                                  control: (base, state) => ({ ...base, fontSize: '12px', minHeight: '36px', borderColor: state.isFocused ? '#003a70' : '#003a70', boxShadow: state.isFocused ? '0 0 0 1px #003a70' : 'none', borderRadius: '0.5rem' }),
                                  menu: (base) => ({ ...base, fontSize: '12px', zIndex: 50 })
                                }}
                                isSearchable
                              />
                            </div>
                            <div>
                              <label className="block text-xs font-bold text-slate-700 mb-1.5">Syarat Pelaksanaan</label>
                              <Select 
                                options={[
                                  { value: "Individu", label: "Individu (Mandiri)" },
                                  { value: "Kelompok", label: "Kelompok (Tim)" }
                                ]}
                                value={{ value: picTaskWorkType, label: picTaskWorkType === "Individu" ? "Individu (Mandiri)" : "Kelompok (Tim)" }}
                                onChange={(option) => setPicTaskWorkType((option?.value as any) || "Individu")}
                                styles={{
                                  control: (base, state) => ({ ...base, fontSize: '12px', minHeight: '36px', borderColor: state.isFocused ? '#003a70' : '#e2e8f0', boxShadow: state.isFocused ? '0 0 0 1px #003a70' : 'none', borderRadius: '0.5rem' }),
                                  menu: (base) => ({ ...base, fontSize: '12px', zIndex: 50 })
                                }}
                                isSearchable
                              />
                            </div>
                          </div>

                          <div>
                            <label className="block text-xs font-bold text-slate-700 mb-1.5">Target Peran Standardisasi (Kompetensi Khusus)</label>
                            <Select 
                              options={[
                                { value: "Semua Peran", label: "-- Semua Peran (Generalist) --" },
                                ...Object.keys(jobdesks).map(rName => ({ value: rName, label: rName }))
                              ]}
                              value={{ value: picTaskTargetRole, label: picTaskTargetRole === "Semua Peran" ? "-- Semua Peran (Generalist) --" : picTaskTargetRole }}
                              onChange={(option) => setPicTaskTargetRole(option?.value || "Semua Peran")}
                              styles={{
                                control: (base, state) => ({ ...base, fontSize: '12px', minHeight: '36px', borderColor: state.isFocused ? '#003a70' : '#003a70', boxShadow: state.isFocused ? '0 0 0 1px #003a70' : 'none', borderRadius: '0.5rem', fontWeight: 'bold' }),
                                menu: (base) => ({ ...base, fontSize: '12px', zIndex: 50, fontWeight: 'normal' })
                              }}
                              isSearchable
                            />
                          </div>

                          <div>
                            <label className="block text-xs font-bold text-slate-700 mb-1.5 flex justify-between items-center">
                              <span>1. Deskripsi / Tujuan Tugas Pokok</span>
                              <span className="text-[10px] text-[#003a70] font-bold">Latar Belakang</span>
                            </label>
                            <textarea 
                              className="w-full text-xs px-3.5 py-2 border border-slate-200 rounded-lg outline-none focus:border-[#003a70] text-slate-800"
                              rows={2}
                              placeholder="Tulis ringkasan latar belakang, instruksi kelola, atau aturan standar tugas..."
                              value={picTaskDesc}
                              onChange={(e) => setPicTaskDesc(e.target.value)}
                              required
                            />
                          </div>

                          <div>
                            <label className="block text-xs font-bold text-slate-700 mb-1.5 flex justify-between items-center">
                              <span>2. Output yang Harus Dikerjakan (Checklist Capaian)</span>
                              <span className="text-[10px] text-[#003a70] font-bold">Butir: {picChecklistItems.length}</span>
                            </label>
                            <p className="text-[10px] text-slate-400 mb-2 leading-snug">
                              Tambahkan, ubah, atau hapus butir checklist capaian teknis harian siswa sebelum dipublikasi.
                            </p>

                            {/* Checklist Item Adder Input */}
                            <div className="flex gap-2 mb-3">
                              <input 
                                type="text" 
                                className="flex-1 text-xs px-3.5 py-2 border border-slate-200 rounded-lg outline-none focus:border-[#003a70] text-slate-800 bg-slate-50/50"
                                placeholder="Tambah butir (misal: Setup VLAN Dosen)..."
                                value={newCheckItemText}
                                onChange={(e) => setNewCheckItemText(e.target.value)}
                                onKeyDown={(e) => {
                                  if (e.key === 'Enter') {
                                    e.preventDefault();
                                    if (newCheckItemText.trim()) {
                                      setPicChecklistItems([...picChecklistItems, newCheckItemText.trim()]);
                                      setNewCheckItemText('');
                                    }
                                  }
                                }}
                              />
                              <button 
                                type="button"
                                onClick={() => {
                                  if (newCheckItemText.trim()) {
                                    setPicChecklistItems([...picChecklistItems, newCheckItemText.trim()]);
                                    setNewCheckItemText('');
                                  }
                                }}
                                className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-extrabold transition-all shrink-0"
                              >
                                ＋ Tambah
                              </button>
                            </div>

                            {/* Checklist Items Interactive List */}
                            <div className="space-y-1.5 max-h-[220px] overflow-y-auto border border-slate-150 p-2 rounded-xl bg-slate-50/50 mb-3">
                              {picChecklistItems.length === 0 ? (
                                <p className="text-[10.5px] text-slate-400 italic text-center py-4">Belum ada butir checklist. Klik tombol "Tambah" di atas.</p>
                              ) : (
                                picChecklistItems.map((item, index) => {
                                  const isEditing = editingCheckItemIdx === index;
                                  return (
                                    <div key={index} className="transition-all">
                                      {isEditing ? (
                                        <div className="flex gap-1.5 items-center bg-amber-50/70 border border-amber-100 p-1.5 rounded-lg w-full">
                                          <input 
                                            type="text" 
                                            className="flex-1 text-[11px] px-2 py-1 border border-slate-300 rounded outline-none text-slate-800 bg-white"
                                            value={editingCheckItemText}
                                            onChange={(e) => setEditingCheckItemText(e.target.value)}
                                            onKeyDown={(e) => {
                                              if (e.key === 'Enter') {
                                                e.preventDefault();
                                                if (editingCheckItemText.trim()) {
                                                  const updated = [...picChecklistItems];
                                                  updated[index] = editingCheckItemText.trim();
                                                  setPicChecklistItems(updated);
                                                  setEditingCheckItemIdx(null);
                                                }
                                              }
                                            }}
                                            autoFocus
                                          />
                                          <button 
                                            type="button" 
                                            onClick={() => {
                                              if (editingCheckItemText.trim()) {
                                                const updated = [...picChecklistItems];
                                                updated[index] = editingCheckItemText.trim();
                                                setPicChecklistItems(updated);
                                                setEditingCheckItemIdx(null);
                                              }
                                            }} 
                                            className="px-2 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded text-[10px] font-bold"
                                            title="Simpan perubahan"
                                          >
                                            ✓
                                          </button>
                                          <button 
                                            type="button" 
                                            onClick={() => setEditingCheckItemIdx(null)} 
                                            className="px-2 py-1 bg-slate-200 hover:bg-slate-300 text-slate-705 rounded text-[10px] font-bold"
                                            title="Batal"
                                          >
                                            ✕
                                          </button>
                                        </div>
                                      ) : (
                                        <div className="flex items-center justify-between p-2 rounded-lg bg-white border border-slate-100/80 shadow-sm text-[11px] leading-relaxed hover:border-slate-200">
                                          <div className="flex items-start gap-2 text-slate-700 max-w-[80%]">
                                            <input 
                                              type="checkbox" 
                                              disabled 
                                              className="rounded text-[#003a70] mt-0.5 border-slate-300 pointer-events-none shrink-0" 
                                            />
                                            <span className="font-semibold text-slate-750 text-left leading-snug">{item}</span>
                                          </div>
                                          <div className="flex gap-1 shrink-0 ml-1">
                                            <button 
                                              type="button" 
                                              onClick={() => {
                                                setEditingCheckItemIdx(index);
                                                setEditingCheckItemText(item);
                                              }} 
                                              className="p-1 hover:bg-slate-100 text-slate-400 hover:text-slate-650 rounded transition-colors"
                                              title="Ubah butir capaian ini"
                                            >
                                              <Pencil className="w-3 h-3" />
                                            </button>
                                            <button 
                                              type="button" 
                                              onClick={() => setPicChecklistItems(picChecklistItems.filter((_, i) => i !== index))} 
                                              className="p-1 hover:bg-red-50 text-slate-400 hover:text-red-600 rounded transition-colors"
                                              title="Hapus butir capaian ini"
                                            >
                                              <Trash2 className="w-3 h-3" />
                                            </button>
                                          </div>
                                        </div>
                                      )}
                                    </div>
                                  );
                                })
                              )}
                            </div>
                          </div>

                          {editingMasterTaskId ? (
                            <div className="flex gap-2">
                              <button 
                                type="submit"
                                disabled={picChecklistItems.length === 0}
                                className="flex-1 py-2.5 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white rounded-xl text-xs font-bold transition-all flex items-center justify-center gap-1.5 shadow-sm font-sans"
                              >
                                ✓ Simpan Perubahan
                              </button>
                              <button 
                                type="button"
                                onClick={handleResetPicTaskForm}
                                className="px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold transition-all border border-slate-200"
                              >
                                Batal
                              </button>
                            </div>
                          ) : (
                            <button 
                              type="submit"
                              disabled={picChecklistItems.length === 0}
                              className="w-full py-2.5 bg-[#003a70] hover:bg-[#002244] disabled:opacity-50 text-white rounded-xl text-xs font-bold transition-all flex items-center justify-center gap-2 shadow-sm font-sans"
                            >
                              <Plus className="w-4 h-4" /> Publikasikan Tugas Pokok
                            </button>
                          )}
                        </form>

                        {/* Fitur Kelola Kategori PIC */}
                        <div className="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-sm space-y-4 w-full">
                          <div className="pb-2 border-b border-slate-100 flex items-center justify-between">
                            <h4 className="font-extrabold text-[#003a70] text-xs uppercase tracking-wider mb-0">
                              📂 Kelola Kategori Pekerjaan (PIC)
                            </h4>
                            <span className="text-[10px] bg-blue-50 text-[#003a70] font-mono font-bold px-2 py-0.5 rounded border border-[#003a70]/10">
                              {categories.length} Kategori
                            </span>
                          </div>
                          
                          <div className="flex gap-2">
                            <input 
                              type="text"
                              className="flex-1 text-xs px-3 py-2 border border-slate-200 rounded-lg outline-none focus:border-[#003a70] text-slate-800"
                              placeholder="Tambah kategori baru..."
                              value={newCategoryName}
                              onChange={(e) => setNewCategoryName(e.target.value)}
                            />
                            <button
                              type="button"
                              onClick={async () => {
                                const name = newCategoryName.trim();
                                if (!name) return;
                                if (categories.some(c => (c || '').toLowerCase() === (name || '').toLowerCase())) {
                                  Swal.fire({ toast: true, position: 'top-end', icon: 'warning', title: 'Kategori tersebut sudah terdaftar!', showConfirmButton: false, timer: 2500 });
                                  return;
                                }
                                const updated = [...categories, name];
                                setCategories(updated);
                                setNewCategoryName('');
                                // Langsung simpan ke Supabase
                                try {
                                  await db.runMutation('categories', 'upsert', [{ name }]);
                                  Swal.fire({ toast: true, position: 'top-end', icon: 'success', title: `Kategori "${name}" berhasil ditambahkan!`, showConfirmButton: false, timer: 2500 });
                                } catch (e: any) {
                                  console.error('Gagal menyimpan kategori ke Supabase:', e);
                                  Swal.fire({ icon: 'error', title: 'Gagal Simpan', text: `Kategori sudah ditambahkan di lokal, tapi gagal sync ke server: ${e.message || e}` });
                                }
                              }}
                              className="px-3 py-2 bg-[#003a70] hover:bg-[#002244] text-white rounded-lg text-xs font-bold shrink-0"
                            >
                              Tambah
                            </button>
                          </div>

                          <div className="space-y-1.5 max-h-[200px] overflow-y-auto pr-1">
                            {categories.length === 0 && (
                              <p className="text-[10px] text-slate-400 italic text-center py-3">Belum ada kategori. Tambahkan kategori baru di atas.</p>
                            )}
                            {categories.map((cat) => {
                              const isSystemCategory = ['jaringan', 'website', 'admin'].includes((cat || '').toLowerCase());
                              const isEditingThis = editingCategoryOld === cat;
                              return (
                                <div key={cat} className={`flex items-center justify-between p-2 rounded-lg border text-xs transition-all ${
                                  isEditingThis ? 'bg-amber-50 border-amber-300' : 'bg-slate-50 border-slate-100 text-slate-700'
                                }`}>
                                  {isEditingThis ? (
                                    <input
                                      type="text"
                                      autoFocus
                                      value={editingCategoryNew}
                                      onChange={e => setEditingCategoryNew(e.target.value)}
                                      onKeyDown={async (e) => {
                                        if (e.key === 'Enter') {
                                          const newName = editingCategoryNew.trim();
                                          if (!newName || newName === cat) { setEditingCategoryOld(null); return; }
                                          if (categories.some(c => c !== cat && (c || '').toLowerCase() === newName.toLowerCase())) {
                                            Swal.fire({ toast: true, position: 'top-end', icon: 'warning', title: 'Nama kategori sudah ada!', showConfirmButton: false, timer: 2000 });
                                            return;
                                          }
                                          // Update state
                                          setCategories(prev => prev.map(c => c === cat ? newName : c));
                                          setMasterTasks(prev => prev.map(m => m.category === cat ? { ...m, category: newName } : m));
                                          setTasks(prev => prev.map(t => t.category === cat ? { ...t, category: newName } : t));
                                          setEditingCategoryOld(null);
                                          // Langsung sync ke Supabase
                                          try {
                                            await db.runMutation('categories', 'delete', null, { column: 'name', value: cat });
                                            await db.runMutation('categories', 'upsert', [{ name: newName }]);
                                            Swal.fire({ toast: true, position: 'top-end', icon: 'success', title: `Kategori diubah menjadi "${newName}"`, showConfirmButton: false, timer: 2500 });
                                          } catch (err: any) {
                                            console.error('Gagal update kategori di Supabase:', err);
                                            Swal.fire({ icon: 'error', title: 'Gagal Update', text: `Perubahan disimpan di lokal, tapi gagal sync ke server: ${err.message || err}` });
                                          }
                                        } else if (e.key === 'Escape') {
                                          setEditingCategoryOld(null);
                                        }
                                      }}
                                      className="flex-1 text-xs px-2 py-0.5 border border-amber-300 rounded outline-none bg-white mr-2"
                                      placeholder="Nama baru..."
                                    />
                                  ) : (
                                    <span className="font-extrabold uppercase tracking-wide text-[10px]">{cat}</span>
                                  )}
                                  <div className="flex items-center gap-1">
                                    {isSystemCategory && !isEditingThis && <span className="text-[8px] text-slate-400 font-mono hidden sm:inline">Sistem</span>}
                                    {isEditingThis ? (
                                      <>
                                        <button
                                          type="button"
                                          onClick={async () => {
                                            const newName = editingCategoryNew.trim();
                                            if (!newName || newName === cat) { setEditingCategoryOld(null); return; }
                                            if (categories.some(c => c !== cat && (c || '').toLowerCase() === newName.toLowerCase())) {
                                              Swal.fire({ toast: true, position: 'top-end', icon: 'warning', title: 'Nama kategori sudah ada!', showConfirmButton: false, timer: 2000 });
                                              return;
                                            }
                                            setCategories(prev => prev.map(c => c === cat ? newName : c));
                                            setMasterTasks(prev => prev.map(m => m.category === cat ? { ...m, category: newName } : m));
                                            setTasks(prev => prev.map(t => t.category === cat ? { ...t, category: newName } : t));
                                            setEditingCategoryOld(null);
                                            try {
                                              await db.runMutation('categories', 'delete', null, { column: 'name', value: cat });
                                              await db.runMutation('categories', 'upsert', [{ name: newName }]);
                                              Swal.fire({ toast: true, position: 'top-end', icon: 'success', title: `Kategori diubah menjadi "${newName}"`, showConfirmButton: false, timer: 2500 });
                                            } catch (err: any) {
                                              Swal.fire({ icon: 'error', title: 'Gagal Update', text: `${err.message || err}` });
                                            }
                                          }}
                                          className="px-2 py-0.5 bg-emerald-500 hover:bg-emerald-600 text-white rounded text-[10px] font-bold"
                                        >Simpan</button>
                                        <button
                                          type="button"
                                          onClick={() => setEditingCategoryOld(null)}
                                          className="px-2 py-0.5 bg-slate-200 hover:bg-slate-300 text-slate-600 rounded text-[10px] font-bold"
                                        >Batal</button>
                                      </>
                                    ) : (
                                      <button
                                        type="button"
                                        onClick={() => { setEditingCategoryOld(cat); setEditingCategoryNew(cat); }}
                                        className="p-1 hover:bg-amber-50 text-slate-400 hover:text-amber-500 rounded transition-all"
                                        title="Edit Nama Kategori"
                                      >
                                        <Pencil className="w-3 h-3" />
                                      </button>
                                    )}
                                    {!isEditingThis && (
                                    <button
                                      type="button"
                                      onClick={() => {
                                        Swal.fire({
                                          title: isSystemCategory ? 'Hapus Kategori Sistem?' : 'Hapus Kategori?',
                                          html: isSystemCategory 
                                            ? `PERHATIAN: Kategori <b>"${cat}"</b> adalah kategori sistem utama. Menghapus kategori ini mungkin berdampak pada fungsionalitas lain.<br/><br/>Apakah Anda yakin ingin menghapus?`
                                            : `Apakah Anda yakin ingin menghapus kategori <b>"${cat}"</b>?`,
                                          icon: 'warning',
                                          showCancelButton: true,
                                          confirmButtonColor: '#d33',
                                          cancelButtonColor: '#3085d6',
                                          confirmButtonText: 'Ya, Hapus!',
                                          cancelButtonText: 'Batal'
                                        }).then((result) => {
                                          if (result.isConfirmed) {
                                            const fallback = categories.find(c => c !== cat) || 'Jaringan';
                                            setCategories(categories.filter(c => c !== cat));
                                            setMasterTasks(prev => prev.map(m => m.category === cat ? { ...m, category: fallback } : m));
                                            setTasks(prev => prev.map(t => t.category === cat ? { ...t, category: fallback } : t));
                                            db.runMutation('master_tasks', 'update', { category: fallback }, { column: 'category', value: cat }).catch(() => {});
                                            db.runMutation('tasks', 'update', { category: fallback }, { column: 'category', value: cat }).catch(() => {});
                                            db.runMutation('categories', 'delete', null, { column: 'name', value: cat })
                                              .then(() => {
                                                Swal.fire({ toast: true, position: 'top-end', icon: 'success', title: `Kategori "${cat}" berhasil dihapus.`, showConfirmButton: false, timer: 3000 });
                                              })
                                              .catch(e => {
                                                Swal.fire({ icon: 'error', title: 'Database Error', text: `Gagal menghapus dari server: ${e.message || e}` });
                                              });
                                          }
                                        });
                                      }}
                                      className="p-1 hover:bg-red-50 text-slate-400 hover:text-red-500 rounded transition-all"
                                      title="Hapus Kategori"
                                    >
                                      <Trash2 className="w-3.5 h-3.5" />
                                    </button>
                                    )}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      </div>

                        {/* Master Tasks Bank List */}
                        <div className="lg:col-span-7 bg-white p-6 rounded-2xl border border-slate-200/80 shadow-sm flex flex-col h-full max-h-[850px]">
                          <div className="pb-3 border-b border-slate-100 flex items-center justify-between mb-4">
                            <h4 className="font-extrabold text-[#003a70] text-xs uppercase tracking-wider mb-0">Bank Tugas Pokok Aktif</h4>
                            <span className="text-[10px] text-slate-500 font-bold">Total: {masterTasks.length} Standarisasi</span>
                          </div>

                          <div className="flex-1 overflow-y-auto space-y-2.5 pr-2">
                            {masterTasks.map(m => {
                              const isEditingThis = editingMasterTaskId === m.id;
                              return (
                                <div 
                                  key={m.id} 
                                  className={`p-3 border rounded-xl transition-all space-y-1.5 ${
                                    isEditingThis 
                                      ? 'border-emerald-500 bg-emerald-50/20 ring-1 ring-emerald-400' 
                                      : 'border-slate-100 bg-slate-50/50 hover:bg-slate-50'
                                  }`}
                                >
                                  <div className="flex items-start justify-between gap-1.5">
                                    <div className="space-y-1 max-w-[75%]">
                                      <span className="font-extrabold text-xs text-slate-850 block leading-tight">
                                        {m.title}
                                      </span>
                                      <div className="flex flex-wrap items-center gap-1">
                                        <span className={`text-[8px] font-extrabold px-1.5 py-0.5 rounded ${
                                          (m.category || '').toLowerCase().includes('jaringan') ? 'bg-sky-100 text-sky-800' :
                                          (m.category || '').toLowerCase().includes('website') ? 'bg-emerald-100 text-emerald-800' :
                                          (m.category || '').toLowerCase().includes('admin') ? 'bg-slate-100 text-slate-800' :
                                          (m.category || '').toLowerCase().includes('desain') ? 'bg-purple-100 text-purple-800' :
                                          'bg-amber-100 text-amber-850'
                                        }`}>
                                          {m.category || 'Lainnya'}
                                        </span>
                                        <span className={`text-[8px] font-extrabold px-1.5 py-0.5 rounded ${
                                          m.workType === 'Kelompok' ? 'bg-purple-100 text-purple-800' :
                                          'bg-amber-100 text-amber-850'
                                        }`}>
                                          {m.workType === 'Kelompok' ? '👥 Kelompok' : '👤 Individu'}
                                        </span>
                                        <span className="text-[8px] font-extrabold px-1.5 py-0.5 rounded bg-blue-105 bg-[#003a70]/10 text-[#003a70] border border-[#003a70]/20">
                                          🎯 {m.targetRole || 'Semua Peran'}
                                        </span>
                                      </div>
                                    </div>

                                    {/* Action Buttons */}
                                    <div className="flex items-center gap-1 shrink-0">
                                      <button
                                        type="button"
                                        onClick={() => handleSelectMasterTaskToEdit(m)}
                                        title="Pilih & Edit Tugas Pokok Ini"
                                        className={`p-1 rounded transition-colors ${
                                          isEditingThis 
                                            ? 'bg-emerald-600 text-white hover:bg-emerald-700' 
                                            : 'bg-white hover:bg-slate-105 text-slate-500 border border-slate-200'
                                        }`}
                                      >
                                        <Pencil className="w-3.5 h-3.5" />
                                      </button>
                                      <button
                                        type="button"
                                        onClick={() => handlePicDeleteMasterTask(m.id, m.title)}
                                        title="Hapus Tugas Pokok"
                                        className="p-1 bg-white hover:bg-red-50 text-red-650 rounded border border-slate-200 hover:border-red-200 transition-colors"
                                      >
                                        <Trash2 className="w-3.5 h-3.5 animate-pulse-slow" />
                                      </button>
                                    </div>
                                  </div>
                                  <p className="text-[10.5px] text-slate-500 font-medium leading-relaxed">
                                    <strong>Tujuan:</strong> {m.description}
                                  </p>
                                  {m.points && m.points.length > 0 && (
                                    <div className="mt-1.5 space-y-1 bg-[#f7fafe] p-2 rounded-lg border border-slate-100">
                                      <div className="text-[8.5px] uppercase font-extrabold text-[#003a70] tracking-wider flex items-center justify-between">
                                        <span>📋 Target Output ({m.points.length} Butir):</span>
                                        <span className="text-[8px] bg-sky-100 px-1 py-0.2 rounded font-mono font-bold text-sky-850">Progres Aktif</span>
                                      </div>
                                      <ul className="list-disc pl-3.5 space-y-0.5 text-[9.5px] text-slate-600 font-medium font-sans">
                                        {m.points.map((pt, idx) => (
                                          <li key={idx}>{pt}</li>
                                        ))}
                                      </ul>
                                    </div>
                                  )}
                                </div>
                              );
                            })}
                          </div>
                        </div>

                      </div>

                      {/* Jobdesk config panel inside PIC workspace */}
                      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 border-t border-slate-200 pt-6">
                        <div className="lg:col-span-6 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm space-y-4">
                          <div className="border-b border-slate-100 pb-2">
                            <h4 className="font-extrabold text-[#003a70] text-xs uppercase tracking-wider mb-1">Master Manajemen Peran & Jobdesk</h4>
                            <p className="text-[10px] text-slate-400">Tambah peran kustom baru & modifikasi deskripsi tanggung jawab yang disinkronkan ke Sheets.</p>
                          </div>

                          <div className="space-y-3">
                            {/* Role selector list with delete capability for non-core roles */}
                            <div className="flex flex-wrap gap-1.5 p-1 bg-slate-50/50 rounded-xl border border-slate-100">
                              {Object.keys(jobdesks).map(rName => {
                                const isCore = ['PIC', 'Ketua', 'Anggota'].includes(rName);
                                const isSelected = selectedRoleForJobdesk === rName;
                                return (
                                  <div 
                                    key={rName} 
                                    className={`inline-flex items-center gap-1 px-2.5 py-1 text-[10px] font-extrabold rounded-lg transition-all border ${
                                      isSelected 
                                        ? 'bg-[#003a70] text-white border-[#003a70]' 
                                        : 'bg-white text-slate-600 hover:bg-slate-50 border-slate-200'
                                    }`}
                                  >
                                    <button
                                      type="button"
                                      onClick={() => setSelectedRoleForJobdesk(rName)}
                                      className="outline-none"
                                    >
                                      {rName}
                                    </button>
                                      <button
                                        type="button"
                                        title={`Hapus peran ${rName}`}
                                        onClick={(e) => {
                                          e.stopPropagation();
                                          Swal.fire({
                                            title: 'Hapus Peran?',
                                            text: isCore
                                              ? `PERHATIAN: Peran "${rName}" adalah peran utama sistem. Menghapus peran ini mungkin memengaruhi fungsionalitas lain. Yakin ingin menghapus?`
                                              : `Apakah Anda yakin ingin menghapus peran kustom "${rName}"? Seluruh data jobdesk ini akan terhapus.`,
                                            icon: 'warning',
                                            showCancelButton: true,
                                            confirmButtonColor: '#d33',
                                            cancelButtonColor: '#3085d6',
                                            confirmButtonText: 'Ya, hapus!',
                                            cancelButtonText: 'Batal'
                                          }).then((result) => {
                                            if (result.isConfirmed) {
                                              const updated = { ...jobdesks };
                                              delete updated[rName];
                                              db.runMutation('jobdesks', 'delete', null, { column: 'rolename', value: rName })
                                                .then(() => console.log("Jobdesk berhasil dihapus dari Supabase."))
                                                .catch(e => console.error("Gagal menghapus jobdesk:", e));
                                              setJobdesks(updated);
                                              if (selectedRoleForJobdesk === rName) {
                                                setSelectedRoleForJobdesk(Object.keys(updated)[0] || '');
                                              }
                                              Swal.fire({
                                                toast: true,
                                                position: 'top-end',
                                                icon: 'success',
                                                title: `Berhasil menghapus peran "${rName}".`,
                                                showConfirmButton: false,
                                                timer: 3000
                                              });
                                            }
                                          });
                                        }}
                                        className={`p-0.5 px-1.5 flex items-center justify-center rounded-full hover:bg-red-500 hover:text-white shrink-0 ${isSelected ? 'text-white/70' : 'text-slate-400'}`}
                                      >
                                        &times;
                                      </button>
                                  </div>
                                );
                              })}
                            </div>

                            {/* Create dynamic role form input */}
                            <div className="flex gap-2">
                              <input
                                type="text"
                                placeholder="Nama peran baru (misal: Koordinator ICT)"
                                className="text-xs px-3 py-1.5 border border-slate-205 rounded-lg outline-none focus:border-[#003a70] text-slate-800 flex-1 bg-white"
                                value={newRoleName}
                                onChange={(e) => setNewRoleName(e.target.value)}
                              />
                              <button
                                type="button"
                                onClick={(e) => {
                                  e.preventDefault();
                                  const name = newRoleName.trim();
                                  if (!name) return;
                                  if (jobdesks[name]) {
                                    Swal.fire({
                                      toast: true,
                                      position: 'top-end',
                                      icon: 'error',
                                      title: `Peran "${name}" sudah terdaftar.`,
                                      showConfirmButton: false,
                                      timer: 3000
                                    });
                                    return;
                                  }
                                  setJobdesks({
                                    ...jobdesks,
                                    [name]: `Tuntutan deskripsi kerja standar untuk peran kustom ${name}:\n1. Melakukan koordinasi spesifik bidang secara berkala\n2. Melaporkan inventarisasi dan status kerja perbaikan ke Ketua\n3. Turut menjaga standardisasi teknis lab Vokasi IT Network`
                                  });
                                  setSelectedRoleForJobdesk(name);
                                  setNewRoleName('');
                                  Swal.fire({
                                    toast: true,
                                    position: 'top-end',
                                    icon: 'success',
                                    title: `Peran kustom baru "${name}" berhasil ditambahkan!`,
                                    showConfirmButton: false,
                                    timer: 3000
                                  });
                                }}
                                className="px-3.5 py-1.5 bg-emerald-600 font-bold hover:bg-emerald-700 text-white text-[10px] rounded-lg transition-all shadow-sm shrink-0"
                              >
                                + Peran Baru
                              </button>
                            </div>

                            {/* Custom textarea to edit selected jobdesks */}
                            <div className="space-y-1.5 pt-1">
                              <div className="flex items-center justify-between">
                                <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">
                                  Kandungan Jobdesk Peran: <span className="text-[#003a70] font-extrabold">[{selectedRoleForJobdesk}]</span>
                                </label>
                                <span className="text-[9px] text-slate-400 italic">Simpan setelah melakukan edit</span>
                              </div>
                              <textarea 
                                className="w-full text-xs font-mono p-3 border border-slate-200 rounded-xl bg-slate-50 text-slate-800 outline-none focus:border-[#003a70] focus:bg-white transition-all leading-normal"
                                rows={5}
                                value={jobdesks[selectedRoleForJobdesk] || ''}
                                onChange={(e) => {
                                  setJobdesks({ ...jobdesks, [selectedRoleForJobdesk]: e.target.value });
                                }}
                                placeholder="Tuliskan butir-butir standar jobdesk untuk peran khusus ini..."
                              />
                              <button 
                                type="button"
                                onClick={() => {
                                  triggerCallSimulation(`Merekam standar jobdesk kustom [${selectedRoleForJobdesk}] ke Sheets...`, () => {
                                    Swal.fire({
                                      toast: true,
                                      position: 'top-end',
                                      icon: 'success',
                                      title: `Jobdesk Peran [${selectedRoleForJobdesk}] tersimpan!`,
                                      showConfirmButton: false,
                                      timer: 3000
                                    });
                                  });
                                }}
                                className="w-full py-2 bg-[#003a70] text-white rounded-lg text-xs font-bold hover:bg-[#002244] transition-all flex items-center justify-center gap-2 shadow-sm font-sans"
                              >
                                <Check className="w-3.5 h-3.5" /> Simpan Perubahan Jobdesk ke Sheets
                              </button>
                            </div>
                          </div>
                        </div>

                        {/* Script configurations (drive ids templates) */}
                        <div className="lg:col-span-6 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
                          <h4 className="font-extrabold text-[#003a70] text-xs uppercase tracking-wider border-b border-slate-100 pb-2 mb-4">Pengaturan Parameter Properties Google APIs</h4>
                          <div className="space-y-4">
                            <div>
                              <label className="block text-[10px] uppercase font-extrabold text-[#003a70] mb-0.5">ID Folder Hasil Laporan Duplikasi (Google Drive)</label>
                              <span className="text-[9px] text-slate-400 block mb-1.5 leading-normal">Folder kearsipan di mana seluruh salinan draf laporan Google Docs siswa secara otomatis ditempatkan.</span>
                              <div className="flex gap-2">
                                <input 
                                  type="text" 
                                  className="flex-1 text-xs font-mono px-3.5 py-2 border border-slate-200 rounded-lg outline-none text-slate-800 bg-slate-50/50 focus:bg-white focus:border-[#003a70]"
                                  value={driveIdInput}
                                  onChange={(e) => setDriveIdInput(e.target.value)}
                                />
                                <button 
                                  type="button"
                                  onClick={() => {
                                    const url = driveIdInput.startsWith('http') 
                                      ? driveIdInput 
                                      : `https://drive.google.com/drive/folders/${driveIdInput}`;
                                    window.open(url, '_blank');
                                  }}
                                  className="px-3 py-2 bg-slate-100 hover:bg-[#003a70]/10 text-[#003a70] hover:text-[#002244] rounded-lg text-xs font-semibold transition-all border border-slate-200 flex items-center justify-center gap-1 shrink-0"
                                  title="Lihat Berkas yang Sudah Tersubmit di Google Drive"
                                >
                                  <ExternalLink className="w-3.5 h-3.5" /> Lihat Hasil
                                </button>
                              </div>
                            </div>
                            <div>
                              <label className="block text-[10px] uppercase font-extrabold text-[#003a70] mb-0.5">ID Template Master Laporan Kerja (Google Docs)</label>
                              <span className="text-[9px] text-slate-400 block mb-1.5 leading-normal">Template dokumen formal UNY yang diduplikasikan langsung untuk diisi progress harian & disubmisi oleh tim magang.</span>
                              <div className="flex gap-2">
                                <input 
                                  type="text" 
                                  className="flex-1 text-xs font-mono px-3.5 py-2 border border-slate-200 rounded-lg outline-none text-slate-800 bg-slate-50/50 focus:bg-white focus:border-[#003a70]"
                                  value={docTemplateInput}
                                  onChange={(e) => setDocTemplateInput(e.target.value)}
                                />
                                <button 
                                  type="button"
                                  onClick={() => {
                                    const url = docTemplateInput.startsWith('http') 
                                      ? docTemplateInput 
                                      : `https://docs.google.com/document/d/${docTemplateInput}/edit`;
                                    window.open(url, '_blank');
                                  }}
                                  className="px-3 py-2 bg-slate-100 hover:bg-[#003a70]/10 text-[#003a70] hover:text-[#002244] rounded-lg text-xs font-semibold transition-all border border-slate-200 flex items-center justify-center gap-1 shrink-0"
                                  title="Modifikasi Template Laporan Google Docs"
                                >
                                  <ExternalLink className="w-3.5 h-3.5" /> Modifikasi
                                </button>
                              </div>
                            </div>
                            <div>
                               <label className="block text-[10px] uppercase font-extrabold text-[#003a70] mb-0.5">ID Template Portofolio Akhir (Google Docs)</label>
                               <span className="text-[9px] text-slate-400 block mb-1.5 leading-normal">Template dokumen Portofolio Rekapitulasi Akhir yang digunakan saat cetak PDF Portofolio. Harus dipisahkan dari template Logbook!</span>
                               <div className="flex gap-2">
                                 <input 
                                   type="text" 
                                   className="flex-1 text-xs font-mono px-3.5 py-2 border border-slate-200 rounded-lg outline-none text-slate-800 bg-slate-50/50 focus:bg-white focus:border-[#003a70]"
                                   placeholder="Masukkan ID Google Docs template Portofolio..."
                                   value={portfolioTemplateInput}
                                   onChange={(e) => setPortfolioTemplateInput(e.target.value)}
                                 />
                                 <button 
                                   type="button"
                                   onClick={() => {
                                     if (!portfolioTemplateInput) return;
                                     const url = portfolioTemplateInput.startsWith('http') 
                                       ? portfolioTemplateInput 
                                       : `https://docs.google.com/document/d/${portfolioTemplateInput}/edit`;
                                     window.open(url, '_blank');
                                   }}
                                   className="px-3 py-2 bg-slate-100 hover:bg-[#003a70]/10 text-[#003a70] hover:text-[#002244] rounded-lg text-xs font-semibold transition-all border border-slate-200 flex items-center justify-center gap-1 shrink-0"
                                   title="Lihat/Edit Template Portofolio"
                                 >
                                   <ExternalLink className="w-3.5 h-3.5" /> Modifikasi
                                 </button>
                               </div>
                            </div>
                            <div>
                              <label className="block text-[10px] uppercase font-extrabold text-[#003a70] mb-0.5">ID Template Sertifikat Kelulusan (Google Slides)</label>
                              <span className="text-[9px] text-slate-400 block mb-1.5 leading-normal">Template desain sertifikat magang bernomor seri dari UNY.</span>
                              <div className="flex gap-2">
                                <input 
                                  type="text" 
                                  className="flex-1 text-xs font-mono px-3.5 py-2 border border-slate-200 rounded-lg outline-none text-slate-800 bg-slate-50/50 focus:bg-white focus:border-[#003a70]"
                                  value={slideTemplateInput}
                                  onChange={(e) => setSlideTemplateInput(e.target.value)}
                                />
                                <button 
                                  type="button"
                                  onClick={() => {
                                    const url = slideTemplateInput.startsWith('http') 
                                      ? slideTemplateInput 
                                      : `https://docs.google.com/presentation/d/${slideTemplateInput}/edit`;
                                    window.open(url, '_blank');
                                  }}
                                  className="px-3 py-2 bg-slate-100 hover:bg-[#003a70]/10 text-[#003a70] hover:text-[#002244] rounded-lg text-xs font-semibold transition-all border border-slate-200 flex items-center justify-center gap-1 shrink-0"
                                  title="Modifikasi Template Sertifikat Google Slides"
                                >
                                  <ExternalLink className="w-3.5 h-3.5" /> Modifikasi
                                </button>
                              </div>
                            </div>
                            <div className="pt-2">
                              <button 
                                onClick={() => {
                                  triggerCallSimulation("Menyimpan script properties Google APIs ke dalam Supabase...", () => {
                                      setPropertiesData([
                                        { propKey: 'SUPABASE_URL', propValue: import.meta.env.VITE_SUPABASE_URL || '', lastUpdated: new Date().toISOString() },
                                        { propKey: 'SUPABASE_ANON_KEY', propValue: import.meta.env.VITE_SUPABASE_ANON_KEY || '', lastUpdated: new Date().toISOString() },
                                        { propKey: 'DRIVE_FOLDER_ID', propValue: driveIdInput, lastUpdated: new Date().toISOString() },
                                        { propKey: 'TEMPLATE_M_DOCS_ID', propValue: docTemplateInput, lastUpdated: new Date().toISOString() },
                                        { propKey: 'TEMPLATE_PORTFOLIO_ID', propValue: portfolioTemplateInput, lastUpdated: new Date().toISOString() },
                                        { propKey: 'TEMPLATE_M_SLIDES_ID', propValue: slideTemplateInput, lastUpdated: new Date().toISOString() }
                                      ]);
                                      customAlert(`Sukses: Parameter Sistem API Google berhasil direkam ke Database Supabase secara Global!`, 'success');
                                  });
                                }}
                                className="w-full sm:w-auto px-5 py-2.5 bg-[#003a70] hover:bg-[#002244] text-white rounded-xl text-xs font-bold transition-all shadow-sm font-sans flex items-center justify-center gap-1.5"
                              >
                                ✓ Simpan Konfigurasi Parameter
                              </button>

                              {/* Live Spreadsheet Properties table display */}
                              <div className="border border-slate-200 rounded-xl overflow-hidden mt-6">
                                <div className="bg-[#003a70]/5 px-3.5 py-2 border-b border-slate-200 text-[10px] font-extrabold text-[#003a70] tracking-wider uppercase flex items-center justify-between">
                                  <span>Live Spreadsheet View: Tab &quot;Properties&quot; (Config)</span>
                                  <span className="text-[8px] bg-emerald-100 text-emerald-850 px-1.5 py-0.5 rounded font-mono font-bold">TEREKAM AKTIF</span>
                                </div>
                                <div className="overflow-x-auto">
                                  <table className="w-full text-left text-[10.5px] border-collapse bg-slate-50/50">
                                    <thead>
                                      <tr className="bg-slate-100/80 text-slate-500 font-bold border-b border-slate-200">
                                        <th className="p-2 pl-3">PropertyKey</th>
                                        <th className="p-2">PropertyValue</th>
                                        <th className="p-2 text-right pr-3">Last Updated</th>
                                      </tr>
                                    </thead>
                                    <tbody className="divide-y divide-slate-100 font-mono text-slate-700">
                                      {propertiesData.map(p => (
                                        <tr key={p.propKey}>
                                          <td className="p-2 pl-3 font-semibold text-[#003a70]">{p.propKey}</td>
                                          <td className="p-2 truncate max-w-[170px]" title={p.propValue}>{p.propValue}</td>
                                          <td className="p-2 text-right pr-3 text-slate-500 whitespace-nowrap">{new Date(p.lastUpdated).toLocaleString('id-ID')}</td>
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>

                      {/* PowerPoint (PPTX) & Google Slides Certificate customizer */}
                      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 border-t border-slate-200 pt-6">
                        <div className="lg:col-span-12 bg-slate-50 border border-slate-200 rounded-3xl p-6 space-y-6">
                          <div className="flex flex-col md:flex-row md:items-center justify-between border-b border-slate-200 pb-4 gap-4">
                            <div>
                              <div className="flex items-center gap-2">
                                <span className="p-1.5 bg-[#003a70]/10 text-[#003a70] rounded-lg">
                                  <Laptop className="w-5 h-5" />
                                </span>
                                <h4 className="font-extrabold text-[#003a70] text-sm uppercase tracking-wider">Kustomisasi Desain Template Sertifikat dari PPTX / Google Slides</h4>
                              </div>
                              <p className="text-xs text-slate-500 mt-1">Panduan lengkap & panel interaktif untuk memetakan penempatan tag sertifikat dari PowerPoint Anda.</p>
                            </div>
                            <div className="flex flex-wrap gap-2">
                              <span className="text-[10px] font-bold px-3 py-1 rounded-full bg-blue-100 text-blue-900 border border-blue-200">
                                📑 Standard: 100% Cocok Dengan Slides & PPTX
                              </span>
                            </div>
                          </div>

                          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
                            {/* Step Explainer Column */}
                            <div className="lg:col-span-5 space-y-4">
                              <div className="bg-white p-4 rounded-xl border border-slate-200 space-y-3.5 shadow-sm">
                                <h5 className="font-bold text-[#003a70] text-xs uppercase tracking-wider">Alur Kerja Desain Mandiri</h5>
                                
                                <div className="space-y-4 text-xs text-slate-650 font-medium">
                                  <div className="flex gap-2.5">
                                    <div className="w-5 h-5 rounded-full bg-[#003a70] text-white flex items-center justify-center font-bold text-[10px] shrink-0">1</div>
                                    <p className="leading-relaxed">
                                      <strong>Desain di PPTX:</strong> Buat slide horizontal 16:9 di PowerPoint dengan ornamen logo UNY, garis emas, tanda tangan, dsb.
                                    </p>
                                  </div>
                                  <div className="flex gap-2.5">
                                    <div className="w-5 h-5 rounded-full bg-[#003a70] text-white flex items-center justify-center font-bold text-[10px] shrink-0">2</div>
                                    <p className="leading-relaxed">
                                      <strong>Tulis Marker Tag:</strong> Taruh box tulisan biasa berisi teks marker, misal: <code className="bg-slate-100 px-1 py-0.5 rounded text-[#054480] font-bold font-mono">{"{{NAMA}}"}</code>, <code className="bg-slate-100 px-1 py-0.5 rounded text-[#054480] font-bold font-mono">{"{{NIM}}"}</code>, <code className="bg-slate-100 px-1 py-0.5 rounded text-[#054480] font-bold font-mono">{"{{PREDIKAT}}"}</code>.
                                    </p>
                                  </div>
                                  <div className="flex gap-2.5">
                                    <div className="w-5 h-5 rounded-full bg-[#003a70] text-white flex items-center justify-center font-bold text-[10px] shrink-0">3</div>
                                    <p className="leading-relaxed">
                                      <strong>Imporkan ke Slides:</strong> Unggah file <code>.pptx</code> Anda ke Google Drive, buka dengan Google Slides, lalu simpan sebagai file Slides asli (mendapat ID file unik).
                                    </p>
                                  </div>
                                  <div className="flex gap-2.5">
                                    <div className="w-5 h-5 rounded-full bg-[#003a70] text-white flex items-center justify-center font-bold text-[10px] shrink-0">4</div>
                                    <p className="leading-relaxed">
                                      <strong>Sinkronkan Tag di Bawah:</strong> Sesuaikan penamaan tag di panel kanan ini. Apps Script Anda akan memindai objek slide dan mengganti tag tersebut secara otomatis!
                                    </p>
                                  </div>
                                </div>
                              </div>

                              {/* Interactive Tag Configurator */}
                              <div className="bg-white p-4 rounded-xl border border-slate-200/80 space-y-3 shadow-sm">
                                <h5 className="font-bold text-[#003a70] text-xs uppercase tracking-wider">Pemetaan Tag Marker (PPTX)</h5>
                                
                                <div className="grid grid-cols-2 gap-2.5">
                                  <div>
                                    <label className="block text-[9px] font-bold text-slate-500 uppercase">Tag Nama Mahasiswa</label>
                                    <input 
                                      type="text" 
                                      className="w-full text-xs font-mono px-2 py-1.5 border border-slate-200 rounded bg-slate-50 text-slate-800" 
                                      value={certTagName} 
                                      onChange={(e) => setCertTagName(e.target.value)} 
                                    />
                                  </div>
                                  <div>
                                    <label className="block text-[9px] font-bold text-slate-500 uppercase">Tag Nomor Sertifikat</label>
                                    <input 
                                      type="text" 
                                      className="w-full text-xs font-mono px-2 py-1.5 border border-slate-200 rounded bg-slate-50 text-slate-800" 
                                      value={certTagCertNum} 
                                      onChange={(e) => setCertTagCertNum(e.target.value)} 
                                    />
                                  </div>
                                  <div>
                                    <label className="block text-[9px] font-bold text-slate-500 uppercase">Tag NIM Mahasiswa</label>
                                    <input 
                                      type="text" 
                                      className="w-full text-xs font-mono px-2 py-1.5 border border-slate-200 rounded bg-slate-50 text-slate-800" 
                                      value={certTagNim} 
                                      onChange={(e) => setCertTagNim(e.target.value)} 
                                    />
                                  </div>
                                  <div>
                                    <label className="block text-[9px] font-bold text-slate-500 uppercase">Tag Predikat Nilai</label>
                                    <input 
                                      type="text" 
                                      className="w-full text-xs font-mono px-2 py-1.5 border border-slate-200 rounded bg-slate-50 text-slate-800" 
                                      value={certTagGrade} 
                                      onChange={(e) => setCertTagGrade(e.target.value)} 
                                    />
                                  </div>
                                </div>
                                <div className="grid grid-cols-3 gap-3 pt-2">
                                  <div>
                                    <label className="block text-[9px] font-bold text-slate-500 uppercase mb-1">Tag Periode</label>
                                    <input 
                                      type="text" 
                                      className="w-full text-xs font-mono px-2 py-1.5 border border-slate-200 rounded bg-slate-50 text-slate-800"
                                      value={certTagPeriode}
                                      onChange={(e) => setCertTagPeriode(e.target.value)}
                                    />
                                  </div>
                                  <div>
                                    <label className="block text-[9px] font-bold text-slate-500 uppercase mb-1">Tag Tanggal Mulai</label>
                                    <input 
                                      type="text" 
                                      className="w-full text-xs font-mono px-2 py-1.5 border border-slate-200 rounded bg-slate-50 text-slate-800"
                                      value={certTagMulai}
                                      onChange={(e) => setCertTagMulai(e.target.value)}
                                    />
                                  </div>
                                  <div>
                                    <label className="block text-[9px] font-bold text-slate-500 uppercase mb-1">Tag Tanggal Selesai</label>
                                    <input 
                                      type="text" 
                                      className="w-full text-xs font-mono px-2 py-1.5 border border-slate-200 rounded bg-slate-50 text-slate-800"
                                      value={certTagSelesai}
                                      onChange={(e) => setCertTagSelesai(e.target.value)}
                                    />
                                  </div>
                                </div>
                              </div>
                            </div>

                            {/* Live WYSIWYG PPT Design Customizer Sandbox */}
                            <div className="lg:col-span-7 bg-white p-4 rounded-xl border border-slate-200 shadow-sm space-y-4">
                              <div className="flex items-center justify-between border-b border-slate-100 pb-2.5">
                                <h5 className="font-bold text-slate-800 text-xs uppercase tracking-wider">Preview Real-Time Kustomisasi Desain</h5>
                              </div>

                              {slideTemplateInput ? (
                                <iframe 
                                  src={`https://docs.google.com/presentation/d/${slideTemplateInput}/embed?rm=minimal`}
                                  className="w-full aspect-[16/10] rounded-xl border border-slate-200 shadow-sm"
                                  allowFullScreen
                                ></iframe>
                              ) : (
                                <div className="w-full aspect-[16/10] rounded-xl border border-slate-200 shadow-sm bg-slate-50 flex items-center justify-center text-slate-500 font-medium text-sm">
                                  ID Template Google Slides belum diisi di Pengaturan API
                                </div>
                              )}

                              <div className="p-3 bg-blue-50 border border-blue-200 rounded-lg flex items-start gap-2 text-[10.5px] text-[#003a70] leading-relaxed font-semibold">
                                <Sparkles className="w-4 h-4 shrink-0 text-blue-600 mt-0.5" />
                                <div>
                                  <strong>Info API Google Slides:</strong> Preview ini merupakan render langsung dari template Google Slides yang diatur melalui ID API parameter properties. Sistem akan memproses duplikasi dan injeksi tag secara aman menggunakan API Slide di backend tanpa localStorage.
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>

                    </div>
                  )}

                </div>
              </div>
            </div>

            {/* Simulated Workspace Information section */}
            

          </div>
        )}

        {/* ===================== TAB 2: CODE.GS SERVER SIDE ===================== */}
        {activeTab === 'code_gs' && (
          <div className="space-y-4">
            <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-100 pb-4 mb-6">
                <div>
                  <h3 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                    <span className="w-2.5 h-2.5 rounded-full bg-amber-500" /> Code.gs
                  </h3>
                  <p className="text-xs text-slate-500 mt-1 font-medium">
                    Script backend server-side yang menangani komunikasi dengan spreadsheet database, folder berkas Google Drive, templates Slides & Docs, dan automasi e-mail PDF.
                  </p>
                </div>
                <button 
                  onClick={() => copyToClipboard(CODE_GS_CONTENT, 'gs')}
                  className="px-4 py-2 bg-[#003a70] text-white rounded-lg text-xs font-bold font-mono transition-all hover:bg-[#002244] flex items-center gap-2 cursor-pointer shadow-sm"
                >
                  {copiedGs ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
                  {copiedGs ? "Copied!" : "Copy Code.gs"}
                </button>
              </div>

              <div className="relative">
                <div className="absolute top-3 right-3 bg-slate-900 border border-slate-800 rounded px-2.5 py-1 text-[10px] text-slate-400 font-mono">
                  JAVASCRIPT (APPS SCRIPT)
                </div>
                <pre className="text-xs font-mono text-slate-100 bg-slate-950 p-5 rounded-xl overflow-x-auto max-h-[500px] leading-relaxed border border-slate-800">
                  {CODE_GS_CONTENT}
                </pre>
              </div>
            </div>

            <div className="bg-white border border-slate-200 p-6 rounded-2xl shadow-sm">
              <h4 className="font-bold text-slate-900 text-sm mb-3">Penjelasan API Server & Google Apps Script:</h4>
              <ul className="space-y-2 text-xs text-slate-600 font-medium">
                <li>🎯 <strong className="text-slate-800">DriveApp:</strong> Digunakan untuk menyalin template portfolio/doc sertifikat dan mengunggah / sharing berkas logbook anggota secara instan.</li>
                <li>🎯 <strong className="text-slate-800">DocumentApp / SlidesApp:</strong> Memanipulasi tag marker text <code>{"<<Nama>>"}</code>, <code>{"<<NIM>>"}</code>, dan <code>{"<<Nilai Akhir>>"}</code>, serta memformulasikan list rekap dalam grid tabel Docs yang rapi.</li>
                <li>🎯 <strong className="text-slate-800">PropertiesService:</strong> Tempat aman untuk menyimpan credential ID template, folder, sehingga tidak ditulis secara hardcode di baris editor.</li>
              </ul>
            </div>
          </div>
        )}

        {/* ===================== TAB 3: INDEX.HTML CLIENT SIDE ===================== */}
        {activeTab === 'index_html' && (
          <div className="space-y-4">
            <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-100 pb-4 mb-6">
                <div>
                  <h3 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                    <span className="w-2.5 h-2.5 rounded-full bg-sky-500" /> Index.html
                  </h3>
                  <p className="text-xs text-slate-500 mt-1 font-medium">
                    File antarmuka client-side utama menggunakan Bootstrap 5 yang akan dievaluasi oleh HtmlService Google Apps Script. Dilengkapi drag-and-drop file encoder ke base64.
                  </p>
                </div>
                <button 
                  onClick={() => copyToClipboard(INDEX_HTML_CONTENT, 'html')}
                  className="px-4 py-2 bg-[#003a70] text-white rounded-lg text-xs font-bold font-mono transition-all hover:bg-[#002244] flex items-center gap-2 cursor-pointer shadow-sm"
                >
                  {copiedHtml ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
                  {copiedHtml ? "Copied!" : "Copy Index.html Template"}
                </button>
              </div>

              <div className="relative">
                <div className="absolute top-3 right-3 bg-slate-900 border border-slate-800 rounded px-2.5 py-1 text-[10px] text-slate-400 font-mono">
                  HTML5 + BOOTSTRAP 5
                </div>
                <pre className="text-xs font-mono text-slate-100 bg-slate-950 p-5 rounded-xl overflow-x-auto max-h-[500px] leading-relaxed border border-slate-800">
                  {INDEX_HTML_CONTENT}
                </pre>
              </div>
            </div>

            <div className="bg-white border border-slate-200 p-6 rounded-2xl shadow-sm">
              <h4 className="font-bold text-slate-900 text-sm mb-3">Formulasional Pengiriman Array File Google Apps Script:</h4>
              <p className="text-xs text-slate-600 leading-relaxed mb-3 font-medium">
                Karena <code>google.script.run</code> tidak mendukung secara langsung pengiriman objek file html <code>&lt;input type="file"&gt;</code> ke server, kita menggunakan <code>FileReader</code> pada runtime browser untuk mengubah berkas menjadi string base64, lalu dikirim dalam bentuk dictionary payload:
              </p>
              <pre className="text-xs font-mono text-emerald-400 bg-slate-950 p-4 rounded-xl border border-slate-800 overflow-x-auto leading-relaxed">
{`const reader = new FileReader();
reader.onload = function(e) {
  const base64Index = e.target.result.indexOf(';base64,') + 8;
  const payload = {
    base64: e.target.result.substring(base64Index),
    mimeType: file.type,
    name: file.name
  };
  google.script.run.completeTaskAndUpload(nim, taskId, description, payload);
};
reader.readAsDataURL(file);`}
              </pre>
            </div>
          </div>
        )}

        {/* ===================== TAB 4: DATABASE SHEET SCHEMA ===================== */}
        {activeTab === 'db_schema' && (
          <div className="space-y-6">
            <div className="bg-white border border-slate-200 p-6 rounded-2xl shadow-sm">
              <h3 className="text-sm font-bold text-slate-900 mb-2 flex items-center gap-2">
                <Database className="w-5 h-5 text-[#003a70]" /> Kode SQL Supabase & Struktur Tabel
              </h3>
              <p className="text-xs text-slate-500 mb-6 font-medium">
                Jalankan kode SQL berikut di SQL Editor pada proyek Supabase Anda. Tabel <b>app_state</b> digunakan sebagai penyimpan konfigurasi parameter secara global (JSON), sementara sisa tabel relasional lainnya disediakan apabila Anda ingin mengembangkan backend API langsung pada PostgREST secara modular.
              </p>

              <div className="relative">
                <button 
                  onClick={() => {
                    navigator.clipboard.writeText(SUPABASE_SQL_CODE);
                    customAlert("Kode SQL berhasil disalin!", "success", "Tersalin");
                  }}
                  className="absolute top-4 right-4 px-3 py-1.5 bg-slate-800 text-slate-200 hover:bg-slate-700 rounded shadow-sm text-xs border border-slate-700 flex items-center gap-2 font-bold"
                >
                  <Copy className="w-3.5 h-3.5" /> Salin Code
                </button>
                <pre className="text-[11px] font-mono text-emerald-400 bg-slate-950 p-5 rounded-xl border border-slate-800 overflow-x-auto leading-relaxed h-[400px]">
{SUPABASE_SQL_CODE}
                </pre>
              </div>
            </div>

            {/* Google Drive and Folder Relations Map diagram */}
            <div className="bg-white border border-slate-200 p-6 rounded-2xl shadow-sm">
              <h4 className="font-bold text-slate-900 text-sm mb-4">Grafis Arsitektur Cloud Integrasi Google Drive & Template Kelulusan:</h4>
              <div className="p-6 border border-slate-100 rounded-2xl bg-slate-50/50 relative flex flex-col md:flex-row items-center justify-center gap-8 text-center text-xs">
                
                <div className="p-4 border border-slate-200 rounded-xl bg-white w-44 shadow-sm">
                  <div className="font-bold text-amber-600 uppercase">Input Bukti Dukung</div>
                  <div className="text-[11px] text-slate-500 mt-1 font-medium">NIM Mahasiswa melakukan submit file / screenshot di seluler</div>
                </div>

                <div className="text-[#003a70] text-lg flex items-center justify-center font-bold">
                  <ChevronRight className="w-6 h-6 rotate-90 md:rotate-0" />
                </div>

                <div className="p-4 border border-slate-200 rounded-xl bg-white w-44 shadow-sm">
                  <div className="font-bold text-[#003a70] uppercase">Script Convert</div>
                  <div className="text-[11px] text-slate-500 mt-1 font-medium">Decode base64 string di Code.gs dan panggil <code>folder.createFile()</code></div>
                </div>

                <div className="text-[#003a70] text-lg flex items-center justify-center font-bold">
                  <ChevronRight className="w-6 h-6 rotate-90 md:rotate-0" />
                </div>

                <div className="p-4 border border-slate-200 rounded-xl bg-white w-44 shadow-sm">
                  <div className="font-bold text-emerald-700 uppercase">Google Drive</div>
                  <div className="text-[11px] text-slate-500 mt-1 font-medium">Berkas tersimpan rapi dan mengembalikan URL download ke Sheets</div>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ===================== TAB 5: DEPLOYMENT GUIDE ===================== */}
        {activeTab === 'deploy_guide' && (
          <div className="space-y-6">
            <div className="bg-white border border-slate-200 p-6 rounded-2xl shadow-sm space-y-6 animate-fade-in">
              <div>
                <h3 className="text-sm font-bold text-slate-950 mb-2 flex items-center gap-2">
                  <BookOpen className="w-5 h-5 text-[#003a70]" /> Panduan Peluncuran Google Apps Script Web App
                </h3>
                <p className="text-xs text-slate-500 font-medium">
                  Ikuti panduan langkah demi langkah di bawah ini untuk meluncurkan sistem manajemen logbook & sertifikat PDF di workspace universitas Anda.
                </p>
              </div>

              {/* Step checklist */}
              <div className="space-y-4">
                
                {/* Step 1 */}
                <div className="flex gap-4 border border-slate-100 p-5 rounded-2xl bg-slate-50/50">
                  <span className="w-8 h-8 rounded-full bg-[#003a70] text-white text-xs font-bold flex items-center justify-center shrink-0">1</span>
                  <div>
                    <h4 className="font-bold text-slate-800 text-sm">Persiapan Google Sheets Database</h4>
                    <p className="text-xs text-slate-500 font-medium leading-relaxed mt-1">
                      Buat spreadsheet baru di Google Drive Anda. Di dalam menu bar Spreadsheet, klik tombol <strong className="text-slate-800">Extensions &gt; Apps Script</strong>. Ini akan membuka web editor khusus untuk Google Apps Script.
                    </p>
                  </div>
                </div>

                {/* Step 2 */}
                <div className="flex gap-4 border border-slate-100 p-5 rounded-2xl bg-slate-50/50">
                  <span className="w-8 h-8 rounded-full bg-[#003a70] text-white text-xs font-bold flex items-center justify-center shrink-0">2</span>
                  <div>
                    <h4 className="font-bold text-slate-800 text-sm">Salin Kode Code.gs & Index.html</h4>
                    <p className="text-xs text-slate-500 font-medium leading-relaxed mt-1">
                      Salin seluruh kode dari tab <strong className="text-[#003a70]">Code.gs</strong> Workspace ini dan timpa file default <code>Code.gs</code> di editor Google. Selanjutnya, ketuk ikon <strong className="text-slate-800">+ (Tambah berkas)</strong> di Apps Script, pilih berkas <strong className="text-slate-800">HTML</strong>, namai file tersebut <strong className="text-slate-800">Index</strong> (pemberian huruf besar kecil sensitif) dan salin kode template dari tab <strong className="text-[#003a70]">Index.html</strong>.
                    </p>
                  </div>
                </div>

                {/* Step 3 */}
                <div className="flex gap-4 border border-slate-100 p-5 rounded-2xl bg-slate-50/50">
                  <span className="w-8 h-8 rounded-full bg-[#003a70] text-white text-xs font-bold flex items-center justify-center shrink-0">3</span>
                  <div>
                    <h4 className="font-bold text-slate-800 text-sm">Pembuatan Folder & Berkas Google Suite Templates</h4>
                    <p className="text-xs text-slate-500 font-medium leading-relaxed mt-1">
                      Buat satu buah folder kosong di Google Drive Anda. Selanjutnya buat master template dokumen portfolio menggunakan Google Docs (berisi teks tag <code>{"<<Nama>>"}</code>, <code>{"<<NIM>>"}</code>, <code>{"<<Nilai Akhir>>"}</code>, dan tanda paragraf table <code>{"<<LogbookTable>>"}</code>) dan template design sertifikat menggunakan Google Slides. Catat masing-masing ID berkas/folder tersebut untuk dikonfigurasi di Settings Properties.
                    </p>
                  </div>
                </div>

                {/* Step 4 */}
                <div className="flex gap-4 border border-slate-100 p-5 rounded-2xl bg-slate-50/50">
                  <span className="w-8 h-8 rounded-full bg-[#003a70] text-white text-xs font-bold flex items-center justify-center shrink-0">4</span>
                  <div>
                    <h4 className="font-bold text-slate-800 text-sm">Rekam parameter di Script Properties</h4>
                    <p className="text-xs text-slate-500 font-medium leading-relaxed mt-1">
                      Di dalam halaman setting/proyek editor Apps Script Anda (ikon gerigi sebelah kiri), temukan bagian <strong className="text-slate-850">Script Properties</strong>. Masukkan tiga parameter berikut secara teliti:
                    </p>
                    <ul className="list-disc pl-5 mt-2 text-xs text-slate-500 font-medium space-y-1">
                      <li><strong>DRIVE_FOLDER_ID:</strong> ID Folder Drive penyimpanan rekap berkas.</li>
                      <li><strong>DOC_TEMPLATE_ID:</strong> ID Template Google Docs portfolio mahasiswa.</li>
                      <li><strong>SLIDE_TEMPLATE_ID:</strong> ID Template Google Slides sertifikat satu lembar.</li>
                    </ul>
                  </div>
                </div>

                {/* Step 5 */}
                <div className="flex gap-4 border border-slate-100 p-5 rounded-2xl bg-slate-50/50">
                  <span className="w-8 h-8 rounded-full bg-[#003a70] text-white text-xs font-bold flex items-center justify-center shrink-0">5</span>
                  <div>
                    <h4 className="font-bold text-slate-800 text-sm">Deploy Web App secara Publik</h4>
                    <p className="text-xs text-slate-500 font-medium leading-relaxed mt-1">
                      Tekan tombol <strong className="text-slate-850">Deploy &gt; New Deployment</strong> di bagian atas editor. Pilih jenis deployment <strong className="text-slate-850">Web App</strong>. Setel opsi "Execute as" sebagai <strong className="text-amber-600 font-bold">Me (akun Anda)</strong>, jalankan "Who has access" ke <strong className="text-slate-850">Anyone (Semua Orang)</strong>. Klik Deploy dan catat URL hasil deploy untuk diakses oleh ketua & anggota magang Anda via browser / smartphone mereka!
                    </p>
                  </div>
                </div>

              </div>
            </div>
          </div>
        )}

      </main>

      {/* ===================== SIMULATOR MODALS REPLICATED IN REACT ===================== */}
      
      {/* 1. Modal: Member completes task */}
      {selectedTaskToComplete && (
        <div className="fixed inset-0 bg-slate-950/80 flex items-center justify-center p-4 z-50 animate-fade-in overflow-y-auto">
          <form 
            onSubmit={handleSimulatedCompleteTask}
            className="bg-white text-slate-900 rounded-3xl shadow-2xl max-w-4xl w-full overflow-hidden flex flex-col my-8"
          >
            <div className="bg-[#003a70] text-white p-5 border-b border-slate-100 flex items-center justify-between">
              <div>
                <h4 className="font-extrabold text-sm uppercase tracking-wide">Penyelesaian Mandat & Isian Google Doc Laporan</h4>
                <p className="text-[10px] text-blue-200 mt-0.5">Sistem akan menyalin template Google Doc dari PIC dan membuat file baru di Drive</p>
              </div>
              <button 
                type="button" 
                onClick={() => setSelectedTaskToComplete(null)}
                className="text-white/60 hover:text-white transition-colors text-lg"
              >
                ✕
              </button>
            </div>

            <div className="p-6 grid grid-cols-1 lg:grid-cols-12 gap-6 overflow-y-auto max-h-[70vh]">
              {/* Left Column: Basic info & uploads */}
              <div className="lg:col-span-5 space-y-4">
                <div className="bg-slate-50 p-4 rounded-xl border border-slate-205/80 text-xs text-slate-600 space-y-1.5">
                  <div className="font-bold text-[#003a70] text-[11px] uppercase tracking-wide border-b pb-1 mb-1.5 flex items-center justify-between">
                    <span>Target Instruksi Kerja</span>
                    <span className={`text-[9px] font-extrabold px-1.5 py-0.5 rounded ${selectedTaskToComplete.workType === 'Kelompok' ? 'bg-purple-100 text-purple-800' : 'bg-emerald-100 text-emerald-800'}`}>
                      {selectedTaskToComplete.workType || 'Individu'}
                    </span>
                  </div>
                  <div><strong>Nama Aktivitas:</strong> {selectedTaskToComplete.taskName}</div>
                  <div><strong>Deskripsi Pokok:</strong> "{selectedTaskToComplete.description}"</div>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1.5">Uraian / Ringkasan Deskripsi Kerja (Sheets Logbook)</label>
                  <textarea 
                    className="w-full text-xs px-3 py-2 border border-slate-200 rounded-lg outline-none focus:border-[#003a70]"
                    rows={4}
                    required
                    placeholder="Uraikan laporan ringkas yang masuk ke Sheets..."
                    value={completeDesc}
                    onChange={(e) => setCompleteDesc(e.target.value)}
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1.5">Unggah Lampiran Bukti Kerja (Simulasi Drive)</label>
                  <div 
                    className="border-2 border-dashed border-slate-200 rounded-xl p-4 text-center cursor-pointer hover:bg-slate-50 transition bg-slate-50/50"
                    onClick={() => {
                      const mockNames = [
                        'config_sw_vokasi_done.png',
                        'rekap_vlan_testing.pdf',
                        'pendaftaran_modul_sc3.png',
                        'backup_database_itnetwork.sql'
                      ];
                      const randomName = mockNames[Math.floor(Math.random() * mockNames.length)];
                      setCompleteFileName(randomName);
                      setCompleteFileUrl('https://images.unsplash.com/photo-1551288049-bebda4e38f71?q=80&w=600');
                    }}
                  >
                    <span className="text-[10px] text-slate-400 block mb-1">Klik untuk mensimulasikan upload tangkapan layar</span>
                    <span className="text-xs font-bold text-sky-900 bg-sky-100 py-1 px-2.5 rounded inline-block mt-1">
                      {completeFileName || 'Ambil screenshoot bukti...'}
                    </span>
                  </div>
                </div>
              </div>

              {/* Right Column: Google Doc Template Fields */}
              <div className="lg:col-span-7 bg-blue-50/35 border border-[#003a70]/10 p-5 rounded-2xl space-y-4">
                <div className="flex items-center gap-1.5 pb-2 border-b border-[#003a70]/10 justify-between">
                  <div className="flex items-center gap-1.5">
                    <span className="w-2.5 h-2.5 bg-blue-600 rounded-full animate-pulse" />
                    <span className="text-xs font-bold text-slate-700 uppercase tracking-widest font-mono">Google Docs Editor Frame</span>
                  </div>
                  <span className="text-[9px] bg-blue-600 text-white font-mono font-bold px-2 py-0.5 rounded">Active Template ID</span>
                </div>

                <div className="space-y-3.5">
                  <div>
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-wide mb-1">Judul Dokumen Laporan Baru</label>
                    <input 
                      type="text" 
                      className="w-full text-xs font-bold px-3 py-2 border border-slate-250 bg-white rounded-lg outline-none focus:border-[#003a70]"
                      required
                      placeholder="Laporan Hasil Konfigurasi Router..."
                      value={docReportTitle}
                      onChange={(e) => setDocReportTitle(e.target.value)}
                    />
                  </div>

                  <div>
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-wide mb-1">I. Latar Belakang & Pembahasan Masalah</label>
                    <textarea 
                      className="w-full text-[11px] px-3 py-1.5 border border-slate-250 bg-white rounded-lg outline-none focus:border-[#003a70] leading-relaxed"
                      rows={2}
                      required
                      placeholder="Tulis latar belakang pengerjaan dan teori pendukung..."
                      value={docReportOverview}
                      onChange={(e) => setDocReportOverview(e.target.value)}
                    />
                  </div>

                  <div>
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-wide mb-1">II. Langkah-Langkah Pengerjaan Teknis (Template PIC)</label>
                    <textarea 
                      className="w-full text-[11px] font-mono px-3 py-1.5 border border-slate-250 bg-white rounded-lg outline-none focus:border-[#003a70] leading-relaxed"
                      rows={3}
                      required
                      placeholder="Langkah 1: ...&#10;Langkah 2: ..."
                      value={docReportSteps}
                      onChange={(e) => setDocReportSteps(e.target.value)}
                    />
                  </div>

                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-wide mb-1">III. Kendala & Hambatan</label>
                      <textarea 
                        className="w-full text-[11px] px-3 py-1.5 border border-slate-250 bg-white rounded-lg outline-none focus:border-[#003a70]"
                        rows={2}
                        required
                        placeholder="Contoh: Hambatan daya listrik..."
                        value={docReportChallenges}
                        onChange={(e) => setDocReportChallenges(e.target.value)}
                      />
                    </div>
                    <div>
                      <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-wide mb-1">IV. Kesimpulan Akhir</label>
                      <textarea 
                        className="w-full text-[11px] px-3 py-1.5 border border-slate-250 bg-white rounded-lg outline-none focus:border-[#003a70]"
                        rows={2}
                        required
                        placeholder="Semua fungsionalitas berjalan normal..."
                        value={docReportConclusion}
                        onChange={(e) => setDocReportConclusion(e.target.value)}
                      />
                    </div>
                  </div>
                </div>
              </div>
            </div>

            <div className="bg-slate-50 px-6 py-4 border-t border-slate-100 flex items-center justify-end gap-2.5">
              <button 
                type="button" 
                onClick={() => setSelectedTaskToComplete(null)}
                className="px-4 py-2 text-xs text-slate-600 rounded-lg hover:bg-slate-100 font-bold"
              >
                Batalkan
              </button>
              <button 
                type="submit" 
                disabled={isSimulatingCall}
                className="px-5 py-2 text-xs bg-[#003a70] hover:bg-[#002244] text-white font-extrabold rounded-lg transition-all shadow-sm disabled:opacity-50 disabled:cursor-not-allowed"
              >
                {isSimulatingCall ? "Menyimpan Dokumen..." : "Simpan & Upload Laporan Google Doc"}
              </button>
            </div>
          </form>
        </div>
      )}

      {/* 1.5. Modal: Predetermined Google Doc template duplicator with Daily Timeline Logs Table */}
      {false && selectedTaskToDuplicateDoc && (
        <div className="fixed inset-0 bg-slate-950/80 flex items-center justify-center p-4 z-50 animate-fade-in overflow-y-auto">
          <div className="bg-slate-100 text-slate-1000 rounded-3xl shadow-2xl max-w-6xl w-full overflow-hidden flex flex-col my-4">
            <div className="bg-[#003a70] text-white p-5 border-b border-slate-100 flex items-center justify-between shadow">
              <div>
                <h4 className="font-extrabold text-sm uppercase tracking-wide flex items-center gap-2">
                  <span className="bg-sky-600 text-white rounded-md px-2 py-0.5 text-[9.5px] font-mono">TEMPLATE MASTER PIC</span>
                  <span>Google Docs Template Duplicator & Editor</span>
                </h4>
                <p className="text-[10px] text-blue-200 mt-0.5 leading-relaxed">
                  Tim magang menyalin template dokumen dengan layout formal UNY, menyisipkan data tim, dan mencatat timeline harian.
                </p>
              </div>
              <button 
                type="button" 
                onClick={() => setSelectedTaskToDuplicateDoc(null)}
                className="text-white/60 hover:text-white transition-colors text-lg px-2"
              >
                ✕
              </button>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 p-5 overflow-y-auto max-h-[75vh]">
              
              {/* Left Side: Forms and dynamic daily table */}
              <div className="lg:col-span-6 space-y-4 max-h-[72vh] overflow-y-auto pr-1">
                
                {/* Form Section Header */}
                <div className="bg-white p-4.5 rounded-2xl border border-slate-200 shadow-sm space-y-3">
                  <h5 className="text-[11.5px] font-bold text-[#003a70] uppercase tracking-wider border-b pb-1.5 flex justify-between items-center">
                    <span>Metatags Template Dokumen [[DOC]]</span>
                    <span className="text-[10px] font-normal text-slate-400">Autofill dr akun & tugas</span>
                  </h5>
                  
                  <div className="space-y-2">
                    <div>
                      <label className="block text-[10px] uppercase font-extrabold text-slate-500 mb-1">Judul Laporan Resmi ([[JUDUL_LAPORAN]])</label>
                      <input 
                        type="text"
                        className="w-full text-xs font-bold px-3 py-2 border border-slate-300 rounded-lg outline-none focus:border-[#003a70] text-slate-800"
                        value={docReportTitle}
                        onChange={(e) => setDocReportTitle(e.target.value)}
                        required
                      />
                    </div>

                    <div>
                      <label className="block text-[10px] uppercase font-extrabold text-slate-500 mb-1">Unit Kerja / Instansi ([[UNIT_KERJA]])</label>
                      <input 
                        type="text"
                        className="w-full text-xs px-3 py-2 border border-slate-205 bg-slate-50 text-slate-400 rounded-lg cursor-not-allowed"
                        value="Fakultas Vokasi Universitas Negeri Yogyakarta"
                        disabled
                      />
                    </div>
                  </div>
                </div>

                {/* Daily Timeline Logs Interface */}
                <div className="bg-white p-4.5 rounded-2xl border border-slate-200 shadow-sm space-y-3.5">
                  <div>
                    <h5 className="text-[11.5px] font-bold text-[#003a70] uppercase tracking-wider flex justify-between items-center">
                      <span>Timeline Pengerjaan Harian</span>
                      <span className="text-[9.5px] bg-[#003a70] text-sky-100 rounded px-1.5 py-0.5">Automasi Tabel</span>
                    </h5>
                    <p className="text-[10px] text-slate-500 leading-relaxed mt-0.5">
                      PIC mewajibkan tim magang merinci progress pekerjaan per jam kerja secara bertahap ke dalam database.
                    </p>
                  </div>

                  {/* Inline Form to Add Log */}
                  <div className="bg-slate-50 hover:bg-slate-100/60 transition p-3 rounded-xl border border-slate-250/70 space-y-2">
                    <span className="text-[10px] font-extrabold text-slate-600 block uppercase tracking-wide">
                      + Input Log Aktivitas Baru
                    </span>
                    <div className="grid grid-cols-1 sm:grid-cols-12 gap-2">
                      <div className="sm:col-span-3">
                        <input 
                          type="text" 
                          placeholder="Tgl (e.g. 21 Juni)" 
                          className="w-full text-xs px-2.5 py-1.5 border border-slate-300 rounded bg-white text-slate-800 outline-none"
                          value={newTimelineDate}
                          onChange={(e) => setNewTimelineDate(e.target.value)}
                        />
                      </div>
                      <div className="sm:col-span-6">
                        <input 
                          type="text" 
                          placeholder="Deskripsi pengerjaan..." 
                          className="w-full text-xs px-2.5 py-1.5 border border-slate-300 rounded bg-white text-slate-800 outline-none"
                          value={newTimelineDesc}
                          onChange={(e) => setNewTimelineDesc(e.target.value)}
                        />
                      </div>
                      <div className="sm:col-span-3">
                        <input 
                          type="number" 
                          placeholder="Jam (e.g. 4)" 
                          className="w-full text-xs px-2.5 py-1.5 border border-slate-300 rounded bg-white text-slate-800 outline-none"
                          value={newTimelineHours}
                          onChange={(e) => setNewTimelineHours(e.target.value)}
                        />
                      </div>
                    </div>
                    <button 
                      type="button"
                      onClick={() => handleAddTimelineRow(selectedTaskToDuplicateDoc.taskId, newTimelineDate, newTimelineDesc, Number(newTimelineHours))}
                      className="w-full py-1.5 bg-[#003a70] hover:bg-[#002244] text-white rounded text-[11px] font-bold transition-all"
                    >
                      Tambahkan ke Tabel Google Doc Template
                    </button>
                  </div>

                  {/* Table View */}
                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-xs border border-slate-200 rounded-lg">
                      <thead>
                        <tr className="bg-slate-50 border-b border-slate-200 text-[#003a70] font-bold text-[10px] uppercase">
                          <th className="p-2 border-r border-slate-200">Tanggal</th>
                          <th className="p-2 border-r border-slate-200">Aktivitas Riil</th>
                          <th className="p-2 border-r border-slate-200 text-center">Durasi</th>
                          <th className="p-2 text-center">Aksi</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 text-[11px] text-slate-700">
                        {!selectedTaskToDuplicateDoc.timelineLogs || selectedTaskToDuplicateDoc.timelineLogs.length === 0 ? (
                          <tr>
                            <td colSpan={4} className="p-3 text-center text-slate-400 italic">
                              Belum ada entri timeline. Ketik pada form di bawah untuk menyisipkan baris tabel laporan harian.
                            </td>
                          </tr>
                        ) : (
                          selectedTaskToDuplicateDoc.timelineLogs.map((log, idx) => (
                            <tr key={idx} className="hover:bg-slate-50/50">
                              <td className="p-2 border-r border-slate-200 font-bold text-[#003a70]">{log.date}</td>
                              <td className="p-2 border-r border-slate-200">{log.description}</td>
                              <td className="p-2 border-r border-slate-200 text-center font-mono font-bold text-slate-700">{log.hours} Jam</td>
                              <td className="p-1 text-center">
                                <button 
                                  type="button"
                                  onClick={() => handleRemoveTimelineRow(selectedTaskToDuplicateDoc.taskId, idx)}
                                  className="text-red-600 hover:text-red-850 hover:underline px-2 py-1 font-bold text-[10px]"
                                >
                                  Hapus
                                </button>
                              </td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>

                {/* Discourse Textareas */}
                <div className="bg-white p-4.5 rounded-2xl border border-slate-200 shadow-sm space-y-3">
                  <h5 className="text-[11.5px] font-bold text-[#003a70] uppercase tracking-wider border-b pb-1">
                    Isi Pembahasan Utama & Resolusi
                  </h5>
                  
                  <div className="space-y-3">
                    <div>
                      <label className="block text-[10px] uppercase font-extrabold text-slate-500 mb-1">I. Pembahasan & Rekap Analisis teknis</label>
                      <textarea 
                        className="w-full text-xs px-3 py-2 border border-slate-300 rounded-lg outline-none focus:border-[#003a70] font-sans text-slate-800"
                        rows={3}
                        required
                        placeholder="Uraikan hasil analisis kelayakan alat, konfigurasi, atau testing yang telah berhasil dikerjakan..."
                        value={docReportOverview}
                        onChange={(e) => setDocReportOverview(e.target.value)}
                      />
                    </div>

                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <label className="block text-[10px] uppercase font-extrabold text-slate-500 mb-1">II. Kendala Lapangan</label>
                        <textarea 
                          className="w-full text-xs px-3 py-2 border border-slate-300 text-slate-800 rounded-lg outline-none focus:border-[#003a70]"
                          rows={2}
                          placeholder="Hambatan fisik di UNY..."
                          value={docReportChallenges}
                          onChange={(e) => setDocReportChallenges(e.target.value)}
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] uppercase font-extrabold text-slate-500 mb-1">III. Tindak Lanjut / Kesimpulan</label>
                        <textarea 
                          className="w-full text-xs px-3 py-2 border border-slate-300 text-slate-800 rounded-lg outline-none focus:border-[#003a70]"
                          rows={2}
                          placeholder="Status serah terima switch..."
                          value={docReportConclusion}
                          onChange={(e) => setDocReportConclusion(e.target.value)}
                        />
                      </div>
                    </div>
                  </div>
                </div>

              </div>

              {/* Right Side: Google Doc Sandbox Preview */}
              <div className="lg:col-span-6 flex flex-col bg-slate-200/50 p-4.5 rounded-2xl border border-slate-250">
                <span className="text-[10px] font-extrabold text-slate-500 uppercase tracking-wide mb-1.5 flex items-center justify-between">
                  <span>Pratinjau Live Duplikasi Google Doc Template</span>
                  <span className="text-[#003a70] font-mono px-2 py-0.5 bg-blue-100 rounded text-[9px] font-bold">STABILIZED WYSIWYG</span>
                </span>

                <div className="flex-1 bg-white border border-slate-300 rounded-2xl shadow p-6 min-h-[460px] font-serif text-slate-900 leading-relaxed text-xs relative overflow-y-auto max-h-[72vh] flex flex-col justify-between">
                  
                  {/* Watermark badge standard PIC UI */}
                  <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 text-slate-100/75 font-extrabold text-3xl rotate-12 select-none uppercase pointer-events-none tracking-widest text-center border-4 border-dashed border-slate-100 p-3 rounded-xl max-w-sm z-0">
                    TEMPLATED BY PIC<br/>DUPLICATED SYSTEM
                  </div>

                  <div className="relative z-10 space-y-4">
                    {/* Doc Header styling representation */}
                    <div className="border-b-2 border-[#003a70] pb-2 text-center">
                      <div className="text-[9.5px] uppercase font-sans font-bold tracking-widest text-[#003a70]">
                        KEMENTERIAN PENDIDIKAN, KEBUDAYAAN, RISET, DAN TEKNOLOGI
                      </div>
                      <div className="text-[10px] uppercase font-sans font-extrabold tracking-wide text-[#003a70]">
                        FAKULTAS VOKASI — UNIVERSITAS NEGERI YOGYAKARTA
                      </div>
                      <div className="text-[8px] font-sans text-slate-400 mt-0.5">
                        Alamat: Yogyakarta, Indonesia | Surel: vokasi@uny.ac.id
                      </div>
                    </div>

                    {/* Logo simulator / Title */}
                    <div className="text-center pt-1">
                      <div className="text-sky-900 bg-sky-50 font-sans border border-sky-100 px-3 py-1 text-[9px] rounded-full inline-block uppercase font-bold mb-3 tracking-wider">
                        [[TEMPLAT_DOKUMEN_SINKRON]]
                      </div>
                      <h4 className="text-sm font-sans font-extrabold text-slate-900 uppercase">
                        {docReportTitle || "[[LAPORAN_JUDUL_UTAMA]]"}
                      </h4>
                      <p className="text-[10px] font-sans text-slate-500 mt-1 font-medium italic">
                        Diajukan oleh: {activeUser?.name} ({activeUser?.nim}) — Peran: {activeUser?.role || 'Umum'}
                      </p>
                    </div>

                    {/* Meta section inside duplicated doc layout */}
                    <div className="bg-slate-50 border border-slate-200 p-2.5 rounded-xl text-[10px] space-y-1 font-sans text-slate-700 leading-normal">
                      <div><strong>Nama Project / Standardisasi:</strong> {selectedTaskToDuplicateDoc.taskName}</div>
                      <div><strong>Grup Kategori:</strong> {selectedTaskToDuplicateDoc.category}</div>
                      <div><strong>Sifat Pekerjaan:</strong> {selectedTaskToDuplicateDoc.workType || 'Standard'} (Terdata di Sheets)</div>
                      <div className="text-emerald-700 font-bold font-mono">✓ Live Drive ID: 1Doc_report_{selectedTaskToDuplicateDoc.taskId}...</div>
                    </div>

                    {/* Pembahasan */}
                    <div className="space-y-1">
                      <h6 className="font-sans font-bold text-[11px] text-slate-900 uppercase tracking-widest">
                        I. Pembahasan & Detail Hasil Kerja
                      </h6>
                      <p className="text-[11px] leading-relaxed text-slate-700 text-justify">
                        {docReportOverview || "Belum ada pembahasan yang dimasukkan. Silakan ketik rincian Pembahasan Utama di kolom formulir kiri."}
                      </p>
                    </div>

                    {/* Timeline logs injected within the template itself! */}
                    <div className="space-y-1 pt-1.5 font-sans">
                      <h6 className="font-bold text-[11px] text-slate-900 uppercase tracking-widest">
                        II. Tabel Rincian Jam Kerja Mandiri & Kelompok (Timeline Harian)
                      </h6>
                      <div className="overflow-x-auto">
                        <table className="w-full text-left text-[9.5px] border border-slate-300">
                          <thead>
                            <tr className="bg-slate-50 font-bold border-b border-slate-300">
                              <th className="p-1 text-slate-800 border-r border-slate-300">Tanggal</th>
                              <th className="p-1 text-slate-800 border-r border-slate-300">Aktivitas Teknis</th>
                              <th className="p-1 text-slate-800 text-center">Durasi Kerja</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-200 text-slate-650">
                            {!selectedTaskToDuplicateDoc.timelineLogs || selectedTaskToDuplicateDoc.timelineLogs.length === 0 ? (
                              <tr>
                                <td colSpan={3} className="p-2 text-center text-slate-400 italic">[[Table Empty: Harap isi log aktivitas harian]]</td>
                              </tr>
                            ) : (
                              selectedTaskToDuplicateDoc.timelineLogs.map((log, idx) => (
                                <tr key={idx}>
                                  <td className="p-1 border-r border-slate-300 font-bold text-sky-900">{log.date}</td>
                                  <td className="p-1 border-r border-slate-300">{log.description}</td>
                                  <td className="p-1 text-center font-bold text-slate-805">{log.hours} J</td>
                                </tr>
                              ))
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>

                    {/* Kendala */}
                    <div className="space-y-1">
                      <h6 className="font-sans font-bold text-[11px] text-slate-900 uppercase tracking-widest">
                        III. Tantangan Kelayakan Sistem
                      </h6>
                      <p className="text-[11px] text-slate-700 leading-relaxed text-left">
                        {docReportChallenges || "Tidak ada kendala sistemik yang berarti selama penugasan."}
                      </p>
                    </div>

                    {/* Penutup */}
                    <div className="space-y-1">
                      <h6 className="font-sans font-bold text-[11px] text-slate-900 uppercase tracking-widest">
                        IV. Tindak Lanjut & Validasi
                      </h6>
                      <p className="text-[11px] text-slate-700 leading-relaxed text-left">
                        {docReportConclusion || "Standardisasi tugas pokok sepenuhnya complete dan disiapkan untuk cetak kode portfolio QR."}
                      </p>
                    </div>

                  </div>

                  {/* Footnote of Duplicated Document Layout */}
                  <div className="border-t border-slate-200 pt-3 text-[10px] font-sans flex justify-between items-center text-slate-400 mt-6 select-none leading-none">
                    <span>Sistem Administrasi Kemahasiswaan UNY</span>
                    <span>Tahun {new Date().getFullYear()}</span>
                  </div>

                </div>
              </div>

            </div>

            {/* Modal Bottom Actions */}
            <div className="bg-slate-200/50 px-6 py-4.5 border-t border-slate-300 flex items-center justify-between gap-2.5">
              <div className="text-[11px] font-bold text-slate-500 font-sans flex items-center gap-1">
                <span>⚠️ Note:</span>
                <span className="font-normal text-slate-400 text-[10px]">Tugas langsung selesai 100% dan link otomatis disetor ke Sheets.</span>
              </div>
              <div className="flex items-center gap-2">
                <button 
                  type="button" 
                  onClick={() => setSelectedTaskToDuplicateDoc(null)}
                  className="px-4 py-2 text-xs text-slate-600 rounded-lg hover:bg-slate-200/80 font-bold transition"
                >
                  Batal / Kembali
                </button>
                <button 
                  type="button"
                  disabled={isSimulatingCall}
                  onClick={async () => {
                    if (!selectedTaskToDuplicateDoc) return;
                    const gasUrl = import.meta.env.VITE_GAS_WEB_APP_URL;
                    if (!gasUrl) {
                      customAlert('API backend Google Apps Script belum terhubung! Cek VITE_GAS_WEB_APP_URL di .env', 'error');
                      return;
                    }
                    if (!docTemplateInput) {
                      customAlert('ID Template Logbook (Google Docs) belum diisi di Pengaturan PIC!', 'error');
                      return;
                    }
                    setIsSimulatingCall(true);
                    try {
                      const payload = {
                        studentNim: activeUser?.nim || '',
                        taskId: selectedTaskToDuplicateDoc.taskId,
                        taskName: selectedTaskToDuplicateDoc.taskName,
                        category: selectedTaskToDuplicateDoc.category || '',
                        description: selectedTaskToDuplicateDoc.description || '',
                        docReportTitle: docReportTitle,
                        docReportOverview: docReportOverview,
                        docReportSteps: docReportSteps,
                        docReportChallenges: docReportChallenges,
                        docReportConclusion: docReportConclusion,
                        timelineLogs: selectedTaskToDuplicateDoc.timelineLogs || []
                      };
                      const res = await fetch(gasUrl, {
                        method: 'POST',
                        body: JSON.stringify({ action: 'generateTaskLogbook', payload })
                      });
                      const data = await res.json();
                      setIsSimulatingCall(false);
                      if (data.success) {
                        setTasks(tasks.map(t => t.taskId === selectedTaskToDuplicateDoc.taskId
                          ? { ...t, status: 'Completed', googleDocUrl: data.fileUrl }
                          : t
                        ));
                        setSelectedTaskToDuplicateDoc(null);
                        customAlert(`GDoc Logbook "${selectedTaskToDuplicateDoc.taskName}" berhasil dibuat! Kategori: ${selectedTaskToDuplicateDoc.category}\n\n${data.fileUrl}`, 'success', 'Logbook Berhasil Dibuat');
                      } else {
                        throw new Error(data.message);
                      }
                    } catch (err: any) {
                      setIsSimulatingCall(false);
                      customAlert('Gagal membuat GDoc Logbook: ' + err.message, 'error');
                    }
                  }}
                  className="px-6 py-2 text-xs bg-[#003a70] hover:bg-[#002244] text-white font-extrabold rounded-xl transition-all shadow flex items-center gap-1 disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  {isSimulatingCall ? '⏳ Memproses...' : '✓ Gandakan Laporan & Kirim Logbook'}
                </button>
              </div>
            </div>

          </div>
        </div>
      )}

      {/* 2. Modal: PIC Grades logbook */}
      {selectedLogbookToGrade && (
        <div className="fixed inset-0 bg-slate-950/80 flex items-center justify-center p-4 z-50">
          <form 
            onSubmit={handleSimulatedGrading}
            className="bg-white text-slate-900 rounded-2xl shadow-xl max-w-md w-full overflow-hidden"
          >
            <div className="bg-sky-950 text-white p-4 border-b border-slate-100 flex items-center justify-between">
              <h4 className="font-bold text-sm">Beri Penilaian Pekerjaan</h4>
              <button 
                type="button" 
                onClick={() => setSelectedLogbookToGrade(null)}
                className="text-white/60 hover:text-white"
              >
                ✕
              </button>
            </div>
            <div className="p-6 space-y-4">
              <div className="space-y-1 bg-slate-50 p-3 rounded-lg border border-slate-100 text-xs text-slate-600 leading-normal">
                <div><strong>Mahasiswa:</strong> {selectedLogbookToGrade.studentName}</div>
                <div><strong>Aktivitas:</strong> {selectedLogbookToGrade.taskName}</div>
                <div><strong>Lap. Teknis:</strong> "{selectedLogbookToGrade.workDescription}"</div>
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Masukkan Nilai Akhir (Skala 100)</label>
                <input 
                  type="number" 
                  className="w-full text-xs px-3 py-2 border border-slate-200 rounded-lg outline-none focus:border-sky-950 font-bold text-emerald-600 font-mono"
                  min="0"
                  max="100"
                  required
                  placeholder="Contoh: 88"
                  value={gradeInput}
                  onChange={(e) => setGradeInput(e.target.value)}
                />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Catatan Audit PIC (Feedback Lapangan)</label>
                <textarea 
                  className="w-full text-xs px-3 py-2 border border-slate-200 rounded-lg outline-none focus:border-sky-950"
                  rows={3}
                  placeholder="Bagus sekali, pemetaan VLAN sudah terstandarisasi..."
                  value={gradeNotes}
                  onChange={(e) => setGradeNotes(e.target.value)}
                />
              </div>
            </div>
            <div className="bg-slate-50 px-6 py-4 border-t border-slate-100 flex items-center justify-end gap-2.5">
              <button 
                type="button" 
                onClick={() => setSelectedLogbookToGrade(null)}
                className="px-3.5 py-1.5 text-xs text-slate-600 rounded-lg hover:bg-slate-100"
              >
                Batal
              </button>
              <button 
                type="submit" 
                className="px-4 py-1.5 text-xs bg-sky-950 text-white font-bold rounded-lg hover:bg-sky-900"
              >
                Simpan Transaksi Nilai
              </button>
            </div>
          </form>
        </div>
      )}

      {/* 3. Modal VISUAL DOC GENERATION (Certificate & Portfolio PDF outputs) */}
      {printDocument && (
        <>
          {/* Print CSS removed — Export PDF now uses a dedicated popup window */}
          
          {/* PRINT-ONLY HTML VERSION (CERTIFICATE & PORTFOLIO) */}
          <div className="hidden w-full max-w-[1056px] mx-auto text-slate-900 font-sans">
             {/* PAGE 1: CERTIFICATE — hidden, content rendered via popup Export PDF */}
             <div className="print-page w-full" style={{ aspectRatio: '16/9', border: '16px double rgba(180,120,0,0.3)', padding: '3rem', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', background: '#fffdf5', pageBreakAfter: 'always' }}>
               <div style={{ width: '100%', display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
                 <h1 style={{ fontFamily: 'Georgia,serif', fontSize: '2.5rem', color: '#92400e', letterSpacing: '0.2em', fontWeight: 'bold', marginBottom: '1rem', textAlign: 'center', textTransform: 'uppercase' }}>Sertifikat Penghargaan</h1>
                 <h3 style={{ fontFamily: 'monospace', fontSize: '0.85rem', color: '#64748b', marginBottom: '2.5rem' }}>NOMOR : {printDocument.nomorSurat || '-'}</h3>
                 <p style={{ fontSize: '0.85rem', color: '#475569', marginBottom: '1rem', fontFamily: 'Georgia,serif', textAlign: 'center' }}>Diberikan Khusus Kepada :</p>
                 <h2 style={{ fontFamily: 'Georgia,serif', fontSize: '3rem', color: '#003a70', fontWeight: 900, fontStyle: 'italic', marginBottom: '1.5rem', textAlign: 'center' }}>
                    {printDocument.studentName}
                 </h2>
                 <p style={{ maxWidth: '60%', textAlign: 'center', fontSize: '0.85rem', fontFamily: 'Georgia,serif', lineHeight: 2, color: '#374151' }}>
                   Telah melaksanakan praktik kerja (magang) berstandar project-based learning di Fakultas Vokasi Universitas Negeri Yogyakarta dan dinyatakan <strong>LULUS</strong> dalam kurun waktu <b>{printDocument.periode || 3} bulan</b> aktif ({formatIdDate(printDocument.tanggalMulai)} s/d {formatIdDate(printDocument.tanggalSelesai)}), dengan menorehkan klasifikasi predikat kelayakan performa:
                 </p>
                 <div style={{ marginTop: '2rem', padding: '1rem 2rem', border: '2px solid rgba(180,120,0,0.2)', textAlign: 'center', borderRadius: '0.75rem', background: 'white', display: 'inline-block' }}>
                    <div style={{ fontSize: '2rem', fontWeight: 800, color: '#d97706', textTransform: 'uppercase', letterSpacing: '0.1em' }}>{getStudentGradesInfo(printDocument.studentNim).predicate}</div>
                    <div style={{ fontFamily: 'monospace', fontSize: '0.75rem', color: '#64748b', marginTop: '0.5rem' }}>NILAI CAPAIAN RATA-RATA: {printDocument.overallGrade} / 100</div>
                 </div>
               </div>
             </div>

             {/* PAGE 2: PORTFOLIO */}
             <div className="print-page w-full p-12 bg-white">
                <h2 className="text-2xl font-bold text-[#003a70] text-center border-b-2 border-slate-200 pb-4 mb-4">PORTOFOLIO REKAPITULASI PROGRAM MAGANG<br/><span className="text-sm text-slate-500 font-normal mt-2 inline-block">Fakultas Vokasi Universitas Negeri Yogyakarta</span></h2>
                
                <div className="flex justify-between items-center bg-slate-50 p-6 rounded-2xl border border-slate-200 mb-8 mt-6">
                    <div className="space-y-2 text-sm text-slate-700">
                       <div><strong>Nama Mahasiswa:</strong> <span className="font-bold text-slate-900">{printDocument.studentName}</span></div>
                       <div><strong>Nomor KTM / NIM:</strong> <span className="font-mono">{printDocument.studentNim}</span></div>
                       <div><strong>Periode Praktik:</strong> <span className="font-bold">{printDocument.periode || 3} Bulan</span> ({formatIdDate(printDocument.tanggalMulai)} s/d {formatIdDate(printDocument.tanggalSelesai)})</div>
                    </div>
                    <div className="text-right border-l-2 border-slate-200 pl-6 space-y-1">
                       <div className="text-xs uppercase text-slate-500 font-bold tracking-wider">Nilai Akhir (Predikat)</div>
                       <div className="text-2xl font-extrabold text-emerald-600">{printDocument.overallGrade} <span className="text-sm font-normal text-emerald-700">/ 100</span></div>
                       <div className="font-bold text-slate-800">{getStudentGradesInfo(printDocument.studentNim).predicate}</div>
                    </div>
                </div>

                <h3 className="font-bold text-slate-800 mb-4 text-xs uppercase tracking-wider">Aktivitas & Pembuktian Rekayasa Kerja</h3>
                <table className="w-full text-left text-sm border-collapse border-y-2 border-slate-800">
                   <thead>
                     <tr className="bg-slate-100/50">
                       <th className="py-3 px-4 border-b border-slate-200 font-bold w-12 text-center text-slate-600 uppercase text-[10px]">No</th>
                       <th className="py-3 px-4 border-b border-slate-200 font-bold text-slate-600 uppercase text-[10px]">Aktivitas / Laporan Ops</th>
                       <th className="py-3 px-4 border-b border-slate-200 font-bold text-center w-32 border-l border-slate-200 text-slate-600 uppercase text-[10px]">QR Dokumen</th>
                     </tr>
                   </thead>
                   <tbody>
                     {printDocument.logs.map((l, i) => (
                        <tr key={i} className="border-b border-slate-200">
                          <td className="py-4 px-4 text-center font-mono text-slate-500 font-bold">{i+1}</td>
                          <td className="py-4 px-4">
                             <div className="font-bold text-slate-800 text-sm">{l.name}</div>
                             <div className="text-[10px] text-slate-500 mt-1 uppercase font-bold bg-slate-200 px-2 py-0.5 inline-block rounded">{l.workType || 'Individu'}</div>
                          </td>
                          <td className="py-4 px-4 text-center border-l border-slate-200">
                             <img src={`https://api.qrserver.com/v1/create-qr-code/?size=100x100&data=${encodeURIComponent(l.googleDocUrl || 'https://docs.google.com')}`} alt="QR" className="w-14 h-14 mx-auto" />
                          </td>
                        </tr>
                     ))}
                     {printDocument.logs.length === 0 && (
                        <tr>
                           <td colSpan={3} className="py-8 text-center text-slate-400 italic">Belum ada portofolio divalidasi.</td>
                        </tr>
                     )}
                   </tbody>
                </table>
                <div className="mt-12 text-right">
                  <div className="inline-block text-center mr-8">
                     <p className="text-xs text-slate-500 mb-16">Yogyakarta, {formatIdDate(printDocument.tanggalSelesai)}<br/>Ketua Praktik Magang,</p>
                     <p className="text-sm font-bold text-slate-800 underline">Prof. Dr. Siswayanto, M.Pd., M.T.</p>
                     <p className="text-xs text-slate-500">NIP. 197410262002101002</p>
                  </div>
                </div>
             </div>
          </div>

          <div className="fixed inset-0 bg-slate-950/90 overflow-y-auto p-6 z-50 flex flex-col justify-center items-center no-print">
            <div className="max-w-4xl w-full bg-slate-100 text-slate-900 rounded-3xl overflow-hidden shadow-2xl flex flex-col">
            
            {/* Modal Header */}
            <header className="bg-sky-950 text-white border-b-4 border-amber-400 p-5 flex items-center justify-between">
              <div>
                <h3 className="font-extrabold text-sm text-amber-400">DOKUMEN HASIL KELULUSAN MAGANG (MOCK PREVIEW)</h3>
                <p className="text-[10px] text-sky-200">Berhasil ditransformasikan secara digital via Slide & Doc API Google Web Service</p>
              </div>
              <button 
                onClick={() => setPrintDocument(null)}
                className="bg-slate-800 text-red-400 px-3 py-1.5 rounded-lg text-xs font-bold"
              >
                Tutup Tinjauan
              </button>
            </header>

            {/* Split layout: Certificate & Portfolio list */}
            <div className="p-6 grid grid-cols-1 lg:grid-cols-2 gap-6 max-h-[600px] overflow-y-auto">
              
              {/* E-Certificate Preview (Slides layout) */}
              <div className="flex flex-col gap-4">
                <div className="relative aspect-[16/10] w-full border border-slate-200 shadow-sm flex flex-col overflow-hidden bg-slate-200 min-h-[460px] rounded-xl group">
                   <iframe 
                     src={`https://docs.google.com/presentation/d/${slideTemplateInput}/embed?rm=minimal`} 
                     className="absolute inset-0 w-full h-full border-0 z-0 bg-white" 
                     title="Certificate Google Slide Template" 
                   />
                   <div className="absolute top-2 right-2 bg-[#003a70]/90 text-white text-[9px] px-2 py-1 rounded font-mono shadow-md backdrop-blur-sm pointer-events-none z-10 transition-opacity opacity-100 group-hover:opacity-0">Tinjauan Template ID Slide</div>
                   <div className="absolute bottom-0 left-0 right-0 bg-slate-900/90 text-white p-3 backdrop-blur-sm shadow-[0_-4px_10px_rgba(0,0,0,0.1)] z-10 transition-transform transform translate-y-0 group-hover:translate-y-full pointer-events-none">
                     <div className="text-[10px] uppercase font-bold text-slate-200 mb-1 border-b border-slate-700 pb-1 flex justify-between">
                       <span>Pemetaan Variabel Slide (Otomatis)</span>
                       <span className="text-[9px] text-slate-500 normal-case font-normal">Arahkan kursor untuk melihat slide penuh</span>
                     </div>
                     <div className="grid grid-cols-2 gap-x-4 gap-y-1 text-[8.5px] font-mono">
                       <div className="flex justify-between items-center"><span className="text-amber-400 font-bold">{"{{NAMA}}"}</span> <span className="font-semibold flex-1 text-right truncate ml-2">{printDocument.studentName}</span></div>
                       <div className="flex justify-between items-center"><span className="text-amber-400 font-bold">{"{{NOMOR}}"}</span> <span className={`${printDocument.nomorSurat ? 'text-emerald-300' : 'text-slate-400'} font-bold flex-1 text-right truncate ml-2`}>{printDocument.nomorSurat || '- (Belum diatur)'}</span></div>
                       <div className="flex justify-between items-center"><span className="text-amber-400 font-bold">{"{{PERIODE}}"}</span> <span className="font-semibold flex-1 text-right truncate ml-2">{printDocument.periode || 3}</span></div>
                       <div className="flex justify-between items-center"><span className="text-amber-400 font-bold">{"{{PREDIKAT}}"}</span> <span className="font-semibold flex-1 text-right truncate ml-2">{getStudentGradesInfo(printDocument.studentNim).predicate}</span></div>
                       <div className="flex justify-between items-center"><span className="text-amber-400 font-bold">{"{{MULAI}}"}</span> <span className="font-semibold flex-1 text-right truncate ml-2">{formatIdDate(printDocument.tanggalMulai || '2026-03-01')}</span></div>
                       <div className="flex justify-between items-center"><span className="text-amber-400 font-bold">{"{{SELESAI}}"}</span> <span className="font-semibold flex-1 text-right truncate ml-2">{formatIdDate(printDocument.tanggalSelesai || '2026-06-01')}</span></div>
                     </div>
                   </div>
                </div>

                <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm mb-auto">
                   <label className="block text-[11px] font-bold text-slate-700 mb-1.5 uppercase tracking-wide">Pembaruan Nomor Surat Cepat</label>
                   <div className="flex flex-col sm:flex-row gap-2">
                     <input 
                       type="text" 
                       className="flex-1 text-xs px-3.5 py-2 border border-slate-200 rounded-lg outline-none focus:border-[#003a70] focus:ring-1 focus:ring-[#003a70]/10 transition-all font-mono font-medium text-slate-800"
                       placeholder="Contoh: MAGANG/FV-UNY/21532/1"
                       value={certQuickEditNomorSurat}
                       onChange={(e) => setCertQuickEditNomorSurat(e.target.value)}
                     />
                     <button 
                       type="button"
                       onClick={handleSaveQuickNomorSurat}
                       className="px-4 py-2 bg-[#003a70] text-white rounded-lg text-xs font-bold hover:bg-[#002244] transition-colors shrink-0"
                     >
                       Simpan Nomor
                     </button>
                   </div>
                </div>
              </div>

              {/* Comprehensive Portfolio (Docs table list) */}
              <div id="doc-print-container" className="bg-white print:border-none print:shadow-none print:rounded-none print:p-0 print:m-0 border border-slate-200 p-6 rounded-2xl shadow flex flex-col justify-between min-h-[500px]">
                <div>
                  <div className="text-center font-bold text-xs uppercase border-b pb-2 mb-4 leading-normal text-slate-800">
                    PORTOFOLIO REKAPITULASI PROGRAM MAGANG <br />
                    <span className="text-[10px] text-slate-400 normal-case font-normal">Fakultas Vokasi Universitas Negeri Yogyakarta</span>
                  </div>

                  <div className="text-xs space-y-1 text-slate-600 mb-4 bg-slate-50 p-3.5 rounded-xl border border-slate-100 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
                    <div className="space-y-0.5">
                      <div><strong>Nama Mahasiswa:</strong> <span className="font-bold text-slate-800">{printDocument.studentName}</span></div>
                      <div><strong>Nomor KTM / NIM:</strong> <span className="font-mono font-semibold text-slate-700">{printDocument.studentNim}</span></div>
                      <div><strong>Periode Magang:</strong> <span className="font-bold text-sky-950">{printDocument.periode || 3} Bulan</span> ({formatIdDate(printDocument.tanggalMulai)} s/d {formatIdDate(printDocument.tanggalSelesai)})</div>
                    </div>
                    <div className="p-2 bg-emerald-50 border border-emerald-150 rounded-lg text-right font-sans shrink-0">
                      <div className="text-[8px] uppercase font-bold text-emerald-800 leading-none">Rerata Nilai</div>
                      <div className="text-sm font-extrabold text-emerald-700 leading-none mt-0.5">{printDocument.overallGrade} / 100.00</div>
                      <div className="text-[8px] font-bold text-slate-500 mt-1 uppercase max-w-[100px] truncate">{getStudentGradesInfo(printDocument.studentNim).predicate}</div>
                    </div>
                  </div>

                  {/* Dynamic Category Donut Chart for professional competencies */}
                  <div className="mb-4">
                    {(() => {
                      const distribution = getCategoryPercentages(printDocument.studentNim);
                      const colors = ['#003a70', '#eab308', '#10b981', '#a855f7', '#ec4899', '#3b82f6'];
                      let cumulativePercent = 0;
                      return (
                        <div className="bg-slate-50/50 p-3 rounded-xl border border-slate-100 flex items-center gap-4">
                          {/* Left: SVG donut */}
                          {distribution.length > 0 ? (
                            <div className="relative w-16 h-16 shrink-0 select-none">
                              <svg viewBox="0 0 36 36" className="w-full h-full transform -rotate-90">
                                <circle cx="18" cy="18" r="15.915" fill="none" stroke="#f1f5f9" strokeWidth="3" />
                                {distribution.map((item, idx) => {
                                  const pct = item.percentage;
                                  const strokeDash = `${pct} ${100 - pct}`;
                                  const strokeOffset = 100 - cumulativePercent;
                                  cumulativePercent += pct;
                                  return (
                                    <circle
                                      key={item.category}
                                      cx="18"
                                      cy="18"
                                      r="15.915"
                                      fill="none"
                                      stroke={colors[idx % colors.length]}
                                      strokeWidth="4"
                                      strokeDasharray={strokeDash}
                                      strokeDashoffset={strokeOffset}
                                      className="transition-all duration-300"
                                    />
                                  );
                                })}
                              </svg>
                              <div className="absolute inset-0 flex flex-col items-center justify-center">
                                <span className="text-[7px] font-bold text-slate-400 uppercase tracking-widest leading-none">Kompeten</span>
                                <span className="text-[10px] font-extrabold text-slate-800 leading-none mt-0.5">{distribution[0]?.percentage}%</span>
                              </div>
                            </div>
                          ) : (
                            <div className="w-14 h-14 border rounded bg-slate-50 flex items-center justify-center text-[8px] text-slate-400 italic">No Logs</div>
                          )}

                          {/* Right: Legend */}
                          <div className="flex-1 space-y-1">
                            <div className="text-[9px] font-extrabold text-[#003a70] uppercase tracking-wider mb-1">
                              Kompetensi Keahlian Profesi Berdasarkan Pekerjaan :
                            </div>
                            {distribution.length === 0 ? (
                              <div className="text-[9px] text-slate-400 italic">Belum ada penyebaran kategori kompetensi terekam.</div>
                            ) : (
                              <div className="grid grid-cols-2 gap-x-2 gap-y-1">
                                {distribution.map((item, idx) => (
                                  <div key={item.category} className="flex items-center justify-between text-[8.5px] font-bold text-slate-700">
                                    <div className="flex items-center gap-1 truncate max-w-[120px] uppercase text-[7.5px]">
                                      <span className="w-1.5 h-1.5 rounded-full inline-block shrink-0" style={{ backgroundColor: colors[idx % colors.length] }} />
                                      <span className="truncate" title={item.category}>{item.category}</span>
                                    </div>
                                    <span className="text-slate-500 font-mono text-[8px] pr-0.5">{item.percentage}%</span>
                                  </div>
                                ))}
                              </div>
                            )}
                          </div>
                        </div>
                      );
                    })()}
                  </div>

                  <div className="text-xs font-bold text-slate-700 mb-2 uppercase tracking-wide text-[10px]">Tabel Daftar Pembuktian Kerja & QR Akses Laporan Docs:</div>
                  <div className="overflow-x-auto max-h-[250px]">
                    <table className="w-full text-left text-[10px] border-collapse">
                      <thead>
                        <tr className="bg-[#003a70] text-white">
                          <th className="p-2">No</th>
                          <th className="p-2">Aktivitas Kerja</th>
                          <th className="p-2 text-center">Nilai</th>
                          <th className="p-2 text-center">QR Akses Google Doc</th>
                        </tr>
                      </thead>
                      <tbody>
                        {printDocument.logs.length === 0 ? (
                          <tr>
                            <td colSpan={4} className="py-4 text-center text-slate-400 italic">Belum ada portofolio dinilai lengkap.</td>
                          </tr>
                        ) : (
                          printDocument.logs.map((l, i) => {
                            const docUrl = l.googleDocUrl || `https://docs.google.com/document/d/1Doc_mock_${i}/edit`;
                            const qrCodeUrl = `https://api.qrserver.com/v1/create-qr-code/?size=60x60&data=${encodeURIComponent(docUrl)}`;
                            return (
                              <tr key={i} className="border-b border-slate-100 hover:bg-slate-50/50">
                                <td className="p-2 font-mono font-bold text-slate-500">{i+1}</td>
                                <td className="p-2">
                                  <div className="font-bold text-slate-800">{l.name}</div>
                                  <div className="flex gap-1.5 mt-0.5">
                                    <span className="text-[8px] bg-slate-100 font-mono px-1 rounded">{l.category}</span>
                                    <span className="text-[8px] bg-purple-50 text-purple-700 px-1 rounded font-bold">{l.workType || 'Individu'}</span>
                                  </div>
                                </td>
                                <td className="p-2 text-center font-extrabold text-[#003a70] text-xs">{l.grade || '-'}</td>
                                <td className="p-2 text-center flex flex-col items-center justify-center">
                                  <a href={docUrl} target="_blank" rel="noreferrer" title="Klik untuk membuka Google Doc" className="block p-0.5 border border-slate-200 rounded hover:border-[#003a70] transition-colors bg-white">
                                    <img 
                                      src={qrCodeUrl} 
                                      alt="QR Code" 
                                      className="w-10 h-10 object-contain"
                                      referrerPolicy="no-referrer"
                                    />
                                  </a>
                                  <span className="text-[7.5px] text-slate-400 mt-0.5 block font-mono">Scan QR / Klik</span>
                                </td>
                              </tr>
                            );
                          })
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>

                <div className="border-t pt-2.5 mt-3 text-[10px] text-slate-400 flex items-center justify-between">
                  <span>Template ID Doc: {docTemplateInput || 'uny_rekap_db_2026'}</span>
                  <span className="font-extrabold text-emerald-600 font-mono bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">PORTFOLIO PDF READY</span>
                </div>
              </div>

            </div>

            {/* Email send mock notification block */}
            <div className="p-4 bg-[#003a70] text-white flex flex-col sm:flex-row items-center justify-between px-6 border-t border-[#002244] gap-4">
              <div className="flex items-center gap-2 text-xs font-mono">
                <Mail className="w-4 h-4 text-sky-400" />
                <span>Siap dikirimkan sebagai lampiran PDF ke email: <strong>{printDocument.email}</strong></span>
              </div>
              <div className="flex items-center gap-2 flex-wrap">
                <button 
                  type="button"
                  onClick={() => {
                    customAlert(`Aksi berhasil disimulasikan: Berkas Sertifikat.pdf dan Portofolio (Desain Estetik) telah dikirim ke alamat email ${printDocument.email}`, 'success', 'Simulasi Kirim Email');
                  }}
                  className="flex items-center gap-1.5 bg-sky-600 hover:bg-sky-500 text-white px-4 py-1.5 rounded-lg text-xs font-bold transition-all shadow-md border border-sky-400"
                >
                  <Send className="w-4 h-4" />
                  Kirim Email (PDF)
                </button>
                <button 
                  type="button"
                  onClick={async () => {
                    if (!printDocument) return;
                    
                    const gasUrl = import.meta.env.VITE_GAS_WEB_APP_URL;
                    if (!gasUrl) {
                      customAlert("Mohon maaf, API backend Google Apps Script belum terhubung!\\n\\nSilakan cek tab Terminal/Log dan ikuti instruksi Deploy Code.gs sebagai Web App, lalu masukkan URL-nya ke .env pada variabel VITE_GAS_WEB_APP_URL.", "error", "Konfigurasi Backend Terputus");
                      return;
                    }
                    setIsSimulatingCall(true);
                    setSimulationMsg(`[GOOGLE APPS SCRIPT] Memanggil API untuk memproses template asli Google Slides & Docs untuk siswa ${printDocument.studentName}...`);
                    
                    try {
                      // Prepare the full data payload from React/Supabase to bypass Google Sheets completely
                      const payload = {
                        studentData: {
                          name: printDocument.studentName,
                          nim: printDocument.studentNim.toString(),
                          email: printDocument.email || '',
                          role: printDocument.role || 'Anggota',
                          periode: printDocument.periode || 3,
                          tanggalMulai: printDocument.tanggalMulai || '',
                          tanggalSelesai: printDocument.tanggalSelesai || '',
                          nomorSurat: printDocument.nomorSurat || '-',
                          overallGrade: printDocument.overallGrade
                        },
                        logbooks: printDocument.logs,
                        driveId: driveIdInput || import.meta.env.VITE_DRIVE_FOLDER_ID || '',
                        portfolioId: portfolioTemplateInput || import.meta.env.VITE_TEMPLATE_PORTFOLIO_ID || '',
                        slideId: slideTemplateInput || import.meta.env.VITE_TEMPLATE_M_SLIDES_ID || ''
                      };

                      // Send a simple POST request (text/plain body bypasses CORS preflight)
                      const res = await fetch(gasUrl, {
                        method: 'POST',
                        body: JSON.stringify({
                          action: 'export_certificate',
                          payload: payload
                        })
                      });
                      
                      const data = await res.json();
                      setIsSimulatingCall(false);
                      
                      if (data.success) {
                        customAlert(`[BERHASIL] Sertifikat dan Portofolio telah diunduh ke komputer Anda.`, 'success', 'Export Selesai');
                        
                        // 1. Download Sertifikat PDF dari Base64 secara otomatis
                        if (data.certBase64) {
                          const link = document.createElement('a');
                          link.href = `data:application/pdf;base64,${data.certBase64}`;
                          link.download = `Sertifikat_${printDocument.studentName.replace(/ /g, '_')}.pdf`;
                          document.body.appendChild(link);
                          link.click();
                          document.body.removeChild(link);
                        }
                        
                        // 2. Download Portofolio PDF dari Base64 secara otomatis (Dihapus, diganti Print Browser)
                        
                        // Memanggil window.print() untuk Portofolio Estetik
                        const originalDisplay = document.getElementById('root')?.style.display;
                        if(document.getElementById('root')) document.getElementById('root')!.style.display = 'none';
                        
                        const printDiv = document.createElement('div');
                        printDiv.id = 'print-section';
                        printDiv.className = 'w-full bg-white';
                        printDiv.innerHTML = document.getElementById('doc-print-container')?.outerHTML || '';
                        document.body.appendChild(printDiv);
                        
                        setTimeout(() => {
                          window.print();
                          document.body.removeChild(printDiv);
                          if(document.getElementById('root')) document.getElementById('root')!.style.display = originalDisplay || '';
                        }, 500);
                        
                      } else {
                        const errMsg = data.message || (data.error ? JSON.stringify(data.error) : JSON.stringify(data));
                        customAlert("Backend gagal memproses Sertifikat: " + errMsg, "error", "Kesalahan Sistem GAS");
                      }
                    } catch (error: any) {
                      setIsSimulatingCall(false);
                      customAlert("Gagal terhubung ke Google Apps Script: " + error.message, "error", "Kesalahan Jaringan / CORS");
                    }
                  }}
                  className="flex items-center gap-1.5 bg-amber-500 hover:bg-amber-400 text-sky-950 px-4 py-1.5 rounded-lg text-xs font-bold transition-all shadow-md"
                >
                  <Download className="w-4 h-4" />
                   Export PDF (Sertifikat & Portofolio)
                </button>
              </div>
            </div>

          </div>
        </div>
        </>
      )}

      {/* 5. Modal: Edit User Profil & Kredensial */}
      {editingUser && (
        <div className="fixed inset-0 bg-slate-950/80 flex items-center justify-center p-4 z-50 animate-fade-in overflow-y-auto">
          <form 
            onSubmit={handleSaveEditedUser}
            className="bg-white text-slate-900 rounded-3xl shadow-2xl max-w-md w-full overflow-hidden flex flex-col my-8"
          >
            <div className="bg-[#003a70] text-white p-5 border-b border-slate-100 flex items-center justify-between">
              <div>
                <h4 className="font-extrabold text-sm uppercase tracking-wide">Edit Detail Pengguna</h4>
                <p className="text-[10px] text-blue-200 mt-0.5">Ubah kredensial login, email, nama, dan NIM di Sheets</p>
              </div>
              <button 
                type="button" 
                onClick={() => setEditingUser(null)}
                className="text-white/60 hover:text-white transition-colors text-lg"
              >
                ✕
              </button>
            </div>

            <div className="p-6 space-y-4 overflow-y-auto max-h-[70vh]">
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1.5">Nama Lengkap Pengguna</label>
                <input 
                  type="text" 
                  className="w-full text-xs px-3.5 py-2 border border-slate-200 rounded-lg outline-none focus:border-[#003a70] focus:ring-1 focus:ring-[#003a70]/10 transition-all font-medium text-slate-800"
                  placeholder="Ketik Nama"
                  value={editUserName}
                  onChange={(e) => setEditUserName(e.target.value)}
                  required
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1.5">Nomor Induk / NIM / NIDN / NIS</label>
                <input 
                  type="text" 
                  className="w-full text-xs px-3.5 py-2 border border-slate-200 rounded-lg outline-none focus:border-[#003a70] focus:ring-1 focus:ring-[#003a70]/10 transition-all font-mono font-medium text-slate-800"
                  placeholder="Ketik NIM baru"
                  value={editUserNim}
                  onChange={(e) => setEditUserNim(e.target.value)}
                  required
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1.5">Alamat Surat Elektronik (Email)</label>
                <input 
                  type="email" 
                  className="w-full text-xs px-3.5 py-2 border border-slate-200 rounded-lg outline-none focus:border-[#003a70] focus:ring-1 focus:ring-[#003a70]/10 transition-all font-medium text-slate-800"
                  placeholder="Ketik Email"
                  value={editUserEmail}
                  onChange={(e) => setEditUserEmail(e.target.value)}
                  required
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1.5">Kata Sandi / Password Baru</label>
                <input 
                  type="text" 
                  className="w-full text-xs px-3.5 py-2 border border-slate-200 rounded-lg outline-none focus:border-[#003a70] focus:ring-1 focus:ring-[#003a70]/10 transition-all text-slate-800"
                  placeholder="Ketik Password baru"
                  value={editUserPassword}
                  onChange={(e) => setEditUserPassword(e.target.value)}
                  required
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1.5">Periode Magang (Bulan)</label>
                <input 
                  type="number" 
                  min="1"
                  max="12"
                  className="w-full text-xs px-3.5 py-2 border border-slate-200 rounded-lg outline-none focus:border-[#003a70] focus:ring-1 focus:ring-[#003a70]/10 transition-all font-medium text-slate-800"
                  placeholder="Contoh: 3"
                  value={editUserPeriode}
                  onChange={(e) => setEditUserPeriode(e.target.value)}
                  required
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1.5">Tanggal Mulai</label>
                  <input 
                    type="date" 
                    className="w-full text-xs px-3.5 py-2 border border-slate-200 rounded-lg outline-none focus:border-[#003a70] focus:ring-1 focus:ring-[#003a70]/10 transition-all text-slate-800 font-mono font-medium"
                    value={editUserTanggalMulai}
                    onChange={(e) => setEditUserTanggalMulai(e.target.value)}
                    required
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1.5">Tanggal Selesai</label>
                  <input 
                    type="date" 
                    className="w-full text-xs px-3.5 py-2 border border-slate-200 rounded-lg outline-none focus:border-[#003a70] focus:ring-1 focus:ring-[#003a70]/10 transition-all text-slate-800 font-mono font-medium"
                    value={editUserTanggalSelesai}
                    onChange={(e) => setEditUserTanggalSelesai(e.target.value)}
                    required
                  />
                </div>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1.5">Nomor Surat Sertifikat</label>
                <input 
                  type="text" 
                  className="w-full text-xs px-3.5 py-2 border border-slate-200 rounded-lg outline-none focus:border-[#003a70] focus:ring-1 focus:ring-[#003a70]/10 transition-all font-mono font-medium text-slate-800"
                  placeholder="Contoh: MAGANG/FV-UNY/21532/1"
                  value={editUserNomorSurat}
                  onChange={(e) => setEditUserNomorSurat(e.target.value)}
                />
              </div>

              <div className="bg-amber-50 rounded-xl border border-amber-100 p-3.5 text-[10.5px] leading-relaxed text-amber-800 font-medium">
                ⚠️ <strong>Catatan Kaskade Database:</strong> Mengubah identitas NIM/NIDN/NIS akan memperbarui relasi baris logbook &amp; surat tugas terkait secara otomatis agar data tetap terintegrasi sempurna di Google Sheets!
              </div>
            </div>

            <div className="bg-slate-50 px-6 py-4 border-t border-slate-100 flex items-center justify-end gap-2.5">
              <button 
                type="button" 
                onClick={() => setEditingUser(null)}
                className="px-3.5 py-1.5 text-xs text-slate-600 rounded-lg hover:bg-slate-100 font-bold"
              >
                Batal
              </button>
              <button 
                type="submit" 
                className="px-4 py-1.5 text-xs bg-[#003a70] text-white font-extrabold rounded-lg hover:bg-[#002244] shadow"
              >
                Simpan Perubahan
              </button>
            </div>
          </form>
        </div>
      )}

      {/* Confirm Dialog Modal */}
      {confirmDialog && (
        <div className="fixed inset-0 bg-slate-900/60 flex items-center justify-center p-4 z-[60] animate-fade-in">
          <div className="bg-white rounded-2xl shadow-2xl max-w-sm w-full overflow-hidden flex flex-col scale-100 animate-in fade-in zoom-in duration-200">
            <div className="bg-amber-50 p-5 flex flex-col items-center border-b border-amber-100">
              <div className="w-12 h-12 bg-amber-100 text-amber-600 rounded-full flex items-center justify-center mb-3">
                <AlertTriangle className="h-6 w-6" />
              </div>
              <h3 className="text-amber-800 font-extrabold text-lg">Konfirmasi</h3>
            </div>
            <div className="p-5 text-center">
              <p className="text-slate-600 font-medium text-sm leading-relaxed mb-6 whitespace-pre-line">{confirmDialog.message}</p>
              <div className="flex gap-3">
                <button 
                  onClick={() => setConfirmDialog(null)}
                  className="flex-1 bg-slate-100 hover:bg-slate-200 text-slate-700 py-2.5 rounded-xl font-bold transition-all shadow-sm"
                >
                  Batal
                </button>
                <button 
                  onClick={() => {
                    confirmDialog.onConfirm();
                    setConfirmDialog(null);
                  }}
                  className="flex-1 bg-emerald-600 hover:bg-emerald-700 text-white py-2.5 rounded-xl font-bold transition-all shadow-sm"
                >
                  Ya, Lanjutkan
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Generic App Modal (Replaces native alert) */}
      {appModal && (
        <div className="fixed inset-0 bg-slate-900/60 flex items-center justify-center p-4 z-[60] animate-fade-in">
          <div className="bg-white rounded-2xl shadow-2xl max-w-sm w-full overflow-hidden flex flex-col scale-100 animate-in fade-in zoom-in duration-200">
            <div className={`${appModal.type === 'error' ? 'bg-red-50 border-red-100' : appModal.type === 'success' ? 'bg-emerald-50 border-emerald-100' : appModal.type === 'warning' ? 'bg-amber-50 border-amber-100' : 'bg-sky-50 border-sky-100'} p-5 flex flex-col items-center border-b`}>
              <div className={`w-12 h-12 rounded-full flex items-center justify-center mb-3 ${appModal.type === 'error' ? 'bg-red-100 text-red-600' : appModal.type === 'success' ? 'bg-emerald-100 text-emerald-600' : appModal.type === 'warning' ? 'bg-amber-100 text-amber-600' : 'bg-sky-100 text-sky-600'}`}>
                {appModal.type === 'error' ? <AlertTriangle className="h-6 w-6" /> : appModal.type === 'success' ? <CheckCircle className="h-6 w-6" /> : <Info className="h-6 w-6" />}
              </div>
              <h3 className={`${appModal.type === 'error' ? 'text-red-700' : appModal.type === 'success' ? 'text-emerald-700' : appModal.type === 'warning' ? 'text-amber-800' : 'text-[#003a70]'} font-extrabold text-lg`}>{appModal.title}</h3>
            </div>
            <div className="p-5 text-center">
              <p className="text-slate-600 font-medium text-sm leading-relaxed mb-6 whitespace-pre-line">{appModal.content}</p>
              <button 
                onClick={() => setAppModal(null)}
                className={`w-full ${appModal.type === 'error' ? 'bg-red-600 hover:bg-red-700' : appModal.type === 'success' ? 'bg-emerald-600 hover:bg-emerald-700' : 'bg-[#003a70] hover:bg-[#002244]'} text-white py-2.5 rounded-xl font-bold transition-all shadow-sm`}
              >
                Tutup
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 4. MOCK INBOX ALERT FOR USER CONVENIENCE */}
      <footer className="border-t border-slate-200 py-6 px-6 text-center text-xs text-slate-500 flex flex-col sm:flex-row items-center justify-between gap-4 bg-slate-50 w-full mt-auto">
        <div>
          <span className="font-semibold text-slate-600">Universitas Negeri Yogyakarta • Falkutas Vokasi</span>
        </div>
        
        <div className="font-semibold text-slate-500">
          made by IT FV UNY @2026
        </div>
      </footer>

    </div>
  );
}

-- Supabase SQL Schema for FV-UNY Logbook App
-- Jalankan kode SQL ini di SQL Editor Supabase Anda

-- 1. Table users (Data Pengguna)
CREATE TABLE users (
  nim VARCHAR PRIMARY KEY,
  name VARCHAR NOT NULL,
  email VARCHAR NOT NULL,
  role VARCHAR NOT NULL, -- 'Anggota', 'Ketua', 'PIC', atau custom
  password VARCHAR NOT NULL,
  periode INTEGER,
  tanggalMulai VARCHAR,
  tanggalSelesai VARCHAR
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
  description TEXT,
  points JSONB, -- list/array dari string checklist
  workType VARCHAR NOT NULL,
  targetRole VARCHAR NOT NULL
);

-- 4. Table tasks (Penugasan ke Mahasiswa)
CREATE TABLE tasks (
  id VARCHAR PRIMARY KEY,
  masterId VARCHAR REFERENCES master_tasks(id) ON DELETE CASCADE,
  assignedNim VARCHAR REFERENCES users(nim) ON DELETE CASCADE,
  taskName VARCHAR NOT NULL,
  dateAssigned VARCHAR NOT NULL,
  status VARCHAR NOT NULL, -- 'Sedang Dikerjakan' atau 'Selesai'
  progress INTEGER DEFAULT 0,
  completedDesc TEXT,
  completedDate VARCHAR,
  googleDocUrl VARCHAR,
  points JSONB -- checklist progres spesifik user
);

-- 5. Table logbooks (Catatan Harian Pekerjaan)
CREATE TABLE logbooks (
  logbookId VARCHAR PRIMARY KEY,
  taskId VARCHAR REFERENCES tasks(id) ON DELETE CASCADE,
  nim VARCHAR REFERENCES users(nim) ON DELETE CASCADE,
  taskName VARCHAR NOT NULL,
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

-- Mengaktifkan RLS (Rules/Policies) agar anon access dapat leluasa jika dibutuhkan
-- Jika Anda butuh full akses melalui anon key, jalankan query berikut:
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

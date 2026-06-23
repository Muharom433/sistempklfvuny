-- Jalankan SQL ini di Supabase SQL Editor
-- https://supabase.com/dashboard -> SQL Editor

-- Tambah dua kolom integer langsung di tabel tasks
-- pointstotal  = total butir checklist task ini
-- checkedcount = berapa butir yang sudah dicentang user

ALTER TABLE tasks
  ADD COLUMN IF NOT EXISTS pointstotal  INTEGER DEFAULT 0,
  ADD COLUMN IF NOT EXISTS checkedcount INTEGER DEFAULT 0,
  ADD COLUMN IF NOT EXISTS percentage   INTEGER DEFAULT 0;

-- Setelah menjalankan SQL ini, buka aplikasi dan lakukan aksi apa saja
-- (centang / uncentang checklist) agar data ter-sync ke kolom baru.
-- Atau bisa refresh halaman -- auto-sync akan berjalan dalam 2 detik.

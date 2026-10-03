import { createClient } from '@supabase/supabase-js';
import ExcelJS from 'exceljs';
import nodemailer from 'nodemailer';
import sharp from 'sharp';

/* =========================================================
   KONFIGURASI DAN PEMERIKSAAN SECRET
========================================================= */

const REQUIRED_SECRETS = [
  'SUPABASE_URL',
  'SUPABASE_SECRET_KEY',
  'GMAIL_USER',
  'GMAIL_APP_PASSWORD',
  'REPORT_EMAIL_TO'
];

for (const secretName of REQUIRED_SECRETS) {
  if (!process.env[secretName]) {
    throw new Error(`Secret ${secretName} belum diisi`);
  }
}

const PHOTO_BUCKET =
  process.env.PHOTO_BUCKET || 'maintenance-photos';

const MAX_REPORTS = 100;
const TIME_ZONE = 'Asia/Jakarta';

/* =========================================================
   KONEKSI SUPABASE
========================================================= */

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SECRET_KEY,
  {
    auth: {
      persistSession: false,
      autoRefreshToken: false
    }
  }
);

/* =========================================================
   KONEKSI GMAIL
========================================================= */

const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: process.env.GMAIL_USER,
    pass: process.env.GMAIL_APP_PASSWORD
  }
});

/* =========================================================
   FUNGSI PENDUKUNG
========================================================= */

function safeText(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value);
}

function formatDateTime(value) {
  if (!value) {
    return '';
  }

  try {
    return new Intl.DateTimeFormat('id-ID', {
      timeZone: TIME_ZONE,
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false
    }).format(new Date(value));
  } catch {
    return safeText(value);
  }
}

function formatDateForFile(value = new Date()) {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: TIME_ZONE,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).formatToParts(value);

  const result = {};

  for (const part of parts) {
    result[part.type] = part.value;
  }

  return `${result.year}-${result.month}-${result.day}`;
}

function applyBorder(cell) {
  const side = {
    style: 'thin',
    color: {
      argb: 'FFD7DEE7'
    }
  };

  cell.border = {
    top: side,
    left: side,
    bottom: side,
    right: side
  };
}

/* =========================================================
   MENGAMBIL DATA SUPABASE
====================

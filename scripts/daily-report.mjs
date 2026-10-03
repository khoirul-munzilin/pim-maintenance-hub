import { createClient } from '@supabase/supabase-js';
import ExcelJS from 'exceljs';
import nodemailer from 'nodemailer';
import sharp from 'sharp';

/* =========================================================
   KONFIGURASI
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
   FUNGSI DASAR
========================================================= */

function safeText(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value);
}

function normalizeText(value) {
  return safeText(value)
    .trim()
    .toLowerCase();
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

function getLatestReport(reports) {
  if (!reports || reports.length === 0) {
    return null;
  }

  return [...reports].sort((a, b) => {
    return new Date(b.created_at) - new Date(a.created_at);
  })[0];
}

function getLatestValue(reports, fieldName) {
  const sortedReports = [...(reports || [])].sort((a, b) => {
    return new Date(b.created_at) - new Date(a.created_at);
  });

  for (const report of sortedReports) {
    if (
      report[fieldName] !== null &&
      report[fieldName] !== undefined &&
      safeText(report[fieldName]).trim() !== ''
    ) {
      return report[fieldName];
    }
  }

  return '';
}

/* =========================================================
   MENGAMBIL SELURUH DATA
========================================================= */

async function fetchAllData() {
  console.log(`Supabase URL: ${process.env.SUPABASE_URL}`);
  console.log(`Bucket foto: ${PHOTO_BUCKET}`);
  console.log(`Mengambil maksimal ${MAX_REPORTS} laporan terbaru`);

  const reportsResult = await supabase
    .from('work_reports')
    .select('*')
    .order('created_at', {
      ascending: false
    })
    .limit(MAX_REPORTS);

  if (reportsResult.error) {
    throw new Error(
      `Gagal mengambil work_reports: ${reportsResult.error.message}`
    );
  }

  const funlocResult = await supabase
    .from('functional_locations')
    .select('*');

  if (funlocResult.error) {
    throw new Error(
      `Gagal mengambil functional_locations: ${funlocResult.error.message}`
    );
  }

  const profilesResult = await supabase
    .from('profiles')
    .select('id, full_name, role, is_active');

  if (profilesResult.error) {
    throw new Error(
      `Gagal mengambil profiles: ${profilesResult.error.message}`
    );
  }

  const photosResult = await supabase
    .from('report_photos')
    .select('*')
    .order('created_at', {
      ascending: true
    });

  if (photosResult.error) {
    throw new Error(
      `Gagal mengambil report_photos: ${photosResult.error.message}`
    );
  }

  const reports = reportsResult.data || [];
  const functionalLocations = funlocResult.data || [];
  const profiles = profilesResult.data || [];
  const photos = photosResult.data || [];

  console.log(
    `DATA SUPABASE DITEMUKAN: ${reports.length} laporan`
  );

  console.log(
    `FOTO SUPABASE DITEMUKAN: ${photos.length} foto`
  );

  return reports.map((report) => {
    const functionalLocation = functionalLocations.find(
      (item) => item.id === report.funloc_id
    );

    const creator = profiles.find(
      (item) => item.id === report.created_by
    );

    const technician = profiles.find(
      (item) => item.id === report.assigned_to
    );

    const reportPhotos = photos.filter(
      (photo) => photo.report_id === report.id
    );

    return {
      ...report,
      functional_location: functionalLocation || null,
      creator_profile: creator || null,
      technician_profile: technician || null,
      report_photos: reportPhotos
    };
  });
}

/* =========================================================
   PENGELOMPOKAN FOTO
========================================================= */

function isInitialPhoto(photo) {
  return [
    'pm',
    'initial',
    'before',
    'documentation',
    'pm_inspection'
  ].includes(
    normalizeText(photo.photo_type)
  );
}

function isResultPhoto(photo) {
  return [
    'cm',
    'result',
    'after',
    'process',
    'cm_result'
  ].includes(
    normalizeText(photo.photo_type)
  );
}

function isVerificationPhoto(photo) {
  return [
    'verification',
    'verified',
    'approval'
  ].includes(
    normalizeText(photo.photo_type)
  );
}

/* =========================================================
   MEMBENTUK SATU RANGKAIAN PM, CM, VERIFIKASI
========================================================= */

function buildWorkGroups(reports) {
  const pmReports = reports.filter(
    (report) =>
      safeText(report.report_type).toUpperCase() === 'PM'
  );

  const cmReports = reports.filter(
    (report) =>
      safeText(report.report_type).toUpperCase() === 'CM'
  );

  const groups = [];
  const usedCmIds = new Set();

  /*
   * Prioritas pertama:
   * Pasangkan CM menggunakan pm_reference.
   */

  for (const pmReport of pmReports) {
    const relatedCmReports = cmReports.filter(
      (cmReport) =>
        cmReport.pm_reference === pmReport.id
    );

    for (const cmReport of relatedCmReports) {
      usedCmIds.add(cmReport.id);
    }

    groups.push(
      createWorkGroup(
        pmReport,
        relatedCmReports
      )
    );
  }

  /*
   * Prioritas kedua:
   * CM tanpa pm_reference dicari PM pada Funloc sama.
   */

  for (const cmReport of cmReports) {
    if (usedCmIds.has(cmReport.id)) {
      continue;
    }

    let matchingPm = null;

    if (cmReport.funloc_id) {
      const candidates = pmReports
        .filter(
          (pmReport) =>
            pmReport.funloc_id === cmReport.funloc_id
        )
        .sort((a, b) => {
          const cmTime = new Date(cmReport.created_at).getTime();

          const firstDifference = Math.abs(
            cmTime - new Date(a.created_at).getTime()
          );

          const secondDifference = Math.abs(
            cmTime - new Date(b.created_at).getTime()
          );

          return firstDifference - secondDifference;
        });

      matchingPm = candidates[0] || null;
    }

    if (matchingPm) {
      const existingGroup = groups.find(
        (group) =>
          group.pmReport?.id === matchingPm.id
      );

      if (existingGroup) {
        existingGroup.cmReports.push(cmReport);
        refreshWorkGroup(existingGroup);
      } else {
        groups.push(
          createWorkGroup(
            matchingPm,
            [cmReport]
          )
        );
      }

      usedCmIds.add(cmReport.id);
      continue;
    }

    groups.push(
      createWorkGroup(
        null,
        [cmReport]
      )
    );

    usedCmIds.add(cmReport.id);
  }

  return groups.sort((a, b) => {
    return (
      new Date(b.createdAt).getTime() -
      new Date(a.createdAt).getTime()
    );
  });
}

function createWorkGroup(
  pmReport,
  cmReports
) {
  const group = {
    pmReport: pmReport || null,
    cmReports: [...(cmReports || [])],
    functionalLocation:
      pmReport?.functional_location ||
      cmReports?.[0]?.functional_location ||
      null,
    createdAt:
      pmReport?.created_at ||
      cmReports?.[0]?.created_at ||
      new Date().toISOString(),
    initialPhotos: [],
    resultPhotos: [],
    verificationPhotos: [],
    reportNumbers: '',
    status: '',
    description: '',
    recommendation: '',
    correctiveAction: '',
    sparePart: '',
    technicianNote: '',
    testResult: '',
    verificationNote: '',
    reporterName: '',
    technicianName: '',
    priority: ''
  };

  refreshWorkGroup(group);

  return group;
}

function refreshWorkGroup(group) {
  const pmPhotos =
    group.pmReport?.report_photos || [];

  const cmPhotos =
    group.cmReports.flatMap(
      (report) => report.report_photos || []
    );

  const allPhotos = [
    ...pmPhotos,
    ...cmPhotos
  ];

  group.initialPhotos =
    allPhotos.filter(isInitialPhoto);

  group.resultPhotos =
    allPhotos.filter(isResultPhoto);

  group.verificationPhotos =
    allPhotos.filter(isVerificationPhoto);

  const cmNumbers = group.cmReports
    .map((report) => report.report_no)
    .filter(Boolean);

  const numbers = [];

  if (group.pmReport?.report_no) {
    numbers.push(group.pmReport.report_no);
  }

  numbers.push(...cmNumbers);

  group.reportNumbers = numbers.join(' → ');

  const latestCm =
    getLatestReport(group.cmReports);

  group.status =
    latestCm?.status ||
    group.pmReport?.status ||
    '';

  group.description =
    group.pmReport?.description ||
    group.cmReports[0]?.description ||
    '';

  group.recommendation =
    group.pmReport?.recommendation ||
    '';

  group.correctiveAction =
    getLatestValue(
      group.cmReports,
      'corrective_action'
    );

  group.sparePart =
    getLatestValue(
      group.cmReports,
      'spare_part'
    );

  group.technicianNote =
    getLatestValue(
      group.cmReports,
      'technician_note'
    );

  group.testResult =
    getLatestValue(
      group.cmReports,
      'test_result'
    );

  group.verificationNote =
    getLatestValue(
      group.cmReports,
      'verification_note'
    );

  group.reporterName =
    group.pmReport
      ?.creator_profile
      ?.full_name ||
    group.cmReports[0]
      ?.creator_profile
      ?.full_name ||
    '';

  group.technicianName =
    latestCm
      ?.technician_profile
      ?.full_name ||
    '';

  group.priority =
    latestCm?.priority ||
    group.pmReport?.priority ||
    '';
}

/* =========================================================
   DOWNLOAD DAN KOMPRESI FOTO
========================================================= */

async function downloadPhoto(storagePath) {
  if (!storagePath) {
    return null;
  }

  try {
    console.log(
      `Mengambil foto: ${storagePath}`
    );

    const downloadResult = await supabase.storage
      .from(PHOTO_BUCKET)
      .download(storagePath);

    if (
      downloadResult.error ||
      !downloadResult.data
    ) {
      console.warn(
        `Foto gagal diambil: ${storagePath}`
      );

      console.warn(
        downloadResult.error?.message ||
        'Data foto kosong'
      );

      return null;
    }

    const originalBuffer =

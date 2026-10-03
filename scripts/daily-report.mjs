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
========================================================= */

async function fetchAllData() {
  console.log(`Supabase URL: ${process.env.SUPABASE_URL}`);
  console.log(`Bucket foto: ${PHOTO_BUCKET}`);
  console.log(`Mengambil maksimal ${MAX_REPORTS} laporan terbaru...`);

  /*
   * Data diambil secara terpisah agar tidak tergantung
   * pada nama foreign key otomatis Supabase.
   */

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
  const funlocs = funlocResult.data || [];
  const profiles = profilesResult.data || [];
  const photos = photosResult.data || [];

  console.log(`DATA SUPABASE DITEMUKAN: ${reports.length} laporan`);
  console.log(`FOTO SUPABASE DITEMUKAN: ${photos.length} foto`);

  /*
   * Menggabungkan laporan, Functional Location,
   * pengguna, teknisi, dan foto berdasarkan UUID.
   */

  const mergedReports = reports.map((report) => {
    const functionalLocation = funlocs.find(
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

  for (const report of mergedReports) {
    console.log(
      `${safeText(report.report_no)}: ` +
      `${report.report_photos.length} foto`
    );

    for (const photo of report.report_photos) {
      console.log(
        `Foto type=${safeText(photo.photo_type)} ` +
        `path=${safeText(photo.storage_path)}`
      );
    }
  }

  return mergedReports;
}

/* =========================================================
   KATEGORI FOTO
========================================================= */

const PHOTO_CATEGORIES = [
  {
    label: 'Kondisi Awal',
    column: 1,
    types: [
      'pm',
      'initial',
      'before',
      'documentation',
      'pm_inspection'
    ]
  },
  {
    label: 'Hasil Pekerjaan',
    column: 3,
    types: [
      'cm',
      'result',
      'after',
      'process',
      'cm_result'
    ]
  },
  {
    label: 'Verifikasi',
    column: 5,
    types: [
      'verification',
      'verified',
      'approval'
    ]
  }
];

function getPhotosByCategory(report, category) {
  return (report.report_photos || [])
    .filter((photo) => {
      const type = safeText(photo.photo_type)
        .trim()
        .toLowerCase();

      return category.types.includes(type);
    })
    .sort((firstPhoto, secondPhoto) => {
      return (
        new Date(firstPhoto.created_at) -
        new Date(secondPhoto.created_at)
      );
    });
}

/* =========================================================
   DOWNLOAD DAN KOMPRESI FOTO
========================================================= */

async function downloadPhoto(storagePath) {
  if (!storagePath) {
    return null;
  }

  try {
    console.log(`Mengambil foto: ${storagePath}`);

    const downloadResult = await supabase.storage
      .from(PHOTO_BUCKET)
      .download(storagePath);

    if (downloadResult.error || !downloadResult.data) {
      console.warn(`Foto gagal diambil: ${storagePath}`);
      console.warn(
        downloadResult.error?.message || 'Data foto kosong'
      );

      return null;
    }

    const originalBuffer = Buffer.from(
      await downloadResult.data.arrayBuffer()
    );

    const compressedBuffer = await sharp(originalBuffer)
      .rotate()
      .resize({
        width: 900,
        height: 650,
        fit: 'inside',
        withoutEnlargement: true
      })
      .jpeg({
        quality: 68,
        mozjpeg: true
      })
      .toBuffer();

    console.log(`Foto berhasil diproses: ${storagePath}`);

    return compressedBuffer;
  } catch (error) {
    console.warn(
      `Gagal memproses foto ${storagePath}: ${error.message}`
    );

    return null;
  }
}

/* =========================================================
   MEMBUAT SHEET MAINTENANCE REPORT
========================================================= */

function createMaintenanceSheet(workbook, reports) {
  const worksheet = workbook.addWorksheet(
    'Maintenance Report',
    {
      views: [
        {
          state: 'frozen',
          ySplit: 1,
          showGridLines: false
        }
      ]
    }
  );

  const columns = [
    { header: 'No.', width: 7 },
    { header: 'Nomor Laporan', width: 24 },
    { header: 'Tanggal', width: 22 },
    { header: 'Jenis', width: 12 },
    { header: 'Functional Location', width: 42 },
    { header: 'Deskripsi Lokasi', width: 30 },
    { header: 'Area', width: 18 },
    { header: 'Prioritas', width: 14 },
    { header: 'Deskripsi / Temuan', width: 40 },
    { header: 'Dampak Operasional', width: 32 },
    { header: 'Status', width: 20 },
    { header: 'Pelapor', width: 22 },
    { header: 'Teknisi', width: 22 },
    { header: 'Penyebab', width: 32 },
    { header: 'Tindakan Perbaikan', width: 40 },
    { header: 'Spare Part', width: 28 },
    { header: 'Catatan Teknisi', width: 34 },
    { header: 'Hasil Test', width: 20 },
    { header: 'Catatan Verifikasi', width: 34 },
    { header: 'Waktu Mulai', width: 22 },
    { header: 'Request Verification', width: 24 },
    { header: 'Waktu Closed', width: 22 }
  ];

  worksheet.columns = columns.map((column, index) => ({
    header: column.header,
    key: `column_${index + 1}`,
    width: column.width
  }));

  worksheet.getRow(1).height = 34;

  worksheet.getRow(1).eachCell((cell) => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {
        argb: 'FF163A5F'
      }
    };

    cell.font = {
      bold: true,
      size: 10,
      color: {
        argb: 'FFFFFFFF'
      }
    };

    cell.alignment = {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true
    };

    applyBorder(cell);
  });

  if (reports.length === 0) {
    worksheet.mergeCells('A2:V4');

    const emptyCell = worksheet.getCell('A2');

    emptyCell.value = 'Belum ada laporan pada database.';

    emptyCell.alignment = {
      horizontal: 'center',
      vertical: 'middle'
    };

    emptyCell.font = {
      italic: true,
      color: {
        argb: 'FF64748B'
      }
    };

    return worksheet;
  }

  reports.forEach((report, index) => {
    const row = worksheet.addRow([
      index + 1,
      safeText(report.report_no),
      formatDateTime(report.created_at),
      safeText(report.report_type),
      safeText(report.functional_location?.funloc_code),
      safeText(report.functional_location?.description),
      safeText(report.functional_location?.area),
      safeText(report.priority),
      safeText(report.description),
      safeText(report.operational_impact),
      safeText(report.status),
      safeText(report.creator_profile?.full_name),
      safeText(report.technician_profile?.full_name),
      safeText(report.failure_cause),
      safeText(
        report.corrective_action ||
        report.recommendation
      ),
      safeText(report.spare_part),
      safeText(report.technician_note),
      safeText(report.test_result),
      safeText(report.verification_note),
      formatDateTime(report.started_at),
      formatDateTime(
        report.verification_requested_at
      ),
      formatDateTime(report.closed_at)
    ]);

    row.height = 48;

    row.eachCell(
      {
        includeEmpty: true
      },
      (cell) => {
        cell.alignment = {
          vertical: 'top',
          wrapText: true
        };

        cell.font = {
          size: 10,
          color: {
            argb: 'FF111827'
          }
        };

        applyBorder(cell);
      }
    );

    const statusCell = row.getCell(11);

    const status = safeText(report.status)
      .trim()
      .toUpperCase();

    if (status === 'CLOSED') {
      statusCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {
          argb: 'FFDCFCE7'
        }
      };

      statusCell.font = {
        bold: true,
        color: {
          argb: 'FF166534'
        }
      };
    } else if (
      status.includes('PROGRESS') ||
      status.includes('ASSIGNED')
    ) {
      statusCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {
          argb: 'FFDBEAFE'
        }
      };

      statusCell.font = {
        bold: true,
        color: {
          argb: 'FF1D4ED8'
        }
      };
    } else {
      statusCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {
          argb: 'FFFEF3C7'
        }
      };

      statusCell.font = {
        bold: true,
        color: {
          argb: 'FF92400E'
        }
      };
    }
  });

  worksheet.autoFilter = {
    from: 'A1',
    to: `V${reports.length + 1}`
  };

  worksheet.pageSetup = {
    orientation: 'landscape',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    paperSize: 9,
    margins: {
      left: 0.25,
      right: 0.25,
      top: 0.5,
      bottom: 0.5,
      header: 0.2,
      footer: 0.2
    }
  };

  return worksheet;
}

/* =========================================================
   MEMBUAT SHEET DOKUMENTASI
========================================================= */

async function createDocumentationSheet(
  workbook,
  reports
) {
  const worksheet = workbook.addWorksheet(
    'Dokumentasi',
    {
      views: [
        {
          showGridLines: false
        }
      ]
    }
  );

  worksheet.getColumn('A').width = 36;
  worksheet.getColumn('B').width = 4;
  worksheet.getColumn('C').width = 36;
  worksheet.getColumn('D').width = 4;
  worksheet.getColumn('E').width = 36;

  if (reports.length === 0) {
    worksheet.mergeCells('A1:E5');

    const emptyCell = worksheet.getCell('A1');

    emptyCell.value =
      'Belum ada laporan dan dokumentasi.';

    emptyCell.alignment = {
      horizontal: 'center',
      vertical: 'middle'
    };

    emptyCell.font = {
      italic: true,
      color: {
        argb: 'FF64748B'
      }
    };

    return worksheet;
  }

  for (
    let reportIndex = 0;
    reportIndex < reports.length;
    reportIndex += 1
  ) {
    const report = reports[reportIndex];

    const startRow = reportIndex * 23 + 1;
    const endImageRow = startRow + 20;

    worksheet.mergeCells(
      startRow,
      1,
      startRow,
      5
    );

    const titleCell = worksheet.getCell(
      startRow,
      1
    );

    titleCell.value =
      `${safeText(report.report_no)} • ` +
      `${safeText(
        report.functional_location?.funloc_code
      )}`;

    titleCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {
        argb: 'FF163A5F'
      }
    };

    titleCell.font = {
      bold: true,
      size: 12,
      color: {
        argb: 'FFFFFFFF'
      }
    };

    titleCell.alignment = {
      vertical: 'middle'
    };

    worksheet.getRow(startRow).height = 26;

    for (const category of PHOTO_CATEGORIES) {
      const headerCell = worksheet.getCell(
        startRow + 1,
        category.column
      );

      headerCell.value = category.label;

      headerCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {
          argb: 'FF0F766E'
        }
      };

      headerCell.font = {
        bold: true,
        color: {
          argb: 'FFFFFFFF'
        }
      };

      headerCell.alignment = {
        horizontal: 'center',
        vertical: 'middle'
      };

      for (
        let rowNumber = startRow + 2;
        rowNumber <= endImageRow;
        rowNumber += 1
      ) {
        worksheet.getRow(rowNumber).height = 18;

        const areaCell = worksheet.getCell(
          rowNumber,
          category.column
        );

        areaCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: 'FFF8FAFC'
          }
        };

        applyBorder(areaCell);
      }

      const categoryPhotos = getPhotosByCategory(
        report,
        category
      );

      console.log(
        `${safeText(report.report_no)} | ` +
        `${category.label}: ` +
        `${categoryPhotos.length} foto`
      );

      const selectedPhoto = categoryPhotos[0];

      if (!selectedPhoto) {
        const noPhotoCell = worksheet.getCell(
          startRow + 10,
          category.column
        );

        noPhotoCell.value = 'Foto belum tersedia';

        noPhotoCell.font = {
          italic: true,
          color: {
            argb: 'FF64748B'
          }
        };

        noPhotoCell.alignment = {
          horizontal: 'center',
          vertical: 'middle'
        };

        continue;
      }

      const imageBuffer = await downloadPhoto(
        selectedPhoto.storage_path
      );

      if (!imageBuffer) {
        const errorPhotoCell = worksheet.getCell(
          startRow + 10,
          category.column
        );

        errorPhotoCell.value =
          'Foto tidak dapat dimuat';

        errorPhotoCell.font = {
          italic: true,
          color: {
            argb: 'FFB91C1C'
          }
        };

        errorPhotoCell.alignment = {
          horizontal: 'center',
          vertical: 'middle'
        };

        continue;
      }

      const imageId = workbook.addImage({
        buffer: imageBuffer,
        extension: 'jpeg'
      });

      worksheet.addImage(imageId, {
        tl: {
          col: category.column - 1 + 0.08,
          row: startRow + 1 + 0.2
        },
        br: {
          col: category.column - 0.08,
          row: endImageRow + 0.8
        },
        editAs: 'oneCell'
      });
    }

    worksheet.mergeCells(
      startRow + 21,
      1,
      startRow + 21,
      5
    );

    const noteCell = worksheet.getCell(
      startRow + 21,
      1
    );

    noteCell.value =
      `Deskripsi: ${safeText(report.description)} | ` +
      `Status: ${safeText(report.status)}`;

    noteCell.font = {
      size: 9,
      color: {
        argb: 'FF475569'
      }
    };

    noteCell.alignment = {
      wrapText: true,
      vertical: 'middle'
    };

    worksheet.getRow(startRow + 21).height = 28;
  }

  worksheet.pageSetup = {
    orientation: 'landscape',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    paperSize: 9,
    margins: {
      left: 0.25,
      right: 0.25,
      top: 0.5,
      bottom: 0.5,
      header: 0.2,
      footer: 0.2
    }
  };

  return worksheet;
}

/* =========================================================
   MENGIRIM EMAIL
========================================================= */

async function sendEmail(workbook, reports) {
  const reportDate = formatDateForFile();

  const fileName =
    `PIM_PM_CM_Report_${reportDate}.xlsx`;

  console.log('Membuat file Excel...');

  const excelBuffer =
    await workbook.xlsx.writeBuffer();

  console.log(
    `Ukuran Excel: ` +
    `${Math.round(excelBuffer.byteLength / 1024)} KB`
  );

  const totalPM = reports.filter(
    (report) =>
      safeText(report.report_type).toUpperCase() ===
      'PM'
  ).length;

  const totalCM = reports.filter(
    (report) =>
      safeText(report.report_type).toUpperCase() ===
      'CM'
  ).length;

  const totalClosed = reports.filter(
    (report) =>
      safeText(report.status).toUpperCase() ===
      'CLOSED'
  ).length;

  const totalOpen =
    reports.length - totalClosed;

  const emailRecipients =
    process.env.REPORT_EMAIL_TO
      .split(',')
      .map((email) => email.trim())
      .filter(Boolean)
      .join(',');

  console.log('Memeriksa koneksi Gmail SMTP...');

  await transporter.verify();

  console.log('Koneksi Gmail SMTP berhasil.');

  const result = await transporter.sendMail({
    from:
      `PIM Maintenance Hub ` +
      `<${process.env.GMAIL_USER}>`,

    to: emailRecipients,

    subject:
      `PIM PM & CM Report - ${reportDate}`,

    html: `
      <p>Yth. Bapak/Ibu,</p>

      <p>
        Berikut kami sampaikan laporan pekerjaan
        Preventive Maintenance dan Corrective Maintenance.
      </p>

      <h3>Ringkasan Laporan</h3>

      <ul>
        <li>
          Total laporan:
          <b>${reports.length}</b>
        </li>

        <li>
          Preventive Maintenance:
          <b>${totalPM}</b>
        </li>

        <li>
          Corrective Maintenance:
          <b>${totalCM}</b>
        </li>

        <li>
          Pekerjaan terbuka:
          <b>${totalOpen}</b>
        </li>

        <li>
          Pekerjaan selesai:
          <b>${totalClosed}</b>
        </li>
      </ul>

      <p>
        File Excel lengkap beserta dokumentasi foto
        terlampir pada email ini.
      </p>

      <p>
        Email ini dikirim otomatis oleh
        PIM Maintenance Hub.
      </p>
    `,

    attachments: [
      {
        filename: fileName,
        content: Buffer.from(excelBuffer),
        contentType:
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      }
    ]
  });

  console.log(
    `Email berhasil dikirim dari: ` +
    `${process.env.GMAIL_USER}`
  );

  console.log(`Penerima: ${emailRecipients}`);
  console.log(`Message ID: ${result.messageId}`);

  console.log(
    `Accepted: ${JSON.stringify(result.accepted)}`
  );

  console.log(
    `Rejected: ${JSON.stringify(result.rejected)}`
  );

  console.log(`Response: ${result.response}`);
  console.log(`Lampiran: ${fileName}`);
}

/* =========================================================
   PROSES UTAMA
========================================================= */

async function main() {
  console.log('========================================');
  console.log('PIM Maintenance Hub Daily Report');
  console.log('========================================');

  const reports = await fetchAllData();

  const totalPhotos = reports.reduce(
    (total, report) =>
      total + report.report_photos.length,
    0
  );

  console.log(
    `TOTAL FINAL: ${reports.length} laporan`
  );

  console.log(
    `TOTAL FOTO FINAL: ${totalPhotos} foto`
  );

  const workbook = new ExcelJS.Workbook();

  workbook.creator = 'PIM Maintenance Hub';
  workbook.company = 'PT Padi Indonesia Maju';
  workbook.created = new Date();
  workbook.modified = new Date();

  createMaintenanceSheet(workbook, reports);

  await createDocumentationSheet(
    workbook,
    reports
  );

  await sendEmail(workbook, reports);

  console.log('========================================');
  console.log('Proses report selesai.');
  console.log('========================================');
}

main().catch((error) => {
  console.error('========================================');
  console.error('Gagal membuat atau mengirim laporan.');
  console.error(error);
  console.error('========================================');

  process.exit(1);
});

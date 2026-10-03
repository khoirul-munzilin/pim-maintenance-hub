import { createClient } from '@supabase/supabase-js';
import ExcelJS from 'exceljs';
import nodemailer from 'nodemailer';
import sharp from 'sharp';

/*
|--------------------------------------------------------------------------
| Pemeriksaan GitHub Secrets
|--------------------------------------------------------------------------
*/

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

/*
|--------------------------------------------------------------------------
| Konfigurasi
|--------------------------------------------------------------------------
*/

const PHOTO_BUCKET =
  process.env.PHOTO_BUCKET || 'maintenance-photos';

const TIME_ZONE = 'Asia/Jakarta';

const MAX_REPORTS = 100;

/*
|--------------------------------------------------------------------------
| Koneksi Supabase
|--------------------------------------------------------------------------
*/

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

/*
|--------------------------------------------------------------------------
| Koneksi Gmail SMTP
|--------------------------------------------------------------------------
*/

const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: process.env.GMAIL_USER,
    pass: process.env.GMAIL_APP_PASSWORD
  }
});

/*
|--------------------------------------------------------------------------
| Fungsi Pendukung
|--------------------------------------------------------------------------
*/

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

function formatDateForFile(value) {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: TIME_ZONE,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).formatToParts(value);

  const dateParts = Object.fromEntries(
    parts.map((part) => [
      part.type,
      part.value
    ])
  );

  return (
    `${dateParts.year}-` +
    `${dateParts.month}-` +
    `${dateParts.day}`
  );
}

function applyThinBorder(cell) {
  cell.border = {
    top: {
      style: 'thin',
      color: {
        argb: 'FFD7DEE7'
      }
    },
    left: {
      style: 'thin',
      color: {
        argb: 'FFD7DEE7'
      }
    },
    bottom: {
      style: 'thin',
      color: {
        argb: 'FFD7DEE7'
      }
    },
    right: {
      style: 'thin',
      color: {
        argb: 'FFD7DEE7'
      }
    }
  };
}

/*
|--------------------------------------------------------------------------
| Mengambil Data Laporan
|--------------------------------------------------------------------------
|
| Untuk pengujian ini, script mengambil 100 laporan terbaru.
| Tidak menggunakan filter tanggal agar data lama tetap masuk.
|
*/

async function fetchReports() {
  const query = `
    *,
    functional_locations (
      funloc_code,
      description,
      area
    ),
    creator:profiles!work_reports_created_by_fkey (
      full_name
    ),
    assignee:profiles!work_reports_assigned_to_fkey (
      full_name
    ),
    report_photos (
      id,
      storage_path,
      photo_type,
      created_at
    )
  `;

  console.log(
    `Mengambil maksimal ${MAX_REPORTS} laporan terbaru...`
  );

  const { data, error } = await supabase
    .from('work_reports')
    .select(query)
    .order('created_at', {
      ascending: false
    })
    .limit(MAX_REPORTS);

  if (error) {
    throw new Error(
      `Gagal mengambil laporan Supabase: ${error.message}`
    );
  }

  const reports = data || [];

  console.log(
    `Data laporan dari Supabase: ${reports.length}`
  );

  return reports;
}

/*
|--------------------------------------------------------------------------
| Menampilkan Informasi Foto di Log
|--------------------------------------------------------------------------
*/

function logReportPhotos(reports) {
  const totalPhotos = reports.reduce(
    (total, report) => {
      return (
        total +
        (report.report_photos?.length || 0)
      );
    },
    0
  );

  console.log(
    `Jumlah metadata foto ditemukan: ${totalPhotos}`
  );

  for (const report of reports) {
    const photos =
      report.report_photos || [];

    console.log(
      `${safeText(report.report_no)}: ` +
      `${photos.length} foto`
    );

    for (const photo of photos) {
      console.log(
        `  Foto: ` +
        `type=${safeText(photo.photo_type)}, ` +
        `path=${safeText(photo.storage_path)}`
      );
    }
  }
}

/*
|--------------------------------------------------------------------------
| Mengambil dan Mengompres Foto
|--------------------------------------------------------------------------
*/

async function downloadAndCompressImage(
  storagePath
) {
  if (!storagePath) {
    return null;
  }

  try {
    console.log(
      `Mengambil foto: ${storagePath}`
    );

    const { data, error } =
      await supabase.storage
        .from(PHOTO_BUCKET)
        .download(storagePath);

    if (error || !data) {
      console.warn(
        `Foto tidak dapat diambil: ${storagePath}`
      );

      console.warn(
        error?.message || 'Data foto kosong'
      );

      return null;
    }

    const originalBuffer = Buffer.from(
      await data.arrayBuffer()
    );

    const compressedBuffer =
      await sharp(originalBuffer)
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

    console.log(
      `Foto berhasil diproses: ${storagePath}`
    );

    return compressedBuffer;
  } catch (error) {
    console.warn(
      `Gagal memproses foto: ${storagePath}`
    );

    console.warn(error.message);

    return null;
  }
}

/*
|--------------------------------------------------------------------------
| Pengelompokan Tipe Foto
|--------------------------------------------------------------------------
|
| Script menerima tipe foto lama maupun baru.
|
| Kondisi Awal:
| initial, pm, documentation, pm_inspection, before
|
| Hasil Pekerjaan:
| result, cm, after, process, cm_result
|
| Verifikasi:
| verification, verified, approval
|
*/

const PHOTO_CATEGORIES = [
  {
    types: [
      'initial',
      'pm',
      'documentation',
      'pm_inspection',
      'before'
    ],
    label: 'Kondisi Awal',
    column: 1
  },
  {
    types: [
      'result',
      'cm',
      'after',
      'process',
      'cm_result'
    ],
    label: 'Hasil Pekerjaan',
    column: 3
  },
  {
    types: [
      'verification',
      'verified',
      'approval'
    ],
    label: 'Verifikasi',
    column: 5
  }
];

function getPhotosByCategory(
  report,
  category
) {
  const photos =
    report.report_photos || [];

  return photos
    .filter((photo) => {
      const photoType = safeText(
        photo.photo_type
      )
        .trim()
        .toLowerCase();

      return category.types.includes(
        photoType
      );
    })
    .sort((firstPhoto, secondPhoto) => {
      return (
        new Date(firstPhoto.created_at) -
        new Date(secondPhoto.created_at)
      );
    });
}

/*
|--------------------------------------------------------------------------
| Membuat Sheet Maintenance Report
|--------------------------------------------------------------------------
*/

function createMaintenanceSheet(
  workbook,
  reports
) {
  const worksheet =
    workbook.addWorksheet(
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

  const headers = [
    'No.',
    'Nomor Laporan',
    'Tanggal',
    'Jenis',
    'Functional Location',
    'Deskripsi Lokasi',
    'Area',
    'Prioritas',
    'Deskripsi / Temuan',
    'Dampak Operasional',
    'Status',
    'Pelapor',
    'Teknisi',
    'Penyebab',
    'Tindakan Perbaikan',
    'Spare Part',
    'Catatan Teknisi',
    'Hasil Test',
    'Catatan Verifikasi',
    'Waktu Mulai',
    'Request Verification',
    'Waktu Closed'
  ];

  const columnWidths = [
    7,
    24,
    22,
    12,
    42,
    30,
    18,
    14,
    40,
    32,
    18,
    22,
    22,
    32,
    40,
    28,
    34,
    20,
    34,
    22,
    24,
    22
  ];

  worksheet.columns =
    headers.map((header, index) => ({
      header,
      key: `column_${index + 1}`,
      width: columnWidths[index]
    }));

  worksheet.getRow(1).height = 34;

  worksheet.autoFilter = {
    from: 'A1',
    to: `V${Math.max(
      reports.length + 1,
      1
    )}`
  };

  worksheet
    .getRow(1)
    .eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {
          argb: 'FF163A5F'
        }
      };

      cell.font = {
        bold: true,
        color: {
          argb: 'FFFFFFFF'
        },
        size: 10
      };

      cell.alignment = {
        horizontal: 'center',
        vertical: 'middle',
        wrapText: true
      };

      applyThinBorder(cell);
    });

  /*
   * Jika tidak ada data.
   */
  if (reports.length === 0) {
    worksheet.mergeCells('A2:V4');

    const emptyCell =
      worksheet.getCell('A2');

    emptyCell.value =
      'Belum ada laporan pada database.';

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
  }

  /*
   * Menambahkan baris laporan.
   */
  reports.forEach(
    (report, index) => {
      const row =
        worksheet.addRow([
          index + 1,

          safeText(
            report.report_no
          ),

          formatDateTime(
            report.created_at
          ),

          safeText(
            report.report_type
          ),

          safeText(
            report
              .functional_locations
              ?.funloc_code
          ),

          safeText(
            report
              .functional_locations
              ?.description
          ),

          safeText(
            report
              .functional_locations
              ?.area
          ),

          safeText(
            report.priority
          ),

          safeText(
            report.description
          ),

          safeText(
            report.operational_impact
          ),

          safeText(
            report.status
          ),

          safeText(
            report.creator?.full_name
          ),

          safeText(
            report.assignee?.full_name
          ),

          safeText(
            report.failure_cause
          ),

          safeText(
            report.corrective_action ||
            report.recommendation
          ),

          safeText(
            report.spare_part
          ),

          safeText(
            report.technician_note
          ),

          safeText(
            report.test_result
          ),

          safeText(
            report.verification_note
          ),

          formatDateTime(
            report.started_at
          ),

          formatDateTime(
            report
              .verification_requested_at
          ),

          formatDateTime(
            report.closed_at
          )
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

          applyThinBorder(cell);
        }
      );

      /*
       * Warna kolom status.
       */
      const statusCell =
        row.getCell(11);

      const status =
        safeText(report.status)
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
        status.includes(
          'PROGRESS'
        ) ||
        status.includes(
          'ASSIGNED'
        )
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
    }
  );

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

/*
|--------------------------------------------------------------------------
| Membuat Sheet Dokumentasi
|--------------------------------------------------------------------------
*/

async function createDocumentationSheet(
  workbook,
  reports
) {
  const worksheet =
    workbook.addWorksheet(
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

  /*
   * Jika tidak ada laporan.
   */
  if (reports.length === 0) {
    worksheet.mergeCells('A1:E5');

    const emptyCell =
      worksheet.getCell('A1');

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

  /*
   * Membuat satu blok dokumentasi
   * untuk setiap laporan.
   */
  for (
    let reportIndex = 0;
    reportIndex < reports.length;
    reportIndex += 1
  ) {
    const report =
      reports[reportIndex];

    const startRow =
      reportIndex * 23 + 1;

    const endImageRow =
      startRow + 20;

    /*
     * Judul laporan.
     */
    worksheet.mergeCells(
      startRow,
      1,
      startRow,
      5
    );

    const reportTitle =
      worksheet.getCell(
        startRow,
        1
      );

    reportTitle.value =
      `${safeText(
        report.report_no
      )} • ` +
      `${safeText(
        report
          .functional_locations
          ?.funloc_code
      )}`;

    reportTitle.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {
        argb: 'FF163A5F'
      }
    };

    reportTitle.font = {
      bold: true,
      color: {
        argb: 'FFFFFFFF'
      },
      size: 12
    };

    reportTitle.alignment = {
      vertical: 'middle'
    };

    worksheet
      .getRow(startRow)
      .height = 26;

    /*
     * Membuat kategori foto.
     */
    for (
      const category
      of PHOTO_CATEGORIES
    ) {
      const heading =
        worksheet.getCell(
          startRow + 1,
          category.column
        );

      heading.value =
        category.label;

      heading.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {
          argb: 'FF0F766E'
        }
      };

      heading.font = {
        bold: true,
        color: {
          argb: 'FFFFFFFF'
        }
      };

      heading.alignment = {
        horizontal: 'center',
        vertical: 'middle'
      };

      /*
       * Membuat area foto.
       */
      for (
        let rowNumber =
          startRow + 2;

        rowNumber <=
          endImageRow;

        rowNumber += 1
      ) {
        worksheet
          .getRow(rowNumber)
          .height = 18;

        const photoAreaCell =
          worksheet.getCell(
            rowNumber,
            category.column
          );

        photoAreaCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: 'FFF8FAFC'
          }
        };

        applyThinBorder(
          photoAreaCell
        );
      }

      /*
       * Mengambil foto kategori.
       */
      const categoryPhotos =
        getPhotosByCategory(
          report,
          category
        );

      console.log(
        `${safeText(
          report.report_no
        )} | ` +
        `${category.label}: ` +
        `${categoryPhotos.length} foto`
      );

      /*
       * Versi ini memakai satu foto
       * pertama dari setiap kategori.
       */
      const selectedPhoto =
        categoryPhotos[0];

      if (!selectedPhoto) {
        const noPhotoCell =
          worksheet.getCell(
            startRow + 10,
            category.column
          );

        noPhotoCell.value =
          'Foto belum tersedia';

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

      const imageBuffer =
        await downloadAndCompressImage(
          selectedPhoto.storage_path
        );

      if (!imageBuffer) {
        const failedPhotoCell =
          worksheet.getCell(
            startRow + 10,
            category.column
          );

        failedPhotoCell.value =
          'Foto tidak dapat dimuat';

        failedPhotoCell.font = {
          italic: true,
          color: {
            argb: 'FFB91C1C'
          }
        };

        failedPhotoCell.alignment = {
          horizontal: 'center',
          vertical: 'middle'
        };

        continue;
      }

      /*
       * Memasukkan foto ke Excel.
       */
      const imageId =
        workbook.addImage({
          buffer: imageBuffer,
          extension: 'jpeg'
        });

      worksheet.addImage(
        imageId,
        {
          tl: {
            col:
              category.column -
              1 +
              0.05,

            row:
              startRow +
              1 +
              0.15
          },

          br: {
            col:
              category.column -
              0.05,

            row:
              endImageRow +
              0.85
          },

          editAs: 'oneCell'
        }
      );
    }

    /*
     * Catatan laporan.
     */
    worksheet.mergeCells(
      startRow + 21,
      1,
      startRow + 21,
      5
    );

    const noteCell =
      worksheet.getCell(
        startRow + 21,
        1
      );

    noteCell.value =
      `Deskripsi: ` +
      `${safeText(
        report.description
      )} | ` +
      `Status: ` +
      `${safeText(
        report.status
      )}`;

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

    worksheet
      .getRow(startRow + 21)
      .height = 28;
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

/*
|--------------------------------------------------------------------------
| Membuat dan Mengirim Email
|--------------------------------------------------------------------------
*/

async function sendReportEmail(
  workbook,
  reports
) {
  const reportDate =
    formatDateForFile(
      new Date()
    );

  const fileName =
    `PIM_PM_CM_Report_` +
    `${reportDate}.xlsx`;

  console.log(
    'Membuat file Excel...'
  );

  const excelBuffer =
    await workbook.xlsx
      .writeBuffer();

  console.log(
    `Ukuran Excel: ` +
    `${Math.round(
      excelBuffer.byteLength /
      1024
    )} KB`
  );

  const totalPM =
    reports.filter(
      (report) =>
        safeText(
          report.report_type
        ).toUpperCase() === 'PM'
    ).length;

  const totalCM =
    reports.filter(
      (report) =>
        safeText(
          report.report_type
        ).toUpperCase() === 'CM'
    ).length;

  const totalClosed =
    reports.filter(
      (report) =>
        safeText(
          report.status
        ).toUpperCase() ===
        'CLOSED'
    ).length;

  const totalOpen =
    reports.length -
    totalClosed;

  const emailRecipients =
    process.env.REPORT_EMAIL_TO
      .split(',')
      .map(
        (email) =>
          email.trim()
      )
      .filter(Boolean)
      .join(',');

  console.log(
    'Memeriksa koneksi Gmail SMTP...'
  );

  await transporter.verify();

  console.log(
    'Koneksi Gmail SMTP berhasil.'
  );

  const result =
    await transporter.sendMail({
      from:
        `PIM Maintenance Hub ` +
        `<${process.env.GMAIL_USER}>`,

      to: emailRecipients,

      subject:
        `PIM PM & CM Report - ` +
        `${reportDate}`,

      html: `
        <p>Yth. Bapak/Ibu,</p>

        <p>
          Berikut kami sampaikan
          laporan pekerjaan Preventive
          Maintenance dan Corrective
          Maintenance.
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
          File Excel lengkap beserta
          dokumentasi foto terlampir.
        </p>

        <p>
          Email ini dikirim otomatis
          oleh PIM Maintenance Hub.
        </p>
      `,

      attachments: [
        {
          filename: fileName,

          content:
            Buffer.from(
              excelBuffer
            ),

          contentType:
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
      ]
    });

  console.log(
    `Email berhasil dikirim dari: ` +
    `${process.env.GMAIL_USER}`
  );

  console.log(
    `Penerima: ` +
    `${emailRecipients}`
  );

  console.log(
    `Message ID: ` +
    `${result.messageId}`
  );

  console.log(
    `Accepted: ` +
    `${JSON.stringify(
      result.accepted
    )}`
  );

  console.log(
    `Rejected: ` +
    `${JSON.stringify(
      result.rejected
    )}`
  );

  console.log(
    `Response: ` +
    `${result.response}`
  );

  console.log(
    `Lampiran: ${fileName}`
  );
}

/*
|--------------------------------------------------------------------------
| Proses Utama
|--------------------------------------------------------------------------
*/

async function main() {
  console.log(
    '========================================'
  );

  console.log(
    'PIM Maintenance Hub Daily Report'
  );

  console.log(
    '========================================'
  );

  console.log(
    `Supabase URL: ` +
    `${process.env.SUPABASE_URL}`
  );

  console.log(
    `Bucket foto: ${PHOTO_BUCKET}`
  );

  /*
   * Mengambil seluruh data terbaru.
   */
  const reports =
    await fetchReports();

  /*
   * Menampilkan data foto di log.
   */
  logReportPhotos(reports);

  /*
   * Membuat workbook.
   */
  const workbook =
    new ExcelJS.Workbook();

  workbook.creator =
    'PIM Maintenance Hub';

  workbook.company =
    'PT Padi Indonesia Maju';

  workbook.created =
    new Date();

  workbook.modified =
    new Date();

  /*
   * Membuat sheet laporan.
   */
  createMaintenanceSheet(
    workbook,
    reports
  );

  /*
   * Membuat sheet dokumentasi.
   */
  await createDocumentationSheet(
    workbook,
    reports
  );

  /*
   * Mengirim email.
   */
  await sendReportEmail(
    workbook,
    reports
  );

  console.log(
    '========================================'
  );

  console.log(
    'Proses report selesai.'
  );

  console.log(
    '========================================'
  );
}

/*
|--------------------------------------------------------------------------
| Menjalankan Program
|--------------------------------------------------------------------------
*/

main().catch((error) => {
  console.error(
    '========================================'
  );

  console.error(
    'Gagal membuat atau mengirim laporan.'
  );

  console.error(
    error
  );

  console.error(
    '========================================'
  );

  process.exit(1);
});

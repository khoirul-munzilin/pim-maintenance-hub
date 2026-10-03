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
| Menentukan Periode Laporan
|--------------------------------------------------------------------------
|
| Periode laporan:
| Kemarin pukul 07.30 WIB
| sampai
| Hari ini pukul 07.29 WIB
|
| 07.30 WIB = 00.30 UTC
|
*/

function getReportPeriod() {
  const now = new Date();

  const end = new Date(now);

  end.setUTCHours(0, 30, 0, 0);

  /*
   * Jika workflow dijalankan secara manual sebelum pukul 07.30 WIB,
   * gunakan batas akhir hari sebelumnya.
   */
  if (now < end) {
    end.setUTCDate(end.getUTCDate() - 1);
  }

  const start = new Date(end);

  start.setUTCDate(start.getUTCDate() - 1);

  return {
    start,
    end
  };
}

/*
|--------------------------------------------------------------------------
| Format Tanggal dan Jam
|--------------------------------------------------------------------------
*/

function formatDateTime(value) {
  if (!value) {
    return '';
  }

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
}

/*
|--------------------------------------------------------------------------
| Format Tanggal untuk Nama File
|--------------------------------------------------------------------------
*/

function formatDateForFile(value) {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: TIME_ZONE,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).formatToParts(value);

  const dateParts = Object.fromEntries(
    parts.map((part) => [part.type, part.value])
  );

  return `${dateParts.year}-${dateParts.month}-${dateParts.day}`;
}

/*
|--------------------------------------------------------------------------
| Mengamankan Nilai Kosong
|--------------------------------------------------------------------------
*/

function safeText(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value);
}

/*
|--------------------------------------------------------------------------
| Format Border Excel
|--------------------------------------------------------------------------
*/

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
| Mengambil dan Mengompres Foto
|--------------------------------------------------------------------------
*/

async function downloadAndCompressImage(storagePath) {
  try {
    const { data, error } = await supabase.storage
      .from(PHOTO_BUCKET)
      .download(storagePath);

    if (error || !data) {
      console.warn(
        `Foto tidak dapat diambil: ${storagePath}`,
        error?.message || ''
      );

      return null;
    }

    const originalBuffer = Buffer.from(
      await data.arrayBuffer()
    );

    /*
     * Foto diperkecil agar ukuran Excel tidak terlalu besar.
     */
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

    return compressedBuffer;
  } catch (error) {
    console.warn(
      `Gagal memproses foto: ${storagePath}`,
      error.message
    );

    return null;
  }
}

/*
|--------------------------------------------------------------------------
| Mengambil Data Laporan dari Supabase
|--------------------------------------------------------------------------
*/

async function fetchReports(start, end) {
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

  const { data, error } = await supabase
    .from('work_reports')
    .select(query)
    .gte('created_at', start.toISOString())
    .lt('created_at', end.toISOString())
    .order('created_at', {
      ascending: true
    });

  if (error) {
    throw new Error(
      `Gagal mengambil laporan Supabase: ${error.message}`
    );
  }

  return data || [];
}

/*
|--------------------------------------------------------------------------
| Membuat Sheet Maintenance Report
|--------------------------------------------------------------------------
*/

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

  worksheet.columns = headers.map(
    (header, index) => ({
      header,
      key: `column_${index + 1}`,
      width: columnWidths[index]
    })
  );

  /*
   * Pengaturan header.
   */
  worksheet.getRow(1).height = 34;

  worksheet.autoFilter = {
    from: 'A1',
    to: `V${Math.max(reports.length + 1, 1)}`
  };

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
   * Memasukkan laporan ke Excel.
   */
  reports.forEach((report, index) => {
    const row = worksheet.addRow([
      index + 1,
      safeText(report.report_no),
      formatDateTime(report.created_at),
      safeText(report.report_type),
      safeText(
        report.functional_locations?.funloc_code
      ),
      safeText(
        report.functional_locations?.description
      ),
      safeText(
        report.functional_locations?.area
      ),
      safeText(report.priority),
      safeText(report.description),
      safeText(report.operational_impact),
      safeText(report.status),
      safeText(report.creator?.full_name),
      safeText(report.assignee?.full_name),
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

        applyThinBorder(cell);
      }
    );

    /*
     * Warna status.
     */
    const statusCell = row.getCell(11);

    const status = safeText(
      report.status
    ).toUpperCase();

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

  /*
   * Pengaturan halaman cetak.
   */
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

  /*
   * Kolom A, C, dan E digunakan untuk foto.
   * Kolom B dan D digunakan sebagai pemisah.
   */
  worksheet.getColumn('A').width = 36;
  worksheet.getColumn('B').width = 4;
  worksheet.getColumn('C').width = 36;
  worksheet.getColumn('D').width = 4;
  worksheet.getColumn('E').width = 36;

  const categories = [
    {
      type: 'initial',
      label: 'Kondisi Awal',
      column: 1
    },
    {
      type: 'result',
      label: 'Hasil Pekerjaan',
      column: 3
    },
    {
      type: 'verification',
      label: 'Verifikasi',
      column: 5
    }
  ];

  /*
   * Jika tidak ada laporan.
   */
  if (reports.length === 0) {
    worksheet.mergeCells('A1:E3');

    const emptyCell = worksheet.getCell('A1');

    emptyCell.value =
      'Tidak ada laporan pada periode ini.';

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
    let index = 0;
    index < reports.length;
    index += 1
  ) {
    const report = reports[index];

    const startRow = index * 23 + 1;

    const endImageRow = startRow + 20;

    /*
     * Judul laporan.
     */
    worksheet.mergeCells(
      startRow,
      1,
      startRow,
      5
    );

    const reportTitle = worksheet.getCell(
      startRow,
      1
    );

    reportTitle.value =
      `${safeText(report.report_no)} • ` +
      `${safeText(
        report.functional_locations?.funloc_code
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

    worksheet.getRow(startRow).height = 26;

    /*
     * Membuat tiga bagian foto.
     */
    for (const category of categories) {
      const heading = worksheet.getCell(
        startRow + 1,
        category.column
      );

      heading.value = category.label;

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
       * Membuat ruang untuk foto.
       */
      for (
        let rowNumber = startRow + 2;
        rowNumber <= endImageRow;
        rowNumber += 1
      ) {
        worksheet.getRow(rowNumber).height = 18;

        const photoAreaCell = worksheet.getCell(
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

        applyThinBorder(photoAreaCell);
      }

      /*
       * Memilih foto sesuai tipe.
       */
      const categoryPhotos = (
        report.report_photos || []
      )
        .filter(
          (photo) =>
            photo.photo_type === category.type
        )
        .sort(
          (firstPhoto, secondPhoto) =>
            new Date(firstPhoto.created_at) -
            new Date(secondPhoto.created_at)
        );

      /*
       * Versi awal memasukkan maksimal satu foto
       * untuk setiap kategori.
       */
      const selectedPhoto = categoryPhotos[0];

      if (!selectedPhoto) {
        const noPhotoCell = worksheet.getCell(
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
       * Menambahkan gambar ke workbook.
       */
      const imageId = workbook.addImage({
        buffer: imageBuffer,
        extension: 'jpeg'
      });

      /*
       * Menempatkan gambar di kolom yang sesuai.
       */
      worksheet.addImage(imageId, {
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
      });
    }

    /*
     * Catatan laporan di bawah foto.
     */
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
      `Deskripsi: ${safeText(
        report.description
      )} | ` +
      `Status: ${safeText(
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

    worksheet.getRow(
      startRow + 21
    ).height = 28;
  }

  /*
   * Pengaturan cetak dokumentasi.
   */
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
| Proses Utama
|--------------------------------------------------------------------------
*/

async function main() {
  /*
   * Menentukan periode report.
   */
  const {
    start,
    end
  } = getReportPeriod();

  console.log(
    `Periode laporan: ` +
    `${start.toISOString()} sampai ` +
    `${end.toISOString()}`
  );

  /*
   * Mengambil data Supabase.
   */
  const reports = await fetchReports(
    start,
    end
  );

  console.log(
    `Jumlah laporan ditemukan: ` +
    `${reports.length}`
  );

  /*
   * Membuat workbook Excel.
   */
  const workbook = new ExcelJS.Workbook();

  workbook.creator =
    'PIM Maintenance Hub';

  workbook.company =
    'PT Padi Indonesia Maju';

  workbook.created = new Date();

  workbook.modified = new Date();

  /*
   * Membuat sheet laporan dan dokumentasi.
   */
  createMaintenanceSheet(
    workbook,
    reports
  );

  await createDocumentationSheet(
    workbook,
    reports
  );

  /*
   * Menentukan nama file report.
   */
  const reportDate =
    formatDateForFile(start);

  const fileName =
    `PIM_PM_CM_Report_${reportDate}.xlsx`;

  /*
   * Mengubah workbook menjadi buffer
   * untuk dilampirkan ke email.
   */
  const excelBuffer =
    await workbook.xlsx.writeBuffer();

  /*
   * Menghitung ringkasan.
   */
  const totalPM = reports.filter(
    (report) =>
      report.report_type === 'PM'
  ).length;

  const totalCM = reports.filter(
    (report) =>
      report.report_type === 'CM'
  ).length;

  const totalClosed = reports.filter(
    (report) =>
      safeText(
        report.status
      ).toUpperCase() === 'CLOSED'
  ).length;

  const totalOpen =
    reports.length - totalClosed;

  /*
   * Memastikan koneksi Gmail berhasil.
   */
  console.log(
    'Memeriksa koneksi Gmail SMTP...'
  );

  await transporter.verify();

  console.log(
    'Koneksi Gmail SMTP berhasil.'
  );

  /*
   * Menyiapkan alamat penerima.
   */
  const emailRecipients =
    process.env.REPORT_EMAIL_TO
      .split(',')
      .map((email) => email.trim())
      .filter(Boolean)
      .join(',');

  /*
   * Mengirim email.
   */
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
          Berikut kami sampaikan laporan
          pekerjaan Preventive Maintenance
          dan Corrective Maintenance
          periode pukul 07.30 WIB.
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
          dokumentasi foto terlampir
          pada email ini.
        </p>

        <p>
          Email ini dikirim otomatis oleh
          PIM Maintenance Hub.
        </p>
      `,

      attachments: [
        {
          filename: fileName,

          content:
            Buffer.from(excelBuffer),

          contentType:
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
      ]
    });

  /*
   * Menampilkan informasi keberhasilan
   * di log GitHub Actions.
   */
  console.log(
    `Email berhasil dikirim dari ` +
    `${process.env.GMAIL_USER}`
  );

  console.log(
    `Penerima: ${emailRecipients}`
  );

  console.log(
    `Message ID: ${result.messageId}`
  );

  console.log(
    `Lampiran: ${fileName}`
  );
}

/*
|--------------------------------------------------------------------------
| Menjalankan Proses
|--------------------------------------------------------------------------
*/

main().catch((error) => {
  console.error(
    'Gagal membuat atau mengirim laporan:'
  );

  console.error(error);

  process.exit(1);
});

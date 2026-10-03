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

  return [...reports].sort((first, second) => {
    return (
      new Date(second.created_at) -
      new Date(first.created_at)
    );
  })[0];
}

function getLatestValue(reports, fieldName) {
  const sortedReports = [...(reports || [])].sort(
    (first, second) => {
      return (
        new Date(second.created_at) -
        new Date(first.created_at)
      );
    }
  );

  for (const report of sortedReports) {
    const value = report[fieldName];

    if (
      value !== null &&
      value !== undefined &&
      safeText(value).trim() !== ''
    ) {
      return value;
    }
  }

  return '';
}

/* =========================================================
   MENGAMBIL DATA SUPABASE
========================================================= */

async function fetchAllData() {
  console.log('========================================');
  console.log(`Supabase URL: ${process.env.SUPABASE_URL}`);
  console.log(`Bucket foto: ${PHOTO_BUCKET}`);
  console.log(
    `Mengambil maksimal ${MAX_REPORTS} laporan terbaru`
  );

  const reportsResult = await supabase
    .from('work_reports')
    .select('*')
    .order('created_at', {
      ascending: false
    })
    .limit(MAX_REPORTS);

  if (reportsResult.error) {
    throw new Error(
      `Gagal mengambil work_reports: ` +
      `${reportsResult.error.message}`
    );
  }

  const funlocResult = await supabase
    .from('functional_locations')
    .select('*');

  if (funlocResult.error) {
    throw new Error(
      `Gagal mengambil functional_locations: ` +
      `${funlocResult.error.message}`
    );
  }

  const profilesResult = await supabase
    .from('profiles')
    .select('id, full_name, role, is_active');

  if (profilesResult.error) {
    throw new Error(
      `Gagal mengambil profiles: ` +
      `${profilesResult.error.message}`
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
      `Gagal mengambil report_photos: ` +
      `${photosResult.error.message}`
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

  const mergedReports = reports.map((report) => {
    const functionalLocation =
      functionalLocations.find(
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
      functional_location:
        functionalLocation || null,
      creator_profile:
        creator || null,
      technician_profile:
        technician || null,
      report_photos:
        reportPhotos
    };
  });

  for (const report of mergedReports) {
    console.log(
      `${safeText(report.report_no)}: ` +
      `${report.report_photos.length} foto`
    );
  }

  return mergedReports;
}

/* =========================================================
   KATEGORI FOTO
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
   MEMBENTUK RANGKAIAN PM DAN CM
========================================================= */

function createWorkGroup(pmReport, cmReports) {
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
    group.cmReports.flatMap((report) => {
      return report.report_photos || [];
    });

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

  const reportNumbers = [];

  if (group.pmReport?.report_no) {
    reportNumbers.push(
      group.pmReport.report_no
    );
  }

  for (const cmReport of group.cmReports) {
    if (cmReport.report_no) {
      reportNumbers.push(
        cmReport.report_no
      );
    }
  }

  group.reportNumbers =
    reportNumbers.join(' → ');

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

function buildWorkGroups(reports) {
  const pmReports = reports.filter(
    (report) => {
      return (
        safeText(report.report_type)
          .trim()
          .toUpperCase() === 'PM'
      );
    }
  );

  const cmReports = reports.filter(
    (report) => {
      return (
        safeText(report.report_type)
          .trim()
          .toUpperCase() === 'CM'
      );
    }
  );

  const groups = [];
  const usedCmIds = new Set();

  /*
   * Tahap pertama:
   * Hubungkan CM berdasarkan pm_reference.
   */

  for (const pmReport of pmReports) {
    const relatedCmReports = cmReports.filter(
      (cmReport) => {
        return (
          cmReport.pm_reference ===
          pmReport.id
        );
      }
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
   * Tahap kedua:
   * Jika CM tidak memiliki pm_reference,
   * cari PM terdekat di Functional Location yang sama.
   */

  for (const cmReport of cmReports) {
    if (usedCmIds.has(cmReport.id)) {
      continue;
    }

    let matchingPm = null;

    if (cmReport.funloc_id) {
      const candidates = pmReports
        .filter((pmReport) => {
          return (
            pmReport.funloc_id ===
            cmReport.funloc_id
          );
        })
        .sort((firstPm, secondPm) => {
          const cmTimestamp =
            new Date(
              cmReport.created_at
            ).getTime();

          const firstDifference =
            Math.abs(
              cmTimestamp -
              new Date(
                firstPm.created_at
              ).getTime()
            );

          const secondDifference =
            Math.abs(
              cmTimestamp -
              new Date(
                secondPm.created_at
              ).getTime()
            );

          return (
            firstDifference -
            secondDifference
          );
        });

      matchingPm =
        candidates[0] || null;
    }

    if (matchingPm) {
      const existingGroup =
        groups.find((group) => {
          return (
            group.pmReport?.id ===
            matchingPm.id
          );
        });

      if (existingGroup) {
        existingGroup.cmReports.push(
          cmReport
        );

        refreshWorkGroup(
          existingGroup
        );
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

  return groups.sort(
    (firstGroup, secondGroup) => {
      return (
        new Date(
          secondGroup.createdAt
        ).getTime() -
        new Date(
          firstGroup.createdAt
        ).getTime()
      );
    }
  );
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

    const downloadResult =
      await supabase.storage
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
      Buffer.from(
        await downloadResult.data
          .arrayBuffer()
      );

    const compressedBuffer =
      await sharp(originalBuffer)
        .rotate()
        .resize({
          width: 420,
          height: 300,
          fit: 'inside',
          withoutEnlargement: true
        })
        .jpeg({
          quality: 58,
          mozjpeg: true
        })
        .toBuffer();

    console.log(
      `Foto berhasil diproses: ${storagePath}`
    );

    return compressedBuffer;
  } catch (error) {
    console.warn(
      `Gagal memproses foto ${storagePath}: ` +
      `${error.message}`
    );

    return null;
  }
}

/* =========================================================
   SHEET MAINTENANCE REPORT
========================================================= */

function createMaintenanceSheet(
  workbook,
  groups
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

  const columns = [
    {
      header: 'No.',
      width: 7
    },
    {
      header: 'Nomor PM / CM',
      width: 42
    },
    {
      header: 'Tanggal',
      width: 22
    },
    {
      header: 'Functional Location',
      width: 42
    },
    {
      header: 'Deskripsi Lokasi',
      width: 30
    },
    {
      header: 'Area',
      width: 18
    },
    {
      header: 'Prioritas',
      width: 14
    },
    {
      header: 'Deskripsi / Temuan',
      width: 40
    },
    {
      header: 'Rekomendasi PM',
      width: 35
    },
    {
      header: 'Status Akhir',
      width: 20
    },
    {
      header: 'Pelapor',
      width: 22
    },
    {
      header: 'Teknisi',
      width: 22
    },
    {
      header: 'Tindakan Perbaikan',
      width: 40
    },
    {
      header: 'Spare Part',
      width: 28
    },
    {
      header: 'Catatan Teknisi',
      width: 34
    },
    {
      header: 'Hasil Test',
      width: 22
    },
    {
      header: 'Catatan Verifikasi',
      width: 34
    }
  ];

  worksheet.columns =
    columns.map(
      (column, index) => ({
        header: column.header,
        key:
          `column_${index + 1}`,
        width: column.width
      })
    );

  worksheet.getRow(1).height = 34;

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

  if (groups.length === 0) {
    worksheet.mergeCells('A2:Q4');

    const emptyCell =
      worksheet.getCell('A2');

    emptyCell.value =
      'Belum ada laporan pada database.';

    emptyCell.alignment = {
      horizontal: 'center',
      vertical: 'middle'
    };

    return worksheet;
  }

  groups.forEach(
    (group, index) => {
      const dateValue =
        group.pmReport?.created_at ||
        group.cmReports[0]
          ?.created_at ||
        '';

      const row =
        worksheet.addRow([
          index + 1,

          safeText(
            group.reportNumbers
          ),

          formatDateTime(
            dateValue
          ),

          safeText(
            group
              .functionalLocation
              ?.funloc_code
          ),

          safeText(
            group
              .functionalLocation
              ?.description
          ),

          safeText(
            group
              .functionalLocation
              ?.area
          ),

          safeText(
            group.priority
          ),

          safeText(
            group.description
          ),

          safeText(
            group.recommendation
          ),

          safeText(
            group.status
          ),

          safeText(
            group.reporterName
          ),

          safeText(
            group.technicianName
          ),

          safeText(
            group.correctiveAction
          ),

          safeText(
            group.sparePart
          ),

          safeText(
            group.technicianNote
          ),

          safeText(
            group.testResult
          ),

          safeText(
            group.verificationNote
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

          applyBorder(cell);
        }
      );

      const statusCell =
        row.getCell(10);

      const status =
        safeText(group.status)
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

  worksheet.autoFilter = {
    from: 'A1',
    to:
      `Q${groups.length + 1}`
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
      top: 0.4,
      bottom: 0.4,
      header: 0.15,
      footer: 0.15
    }
  };

  return worksheet;
}

/* =========================================================
   SHEET DOKUMENTASI
========================================================= */

async function createDocumentationSheet(
  workbook,
  groups
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

  worksheet
    .getColumn('A')
    .width = 34;

  worksheet
    .getColumn('B')
    .width = 4;

  worksheet
    .getColumn('C')
    .width = 34;

  worksheet
    .getColumn('D')
    .width = 4;

  worksheet
    .getColumn('E')
    .width = 34;

  if (groups.length === 0) {
    worksheet.mergeCells('A1:E4');

    const emptyCell =
      worksheet.getCell('A1');

    emptyCell.value =
      'Belum ada dokumentasi pekerjaan.';

    emptyCell.alignment = {
      horizontal: 'center',
      vertical: 'middle'
    };

    return worksheet;
  }

  const categories = [
    {
      key: 'initialPhotos',
      label: 'Kondisi Awal',
      column: 1
    },
    {
      key: 'resultPhotos',
      label: 'Hasil Pekerjaan',
      column: 3
    },
    {
      key: 'verificationPhotos',
      label: 'Verifikasi',
      column: 5
    }
  ];

  /*
   * Setiap rangkaian menggunakan 13 baris.
   */

  for (
    let groupIndex = 0;
    groupIndex < groups.length;
    groupIndex += 1
  ) {
    const group =
      groups[groupIndex];

    const startRow =
      groupIndex * 13 + 1;

    const headerRow =
      startRow + 1;

    const photoStartRow =
      startRow + 2;

    const photoEndRow =
      startRow + 10;

    const informationRow =
      startRow + 11;

    worksheet.mergeCells(
      startRow,
      1,
      startRow,
      5
    );

    const titleCell =
      worksheet.getCell(
        startRow,
        1
      );

    titleCell.value =
      `${safeText(
        group.reportNumbers
      )} • ` +
      `${safeText(
        group
          .functionalLocation
          ?.funloc_code
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
      size: 11,
      color: {
        argb: 'FFFFFFFF'
      }
    };

    titleCell.alignment = {
      vertical: 'middle'
    };

    worksheet
      .getRow(startRow)
      .height = 24;

    for (
      const category
      of categories
    ) {
      const headerCell =
        worksheet.getCell(
          headerRow,
          category.column
        );

      headerCell.value =
        category.label;

      headerCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {
          argb: 'FF367F78'
        }
      };

      headerCell.font = {
        bold: true,
        size: 10,
        color: {
          argb: 'FFFFFFFF'
        }
      };

      headerCell.alignment = {
        horizontal: 'center',
        vertical: 'middle'
      };

      worksheet
        .getRow(headerRow)
        .height = 20;

      for (
        let rowNumber =
          photoStartRow;
        rowNumber <=
          photoEndRow;
        rowNumber += 1
      ) {
        worksheet
          .getRow(rowNumber)
          .height = 16;

        const areaCell =
          worksheet.getCell(
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

      const photos =
        group[category.key] || [];

      console.log(
        `${group.reportNumbers} | ` +
        `${category.label}: ` +
        `${photos.length} foto`
      );

      const selectedPhoto =
        photos[0];

      if (!selectedPhoto) {
        const emptyPhotoCell =
          worksheet.getCell(
            startRow + 6,
            category.column
          );

        emptyPhotoCell.value =
          'Foto belum tersedia';

        emptyPhotoCell.font = {
          italic: true,
          size: 9,
          color: {
            argb: 'FF64748B'
          }
        };

        emptyPhotoCell.alignment = {
          horizontal: 'center',
          vertical: 'middle'
        };

        continue;
      }

      const imageBuffer =
        await downloadPhoto(
          selectedPhoto.storage_path
        );

      if (!imageBuffer) {
        const errorPhotoCell =
          worksheet.getCell(
            startRow + 6,
            category.column
          );

        errorPhotoCell.value =
          'Foto tidak dapat dimuat';

        errorPhotoCell.font = {
          italic: true,
          size: 9,
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
              0.08,

            row:
              photoStartRow -
              1 +
              0.12
          },

          ext: {
            width: 215,
            height: 145
          },

          editAs: 'oneCell'
        }
      );
    }

    worksheet.mergeCells(
      informationRow,
      1,
      informationRow,
      5
    );

    const informationCell =
      worksheet.getCell(
        informationRow,
        1
      );

    informationCell.value =
      `Temuan: ` +
      `${safeText(
        group.description
      )} | ` +
      `Status Akhir: ` +
      `${safeText(
        group.status
      )}`;

    informationCell.font = {
      size: 9,
      color: {
        argb: 'FF475569'
      }
    };

    informationCell.alignment = {
      vertical: 'middle',
      wrapText: true
    };

    worksheet
      .getRow(informationRow)
      .height = 24;

    worksheet
      .getRow(startRow + 12)
      .height = 8;
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
      top: 0.4,
      bottom: 0.4,
      header: 0.15,
      footer: 0.15
    }
  };

  return worksheet;
}

/* =========================================================
   MENGIRIM EMAIL
========================================================= */

async function sendEmail(
  workbook,
  groups
) {
  const reportDate =
    formatDateForFile();

  const fileName =
    `PIM_PM_CM_Report_` +
    `${reportDate}.xlsx`;

  console.log(
    'Membuat file Excel'
  );

  const excelBuffer =
    await workbook.xlsx
      .writeBuffer();

  const fileSizeKb =
    Math.round(
      excelBuffer.byteLength /
      1024
    );

  console.log(
    `Ukuran Excel: ` +
    `${fileSizeKb} KB`
  );

  const totalPm =
    groups.filter(
      (group) =>
        Boolean(group.pmReport)
    ).length;

  const totalCm =
    groups.reduce(
      (total, group) => {
        return (
          total +
          group.cmReports.length
        );
      },
      0
    );

  const totalClosed =
    groups.filter(
      (group) => {
        return (
          safeText(group.status)
            .toUpperCase() ===
          'CLOSED'
        );
      }
    ).length;

  const totalOpen =
    groups.length -
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
    'Memeriksa koneksi Gmail SMTP'
  );

  await transporter.verify();

  console.log(
    'Koneksi Gmail SMTP berhasil'
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
          Berikut laporan pekerjaan
          Preventive Maintenance dan
          Corrective Maintenance.
        </p>

        <h3>Ringkasan Laporan</h3>

        <ul>
          <li>
            Total rangkaian pekerjaan:
            <b>${groups.length}</b>
          </li>

          <li>
            Preventive Maintenance:
            <b>${totalPm}</b>
          </li>

          <li>
            Corrective Maintenance:
            <b>${totalCm}</b>
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
          File Excel beserta dokumentasi
          foto terlampir pada email ini.
        </p>

        <p>
          Email dikirim otomatis oleh
          PIM Maintenance Hub.
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

/* =========================================================
   PROSES UTAMA
========================================================= */

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

  const reports =
    await fetchAllData();

  const groups =
    buildWorkGroups(reports);

  const totalPhotos =
    groups.reduce(
      (total, group) => {
        return (
          total +
          group.initialPhotos.length +
          group.resultPhotos.length +
          group.verificationPhotos.length
        );
      },
      0
    );

  console.log(
    `TOTAL DATA ASLI: ` +
    `${reports.length} laporan`
  );

  console.log(
    `TOTAL RANGKAIAN: ` +
    `${groups.length} baris`
  );

  console.log(
    `TOTAL FOTO: ` +
    `${totalPhotos} foto`
  );

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

  createMaintenanceSheet(
    workbook,
    groups
  );

  await createDocumentationSheet(
    workbook,
    groups
  );

  await sendEmail(
    workbook,
    groups
  );

  console.log(
    '========================================'
  );

  console.log(
    'Proses report selesai'
  );

  console.log(
    '========================================'
  );
}

main().catch((error) => {
  console.error(
    '========================================'
  );

  console.error(
    'Gagal membuat atau mengirim laporan'
  );

  console.error(error);

  console.error(
    '========================================'
  );

  process.exit(1);
});

/* =========================================================
   PIM MAINTENANCE HUB
   MANUAL XLSX REPORT DENGAN DOKUMENTASI FOTO
========================================================= */

const MANUAL_REPORT_BUCKET = 'maintenance-photos';
const MANUAL_REPORT_LIMIT = 100;

function manualSafeText(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value);
}

function manualFormatDateTime(value) {
  if (!value) {
    return '';
  }

  try {
    return new Intl.DateTimeFormat('id-ID', {
      timeZone: 'Asia/Jakarta',
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false
    }).format(new Date(value));
  } catch {
    return manualSafeText(value);
  }
}

function manualDateForFile() {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Jakarta',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).formatToParts(new Date());

  const result = {};

  for (const part of parts) {
    result[part.type] = part.value;
  }

  return `${result.year}-${result.month}-${result.day}`;
}

function manualApplyBorder(cell) {
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

function manualShowMessage(message) {
  if (typeof toast === 'function') {
    toast(message);
    return;
  }

  alert(message);
}

/* =========================================================
   MENGAMBIL DATA SUPABASE
========================================================= */

async function manualFetchAllData() {
  if (typeof db === 'undefined' || !db) {
    throw new Error(
      'Koneksi Supabase belum tersedia. Silakan login kembali.'
    );
  }

  const sessionResult = await db.auth.getSession();

  if (
    sessionResult.error ||
    !sessionResult.data.session
  ) {
    throw new Error(
      'Sesi login sudah berakhir. Silakan login kembali.'
    );
  }

  const reportsResult = await db
    .from('work_reports')
    .select('*')
    .order('created_at', {
      ascending: false
    })
    .limit(MANUAL_REPORT_LIMIT);

  if (reportsResult.error) {
    throw new Error(
      `Gagal mengambil laporan: ${reportsResult.error.message}`
    );
  }

  const funlocResult = await db
    .from('functional_locations')
    .select('*');

  if (funlocResult.error) {
    throw new Error(
      `Gagal mengambil Functional Location: ${funlocResult.error.message}`
    );
  }

  const profilesResult = await db
    .from('profiles')
    .select('id, full_name, role, is_active');

  if (profilesResult.error) {
    throw new Error(
      `Gagal mengambil profil: ${profilesResult.error.message}`
    );
  }

  const photosResult = await db
    .from('report_photos')
    .select('*')
    .order('created_at', {
      ascending: true
    });

  if (photosResult.error) {
    throw new Error(
      `Gagal mengambil metadata foto: ${photosResult.error.message}`
    );
  }

  const reports = reportsResult.data || [];
  const funlocs = funlocResult.data || [];
  const profiles = profilesResult.data || [];
  const photos = photosResult.data || [];

  return reports.map((report) => {
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
}

/* =========================================================
   KONVERSI FOTO UNTUK EXCEL
========================================================= */

async function manualBlobToJpegBase64(blob) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    const objectUrl = URL.createObjectURL(blob);

    image.onload = () => {
      try {
        const maximumWidth = 900;
        const maximumHeight = 650;

        let width = image.naturalWidth;
        let height = image.naturalHeight;

        const scale = Math.min(
          maximumWidth / width,
          maximumHeight / height,
          1
        );

        width = Math.round(width * scale);
        height = Math.round(height * scale);

        const canvas = document.createElement('canvas');

        canvas.width = width;
        canvas.height = height;

        const context = canvas.getContext('2d');

        context.fillStyle = '#ffffff';
        context.fillRect(0, 0, width, height);

        context.drawImage(
          image,
          0,
          0,
          width,
          height
        );

        const dataUrl = canvas.toDataURL(
          'image/jpeg',
          0.68
        );

        URL.revokeObjectURL(objectUrl);

        resolve(dataUrl);
      } catch (error) {
        URL.revokeObjectURL(objectUrl);
        reject(error);
      }
    };

    image.onerror = () => {
      URL.revokeObjectURL(objectUrl);

      reject(
        new Error('File foto tidak dapat dibaca.')
      );
    };

    image.src = objectUrl;
  });
}

async function manualDownloadPhoto(storagePath) {
  if (!storagePath) {
    return null;
  }

  try {
    const downloadResult = await db.storage
      .from(MANUAL_REPORT_BUCKET)
      .download(storagePath);

    if (
      downloadResult.error ||
      !downloadResult.data
    ) {
      console.warn(
        'Foto tidak dapat diunduh:',
        storagePath,
        downloadResult.error?.message
      );

      return null;
    }

    return await manualBlobToJpegBase64(
      downloadResult.data
    );
  } catch (error) {
    console.warn(
      'Foto tidak dapat diproses:',
      storagePath,
      error.message
    );

    return null;
  }
}

/* =========================================================
   KATEGORI FOTO
========================================================= */

const MANUAL_PHOTO_CATEGORIES = [
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

function manualGetPhotosByCategory(
  report,
  category
) {
  return (report.report_photos || [])
    .filter((photo) => {
      const photoType = manualSafeText(
        photo.photo_type
      )
        .trim()
        .toLowerCase();

      return category.types.includes(photoType);
    })
    .sort((firstPhoto, secondPhoto) => {
      return (
        new Date(firstPhoto.created_at) -
        new Date(secondPhoto.created_at)
      );
    });
}

/* =========================================================
   SHEET MAINTENANCE REPORT
========================================================= */

function manualCreateMaintenanceSheet(
  workbook,
  reports
) {
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

  worksheet.columns = columns.map(
    (column, index) => ({
      header: column.header,
      key: `column_${index + 1}`,
      width: column.width
    })
  );

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

    manualApplyBorder(cell);
  });

  if (reports.length === 0) {
    worksheet.mergeCells('A2:V4');

    const emptyCell = worksheet.getCell('A2');

    emptyCell.value =
      'Belum ada laporan pada database.';

    emptyCell.alignment = {
      horizontal: 'center',
      vertical: 'middle'
    };

    return worksheet;
  }

  reports.forEach((report, index) => {
    const row = worksheet.addRow([
      index + 1,
      manualSafeText(report.report_no),
      manualFormatDateTime(report.created_at),
      manualSafeText(report.report_type),
      manualSafeText(
        report.functional_location?.funloc_code
      ),
      manualSafeText(
        report.functional_location?.description
      ),
      manualSafeText(
        report.functional_location?.area
      ),
      manualSafeText(report.priority),
      manualSafeText(report.description),
      manualSafeText(report.operational_impact),
      manualSafeText(report.status),
      manualSafeText(
        report.creator_profile?.full_name
      ),
      manualSafeText(
        report.technician_profile?.full_name
      ),
      manualSafeText(report.failure_cause),
      manualSafeText

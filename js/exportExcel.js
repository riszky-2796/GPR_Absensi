import { EXCEL_COLORS, MONTH_NAMES } from './config.js';
import { formatMonthLabel, parseMonthKey } from './calculator.js';

const THIN_BORDER = {
  top: { style: 'thin', color: { argb: EXCEL_COLORS.BORDER_BLACK } },
  left: { style: 'thin', color: { argb: EXCEL_COLORS.BORDER_BLACK } },
  bottom: { style: 'thin', color: { argb: EXCEL_COLORS.BORDER_BLACK } },
  right: { style: 'thin', color: { argb: EXCEL_COLORS.BORDER_BLACK } },
};

/**
 * Menghasilkan dan mengunduh file .xlsx rekap bulanan berwarna menggunakan ExcelJS
 * sesuai spesifikasi visual PRD Section 7.7.
 *
 * @param {{
 *   className: string,
 *   monthKey: string,
 *   categories: import('./state.js').Category[],
 *   rows: import('./calculator.js').RecapRow[],
 *   classSummary: import('./calculator.js').ClassSummary
 * }} params
 */
export async function exportMonthlyRecapToExcel({
  className,
  monthKey,
  categories,
  rows,
  classSummary,
}) {
  if (typeof window.ExcelJS === 'undefined') {
    throw new Error('Library ExcelJS belum termuat. Periksa koneksi internet lalu muat ulang halaman.');
  }

  const sortedCategories = categories.slice().sort((a, b) => a.order_index - b.order_index);
  const workbook = new window.ExcelJS.Workbook();
  const monthLabel = formatMonthLabel(monthKey);
  const worksheet = workbook.addWorksheet(`Rekap ${monthLabel}`.slice(0, 31));

  // Susun daftar header baku sesuai PRD
  const headers = [
    'NO',
    'NAMA',
    ...sortedCategories.map((c) => c.name),
    'HADIR',
    'IZIN',
    'SAKIT',
    'ALFA',
    'TOTAL PERTEMUAN',
    'PRESENTASI KEHADIRAN',
    'PRESENTASI IZIN',
    'PRESENTASI SAKIT',
    'PRESENTASI ALFA',
  ];

  const totalCols = headers.length;
  const colNameIndex = 2;
  const colAlfaPctIndex = totalCols;

  // Baris 1: Judul Rekap (Merged dari kolom 1 sampai kolom terakhir)
  const titleText = `REKAP ABSENSI ${String(className || 'GPR ABSENSI').toUpperCase()} BULAN ${monthLabel}`;
  const titleRow = worksheet.addRow([titleText]);
  worksheet.mergeCells(1, 1, 1, totalCols);
  titleRow.height = 28;

  const titleCell = titleRow.getCell(1);
  titleCell.font = { name: 'Calibri', size: 12, bold: true, color: { argb: EXCEL_COLORS.TEXT_BLACK } };
  titleCell.alignment = { vertical: 'middle', horizontal: 'center' };

  // Baris 2: Header Tabel
  const headerRow = worksheet.addRow(headers);
  headerRow.height = 32;
  headerRow.eachCell((cell) => {
    cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: EXCEL_COLORS.TEXT_BLACK } };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: EXCEL_COLORS.HEADER_BG },
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = THIN_BORDER;
  });

  // Baris 3..N: Data Siswa (Zebra Biru-Putih + Highlight Merah jika PRESENTASI ALFA > 50%)
  let maxNameLength = 20;

  rows.forEach((row, idx) => {
    if (row.name.length > maxNameLength) {
      maxNameLength = row.name.length;
    }

    const rowValues = [
      row.no,
      row.name,
      ...sortedCategories.map((c) => row.categoryValues[c.category_id] || '-'),
      row.hadir,
      row.izin,
      row.sakit,
      row.alfa,
      row.totalPertemuan,
      row.presentasiKehadiran / 100,
      row.presentasiIzin / 100,
      row.presentasiSakit / 100,
      row.presentasiAlfa / 100,
    ];

    const excelRow = worksheet.addRow(rowValues);
    excelRow.height = 21;

    const zebraBg = idx % 2 === 0 ? EXCEL_COLORS.ZEBRA_ODD_BG : EXCEL_COLORS.ZEBRA_EVEN_BG;

    excelRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      cell.font = { name: 'Calibri', size: 11, color: { argb: EXCEL_COLORS.TEXT_BLACK } };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: zebraBg },
      };
      cell.alignment = {
        vertical: 'middle',
        horizontal: colNumber === colNameIndex ? 'left' : 'center',
      };
      cell.border = THIN_BORDER;

      if (colNumber >= totalCols - 3) {
        cell.numFmt = '0%';
      }
    });

    // Aturan Label Merah: hanya sel NAMA dan sel PRESENTASI ALFA jika PRESENTASI ALFA > 50%
    if (row.isWarningRed) {
      const redFill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: EXCEL_COLORS.WARNING_RED_BG },
      };
      excelRow.getCell(colNameIndex).fill = redFill;
      excelRow.getCell(colAlfaPctIndex).fill = redFill;
    }
  });

  // Baris Terakhir: PRESENTASI KELAS
  const mergeEndCol = totalCols - 4; // sampai kolom TOTAL PERTEMUAN
  const classRowValues = new Array(totalCols).fill('');
  classRowValues[0] = 'PRESENTASI KELAS';
  classRowValues[totalCols - 4] = classSummary.presentasiKehadiran / 100;
  classRowValues[totalCols - 3] = classSummary.presentasiIzin / 100;
  classRowValues[totalCols - 2] = classSummary.presentasiSakit / 100;
  classRowValues[totalCols - 1] = classSummary.presentasiAlfa / 100;

  const classRow = worksheet.addRow(classRowValues);
  classRow.height = 24;
  worksheet.mergeCells(classRow.number, 1, classRow.number, mergeEndCol);

  classRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: EXCEL_COLORS.TEXT_BLACK } };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: EXCEL_COLORS.ZEBRA_EVEN_BG },
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = THIN_BORDER;

    if (colNumber >= totalCols - 3) {
      cell.numFmt = '0%';
    }
  });

  // Pengaturan Lebar Kolom Otomatis (Auto-fit dengan batas minimum agar teks tidak terpotong)
  worksheet.columns = headers.map((header, i) => {
    if (i === 0) return { width: 6 }; // NO
    if (i === 1) return { width: Math.min(Math.max(maxNameLength + 4, 24), 42) }; // NAMA
    if (i >= 2 && i < 2 + sortedCategories.length) {
      return { width: Math.max(header.length + 4, 14) }; // Kategori Dinamis
    }
    if (header === 'TOTAL PERTEMUAN') return { width: 18 };
    if (header.startsWith('PRESENTASI')) return { width: 16 };
    return { width: 10 }; // HADIR, IZIN, SAKIT, ALFA
  });

  // Unduh File di Browser
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });

  const { year, monthIndex } = parseMonthKey(monthKey);
  const monthName = MONTH_NAMES[monthIndex] || monthKey;
  const safeClass = String(className || 'GPR_Absensi')
    .trim()
    .replace(/[^a-zA-Z0-9_-]+/g, '_');
  const fileName = `Rekap_Absensi_${safeClass}_${monthName}_${year}.xlsx`;

  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
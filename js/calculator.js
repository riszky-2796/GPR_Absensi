import {
  ALFA_WARNING_THRESHOLD,
  DAY_SHORT_NAMES,
  MONTH_NAMES,
  STATUS,
} from './config.js';

/**
 * @typedef {Object} RecapRow
 * @property {number} no
 * @property {string} student_id
 * @property {string} name
 * @property {Record<string, string>} categoryValues
 * @property {number} hadir
 * @property {number} izin
 * @property {number} sakit
 * @property {number} alfa
 * @property {number} totalPertemuan
 * @property {number} presentasiKehadiran
 * @property {number} presentasiIzin
 * @property {number} presentasiSakit
 * @property {number} presentasiAlfa
 * @property {boolean} isWarningRed
 */

/**
 * @typedef {Object} ClassSummary
 * @property {number} totalHadir
 * @property {number} totalIzin
 * @property {number} totalSakit
 * @property {number} totalAlfa
 * @property {number} totalPertemuan
 * @property {number} presentasiKehadiran
 * @property {number} presentasiIzin
 * @property {number} presentasiSakit
 * @property {number} presentasiAlfa
 */

/**
 * Memecah string YYYY-MM menjadi angka tahun dan index bulan (0-11).
 * @param {string} monthKey
 * @returns {{ year: number, monthIndex: number }}
 */
export function parseMonthKey(monthKey) {
  const [yStr, mStr] = String(monthKey).split('-');
  return {
    year: Number(yStr),
    monthIndex: Number(mStr) - 1,
  };
}

/**
 * Mengubah YYYY-MM menjadi label kapital Bahasa Indonesia (contoh: "OKTOBER 2026").
 * @param {string} monthKey
 * @returns {string}
 */
export function formatMonthLabel(monthKey) {
  const { year, monthIndex } = parseMonthKey(monthKey);
  const name = MONTH_NAMES[monthIndex] || '';
  return `${name} ${year}`;
}

/**
 * Menggeser bulan mundur (-1) atau maju (+1) dari YYYY-MM.
 * @param {string} monthKey
 * @param {number} delta
 * @returns {string}
 */
export function shiftMonthKey(monthKey, delta) {
  const { year, monthIndex } = parseMonthKey(monthKey);
  const shifted = new Date(year, monthIndex + delta, 1);
  const y = shifted.getFullYear();
  const m = String(shifted.getMonth() + 1).padStart(2, '0');
  return `${y}-${m}`;
}

/**
 * Mengubah YYYY-MM-DD menjadi label singkat hari dan tanggal (aman zona waktu).
 * @param {string} isoDate
 * @returns {{ dayShort: string, dayNum: string, fullLabel: string }}
 */
export function formatDateHeader(isoDate) {
  const [y, m, d] = String(isoDate).split('-').map(Number);
  const dateObj = new Date(y, m - 1, d);
  const dayShort = DAY_SHORT_NAMES[dateObj.getDay()] || '';
  const dayNum = String(d).padStart(2, '0');
  return {
    dayShort,
    dayNum,
    fullLabel: `${dayShort}, ${dayNum}/${String(m).padStart(2, '0')}`,
  };
}

/**
 * Men-generate seluruh tanggal pada bulan YYYY-MM sesuai pola hari rutin (0=Minggu..6=Sabtu).
 * @param {string} monthKey
 * @param {number[]} defaultDays
 * @returns {string[]}
 */
export function generateRoutineDates(monthKey, defaultDays) {
  if (!Array.isArray(defaultDays) || defaultDays.length === 0) return [];
  const { year, monthIndex } = parseMonthKey(monthKey);
  const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
  const daySet = new Set(defaultDays.map(Number));
  const mm = String(monthIndex + 1).padStart(2, '0');

  const result = [];
  for (let d = 1; d <= daysInMonth; d++) {
    const dt = new Date(year, monthIndex, d);
    if (daySet.has(dt.getDay())) {
      result.push(`${year}-${mm}-${String(d).padStart(2, '0')}`);
    }
  }
  return result;
}

/**
 * Menentukan daftar tanggal aktif final untuk satu bulan.
 * Jika bulan sudah tersimpan di sheet Schedules, gunakan tanggal tersimpan agar perubahan
 * pola rutin di masa depan tidak menimpa riwayat bulan lalu.
 *
 * @param {string} monthKey
 * @param {number[]} defaultDays
 * @param {{ dates: string[], added_dates: string[], removed_dates: string[] } | null} savedSchedule
 * @returns {{ dates: string[], added_dates: string[], removed_dates: string[], isSaved: boolean }}
 */
export function resolveMonthDates(monthKey, defaultDays, savedSchedule) {
  if (savedSchedule && Array.isArray(savedSchedule.dates)) {
    return {
      dates: savedSchedule.dates.slice().sort(),
      added_dates: Array.isArray(savedSchedule.added_dates) ? savedSchedule.added_dates.slice() : [],
      removed_dates: Array.isArray(savedSchedule.removed_dates) ? savedSchedule.removed_dates.slice() : [],
      isSaved: true,
    };
  }

  return {
    dates: generateRoutineDates(monthKey, defaultDays),
    added_dates: [],
    removed_dates: [],
    isSaved: false,
  };
}

/**
 * Menentukan daftar siswa yang tampil pada bulan tertentu:
 * seluruh siswa aktif + siswa non-aktif yang pernah memiliki catatan absensi di bulan tersebut.
 *
 * @param {import('./state.js').Student[]} allStudents
 * @param {Record<string, string>} monthAttendanceMap
 * @returns {import('./state.js').Student[]}
 */
export function getMonthStudents(allStudents, monthAttendanceMap) {
  const historicalIds = new Set();
  Object.keys(monthAttendanceMap || {}).forEach((cellKey) => {
    const [studentId] = cellKey.split('|');
    if (studentId && monthAttendanceMap[cellKey]) {
      historicalIds.add(studentId);
    }
  });

  return allStudents.filter((s) => s.is_active || historicalIds.has(s.student_id));
}

/**
 * Sumber tunggal perhitungan rekap bulanan (dipakai oleh pratinjau UI dan Export Excel).
 *
 * @param {{
 *   students: import('./state.js').Student[],
 *   categories: import('./state.js').Category[],
 *   subCategories: import('./state.js').SubCategory[],
 *   activeDates: string[],
 *   attendanceMap: Record<string, string>
 * }} params
 * @returns {{ rows: RecapRow[], classSummary: ClassSummary }}
 */
export function buildMonthlyRecap({
  students,
  categories,
  subCategories,
  activeDates,
  attendanceMap,
}) {
  const subNameById = {};
  subCategories.forEach((sub) => {
    subNameById[sub.subcategory_id] = sub.name;
  });

  const sortedCategories = categories.slice().sort((a, b) => a.order_index - b.order_index);
  const map = attendanceMap || {};

  let sumH = 0;
  let sumI = 0;
  let sumS = 0;
  let sumA = 0;
  let sumTotal = 0;

  const rows = students.map((student, idx) => {
    let h = 0;
    let i = 0;
    let s = 0;
    let a = 0;

    activeDates.forEach((date) => {
      const val = map[`${student.student_id}|${date}`];
      if (val === STATUS.H) h++;
      else if (val === STATUS.I) i++;
      else if (val === STATUS.S) s++;
      else if (val === STATUS.A) a++;
    });

    const totalPertemuan = h + i + s + a;
    const presentasiKehadiran = totalPertemuan > 0 ? Math.round((h / totalPertemuan) * 100) : 0;
    const presentasiIzin = totalPertemuan > 0 ? Math.round((i / totalPertemuan) * 100) : 0;
    const presentasiSakit = totalPertemuan > 0 ? Math.round((s / totalPertemuan) * 100) : 0;
    const presentasiAlfa = totalPertemuan > 0 ? Math.round((a / totalPertemuan) * 100) : 0;

    sumH += h;
    sumI += i;
    sumS += s;
    sumA += a;
    sumTotal += totalPertemuan;

    const categoryValues = {};
    sortedCategories.forEach((cat) => {
      const subId = student.categories?.[cat.category_id];
      categoryValues[cat.category_id] = subId && subNameById[subId] ? subNameById[subId] : '-';
    });

    return {
      no: idx + 1,
      student_id: student.student_id,
      name: student.name,
      categoryValues,
      hadir: h,
      izin: i,
      sakit: s,
      alfa: a,
      totalPertemuan,
      presentasiKehadiran,
      presentasiIzin,
      presentasiSakit,
      presentasiAlfa,
      isWarningRed: presentasiAlfa > ALFA_WARNING_THRESHOLD,
    };
  });

  const classSummary = {
    totalHadir: sumH,
    totalIzin: sumI,
    totalSakit: sumS,
    totalAlfa: sumA,
    totalPertemuan: sumTotal,
    presentasiKehadiran: sumTotal > 0 ? Math.round((sumH / sumTotal) * 100) : 0,
    presentasiIzin: sumTotal > 0 ? Math.round((sumI / sumTotal) * 100) : 0,
    presentasiSakit: sumTotal > 0 ? Math.round((sumS / sumTotal) * 100) : 0,
    presentasiAlfa: sumTotal > 0 ? Math.round((sumA / sumTotal) * 100) : 0,
  };

  return { rows, classSummary };
}
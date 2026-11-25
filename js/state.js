import { STORAGE_KEYS } from './config.js';

/**
 * @typedef {Object} Category
 * @property {string} category_id
 * @property {string} name
 * @property {number} order_index
 */

/**
 * @typedef {Object} SubCategory
 * @property {string} subcategory_id
 * @property {string} category_id
 * @property {string} name
 */

/**
 * @typedef {Object} Student
 * @property {string} student_id
 * @property {string} name
 * @property {Record<string, string>} categories
 * @property {boolean} is_active
 * @property {string} created_at
 */

/**
 * @typedef {Object} MonthSchedule
 * @property {string} month_key
 * @property {string[]} dates
 * @property {string[]} added_dates
 * @property {string[]} removed_dates
 */

/**
 * @typedef {Object} PendingAttendanceChange
 * @property {string} month_key
 * @property {string} student_id
 * @property {string} date
 * @property {string} status
 */

/**
 * @typedef {Object} KasSettings
 * @property {string} title
 * @property {number} fee_amount
 * @property {string[]} categories_in
 * @property {string[]} categories_out
 * @property {Record<string, number[]>} months_by_year
 */

/**
 * @typedef {Object} KasIuranRecord
 * @property {string} iuran_key
 * @property {string} student_id
 * @property {string} month_key
 * @property {number} amount
 * @property {string} paid_at
 */

/**
 * @typedef {Object} KasTransaction
 * @property {string} tx_id
 * @property {'in' | 'out'} type
 * @property {number} amount
 * @property {string} title
 * @property {string} category
 * @property {string} date
 * @property {string} note
 * @property {string} created_at
 */

/** @returns {string} Format YYYY-MM bulan berjalan */
export function getCurrentMonthKey() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  return `${y}-${m}`;
}

export const state = {
  activeTab: localStorage.getItem(STORAGE_KEYS.LAST_TAB) || 'absensi',
  activeMonthKey: getCurrentMonthKey(),
  /** @type {string | null} */
  selectedCategoryId: null,
  /** @type {'idle' | 'saving' | 'error'} */
  syncStatus: 'idle',
  settings: {
    class_name: 'GPR Absensi',
    /** @type {number[]} */
    default_days: [],
  },
  /** @type {Category[]} */
  categories: [],
  /** @type {SubCategory[]} */
  subCategories: [],
  /** @type {Student[]} */
  students: [],
  /** @type {Record<string, MonthSchedule | null>} */
  schedulesByMonth: {},
  /** @type {Record<string, Record<string, string>>} */
  attendanceByMonth: {},
  /** @type {Record<string, boolean>} */
  loadedMonths: {},
  /** @type {Map<string, PendingAttendanceChange>} */
  pendingQueue: new Map(),

  // State Modul Kas (V2.1)
  kas: {
    loaded: false,
    /** @type {'dashboard' | 'iuran' | 'form' | 'laporan'} */
    subView: 'dashboard',
    /** @type {'in' | 'out'} */
    formType: 'in',
    /** @type {string | null} */
    editTxId: null,
    activeYear: new Date().getFullYear(),
    /** @type {KasSettings} */
    settings: {
      title: 'Iuran Kas Bulanan',
      fee_amount: 20000,
      categories_in: ['Donasi & Sukarela', 'Sisa Kegiatan', 'Lainnya'],
      categories_out: ['Konsumsi', 'Perlengkapan', 'Kegiatan', 'Lainnya'],
      months_by_year: {},
    },
    /** @type {KasIuranRecord[]} */
    iuran: [],
    /** @type {KasTransaction[]} */
    transactions: [],
  },
};

/** @param {string} tab */
export function setActiveTab(tab) {
  state.activeTab = tab;
  localStorage.setItem(STORAGE_KEYS.LAST_TAB, tab);
}

/**
 * @param {{ settings: any, categories: Category[], subCategories: SubCategory[], students: Student[] }} data
 */
export function setBootstrapData(data) {
  state.settings = {
    class_name: data.settings?.class_name || 'GPR Absensi',
    default_days: Array.isArray(data.settings?.default_days)
      ? data.settings.default_days.map(Number)
      : [],
  };
  state.categories = Array.isArray(data.categories) ? data.categories : [];
  state.subCategories = Array.isArray(data.subCategories) ? data.subCategories : [];
  state.students = Array.isArray(data.students) ? data.students : [];
}

/**
 * @param {string} monthKey
 * @param {{ schedule: MonthSchedule | null, attendance: Array<{ student_id: string, date: string, status: string }> }} data
 */
export function setMonthData(monthKey, data) {
  state.schedulesByMonth[monthKey] = data.schedule || null;

  const map = {};
  if (Array.isArray(data.attendance)) {
    data.attendance.forEach((item) => {
      if (item.student_id && item.date && item.status) {
        map[`${item.student_id}|${item.date}`] = item.status;
      }
    });
  }

  state.pendingQueue.forEach((change) => {
    if (change.month_key === monthKey) {
      const cellKey = `${change.student_id}|${change.date}`;
      if (change.status) {
        map[cellKey] = change.status;
      } else {
        delete map[cellKey];
      }
    }
  });

  state.attendanceByMonth[monthKey] = map;
  state.loadedMonths[monthKey] = true;
}

/**
 * @param {{ kasSettings: KasSettings, iuran: KasIuranRecord[], transactions: KasTransaction[] }} data
 */
export function setKasData(data) {
  if (data.kasSettings) {
    state.kas.settings = {
      title: data.kasSettings.title || 'Iuran Kas Bulanan',
      fee_amount: Number(data.kasSettings.fee_amount) || 20000,
      categories_in: Array.isArray(data.kasSettings.categories_in)
        ? data.kasSettings.categories_in
        : ['Donasi & Sukarela', 'Sisa Kegiatan', 'Lainnya'],
      categories_out: Array.isArray(data.kasSettings.categories_out)
        ? data.kasSettings.categories_out
        : ['Konsumsi', 'Perlengkapan', 'Kegiatan', 'Lainnya'],
      months_by_year:
        data.kasSettings.months_by_year && typeof data.kasSettings.months_by_year === 'object'
          ? data.kasSettings.months_by_year
          : {},
    };
  }
  state.kas.iuran = Array.isArray(data.iuran) ? data.iuran : [];
  state.kas.transactions = Array.isArray(data.transactions) ? data.transactions : [];
  state.kas.loaded = true;
}

/**
 * @param {string} categoryId
 * @returns {SubCategory[]}
 */
export function getSubCategoriesByCategory(categoryId) {
  return state.subCategories.filter((sub) => sub.category_id === categoryId);
}

/**
 * @param {string | undefined} subCategoryId
 * @returns {string}
 */
export function getSubCategoryName(subCategoryId) {
  if (!subCategoryId) return '-';
  const found = state.subCategories.find((sub) => sub.subcategory_id === subCategoryId);
  return found ? found.name : '-';
}

/**
 * @param {string} monthKey
 * @param {string} studentId
 * @param {string} date
 * @returns {string}
 */
export function getAttendanceValue(monthKey, studentId, date) {
  const monthMap = state.attendanceByMonth[monthKey];
  if (!monthMap) return '';
  return monthMap[`${studentId}|${date}`] || '';
}

/**
 * @param {string} monthKey
 * @param {string} studentId
 * @param {string} date
 * @param {string} status
 */
export function setAttendanceValue(monthKey, studentId, date, status) {
  if (!state.attendanceByMonth[monthKey]) {
    state.attendanceByMonth[monthKey] = {};
  }
  const cellKey = `${studentId}|${date}`;
  if (status) {
    state.attendanceByMonth[monthKey][cellKey] = status;
  } else {
    delete state.attendanceByMonth[monthKey][cellKey];
  }
}
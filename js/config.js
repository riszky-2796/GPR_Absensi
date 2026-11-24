/**
 * Konfigurasi pusat dan konstanta aplikasi GPR Absensi.
 */

export const API_BASE_URL = 'https://script.google.com/macros/s/AKfycbwEXi8rLaGLugR-q5TjG2_Vc0PZ9NdwEXKVKnJKkGYdBzVWnR47oWBwV2BFpGyFUnCBhg/exec';

export const DEBOUNCE_DELAY_MS = 1500;
export const ALFA_WARNING_THRESHOLD = 50;

export const STATUS = Object.freeze({
  H: 'H',
  I: 'I',
  S: 'S',
  A: 'A',
});

export const STATUS_LIST = Object.freeze([STATUS.H, STATUS.I, STATUS.S, STATUS.A]);

export const STORAGE_KEYS = Object.freeze({
  LAST_TAB: 'gpr_absensi_last_tab',
});

export const MONTH_NAMES = Object.freeze([
  'JANUARI',
  'FEBRUARI',
  'MARET',
  'APRIL',
  'MEI',
  'JUNI',
  'JULI',
  'AGUSTUS',
  'SEPTEMBER',
  'OKTOBER',
  'NOVEMBER',
  'DESEMBER',
]);

export const DAY_NAMES = Object.freeze([
  'Minggu',
  'Senin',
  'Selasa',
  'Rabu',
  'Kamis',
  'Jumat',
  'Sabtu',
]);

export const DAY_SHORT_NAMES = Object.freeze([
  'Min',
  'Sen',
  'Sel',
  'Rab',
  'Kam',
  'Jum',
  'Sab',
]);

export const EXCEL_COLORS = Object.freeze({
  HEADER_BG: 'FFBDD7EE',
  ZEBRA_ODD_BG: 'FFBDD7EE',
  ZEBRA_EVEN_BG: 'FFFFFFFF',
  WARNING_RED_BG: 'FFFF0000',
  BORDER_BLACK: 'FF000000',
  TEXT_BLACK: 'FF000000',
});
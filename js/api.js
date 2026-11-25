import { API_BASE_URL } from './config.js';

/**
 * Mengirim request POST ke Google Apps Script Web App (tanpa PIN).
 * Menggunakan text/plain agar bebas dari blokir CORS preflight browser.
 *
 * @param {string} action
 * @param {Record<string, any>} [payload={}]
 * @returns {Promise<any>}
 */
async function postGas(action, payload = {}) {
  if (!API_BASE_URL || !API_BASE_URL.startsWith('https://script.google.com/macros/s/')) {
    throw new Error(
      'URL Google Apps Script belum dipasang di js/config.js. Silakan tempel URL /exec Anda.'
    );
  }

  let response;
  try {
    response = await fetch(API_BASE_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action, payload }),
    });
  } catch (err) {
    throw new Error('Gagal terhubung ke server. Periksa koneksi internet atau URL Web App.');
  }

  if (!response.ok) {
    throw new Error(`Server merespon error (HTTP ${response.status}).`);
  }

  let result;
  try {
    result = await response.json();
  } catch (err) {
    throw new Error('Respon server tidak valid. Pastikan akses Web App GAS diatur ke "Anyone".');
  }

  if (!result.ok) {
    throw new Error(result.error || 'Terjadi kesalahan saat memproses data.');
  }

  return result.data;
}

export const api = {
  // Absensi (V1)
  getBootstrap: () => postGas('getBootstrap'),
  getMonthData: (month_key) => postGas('getMonthData', { month_key }),
  saveSettings: (payload) => postGas('saveSettings', payload),
  saveCategory: (payload) => postGas('saveCategory', payload),
  deleteCategory: (category_id, force = false) =>
    postGas('deleteCategory', { category_id, force }),
  saveSubCategory: (payload) => postGas('saveSubCategory', payload),
  deleteSubCategory: (subcategory_id, force = false) =>
    postGas('deleteSubCategory', { subcategory_id, force }),
  saveStudent: (payload) => postGas('saveStudent', payload),
  deleteStudent: (student_id) => postGas('deleteStudent', { student_id }),
  saveSchedule: (payload) => postGas('saveSchedule', payload),
  batchSaveAttendance: (month_key, changes) =>
    postGas('batchSaveAttendance', { month_key, changes }),

  // Kas (V2)
  getKasData: () => postGas('getKasData'),
  saveKasSettings: (payload) => postGas('saveKasSettings', payload),
  toggleKasIuran: (payload) => postGas('toggleKasIuran', payload),
  saveKasTransaction: (payload) => postGas('saveKasTransaction', payload),
  deleteKasTransaction: (tx_id) => postGas('deleteKasTransaction', { tx_id }),
};
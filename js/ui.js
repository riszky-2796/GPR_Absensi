import { DAY_NAMES, STATUS_LIST } from './config.js';
import {
  state,
  getSubCategoriesByCategory,
  getSubCategoryName,
  getAttendanceValue,
} from './state.js';
import {
  formatDateHeader,
  formatMonthLabel,
  getMonthStudents,
  parseMonthKey,
  resolveMonthDates,
  buildMonthlyRecap,
} from './calculator.js';

const overlayDom = {
  popoverRoot: document.getElementById('popover-root'),
  modalRoot: document.getElementById('modal-root'),
  loadingOverlay: document.getElementById('loading-overlay'),
  loadingText: document.getElementById('loading-text'),
  toastContainer: document.getElementById('toast-container'),
};

/** @type {HTMLElement | null} */
let activeCellElement = null;

/**
 * Sanitasi teks sebelum masuk ke template HTML untuk mencegah XSS.
 * @param {any} value
 * @returns {string}
 */
export function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Menampilkan notifikasi toast singkat.
 * @param {string} message
 * @param {'info' | 'success' | 'error'} [type='info']
 */
export function showToast(message, type = 'info') {
  const el = document.createElement('div');
  el.className = `toast ${
    type === 'error' ? 'toast-error' : type === 'success' ? 'toast-success' : ''
  }`;
  el.textContent = message;
  overlayDom.toastContainer.appendChild(el);
  setTimeout(() => el.remove(), 3200);
}

/**
 * Menampilkan atau menyembunyikan overlay loading.
 * @param {boolean} visible
 * @param {string} [message='Memuat data...']
 */
export function setLoading(visible, message = 'Memuat data...') {
  overlayDom.loadingText.textContent = message;
  overlayDom.loadingOverlay.classList.toggle('hidden', !visible);
}

/* ==================== MODAL REUSABLE ==================== */

export function closeModal() {
  overlayDom.modalRoot.innerHTML = '';
}

/**
 * Membuka modal dialog reusable.
 * @param {{ title: string, bodyHtml: string, footerHtml?: string, onMount?: (card: HTMLElement) => void }} options
 */
export function openModal({ title, bodyHtml, footerHtml = '', onMount }) {
  closePopover();
  overlayDom.modalRoot.innerHTML = `
    <div class="modal-backdrop" id="active-modal-backdrop">
      <div class="modal-card" role="dialog" aria-modal="true" aria-label="${escapeHtml(title)}">
        <div class="modal-header">
          <h3 class="modal-title">${escapeHtml(title)}</h3>
          <button type="button" class="icon-btn" data-close-modal aria-label="Tutup">✕</button>
        </div>
        <div class="modal-body">${bodyHtml}</div>
        ${footerHtml ? `<div class="modal-footer">${footerHtml}</div>` : ''}
      </div>
    </div>
  `;

  const backdrop = document.getElementById('active-modal-backdrop');
  const card = backdrop.querySelector('.modal-card');

  backdrop.addEventListener('click', (e) => {
    if (e.target === backdrop || e.target.closest('[data-close-modal]')) {
      closeModal();
    }
  });

  if (typeof onMount === 'function') {
    onMount(card);
  }
}

/**
 * Modal konfirmasi aksi berbahaya / proteksi penghapusan.
 * @param {{ title: string, message: string, confirmText?: string, isDanger?: boolean, onConfirm: () => Promise<void> | void }} options
 */
export function openConfirmModal({
  title,
  message,
  confirmText = 'Lanjutkan',
  isDanger = true,
  onConfirm,
}) {
  openModal({
    title,
    bodyHtml: `<div class="warning-box">${escapeHtml(message)}</div>`,
    footerHtml: `
      <button type="button" class="btn btn-secondary" data-close-modal>Batal</button>
      <button type="button" id="modal-confirm-btn" class="btn ${
        isDanger ? 'btn-danger' : 'btn-primary'
      }">${escapeHtml(confirmText)}</button>
    `,
    onMount: (card) => {
      const btn = card.querySelector('#modal-confirm-btn');
      btn.addEventListener('click', async () => {
        btn.disabled = true;
        btn.textContent = 'Memproses...';
        try {
          await onConfirm();
          closeModal();
        } catch (err) {
          btn.disabled = false;
          btn.textContent = confirmText;
          showToast(err.message || 'Gagal memproses aksi.', 'error');
        }
      });
    },
  });
}

/* ==================== POPOVER STATUS ABSENSI ==================== */

export function closePopover() {
  if (activeCellElement) {
    activeCellElement.classList.remove('is-active');
    activeCellElement = null;
  }
  overlayDom.popoverRoot.innerHTML = '';
}

/**
 * Membuka popover pilihan H / I / S / A / Kosongkan tepat di dekat sel yang diketuk.
 * @param {HTMLElement} cellEl
 * @param {(status: string) => void} onSelect
 */
export function openStatusPopover(cellEl, onSelect) {
  const isSameCell = activeCellElement === cellEl;
  closePopover();
  if (isSameCell) return;

  activeCellElement = cellEl;
  cellEl.classList.add('is-active');

  const menu = document.createElement('div');
  menu.className = 'popover-menu';
  menu.setAttribute('role', 'menu');
  menu.innerHTML = `
    <div class="popover-row">
      ${STATUS_LIST.map(
        (st) =>
          `<button type="button" class="popover-btn popover-btn-${st.toLowerCase()}" data-pick-status="${st}">${st}</button>`
      ).join('')}
    </div>
    <button type="button" class="popover-btn-clear" data-pick-status="">Kosongkan</button>
  `;

  overlayDom.popoverRoot.appendChild(menu);

  // Hitung posisi agar tidak terpotong tepi layar HP/Desktop
  const cellRect = cellEl.getBoundingClientRect();
  const menuRect = menu.getBoundingClientRect();
  const margin = 8;

  let left = cellRect.left + cellRect.width / 2 - menuRect.width / 2;
  left = Math.max(margin, Math.min(left, window.innerWidth - menuRect.width - margin));

  let top = cellRect.bottom + 6;
  if (top + menuRect.height + margin > window.innerHeight) {
    top = Math.max(margin, cellRect.top - menuRect.height - 6);
  }

  menu.style.left = `${Math.round(left)}px`;
  menu.style.top = `${Math.round(top)}px`;

  menu.addEventListener('click', (e) => {
    const btn = e.target.closest('[data-pick-status]');
    if (!btn) return;
    const picked = btn.dataset.pickStatus;
    closePopover();
    onSelect(picked);
  });
}

/**
 * Memperbarui badge pada 1 sel absensi secara instan tanpa mereset scroll tabel.
 * @param {HTMLElement} cellEl
 * @param {string} status
 */
export function updateAttendanceCellDom(cellEl, status) {
  cellEl.innerHTML = renderStatusBadgeHtml(status);
}

/**
 * @param {string} status
 * @returns {string}
 */
function renderStatusBadgeHtml(status) {
  if (!status) {
    return `<span class="badge-status badge-status-empty">·</span>`;
  }
  const lower = status.toLowerCase();
  return `<span class="badge-status badge-status-${escapeHtml(lower)}">${escapeHtml(status)}</span>`;
}

/* ==================== RENDER TAB 1: ABSENSI ==================== */

/**
 * @param {HTMLElement} container
 */
export function renderAbsensiView(container) {
  const monthKey = state.activeMonthKey;
  const monthLabel = formatMonthLabel(monthKey);
  const monthAttendance = state.attendanceByMonth[monthKey] || {};
  const resolved = resolveMonthDates(
    monthKey,
    state.settings.default_days,
    state.schedulesByMonth[monthKey] || null
  );
  const students = getMonthStudents(state.students, monthAttendance);

  let bodyHtml = '';

  if (students.length === 0) {
    bodyHtml = `
      <div class="empty-state">
        <strong>Belum ada data siswa</strong>
        Tambahkan siswa terlebih dahulu di menu <b>Siswa</b> untuk mulai mengabsen.
        <div style="margin-top: 12px;">
          <button type="button" class="btn btn-primary btn-sm" data-goto-tab="siswa">Buka Menu Siswa</button>
        </div>
      </div>
    `;
  } else if (resolved.dates.length === 0) {
    bodyHtml = `
      <div class="empty-state">
        <strong>Belum ada jadwal pertemuan di bulan ${escapeHtml(monthLabel)}</strong>
        Atur pola hari rutin atau tambah tanggal pertemuan manual pada menu <b>Pengaturan &amp; Rekap</b>.
        <div style="margin-top: 12px;">
          <button type="button" class="btn btn-primary btn-sm" data-goto-tab="pengaturan">Atur Jadwal Sekarang</button>
        </div>
      </div>
    `;
  } else {
    const dateHeaders = resolved.dates
      .map((isoDate) => {
        const { dayShort, dayNum } = formatDateHeader(isoDate);
        return `
          <th class="col-date" title="${escapeHtml(isoDate)}">
            <span class="date-day">${escapeHtml(dayShort)}</span>
            <span class="date-num">${escapeHtml(dayNum)}</span>
          </th>
        `;
      })
      .join('');

    const studentRows = students
      .map((student, index) => {
        const cells = resolved.dates
          .map((isoDate) => {
            const val = getAttendanceValue(monthKey, student.student_id, isoDate);
            return `
              <td
                class="absen-cell"
                data-student-id="${escapeHtml(student.student_id)}"
                data-date="${escapeHtml(isoDate)}"
              >
                ${renderStatusBadgeHtml(val)}
              </td>
            `;
          })
          .join('');

        return `
          <tr>
            <td class="sticky-col-no">${index + 1}</td>
            <td class="sticky-col-name" title="${escapeHtml(student.name)}">${escapeHtml(student.name)}</td>
            ${cells}
          </tr>
        `;
      })
      .join('');

    bodyHtml = `
      <div class="table-wrap" id="absensi-table-wrap">
        <table class="data-table">
          <thead>
            <tr>
              <th class="sticky-col-no">NO</th>
              <th class="sticky-col-name">NAMA</th>
              ${dateHeaders}
            </tr>
          </thead>
          <tbody>
            ${studentRows}
          </tbody>
        </table>
      </div>
      <p class="helper-text">
        Ketuk kotak tanggal untuk memilih <b>H</b> (Hadir), <b>I</b> (Izin), <b>S</b> (Sakit), atau <b>A</b> (Alfa). Perubahan tersimpan otomatis ke Google Sheets.
      </p>
    `;
  }

  container.innerHTML = `
    <div class="card">
      <div class="toolbar">
        <div class="month-picker">
          <button type="button" class="icon-btn" data-month-shift="-1" aria-label="Bulan Sebelumnya">‹</button>
          <span class="month-label">${escapeHtml(monthLabel)}</span>
          <button type="button" class="icon-btn" data-month-shift="1" aria-label="Bulan Berikutnya">›</button>
        </div>
        <div class="row-wrap">
          <button type="button" class="btn btn-secondary btn-sm" data-goto-tab="pengaturan">Atur Jadwal</button>
          <button type="button" class="btn btn-primary btn-sm" data-export-excel>Rekap &amp; Download</button>
        </div>
      </div>
    </div>

    <div class="card">
      <div class="card-header">
        <div>
          <h2 class="card-title">Absensi Bulan ${escapeHtml(monthLabel)}</h2>
          <p class="card-desc">${students.length} siswa · ${resolved.dates.length} pertemuan terjadwal</p>
        </div>
      </div>
      ${bodyHtml}
    </div>
  `;
}

/* ==================== RENDER TAB 2: SISWA ==================== */

/**
 * @param {HTMLElement} container
 */
export function renderSiswaView(container) {
  const activeStudents = state.students.filter((s) => s.is_active);
  const sortedCategories = state.categories.slice().sort((a, b) => a.order_index - b.order_index);

  let contentHtml = '';
  if (activeStudents.length === 0) {
    contentHtml = `
      <div class="empty-state">
        <strong>Daftar siswa masih kosong</strong>
        Klik tombol <b>+ Tambah Siswa</b> untuk mulai memasukkan nama siswa beserta kategorinya.
      </div>
    `;
  } else {
    const catHeaders = sortedCategories.map((c) => `<th>${escapeHtml(c.name)}</th>`).join('');
    const rowsHtml = activeStudents
      .map((student, idx) => {
        const catCells = sortedCategories
          .map((c) => {
            const subName = getSubCategoryName(student.categories?.[c.category_id]);
            return `<td><span class="cat-pill">${escapeHtml(subName)}</span></td>`;
          })
          .join('');

        return `
          <tr>
            <td>${idx + 1}</td>
            <td class="text-left"><strong>${escapeHtml(student.name)}</strong></td>
            ${catCells}
            <td>
              <div class="action-group">
                <button type="button" class="btn btn-secondary btn-sm" data-edit-student="${escapeHtml(student.student_id)}">Edit</button>
                <button type="button" class="btn btn-outline-danger btn-sm" data-delete-student="${escapeHtml(student.student_id)}">Hapus</button>
              </div>
            </td>
          </tr>
        `;
      })
      .join('');

    contentHtml = `
      <div class="table-wrap">
        <table class="data-table">
          <thead>
            <tr>
              <th style="width: 48px;">NO</th>
              <th class="text-left">NAMA SISWA</th>
              ${catHeaders}
              <th style="width: 140px;">AKSI</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHtml}
          </tbody>
        </table>
      </div>
    `;
  }

  container.innerHTML = `
    <div class="card">
      <div class="card-header">
        <div>
          <h2 class="card-title">Daftar Siswa</h2>
          <p class="card-desc">Total ${activeStudents.length} siswa aktif terdaftar.</p>
        </div>
        <button type="button" class="btn btn-primary btn-sm" id="btn-add-student">+ Tambah Siswa</button>
      </div>
      ${
        sortedCategories.length === 0
          ? `<p class="helper-text" style="margin-bottom: 12px;">Tips: Buat kategori dulu di menu <b>Kategori</b> (mis. Jenis Kelamin / Kelompok) agar bisa langsung dipilih saat menambah siswa.</p>`
          : ''
      }
      ${contentHtml}
    </div>
  `;
}

/* ==================== RENDER TAB 3: KATEGORI 2 LAPIS ==================== */

/**
 * @param {HTMLElement} container
 */
export function renderKategoriView(container) {
  const sortedCategories = state.categories.slice().sort((a, b) => a.order_index - b.order_index);

  // Lapis 2: Masuk ke dalam detail satu Kategori Induk untuk kelola Sub-Kategori
  if (state.selectedCategoryId) {
    const parentCat = sortedCategories.find((c) => c.category_id === state.selectedCategoryId);
    if (!parentCat) {
      state.selectedCategoryId = null;
    } else {
      const subs = getSubCategoriesByCategory(parentCat.category_id);
      const subItemsHtml =
        subs.length === 0
          ? `
            <div class="empty-state">
              <strong>Belum ada sub-kategori</strong>
              Tambahkan pilihan sub-kategori untuk <b>${escapeHtml(parentCat.name)}</b> (mis. L, P atau K 2, K 3).
            </div>
          `
          : `
            <div class="table-wrap">
              <table class="data-table">
                <thead>
                  <tr>
                    <th style="width: 48px;">NO</th>
                    <th class="text-left">NAMA SUB-KATEGORI</th>
                    <th style="width: 140px;">AKSI</th>
                  </tr>
                </thead>
                <tbody>
                  ${subs
                    .map(
                      (sub, idx) => `
                        <tr>
                          <td>${idx + 1}</td>
                          <td class="text-left"><strong>${escapeHtml(sub.name)}</strong></td>
                          <td>
                            <div class="action-group">
                              <button type="button" class="btn btn-secondary btn-sm" data-edit-sub="${escapeHtml(sub.subcategory_id)}">Edit</button>
                              <button type="button" class="btn btn-outline-danger btn-sm" data-delete-sub="${escapeHtml(sub.subcategory_id)}">Hapus</button>
                            </div>
                          </td>
                        </tr>
                      `
                    )
                    .join('')}
                </tbody>
              </table>
            </div>
          `;

      container.innerHTML = `
        <div class="card">
          <div class="breadcrumb-bar">
            <button type="button" class="btn btn-secondary btn-sm" id="btn-back-categories">← Semua Kategori</button>
            <div>
              <h2 class="card-title">${escapeHtml(parentCat.name)}</h2>
              <p class="card-desc">Kelola pilihan sub-kategori (Lapis 2).</p>
            </div>
          </div>

          <form id="form-add-subcategory" class="inline-form" style="margin-bottom: 16px;">
            <input
              type="text"
              id="input-sub-name"
              class="input-control"
              placeholder="Nama sub-kategori baru (mis. L / P / K 2)..."
              required
            />
            <button type="submit" class="btn btn-primary">+ Tambah</button>
          </form>

          ${subItemsHtml}
        </div>
      `;
      return;
    }
  }

  // Lapis 1: Daftar Kategori Induk
  const cardsHtml =
    sortedCategories.length === 0
      ? `
        <div class="empty-state">
          <strong>Belum ada kategori induk</strong>
          Buat kategori induk terlebih dahulu (contoh: <b>Jenis Kelamin</b> atau <b>KELOMPOK</b>), lalu klik untuk mengisi sub-kategorinya.
        </div>
      `
      : `
        <div class="category-list">
          ${sortedCategories
            .map((cat) => {
              const subs = getSubCategoriesByCategory(cat.category_id);
              const chips =
                subs.length > 0
                  ? subs
                      .map((s) => `<span class="sub-chip">${escapeHtml(s.name)}</span>`)
                      .join('')
                  : `<span class="helper-text">Belum ada sub-kategori</span>`;

              return `
                <div class="category-card">
                  <div class="category-card-head">
                    <div>
                      <h3 class="category-name">${escapeHtml(cat.name)}</h3>
                      <div class="category-meta">${subs.length} sub-kategori</div>
                    </div>
                    <button type="button" class="btn btn-primary btn-sm" data-open-category="${escapeHtml(cat.category_id)}">
                      Kelola Sub →
                    </button>
                  </div>
                  <div class="sub-chip-list">${chips}</div>
                  <div class="row-wrap" style="margin-top: auto; padding-top: 6px; border-top: 1px solid var(--border);">
                    <button type="button" class="btn btn-secondary btn-sm" data-edit-category="${escapeHtml(cat.category_id)}">Ubah Nama</button>
                    <button type="button" class="btn btn-outline-danger btn-sm" data-delete-category="${escapeHtml(cat.category_id)}">Hapus</button>
                  </div>
                </div>
              `;
            })
            .join('')}
        </div>
      `;

  container.innerHTML = `
    <div class="card">
      <div class="card-header">
        <div>
          <h2 class="card-title">Kategori Siswa (2 Lapis)</h2>
          <p class="card-desc">Buat kategori induk, lalu tekan <b>Kelola Sub</b> untuk mengisi pilihan di dalamnya.</p>
        </div>
        <button type="button" class="btn btn-primary btn-sm" id="btn-add-category">+ Kategori Baru</button>
      </div>
      ${cardsHtml}
    </div>
  `;
}

/* ==================== RENDER TAB 4: PENGATURAN & REKAP ==================== */

/**
 * Menghasilkan HTML tabel pratinjau rekap bulanan (sumber angka sama persis dengan ExcelJS).
 * @param {string} monthKey
 * @returns {string}
 */
function renderRecapPreviewHtml(monthKey) {
  const monthAttendance = state.attendanceByMonth[monthKey] || {};
  const resolved = resolveMonthDates(
    monthKey,
    state.settings.default_days,
    state.schedulesByMonth[monthKey] || null
  );
  const students = getMonthStudents(state.students, monthAttendance);

  if (students.length === 0) {
    return `
      <div class="empty-state">
        <strong>Belum ada data untuk direkap</strong>
        Tambahkan siswa dan isi absensi terlebih dahulu.
      </div>
    `;
  }

  const sortedCategories = state.categories.slice().sort((a, b) => a.order_index - b.order_index);
  const { rows, classSummary } = buildMonthlyRecap({
    students,
    categories: sortedCategories,
    subCategories: state.subCategories,
    activeDates: resolved.dates,
    attendanceMap: monthAttendance,
  });

  const totalCols = 11 + sortedCategories.length;
  const mergeClassCols = totalCols - 4;
  const titleText = `REKAP ABSENSI ${String(state.settings.class_name || 'GPR ABSENSI').toUpperCase()} BULAN ${formatMonthLabel(monthKey)}`;

  const bodyRows = rows
    .map((r, idx) => {
      const zebraClass = idx % 2 === 0 ? 'recap-zebra-odd' : 'recap-zebra-even';
      const warnClass = r.isWarningRed ? 'recap-warn-red' : '';
      const catCells = sortedCategories
        .map((c) => `<td>${escapeHtml(r.categoryValues[c.category_id] || '-')}</td>`)
        .join('');

      return `
        <tr class="${zebraClass}">
          <td>${r.no}</td>
          <td class="text-left ${warnClass}">${escapeHtml(r.name)}</td>
          ${catCells}
          <td>${r.hadir}</td>
          <td>${r.izin}</td>
          <td>${r.sakit}</td>
          <td>${r.alfa}</td>
          <td>${r.totalPertemuan}</td>
          <td>${r.presentasiKehadiran}%</td>
          <td>${r.presentasiIzin}%</td>
          <td>${r.presentasiSakit}%</td>
          <td class="${warnClass}">${r.presentasiAlfa}%</td>
        </tr>
      `;
    })
    .join('');

  return `
    <div class="table-wrap">
      <table class="recap-table">
        <thead>
          <tr>
            <th colspan="${totalCols}" class="recap-title-cell">${escapeHtml(titleText)}</th>
          </tr>
          <tr>
            <th>NO</th>
            <th>NAMA</th>
            ${sortedCategories.map((c) => `<th>${escapeHtml(c.name)}</th>`).join('')}
            <th>HADIR</th>
            <th>IZIN</th>
            <th>SAKIT</th>
            <th>ALFA</th>
            <th>TOTAL PERTEMUAN</th>
            <th>PRESENTASI KEHADIRAN</th>
            <th>PRESENTASI IZIN</th>
            <th>PRESENTASI SAKIT</th>
            <th>PRESENTASI ALFA</th>
          </tr>
        </thead>
        <tbody>
          ${bodyRows}
          <tr class="recap-class-row">
            <td colspan="${mergeClassCols}">PRESENTASI KELAS</td>
            <td>${classSummary.presentasiKehadiran}%</td>
            <td>${classSummary.presentasiIzin}%</td>
            <td>${classSummary.presentasiSakit}%</td>
            <td>${classSummary.presentasiAlfa}%</td>
          </tr>
        </tbody>
      </table>
    </div>
  `;
}

/**
 * @param {HTMLElement} container
 */
export function renderPengaturanView(container) {
  const monthKey = state.activeMonthKey;
  const monthLabel = formatMonthLabel(monthKey);
  const { year, monthIndex } = parseMonthKey(monthKey);
  const mm = String(monthIndex + 1).padStart(2, '0');
  const lastDay = new Date(year, monthIndex + 1, 0).getDate();
  const minDate = `${year}-${mm}-01`;
  const maxDate = `${year}-${mm}-${String(lastDay).padStart(2, '0')}`;

  const defaultDaySet = new Set(state.settings.default_days.map(Number));
  const resolved = resolveMonthDates(
    monthKey,
    state.settings.default_days,
    state.schedulesByMonth[monthKey] || null
  );
  const addedSet = new Set(resolved.added_dates);

  const dayButtonsHtml = DAY_NAMES.map(
    (dayName, idx) => `
      <button
        type="button"
        class="day-toggle-btn ${defaultDaySet.has(idx) ? 'selected' : ''}"
        data-toggle-day="${idx}"
      >
        ${escapeHtml(dayName)}
      </button>
    `
  ).join('');

  const dateChipsHtml =
    resolved.dates.length === 0
      ? `<p class="helper-text">Belum ada tanggal pertemuan pada bulan ${escapeHtml(monthLabel)}. Pilih pola hari rutin di atas atau tambah tanggal manual di bawah.</p>`
      : resolved.dates
          .map((isoDate) => {
            const { fullLabel } = formatDateHeader(isoDate);
            const isAdded = addedSet.has(isoDate);
            return `
              <span class="schedule-chip ${isAdded ? 'is-added' : ''}" title="${escapeHtml(isoDate)}">
                <span>${escapeHtml(fullLabel)}</span>
                <button
                  type="button"
                  class="schedule-chip-remove"
                  data-remove-date="${escapeHtml(isoDate)}"
                  aria-label="Hapus tanggal ${escapeHtml(isoDate)}"
                >✕</button>
              </span>
            `;
          })
          .join('');

  container.innerHTML = `
    <!-- Kartu 1: Identitas Kelas & Pola Hari Rutin -->
    <div class="card">
      <div class="card-header">
        <div>
          <h2 class="card-title">Identitas Kelas &amp; Pola Hari Rutin</h2>
          <p class="card-desc">Nama kelas tampil pada judul rekap Excel. Pola hari rutin otomatis mengisi tanggal tiap bulan baru.</p>
        </div>
      </div>

      <form id="form-settings">
        <div class="form-group">
          <label class="form-label" for="input-class-name">Nama Kelas / Grup</label>
          <input
            type="text"
            id="input-class-name"
            class="input-control"
            value="${escapeHtml(state.settings.class_name)}"
            placeholder="Contoh: GPR C"
            required
          />
        </div>

        <div class="form-group">
          <label class="form-label">Pola Hari Pertemuan Rutin</label>
          <div class="day-toggle-grid" id="day-toggle-grid">
            ${dayButtonsHtml}
          </div>
          <p class="helper-text">Mengubah pola hari rutin tidak akan menimpa jadwal bulan yang sudah memiliki catatan absensi.</p>
        </div>

        <button type="submit" id="btn-save-settings" class="btn btn-primary">Simpan Pengaturan</button>
      </form>
    </div>

    <!-- Kartu 2: Jadwal Hybrid Bulanan (Override Tanggal Spesifik) -->
    <div class="card">
      <div class="card-header">
        <div>
          <h2 class="card-title">Jadwal Pertemuan Bulan ${escapeHtml(monthLabel)}</h2>
          <p class="card-desc">Hapus tanggal jika libur atau tambahkan tanggal khusus untuk pertemuan tambahan.</p>
        </div>
        <div class="month-picker">
          <button type="button" class="icon-btn" data-month-shift="-1" aria-label="Bulan Sebelumnya">‹</button>
          <span class="month-label">${escapeHtml(monthLabel)}</span>
          <button type="button" class="icon-btn" data-month-shift="1" aria-label="Bulan Berikutnya">›</button>
        </div>
      </div>

      <div class="schedule-chip-list" style="margin-bottom: 14px;">
        ${dateChipsHtml}
      </div>

      <form id="form-add-manual-date" class="inline-form">
        <input
          type="date"
          id="input-manual-date"
          class="input-control"
          min="${minDate}"
          max="${maxDate}"
          required
        />
        <button type="submit" class="btn btn-secondary">+ Tambah Tanggal</button>
      </form>

      <div class="row-wrap" style="margin-top: 12px;">
        <button type="button" class="btn btn-ghost btn-sm" id="btn-reset-schedule">
          Reset Bulan Ini ke Pola Hari Rutin
        </button>
      </div>
    </div>

    <!-- Kartu 3: Pratinjau Rekap & Download Excel -->
    <div class="card">
      <div class="card-header">
        <div>
          <h2 class="card-title">Pratinjau Rekap ${escapeHtml(monthLabel)}</h2>
          <p class="card-desc">Sel <b>NAMA</b> dan <b>PRESENTASI ALFA</b> otomatis ditandai merah jika <b>PRESENTASI ALFA &gt; 50%</b>.</p>
        </div>
        <button type="button" class="btn btn-primary" data-export-excel>Rekap &amp; Download (.xlsx)</button>
      </div>
      ${renderRecapPreviewHtml(monthKey)}
    </div>
  `;
}
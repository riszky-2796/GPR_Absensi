import { DEBOUNCE_DELAY_MS } from './config.js';
import { api } from './api.js';
import {
  state,
  setActiveTab,
  setBootstrapData,
  setMonthData,
  setKasData,
  setAttendanceValue,
  getSubCategoriesByCategory,
} from './state.js';
import {
  buildMonthlyRecap,
  generateRoutineDates,
  getMonthStudents,
  resolveMonthDates,
  shiftMonthKey,
} from './calculator.js';
import { exportMonthlyRecapToExcel } from './exportExcel.js';
import {
  escapeHtml,
  showToast,
  setLoading,
  openModal,
  closeModal,
  openConfirmModal,
  openStatusPopover,
  closePopover,
  updateAttendanceCellDom,
  renderAbsensiView,
  renderSiswaView,
  renderKategoriView,
  renderPengaturanView,
} from './ui.js';
import { renderKasView, bindKasEvents } from './kas.js';
import { initPwaInstall } from './pwa.js';

const dom = {
  headerClassName: document.getElementById('header-class-name'),
  syncPill: document.getElementById('sync-pill'),
  viewAbsensi: document.getElementById('view-absensi'),
  viewKas: document.getElementById('view-kas'),
  viewSiswa: document.getElementById('view-siswa'),
  viewKategori: document.getElementById('view-kategori'),
  viewPengaturan: document.getElementById('view-pengaturan'),
};

/** @type {number | null} */
let debounceTimer = null;
let isFlushingQueue = false;

/* ==================== SYNC STATUS & BATCH QUEUE ==================== */

/** @param {'idle' | 'saving' | 'error'} status */
function updateSyncPill(status) {
  state.syncStatus = status;
  dom.syncPill.className = 'sync-pill';
  if (status === 'saving') {
    dom.syncPill.classList.add('sync-saving');
    dom.syncPill.textContent = 'Menyimpan...';
  } else if (status === 'error') {
    dom.syncPill.classList.add('sync-error');
    dom.syncPill.textContent = 'Gagal - Coba Lagi';
  } else {
    dom.syncPill.classList.add('sync-idle');
    dom.syncPill.textContent = 'Tersimpan ✓';
  }
}

/** @param {string} monthKey */
async function ensureMonthScheduleLocked(monthKey) {
  if (state.schedulesByMonth[monthKey]) return;
  const resolved = resolveMonthDates(monthKey, state.settings.default_days, null);
  const saved = await api.saveSchedule({
    month_key: monthKey,
    dates: resolved.dates,
    added_dates: resolved.added_dates,
    removed_dates: resolved.removed_dates,
  });
  state.schedulesByMonth[monthKey] = saved;
}

async function flushPendingQueue() {
  if (isFlushingQueue || state.pendingQueue.size === 0) return;
  isFlushingQueue = true;
  updateSyncPill('saving');

  const snapshotEntries = Array.from(state.pendingQueue.entries());
  /** @type {Record<string, Array<{ student_id: string, date: string, status: string }>>} */
  const groupedByMonth = {};

  snapshotEntries.forEach(([, item]) => {
    if (!groupedByMonth[item.month_key]) {
      groupedByMonth[item.month_key] = [];
    }
    groupedByMonth[item.month_key].push({
      student_id: item.student_id,
      date: item.date,
      status: item.status,
    });
  });

  try {
    for (const monthKey of Object.keys(groupedByMonth)) {
      await ensureMonthScheduleLocked(monthKey);
      await api.batchSaveAttendance(monthKey, groupedByMonth[monthKey]);

      snapshotEntries.forEach(([key, sentItem]) => {
        if (sentItem.month_key !== monthKey) return;
        const currentItem = state.pendingQueue.get(key);
        if (currentItem && currentItem.status === sentItem.status) {
          state.pendingQueue.delete(key);
        }
      });
    }

    if (state.pendingQueue.size > 0) {
      isFlushingQueue = false;
      scheduleFlushQueue();
      return;
    }

    updateSyncPill('idle');
  } catch (err) {
    updateSyncPill('error');
    showToast('Gagal menyimpan absensi. Klik "Gagal - Coba Lagi" di atas.', 'error');
  } finally {
    isFlushingQueue = false;
  }
}

function scheduleFlushQueue() {
  updateSyncPill('saving');
  if (debounceTimer) clearTimeout(debounceTimer);
  debounceTimer = setTimeout(() => {
    debounceTimer = null;
    flushPendingQueue();
  }, DEBOUNCE_DELAY_MS);
}

/**
 * @param {HTMLElement} cellEl
 * @param {string} studentId
 * @param {string} date
 * @param {string} status
 */
function handlePickAttendanceStatus(cellEl, studentId, date, status) {
  const monthKey = state.activeMonthKey;
  setAttendanceValue(monthKey, studentId, date, status);
  updateAttendanceCellDom(cellEl, status);

  const queueKey = `${monthKey}|${studentId}|${date}`;
  state.pendingQueue.set(queueKey, {
    month_key: monthKey,
    student_id: studentId,
    date,
    status,
  });

  scheduleFlushQueue();
}

/* ==================== NAVIGASI, HEADER & RENDER ==================== */

function syncHeaderTitle() {
  const title = state.settings.class_name || 'GPR Absensi';
  dom.headerClassName.textContent = title;
  document.title = `${title} — Absensi & Kas`;
}

function renderActiveView() {
  closePopover();
  syncHeaderTitle();

  if (state.activeTab === 'absensi') {
    renderAbsensiView(dom.viewAbsensi);
  } else if (state.activeTab === 'kas') {
    renderKasView(dom.viewKas);
  } else if (state.activeTab === 'siswa') {
    renderSiswaView(dom.viewSiswa);
  } else if (state.activeTab === 'kategori') {
    renderKategoriView(dom.viewKategori);
  } else if (state.activeTab === 'pengaturan') {
    renderPengaturanView(dom.viewPengaturan);
  }
}

/**
 * @param {string} monthKey
 * @param {boolean} [forceReload=false]
 */
async function ensureMonthLoaded(monthKey, forceReload = false) {
  if (!forceReload && state.loadedMonths[monthKey]) return;
  setLoading(true, 'Memuat data bulan...');
  try {
    const data = await api.getMonthData(monthKey);
    setMonthData(monthKey, data);
  } finally {
    setLoading(false);
  }
}

/** @param {boolean} [forceReload=false] */
async function ensureKasLoaded(forceReload = false) {
  if (!forceReload && state.kas.loaded) return;
  setLoading(true, 'Memuat data kas...');
  try {
    const data = await api.getKasData();
    setKasData(data);
  } finally {
    setLoading(false);
  }
}

/** @param {string} tabName */
async function switchTab(tabName) {
  const validTabs = ['absensi', 'kas', 'siswa', 'kategori', 'pengaturan'];
  const target = validTabs.includes(tabName) ? tabName : 'absensi';
  setActiveTab(target);

  document.querySelectorAll('[data-tab]').forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.tab === target);
  });

  document.querySelectorAll('[data-view]').forEach((section) => {
    const isCurrent = section.dataset.view === target;
    section.classList.toggle('hidden', !isCurrent);
    section.classList.toggle('active', isCurrent);
  });

  if (target === 'absensi' || target === 'pengaturan') {
    try {
      await ensureMonthLoaded(state.activeMonthKey);
    } catch (err) {
      showToast(err.message, 'error');
    }
  } else if (target === 'kas') {
    try {
      await ensureKasLoaded();
    } catch (err) {
      showToast(err.message, 'error');
    }
  }

  renderActiveView();
}

/** @param {number} delta */
async function changeActiveMonth(delta) {
  closePopover();
  if (state.pendingQueue.size > 0) {
    await flushPendingQueue();
  }

  const nextMonth = shiftMonthKey(state.activeMonthKey, delta);
  state.activeMonthKey = nextMonth;
  try {
    await ensureMonthLoaded(nextMonth);
    renderActiveView();
  } catch (err) {
    showToast(err.message, 'error');
  }
}

/* ==================== CRUD KATEGORI 2 LAPIS ==================== */

/** @param {import('./state.js').Category | null} [existingCat=null] */
function openCategoryModal(existingCat = null) {
  const isEdit = Boolean(existingCat);
  openModal({
    title: isEdit ? 'Ubah Nama Kategori' : 'Tambah Kategori Induk',
    bodyHtml: `
      <form id="modal-category-form">
        <div class="form-group">
          <label class="form-label" for="modal-cat-name">Nama Kategori</label>
          <input
            type="text"
            id="modal-cat-name"
            class="input-control"
            value="${escapeHtml(existingCat?.name || '')}"
            placeholder="Contoh: Jenis Kelamin / KELOMPOK"
            required
          />
        </div>
      </form>
    `,
    footerHtml: `
      <button type="button" class="btn btn-secondary" data-close-modal>Batal</button>
      <button type="submit" form="modal-category-form" id="modal-cat-submit" class="btn btn-primary">
        Simpan
      </button>
    `,
    onMount: (card) => {
      const form = card.querySelector('#modal-category-form');
      const input = card.querySelector('#modal-cat-name');
      const btn = card.querySelector('#modal-cat-submit');
      input.focus();

      form.addEventListener('submit', async (e) => {
        e.preventDefault();
        const name = input.value.trim();
        if (!name) return;

        btn.disabled = true;
        btn.textContent = 'Menyimpan...';
        try {
          const saved = await api.saveCategory({
            category_id: existingCat?.category_id,
            name,
          });
          if (isEdit) {
            const idx = state.categories.findIndex((c) => c.category_id === saved.category_id);
            if (idx >= 0) state.categories[idx] = saved;
          } else {
            state.categories.push(saved);
          }
          closeModal();
          renderActiveView();
          showToast(isEdit ? 'Kategori diperbarui.' : 'Kategori berhasil ditambahkan.', 'success');
        } catch (err) {
          btn.disabled = false;
          btn.textContent = 'Simpan';
          showToast(err.message, 'error');
        }
      });
    },
  });
}

/** @param {string} categoryId */
async function handleDeleteCategory(categoryId) {
  const cat = state.categories.find((c) => c.category_id === categoryId);
  if (!cat) return;

  setLoading(true, 'Memeriksa keterkaitan data...');
  try {
    const res = await api.deleteCategory(categoryId, false);
    setLoading(false);

    if (res.requiresConfirmation) {
      openConfirmModal({
        title: 'Kategori Masih Digunakan',
        message: `Kategori "${cat.name}" masih digunakan oleh ${res.usedCount} siswa aktif. Jika dihapus, seluruh sub-kategorinya akan ikut terhapus dan nilai pada siswa terkait akan diubah menjadi "-".`,
        confirmText: 'Hapus & Ubah Jadi "-"',
        onConfirm: async () => {
          await api.deleteCategory(categoryId, true);
          const fresh = await api.getBootstrap();
          setBootstrapData(fresh);
          if (state.selectedCategoryId === categoryId) {
            state.selectedCategoryId = null;
          }
          renderActiveView();
          showToast('Kategori dihapus.', 'success');
        },
      });
      return;
    }

    const fresh = await api.getBootstrap();
    setBootstrapData(fresh);
    if (state.selectedCategoryId === categoryId) {
      state.selectedCategoryId = null;
    }
    renderActiveView();
    showToast('Kategori dihapus.', 'success');
  } catch (err) {
    setLoading(false);
    showToast(err.message, 'error');
  }
}

/** @param {import('./state.js').SubCategory} existingSub */
function openEditSubCategoryModal(existingSub) {
  openModal({
    title: 'Ubah Nama Sub-Kategori',
    bodyHtml: `
      <form id="modal-sub-form">
        <div class="form-group">
          <label class="form-label" for="modal-sub-name">Nama Sub-Kategori</label>
          <input
            type="text"
            id="modal-sub-name"
            class="input-control"
            value="${escapeHtml(existingSub.name)}"
            required
          />
        </div>
      </form>
    `,
    footerHtml: `
      <button type="button" class="btn btn-secondary" data-close-modal>Batal</button>
      <button type="submit" form="modal-sub-form" id="modal-sub-submit" class="btn btn-primary">
        Simpan
      </button>
    `,
    onMount: (card) => {
      const form = card.querySelector('#modal-sub-form');
      const input = card.querySelector('#modal-sub-name');
      const btn = card.querySelector('#modal-sub-submit');
      input.focus();

      form.addEventListener('submit', async (e) => {
        e.preventDefault();
        const name = input.value.trim();
        if (!name) return;

        btn.disabled = true;
        btn.textContent = 'Menyimpan...';
        try {
          const saved = await api.saveSubCategory({
            subcategory_id: existingSub.subcategory_id,
            category_id: existingSub.category_id,
            name,
          });
          const idx = state.subCategories.findIndex(
            (s) => s.subcategory_id === saved.subcategory_id
          );
          if (idx >= 0) state.subCategories[idx] = saved;
          closeModal();
          renderActiveView();
          showToast('Sub-kategori diperbarui.', 'success');
        } catch (err) {
          btn.disabled = false;
          btn.textContent = 'Simpan';
          showToast(err.message, 'error');
        }
      });
    },
  });
}

/** @param {string} subCategoryId */
async function handleDeleteSubCategory(subCategoryId) {
  const sub = state.subCategories.find((s) => s.subcategory_id === subCategoryId);
  if (!sub) return;

  setLoading(true, 'Memeriksa keterkaitan data...');
  try {
    const res = await api.deleteSubCategory(subCategoryId, false);
    setLoading(false);

    if (res.requiresConfirmation) {
      openConfirmModal({
        title: 'Sub-Kategori Masih Digunakan',
        message: `Sub-kategori "${sub.name}" masih digunakan oleh ${res.usedCount} siswa aktif. Jika dihapus, nilai kategori pada siswa terkait akan diubah menjadi "-".`,
        confirmText: 'Hapus & Ubah Jadi "-"',
        onConfirm: async () => {
          await api.deleteSubCategory(subCategoryId, true);
          const fresh = await api.getBootstrap();
          setBootstrapData(fresh);
          renderActiveView();
          showToast('Sub-kategori dihapus.', 'success');
        },
      });
      return;
    }

    const fresh = await api.getBootstrap();
    setBootstrapData(fresh);
    renderActiveView();
    showToast('Sub-kategori dihapus.', 'success');
  } catch (err) {
    setLoading(false);
    showToast(err.message, 'error');
  }
}

/* ==================== CRUD SISWA ==================== */

/** @param {import('./state.js').Student | null} [existingStudent=null] */
function openStudentModal(existingStudent = null) {
  const isEdit = Boolean(existingStudent);
  const sortedCategories = state.categories.slice().sort((a, b) => a.order_index - b.order_index);

  const categoryFieldsHtml = sortedCategories
    .map((cat) => {
      const subs = getSubCategoriesByCategory(cat.category_id);
      const currentVal = existingStudent?.categories?.[cat.category_id] || '';
      const optionsHtml = subs
        .map(
          (s) =>
            `<option value="${escapeHtml(s.subcategory_id)}" ${
              s.subcategory_id === currentVal ? 'selected' : ''
            }>${escapeHtml(s.name)}</option>`
        )
        .join('');

      return `
        <div class="form-group">
          <label class="form-label" for="stu-cat-${escapeHtml(cat.category_id)}">
            ${escapeHtml(cat.name)}
          </label>
          <select
            id="stu-cat-${escapeHtml(cat.category_id)}"
            class="select-control"
            data-student-category="${escapeHtml(cat.category_id)}"
          >
            <option value="">- Pilih ${escapeHtml(cat.name)} -</option>
            ${optionsHtml}
          </select>
        </div>
      `;
    })
    .join('');

  openModal({
    title: isEdit ? 'Edit Data Siswa' : 'Tambah Siswa Baru',
    bodyHtml: `
      <form id="modal-student-form">
        <div class="form-group">
          <label class="form-label" for="modal-student-name">Nama Lengkap Siswa</label>
          <input
            type="text"
            id="modal-student-name"
            class="input-control"
            value="${escapeHtml(existingStudent?.name || '')}"
            placeholder="Masukkan nama siswa..."
            required
          />
        </div>
        ${categoryFieldsHtml}
      </form>
    `,
    footerHtml: `
      <button type="button" class="btn btn-secondary" data-close-modal>Batal</button>
      <button type="submit" form="modal-student-form" id="modal-student-submit" class="btn btn-primary">
        Simpan
      </button>
    `,
    onMount: (card) => {
      const form = card.querySelector('#modal-student-form');
      const nameInput = card.querySelector('#modal-student-name');
      const btn = card.querySelector('#modal-student-submit');
      nameInput.focus();

      form.addEventListener('submit', async (e) => {
        e.preventDefault();
        const name = nameInput.value.trim();
        if (!name) return;

        const categories = {};
        card.querySelectorAll('[data-student-category]').forEach((select) => {
          if (select.value) {
            categories[select.dataset.studentCategory] = select.value;
          }
        });

        btn.disabled = true;
        btn.textContent = 'Menyimpan...';
        try {
          const saved = await api.saveStudent({
            student_id: existingStudent?.student_id,
            name,
            categories,
          });
          if (isEdit) {
            const idx = state.students.findIndex((s) => s.student_id === saved.student_id);
            if (idx >= 0) state.students[idx] = saved;
          } else {
            state.students.push(saved);
          }
          closeModal();
          renderActiveView();
          showToast(isEdit ? 'Data siswa diperbarui.' : 'Siswa berhasil ditambahkan.', 'success');
        } catch (err) {
          btn.disabled = false;
          btn.textContent = 'Simpan';
          showToast(err.message, 'error');
        }
      });
    },
  });
}

/** @param {string} studentId */
function handleDeleteStudent(studentId) {
  const student = state.students.find((s) => s.student_id === studentId);
  if (!student) return;

  openConfirmModal({
    title: 'Hapus Siswa',
    message: `Hapus "${student.name}" dari daftar siswa aktif? Riwayat absensi & kas pada bulan-bulan sebelumnya akan tetap tersimpan aman.`,
    confirmText: 'Ya, Hapus',
    onConfirm: async () => {
      await api.deleteStudent(studentId);
      student.is_active = false;
      renderActiveView();
      showToast('Siswa dihapus dari daftar aktif.', 'success');
    },
  });
}

/* ==================== PENGATURAN & JADWAL HYBRID ==================== */

/** @param {string} isoDate */
async function handleAddManualDate(isoDate) {
  const monthKey = state.activeMonthKey;
  if (!isoDate || !isoDate.startsWith(`${monthKey}-`)) {
    showToast('Tanggal harus berada di dalam bulan yang sedang dipilih.', 'error');
    return;
  }

  const resolved = resolveMonthDates(
    monthKey,
    state.settings.default_days,
    state.schedulesByMonth[monthKey] || null
  );

  if (resolved.dates.includes(isoDate)) {
    showToast('Tanggal tersebut sudah ada di jadwal.', 'info');
    return;
  }

  const routineSet = new Set(generateRoutineDates(monthKey, state.settings.default_days));
  const nextDates = [...resolved.dates, isoDate].sort();
  const nextRemoved = resolved.removed_dates.filter((d) => d !== isoDate);
  const nextAdded = routineSet.has(isoDate)
    ? resolved.added_dates
    : Array.from(new Set([...resolved.added_dates, isoDate])).sort();

  setLoading(true, 'Menyimpan jadwal...');
  try {
    const saved = await api.saveSchedule({
      month_key: monthKey,
      dates: nextDates,
      added_dates: nextAdded,
      removed_dates: nextRemoved,
    });
    state.schedulesByMonth[monthKey] = saved;
    renderActiveView();
    showToast('Tanggal tambahan berhasil disimpan.', 'success');
  } catch (err) {
    showToast(err.message, 'error');
  } finally {
    setLoading(false);
  }
}

/** @param {string} isoDate */
async function handleRemoveScheduleDate(isoDate) {
  const monthKey = state.activeMonthKey;
  const resolved = resolveMonthDates(
    monthKey,
    state.settings.default_days,
    state.schedulesByMonth[monthKey] || null
  );

  const routineSet = new Set(generateRoutineDates(monthKey, state.settings.default_days));
  const nextDates = resolved.dates.filter((d) => d !== isoDate);
  const nextAdded = resolved.added_dates.filter((d) => d !== isoDate);
  const nextRemoved = routineSet.has(isoDate)
    ? Array.from(new Set([...resolved.removed_dates, isoDate])).sort()
    : resolved.removed_dates;

  setLoading(true, 'Menyimpan jadwal...');
  try {
    const saved = await api.saveSchedule({
      month_key: monthKey,
      dates: nextDates,
      added_dates: nextAdded,
      removed_dates: nextRemoved,
    });
    state.schedulesByMonth[monthKey] = saved;
    renderActiveView();
    showToast('Tanggal pertemuan dihapus.', 'success');
  } catch (err) {
    showToast(err.message, 'error');
  } finally {
    setLoading(false);
  }
}

async function handleResetMonthSchedule() {
  const monthKey = state.activeMonthKey;
  const routineDates = generateRoutineDates(monthKey, state.settings.default_days);

  setLoading(true, 'Mereset jadwal bulan...');
  try {
    const saved = await api.saveSchedule({
      month_key: monthKey,
      dates: routineDates,
      added_dates: [],
      removed_dates: [],
    });
    state.schedulesByMonth[monthKey] = saved;
    renderActiveView();
    showToast('Jadwal bulan ini dikembalikan ke pola hari rutin.', 'success');
  } catch (err) {
    showToast(err.message, 'error');
  } finally {
    setLoading(false);
  }
}

/* ==================== EXPORT REKAP ABSENSI .XLSX ==================== */

async function handleExportExcel() {
  const monthKey = state.activeMonthKey;
  if (state.pendingQueue.size > 0) {
    await flushPendingQueue();
  }

  try {
    await ensureMonthLoaded(monthKey);
    const monthAttendance = state.attendanceByMonth[monthKey] || {};
    const resolved = resolveMonthDates(
      monthKey,
      state.settings.default_days,
      state.schedulesByMonth[monthKey] || null
    );
    const students = getMonthStudents(state.students, monthAttendance);

    if (students.length === 0) {
      showToast('Belum ada data siswa untuk diekspor.', 'error');
      return;
    }

    const { rows, classSummary } = buildMonthlyRecap({
      students,
      categories: state.categories,
      subCategories: state.subCategories,
      activeDates: resolved.dates,
      attendanceMap: monthAttendance,
    });

    await exportMonthlyRecapToExcel({
      className: state.settings.class_name,
      monthKey,
      categories: state.categories,
      rows,
      classSummary,
    });

    showToast('File rekap .xlsx berhasil diunduh.', 'success');
  } catch (err) {
    showToast(err.message || 'Gagal mengekspor file Excel.', 'error');
  }
}

/* ==================== BOOTSTRAP LANGSUNG ==================== */

function renderBootstrapError(message) {
  const activeSection =
    document.querySelector(`[data-view="${state.activeTab}"]`) || dom.viewAbsensi;
  document.querySelectorAll('[data-view]').forEach((section) => {
    const isCurrent = section === activeSection;
    section.classList.toggle('hidden', !isCurrent);
    section.classList.toggle('active', isCurrent);
  });

  activeSection.innerHTML = `
    <div class="card">
      <div class="empty-state">
        <strong>Gagal terhubung ke Google Sheets</strong>
        <p class="form-error" style="max-width: 460px; margin: 10px auto 14px;">${escapeHtml(message)}</p>
        <button type="button" class="btn btn-primary btn-sm" id="btn-retry-bootstrap">
          Coba Muat Ulang
        </button>
      </div>
    </div>
  `;
}

async function bootstrapApplication() {
  setLoading(true, 'Menghubungkan ke Google Sheets...');
  try {
    const [bootstrapData, monthData] = await Promise.all([
      api.getBootstrap(),
      api.getMonthData(state.activeMonthKey),
    ]);
    setBootstrapData(bootstrapData);
    setMonthData(state.activeMonthKey, monthData);
    updateSyncPill('idle');
    await switchTab(state.activeTab);
  } catch (err) {
    updateSyncPill('error');
    renderBootstrapError(err.message);
    showToast(err.message, 'error');
  } finally {
    setLoading(false);
  }
}

/* ==================== EVENT DELEGATION ==================== */

function bindEvents() {
  document.querySelectorAll('[data-tab]').forEach((btn) => {
    btn.addEventListener('click', () => {
      if (btn.dataset.tab === 'kas') {
        state.kas.subView = 'dashboard';
        state.kas.editTxId = null;
      }
      switchTab(btn.dataset.tab);
    });
  });

  dom.syncPill.addEventListener('click', () => {
    if (state.syncStatus === 'error' || state.pendingQueue.size > 0) {
      flushPendingQueue();
    }
  });

  bindKasEvents(dom.viewKas);

  const main = document.getElementById('main-content');
  main.addEventListener('click', async (e) => {
    const target = e.target;

    if (target.closest('#btn-retry-bootstrap')) {
      bootstrapApplication();
      return;
    }

    const gotoBtn = target.closest('[data-goto-tab]');
    if (gotoBtn) {
      switchTab(gotoBtn.dataset.gotoTab);
      return;
    }

    const shiftBtn = target.closest('[data-month-shift]');
    if (shiftBtn) {
      changeActiveMonth(Number(shiftBtn.dataset.monthShift));
      return;
    }

    if (target.closest('[data-export-excel]')) {
      handleExportExcel();
      return;
    }

    const absenCell = target.closest('.absen-cell');
    if (absenCell) {
      const studentId = absenCell.dataset.studentId;
      const date = absenCell.dataset.date;
      openStatusPopover(absenCell, (pickedStatus) => {
        handlePickAttendanceStatus(absenCell, studentId, date, pickedStatus);
      });
      return;
    }

    if (target.closest('#btn-add-student')) {
      openStudentModal(null);
      return;
    }
    const editStuBtn = target.closest('[data-edit-student]');
    if (editStuBtn) {
      const stu = state.students.find((s) => s.student_id === editStuBtn.dataset.editStudent);
      if (stu) openStudentModal(stu);
      return;
    }
    const delStuBtn = target.closest('[data-delete-student]');
    if (delStuBtn) {
      handleDeleteStudent(delStuBtn.dataset.deleteStudent);
      return;
    }

    if (target.closest('#btn-add-category')) {
      openCategoryModal(null);
      return;
    }
    const openCatBtn = target.closest('[data-open-category]');
    if (openCatBtn) {
      state.selectedCategoryId = openCatBtn.dataset.openCategory;
      renderActiveView();
      return;
    }
    if (target.closest('#btn-back-categories')) {
      state.selectedCategoryId = null;
      renderActiveView();
      return;
    }
    const editCatBtn = target.closest('[data-edit-category]');
    if (editCatBtn) {
      const cat = state.categories.find((c) => c.category_id === editCatBtn.dataset.editCategory);
      if (cat) openCategoryModal(cat);
      return;
    }
    const delCatBtn = target.closest('[data-delete-category]');
    if (delCatBtn) {
      handleDeleteCategory(delCatBtn.dataset.deleteCategory);
      return;
    }
    const editSubBtn = target.closest('[data-edit-sub]');
    if (editSubBtn) {
      const sub = state.subCategories.find(
        (s) => s.subcategory_id === editSubBtn.dataset.editSub
      );
      if (sub) openEditSubCategoryModal(sub);
      return;
    }
    const delSubBtn = target.closest('[data-delete-sub]');
    if (delSubBtn) {
      handleDeleteSubCategory(delSubBtn.dataset.deleteSub);
      return;
    }

    const dayBtn = target.closest('[data-toggle-day]');
    if (dayBtn) {
      dayBtn.classList.toggle('selected');
      return;
    }
    const removeDateBtn = target.closest('[data-remove-date]');
    if (removeDateBtn) {
      handleRemoveScheduleDate(removeDateBtn.dataset.removeDate);
      return;
    }
    if (target.closest('#btn-reset-schedule')) {
      handleResetMonthSchedule();
      return;
    }
  });

  main.addEventListener('submit', async (e) => {
    const form = e.target;

    if (form.id === 'form-add-subcategory') {
      e.preventDefault();
      const input = form.querySelector('#input-sub-name');
      const name = input.value.trim();
      if (!name || !state.selectedCategoryId) return;

      const submitBtn = form.querySelector('button[type="submit"]');
      submitBtn.disabled = true;
      try {
        const saved = await api.saveSubCategory({
          category_id: state.selectedCategoryId,
          name,
        });
        state.subCategories.push(saved);
        renderActiveView();
        showToast('Sub-kategori ditambahkan.', 'success');
      } catch (err) {
        submitBtn.disabled = false;
        showToast(err.message, 'error');
      }
      return;
    }

    if (form.id === 'form-settings') {
      e.preventDefault();
      const classNameInput = form.querySelector('#input-class-name');
      const className = classNameInput.value.trim() || 'GPR Absensi';
      const selectedDays = Array.from(form.querySelectorAll('.day-toggle-btn.selected')).map((b) =>
        Number(b.dataset.toggleDay)
      );

      const btn = form.querySelector('#btn-save-settings');
      btn.disabled = true;
      btn.textContent = 'Menyimpan...';
      try {
        await api.saveSettings({
          class_name: className,
          default_days: selectedDays,
        });
        state.settings.class_name = className;
        state.settings.default_days = selectedDays;
        renderActiveView();
        showToast('Pengaturan kelas & hari rutin tersimpan.', 'success');
      } catch (err) {
        btn.disabled = false;
        btn.textContent = 'Simpan Pengaturan';
        showToast(err.message, 'error');
      }
      return;
    }

    if (form.id === 'form-add-manual-date') {
      e.preventDefault();
      const dateInput = form.querySelector('#input-manual-date');
      if (dateInput.value) {
        await handleAddManualDate(dateInput.value);
      }
    }
  });

  document.addEventListener('click', (e) => {
    if (!e.target.closest('.absen-cell') && !e.target.closest('.popover-menu')) {
      closePopover();
    }
  });
  window.addEventListener('scroll', closePopover, true);
  window.addEventListener('resize', closePopover);

  window.addEventListener('beforeunload', (e) => {
    if (state.pendingQueue.size > 0) {
      e.preventDefault();
      e.returnValue = '';
    }
  });
}

function init() {
  initPwaInstall();
  bindEvents();
  bootstrapApplication();
}

init();
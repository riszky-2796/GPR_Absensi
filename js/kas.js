import { EXCEL_COLORS, MONTH_NAMES } from './config.js';
import { api } from './api.js';
import { state } from './state.js';
import {
  escapeHtml,
  showToast,
  setLoading,
  openModal,
  closeModal,
  openConfirmModal,
} from './ui.js';

const MONTH_SHORT = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];

const THIN_BORDER = {
  top: { style: 'thin', color: { argb: EXCEL_COLORS.BORDER_BLACK } },
  left: { style: 'thin', color: { argb: EXCEL_COLORS.BORDER_BLACK } },
  bottom: { style: 'thin', color: { argb: EXCEL_COLORS.BORDER_BLACK } },
  right: { style: 'thin', color: { argb: EXCEL_COLORS.BORDER_BLACK } },
};

/* ==================== HELPERS FORMAT & HITUNG ==================== */

/** @param {number} amount */
function formatRupiah(amount) {
  const num = Math.round(Number(amount) || 0);
  return `Rp ${num.toLocaleString('id-ID')}`;
}

/** @param {string | number} raw */
function formatDots(raw) {
  const digits = String(raw ?? '').replace(/\D/g, '');
  if (!digits) return '';
  return Number(digits).toLocaleString('id-ID');
}

/** @param {string} formatted */
function parseDots(formatted) {
  const digits = String(formatted ?? '').replace(/\D/g, '');
  return digits ? Number(digits) : 0;
}

/** @returns {string} Tanggal lokal hari ini YYYY-MM-DD */
function getTodayIso() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

/** @param {string} isoStr YYYY-MM-DD atau ISO timestamp */
function extractDatePart(isoStr) {
  if (!isoStr) return getTodayIso();
  const match = String(isoStr).match(/^(\d{4}-\d{2}-\d{2})/);
  return match ? match[1] : getTodayIso();
}

/** @param {string} isoDate YYYY-MM-DD */
function formatDateId(isoDate) {
  const part = extractDatePart(isoDate);
  const [y, m, d] = part.split('-').map(Number);
  const mon = MONTH_SHORT[(m || 1) - 1] || '';
  return `${String(d || 1).padStart(2, '0')} ${mon} ${y}`;
}

/** @param {string} monthKey YYYY-MM */
function formatMonthShortYear(monthKey) {
  const [y, m] = String(monthKey).split('-');
  const mon = MONTH_SHORT[Number(m) - 1] || m;
  return `${mon} ${y}`;
}

/** @param {string} studentId */
function getStudentNameById(studentId) {
  const found = state.students.find((s) => s.student_id === studentId);
  return found ? found.name : 'Siswa';
}

/** @param {boolean} isPaid */
function renderIuranBadgeHtml(isPaid) {
  if (isPaid) {
    return `<span class="badge-status badge-status-h">✓</span>`;
  }
  return `<span class="badge-status badge-status-empty">·</span>`;
}

/**
 * Memperbarui angka kolom LUNAS (X/12) pada 1 baris siswa tanpa merender ulang seluruh tabel.
 * @param {HTMLElement} cellEl
 * @param {string} studentId
 */
function updateRowPaidCount(cellEl, studentId) {
  const countCell = cellEl.closest('tr')?.querySelector('[data-kas-student-count]');
  if (!countCell) return;
  const yearPrefix = `${state.kas.activeYear}-`;
  const count = state.kas.iuran.filter(
    (r) => r.student_id === studentId && r.month_key.startsWith(yearPrefix)
  ).length;
  countCell.innerHTML = `<strong>${count}/12</strong>`;
}

/**
 * Menghitung saldo total, pemasukan, pengeluaran, dan gabungan arus transaksi (Single Source of Truth).
 */
function computeKasSummary() {
  let totalIuran = 0;
  let totalTxIn = 0;
  let totalOut = 0;

  state.kas.iuran.forEach((r) => {
    totalIuran += Number(r.amount) || 0;
  });

  state.kas.transactions.forEach((tx) => {
    const amt = Number(tx.amount) || 0;
    if (tx.type === 'out') totalOut += amt;
    else totalTxIn += amt;
  });

  const totalIn = totalIuran + totalTxIn;
  const balance = totalIn - totalOut;

  const combined = [];

  state.kas.iuran.forEach((r) => {
    const datePart = extractDatePart(r.paid_at);
    combined.push({
      kind: 'iuran',
      id: r.iuran_key,
      type: 'in',
      amount: Number(r.amount) || 0,
      title: `${getStudentNameById(r.student_id)} - ${formatMonthShortYear(r.month_key)}`,
      category: 'Iuran Kas',
      date: datePart,
      note: '',
      sortKey: r.paid_at || `${datePart}T00:00:00`,
    });
  });

  state.kas.transactions.forEach((tx) => {
    const datePart = extractDatePart(tx.date);
    const timeSuffix =
      tx.created_at && tx.created_at.length >= 19 ? tx.created_at.slice(10) : 'T12:00:00';
    combined.push({
      kind: 'tx',
      id: tx.tx_id,
      type: tx.type === 'out' ? 'out' : 'in',
      amount: Number(tx.amount) || 0,
      title: tx.title,
      category: tx.category || 'Lainnya',
      date: datePart,
      note: tx.note || '',
      sortKey: `${datePart}${timeSuffix}`,
    });
  });

  combined.sort((a, b) => (b.sortKey > a.sortKey ? 1 : b.sortKey < a.sortKey ? -1 : 0));

  return { totalIuran, totalTxIn, totalIn, totalOut, balance, combined };
}

/* ==================== RENDER SUB-HALAMAN ==================== */

function renderDashboardHtml() {
  const { totalIn, totalOut, balance, combined } = computeKasSummary();
  const recent = combined.slice(0, 15);

  const txListHtml =
    recent.length === 0
      ? `
        <div class="empty-state">
          <strong>Belum ada transaksi kas</strong>
          Mulai catat lewat tombol <b>Bayar Iuran</b>, <b>Kas Masuk</b>, atau <b>Kas Keluar</b> di atas.
        </div>
      `
      : `
        <div class="kas-tx-list">
          ${recent
            .map((item) => {
              const isOut = item.type === 'out';
              const sign = isOut ? '-' : '+';
              const arrow = isOut ? '↑' : '↓';
              const clickableAttr =
                item.kind === 'tx'
                  ? `class="kas-tx-item is-clickable" data-kas-edit-tx="${escapeHtml(item.id)}" title="Ketuk untuk edit / hapus"`
                  : `class="kas-tx-item"`;

              return `
                <div ${clickableAttr}>
                  <div class="kas-tx-left">
                    <span class="kas-tx-icon ${isOut ? 'out' : 'in'}" aria-hidden="true">${arrow}</span>
                    <div class="kas-tx-info">
                      <div class="kas-tx-title">${escapeHtml(item.title)}</div>
                      <div class="kas-tx-meta">${escapeHtml(formatDateId(item.date))} · ${escapeHtml(item.category)}</div>
                    </div>
                  </div>
                  <div class="kas-tx-amount ${isOut ? 'out' : 'in'}">
                    ${sign}${escapeHtml(formatRupiah(item.amount))}
                  </div>
                </div>
              `;
            })
            .join('')}
        </div>
      `;

  return `
    <!-- Kartu Hero Saldo Kas -->
    <div class="kas-hero">
      <div class="kas-hero-top">
        <span class="kas-hero-label">Saldo Kas Saat Ini</span>
        <span class="kas-online-pill"><span class="kas-online-dot"></span> Online</span>
      </div>
      <div class="kas-balance">${escapeHtml(formatRupiah(balance))}</div>
      <div class="kas-hero-split">
        <div class="kas-hero-stat">
          <span class="kas-stat-icon in" aria-hidden="true">↓</span>
          <div style="min-width: 0;">
            <div class="kas-stat-caption">Pemasukan</div>
            <div class="kas-stat-num">${escapeHtml(formatRupiah(totalIn))}</div>
          </div>
        </div>
        <div class="kas-hero-stat">
          <span class="kas-stat-icon out" aria-hidden="true">↑</span>
          <div style="min-width: 0;">
            <div class="kas-stat-caption">Pengeluaran</div>
            <div class="kas-stat-num">${escapeHtml(formatRupiah(totalOut))}</div>
          </div>
        </div>
      </div>
    </div>

    <!-- 3 Tombol Aksi Cepat -->
    <div class="kas-quick-grid">
      <button type="button" class="kas-quick-btn" data-kas-nav="iuran">
        <span class="kas-quick-icon iuran" aria-hidden="true">✓</span>
        <span>Bayar Iuran</span>
      </button>
      <button type="button" class="kas-quick-btn" data-kas-open-form="in">
        <span class="kas-quick-icon in" aria-hidden="true">+</span>
        <span>Kas Masuk</span>
      </button>
      <button type="button" class="kas-quick-btn" data-kas-open-form="out">
        <span class="kas-quick-icon out" aria-hidden="true">−</span>
        <span>Kas Keluar</span>
      </button>
    </div>

    <!-- Menu Pengelolaan (Ringkas) -->
    <div>
      <div class="kas-section-label">Menu Pengelolaan</div>
      <div class="card kas-menu-card">
        <button type="button" class="kas-menu-item" data-kas-open-settings>
          <div class="kas-menu-left">
            <span class="kas-menu-badge blue" aria-hidden="true">⚙</span>
            <div>
              <div class="kas-menu-title">Atur Nominal Iuran</div>
              <div class="kas-menu-sub">${escapeHtml(state.kas.settings.title)} · ${escapeHtml(formatRupiah(state.kas.settings.fee_amount))}/bln</div>
            </div>
          </div>
          <span class="kas-chevron" aria-hidden="true">›</span>
        </button>

        <button type="button" class="kas-menu-item" data-kas-nav="laporan">
          <div class="kas-menu-left">
            <span class="kas-menu-badge purple" aria-hidden="true">📊</span>
            <div>
              <div class="kas-menu-title">Laporan &amp; Ekspor (.xlsx)</div>
              <div class="kas-menu-sub">Rekap iuran &amp; buku kas masuk/keluar</div>
            </div>
          </div>
          <span class="kas-chevron" aria-hidden="true">›</span>
        </button>
      </div>
    </div>

    <!-- Transaksi Terbaru -->
    <div>
      <div class="row-between" style="margin-bottom: 8px;">
        <div class="kas-section-label" style="margin: 0;">Transaksi Terbaru</div>
        ${
          combined.length > 0
            ? `<button type="button" class="kas-add-cat-btn" data-kas-nav="laporan">Lihat Semua</button>`
            : ''
        }
      </div>
      <div class="card" style="padding: 8px 14px;">
        ${txListHtml}
      </div>
    </div>
  `;
}

function renderIuranHtml() {
  const year = state.kas.activeYear;
  const activeStudents = state.students.filter((s) => s.is_active);
  const paidSet = new Set(state.kas.iuran.map((r) => r.iuran_key));

  let tableBodyHtml = '';

  if (activeStudents.length === 0) {
    tableBodyHtml = `
      <div class="empty-state">
        <strong>Belum ada data siswa</strong>
        Tambahkan siswa di menu <b>Siswa</b> terlebih dahulu untuk mencatat pembayaran iuran.
        <div style="margin-top: 12px;">
          <button type="button" class="btn btn-primary btn-sm" data-goto-tab="siswa">Buka Menu Siswa</button>
        </div>
      </div>
    `;
  } else {
    const monthHeaders = MONTH_SHORT.map(
      (label) => `<th class="col-date">${escapeHtml(label.toUpperCase())}</th>`
    ).join('');

    const studentRows = activeStudents
      .map((student, index) => {
        let paidCountYear = 0;

        const monthCells = MONTH_SHORT.map((label, mIdx) => {
          const mm = String(mIdx + 1).padStart(2, '0');
          const monthKey = `${year}-${mm}`;
          const isPaid = paidSet.has(`${student.student_id}|${monthKey}`);
          if (isPaid) paidCountYear++;

          return `
            <td
              class="absen-cell"
              data-kas-toggle-iuran
              data-student-id="${escapeHtml(student.student_id)}"
              data-month-key="${escapeHtml(monthKey)}"
              data-paid="${isPaid ? '1' : '0'}"
              title="${escapeHtml(student.name)} — ${label} ${year}"
            >
              ${renderIuranBadgeHtml(isPaid)}
            </td>
          `;
        }).join('');

        return `
          <tr>
            <td class="sticky-col-no">${index + 1}</td>
            <td class="sticky-col-name" title="${escapeHtml(student.name)}">${escapeHtml(student.name)}</td>
            ${monthCells}
            <td data-kas-student-count><strong>${paidCountYear}/12</strong></td>
          </tr>
        `;
      })
      .join('');

    tableBodyHtml = `
      <div class="table-wrap">
        <table class="data-table">
          <thead>
            <tr>
              <th class="sticky-col-no">NO</th>
              <th class="sticky-col-name">NAMA</th>
              ${monthHeaders}
              <th>LUNAS</th>
            </tr>
          </thead>
          <tbody>
            ${studentRows}
          </tbody>
        </table>
      </div>
      <p class="helper-text">
        Ketuk kotak bulan kosong untuk menandai <b>Lunas (✓)</b>. Ketuk kotak bertanda <b>✓</b> jika ingin membatalkan pembayaran.
      </p>
    `;
  }

  return `
    <div class="kas-sub-header">
      <button type="button" class="btn btn-secondary btn-sm" data-kas-nav="dashboard">← Kembali</button>
      <h2 class="kas-sub-title">Pencatatan Iuran Siswa</h2>
      <span style="width: 68px;"></span>
    </div>

    <!-- Banner Info Tagihan & Pilih Tahun -->
    <div class="kas-iuran-banner">
      <div class="kas-iuran-banner-top">
        <div>
          <h3 class="kas-iuran-title">${escapeHtml(state.kas.settings.title)}</h3>
          <div class="kas-iuran-fee">${escapeHtml(formatRupiah(state.kas.settings.fee_amount))}</div>
        </div>
        <button
          type="button"
          class="kas-banner-icon-btn"
          data-kas-open-settings
          title="Ubah Judul & Nominal Iuran"
          aria-label="Ubah Nominal Iuran"
        >⚙</button>
      </div>

      <div class="kas-year-bar">
        <div class="kas-year-picker">
          <button type="button" class="kas-year-btn" data-kas-year-shift="-1" aria-label="Tahun Sebelumnya">‹</button>
          <span class="kas-year-label">${year}</span>
          <button type="button" class="kas-year-btn" data-kas-year-shift="1" aria-label="Tahun Berikutnya">›</button>
        </div>
        <span class="kas-period-pill">Bulanan (Jan – Des)</span>
      </div>
    </div>

    <!-- Tabel Iuran (Sama Seperti Tabel Absensi) -->
    <div class="card">
      <div class="card-header">
        <div>
          <h3 class="card-title">Tabel Iuran Tahun ${year}</h3>
          <p class="card-desc">${activeStudents.length} siswa aktif · Nominal ${escapeHtml(formatRupiah(state.kas.settings.fee_amount))}/bulan</p>
        </div>
        <div class="kas-legend">
          <span><span class="kas-dot paid"></span>Lunas (✓)</span>
          <span><span class="kas-dot unpaid"></span>Belum (·)</span>
        </div>
      </div>
      ${tableBodyHtml}
    </div>
  `;
}

function renderFormHtml() {
  const editTx = state.kas.editTxId
    ? state.kas.transactions.find((t) => t.tx_id === state.kas.editTxId) || null
    : null;

  const isOut = state.kas.formType === 'out';
  const categories = isOut ? state.kas.settings.categories_out : state.kas.settings.categories_in;
  const selectedCat =
    editTx && editTx.type === state.kas.formType && categories.includes(editTx.category)
      ? editTx.category
      : categories[0] || 'Lainnya';

  const catOptions = Array.from(new Set([...categories, selectedCat]))
    .filter(Boolean)
    .map(
      (cat) =>
        `<option value="${escapeHtml(cat)}" ${cat === selectedCat ? 'selected' : ''}>${escapeHtml(cat)}</option>`
    )
    .join('');

  const pageTitle = editTx
    ? 'Edit Transaksi Kas'
    : isOut
      ? 'Catat Pengeluaran Kas'
      : 'Catat Pemasukan Kas';

  return `
    <div class="kas-sub-header">
      <button type="button" class="btn btn-secondary btn-sm" data-kas-nav="dashboard">← Kembali</button>
      <h2 class="kas-sub-title">${escapeHtml(pageTitle)}</h2>
      <span style="width: 68px;"></span>
    </div>

    <div class="kas-segment">
      <button
        type="button"
        class="kas-segment-btn ${!isOut ? 'active-in' : ''}"
        data-kas-switch-type="in"
      >
        Pemasukan (+)
      </button>
      <button
        type="button"
        class="kas-segment-btn ${isOut ? 'active-out' : ''}"
        data-kas-switch-type="out"
      >
        Pengeluaran (−)
      </button>
    </div>

    <div class="card">
      <form id="kas-tx-form">
        <div class="form-group">
          <label class="form-label" for="kas-input-amount">Nominal Transaksi (Rp)</label>
          <div class="kas-nominal-box">
            <span class="kas-rp-prefix">Rp</span>
            <input
              type="text"
              inputmode="numeric"
              id="kas-input-amount"
              class="kas-nominal-input"
              value="${editTx ? escapeHtml(formatDots(editTx.amount)) : ''}"
              placeholder="0"
              autocomplete="off"
              required
            />
          </div>
        </div>

        <div class="form-group">
          <label class="form-label" for="kas-input-title">Judul / Keperluan</label>
          <input
            type="text"
            id="kas-input-title"
            class="input-control"
            value="${editTx ? escapeHtml(editTx.title) : ''}"
            placeholder="${isOut ? 'Misal: Beli konsumsi rapat' : 'Misal: Donasi sukarela wali murid'}"
            required
          />
        </div>

        <div class="form-group">
          <div class="kas-cat-head">
            <label class="form-label" for="kas-input-category" style="margin: 0;">Kategori</label>
            <button type="button" class="kas-add-cat-btn" data-kas-add-category>+ Tambah</button>
          </div>
          <select id="kas-input-category" class="select-control">
            ${catOptions}
          </select>
        </div>

        <div class="form-group">
          <label class="form-label" for="kas-input-date">Tanggal Transaksi</label>
          <input
            type="date"
            id="kas-input-date"
            class="input-control"
            value="${editTx ? escapeHtml(extractDatePart(editTx.date)) : getTodayIso()}"
            required
          />
        </div>

        <div class="form-group">
          <label class="form-label" for="kas-input-note">Catatan Tambahan (Opsional)</label>
          <input
            type="text"
            id="kas-input-note"
            class="input-control"
            value="${editTx ? escapeHtml(editTx.note) : ''}"
            placeholder="Rincian nota atau rincian item..."
          />
        </div>

        <button
          type="submit"
          id="kas-submit-btn"
          class="btn ${isOut ? 'btn-danger' : 'btn-success'} btn-full"
        >
          ${editTx ? 'Simpan Perubahan' : isOut ? 'Simpan Pengeluaran' : 'Simpan Pemasukan'}
        </button>

        ${
          editTx
            ? `
              <button
                type="button"
                class="btn btn-outline-danger btn-full"
                style="margin-top: 10px;"
                data-kas-delete-tx="${escapeHtml(editTx.tx_id)}"
              >
                Hapus Transaksi Ini
              </button>
            `
            : ''
        }
      </form>
    </div>
  `;
}

function renderLaporanHtml() {
  const year = state.kas.activeYear;
  const { totalIuran, totalTxIn, totalIn, totalOut, balance, combined } = computeKasSummary();

  const monthlyRows = MONTH_SHORT.map((label, idx) => {
    const mk = `${year}-${String(idx + 1).padStart(2, '0')}`;
    let masuk = 0;
    let keluar = 0;
    combined.forEach((item) => {
      if (item.date.startsWith(mk)) {
        if (item.type === 'out') keluar += item.amount;
        else masuk += item.amount;
      }
    });
    return { label, masuk, keluar, net: masuk - keluar };
  }).filter((r) => r.masuk > 0 || r.keluar > 0);

  const monthlyTableHtml =
    monthlyRows.length === 0
      ? `<p class="helper-text">Belum ada mutasi kas tercatat pada tahun ${year}.</p>`
      : `
        <div class="table-wrap">
          <table class="data-table">
            <thead>
              <tr>
                <th class="text-left">BULAN (${year})</th>
                <th>MASUK</th>
                <th>KELUAR</th>
                <th>SELISIH</th>
              </tr>
            </thead>
            <tbody>
              ${monthlyRows
                .map(
                  (r) => `
                    <tr>
                      <td class="text-left"><strong>${r.label} ${year}</strong></td>
                      <td style="color: #15803d; font-weight: 600;">+${escapeHtml(formatRupiah(r.masuk))}</td>
                      <td style="color: var(--danger); font-weight: 600;">−${escapeHtml(formatRupiah(r.keluar))}</td>
                      <td><strong>${escapeHtml(formatRupiah(r.net))}</strong></td>
                    </tr>
                  `
                )
                .join('')}
            </tbody>
          </table>
        </div>
      `;

  const allHistoryHtml =
    combined.length === 0
      ? `<div class="empty-state"><strong>Belum ada riwayat transaksi</strong></div>`
      : `
        <div class="kas-tx-list">
          ${combined
            .map((item) => {
              const isOut = item.type === 'out';
              const clickableAttr =
                item.kind === 'tx'
                  ? `class="kas-tx-item is-clickable" data-kas-edit-tx="${escapeHtml(item.id)}"`
                  : `class="kas-tx-item"`;
              return `
                <div ${clickableAttr}>
                  <div class="kas-tx-left">
                    <span class="kas-tx-icon ${isOut ? 'out' : 'in'}">${isOut ? '↑' : '↓'}</span>
                    <div class="kas-tx-info">
                      <div class="kas-tx-title">${escapeHtml(item.title)}</div>
                      <div class="kas-tx-meta">
                        ${escapeHtml(formatDateId(item.date))} · ${escapeHtml(item.category)}
                        ${item.note ? ` · ${escapeHtml(item.note)}` : ''}
                      </div>
                    </div>
                  </div>
                  <div class="kas-tx-amount ${isOut ? 'out' : 'in'}">
                    ${isOut ? '-' : '+'}${escapeHtml(formatRupiah(item.amount))}
                  </div>
                </div>
              `;
            })
            .join('')}
        </div>
      `;

  return `
    <div class="kas-sub-header">
      <button type="button" class="btn btn-secondary btn-sm" data-kas-nav="dashboard">← Kembali</button>
      <h2 class="kas-sub-title">Laporan &amp; Ekspor Kas</h2>
      <span style="width: 68px;"></span>
    </div>

    <div class="card">
      <div class="card-header">
        <div>
          <h3 class="card-title">Ringkasan Kas Total</h3>
          <p class="card-desc">Iuran Siswa: ${escapeHtml(formatRupiah(totalIuran))} · Kas Masuk Lain: ${escapeHtml(formatRupiah(totalTxIn))}</p>
        </div>
        <div class="row-wrap">
          <div class="month-picker">
            <button type="button" class="icon-btn" data-kas-year-shift="-1" aria-label="Tahun Sebelumnya">‹</button>
            <span class="month-label" style="min-width: 64px;">${year}</span>
            <button type="button" class="icon-btn" data-kas-year-shift="1" aria-label="Tahun Berikutnya">›</button>
          </div>
          <button type="button" class="btn btn-primary btn-sm" data-kas-export-excel>
            Download Laporan (.xlsx)
          </button>
        </div>
      </div>

      <div class="stat-grid">
        <div class="stat-box">
          <div class="stat-label">Saldo Kas Saat Ini</div>
          <div class="stat-value" style="color: var(--primary);">${escapeHtml(formatRupiah(balance))}</div>
        </div>
        <div class="stat-box">
          <div class="stat-label">Total Pemasukan</div>
          <div class="stat-value" style="color: #15803d;">${escapeHtml(formatRupiah(totalIn))}</div>
        </div>
        <div class="stat-box">
          <div class="stat-label">Total Pengeluaran</div>
          <div class="stat-value" style="color: var(--danger);">${escapeHtml(formatRupiah(totalOut))}</div>
        </div>
      </div>
    </div>

    <div class="card">
      <h3 class="card-title" style="margin-bottom: 12px;">Ringkasan Arus Kas Tahun ${year}</h3>
      ${monthlyTableHtml}
    </div>

    <div class="card">
      <div class="card-header">
        <div>
          <h3 class="card-title">Seluruh Riwayat Transaksi (${combined.length})</h3>
          <p class="card-desc">Ketuk catatan Kas Masuk / Kas Keluar untuk mengedit atau menghapus.</p>
        </div>
      </div>
      ${allHistoryHtml}
    </div>
  `;
}

/**
 * Merender tampilan aktif di dalam Tab Kas.
 * @param {HTMLElement} container
 */
export function renderKasView(container) {
  if (!container) return;
  if (state.kas.subView === 'iuran') {
    container.innerHTML = renderIuranHtml();
  } else if (state.kas.subView === 'form') {
    container.innerHTML = renderFormHtml();
  } else if (state.kas.subView === 'laporan') {
    container.innerHTML = renderLaporanHtml();
  } else {
    container.innerHTML = renderDashboardHtml();
  }
}

/* ==================== MODAL PENGATURAN & KATEGORI KAS ==================== */

/** @param {HTMLElement} container */
function openKasSettingsModal(container) {
  openModal({
    title: 'Pengaturan Iuran Kas',
    bodyHtml: `
      <form id="kas-settings-form">
        <div class="form-group">
          <label class="form-label" for="kas-set-title">Nama Tagihan Iuran</label>
          <input
            type="text"
            id="kas-set-title"
            class="input-control"
            value="${escapeHtml(state.kas.settings.title)}"
            placeholder="Contoh: Iuran Kas Bulanan"
            required
          />
        </div>
        <div class="form-group">
          <label class="form-label" for="kas-set-fee">Nominal Tagihan per Bulan (Rp)</label>
          <div class="kas-nominal-box">
            <span class="kas-rp-prefix">Rp</span>
            <input
              type="text"
              inputmode="numeric"
              id="kas-set-fee"
              class="kas-nominal-input"
              value="${escapeHtml(formatDots(state.kas.settings.fee_amount))}"
              required
            />
          </div>
          <p class="helper-text">Perubahan nominal hanya berlaku untuk pembayaran baru ke depan (saldo lama tetap aman).</p>
        </div>
      </form>
    `,
    footerHtml: `
      <button type="button" class="btn btn-secondary" data-close-modal>Batal</button>
      <button type="submit" form="kas-settings-form" id="kas-set-submit" class="btn btn-primary">Simpan</button>
    `,
    onMount: (card) => {
      const form = card.querySelector('#kas-settings-form');
      const titleInput = card.querySelector('#kas-set-title');
      const feeInput = card.querySelector('#kas-set-fee');
      const btn = card.querySelector('#kas-set-submit');

      feeInput.addEventListener('input', () => {
        feeInput.value = formatDots(feeInput.value);
      });

      form.addEventListener('submit', async (e) => {
        e.preventDefault();
        const title = titleInput.value.trim() || 'Iuran Kas Bulanan';
        const fee = parseDots(feeInput.value);

        btn.disabled = true;
        btn.textContent = 'Menyimpan...';
        try {
          const updated = await api.saveKasSettings({ title, fee_amount: fee });
          state.kas.settings = updated;
          closeModal();
          renderKasView(container);
          showToast('Pengaturan nominal iuran tersimpan.', 'success');
        } catch (err) {
          btn.disabled = false;
          btn.textContent = 'Simpan';
          showToast(err.message, 'error');
        }
      });
    },
  });
}

/** @param {HTMLElement} container */
function openAddKasCategoryModal(container) {
  const isOut = state.kas.formType === 'out';
  openModal({
    title: isOut ? 'Tambah Kategori Pengeluaran' : 'Tambah Kategori Pemasukan',
    bodyHtml: `
      <form id="kas-add-cat-form">
        <div class="form-group">
          <label class="form-label" for="kas-new-cat-name">Nama Kategori Baru</label>
          <input
            type="text"
            id="kas-new-cat-name"
            class="input-control"
            placeholder="${isOut ? 'Misal: Fotokopi / Kebersihan' : 'Misal: Sponsor / Hasil Usaha'}"
            required
          />
        </div>
      </form>
    `,
    footerHtml: `
      <button type="button" class="btn btn-secondary" data-close-modal>Batal</button>
      <button type="submit" form="kas-add-cat-form" id="kas-new-cat-submit" class="btn btn-primary">Tambah</button>
    `,
    onMount: (card) => {
      const form = card.querySelector('#kas-add-cat-form');
      const input = card.querySelector('#kas-new-cat-name');
      const btn = card.querySelector('#kas-new-cat-submit');
      input.focus();

      form.addEventListener('submit', async (e) => {
        e.preventDefault();
        const name = input.value.trim();
        if (!name) return;

        const key = isOut ? 'categories_out' : 'categories_in';
        const currentList = state.kas.settings[key] || [];
        if (currentList.some((c) => c.toLowerCase() === name.toLowerCase())) {
          showToast('Kategori tersebut sudah ada.', 'info');
          return;
        }

        const nextList = [...currentList, name];
        btn.disabled = true;
        btn.textContent = 'Menyimpan...';
        try {
          const updated = await api.saveKasSettings({ [key]: nextList });
          state.kas.settings = updated;
          closeModal();

          const select = container.querySelector('#kas-input-category');
          if (select) {
            const opt = document.createElement('option');
            opt.value = name;
            opt.textContent = name;
            opt.selected = true;
            select.appendChild(opt);
          }
          showToast(`Kategori "${name}" ditambahkan.`, 'success');
        } catch (err) {
          btn.disabled = false;
          btn.textContent = 'Tambah';
          showToast(err.message, 'error');
        }
      });
    },
  });
}

/* ==================== EXPORT LAPORAN KAS (.XLSX 2 SHEET) ==================== */

async function exportKasToExcel() {
  if (typeof window.ExcelJS === 'undefined') {
    throw new Error('Library ExcelJS belum termuat. Periksa koneksi internet lalu muat ulang halaman.');
  }

  const year = state.kas.activeYear;
  const className = String(state.settings.class_name || 'GPR ABSENSI').toUpperCase();
  const workbook = new window.ExcelJS.Workbook();

  /* ---------- SHEET 1: REKAP IURAN SISWA ---------- */
  const wsIuran = workbook.addWorksheet(`Rekap Iuran ${year}`);
  const activeStudents = state.students.filter((s) => s.is_active);
  const iuranMap = {};
  state.kas.iuran.forEach((r) => {
    iuranMap[r.iuran_key] = r;
  });

  const iuranHeaders = ['NO', 'NAMA SISWA', ...MONTH_SHORT.map((m) => m.toUpperCase()), 'TOTAL LUNAS', 'TOTAL BAYAR (Rp)'];
  const totalCols1 = iuranHeaders.length;

  const titleRow1 = wsIuran.addRow([`REKAP IURAN KAS ${className} TAHUN ${year}`]);
  wsIuran.mergeCells(1, 1, 1, totalCols1);
  titleRow1.height = 26;
  titleRow1.getCell(1).font = { name: 'Calibri', size: 12, bold: true };
  titleRow1.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };

  const headerRow1 = wsIuran.addRow(iuranHeaders);
  headerRow1.height = 28;
  headerRow1.eachCell((cell) => {
    cell.font = { name: 'Calibri', size: 11, bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: EXCEL_COLORS.HEADER_BG } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = THIN_BORDER;
  });

  let grandTotalIuranYear = 0;
  activeStudents.forEach((student, idx) => {
    let countPaid = 0;
    let sumPaid = 0;
    const monthCells = MONTH_SHORT.map((_, mIdx) => {
      const mk = `${year}-${String(mIdx + 1).padStart(2, '0')}`;
      const rec = iuranMap[`${student.student_id}|${mk}`];
      if (rec) {
        countPaid++;
        sumPaid += Number(rec.amount) || 0;
        return '✓';
      }
      return '-';
    });

    grandTotalIuranYear += sumPaid;
    const row = wsIuran.addRow([idx + 1, student.name, ...monthCells, `${countPaid}/12`, sumPaid]);
    row.height = 21;
    const bg = idx % 2 === 0 ? EXCEL_COLORS.ZEBRA_ODD_BG : EXCEL_COLORS.ZEBRA_EVEN_BG;

    row.eachCell((cell, colNum) => {
      cell.font = { name: 'Calibri', size: 11 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
      cell.alignment = { vertical: 'middle', horizontal: colNum === 2 ? 'left' : colNum === totalCols1 ? 'right' : 'center' };
      cell.border = THIN_BORDER;
      if (colNum === totalCols1) cell.numFmt = '#,##0';
    });
  });

  const foot1 = new Array(totalCols1).fill('');
  foot1[0] = `TOTAL IURAN TERKUMPUL TAHUN ${year}`;
  foot1[totalCols1 - 1] = grandTotalIuranYear;
  const footRow1 = wsIuran.addRow(foot1);
  wsIuran.mergeCells(footRow1.number, 1, footRow1.number, totalCols1 - 1);
  footRow1.height = 24;
  footRow1.eachCell({ includeEmpty: true }, (cell, colNum) => {
    cell.font = { name: 'Calibri', size: 11, bold: true };
    cell.alignment = { vertical: 'middle', horizontal: colNum === totalCols1 ? 'right' : 'center' };
    cell.border = THIN_BORDER;
    if (colNum === totalCols1) cell.numFmt = '#,##0';
  });

  wsIuran.columns = [
    { width: 6 },
    { width: 28 },
    ...MONTH_SHORT.map(() => ({ width: 7 })),
    { width: 13 },
    { width: 18 },
  ];

  /* ---------- SHEET 2: BUKU KAS MASUK & KELUAR ---------- */
  const wsBuku = workbook.addWorksheet('Buku Kas');
  const { combined } = computeKasSummary();
  const chronological = combined.slice().reverse();

  const bukuHeaders = ['NO', 'TANGGAL', 'JENIS', 'KATEGORI', 'KETERANGAN', 'MASUK (Rp)', 'KELUAR (Rp)', 'SALDO (Rp)'];
  const titleRow2 = wsBuku.addRow([`BUKU KAS ${className}`]);
  wsBuku.mergeCells(1, 1, 1, bukuHeaders.length);
  titleRow2.height = 26;
  titleRow2.getCell(1).font = { name: 'Calibri', size: 12, bold: true };
  titleRow2.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };

  const headerRow2 = wsBuku.addRow(bukuHeaders);
  headerRow2.height = 28;
  headerRow2.eachCell((cell) => {
    cell.font = { name: 'Calibri', size: 11, bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: EXCEL_COLORS.HEADER_BG } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = THIN_BORDER;
  });

  let runningBalance = 0;
  let sumIn = 0;
  let sumOut = 0;

  chronological.forEach((item, idx) => {
    const masuk = item.type === 'in' ? item.amount : 0;
    const keluar = item.type === 'out' ? item.amount : 0;
    sumIn += masuk;
    sumOut += keluar;
    runningBalance += masuk - keluar;

    const ket = item.note ? `${item.title} (${item.note})` : item.title;
    const row = wsBuku.addRow([
      idx + 1,
      formatDateId(item.date),
      item.type === 'in' ? 'Masuk' : 'Keluar',
      item.category,
      ket,
      masuk || '',
      keluar || '',
      runningBalance,
    ]);
    row.height = 21;
    const bg = idx % 2 === 0 ? EXCEL_COLORS.ZEBRA_ODD_BG : EXCEL_COLORS.ZEBRA_EVEN_BG;

    row.eachCell({ includeEmpty: true }, (cell, colNum) => {
      cell.font = { name: 'Calibri', size: 11 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
      cell.alignment = {
        vertical: 'middle',
        horizontal: colNum === 5 ? 'left' : colNum >= 6 ? 'right' : 'center',
      };
      cell.border = THIN_BORDER;
      if (colNum >= 6 && cell.value !== '') cell.numFmt = '#,##0';
    });
  });

  const footRow2 = wsBuku.addRow(['TOTAL & SALDO AKHIR', '', '', '', '', sumIn, sumOut, runningBalance]);
  wsBuku.mergeCells(footRow2.number, 1, footRow2.number, 5);
  footRow2.height = 24;
  footRow2.eachCell({ includeEmpty: true }, (cell, colNum) => {
    cell.font = { name: 'Calibri', size: 11, bold: true };
    cell.alignment = { vertical: 'middle', horizontal: colNum >= 6 ? 'right' : 'center' };
    cell.border = THIN_BORDER;
    if (colNum >= 6) cell.numFmt = '#,##0';
  });

  wsBuku.columns = [
    { width: 6 },
    { width: 15 },
    { width: 10 },
    { width: 18 },
    { width: 34 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
  ];

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const safeClass = String(state.settings.class_name || 'GPR_Absensi')
    .trim()
    .replace(/[^a-zA-Z0-9_-]+/g, '_');
  const todayMonth = MONTH_NAMES[new Date().getMonth()] || String(year);
  const fileName = `Laporan_Kas_${safeClass}_${todayMonth}_${year}.xlsx`;

  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/* ==================== EVENT BINDING TAB KAS ==================== */

/**
 * Mengikat seluruh event klik, input, dan submit khusus di dalam #view-kas.
 * @param {HTMLElement} container
 */
export function bindKasEvents(container) {
  if (!container) return;

  container.addEventListener('input', (e) => {
    const input = e.target;
    if (input && input.id === 'kas-input-amount') {
      input.value = formatDots(input.value);
    }
  });

  container.addEventListener('click', async (e) => {
    const target = e.target;

    const navBtn = target.closest('[data-kas-nav]');
    if (navBtn) {
      state.kas.subView = navBtn.dataset.kasNav;
      state.kas.editTxId = null;
      renderKasView(container);
      return;
    }

    const formBtn = target.closest('[data-kas-open-form]');
    if (formBtn) {
      state.kas.formType = formBtn.dataset.kasOpenForm === 'out' ? 'out' : 'in';
      state.kas.editTxId = null;
      state.kas.subView = 'form';
      renderKasView(container);
      return;
    }

    const switchBtn = target.closest('[data-kas-switch-type]');
    if (switchBtn) {
      const nextType = switchBtn.dataset.kasSwitchType === 'out' ? 'out' : 'in';
      if (state.kas.formType !== nextType) {
        state.kas.formType = nextType;
        renderKasView(container);
      }
      return;
    }

    if (target.closest('[data-kas-open-settings]')) {
      openKasSettingsModal(container);
      return;
    }

    if (target.closest('[data-kas-add-category]')) {
      openAddKasCategoryModal(container);
      return;
    }

    const yearBtn = target.closest('[data-kas-year-shift]');
    if (yearBtn) {
      state.kas.activeYear += Number(yearBtn.dataset.kasYearShift);
      renderKasView(container);
      return;
    }

    const editTxEl = target.closest('[data-kas-edit-tx]');
    if (editTxEl) {
      const tx = state.kas.transactions.find((t) => t.tx_id === editTxEl.dataset.kasEditTx);
      if (tx) {
        state.kas.editTxId = tx.tx_id;
        state.kas.formType = tx.type === 'out' ? 'out' : 'in';
        state.kas.subView = 'form';
        renderKasView(container);
      }
      return;
    }

    const delTxBtn = target.closest('[data-kas-delete-tx]');
    if (delTxBtn) {
      const txId = delTxBtn.dataset.kasDeleteTx;
      const tx = state.kas.transactions.find((t) => t.tx_id === txId);
      if (!tx) return;

      openConfirmModal({
        title: 'Hapus Transaksi Kas',
        message: `Hapus catatan "${tx.title}" senilai ${formatRupiah(tx.amount)}? Saldo kas akan otomatis disesuaikan.`,
        confirmText: 'Ya, Hapus',
        onConfirm: async () => {
          await api.deleteKasTransaction(txId);
          state.kas.transactions = state.kas.transactions.filter((t) => t.tx_id !== txId);
          state.kas.editTxId = null;
          state.kas.subView = 'dashboard';
          renderKasView(container);
          showToast('Transaksi dihapus.', 'success');
        },
      });
      return;
    }

    // 1-Klik Sel Tabel Iuran (Update langsung di sel tanpa mereset posisi scroll horizontal)
    const iuranCell = target.closest('[data-kas-toggle-iuran]');
    if (iuranCell) {
      if (iuranCell.dataset.busy === '1') return;
      const studentId = iuranCell.dataset.studentId;
      const monthKey = iuranCell.dataset.monthKey;
      const alreadyPaid = iuranCell.dataset.paid === '1';

      if (!alreadyPaid) {
        iuranCell.dataset.busy = '1';
        iuranCell.style.opacity = '0.5';
        try {
          const res = await api.toggleKasIuran({
            student_id: studentId,
            month_key: monthKey,
            paid: true,
            amount: state.kas.settings.fee_amount,
          });
          if (res.record) {
            state.kas.iuran = state.kas.iuran.filter((r) => r.iuran_key !== res.record.iuran_key);
            state.kas.iuran.push(res.record);
          }
          iuranCell.dataset.paid = '1';
          iuranCell.innerHTML = renderIuranBadgeHtml(true);
          updateRowPaidCount(iuranCell, studentId);
          showToast(`Lunas: ${getStudentNameById(studentId)} (${formatMonthShortYear(monthKey)})`, 'success');
        } catch (err) {
          showToast(err.message, 'error');
        } finally {
          delete iuranCell.dataset.busy;
          iuranCell.style.opacity = '';
        }
      } else {
        openConfirmModal({
          title: 'Batalkan Status Lunas?',
          message: `Batalkan pembayaran iuran ${getStudentNameById(studentId)} untuk bulan ${formatMonthShortYear(monthKey)}? Saldo kas akan dikurangi kembali.`,
          confirmText: 'Ya, Batalkan',
          onConfirm: async () => {
            const res = await api.toggleKasIuran({
              student_id: studentId,
              month_key: monthKey,
              paid: false,
            });
            state.kas.iuran = state.kas.iuran.filter((r) => r.iuran_key !== res.iuran_key);
            iuranCell.dataset.paid = '0';
            iuranCell.innerHTML = renderIuranBadgeHtml(false);
            updateRowPaidCount(iuranCell, studentId);
            showToast('Status lunas dibatalkan.');
          },
        });
      }
      return;
    }

    if (target.closest('[data-kas-export-excel]')) {
      setLoading(true, 'Menyiapkan laporan Excel...');
      try {
        await exportKasToExcel();
        showToast('Laporan Kas (.xlsx) berhasil diunduh.', 'success');
      } catch (err) {
        showToast(err.message, 'error');
      } finally {
        setLoading(false);
      }
    }
  });

  container.addEventListener('submit', async (e) => {
    const form = e.target;
    if (!form || form.id !== 'kas-tx-form') return;
    e.preventDefault();

    const amount = parseDots(form.querySelector('#kas-input-amount').value);
    const title = form.querySelector('#kas-input-title').value.trim();
    const category = form.querySelector('#kas-input-category').value;
    const date = form.querySelector('#kas-input-date').value;
    const note = form.querySelector('#kas-input-note').value.trim();

    if (amount <= 0) {
      showToast('Nominal transaksi harus lebih dari Rp 0.', 'error');
      return;
    }
    if (!title) {
      showToast('Judul / keperluan wajib diisi.', 'error');
      return;
    }

    const submitBtn = form.querySelector('#kas-submit-btn');
    submitBtn.disabled = true;
    submitBtn.textContent = 'Menyimpan...';

    try {
      const saved = await api.saveKasTransaction({
        tx_id: state.kas.editTxId || undefined,
        type: state.kas.formType,
        amount,
        title,
        category,
        date,
        note,
      });

      if (state.kas.editTxId) {
        const idx = state.kas.transactions.findIndex((t) => t.tx_id === saved.tx_id);
        if (idx >= 0) state.kas.transactions[idx] = saved;
      } else {
        state.kas.transactions.push(saved);
      }

      const wasEdit = Boolean(state.kas.editTxId);
      state.kas.editTxId = null;
      state.kas.subView = 'dashboard';
      renderKasView(container);
      showToast(wasEdit ? 'Transaksi diperbarui.' : 'Transaksi kas tersimpan.', 'success');
    } catch (err) {
      submitBtn.disabled = false;
      submitBtn.textContent = state.kas.formType === 'out' ? 'Simpan Pengeluaran' : 'Simpan Pemasukan';
      showToast(err.message, 'error');
    }
  });
}
/**
 * GPR Absensi — Backend API (Google Apps Script)
 * Database : Google Sheets (6 Tab)
 * Auth     : PIN disimpan di Project Settings > Script Properties dengan key "APP_PIN"
 */

const SHEETS = {
  SETTINGS: {
    name: 'Settings',
    headers: ['key', 'value'],
  },
  CATEGORIES: {
    name: 'Categories',
    headers: ['category_id', 'name', 'order_index'],
  },
  SUBCATEGORIES: {
    name: 'SubCategories',
    headers: ['subcategory_id', 'category_id', 'name'],
  },
  STUDENTS: {
    name: 'Students',
    headers: ['student_id', 'name', 'categories_json', 'is_active', 'created_at'],
  },
  SCHEDULES: {
    name: 'Schedules',
    headers: ['month_key', 'dates_json', 'added_dates_json', 'removed_dates_json'],
  },
  ATTENDANCE: {
    name: 'Attendance',
    headers: ['record_key', 'month_key', 'student_id', 'date', 'status', 'updated_at'],
  },
};

const VALID_STATUS = { H: true, I: true, S: true, A: true, '': true };

/**
 * Jalankan fungsi ini SEKALI dari editor Apps Script untuk membuat 6 tab database otomatis.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(SHEETS).forEach(function (key) {
    const def = SHEETS[key];
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
    }

    if (sheet.getLastRow() === 0) {
      sheet
        .getRange(1, 1, 1, def.headers.length)
        .setValues([def.headers])
        .setFontWeight('bold')
        .setBackground('#BDD7EE');
      sheet.setFrozenRows(1);
    }

    // Format seluruh kolom sebagai Plain Text agar format tanggal YYYY-MM-DD tidak berubah otomatis
    sheet.getRange(1, 1, sheet.getMaxRows(), def.headers.length).setNumberFormat('@');
  });

  // Isi pengaturan awal default jika tab Settings masih kosong
  const settingsSheet = ss.getSheetByName(SHEETS.SETTINGS.name);
  if (settingsSheet.getLastRow() === 1) {
    const defaults = [
      ['class_name', 'GPR C'],
      ['default_days', '[]'],
    ];
    settingsSheet.getRange(2, 1, defaults.length, 2).setNumberFormat('@').setValues(defaults);
  }
}

/**
 * Health check publik (tidak membocorkan data apa pun).
 * Buka URL Web App (/exec) di browser untuk memastikan deployment aktif.
 */
function doGet() {
  const pinConfigured = Boolean(getAppPin());
  return jsonResponse({
    ok: true,
    service: 'GPR Absensi API',
    pinConfigured: pinConfigured,
  });
}

/**
 * Endpoint utama seluruh operasi aplikasi (Read & Write terproteksi PIN).
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ ok: false, error: 'Payload request kosong.' });
    }

    const body = JSON.parse(e.postData.contents);
    const savedPin = getAppPin();

    if (!savedPin) {
      return jsonResponse({
        ok: false,
        error: 'APP_PIN belum diatur di Properti Skrip Apps Script.',
      });
    }

    if (String(body.pin || '') !== savedPin) {
      return jsonResponse({
        ok: false,
        unauthorized: true,
        error: 'PIN salah. Silakan periksa kembali PIN Anda.',
      });
    }

    const action = String(body.action || '');
    const payload = body.payload || {};

    if (action === 'verifyPin') {
      return jsonResponse({ ok: true });
    }

    // Operasi baca (tanpa lock)
    if (action === 'getBootstrap') {
      return jsonResponse({ ok: true, data: handleGetBootstrap() });
    }
    if (action === 'getMonthData') {
      return jsonResponse({ ok: true, data: handleGetMonthData(payload) });
    }

    // Operasi tulis (menggunakan ScriptLock untuk mencegah tabrakan data)
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      let data;
      switch (action) {
        case 'saveSettings':
          data = handleSaveSettings(payload);
          break;
        case 'saveCategory':
          data = handleSaveCategory(payload);
          break;
        case 'deleteCategory':
          data = handleDeleteCategory(payload);
          break;
        case 'saveSubCategory':
          data = handleSaveSubCategory(payload);
          break;
        case 'deleteSubCategory':
          data = handleDeleteSubCategory(payload);
          break;
        case 'saveStudent':
          data = handleSaveStudent(payload);
          break;
        case 'deleteStudent':
          data = handleDeleteStudent(payload);
          break;
        case 'saveSchedule':
          data = handleSaveSchedule(payload);
          break;
        case 'batchSaveAttendance':
          data = handleBatchSaveAttendance(payload);
          break;
        default:
          return jsonResponse({ ok: false, error: 'Aksi tidak dikenali: ' + action });
      }
      return jsonResponse({ ok: true, data: data });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return jsonResponse({
      ok: false,
      error: err && err.message ? err.message : 'Terjadi kesalahan pada server Apps Script.',
    });
  }
}

/* ===========================
 * HANDLER DATA & LOGIKA CRUD
 * =========================== */

function handleGetBootstrap() {
  const settingsRows = readSheetObjects(SHEETS.SETTINGS.name);
  const settingsMap = {};
  settingsRows.forEach(function (r) {
    settingsMap[r.key] = r.value;
  });

  const settings = {
    class_name: settingsMap.class_name || 'GPR Absensi',
    default_days: safeJsonParse(settingsMap.default_days, []),
  };

  const categories = readSheetObjects(SHEETS.CATEGORIES.name)
    .map(function (r) {
      return {
        category_id: r.category_id,
        name: r.name,
        order_index: Number(r.order_index) || 0,
      };
    })
    .sort(function (a, b) {
      return a.order_index - b.order_index;
    });

  const subCategories = readSheetObjects(SHEETS.SUBCATEGORIES.name).map(function (r) {
    return {
      subcategory_id: r.subcategory_id,
      category_id: r.category_id,
      name: r.name,
    };
  });

  const students = readSheetObjects(SHEETS.STUDENTS.name).map(function (r) {
    return {
      student_id: r.student_id,
      name: r.name,
      categories: safeJsonParse(r.categories_json, {}),
      is_active: String(r.is_active).toUpperCase() === 'TRUE',
      created_at: r.created_at || '',
    };
  });

  return {
    settings: settings,
    categories: categories,
    subCategories: subCategories,
    students: students,
  };
}

function handleSaveSettings(payload) {
  const sheet = getSheet(SHEETS.SETTINGS.name);
  const rows = readSheetObjects(SHEETS.SETTINGS.name);
  const updates = {};

  if (typeof payload.class_name === 'string') {
    updates.class_name = payload.class_name.trim() || 'GPR Absensi';
  }
  if (Array.isArray(payload.default_days)) {
    const cleanDays = payload.default_days
      .map(Number)
      .filter(function (d) {
        return d >= 0 && d <= 6;
      });
    updates.default_days = JSON.stringify(cleanDays);
  }

  Object.keys(updates).forEach(function (key) {
    const existing = rows.find(function (r) {
      return r.key === key;
    });
    if (existing) {
      sheet.getRange(existing._row, 2).setNumberFormat('@').setValue(updates[key]);
    } else {
      appendTextRows(sheet, [[key, updates[key]]]);
    }
  });

  return {
    class_name: updates.class_name,
    default_days: updates.default_days ? safeJsonParse(updates.default_days, []) : undefined,
  };
}

function handleSaveCategory(payload) {
  const name = String(payload.name || '').trim();
  if (!name) throw new Error('Nama kategori wajib diisi.');

  const sheet = getSheet(SHEETS.CATEGORIES.name);
  const rows = readSheetObjects(SHEETS.CATEGORIES.name);

  if (payload.category_id) {
    const existing = rows.find(function (r) {
      return r.category_id === payload.category_id;
    });
    if (!existing) throw new Error('Kategori tidak ditemukan.');
    const orderIndex =
      payload.order_index !== undefined ? Number(payload.order_index) : Number(existing.order_index) || 0;
    sheet
      .getRange(existing._row, 2, 1, 2)
      .setNumberFormat('@')
      .setValues([[name, String(orderIndex)]]);
    return { category_id: existing.category_id, name: name, order_index: orderIndex };
  }

  const newId = generateId('cat');
  const nextOrder = rows.length + 1;
  appendTextRows(sheet, [[newId, name, String(nextOrder)]]);
  return { category_id: newId, name: name, order_index: nextOrder };
}

function handleDeleteCategory(payload) {
  const categoryId = String(payload.category_id || '').trim();
  const force = Boolean(payload.force);
  if (!categoryId) throw new Error('ID kategori tidak valid.');

  const studentSheet = getSheet(SHEETS.STUDENTS.name);
  const studentRows = readSheetObjects(SHEETS.STUDENTS.name);
  const affectedStudents = studentRows.filter(function (s) {
    const cats = safeJsonParse(s.categories_json, {});
    return String(s.is_active).toUpperCase() === 'TRUE' && Boolean(cats[categoryId]);
  });

  if (affectedStudents.length > 0 && !force) {
    return { requiresConfirmation: true, usedCount: affectedStudents.length };
  }

  // Hapus relasi kategori dari seluruh siswa yang memakainya (nilai menjadi kosong / "-")
  studentRows.forEach(function (s) {
    const cats = safeJsonParse(s.categories_json, {});
    if (Object.prototype.hasOwnProperty.call(cats, categoryId)) {
      delete cats[categoryId];
      studentSheet.getRange(s._row, 3).setNumberFormat('@').setValue(JSON.stringify(cats));
    }
  });

  // Hapus seluruh sub-kategori di bawah kategori induk ini (dari bawah ke atas)
  deleteMatchingRows(SHEETS.SUBCATEGORIES.name, function (r) {
    return r.category_id === categoryId;
  });

  // Hapus kategori induk
  deleteMatchingRows(SHEETS.CATEGORIES.name, function (r) {
    return r.category_id === categoryId;
  });

  return { deleted: true, category_id: categoryId };
}

function handleSaveSubCategory(payload) {
  const categoryId = String(payload.category_id || '').trim();
  const name = String(payload.name || '').trim();
  if (!categoryId || !name) throw new Error('Kategori induk dan nama sub-kategori wajib diisi.');

  const sheet = getSheet(SHEETS.SUBCATEGORIES.name);
  const rows = readSheetObjects(SHEETS.SUBCATEGORIES.name);

  if (payload.subcategory_id) {
    const existing = rows.find(function (r) {
      return r.subcategory_id === payload.subcategory_id;
    });
    if (!existing) throw new Error('Sub-kategori tidak ditemukan.');
    sheet.getRange(existing._row, 3).setNumberFormat('@').setValue(name);
    return { subcategory_id: existing.subcategory_id, category_id: existing.category_id, name: name };
  }

  const newId = generateId('sub');
  appendTextRows(sheet, [[newId, categoryId, name]]);
  return { subcategory_id: newId, category_id: categoryId, name: name };
}

function handleDeleteSubCategory(payload) {
  const subId = String(payload.subcategory_id || '').trim();
  const force = Boolean(payload.force);
  if (!subId) throw new Error('ID sub-kategori tidak valid.');

  const studentSheet = getSheet(SHEETS.STUDENTS.name);
  const studentRows = readSheetObjects(SHEETS.STUDENTS.name);
  const affectedActive = studentRows.filter(function (s) {
    if (String(s.is_active).toUpperCase() !== 'TRUE') return false;
    const cats = safeJsonParse(s.categories_json, {});
    return Object.keys(cats).some(function (k) {
      return cats[k] === subId;
    });
  });

  if (affectedActive.length > 0 && !force) {
    return { requiresConfirmation: true, usedCount: affectedActive.length };
  }

  // Kosongkan nilai sub-kategori pada siswa terkait (tampil sebagai "-")
  studentRows.forEach(function (s) {
    const cats = safeJsonParse(s.categories_json, {});
    let changed = false;
    Object.keys(cats).forEach(function (k) {
      if (cats[k] === subId) {
        delete cats[k];
        changed = true;
      }
    });
    if (changed) {
      studentSheet.getRange(s._row, 3).setNumberFormat('@').setValue(JSON.stringify(cats));
    }
  });

  deleteMatchingRows(SHEETS.SUBCATEGORIES.name, function (r) {
    return r.subcategory_id === subId;
  });

  return { deleted: true, subcategory_id: subId };
}

function handleSaveStudent(payload) {
  const name = String(payload.name || '').trim();
  if (!name) throw new Error('Nama siswa wajib diisi.');

  const categories = payload.categories && typeof payload.categories === 'object' ? payload.categories : {};
  const categoriesJson = JSON.stringify(categories);
  const sheet = getSheet(SHEETS.STUDENTS.name);
  const rows = readSheetObjects(SHEETS.STUDENTS.name);

  if (payload.student_id) {
    const existing = rows.find(function (r) {
      return r.student_id === payload.student_id;
    });
    if (!existing) throw new Error('Data siswa tidak ditemukan.');
    sheet
      .getRange(existing._row, 2, 1, 2)
      .setNumberFormat('@')
      .setValues([[name, categoriesJson]]);
    return {
      student_id: existing.student_id,
      name: name,
      categories: categories,
      is_active: String(existing.is_active).toUpperCase() === 'TRUE',
      created_at: existing.created_at,
    };
  }

  const newId = generateId('stu');
  const now = new Date().toISOString();
  appendTextRows(sheet, [[newId, name, categoriesJson, 'TRUE', now]]);
  return {
    student_id: newId,
    name: name,
    categories: categories,
    is_active: true,
    created_at: now,
  };
}

function handleDeleteStudent(payload) {
  const studentId = String(payload.student_id || '').trim();
  if (!studentId) throw new Error('ID siswa tidak valid.');

  const sheet = getSheet(SHEETS.STUDENTS.name);
  const rows = readSheetObjects(SHEETS.STUDENTS.name);
  const existing = rows.find(function (r) {
    return r.student_id === studentId;
  });
  if (!existing) throw new Error('Data siswa tidak ditemukan.');

  // Soft delete agar riwayat absensi bulan sebelumnya tetap aman
  sheet.getRange(existing._row, 4).setNumberFormat('@').setValue('FALSE');
  return { deleted: true, student_id: studentId };
}

function handleGetMonthData(payload) {
  const monthKey = String(payload.month_key || '').trim();
  if (!/^\d{4}-\d{2}$/.test(monthKey)) {
    throw new Error('Format month_key harus YYYY-MM.');
  }

  const scheduleRows = readSheetObjects(SHEETS.SCHEDULES.name);
  const foundSchedule = scheduleRows.find(function (r) {
    return r.month_key === monthKey;
  });

  const schedule = foundSchedule
    ? {
        month_key: foundSchedule.month_key,
        dates: safeJsonParse(foundSchedule.dates_json, []),
        added_dates: safeJsonParse(foundSchedule.added_dates_json, []),
        removed_dates: safeJsonParse(foundSchedule.removed_dates_json, []),
      }
    : null;

  const attendanceRows = readSheetObjects(SHEETS.ATTENDANCE.name);
  const attendance = [];
  attendanceRows.forEach(function (r) {
    if (r.month_key === monthKey && r.status) {
      attendance.push({
        record_key: r.record_key,
        month_key: r.month_key,
        student_id: r.student_id,
        date: r.date,
        status: r.status,
      });
    }
  });

  return {
    month_key: monthKey,
    schedule: schedule,
    attendance: attendance,
  };
}

function handleSaveSchedule(payload) {
  const monthKey = String(payload.month_key || '').trim();
  if (!/^\d{4}-\d{2}$/.test(monthKey)) {
    throw new Error('Format month_key harus YYYY-MM.');
  }

  const dates = Array.isArray(payload.dates) ? payload.dates.slice().sort() : [];
  const addedDates = Array.isArray(payload.added_dates) ? payload.added_dates.slice().sort() : [];
  const removedDates = Array.isArray(payload.removed_dates) ? payload.removed_dates.slice().sort() : [];

  const sheet = getSheet(SHEETS.SCHEDULES.name);
  const rows = readSheetObjects(SHEETS.SCHEDULES.name);
  const existing = rows.find(function (r) {
    return r.month_key === monthKey;
  });

  const rowValues = [
    monthKey,
    JSON.stringify(dates),
    JSON.stringify(addedDates),
    JSON.stringify(removedDates),
  ];

  if (existing) {
    sheet.getRange(existing._row, 1, 1, 4).setNumberFormat('@').setValues([rowValues]);
  } else {
    appendTextRows(sheet, [rowValues]);
  }

  return {
    month_key: monthKey,
    dates: dates,
    added_dates: addedDates,
    removed_dates: removedDates,
  };
}

function handleBatchSaveAttendance(payload) {
  const monthKey = String(payload.month_key || '').trim();
  const changes = Array.isArray(payload.changes) ? payload.changes : [];
  if (!/^\d{4}-\d{2}$/.test(monthKey)) {
    throw new Error('Format month_key harus YYYY-MM.');
  }
  if (changes.length === 0) {
    return { saved: 0 };
  }

  // Ambil perubahan terakhir per sel jika terjadi ketukan berulang pada sel yang sama
  const deduped = {};
  changes.forEach(function (item) {
    const studentId = String(item.student_id || '').trim();
    const date = String(item.date || '').trim();
    const status = String(item.status || '').trim().toUpperCase();
    if (!studentId || !date) return;
    if (!VALID_STATUS[status]) {
      throw new Error('Status absensi tidak valid: ' + status);
    }
    const recordKey = monthKey + '|' + studentId + '|' + date;
    deduped[recordKey] = {
      record_key: recordKey,
      month_key: monthKey,
      student_id: studentId,
      date: date,
      status: status,
    };
  });

  const sheet = getSheet(SHEETS.ATTENDANCE.name);
  const existingRows = readSheetObjects(SHEETS.ATTENDANCE.name);
  const rowIndexByKey = {};
  existingRows.forEach(function (r) {
    rowIndexByKey[r.record_key] = r._row;
  });

  const now = new Date().toISOString();
  const toAppend = [];
  let savedCount = 0;

  Object.keys(deduped).forEach(function (recordKey) {
    const item = deduped[recordKey];
    const rowIdx = rowIndexByKey[recordKey];

    if (rowIdx) {
      sheet
        .getRange(rowIdx, 5, 1, 2)
        .setNumberFormat('@')
        .setValues([[item.status, now]]);
      savedCount++;
    } else if (item.status !== '') {
      toAppend.push([
        item.record_key,
        item.month_key,
        item.student_id,
        item.date,
        item.status,
        now,
      ]);
      savedCount++;
    }
  });

  if (toAppend.length > 0) {
    appendTextRows(sheet, toAppend);
  }

  return { saved: savedCount };
}

/* ===========================
 * UTILITY HELPERS
 * =========================== */

function getAppPin() {
  return String(PropertiesService.getScriptProperties().getProperty('APP_PIN') || '').trim();
}

function getSheet(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Tab "' + sheetName + '" tidak ditemukan. Jalankan fungsi setupDatabase() terlebih dahulu.');
  }
  return sheet;
}

function readSheetObjects(sheetName) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const values = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getDisplayValues();
  const headers = values[0];
  const results = [];

  for (let i = 1; i < values.length; i++) {
    const rowObj = { _row: i + 1 };
    for (let j = 0; j < headers.length; j++) {
      rowObj[headers[j]] = values[i][j];
    }
    results.push(rowObj);
  }
  return results;
}

function appendTextRows(sheet, rows) {
  if (!rows.length) return;
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setNumberFormat('@').setValues(rows);
}

function deleteMatchingRows(sheetName, predicate) {
  const sheet = getSheet(sheetName);
  const rows = readSheetObjects(sheetName);
  for (let i = rows.length - 1; i >= 0; i--) {
    if (predicate(rows[i])) {
      sheet.deleteRow(rows[i]._row);
    }
  }
}

function generateId(prefix) {
  return prefix + '_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 6);
}

function safeJsonParse(str, fallback) {
  if (!str) return fallback;
  try {
    return JSON.parse(str);
  } catch (e) {
    return fallback;
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
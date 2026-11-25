import { showToast } from './ui.js';

/** @type {any} */
let deferredInstallPrompt = null;

/**
 * Mengatur tombol ubah orientasi ke Mode Landscape (Fullscreen + Orientation Lock).
 */
function initLandscapeToggle() {
  const btn = document.getElementById('btn-landscape');
  const label = document.getElementById('btn-landscape-label');
  if (!btn) return;

  let isLandscapeActive = false;

  function syncButtonUi(active) {
    isLandscapeActive = active;
    btn.classList.toggle('btn-landscape-active', active);
    btn.title = active ? 'Keluar Mode Landscape' : 'Ubah ke Mode Landscape';
    btn.setAttribute('aria-label', btn.title);
    if (label) {
      label.textContent = active ? 'Normal' : 'Landscape';
    }
  }

  async function enterLandscape() {
    const docEl = document.documentElement;
    let enteredFullscreen = false;

    try {
      if (!document.fullscreenElement && docEl.requestFullscreen) {
        await docEl.requestFullscreen();
        enteredFullscreen = true;
      }
    } catch (err) {
      // Lanjutkan mencoba orientation lock meski fullscreen ditolak
    }

    try {
      if (screen.orientation && typeof screen.orientation.lock === 'function') {
        await screen.orientation.lock('landscape');
        syncButtonUi(true);
        return;
      }
    } catch (err) {
      if (enteredFullscreen) {
        syncButtonUi(true);
        return;
      }
      showToast('Browser tidak mendukung putar otomatis. Putar fisik HP Anda ke posisi mendatar.', 'info');
      return;
    }

    if (enteredFullscreen) {
      syncButtonUi(true);
    } else {
      showToast('Putar fisik HP Anda ke posisi mendatar untuk tampilan landscape.', 'info');
    }
  }

  async function exitLandscape() {
    try {
      if (screen.orientation && typeof screen.orientation.unlock === 'function') {
        screen.orientation.unlock();
      }
    } catch (err) {
      // Abaikan
    }

    try {
      if (document.fullscreenElement && document.exitFullscreen) {
        await document.exitFullscreen();
      }
    } catch (err) {
      // Abaikan
    }

    syncButtonUi(false);
  }

  btn.addEventListener('click', () => {
    if (isLandscapeActive) {
      exitLandscape();
    } else {
      enterLandscape();
    }
  });

  document.addEventListener('fullscreenchange', () => {
    if (!document.fullscreenElement && isLandscapeActive) {
      try {
        if (screen.orientation && typeof screen.orientation.unlock === 'function') {
          screen.orientation.unlock();
        }
      } catch (err) {
        // Abaikan
      }
      syncButtonUi(false);
    }
  });
}

/**
 * Mengaktifkan Service Worker, tombol Install PWA, dan tombol Mode Landscape.
 */
export function initPwaInstall() {
  initLandscapeToggle();

  const installBtn = document.getElementById('btn-install-app');

  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker.register('./sw.js').catch(() => {});
    });
  }

  const isStandalone =
    window.matchMedia('(display-mode: standalone)').matches ||
    window.navigator.standalone === true;

  if (isStandalone || !installBtn) return;

  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredInstallPrompt = e;
    installBtn.classList.remove('hidden');
  });

  installBtn.addEventListener('click', async () => {
    if (!deferredInstallPrompt) return;
    deferredInstallPrompt.prompt();
    const { outcome } = await deferredInstallPrompt.userChoice;
    deferredInstallPrompt = null;

    if (outcome === 'accepted') {
      installBtn.classList.add('hidden');
      showToast('Menginstal aplikasi ke perangkat Anda...', 'success');
    }
  });

  window.addEventListener('appinstalled', () => {
    deferredInstallPrompt = null;
    installBtn.classList.add('hidden');
    showToast('GPR Absensi berhasil diinstal.', 'success');
  });
}
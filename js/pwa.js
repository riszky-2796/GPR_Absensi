import { showToast } from './ui.js';

/** @type {any} */
let deferredInstallPrompt = null;

/**
 * Mengaktifkan Service Worker dan tombol "Install App" langsung di dalam website.
 */
export function initPwaInstall() {
  const installBtn = document.getElementById('btn-install-app');

  // 1. Daftarkan Service Worker
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker.register('./sw.js').catch(() => {});
    });
  }

  // Jangan tampilkan tombol jika sudah berjalan dalam mode aplikasi terinstal
  const isStandalone =
    window.matchMedia('(display-mode: standalone)').matches ||
    window.navigator.standalone === true;

  if (isStandalone || !installBtn) return;

  // 2. Tangkap event instalasi resmi dari Chrome/Edge Android & Desktop
  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredInstallPrompt = e;
    installBtn.classList.remove('hidden');
  });

  // 3. Jalankan dialog instalasi saat tombol diketuk
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
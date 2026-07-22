(() => {
  const installButtons = [document.getElementById('installBtn'), document.getElementById('installBtnHero')].filter(Boolean);
  const help = document.getElementById('installHelp');
  const toast = document.getElementById('toast');
  let deferredPrompt = null;

  const showToast = (text) => {
    toast.textContent = text;
    toast.classList.add('show');
    clearTimeout(showToast.timer);
    showToast.timer = setTimeout(() => toast.classList.remove('show'), 2600);
  };

  window.addEventListener('beforeinstallprompt', (event) => {
    event.preventDefault();
    deferredPrompt = event;
    installButtons.forEach(btn => btn.hidden = false);
    if (help) help.hidden = true;
  });

  installButtons.forEach(btn => btn.addEventListener('click', async () => {
    if (!deferredPrompt) {
      showToast('Chrome menyusidan “Ilovani o‘rnatish” ni tanlang.');
      return;
    }
    deferredPrompt.prompt();
    await deferredPrompt.userChoice;
    deferredPrompt = null;
    installButtons.forEach(b => b.hidden = true);
  }));

  document.querySelectorAll('.module').forEach(button => button.addEventListener('click', () => {
    showToast(`${button.dataset.module} bo‘limi keyingi modul bilan ulanadi.`);
  }));

  window.addEventListener('appinstalled', () => showToast('Mebel360° muvaffaqiyatli o‘rnatildi!'));

  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => navigator.serviceWorker.register('/service-worker.js?v=1.0.0'));
  }
})();

(() => {
  const $ = (selector) => document.querySelector(selector);
  const nativeFetch = window.fetch.bind(window);
  const isAdmin = Boolean(document.querySelector('#dataOpsHost'));
  let toastTimer = null;
  let authRedirectPending = false;

  function toast(message, ok = true) {
    const el = $('#agingToast');
    if (!el) return;
    clearTimeout(toastTimer);
    el.textContent = message;
    el.className = `fixed bottom-4 right-4 z-[120] max-w-sm rounded-xl border px-4 py-3 text-sm shadow-2xl ${
      ok
        ? 'border-emerald-600 bg-emerald-950 text-emerald-200'
        : 'border-red-600 bg-red-950 text-red-200'
    }`;
    el.classList.remove('hidden');
    toastTimer = setTimeout(() => el.classList.add('hidden'), 3500);
  }

  window.showAgingToast = (message, tone = 'success') => toast(message, tone !== 'error' && tone !== false);

  function adminLoginUrl() {
    const next = window.location.pathname + window.location.search + window.location.hash;
    return `/login?next=${encodeURIComponent(next)}`;
  }

  function scheduleAdminRelogin(message = 'Your Admin session expired. Please sign in again.') {
    if (authRedirectPending) return;
    authRedirectPending = true;
    toast(message, false);
    setTimeout(() => window.location.assign(adminLoginUrl()), 700);
  }

  async function appFetch(url, options = {}) {
    const response = await nativeFetch(url, { credentials: 'same-origin', cache: 'no-store', ...options });
    if (isAdmin && String(url).startsWith('/admin/') && (response.status === 401 || response.status === 403)) {
      let data = {};
      try { data = await response.clone().json(); } catch (_) {}
      scheduleAdminRelogin(data.error || 'Admin access needs to be refreshed.');
    }
    return response;
  }

  async function verifyAdminSession() {
    if (!isAdmin) return;
    try {
      const response = await nativeFetch('/api/session', { credentials: 'same-origin', cache: 'no-store' });
      const data = await response.json();
      if (data.configuration_required) {
        toast('Vercel Admin session setup is incomplete. Configure SCM_SECRET_KEY and redeploy.', false);
        return;
      }
      if (!response.ok || !data.is_admin) scheduleAdminRelogin('Admin access needs to be refreshed. Please sign in again.');
    } catch (_) {}
  }

  function responseFilename(response, fallbackName) {
    const disposition = response.headers.get('content-disposition') || '';
    const utf = disposition.match(/filename\*=UTF-8''([^;]+)/i);
    const plain = disposition.match(/filename="?([^";]+)"?/i);
    const raw = utf?.[1] || plain?.[1];
    if (!raw) return fallbackName;
    try {
      return decodeURIComponent(raw.replaceAll('"', '').trim());
    } catch (_) {
      return raw.replaceAll('"', '').trim() || fallbackName;
    }
  }

  async function downloadResponse(response, fallbackName) {
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = responseFilename(response, fallbackName);
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function applySidebarPreference() {
    const collapsed = localStorage.getItem('scm-sidebar-collapsed') === '1';
    document.body.classList.toggle('sidebar-collapsed', collapsed);
    document.body.classList.toggle('sidebar-expanded', !collapsed);
    const button = $('#sidebarToggle');
    if (button) button.title = collapsed ? 'Expand sidebar' : 'Collapse sidebar';
  }

  function toggleSidebar() {
    if (window.innerWidth < 1024) {
      document.body.classList.toggle('sidebar-mobile-open');
      return;
    }
    const collapsed = !document.body.classList.contains('sidebar-collapsed');
    localStorage.setItem('scm-sidebar-collapsed', collapsed ? '1' : '0');
    applySidebarPreference();
    setTimeout(() => window.dispatchEvent(new Event('resize')), 230);
  }

  function syncThemeButton() {
    const light = document.documentElement.dataset.theme === 'light';
    const label = $('#themeLabel');
    if (label) label.textContent = light ? 'Dark Mode' : 'Light Mode';
    const icon = $('#themeIcon');
    if (!icon) return;
    icon.innerHTML = light
      ? '<svg viewBox="0 0 24 24"><path d="M21 12.8A8.5 8.5 0 1 1 11.2 3 7 7 0 0 0 21 12.8Z"/></svg>'
      : '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M4.93 19.07l1.41-1.41M17.66 6.34l1.41-1.41"/></svg>';
  }

  function setTheme(theme) {
    document.documentElement.dataset.theme = theme;
    localStorage.setItem('scm-theme', theme);
    syncThemeButton();
    window.dispatchEvent(new CustomEvent('scm-theme-change', { detail: { theme } }));
  }

  const clearModal = $('#clearDataModal');
  function openClearData() {
    if (!clearModal) return;
    clearModal.classList.remove('hidden');
    clearModal.classList.add('flex');
    const input = $('#clearDataConfirmation');
    if (input) {
      input.value = '';
      setTimeout(() => input.focus(), 80);
    }
  }

  function closeClearData() {
    if (!clearModal) return;
    clearModal.classList.add('hidden');
    clearModal.classList.remove('flex');
  }

  applySidebarPreference();
  syncThemeButton();
  verifyAdminSession();

  $('#sidebarToggle')?.addEventListener('click', toggleSidebar);
  $('#mobileSidebarBtn')?.addEventListener('click', () => document.body.classList.add('sidebar-mobile-open'));
  $('#sidebarBackdrop')?.addEventListener('click', () => document.body.classList.remove('sidebar-mobile-open'));
  window.addEventListener('resize', () => {
    if (window.innerWidth >= 1024) document.body.classList.remove('sidebar-mobile-open');
  });
  $('#themeBtn')?.addEventListener('click', () =>
    setTheme(document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark')
  );

  $('#pptBtn')?.addEventListener('click', async () => {
    const button = $('#pptBtn');
    const label = button?.querySelector('.sidebar-label');
    const oldText = label?.textContent || 'Export Deck';
    if (button) button.disabled = true;
    if (label) label.textContent = 'Generating Deck…';
    toast('Generating the executive PowerPoint deck…');
    try {
      const response = await appFetch('/admin/export/pptx');
      if (!response.ok) {
        let data = {};
        try { data = await response.json(); } catch (_) {}
        toast(data.error || 'PowerPoint export failed.', false);
        return;
      }
      await downloadResponse(response, 'SCM_Executive_Review.pptx');
      toast('PowerPoint deck generated successfully.');
    } catch (error) {
      toast(`PowerPoint export failed. ${error?.message || ''}`.trim(), false);
    } finally {
      if (button) button.disabled = false;
      if (label) label.textContent = oldText;
    }
  });

  $('#clearAllDataBtn')?.addEventListener('click', openClearData);
  $('#cancelClearDataBtn')?.addEventListener('click', closeClearData);
  clearModal?.addEventListener('click', (event) => {
    if (event.target === clearModal) closeClearData();
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && clearModal?.classList.contains('flex')) closeClearData();
  });

  $('#confirmClearDataBtn')?.addEventListener('click', async () => {
    const confirmation = String($('#clearDataConfirmation')?.value || '').trim();
    if (confirmation !== 'CLEAR DATA') {
      toast('Type CLEAR DATA exactly to confirm.', false);
      return;
    }
    const button = $('#confirmClearDataBtn');
    if (button) {
      button.disabled = true;
      button.textContent = 'Clearing…';
    }
    try {
      const response = await appFetch('/admin/clear-data', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ confirmation }),
      });
      let data = {};
      try { data = await response.json(); } catch (_) {}
      if (!response.ok) {
        toast(data.error || 'Clear Data failed.', false);
        return;
      }
      toast(data.message || 'All user data cleared.');
      setTimeout(() => { window.location.href = '/'; }, 500);
    } catch (error) {
      toast(`Clear Data failed. ${error?.message || ''}`.trim(), false);
    } finally {
      if (button) {
        button.disabled = false;
        button.textContent = 'Clear All Data';
      }
    }
  });
})();

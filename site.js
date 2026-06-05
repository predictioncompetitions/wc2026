// ── CONFIG ────────────────────────────────────────────────────────────────────
var SCRIPT_URL = 'https://api.micksworldcup2026.com';

// ── NAV ───────────────────────────────────────────────────────────────────────
async function initNav() {
  const toggle = document.getElementById('navToggle');
  const mobile = document.getElementById('navMobile');
  if (toggle && mobile) {
    toggle.addEventListener('click', () => {
      mobile.classList.toggle('open');
      toggle.textContent = mobile.classList.contains('open') ? '✕' : '☰';
    });
    mobile.querySelectorAll('a').forEach(a => {
      a.addEventListener('click', () => {
        mobile.classList.remove('open');
        toggle.textContent = '☰';
      });
    });
  }
  const path = window.location.pathname.split('/').pop() || 'index.html';
  document.querySelectorAll('.nav-links a, .nav-mobile a').forEach(a => {
    const href = a.getAttribute('href');
    if (href === path || (path === '' && href === 'index.html')) {
      a.classList.add('active');
    }
  });
  try {
    const data = await apiFetch('getInitialData');
    const isLive = data.tournamentStarted === true;
    const entriesOpen = String(data.entryStatus || '').toUpperCase() === 'OPEN';
    if (!entriesOpen) {
      document.querySelectorAll('[data-phase="pre"]').forEach(el => el.style.display = 'none');
    }
    if (!isLive) {
      document.querySelectorAll('[data-phase="live"]').forEach(el => el.style.display = 'none');
    } else {
      document.querySelectorAll('[data-phase="live"]').forEach(el => el.style.display = 'block');
    }
  } catch(e) {
    document.querySelectorAll('[data-phase="live"]').forEach(el => el.style.display = 'none');
  }
}

async function apiFetch(action, params = {}, method = 'GET') {
  if (method === 'POST') {
    const res = await fetch(SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action, payload: params })
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return res.json();
  }
  const url = new URL(SCRIPT_URL);
  url.searchParams.set('action', action);
  Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));
  const res = await fetch(url.toString());
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return res.json();
}

// ── FORMAT HELPERS ────────────────────────────────────────────────────────────
function escSite(str) {
  return String(str || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function pct(n, total) {
  if (!total) return '0%';
  return Math.round((n / total) * 100) + '%';
}

// ── INIT ──────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', initNav);

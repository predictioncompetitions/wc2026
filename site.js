// ── CONFIG ────────────────────────────────────────────
// Replace this with your Apps Script Web App URL after deploying
const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzThbgrqjQytw8ak4PY3GcnxhF-JdnF1CcossbeB4NaXbKKIPfryU1oOm5Y6r9HSw/exec';

// ── NAV ───────────────────────────────────────────────
function initNav() {
  const toggle = document.getElementById('navToggle');
  const mobile = document.getElementById('navMobile');
  if (!toggle || !mobile) return;

  toggle.addEventListener('click', () => {
    mobile.classList.toggle('open');
    toggle.textContent = mobile.classList.contains('open') ? '✕' : '☰';
  });

  // Close mobile nav when a link is clicked
  mobile.querySelectorAll('a').forEach(a => {
    a.addEventListener('click', () => {
      mobile.classList.remove('open');
      toggle.textContent = '☰';
    });
  });

  // Mark current page as active
  const path = window.location.pathname.split('/').pop() || 'index.html';
  document.querySelectorAll('.nav-links a, .nav-mobile a').forEach(a => {
    const href = a.getAttribute('href');
    if (href === path || (path === '' && href === 'index.html')) {
      a.classList.add('active');
    }
  });
}

// ── API ───────────────────────────────────────────────
async function apiFetch(action, params = {}) {
  const url = new URL(SCRIPT_URL);
  url.searchParams.set('action', action);
  Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));

  const res = await fetch(url.toString());
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return res.json();
}

// ── FORMAT HELPERS ────────────────────────────────────
function esc(str) {
  return String(str || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function pct(n, total) {
  if (!total) return '0%';
  return Math.round((n / total) * 100) + '%';
}

// ── INIT ──────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', initNav);

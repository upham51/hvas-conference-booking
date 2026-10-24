// Shared API helper: talks to the Apps Script Web App.
// GET requests are plain fetches. POST requests use a text/plain body
// (Apps Script quirk) so the browser never sends a CORS preflight.
//
// Under load Google sometimes answers with an HTML error page instead of our
// JSON, even when the request went through. So every call is tried twice.
// POSTs carry a requestId that stays the same on the retry; the backend
// remembers the first answer for that id, so a retried booking or cancel
// never happens twice.

function hvaRequestId() {
  if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
  return 'r' + Date.now().toString(36) + Math.random().toString(36).slice(2, 12);
}

// Escapes text before it goes into innerHTML, so a name typed into the form
// can never run as code on the teacher or principal pages.
function esc(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

const HVA_API = {
  async _fetchJSON(url, init) {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), 90000);
    try {
      const res = await fetch(url, { ...init, signal: ctrl.signal });
      const text = await res.text();
      return JSON.parse(text); // throws on an HTML error page
    } finally {
      clearTimeout(timer);
    }
  },
  async _twice(fn) {
    try {
      return await fn();
    } catch (err) {
      await new Promise((r) => setTimeout(r, 1500));
      return fn();
    }
  },
  async get(params) {
    const base = window.HVA_CONFIG.API_BASE_URL;
    const url = base + '?' + new URLSearchParams(params).toString();
    return this._twice(() => this._fetchJSON(url, { method: 'GET' }));
  },
  async post(payload) {
    const base = window.HVA_CONFIG.API_BASE_URL;
    const body = JSON.stringify({ requestId: hvaRequestId(), ...payload });
    return this._twice(() => this._fetchJSON(base, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body,
    }));
  },
};

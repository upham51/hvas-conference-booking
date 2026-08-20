// Shared API helper: talks to the Apps Script Web App.
// GET requests are plain fetches. POST requests use a text/plain body
// (Apps Script quirk) so the browser never sends a CORS preflight.
const HVA_API = {
  async get(params) {
    const base = window.HVA_CONFIG.API_BASE_URL;
    const url = base + '?' + new URLSearchParams(params).toString();
    const res = await fetch(url, { method: 'GET' });
    return res.json();
  },
  async post(payload) {
    const base = window.HVA_CONFIG.API_BASE_URL;
    const res = await fetch(base, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload),
    });
    return res.json();
  },
};

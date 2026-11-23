const CFG = window.APP_CONFIG || {};

function isConfigured() {
  return /^https:\/\/script\.google\.com\/.+\/exec$/.test(CFG.SCRIPT_URL || '') && !/여기에/.test(CFG.SCRIPT_URL);
}

async function apiGet(params) {
  const url = CFG.SCRIPT_URL + '?' + new URLSearchParams(params).toString();
  const res = await fetch(url);
  return parse(res);
}

// Content-Type을 text/plain으로 보내야 브라우저가 사전요청(CORS preflight)을 하지 않습니다.
async function apiPost(body) {
  const res = await fetch(CFG.SCRIPT_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(body),
  });
  return parse(res);
}

async function parse(res) {
  let data;
  try {
    data = await res.json();
  } catch (e) {
    throw new Error('서버 응답을 읽지 못했습니다. 웹 앱 배포 설정(액세스: 모든 사용자)을 확인하세요.');
  }
  if (!data.ok) throw new Error(data.error || '알 수 없는 오류');
  return data;
}

function $(sel, root) { return (root || document).querySelector(sel); }

function esc(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function fmtDate(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d)) return '';
  return d.toLocaleString('ko-KR', { month: 'long', day: 'numeric', weekday: 'short', hour: '2-digit', minute: '2-digit' });
}

function fmtSize(bytes) {
  if (bytes < 1024) return bytes + 'B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(0) + 'KB';
  return (bytes / 1024 / 1024).toFixed(1) + 'MB';
}

function toast(msg, type) {
  let el = $('#toast');
  if (!el) {
    el = document.createElement('div');
    el.id = 'toast';
    document.body.appendChild(el);
  }
  el.textContent = msg;
  el.className = 'show ' + (type || '');
  clearTimeout(el._t);
  el._t = setTimeout(() => (el.className = ''), 3200);
}

document.addEventListener('DOMContentLoaded', () => {
  document.querySelectorAll('[data-title]').forEach(el => (el.textContent = CFG.TITLE || el.textContent));
  if (CFG.TITLE) document.title = CFG.TITLE + (document.body.dataset.page === 'admin' ? ' · 관리' : '');
});

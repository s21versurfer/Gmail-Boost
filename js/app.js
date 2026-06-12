/* ── State ── */
let cols = [], rows = [], previews = [], cur = 0, imgs = {};
let savedRange = null;
let editor = null;  // DOM 준비 후 바인딩

/* ── DOMContentLoaded ── */
document.addEventListener('DOMContentLoaded', () => {
  editor = document.getElementById('body-editor');
  editor.addEventListener('blur', saveRange);
  editor.addEventListener('keyup', () => { updateTags(); saveRange(); });
  editor.addEventListener('mouseup', saveRange);
  editor.addEventListener('input', updateTags);
  updateTags();
});

/* ── Tab switching ── */
function sw(name) {
  const names = ['tpl', 'img', 'data', 'map'];
  names.forEach((t, i) => {
    document.getElementById('t-' + t).classList.add('hidden');
    document.querySelectorAll('.nav-tabs .tab')[i].classList.remove('active');
  });
  document.getElementById('t-' + name).classList.remove('hidden');
  document.querySelectorAll('.nav-tabs .tab')[names.indexOf(name)].classList.add('active');
  if (name === 'map') updateMapping();
}

/* ── Editor: range save/restore ── */
function saveRange() {
  const sel = window.getSelection();
  if (sel && sel.rangeCount) savedRange = sel.getRangeAt(0).cloneRange();
}
function restoreRange() {
  if (!editor) return;
  editor.focus();
  if (savedRange) {
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(savedRange);
  }
}

/* ── Toolbar commands ── */
function fmt(cmd) { if (!editor) return; editor.focus(); document.execCommand(cmd, false, null); updateTags(); }

function setFontSize(v) {
  if (!editor) return;
  editor.focus();
  document.execCommand('fontSize', false, '7');
  document.querySelectorAll('#body-editor font[size="7"]').forEach(el => {
    el.removeAttribute('size');
    el.style.fontSize = v;
  });
}

function setColor(cmd, val) { if (!editor) return; editor.focus(); document.execCommand(cmd, false, val); }
function clearFmt() { if (!editor) return; editor.focus(); document.execCommand('removeFormat', false, null); }
function insertList() { if (!editor) return; editor.focus(); document.execCommand('insertUnorderedList', false, null); }

function insertHR() {
  if (!editor) return;
  editor.focus();
  document.execCommand('insertHTML', false, '<hr style="border:none;border-top:1px solid #e0e0e0;margin:12px 0">');
}

function insertTable() {
  if (!editor) return;
  const html = `<table style="width:100%;border-collapse:collapse;font-size:13px;margin:8px 0">
  <tr>
    <th style="border:1px solid #ddd;padding:6px 10px;background:#f5f5f5;text-align:left">항목</th>
    <th style="border:1px solid #ddd;padding:6px 10px;background:#f5f5f5;text-align:left">내용</th>
  </tr>
  <tr>
    <td style="border:1px solid #ddd;padding:6px 10px">예시 1</td>
    <td style="border:1px solid #ddd;padding:6px 10px">내용 1</td>
  </tr>
  <tr>
    <td style="border:1px solid #ddd;padding:6px 10px">예시 2</td>
    <td style="border:1px solid #ddd;padding:6px 10px">내용 2</td>
  </tr>
</table>`;
  editor.focus();
  document.execCommand('insertHTML', false, html);
}

/* ── Variable detection ── */
function getVars() {
  const subjTxt = document.getElementById('subj').value || '';
  // innerText로 읽으면 HTML 엔티티 변환 없이 순수 텍스트로 {{변수}} 감지 가능
  const bodyTxt = editor ? (editor.innerText || '') : '';
  const combined = subjTxt + bodyTxt;
  const m = [...combined.matchAll(/\{\{([^{}]+)\}\}/g)];
  return [...new Set(m.map(x => x[1].trim()).filter(Boolean))];
}

function updateTags() {
  const v = getVars();
  const row = document.getElementById('tag-row');
  if (!row) return;
  row.innerHTML = v.length
    ? v.map(x => `<button class="tag" onclick="insertVar('${x.replace(/'/g, "\\'")}')">\{\{${x}\}\}</button>`).join('')
    : '<span class="empty-note">변수 없음</span>';
}

/* 에디터에 {{변수}} 순수 텍스트로 삽입 */
function insertVar(varName) {
  restoreRange();
  document.execCommand('insertText', false, '{{' + varName + '}}');
  saveRange();
  updateTags();
}

/* ── Images ── */
function loadImages(input) {
  Array.from(input.files).forEach(f => {
    const r = new FileReader();
    r.onload = e => { imgs[f.name] = e.target.result; renderImgGrid(); };
    r.readAsDataURL(f);
  });
}

function renderImgGrid() {
  const g = document.getElementById('img-grid');
  const keys = Object.keys(imgs);
  if (!keys.length) { g.innerHTML = '<p class="empty-note">이미지 없음</p>'; return; }
  g.innerHTML = keys.map(k => `
    <div class="img-thumb" onclick="insertImg('${k.replace(/'/g, "\\'")}')">
      <img src="${imgs[k]}" alt="${k}">
      <div class="img-ins">삽입</div>
      <div class="img-name">${k}</div>
    </div>`).join('');
}

function insertImg(name) {
  sw('tpl');
  setTimeout(() => {
    restoreRange();
    document.execCommand('insertHTML', false,
      `<img src="${imgs[name]}" alt="${name}" style="max-width:100%;border-radius:4px;display:block;margin:8px 0">`
    );
    saveRange();
  }, 60);
}

function insertImgUrl() {
  const url = document.getElementById('img-url').value.trim();
  if (!url) return;
  sw('tpl');
  setTimeout(() => {
    restoreRange();
    document.execCommand('insertHTML', false,
      `<img src="${url}" alt="image" style="max-width:100%;border-radius:4px;display:block;margin:8px 0">`
    );
    document.getElementById('img-url').value = '';
    saveRange();
  }, 60);
}

/* ── File loading ── */
function loadFile(input) {
  const f = input.files[0];
  if (!f) return;
  document.getElementById('file-name').textContent = f.name;
  const r = new FileReader();
  r.onload = e => {
    try {
      if (f.name.toLowerCase().endsWith('.csv')) {
        parseCSV(e.target.result);
      } else {
        const wb = XLSX.read(e.target.result, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
        cols = data[0] ? data[0].map(String) : [];
        rows = data.slice(1).filter(row => row.some(c => c !== undefined && c !== ''));
      }
      showCols();
      updateEmailSel();
      updateMapping();
    } catch (err) {
      document.getElementById('file-name').textContent = '오류: ' + err.message;
    }
  };
  f.name.toLowerCase().endsWith('.csv') ? r.readAsText(f, 'UTF-8') : r.readAsBinaryString(f);
}

function parseCSV(txt) {
  const lines = txt.trim().split(/\r?\n/);
  cols = lines[0].split(',').map(c => c.replace(/^"|"$/g, '').trim());
  rows = lines.slice(1).map(line => {
    const cells = [];
    let cur2 = '', inQ = false;
    for (let ch of line) {
      if (ch === '"') { inQ = !inQ; }
      else if (ch === ',' && !inQ) { cells.push(cur2.trim()); cur2 = ''; }
      else cur2 += ch;
    }
    cells.push(cur2.trim());
    return cells;
  }).filter(r => r.some(c => c !== ''));
}

/* ── 열 문자 변환 (0→A, 25→Z, 26→AA ...) ── */
function colLetter(i) {
  let s = '', n = i + 1;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function showCols() {
  document.getElementById('col-area').classList.remove('hidden');
  document.getElementById('col-tags').innerHTML =
    cols.map((c, i) => `<span class="tag tag-static col-chip"><span class="col-letter">${colLetter(i)}</span>${c}</span>`).join('');
  document.getElementById('data-stat').textContent = `${cols.length}개 열 · ${rows.length}개 행 로드됨`;
}

function updateEmailSel() {
  const s = document.getElementById('email-col');
  s.innerHTML = '<option value="">— 선택 —</option>' +
    cols.map((c, i) => `<option value="${i}">${colLetter(i)}열 — ${c}</option>`).join('');
}

/* ── Mapping ── */
function updateMapping() {
  const vars = getVars();
  const area = document.getElementById('mapping-area');
  if (!area) return;

  if (!vars.length && !cols.length) {
    area.innerHTML = '<p class="empty-note">템플릿과 데이터를 먼저 입력하세요</p>';
    return;
  }
  if (!vars.length) {
    area.innerHTML = '<p class="empty-note">템플릿에 {{변수}}가 없습니다</p>';
    return;
  }
  if (!cols.length) {
    area.innerHTML = '<p class="empty-note">데이터 파일을 먼저 업로드하세요</p>';
    return;
  }

  area.innerHTML = vars.map(v => {
    // 변수명과 헤더명이 같으면 자동 매칭
    const ai = cols.findIndex(c => c.trim().toLowerCase() === v.trim().toLowerCase());
    const opts = cols.map((c, i) =>
      `<option value="${i}"${i === ai ? ' selected' : ''}>${colLetter(i)}열 — ${c}</option>`
    ).join('');
    return `<div class="mapping-row">
      <span class="m-tag">{{${v}}}</span>
      <span class="m-arrow">→</span>
      <select class="m-sel" data-var="${v}">
        <option value="">— 선택 안 함 —</option>
        ${opts}
      </select>
    </div>`;
  }).join('');
}

/* ── Template apply ── */
let currentRow = [];

function applyTmpl(tmpl, vm) {
  // innerHTML 기준이므로 innerText {{변수}} 가 텍스트노드로 그대로 존재함
  // 단, HTML 엔티티로 인코딩된 경우도 처리
  const decoded = tmpl
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
  return decoded.replace(/\{\{([^{}]+)\}\}/g, (match, k) => {
    k = k.trim();
    const ci = vm[k];
    if (ci !== undefined && ci !== '') {
      const val = currentRow[parseInt(ci)];
      return val !== undefined ? String(val) : match;
    }
    return match;
  });
}

/* ── Generate previews ── */
function genPreviews() {
  const subj = document.getElementById('subj').value || '';
  const bodyHTML = editor ? editor.innerHTML : '';
  const emailColVal = document.getElementById('email-col').value;
  const emailColIdx = emailColVal !== '' ? parseInt(emailColVal) : -1;
  const rs = parseInt(document.getElementById('row-s').value) - 2;
  const re = parseInt(document.getElementById('row-e').value);

  const vm = {};
  document.querySelectorAll('[data-var]').forEach(s => {
    if (s.value !== '') vm[s.dataset.var] = s.value;
  });

  let wr = rows.slice(Math.max(0, rs));
  if (re > 0) {
    const startRow = parseInt(document.getElementById('row-s').value);
    wr = wr.slice(0, re - startRow + 1);
  }

  if (!wr.length) { showMapStatus('데이터 행이 없습니다. 파일과 행 범위를 확인하세요.', 'warn'); return; }

  previews = wr.map((row, i) => {
    currentRow = row;
    const email = emailColIdx >= 0 ? String(row[emailColIdx] || '') : '';
    return {
      i: i + 1,
      to: email,
      subj: applyTmpl(subj, vm),
      body: applyTmpl(bodyHTML, vm),
      valid: /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)
    };
  });
  currentRow = [];

  cur = 0;
  renderPrev();
  document.getElementById('send-bar').style.display = 'flex';
  document.getElementById('send-info').textContent =
    `총 ${previews.length}개 (유효: ${previews.filter(p => p.valid).length}개)`;
  showMapStatus(`${previews.length}개 미리보기 생성 완료`, 'ok');
}

function showMapStatus(msg, type) {
  document.getElementById('map-status').innerHTML =
    `<div class="status s-${type}">${msg}</div>`;
}

/* ── Preview rendering ── */
function renderPrev() {
  if (!previews.length) return;
  const p = previews[cur];
  document.getElementById('prev-count').textContent = `${cur + 1} / ${previews.length}`;
  const badge = document.getElementById('cur-badge');
  badge.style.display = 'inline-block';
  badge.className = 'badge ' + (p.valid ? 'badge-ok' : 'badge-warn');
  badge.textContent = p.valid ? '유효' : '주소 없음';
  document.getElementById('prev-area').innerHTML = `
    <div class="email-frame">
      <div class="email-meta">
        <div class="email-meta-row"><span class="meta-label">받는 이</span><span class="meta-val">${p.to || '(없음)'}</span></div>
        <div class="email-meta-row"><span class="meta-label">제목</span><span class="meta-val">${p.subj || '(없음)'}</span></div>
      </div>
      <div class="email-body-preview">${p.body}</div>
    </div>`;
}

function nav(d) {
  if (!previews.length) return;
  cur = (cur + d + previews.length) % previews.length;
  renderPrev();
}

/* ── Send / Export ── */
function sendGmail() {
  if (!previews.length) return;
  const p = previews[cur];
  if (!p.valid) { showMapStatus('현재 항목의 이메일 주소가 없습니다.', 'warn'); return; }
  const plain = p.body.replace(/<br\s*\/?>/gi, '\n').replace(/<\/p>/gi, '\n').replace(/<[^>]+>/g, '');
  const url = `https://mail.google.com/mail/?view=cm&fs=1&to=${encodeURIComponent(p.to)}&su=${encodeURIComponent(p.subj)}&body=${encodeURIComponent(plain)}`;
  window.open(url, '_blank');
}

function exportHTML() {
  if (!previews.length) return;
  const p = previews[cur];
  const html = `<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>${p.subj}</title>
  <style>
    body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;max-width:620px;margin:40px auto;font-size:14px;line-height:1.75;color:#1a1a18}
    .meta{color:#888;font-size:12px;margin-bottom:16px;padding-bottom:12px;border-bottom:1px solid #eee}
  </style>
</head>
<body>
  <div class="meta">To: ${p.to}<br>Subject: ${p.subj}</div>
  ${p.body}
</body>
</html>`;
  const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `email_${String(cur + 1).padStart(3, '0')}.html`;
  a.click();
  URL.revokeObjectURL(url);
}

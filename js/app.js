/* ── State ── */
let cols = [], rows = [], previews = [], cur = 0, imgs = {};
let savedRange = null;

/* editor를 전역변수로 캐싱하지 않고, 매번 직접 조회 */
function ed() { return document.getElementById('body-editor'); }

/* ── 초기화: script가 </body> 직전에 있으므로 DOM은 이미 준비됨 ── */
(function init() {
  const el = ed();
  if (!el) return;
  el.addEventListener('blur', saveRange);
  el.addEventListener('keyup', function() { updateTags(); saveRange(); });
  el.addEventListener('mouseup', saveRange);
  el.addEventListener('input', updateTags);
  updateTags();
})();

/* ── Tab switching ── */
function sw(name) {
  ['tpl','img','data','map'].forEach(function(t, i) {
    document.getElementById('t-' + t).classList.add('hidden');
    document.querySelectorAll('.nav-tabs .tab')[i].classList.remove('active');
  });
  document.getElementById('t-' + name).classList.remove('hidden');
  document.querySelectorAll('.nav-tabs .tab')[['tpl','img','data','map'].indexOf(name)].classList.add('active');
  if (name === 'map') updateMapping();
}

/* ── Editor range ── */
function saveRange() {
  var sel = window.getSelection();
  if (sel && sel.rangeCount) savedRange = sel.getRangeAt(0).cloneRange();
}
function restoreRange() {
  var el = ed();
  if (!el) return;
  el.focus();
  if (savedRange) {
    var sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(savedRange);
  }
}

/* ── Toolbar ── */
function fmt(cmd) { ed().focus(); document.execCommand(cmd, false, null); updateTags(); }

function setFontSize(v) {
  ed().focus();
  document.execCommand('fontSize', false, '7');
  document.querySelectorAll('#body-editor font[size="7"]').forEach(function(el) {
    el.removeAttribute('size'); el.style.fontSize = v;
  });
}

function setColor(cmd, val) { ed().focus(); document.execCommand(cmd, false, val); }
function clearFmt() { ed().focus(); document.execCommand('removeFormat', false, null); }
function insertList() { ed().focus(); document.execCommand('insertUnorderedList', false, null); }

function insertHR() {
  ed().focus();
  document.execCommand('insertHTML', false, '<hr style="border:none;border-top:1px solid #e0e0e0;margin:12px 0">');
}

function insertTable() {
  ed().focus();
  document.execCommand('insertHTML', false,
    '<table style="width:100%;border-collapse:collapse;font-size:13px;margin:8px 0">' +
    '<tr><th style="border:1px solid #ddd;padding:6px 10px;background:#f5f5f5;text-align:left">항목</th>' +
    '<th style="border:1px solid #ddd;padding:6px 10px;background:#f5f5f5;text-align:left">내용</th></tr>' +
    '<tr><td style="border:1px solid #ddd;padding:6px 10px">예시 1</td><td style="border:1px solid #ddd;padding:6px 10px">내용 1</td></tr>' +
    '<tr><td style="border:1px solid #ddd;padding:6px 10px">예시 2</td><td style="border:1px solid #ddd;padding:6px 10px">내용 2</td></tr>' +
    '</table>');
}

/* ── Variable detection: innerText로 순수 텍스트에서 {{변수}} 읽기 ── */
function getVars() {
  var subjTxt = document.getElementById('subj').value || '';
  var edEl = ed();
  var bodyTxt = edEl ? (edEl.innerText || edEl.textContent || '') : '';
  var combined = subjTxt + bodyTxt;
  var matches = [];
  var re = /\{\{([^{}]+)\}\}/g;
  var m;
  while ((m = re.exec(combined)) !== null) {
    var v = m[1].trim();
    if (v && matches.indexOf(v) === -1) matches.push(v);
  }
  return matches;
}

function updateTags() {
  var v = getVars();
  var row = document.getElementById('tag-row');
  if (!row) return;
  if (v.length) {
    row.innerHTML = v.map(function(x) {
      return '<button class="tag" onclick="insertVar(\'' + x.replace(/\\/g,'\\\\').replace(/'/g,"\\'") + '\')">{{' + x + '}}</button>';
    }).join('');
  } else {
    row.innerHTML = '<span class="empty-note">변수 없음</span>';
  }
}

function insertVar(varName) {
  restoreRange();
  document.execCommand('insertText', false, '{{' + varName + '}}');
  saveRange();
  updateTags();
}

/* ── Images ── */
function loadImages(input) {
  Array.from(input.files).forEach(function(f) {
    var r = new FileReader();
    r.onload = function(e) { imgs[f.name] = e.target.result; renderImgGrid(); };
    r.readAsDataURL(f);
  });
}

function renderImgGrid() {
  var g = document.getElementById('img-grid');
  var keys = Object.keys(imgs);
  if (!keys.length) { g.innerHTML = '<p class="empty-note">이미지 없음</p>'; return; }
  g.innerHTML = keys.map(function(k) {
    return '<div class="img-thumb" onclick="insertImg(\'' + k.replace(/\\/g,'\\\\').replace(/'/g,"\\'") + '\')">' +
      '<img src="' + imgs[k] + '" alt="' + k + '">' +
      '<div class="img-ins">삽입</div>' +
      '<div class="img-name">' + k + '</div></div>';
  }).join('');
}

function insertImg(name) {
  sw('tpl');
  setTimeout(function() {
    restoreRange();
    document.execCommand('insertHTML', false,
      '<img src="' + imgs[name] + '" alt="' + name + '" style="max-width:100%;border-radius:4px;display:block;margin:8px 0">');
    saveRange();
  }, 60);
}

function insertImgUrl() {
  var url = document.getElementById('img-url').value.trim();
  if (!url) return;
  sw('tpl');
  setTimeout(function() {
    restoreRange();
    document.execCommand('insertHTML', false,
      '<img src="' + url + '" alt="image" style="max-width:100%;border-radius:4px;display:block;margin:8px 0">');
    document.getElementById('img-url').value = '';
    saveRange();
  }, 60);
}

/* ── File loading ── */
function loadFile(input) {
  var f = input.files[0];
  if (!f) return;
  document.getElementById('file-name').textContent = f.name;
  var r = new FileReader();
  r.onload = function(e) {
    try {
      if (f.name.toLowerCase().endsWith('.csv')) {
        parseCSV(e.target.result);
      } else {
        var wb = XLSX.read(e.target.result, { type: 'binary' });
        var ws = wb.Sheets[wb.SheetNames[0]];
        var data = XLSX.utils.sheet_to_json(ws, { header: 1 });
        cols = data[0] ? data[0].map(String) : [];
        rows = data.slice(1).filter(function(row) {
          return row.some(function(c) { return c !== undefined && c !== ''; });
        });
      }
      showCols();
      updateEmailSel();
      updateMapping();
    } catch(err) {
      document.getElementById('file-name').textContent = '오류: ' + err.message;
    }
  };
  f.name.toLowerCase().endsWith('.csv') ? r.readAsText(f, 'UTF-8') : r.readAsBinaryString(f);
}

function parseCSV(txt) {
  var lines = txt.trim().split(/\r?\n/);
  cols = lines[0].split(',').map(function(c) { return c.replace(/^"|"$/g,'').trim(); });
  rows = lines.slice(1).map(function(line) {
    var cells = [], cur2 = '', inQ = false;
    for (var i = 0; i < line.length; i++) {
      var ch = line[i];
      if (ch === '"') { inQ = !inQ; }
      else if (ch === ',' && !inQ) { cells.push(cur2.trim()); cur2 = ''; }
      else { cur2 += ch; }
    }
    cells.push(cur2.trim());
    return cells;
  }).filter(function(r) { return r.some(function(c) { return c !== ''; }); });
}

/* ── 열 문자 변환 ── */
function colLetter(i) {
  var s = '', n = i + 1;
  while (n > 0) {
    var r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function showCols() {
  document.getElementById('col-area').classList.remove('hidden');
  document.getElementById('col-tags').innerHTML = cols.map(function(c, i) {
    return '<span class="tag tag-static col-chip"><span class="col-letter">' + colLetter(i) + '</span>' + c + '</span>';
  }).join('');
  document.getElementById('data-stat').textContent = cols.length + '개 열 · ' + rows.length + '개 행 로드됨';
}

function updateEmailSel() {
  var s = document.getElementById('email-col');
  s.innerHTML = '<option value="">— 선택 —</option>' + cols.map(function(c, i) {
    return '<option value="' + i + '">' + colLetter(i) + '열 — ' + c + '</option>';
  }).join('');
}

/* ── Mapping ── */
function updateMapping() {
  var vars = getVars();
  var area = document.getElementById('mapping-area');
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

  area.innerHTML = vars.map(function(v) {
    var ai = -1;
    for (var i = 0; i < cols.length; i++) {
      if (cols[i].trim().toLowerCase() === v.trim().toLowerCase()) { ai = i; break; }
    }
    var opts = cols.map(function(c, i) {
      return '<option value="' + i + '"' + (i === ai ? ' selected' : '') + '>' + colLetter(i) + '열 — ' + c + '</option>';
    }).join('');
    return '<div class="mapping-row">' +
      '<span class="m-tag">{{' + v + '}}</span>' +
      '<span class="m-arrow">→</span>' +
      '<select class="m-sel" data-var="' + v + '">' +
      '<option value="">— 선택 안 함 —</option>' + opts +
      '</select></div>';
  }).join('');
}

/* ── Template apply ── */
var currentRow = [];

function applyTmpl(tmpl, vm) {
  var decoded = tmpl.replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&amp;/g,'&');
  return decoded.replace(/\{\{([^{}]+)\}\}/g, function(match, k) {
    k = k.trim();
    var ci = vm[k];
    if (ci !== undefined && ci !== '') {
      var val = currentRow[parseInt(ci)];
      return val !== undefined ? String(val) : match;
    }
    return match;
  });
}

/* ── Generate previews ── */
function genPreviews() {
  var subj = document.getElementById('subj').value || '';
  var edEl = ed();
  var bodyHTML = edEl ? edEl.innerHTML : '';
  var emailColVal = document.getElementById('email-col').value;
  var emailColIdx = emailColVal !== '' ? parseInt(emailColVal) : -1;
  var rs = parseInt(document.getElementById('row-s').value) - 2;
  var re = parseInt(document.getElementById('row-e').value);

  var vm = {};
  document.querySelectorAll('[data-var]').forEach(function(s) {
    if (s.value !== '') vm[s.dataset.var] = s.value;
  });

  var wr = rows.slice(Math.max(0, rs));
  if (re > 0) {
    var startRow = parseInt(document.getElementById('row-s').value);
    wr = wr.slice(0, re - startRow + 1);
  }

  if (!wr.length) { showMapStatus('데이터 행이 없습니다. 파일과 행 범위를 확인하세요.', 'warn'); return; }

  previews = wr.map(function(row, i) {
    currentRow = row;
    var email = emailColIdx >= 0 ? String(row[emailColIdx] || '') : '';
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
    '총 ' + previews.length + '개 (유효: ' + previews.filter(function(p) { return p.valid; }).length + '개)';
  showMapStatus(previews.length + '개 미리보기 생성 완료', 'ok');
}

function showMapStatus(msg, type) {
  document.getElementById('map-status').innerHTML = '<div class="status s-' + type + '">' + msg + '</div>';
}

/* ── Preview ── */
function renderPrev() {
  if (!previews.length) return;
  var p = previews[cur];
  document.getElementById('prev-count').textContent = (cur + 1) + ' / ' + previews.length;
  var badge = document.getElementById('cur-badge');
  badge.style.display = 'inline-block';
  badge.className = 'badge ' + (p.valid ? 'badge-ok' : 'badge-warn');
  badge.textContent = p.valid ? '유효' : '주소 없음';
  document.getElementById('prev-area').innerHTML =
    '<div class="email-frame">' +
    '<div class="email-meta">' +
    '<div class="email-meta-row"><span class="meta-label">받는 이</span><span class="meta-val">' + (p.to || '(없음)') + '</span></div>' +
    '<div class="email-meta-row"><span class="meta-label">제목</span><span class="meta-val">' + (p.subj || '(없음)') + '</span></div>' +
    '</div><div class="email-body-preview">' + p.body + '</div></div>';
}

function nav(d) {
  if (!previews.length) return;
  cur = (cur + d + previews.length) % previews.length;
  renderPrev();
}

/* ── Send / Export ── */
function sendGmail() {
  if (!previews.length) return;
  var p = previews[cur];
  if (!p.valid) { showMapStatus('현재 항목의 이메일 주소가 없습니다.', 'warn'); return; }
  var plain = p.body.replace(/<br\s*\/?>/gi,'\n').replace(/<\/p>/gi,'\n').replace(/<[^>]+>/g,'');
  var url = 'https://mail.google.com/mail/?view=cm&fs=1&to=' + encodeURIComponent(p.to) +
    '&su=' + encodeURIComponent(p.subj) + '&body=' + encodeURIComponent(plain);
  window.open(url, '_blank');
}

function exportHTML() {
  if (!previews.length) return;
  var p = previews[cur];
  var html = '<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">' +
    '<title>' + p.subj + '</title>' +
    '<style>body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;max-width:620px;margin:40px auto;font-size:14px;line-height:1.75;color:#1a1a18}.meta{color:#888;font-size:12px;margin-bottom:16px;padding-bottom:12px;border-bottom:1px solid #eee}</style>' +
    '</head><body><div class="meta">To: ' + p.to + '<br>Subject: ' + p.subj + '</div>' + p.body + '</body></html>';
  var blob = new Blob([html], { type: 'text/html;charset=utf-8' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = 'email_' + String(cur + 1).padStart(3, '0') + '.html';
  a.click();
  URL.revokeObjectURL(url);
}

/**
 * 🌙 저녁 3시간 — 구글 시트 동기화 스크립트
 *
 * 설치 (처음 한 번)
 *  1) 아래 SECRET 을 원하는 비밀번호로 바꾸고 저장(Ctrl+S)
 *  2) 위쪽 함수 선택에서 setup 을 고르고 ▶실행 → 권한 허용 (시트 모양이 만들어져요)
 *  3) 배포 → 새 배포 → 유형: 웹 앱 / 실행 사용자: 나 / 액세스 권한: 모든 사용자 → 배포
 *  4) 나온 "웹 앱 URL" 과 비밀번호를 앱 ⚙️설정에 넣고 [연결하고 전부 보내기]
 *
 * 코드를 고친 뒤에는: 배포 → 배포 관리 → ✏️수정 → 버전: 새 버전 → 배포
 */
const SECRET = '여기에-비밀번호';

const CATS = {
  biz:   { name: '사업',   emoji: '💼', color: '#4C6EF5' },
  ai:    { name: 'AI학습', emoji: '🤖', color: '#9B5DE5' },
  read:  { name: '독서',   emoji: '📚', color: '#12B886' },
  gym:   { name: '운동',   emoji: '🏃', color: '#F76707' },
  video: { name: '영상',   emoji: '🎬', color: '#E64980' },
  etc:   { name: '기타',   emoji: '✨', color: '#868E96' },
};
const ORDER = ['biz', 'ai', 'read', 'gym', 'video', 'etc'];
const ST = {
  done: { mark: '✓', score: 1,   label: '했음' },
  part: { mark: '△', score: 0.5, label: '일부' },
  miss: { mark: '✕', score: 0,   label: '못함' },
};
const DOW = ['일', '월', '화', '수', '목', '금', '토'];
const LOG = '저녁기록', YEAR = '한눈에 1년', DATA = '_data';
const HEAD = 4;          // 저녁기록 시트에서 데이터가 시작되는 행
const KEY_RE = /^\d{4}-\d{2}-\d{2}$/;

/* ───────── 웹 앱 입구 ───────── */
function doGet() {
  return ContentService.createTextOutput('🌙 저녁 3시간 동기화 서버가 작동 중이에요. 이 URL을 앱 설정에 넣으세요.');
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: '잘못된 요청이에요' }); }
  if (SECRET === '여기에-비밀번호') return out_({ ok: false, error: '스크립트 맨 위 SECRET 비밀번호를 먼저 바꿔 주세요' });
  if (String(req.key || '') !== SECRET) return out_({ ok: false, error: '비밀번호가 맞지 않아요' });
  try {
    if (req.action === 'ping') return out_({ ok: true });
    if (req.action === 'load') return out_({ ok: true, days: readData_() });
    if (req.action === 'save') {
      const lock = LockService.getScriptLock();
      lock.waitLock(25000);
      try {
        const days = readData_();
        const inc = req.days || {};
        let n = 0;
        Object.keys(inc).forEach(function (k) {
          if (!KEY_RE.test(k)) return;
          const d = clean_(inc[k]);
          if (!days[k] || d.u >= (days[k].u || 0)) { days[k] = d; n++; }
        });
        const p = PropertiesService.getScriptProperties();
        if (req.today && KEY_RE.test(req.today)) p.setProperty('TODAY', req.today);
        if (n) writeData_(days);
        renderAll_(days);
        return out_({ ok: true, saved: n });
      } finally {
        lock.releaseLock();
      }
    }
    return out_({ ok: false, error: '알 수 없는 요청이에요' });
  } catch (err) {
    return out_({ ok: false, error: '시트 오류: ' + err.message });
  }
}

/** 편집기에서 한 번 실행: 시트 모양 만들기 + 권한 허용 */
function setup() {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  PropertiesService.getScriptProperties().setProperty('TODAY', today);
  renderAll_(readData_());
  SpreadsheetApp.getActive().toast('🌙 저녁 3시간 시트 준비 완료!');
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ───────── 원본 데이터 (숨김 시트 _data) ───────── */
function dataSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DATA);
  if (!sh) {
    sh = ss.insertSheet(DATA);
    sh.getRange(1, 1, 1, 3).setValues([['date', 'json', 'updated']]);
    sh.hideSheet();
  }
  return sh;
}
function readData_() {
  const sh = dataSheet_();
  const n = sh.getLastRow() - 1;
  const days = {};
  if (n < 1) return days;
  sh.getRange(2, 1, n, 2).getDisplayValues().forEach(function (r) {
    if (!KEY_RE.test(r[0])) return;
    try { days[r[0]] = clean_(JSON.parse(r[1])); } catch (e) {}
  });
  return days;
}
function writeData_(days) {
  const sh = dataSheet_();
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 3).clearContent();
  const keys = Object.keys(days).sort();
  if (!keys.length) return;
  ensureRows_(sh, keys.length + 1);
  sh.getRange(2, 1, keys.length, 2).setNumberFormat('@');
  sh.getRange(2, 1, keys.length, 3).setValues(keys.map(function (k) {
    return [k, JSON.stringify(days[k]), days[k].u ? new Date(days[k].u) : ''];
  }));
}
function clean_(d) {
  d = d || {};
  const s = (Array.isArray(d.s) ? d.s : []).slice(0, 3).map(function (x) {
    x = x || {};
    return { c: CATS[x.c] ? x.c : '', m: String(x.m || '').slice(0, 60), st: ST[x.st] ? x.st : '' };
  });
  while (s.length < 3) s.push({ c: '', m: '', st: '' });
  return { s: s, u: Number(d.u) || 0 };
}

/* ───────── 그리기 ───────── */
function renderAll_(days) {
  const p = PropertiesService.getScriptProperties();
  const today = p.getProperty('TODAY') || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  renderLog_(days, today);
  renderYear_(days, today);
}

/** 시트 1: 저녁기록 — 앱의 주간 보드와 같은 모양 (하루 1줄, 칸 3개) */
function renderLog_(days, today) {
  const sh = sheet_(LOG, 0);
  sh.clear();
  ensureCols_(sh, 9);

  // 머리말
  sh.getRange(1, 1, 3, 7).setValues([
    ['🌙 저녁 3시간 기록', '', '', '', '', '', ''],
    ['진한 색 = 했음 · 반만 진함 = 일부 · 연한 색 = 체크 전 · 회색 취소선 = 못함 · 달성: 1시간 김 / 2시간 새 / 3시간 별', '', '', '', '', '', ''],
    ['주차', '날짜', '요일', '1시간', '2시간', '3시간', '달성'],
  ]);
  sh.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  sh.getRange(2, 1).setFontSize(10).setFontColor('#868E96');
  sh.getRange(3, 1, 1, 7).setBackground('#212529').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sh.setFrozenRows(3);

  // 범례 (오른쪽)
  const lg = [['카테고리', '#FFFFFF', '#212529', 'bold', 'none']];
  ORDER.forEach(function (id) { const c = CATS[id]; lg.push([c.emoji + ' ' + c.name, c.color, '#FFFFFF', 'bold', 'none']); });
  lg.push(['', '#FFFFFF', '#212529', 'normal', 'none']);
  lg.push(['상태 (예: 💼 사업)', '#FFFFFF', '#212529', 'bold', 'none']);
  ['done', 'part', '', 'miss'].forEach(function (st) {
    const cell = slotCell_({ c: 'biz', m: '', st: st });
    cell[0] += st === 'done' ? '  ← 했음' : st === 'part' ? '  ← 일부' : st === 'miss' ? '  ← 못함' : '  ← 체크 전';
    lg.push(cell);
  });
  const lr = sh.getRange(1, 9, lg.length, 1);
  lr.setValues(lg.map(function (x) { return [x[0]]; }))
    .setBackgrounds(lg.map(function (x) { return [x[1]]; }))
    .setFontColors(lg.map(function (x) { return [x[2]]; }))
    .setFontWeights(lg.map(function (x) { return [x[3]]; }))
    .setFontLines(lg.map(function (x) { return [x[4]]; }));

  // 주 단위로 하루 1줄
  const keys = Object.keys(days).filter(function (k) { return hasPlan_(days[k]); }).sort();
  const first = monday_(parseK_(keys.length && keys[0] < today ? keys[0] : today));
  const lastKey = keys.length && keys[keys.length - 1] > today ? keys[keys.length - 1] : today;
  const last = monday_(parseK_(lastKey));

  const V = [], B = [], F = [], W = [], L = [], weekTops = [];
  let band = 0;
  for (let mon = first; mon <= last; mon = addDays_(mon, 7)) {
    const wk = [];
    for (let i = 0; i < 7; i++) wk.push(keyOf_(addDays_(mon, i)));
    let planned = 0, score = 0;
    wk.forEach(function (k) { const x = info_(days[k]); if (k <= today) planned += x.planned; score += x.score; });
    const rate = planned ? Math.round(score / planned * 100) : null;
    const bandBg = band++ % 2 ? '#F5F3FF' : '#FFFFFF';
    const aBg = rate === null ? '#F1F3F5' : mix_('#2F9E44', 0.12 + 0.5 * rate / 100);
    weekTops.push(HEAD + V.length);

    wk.forEach(function (k, i) {
      const dt = addDays_(mon, i), d = days[k], dow = dt.getDay();
      const isToday = k === today;
      const v = [i === 0 ? weekLabel_(mon) + (rate === null ? '' : '\n달성 ' + rate + '%') : '',
        (dt.getMonth() + 1) + '/' + dt.getDate(), DOW[dow]];
      const b = [aBg, isToday ? '#FFE8A3' : bandBg, isToday ? '#FFE8A3' : bandBg];
      const f = ['#1B4332', '#495057', dow === 6 ? '#1C7ED6' : dow === 0 ? '#E03131' : '#495057'];
      const w = [i === 0 ? 'bold' : 'normal', isToday ? 'bold' : 'normal', 'bold'];
      const l = ['none', 'none', 'none'];
      for (let j = 0; j < 3; j++) {
        const c = slotCell_(d && d.s[j]);
        v.push(c[0]); b.push(c[1]); f.push(c[2]); w.push(c[3]); l.push(c[4]);
      }
      const ksb = kimSaeByeol_(d);
      v.push(ksb); b.push(ksb === '김새별' ? '#FFE066' : bandBg); f.push('#212529'); w.push('bold'); l.push('none');
      V.push(v); B.push(b); F.push(f); W.push(w); L.push(l);
    });
  }

  const n = V.length;
  ensureRows_(sh, HEAD + n);
  const r = sh.getRange(HEAD, 1, n, 7);
  r.setNumberFormat('@')
    .setValues(V).setBackgrounds(B).setFontColors(F).setFontWeights(W).setFontLines(L)
    .setVerticalAlignment('middle').setWrap(true).setFontSize(10);
  sh.getRange(HEAD, 1, n, 3).setHorizontalAlignment('center');
  sh.getRange(HEAD, 7, n, 1).setHorizontalAlignment('center').setFontSize(12);
  r.setBorder(true, true, true, true, true, true, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRangeList(weekTops.map(function (x) { return 'A' + x + ':G' + x; }))
    .setBorder(true, null, null, null, null, null, '#495057', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sh.setColumnWidth(1, 80);
  sh.setColumnWidth(2, 52);
  sh.setColumnWidth(3, 36);
  sh.setColumnWidths(4, 3, 200);
  sh.setColumnWidth(7, 64);
  sh.setColumnWidth(8, 18);
  sh.setColumnWidth(9, 150);
}

/** 시트 2: 한눈에 1년 — 12개월 × 31일 색칠 + 카테고리별 시간 */
function renderYear_(days, today) {
  const sh = sheet_(YEAR, 1);
  sh.clear();
  sh.clearNotes();
  const COLS = 41;                 // A월 | B~AF 1~31일 | AG 간격 | AH~AM 카테고리 | AN 합계 | AO 달성률
  const GAP = 32, CAT0 = 33, TOT = 39, RATE = 40;
  ensureCols_(sh, COLS);

  const keys = Object.keys(days).filter(function (k) { return hasPlan_(days[k]); }).sort();
  const yT = Number(today.slice(0, 4));
  const y0 = keys.length ? Math.min(Number(keys[0].slice(0, 4)), yT) : yT;
  const y1 = keys.length ? Math.max(Number(keys[keys.length - 1].slice(0, 4)), yT) : yT;

  const V = [], B = [], F = [], W = [], N = [];
  function row(bg) {
    const o = { v: [], b: [], f: [], w: [], n: [] };
    for (let i = 0; i < COLS; i++) { o.v.push(''); o.b.push(bg || '#FFFFFF'); o.f.push('#212529'); o.w.push('normal'); o.n.push(''); }
    return o;
  }
  function push(o) { V.push(o.v); B.push(o.b); F.push(o.f); W.push(o.w); N.push(o.n); }

  const t1 = row(); t1.v[0] = '🌙 한눈에 1년 — 매일 저녁을 무엇으로 채웠나'; t1.w[0] = 'bold'; push(t1);
  const t2 = row(); t2.v[0] = '그날 가장 많이 한 카테고리 색 · 진할수록 3시간을 꽉 채운 날 · 칸에 마우스를 올리면 그날 기록이 보여요'; t2.f[0] = '#868E96'; push(t2);
  const blockStarts = [];

  for (let Y = y1; Y >= y0; Y--) {
    blockStarts.push(V.length + 1);
    const h = row('#343A40');
    h.v[0] = Y + '년'; h.b[0] = '#212529'; h.f[0] = '#FFFFFF'; h.w[0] = 'bold';
    for (let d = 1; d <= 31; d++) { h.v[d] = d; h.f[d] = '#ADB5BD'; }
    h.b[GAP] = '#FFFFFF';
    ORDER.forEach(function (id, j) { const c = CATS[id]; h.v[CAT0 + j] = c.emoji + ' ' + c.name; h.b[CAT0 + j] = c.color; h.f[CAT0 + j] = '#FFFFFF'; h.w[CAT0 + j] = 'bold'; });
    h.v[TOT] = '합계'; h.v[RATE] = '달성률';
    [TOT, RATE].forEach(function (x) { h.b[x] = '#212529'; h.f[x] = '#FFFFFF'; h.w[x] = 'bold'; });
    push(h);

    const yr = { planned: 0, score: 0, byCat: {} };
    for (let m = 0; m < 12; m++) {
      const o = row();
      o.v[0] = (m + 1) + '월'; o.w[0] = 'bold'; o.b[0] = '#F1F3F5';
      const mo = { planned: 0, score: 0, byCat: {} };
      for (let dd = 1; dd <= 31; dd++) {
        const dt = new Date(Y, m, dd);
        if (dt.getMonth() !== m) continue;               // 없는 날짜(2/30 등)는 흰 칸
        const k = keyOf_(dt), d = days[k], x = info_(d);
        if (x.score > 0) o.b[dd] = mix_(CATS[x.dom].color, 0.3 + 0.7 * x.score / 3);
        else if (x.planned && x.checked) o.b[dd] = '#CED4DA';
        else if (x.planned) o.b[dd] = mix_(CATS[firstCat_(d)].color, 0.15);
        else o.b[dd] = '#F1F3F5';
        if (k === today) { o.v[dd] = '●'; o.f[dd] = '#212529'; }
        if (x.planned) o.n[dd] = noteOf_(dt, d);
        if (k <= today) mo.planned += x.planned;   // 달성률은 오늘까지의 계획만
        mo.score += x.score;
        for (const c in x.byCat) mo.byCat[c] = (mo.byCat[c] || 0) + x.byCat[c];
      }
      ORDER.forEach(function (id, j) {
        const hrs = mo.byCat[id] || 0;
        o.v[CAT0 + j] = hrs || ''; o.b[CAT0 + j] = mix_(CATS[id].color, 0.12); o.f[CAT0 + j] = CATS[id].color; o.w[CAT0 + j] = 'bold';
        yr.byCat[id] = (yr.byCat[id] || 0) + hrs;
      });
      o.v[TOT] = mo.score || ''; o.w[TOT] = 'bold'; o.b[TOT] = '#F1F3F5';
      o.v[RATE] = mo.planned ? Math.round(mo.score / mo.planned * 100) + '%' : ''; o.b[RATE] = '#F1F3F5';
      yr.planned += mo.planned; yr.score += mo.score;
      push(o);
    }
    const s = row('#F8F9FA');
    s.v[0] = '합계'; s.w[0] = 'bold'; s.b[0] = '#E9ECEF';
    s.b[GAP] = '#FFFFFF';
    ORDER.forEach(function (id, j) { s.v[CAT0 + j] = (yr.byCat[id] || 0) + '시간'; s.b[CAT0 + j] = mix_(CATS[id].color, 0.3); s.f[CAT0 + j] = '#212529'; s.w[CAT0 + j] = 'bold'; });
    s.v[TOT] = yr.score + '시간'; s.v[RATE] = yr.planned ? Math.round(yr.score / yr.planned * 100) + '%' : '–';
    [TOT, RATE].forEach(function (x) { s.b[x] = '#E9ECEF'; s.w[x] = 'bold'; });
    push(s);
    push(row());
  }

  ensureRows_(sh, V.length);
  const r = sh.getRange(1, 1, V.length, COLS);
  r.setValues(V).setBackgrounds(B).setFontColors(F).setFontWeights(W).setNotes(N)
    .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(1, 1, 2, 1).setHorizontalAlignment('left');
  sh.getRange(1, 1).setFontSize(16);
  sh.getRange(3, 1, V.length - 2, COLS).setBorder(true, true, true, true, true, true, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeights(3, V.length - 2, 24);
  sh.setColumnWidth(1, 52);
  sh.setColumnWidths(2, 31, 24);
  sh.setColumnWidth(GAP + 1, 10);
  sh.setColumnWidths(CAT0 + 1, 6, 78);
  sh.setColumnWidths(TOT + 1, 2, 62);
}

/* ───────── 도우미 ───────── */
function slotCell_(s) {               // [글자, 배경, 글자색, 굵기, 취소선]
  if (!s || !s.c) return ['', '#FFFFFF', '#ADB5BD', 'normal', 'none'];
  const c = CATS[s.c];
  const text = c.emoji + ' ' + c.name + (s.m ? ' · ' + s.m : '');
  if (s.st === 'done') return ['✓ ' + text, c.color, '#FFFFFF', 'bold', 'none'];
  if (s.st === 'part') return ['△ ' + text, mix_(c.color, 0.45), '#212529', 'bold', 'none'];
  if (s.st === 'miss') return ['✕ ' + text, '#E9ECEF', '#868E96', 'normal', 'line-through'];
  return [text, mix_(c.color, 0.14), c.color, 'normal', 'none'];
}
/** 달성 표시: 1시간 했음 = 김, 2시간 = 새, 3시간 = 별 → 모두 했으면 '김새별' */
function kimSaeByeol_(d) {
  if (!d) return '';
  const word = ['김', '새', '별'];
  return d.s.map(function (s, j) { return s.c && s.st === 'done' ? word[j] : ''; }).join('');
}
function noteOf_(dt, d) {
  return (dt.getMonth() + 1) + '/' + dt.getDate() + ' (' + DOW[dt.getDay()] + ')\n' + d.s.map(function (s, j) {
    if (!s.c) return (j + 1) + '시간  –';
    const c = CATS[s.c];
    return (j + 1) + '시간  ' + (s.st ? ST[s.st].mark + ' ' : '') + c.emoji + ' ' + c.name + (s.m ? ' · ' + s.m : '');
  }).join('\n');
}
function info_(d) {
  const r = { planned: 0, checked: 0, score: 0, byCat: {}, dom: null };
  if (!d) return r;
  d.s.forEach(function (s) {
    if (!s.c) return;
    r.planned++;
    if (s.st) r.checked++;
    const sc = s.st ? ST[s.st].score : 0;
    r.score += sc;
    if (sc) r.byCat[s.c] = (r.byCat[s.c] || 0) + sc;
  });
  let bv = 0;
  for (const c in r.byCat) if (r.byCat[c] > bv) { bv = r.byCat[c]; r.dom = c; }
  return r;
}
function hasPlan_(d) { return !!(d && d.s.some(function (s) { return s.c; })); }
function firstCat_(d) { for (let i = 0; i < 3; i++) if (d.s[i].c) return d.s[i].c; return 'etc'; }
function weekLabel_(mon) { const t = addDays_(mon, 3); return (t.getMonth() + 1) + '월 ' + Math.ceil(t.getDate() / 7) + '주'; }
function mix_(hex, t) {               // 흰색과 섞기 (t=1 이면 원래 색)
  const n = parseInt(hex.slice(1), 16);
  const f = function (c) { return ('0' + Math.round(255 + (c - 255) * t).toString(16)).slice(-2); };
  return '#' + f(n >> 16 & 255) + f(n >> 8 & 255) + f(n & 255);
}
function pad_(n) { return (n < 10 ? '0' : '') + n; }
function keyOf_(d) { return d.getFullYear() + '-' + pad_(d.getMonth() + 1) + '-' + pad_(d.getDate()); }
function parseK_(k) { const p = k.split('-').map(Number); return new Date(p[0], p[1] - 1, p[2]); }
function addDays_(d, n) { const x = new Date(d.getFullYear(), d.getMonth(), d.getDate()); x.setDate(x.getDate() + n); return x; }
function monday_(d) { return addDays_(d, -((d.getDay() + 6) % 7)); }

function sheet_(name, index) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  // 처음: 비어 있는 기본 시트(시트1/Sheet1)가 있으면 그걸 이름만 바꿔 씀
  const blank = ss.getSheets().filter(function (s) { return /^(시트|Sheet)\d+$/.test(s.getName()) && s.getLastRow() === 0; })[0];
  if (blank) { blank.setName(name); sh = blank; }
  else sh = ss.insertSheet(name);
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(index + 1);
  return sh;
}
function ensureRows_(sh, n) { const m = sh.getMaxRows(); if (m < n) sh.insertRowsAfter(m, n - m); }
function ensureCols_(sh, n) { const m = sh.getMaxColumns(); if (m < n) sh.insertColumnsAfter(m, n - m); }

// 흉터 · Code.gs (기록 v2)
// 값을 순서대로 적지 않고, 열 이름을 보고 맞는 칸에 적는다.
// 기록마다 값이 빠지거나 순서가 달라도 다른 열로 밀리지 않는다.
// 열 구성이 예전과 다른 시트는 이름 뒤에 '(이전)'을 붙여 옆으로 옮기고 새로 만든다.

const HTML_FILE = 'index'; // 올려 둔 HTML 파일 이름 (.html 빼고)
const SHEET_ID = '';       // 다른 시트에 적으려면 그 시트 주소의 ID를 넣는다. 비우면 이 스크립트가 붙은 시트를 쓴다.

const REGN = ['들판', '숲', '폐허', '고개', '바닷가', '동굴', '섬'];
const BASE = [['ts', '시각', 'd'], ['dy', '날짜', 't'], ['u', '사용자', 't'], ['v', '버전', 't'], ['e', '기록 ID', 't']];
const RUN = [['r', '판 ID', 't'], ['rn', '몇 번째 판', 'n'], ['reg', '지역', 't']];
const SCHEMA = {
  open: { name: '접속', cols: [['first', '처음 온 날', 't'], ['fresh', '새 사용자', 'n'], ['gap', '지난 접속 뒤 일수', 'n'], ['runs', '누적 판 수', 'n'], ['phone', '휴대폰', 'n'], ['scr', '화면 크기', 't'], ['rfx', '효과 줄이기', 'n'], ['vm', '음악', 'n'], ['vs', '효과음', 'n'], ['save', '이어하기 있음', 'n'], ['stor', '저장 가능', 'n']] },
  run: { name: '판', cols: RUN.concat([['res', '결과', 't'], ['fw', '이긴 결투', 'n'], ['sec', '결투 시간(초)', 'n'], ['wall', '플레이 시간(초)', 'n'], ['aw', '각성', 'n'], ['en', '버팀', 'n'], ['c2', '연속 처치', 'n'], ['sc', '흉터', 'n'], ['pet', '동료', 't'], ['plv', '친밀', 'n'], ['wp', '무기', 't'], ['wl', '강화', 'n'], ['tr', '수련 공/막/피', 't']]) },
  fight: { name: '결투', cols: RUN.concat([['ly', '칸', 'n'], ['kind', '종류', 't'], ['foe', '상대', 't'], ['res', '결과', 't'], ['hp0', '시작 체력', 'n'], ['hp', '끝 체력', 'n'], ['mhp', '최대 체력', 'n'], ['ehp', '상대 남은 체력(%)', 'n'], ['sec', '시간(초)', 'n'], ['hv', '강타', 'n'], ['pf', '완벽 강타', 'n'], ['pa', '받아치기', 'n'], ['dg', '피하기', 'n'], ['bl', '막기', 'n'], ['ht', '맞음', 'n'], ['aw', '각성', 'n'], ['en', '버팀', 'n'], ['gh', '머리 흉터', 'n'], ['gt', '몸통 흉터', 'n'], ['ga', '팔 흉터', 'n'], ['gl', '다리 흉터', 'n'], ['lp', '쓰러진 부위', 't'], ['pet', '동료', 't'], ['wp', '무기', 't'], ['wl', '강화', 'n']]) },
  map: { name: '지도', cols: RUN.concat([['ly', '칸', 'n'], ['kind', '구분', 't'], ['at', '장소', 't'], ['pick', '고른 것', 't'], ['opts', '갈 수 있던 칸', 't']]) },
  err: { name: '오류', cols: [['msg', '내용', 't'], ['at', '위치', 't'], ['ph', '화면', 't']] }
};
for (const k in SCHEMA) {
  const S = SCHEMA[k];
  S.all = BASE.concat(S.cols);
  S.head = S.all.map(c => c[1]);
  S.fmt = S.all.map(c => c[2] === 'n' ? '0' : c[2] === 'd' ? 'yyyy-mm-dd hh:mm:ss' : '@');
}
const BK = { name: '백업', head: ['사용자', '맡긴 시각', '버전', '크기', '데이터'], fmt: ['@', 'yyyy-mm-dd hh:mm:ss', '@', '0', '@'] };
const SUM = '요약', SUMK = '흉터 요약 v2';
let MIG = false;

// ----- 웹 앱 -----
function doGet() {
  return HtmlService.createHtmlOutputFromFile(HTML_FILE).setTitle('흉터')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover');
}

// 편집기에서 한 번 실행: 권한 허용 + 시트·요약 만들기
function setup() {
  const ss = book_();
  sum_(ss, true);
  tab_(ss, BK.name, BK.head);
  Logger.log(ss.getUrl());
}

// ----- 시트 도구 -----
function book_() {
  if (SHEET_ID) return SpreadsheetApp.openById(SHEET_ID);
  const a = SpreadsheetApp.getActiveSpreadsheet();
  if (a) return a;
  const P = PropertiesService.getScriptProperties(), id = P.getProperty('BOOK_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  const b = SpreadsheetApp.create('흉터 플레이 기록');
  P.setProperty('BOOK_ID', b.getId());
  return b;
}
function letter_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }
function uniq_(ss, n) { let k = n, i = 2; while (ss.getSheetByName(k)) k = n + ' ' + (i++); return k; }
function fit_(sh, r, c) {
  if (sh.getMaxRows() < r) sh.insertRowsAfter(sh.getMaxRows(), r - sh.getMaxRows());
  if (sh.getMaxColumns() < c) sh.insertColumnsAfter(sh.getMaxColumns(), c - sh.getMaxColumns());
}
function head_(sh) {
  if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) return [];
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  while (h.length && h[h.length - 1] === '') h.pop();
  return h;
}
// 열 이름이 같으면 그대로, 뒤에 열만 늘었으면 이어 붙이고, 다르면 예전 시트를 옮기고 새로 만든다
function tab_(ss, name, head) {
  let sh = ss.getSheetByName(name);
  if (sh && sh.getLastRow() > 0) {
    const cur = head_(sh), pre = cur.length <= head.length && cur.every((x, i) => x === head[i]);
    if (pre && cur.length === head.length) return sh;
    if (!pre) { sh.setName(uniq_(ss, name + ' (이전)')); MIG = true; sh = null; }
  }
  if (!sh) sh = ss.insertSheet(name);
  fit_(sh, 1, head.length);
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}
function put_(sh, rows, fmt) {
  const r0 = Math.max(2, sh.getLastRow() + 1), n = rows.length, w = fmt.length;
  fit_(sh, r0 + n - 1, w);
  const rg = sh.getRange(r0, 1, n, w);
  rg.setNumberFormats(rows.map(() => fmt));
  rg.setValues(rows);
}
function cell_(c, o) {
  const k = c[0], t = c[2];
  let x = k === 'ts' ? new Date() : o[k];
  if (k === 'reg' && x != null && REGN[x] != null) x = REGN[x];
  if (k === 'ly') x = typeof x === 'number' && x >= 0 ? x + 1 : '';
  if (x == null || x === '') return '';
  if (t === 'd') return x;
  if (t === 'n') { x = Number(x); return isFinite(x) ? x : ''; }
  x = String(x).slice(0, 500);
  return /^[=+\-@]/.test(x) ? ' ' + x : x;
}

// ----- 플레이 기록 -----
function logBatch(b) {
  if (!Array.isArray(b) || !b.length) return 0;
  const L = LockService.getScriptLock();
  L.waitLock(20000);
  try {
    const ss = book_(), C = CacheService.getScriptCache(), ids = [];
    b.forEach(o => { if (o && o.e) ids.push(String(o.e).slice(0, 200)); });
    const seen = ids.length ? C.getAll(ids) : {}, mark = {}, by = {};
    b.forEach(o => {
      if (!o || typeof o !== 'object') return;
      const S = SCHEMA[o.k];
      if (!S) return;
      const id = o.e ? String(o.e).slice(0, 200) : '';
      if (id) { if (seen[id] || mark[id]) return; mark[id] = '1'; }
      (by[o.k] = by[o.k] || []).push(S.all.map(c => cell_(c, o)));
    });
    for (const k in by) { const S = SCHEMA[k]; put_(tab_(ss, S.name, S.head), by[k], S.fmt); }
    if (Object.keys(mark).length) C.putAll(mark, 21600);
    sum_(ss, false);
    return b.length;
  } finally { L.releaseLock(); }
}

// ----- 요약 (수식이라 저절로 갱신된다) -----
function rng_(k, a, b) {
  const S = SCHEMA[k], i = x => S.all.findIndex(c => c[0] === x) + 1;
  return `'${S.name}'!${letter_(i(a))}2:${letter_(i(b || a))}`;
}
function sig_() { return Object.keys(SCHEMA).map(k => SCHEMA[k].name + ':' + SCHEMA[k].head.join('|')).join('/'); }
function sum_(ss, force) {
  const P = PropertiesService.getScriptProperties(), sig = sig_();
  let sh = ss.getSheetByName(SUM);
  if (!force && !MIG && sh && sh.getRange(1, 1).getDisplayValue() === SUMK && P.getProperty('SUMSIG') === sig) return sh;
  for (const k in SCHEMA) tab_(ss, SCHEMA[k].name, SCHEMA[k].head);
  if (sh && sh.getRange(1, 1).getDisplayValue() !== SUMK) { sh.setName(uniq_(ss, SUM + ' (이전)')); sh = null; }
  if (sh) sh.clear(); else sh = ss.insertSheet(SUM, 0);
  const F = v => `COUNTIF(${rng_('fight', 'res')},"${v}")`, R0 = 3;
  const S0 = [
    ['전체 사용자', `=COUNTUNIQUE(${rng_('open', 'u')})`],
    ['오늘 온 사용자', `=IFERROR(COUNTUNIQUE(FILTER(${rng_('open', 'u')},${rng_('open', 'dy')}=TEXT(TODAY(),"yyyy-mm-dd"))),0)`],
    ['새 사용자', `=COUNTIF(${rng_('open', 'fresh')},1)`],
    ['다음 날 다시 온 횟수', `=COUNTIF(${rng_('open', 'gap')},1)`],
    ['시작한 판', `=COUNTIF(${rng_('run', 'res')},"시작")`],
    ['이긴 판', `=COUNTIF(${rng_('run', 'res')},"승리")`],
    ['진 판', `=COUNTIF(${rng_('run', 'res')},"패배")`],
    ['판 승률', ''],
    ['결투 승률', `=IFERROR(${F('승리')}/(${F('승리')}+${F('패배')}),0)`],
    ['결투에서 나감', `=${F('나감')}`],
    ['끊긴 결투', `=${F('끊김')}`],
    ['오류', `=COUNTA(${rng_('err', 'e')})`]];
  const rw = l => R0 + S0.findIndex(x => x[0] === l);
  S0[rw('판 승률') - R0][1] = `=IFERROR(B${rw('이긴 판')}/(B${rw('이긴 판')}+B${rw('진 판')}),0)`;
  fit_(sh, R0 + S0.length, 24);
  sh.getRange(1, 1).setValue(SUMK).setFontWeight('bold');
  sh.getRange(R0, 1, S0.length, 1).setValues(S0.map(x => [x[0]]));
  sh.getRange(R0, 2, S0.length, 1).setFormulas(S0.map(x => [x[1]]));
  ['판 승률', '결투 승률'].forEach(l => sh.getRange(rw(l), 2).setNumberFormat('0%'));
  const T = [
    [4, '날짜별 사용자', `QUERY(UNIQUE(${rng_('open', 'dy', 'u')}),"select Col1, count(Col2) where Col1 <> '' group by Col1 order by Col1 desc label Col1 '날짜', count(Col2) '사용자'",0)`],
    [7, '상대별 결투', `QUERY({${rng_('fight', 'foe', 'res')}},"select Col1, count(Col1) where Col1 <> '' group by Col1 pivot Col2 label Col1 '상대'",0)`],
    [13, '지역별 판', `QUERY({${rng_('run', 'reg', 'res')}},"select Col1, count(Col1) where Col1 <> '' and Col2 <> '시작' group by Col1 pivot Col2 label Col1 '지역'",0)`],
    [17, '쓰러진 부위', `QUERY({${rng_('fight', 'lp')}},"select Col1, count(Col1) where Col1 <> '' group by Col1 order by count(Col1) desc label Col1 '부위', count(Col1) '횟수'",0)`],
    [20, '지도에서 고른 칸', `QUERY({${rng_('map', 'kind', 'pick')}},"select Col3, count(Col3) where Col1 = '칸' group by Col3 order by count(Col3) desc label Col3 '칸', count(Col3) '횟수'",0)`],
    [23, '자주 나는 오류', `QUERY({${rng_('err', 'msg')}},"select Col1, count(Col1) where Col1 <> '' group by Col1 order by count(Col1) desc limit 10 label Col1 '오류', count(Col1) '횟수'",0)`]];
  T.forEach(t => {
    sh.getRange(1, t[0]).setValue(t[1]).setFontWeight('bold');
    sh.getRange(2, t[0]).setFormula(`=IFERROR(${t[2]},"아직 없음")`);
  });
  sh.setColumnWidth(1, 150);
  P.setProperty('SUMSIG', sig);
  MIG = false;
  return sh;
}

// ----- 백업 코드 -----
function okBk_(s) { try { const o = JSON.parse(String(s)); return !!(o && o.d && typeof o.d === 'object'); } catch (e) { return false; } }
function findRow_(sh, code) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  const v = sh.getRange(2, 1, n, 1).getDisplayValues();
  for (let i = 0; i < n; i++) if (String(v[i][0]).toLowerCase() === code) return i + 2;
  return 0;
}
function backupSave(uid, s) {
  uid = String(uid || '').toLowerCase(); s = String(s || '');
  if (!/^[a-z0-9]{8}$/.test(uid) || !s || s.length > 49000 || !okBk_(s)) return false;
  const L = LockService.getScriptLock();
  L.waitLock(20000);
  try {
    const ss = book_(), sh = tab_(ss, BK.name, BK.head), v = String(JSON.parse(s).v || '');
    const r = findRow_(sh, uid) || Math.max(2, sh.getLastRow() + 1);
    fit_(sh, r, 5);
    const rg = sh.getRange(r, 1, 1, 5);
    rg.setNumberFormats([BK.fmt]);
    rg.setValues([[uid, new Date(), v, s.length, s]]);
    return true;
  } finally { L.releaseLock(); }
}
function backupLoad(code) {
  code = String(code || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  if (!/^[a-z0-9]{8}$/.test(code)) return '';
  const ss = book_(), sh = ss.getSheetByName(BK.name);
  if (sh && head_(sh).join('|') === BK.head.join('|')) {
    const r = findRow_(sh, code);
    if (r) return String(sh.getRange(r, 5).getValue() || '');
  }
  return old_(ss, code);
}
// 예전 방식으로 맡긴 백업도 찾아본다
function old_(ss, code) {
  const P = PropertiesService.getScriptProperties().getProperties();
  for (const k in P) if (k.toLowerCase().indexOf(code) >= 0 && okBk_(P[k])) return P[k];
  for (const sh of ss.getSheets()) {
    if (!/백업|backup/i.test(sh.getName()) || sh.getLastRow() < 1) continue;
    const v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
    for (const r of v) {
      if (!r.some(x => String(x).toLowerCase().replace(/[^a-z0-9]/g, '') === code)) continue;
      for (const x of r) if (okBk_(x)) return String(x);
    }
  }
  return '';
}

/**
 * 구글시트 읽기/쓰기. 탭 구성:
 *   인원 · 근태코드 · 조직표기 · 설정 · 변경이력 · 일정_yyyy-MM (월별)
 */

var SHEET = { people: '인원', codes: '근태코드', org: '조직표기', settings: '설정', log: '변경이력' };
var HEAD = {
  people: ['ID', '팀', '소속', '구분1', '구분2', '구분3', '성명', '순서', '관리자', '이메일', '비밀번호', '재직', '제외일'],
  codes: ['표기', '배경색', '글자색', '근무형태', '기본형', '집계묶음', '자주씀', '사용'],
  org: ['경로', '구분자'],
  settings: ['항목', '값'],
  log: ['시각', '수정자ID', '수정자', '대상ID', '대상자', '날짜', '이전', '이후'],
};
var MONTH_PREFIX = '일정_';
var TZ = 'Asia/Seoul';

function ss_() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name, header) {
  var ss = ss_(), sh = ss.getSheetByName(name);
  if (!sh && header) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#eeeeee');
    sh.setFrozenRows(1);
  }
  return sh;
}

function yes_(v) { return v === true || String(v).trim().toUpperCase() === 'Y'; }
function yn_(b) { return b ? 'Y' : ''; }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function cellDate_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v || ''); }

/** 머리글 이름으로 읽은 행 목록. 각 행에 _row(시트 행 번호)가 붙는다. */
function readRows_(name, header) {
  var sh = sheet_(name, header), n = sh.getLastRow();
  if (n < 2) return [];
  var vals = sh.getRange(1, 1, n, sh.getLastColumn()).getValues();
  var h = vals[0].map(String);
  return vals.slice(1).map(function (r, i) {
    var o = { _row: i + 2 };
    h.forEach(function (k, j) { o[k] = r[j]; });
    return o;
  });
}

/* ---------- 설정 ---------- */
function getSetting_(key, dflt) {
  var row = readRows_(SHEET.settings, HEAD.settings).filter(function (r) { return r['항목'] === key; })[0];
  return row && row['값'] !== '' ? row['값'] : dflt;
}
function setSetting_(key, value) {
  var sh = sheet_(SHEET.settings, HEAD.settings);
  var row = readRows_(SHEET.settings, HEAD.settings).filter(function (r) { return r['항목'] === key; })[0];
  if (row) sh.getRange(row._row, 2).setValue(value);
  else sh.appendRow([key, value]);
}

/* ---------- 인원 ---------- */
function toPerson_(r) {
  return {
    id: String(r['ID']), team: String(r['팀'] || ''), org: String(r['소속'] || ''), g1: String(r['구분1'] || ''),
    g2: String(r['구분2'] || ''), g3: String(r['구분3'] || ''), name: String(r['성명'] || ''), order: +r['순서'] || 0,
    admin: yes_(r['관리자']), email: String(r['이메일'] || ''), pwHash: String(r['비밀번호'] || ''),
    active: yes_(r['재직']), removedAt: cellDate_(r['제외일']), _row: r._row,
  };
}
/** 전체 인원 (제외된 사람 포함), 순서대로 */
function allPeople_() {
  return readRows_(SHEET.people, HEAD.people).filter(function (r) { return r['ID'] !== ''; })
    .map(toPerson_).sort(function (a, b) { return a.order - b.order; });
}
function activePeople_() { return allPeople_().filter(function (p) { return p.active; }); }

function writePerson_(p) {
  var sh = sheet_(SHEET.people, HEAD.people);
  var row = [p.id, p.team, p.org, p.g1, p.g2, p.g3, p.name, p.order, yn_(p.admin), p.email, p.pwHash, yn_(p.active), p.removedAt || ''];
  if (p._row) sh.getRange(p._row, 1, 1, row.length).setValues([row]);
  else { sh.appendRow(row); }
}

function newId_() { return 'P' + Utilities.getUuid().replace(/-/g, '').slice(0, 10); }

/* ---------- 근태코드 ---------- */
function getCodes_() {
  var rows = readRows_(SHEET.codes, HEAD.codes);
  if (!rows.length) return DEFAULT_CODES;
  return rows.filter(function (r) { return String(r['표기']).trim(); }).map(function (r) {
    return {
      code: String(r['표기']), bg: String(r['배경색'] || '#ffffff'), fg: String(r['글자색'] || '#000000'),
      work: yes_(r['근무형태']), base: String(r['기본형'] || ''), group: String(r['집계묶음'] || ''),
      primary: yes_(r['자주씀']), active: r['사용'] === '' ? true : yes_(r['사용']),
    };
  });
}
function writeCodes_(codes) {
  var sh = sheet_(SHEET.codes, HEAD.codes);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, HEAD.codes.length).clearContent();
  if (!codes.length) return;
  sh.getRange(2, 1, codes.length, HEAD.codes.length).setValues(codes.map(function (c) {
    return [c.code, c.bg, c.fg, yn_(c.work), c.base || '', c.group || '', yn_(c.primary), yn_(c.active)];
  }));
  sh.getRange(2, 1, codes.length, 1).setWrap(true);
}

/* ---------- 조직표기 ---------- */
function getOrgFormats_() {
  var out = {};
  readRows_(SHEET.org, HEAD.org).forEach(function (r) {
    var s = String(r['구분자']);
    out[String(r['경로'])] = s === '줄바꿈' ? '\n' : s === '붙임' ? '' : ' ';
  });
  return out;
}
function writeOrgFormats_(map) {
  var sh = sheet_(SHEET.org, HEAD.org);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  var rows = Object.keys(map).map(function (k) { return [k, map[k] === '\n' ? '줄바꿈' : map[k] === '' ? '붙임' : '공백']; });
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setValues(rows);
}

/* ---------- 월별 일정 탭 ---------- */
function monthSheetName_(key) { return MONTH_PREFIX + key; }
function monthSheet_(key) { return ss_().getSheetByName(monthSheetName_(key)); }
function existingMonths_() {
  return ss_().getSheets().map(function (s) { return s.getName(); })
    .filter(function (n) { return n.indexOf(MONTH_PREFIX) === 0; })
    .map(function (n) { return n.slice(MONTH_PREFIX.length); }).sort();
}

/** 월 탭 읽기 → { days, rows: { personId: { row, values: [...] } } } */
function readMonth_(key) {
  var sh = monthSheet_(key);
  if (!sh) return null;
  var p = parseMonthKey_(key), days = daysInMonth_(p.y, p.m), n = sh.getLastRow();
  var out = { days: days, rows: {} };
  if (n < 2) return out;
  sh.getRange(2, 1, n - 1, days + 2).getValues().forEach(function (r, i) {
    if (r[0] === '') return;
    out.rows[String(r[0])] = { row: i + 2, values: r.slice(2).map(function (v) { return v == null ? '' : String(v); }) };
  });
  return out;
}

/** 새 월 탭 만들기. rows: [{ id, name, values }] */
function createMonthSheet_(key, rows) {
  var p = parseMonthKey_(key), days = daysInMonth_(p.y, p.m);
  var header = ['ID', '성명'];
  for (var d = 1; d <= days; d++) header.push(d);
  var ss = ss_(), name = monthSheetName_(key);
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#eeeeee');
  // 머리글 요일 색 (토 파랑, 일 빨강)
  for (d = 1; d <= days; d++) {
    var wi = weekdayIdx_(ymd_(p.y, p.m, d));
    if (wi >= 5) sh.getRange(1, d + 2).setFontColor(wi === 5 ? '#0000ff' : '#ff0000');
  }
  if (rows.length) {
    sh.getRange(2, 1, rows.length, header.length).setValues(rows.map(function (r) {
      return [r.id, r.name].concat(r.values);
    }));
  }
  sh.setFrozenRows(1); sh.setFrozenColumns(2);
  sh.hideColumns(1);
  sh.setColumnWidths(3, days, 52);
  sh.getRange(1, 1, Math.max(rows.length + 1, 2), header.length).setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle');
  return sh;
}

/** 사람 한 명의 행을 월 탭에 추가 */
function appendMonthRow_(key, id, name, values) {
  var sh = monthSheet_(key);
  if (!sh) return;
  sh.appendRow([id, name].concat(values));
}

/* ---------- 변경이력 ---------- */
function logLength_() {
  var c = CacheService.getScriptCache().get('logLen');
  if (c !== null) return +c;
  var n = Math.max(sheet_(SHEET.log, HEAD.log).getLastRow() - 1, 0);
  CacheService.getScriptCache().put('logLen', String(n), 21600);
  return n;
}
function appendLog_(rows) {
  if (!rows.length) return;
  var sh = sheet_(SHEET.log, HEAD.log);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEAD.log.length).setValues(rows);
  CacheService.getScriptCache().put('logLen', String(sh.getLastRow() - 1), 21600);
}
/** 1-기준 이력 번호 from..to (포함) */
function readLog_(from, to) {
  if (to < from) return [];
  return sheet_(SHEET.log, HEAD.log).getRange(from + 1, 1, to - from + 1, HEAD.log.length).getValues().map(function (r) {
    return { at: String(r[0]), actorId: String(r[1]), actor: String(r[2]), personId: String(r[3]), target: String(r[4]),
      date: cellDate_(r[5]), before: String(r[6]), after: String(r[7]) };
  });
}

/* ---------- 구조 버전 (명단·근태목록·대량 변경 시 화면 전체 새로고침) ---------- */
function structVersion_() { return +(PropertiesService.getScriptProperties().getProperty('STRUCT_VERSION') || 1); }
function bumpStruct_() { PropertiesService.getScriptProperties().setProperty('STRUCT_VERSION', String(structVersion_() + 1)); }

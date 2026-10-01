/**
 * 웹앱 진입점과 화면(Index.html)에서 google.script.run 으로 부르는 함수들.
 * 웹앱은 "배포한 사람 권한으로 실행"되므로 팀원은 시트에 직접 접근하지 않는다.
 */

var TOKEN_DAYS = 60;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(getSetting_('제목', '근태계획'))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** 화면에 보여줄 두 달 (이번 달 N, 다음 달 N+1) */
function visibleMonths_() {
  var n = today_().slice(0, 7);
  return [monthView_(n), monthView_(addMonths_(n, 1))];
}

function publicPerson_(p) {
  return { id: p.id, team: p.team, org: p.org, g1: p.g1, g2: p.g2, g3: p.g3, name: p.name, admin: p.admin };
}

function getBootstrap() {
  return {
    title: String(getSetting_('제목', '근태계획')),
    codes: getCodes_(),
    people: activePeople_().map(publicPerson_),
    months: visibleMonths_(),
    today: today_(),
  };
}

/** 두 달 화면에 필요한 모든 날짜의 일정. version 은 변경이력 개수(폴링 기준점). */
function getSchedule() {
  var version = logLength_();           // 데이터보다 먼저 읽어야 사이에 생긴 변경을 놓치지 않는다
  var sv = structVersion_();
  var keys = [];
  visibleMonths_().forEach(function (v) {
    monthsInView_(v).forEach(function (k) { if (keys.indexOf(k) < 0) keys.push(k); });
  });
  var cells = {}, generated = {};
  keys.forEach(function (k) {
    var m = readMonth_(k);
    if (!m) return;
    var p = parseMonthKey_(k);
    for (var d = 1; d <= m.days; d++) generated[ymd_(p.y, p.m, d)] = true;
    Object.keys(m.rows).forEach(function (id) {
      var c = cells[id] = cells[id] || {};
      m.rows[id].values.forEach(function (v, i) { if (v) c[ymd_(p.y, p.m, i + 1)] = v; });
    });
  });
  return { cells: cells, generated: generated, version: version, sv: sv };
}

/** 폴링: 마지막으로 받은 이후의 변경만 돌려준다. */
function getChanges(version, sv) {
  if (+sv !== structVersion_()) return { reload: true };
  var len = logLength_();
  if (len <= version) return { changes: [], version: len };
  if (len - version > 300) return { reload: true };
  var changes = readLog_(version + 1, len).map(function (r) {
    return { personId: r.personId, date: r.date, code: r.after };
  });
  return { changes: changes, version: len };
}

/**
 * 근태 한 칸 변경.
 * req: { actorId, token, personId, date, code, prev }
 * prev 가 현재 값과 다르면 덮어쓰지 않고 { conflict, current } 를 돌려준다.
 */
function setCell(req) {
  var people = allPeople_();
  var actor = findActive_(people, req.actorId), target = findActive_(people, req.personId);
  if (req.actorId !== req.personId) requireAdmin_(actor, req.token);
  var code = String(req.code || '');
  if (code && !getCodes_().some(function (c) { return c.code === code && c.active; })) throw new Error('없는 근태입니다: ' + code);
  var date = String(req.date), key = monthOfDate_(date);
  var sh = monthSheet_(key);
  if (!sh) throw new Error(key + ' 일정이 아직 생성되지 않았습니다.');

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var col = parseYmd_(date).d + 2;
    var row = findMonthRow_(sh, target.id);
    if (!row) {
      appendMonthRow_(key, target.id, target.name, defaultMonthValues_(key, ''));
      row = sh.getLastRow();
    }
    var cell = sh.getRange(row, col);
    var cur = cell.getValue();
    cur = cur == null ? '' : String(cur);
    if (cur !== String(req.prev || '')) return { conflict: true, current: cur };
    if (cur === code) return { ok: true };
    cell.setValue(code);
    appendLog_([[now_(), actor.id, actor.name, target.id, target.name, date, cur, code]]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function findMonthRow_(sh, id) {
  var n = sh.getLastRow();
  if (n < 2) return 0;
  var ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return 0;
}

function findActive_(people, id) {
  var p = people.filter(function (x) { return x.id === id && x.active; })[0];
  if (!p) throw new Error('명단에 없는 사람입니다. 화면을 새로고침해 주세요.');
  return p;
}

/* ---------- 관리자 인증 ---------- */
function hashPw_(pw) {
  var props = PropertiesService.getScriptProperties();
  var salt = props.getProperty('PW_SALT');
  if (!salt) { salt = Utilities.getUuid(); props.setProperty('PW_SALT', salt); }
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pw, Utilities.Charset.UTF_8));
}

function adminLogin(id, pw) {
  var p = allPeople_().filter(function (x) { return x.id === id && x.active; })[0];
  Utilities.sleep(300); // 무차별 대입 완화
  if (!p || !p.admin || !p.pwHash || hashPw_(String(pw || '')) !== p.pwHash) throw new Error('비밀번호가 맞지 않습니다.');
  var token = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty('TOK_' + token, id + '|' + (Date.now() + TOKEN_DAYS * 86400000));
  return token;
}

function checkAdminToken(id, token) {
  var p = allPeople_().filter(function (x) { return x.id === id && x.active; })[0];
  return !!(p && p.admin && tokenValid_(id, token));
}

function tokenValid_(id, token) {
  if (!token) return false;
  var v = PropertiesService.getScriptProperties().getProperty('TOK_' + token);
  if (!v) return false;
  var parts = v.split('|');
  return parts[0] === id && +parts[1] > Date.now();
}

function requireAdmin_(actor, token) {
  if (!actor.admin || !tokenValid_(actor.id, token)) throw new Error('관리자만 할 수 있습니다. 관리자로 다시 접속해 주세요.');
}

function adminActor_(actorId, token) {
  var actor = findActive_(allPeople_(), actorId);
  requireAdmin_(actor, token);
  return actor;
}

function cleanupTokens_() {
  var props = PropertiesService.getScriptProperties(), all = props.getProperties();
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('TOK_') === 0 && +all[k].split('|')[1] < Date.now()) props.deleteProperty(k);
  });
}

/* ---------- 관리자: 인원 ---------- */
function savePerson(actorId, token, data) {
  adminActor_(actorId, token);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var people = activePeople_();
    var name = String(data.name || '').trim();
    if (!name) throw new Error('성명을 입력하세요.');
    var fields = { name: name, team: str_(data.team), org: str_(data.org), g1: str_(data.g1), g2: str_(data.g2), g3: str_(data.g3),
      email: str_(data.email) };
    var p;
    if (data.id) {
      p = people.filter(function (x) { return x.id === data.id; })[0];
      if (!p) throw new Error('명단에 없는 사람입니다.');
      if (p.id === actorId && !data.admin) throw new Error('본인의 관리자 권한은 해제할 수 없습니다. 다른 관리자에게 요청하세요.');
      Object.keys(fields).forEach(function (k) { p[k] = fields[k]; });
    } else {
      p = fields; p.id = newId_(); p.active = true; p.order = insertOrder_(people, fields);
    }
    p.admin = !!data.admin;
    if (!p.admin) p.pwHash = '';
    else if (data.password) p.pwHash = hashPw_(String(data.password));
    else if (!p.pwHash) throw new Error('관리자 비밀번호를 입력하세요.');
    writePerson_(p);
    if (!data.id) {
      renumber_();
      // 이번 달부터 생성된 달에는 기본 근태(주간, 일요일 빈칸)로 행 추가
      var cur = today_().slice(0, 7);
      existingMonths_().filter(function (k) { return k >= cur; }).forEach(function (k) {
        appendMonthRow_(k, p.id, p.name, defaultMonthValues_(k, DEFAULT_CODE));
      });
    }
    bumpStruct_();
    return p.id;
  } finally { lock.releaseLock(); }
}

/** 관리자 편집 화면용 (공개 명단에는 없는 이메일) */
function getPersonDetail(actorId, token, id) {
  adminActor_(actorId, token);
  var p = activePeople_().filter(function (x) { return x.id === id; })[0];
  return p ? { email: p.email } : {};
}

function str_(v) { return String(v == null ? '' : v).trim(); }

/** 같은 소속·구분의 마지막 사람 바로 뒤 순서 (없으면 맨 끝) */
function insertOrder_(people, f) {
  var best = -1;
  [['org', 'g1', 'g2', 'g3'], ['org', 'g1', 'g2'], ['org', 'g1'], ['org']].some(function (keys) {
    people.forEach(function (x) {
      if (keys.every(function (k) { return x[k] === f[k]; })) best = Math.max(best, x.order);
    });
    return best >= 0;
  });
  var max = people.reduce(function (m, x) { return Math.max(m, x.order); }, 0);
  return best >= 0 ? best + 0.5 : max + 1;
}

/** 순서를 1, 2, 3… 으로 다시 매긴다 (제외된 사람은 맨 뒤) */
function renumber_() {
  var all = allPeople_();
  var sorted = all.filter(function (p) { return p.active; }).concat(all.filter(function (p) { return !p.active; }));
  var sh = sheet_(SHEET.people, HEAD.people), col = HEAD.people.indexOf('순서') + 1;
  sorted.forEach(function (p, i) { sh.getRange(p._row, col).setValue(i + 1); });
}

function movePerson(actorId, token, id, dir) {
  adminActor_(actorId, token);
  var people = activePeople_();
  var i = people.map(function (p) { return p.id; }).indexOf(id), j = i + (dir < 0 ? -1 : 1);
  if (i < 0 || j < 0 || j >= people.length) return;
  var sh = sheet_(SHEET.people, HEAD.people), col = HEAD.people.indexOf('순서') + 1;
  sh.getRange(people[i]._row, col).setValue(people[j].order);
  sh.getRange(people[j]._row, col).setValue(people[i].order);
  bumpStruct_();
}

function removePerson(actorId, token, id) {
  adminActor_(actorId, token);
  if (id === actorId) throw new Error('본인은 뺄 수 없습니다.');
  var p = activePeople_().filter(function (x) { return x.id === id; })[0];
  if (!p) throw new Error('명단에 없는 사람입니다.');
  p.active = false; p.admin = false; p.pwHash = ''; p.removedAt = today_();
  writePerson_(p);
  renumber_();
  bumpStruct_();
}

/* ---------- 관리자: 근태 목록 ---------- */
function saveCode(actorId, token, data) {
  adminActor_(actorId, token);
  var code = String(data.code || '').trim();
  if (!code) throw new Error('표기를 입력하세요.');
  var codes = getCodes_();
  var c = codes.filter(function (x) { return x.code === code; })[0];
  if (!c) { c = { code: code }; codes.push(c); }
  c.bg = data.bg || '#ffffff'; c.fg = data.fg || '#000000'; c.work = !!data.work; c.base = str_(data.base);
  c.group = str_(data.group); c.primary = !!data.primary; c.active = data.active !== false;
  writeCodes_(codes);
  bumpStruct_();
}

/* ---------- 관리자: 변경 이력 ---------- */
function getHistory(actorId, token, limit) {
  adminActor_(actorId, token);
  var len = logLength_(), from = Math.max(1, len - (limit || 100) + 1);
  return readLog_(from, len).reverse();
}

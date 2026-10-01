/**
 * 기존 엑셀 근무표(첨부 양식)에서 명단·조직 구조·일정 가져오기.
 * 사용법: 이 스프레드시트에서 [파일 > 가져오기 > 업로드 > 새 시트 삽입]으로 엑셀을 올린 뒤
 *         [근태관리 > 엑셀 양식에서 가져오기] 실행.
 * 같은 사람(성명+팀)은 기존 ID·관리자 설정을 유지하므로 여러 번 가져와도 안전하다.
 */

function importFromTemplateMenu() {
  var ui = SpreadsheetApp.getUi();
  var candidates = ss_().getSheets().filter(function (s) {
    return String(s.getRange('A3').getValue()).trim() === 'No.' && String(s.getRange('F3').getValue()).trim() === '성명';
  });
  if (!candidates.length) {
    ui.alert('가져올 근무표 탭을 찾지 못했습니다.\n\n[파일 > 가져오기 > 업로드]에서 엑셀 파일을 고르고 "새 시트 삽입"을 선택한 뒤 다시 실행하세요.');
    return;
  }
  var sh = candidates[0];
  if (candidates.length > 1) {
    var names = candidates.map(function (s) { return s.getName(); });
    var res = ui.prompt('어느 탭을 가져올까요?', names.join(', '), ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) return;
    sh = ss_().getSheetByName(res.getResponseText().trim());
    if (!sh) { ui.alert('탭 이름이 맞지 않습니다.'); return; }
  }
  var parsed = readTemplateSheet_(sh);
  var ok = ui.alert('가져오기 확인',
    '"' + sh.getName() + '" 탭에서 ' + parsed.people.length + '명, ' + parsed.monthKey + ' 일정(' +
    parsed.dates[0] + ' ~ ' + parsed.dates[parsed.dates.length - 1] + ')을 가져옵니다.\n\n' +
    '현재 명단은 이 표의 명단으로 바뀝니다 (같은 성명·팀은 관리자 설정 유지). 계속할까요?', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  var result = importTemplate_(parsed, sh);
  ui.alert('가져오기 완료',
    result.people + '명, ' + result.cells + '칸을 가져왔습니다.\n새로 만든 일정 탭: ' + (result.made.join(', ') || '없음') +
    '\n\n다음 단계: [근태관리 > 관리자 지정]으로 관리자를 정하세요. 가져온 원본 탭("' + sh.getName() + '")은 지워도 됩니다.', ui.ButtonSet.OK);
}

function readTemplateSheet_(sh) {
  var range = sh.getDataRange();
  var merged = range.getMergedRanges().map(function (r) {
    return { row: r.getRow(), col: r.getColumn(), nrows: r.getNumRows(), ncols: r.getNumColumns() };
  });
  return parseTemplate_(range.getValues(), merged);
}

function importTemplate_(parsed, sh) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // 1) 명단: 성명+팀이 같으면 기존 사람 유지
    var existing = allPeople_(), used = {};
    var keyOf = function (p) { return p.name + '|' + p.team; };
    var people = parsed.people.map(function (x, i) {
      var old = existing.filter(function (e) { return keyOf(e) === keyOf(x) && !used[e.id]; })[0];
      if (old) used[old.id] = true;
      return {
        id: old ? old.id : newId_(), team: x.team, org: x.org, g1: x.g1, g2: x.g2, g3: x.g3, name: x.name, order: i + 1,
        admin: old ? old.admin : false, email: old ? old.email : '', pwHash: old ? old.pwHash : '', active: true, removedAt: '',
        values: x.values,
      };
    });
    var today = today_();
    existing.forEach(function (e) {
      if (used[e.id]) return;
      e.active = false; e.admin = false; e.pwHash = ''; e.removedAt = e.removedAt || today; e.order = people.length + e.order;
      people.push(e);
    });
    var psh = sheet_(SHEET.people, HEAD.people);
    if (psh.getLastRow() > 1) psh.getRange(2, 1, psh.getLastRow() - 1, HEAD.people.length).clearContent();
    psh.getRange(2, 1, people.length, HEAD.people.length).setValues(people.map(function (p) {
      return [p.id, p.team, p.org, p.g1, p.g2, p.g3, p.name, p.order, yn_(p.admin), p.email, p.pwHash, yn_(p.active), p.removedAt || ''];
    }));
    var imported = people.filter(function (p) { return p.values; });

    // 2) 조직 표기, 제목
    writeOrgFormats_(parsed.orgFormats);
    setSetting_('제목', parsed.title);

    // 3) 근태 목록에 없는 표기는 원본 칸 색 그대로 추가
    var codes = getCodes_(), known = {};
    codes.forEach(function (c) { known[c.code] = true; });
    var bgs = sh.getRange(5, 7, imported.length, parsed.dates.length).getBackgrounds();
    var fgs = sh.getRange(5, 7, imported.length, parsed.dates.length).getFontColors();
    imported.forEach(function (p, i) {
      p.values.forEach(function (v, j) {
        if (v && !known[v]) {
          known[v] = true;
          codes.push({ code: v, bg: bgs[i][j] || '#ffffff', fg: fgs[i][j] || '#000000', work: false, base: '', group: '', primary: false, active: true });
        }
      });
    });
    writeCodes_(codes);

    // 4) 일정: 표의 날짜를 달별로 나눠서 기록. 다음 달은 기본값 생성 후 덮어쓴다.
    var byMonth = {};
    parsed.dates.forEach(function (d, j) { (byMonth[monthOfDate_(d)] = byMonth[monthOfDate_(d)] || []).push(j); });
    var cells = 0;
    var write = function (k) {
      var p = parseMonthKey_(k), days = daysInMonth_(p.y, p.m), m = readMonth_(k);
      if (!m) {
        createMonthSheet_(k, imported.map(function (x) { return { id: x.id, name: x.name, values: new Array(days).fill('') }; }));
        m = readMonth_(k);
      }
      var msh = monthSheet_(k);
      imported.forEach(function (x) {
        var row = m.rows[x.id];
        var vals = row ? row.values.slice() : new Array(days).fill('');
        byMonth[k].forEach(function (j) { vals[parseYmd_(parsed.dates[j]).d - 1] = x.values[j]; cells++; });
        if (row) msh.getRange(row.row, 3, 1, days).setValues([vals]);
        else msh.appendRow([x.id, x.name].concat(vals));
      });
    };
    Object.keys(byMonth).filter(function (k) { return k <= parsed.monthKey; }).forEach(write);
  } finally { lock.releaseLock(); }

  var made = ensureMonths_();
  Object.keys(byMonth).filter(function (k) { return k > parsed.monthKey && monthSheet_(k); }).forEach(write);
  bumpStruct_();
  return { people: imported.length, cells: cells, made: made };
}

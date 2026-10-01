/**
 * 자동 실행 작업 (트리거는 Setup.gs 의 installTriggers_ 가 설치)
 *   dailyJob          매일 새벽 1시: 이번 달·다음 달 일정 확인, 21일 이후면 다다음 달 생성
 *   sendDailySummary  매일 15시: 관리자에게 변경 요약 메일
 */

var GENERATE_DAY = 21;

function dailyJob() {
  ensureMonths_();
  cleanupTokens_();
}

/** 이번 달(N)·다음 달(N+1) 일정이 없으면 만들고, 21일부터는 다다음 달(N+2)도 만든다. */
function ensureMonths_() {
  var today = today_(), n = today.slice(0, 7);
  var keys = [n, addMonths_(n, 1)];
  if (parseYmd_(today).d >= GENERATE_DAY) keys.push(addMonths_(n, 2));
  var made = [];
  keys.forEach(function (k) { if (!monthSheet_(k) && generateMonth_(k)) made.push(k); });
  return made;
}

/** 월 탭 생성: 사람마다 전 달 가장 많았던 근무형태로 월~토 채움, 일요일 빈칸. */
function generateMonth_(key) {
  var people = activePeople_();
  if (!people.length) return false;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (monthSheet_(key)) return false;
    var codes = getCodes_();
    var prevKey = addMonths_(key, -1), prev = readMonth_(prevKey);
    var pp = parseMonthKey_(prevKey), prevDates = [];
    for (var d = 1; d <= daysInMonth_(pp.y, pp.m); d++) prevDates.push(ymd_(pp.y, pp.m, d));
    var rows = people.map(function (p) {
      var pr = prev && prev.rows[p.id];
      var code = pr ? defaultCodeFor_(pr.values, prevDates, codes) : DEFAULT_CODE;
      return { id: p.id, name: p.name, values: defaultMonthValues_(key, code) };
    });
    createMonthSheet_(key, rows);
    bumpStruct_();
    return true;
  } finally { lock.releaseLock(); }
}

/** 관리자 이메일 (없으면 시트 소유자) */
function adminEmails_() {
  var emails = activePeople_().filter(function (p) { return p.admin && p.email; }).map(function (p) { return p.email; });
  if (!emails.length) {
    var owner = Session.getEffectiveUser().getEmail();
    if (owner) emails.push(owner);
  }
  return emails;
}

function sendDailySummary() {
  var props = PropertiesService.getScriptProperties();
  var from = +(props.getProperty('MAIL_LOG_POS') || 0) + 1, to = logLength_();
  if (to < from) return 0;
  var rows = readLog_(from, to);
  var emails = adminEmails_();
  if (!emails.length) return 0;
  var title = String(getSetting_('제목', '근태계획'));
  var esc = function (s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/\n/g, ''); };
  var md = function (s) { var p = parseYmd_(s); return p.m + '/' + p.d + '(' + WEEKDAYS[weekdayIdx_(s)] + ')'; };
  rows.sort(function (a, b) { return a.target < b.target ? -1 : a.target > b.target ? 1 : a.date < b.date ? -1 : 1; });
  var td = 'style="border:1px solid #ccc;padding:4px 8px"';
  var html = '<p>' + esc(title) + ' 근태계획 변경 ' + rows.length + '건 (' + esc(rows[0].at) + ' ~ ' + esc(rows[rows.length - 1].at) + ')</p>' +
    '<table style="border-collapse:collapse;font-size:13px"><tr><th ' + td + '>대상자</th><th ' + td + '>날짜</th><th ' + td + '>변경</th><th ' + td + '>수정한 사람</th><th ' + td + '>시각</th></tr>' +
    rows.map(function (r) {
      return '<tr><td ' + td + '>' + esc(r.target) + '</td><td ' + td + '>' + md(r.date) + '</td><td ' + td + '>' +
        esc(r.before || '빈칸') + ' → <b>' + esc(r.after || '빈칸') + '</b></td><td ' + td + '>' + esc(r.actor) +
        (r.actorId !== r.personId ? ' (대리)' : '') + '</td><td ' + td + '>' + esc(r.at.slice(5, 16)) + '</td></tr>';
    }).join('') + '</table>' +
    (ScriptApp.getService().getUrl() ? '<p><a href="' + ScriptApp.getService().getUrl() + '">근태계획 열기</a></p>' : '');
  MailApp.sendEmail({
    to: emails.join(','),
    subject: '[' + title + '] 근태 변경 요약 ' + rows.length + '건 (' + Utilities.formatDate(new Date(), TZ, 'M/d HH:mm') + ')',
    htmlBody: html,
  });
  props.setProperty('MAIL_LOG_POS', String(to));
  return rows.length;
}

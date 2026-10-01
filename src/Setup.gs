/**
 * 스프레드시트 메뉴와 처음 설정.
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('근태관리')
    .addItem('1. 처음 설정 (탭·자동실행 만들기)', 'setupMenu')
    .addItem('2. 엑셀 양식에서 가져오기', 'importFromTemplateMenu')
    .addItem('3. 관리자 지정 / 비밀번호 변경', 'setAdminMenu')
    .addSeparator()
    .addItem('이번 달·다음 달 일정 지금 확인/생성', 'ensureMonthsMenu')
    .addItem('변경 요약 메일 지금 보내기', 'sendSummaryMenu')
    .addItem('웹앱 주소 보기', 'showUrlMenu')
    .addToUi();
}

function setupMenu() {
  setup_();
  SpreadsheetApp.getUi().alert('처음 설정 완료',
    '탭과 자동 실행(매일 새벽 1시 일정 생성, 매일 15시 요약 메일)을 만들었습니다.\n\n' +
    '다음 단계: [근태관리 > 엑셀 양식에서 가져오기] 또는 [인원] 탭에 명단 입력', SpreadsheetApp.getUi().ButtonSet.OK);
}

function setup_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), props = PropertiesService.getScriptProperties();
  props.setProperty('SPREADSHEET_ID', ss.getId());
  if (!props.getProperty('PW_SALT')) props.setProperty('PW_SALT', Utilities.getUuid());
  sheet_(SHEET.people, HEAD.people);
  var codeSheet = sheet_(SHEET.codes, HEAD.codes);
  if (codeSheet.getLastRow() < 2) writeCodes_(DEFAULT_CODES);
  sheet_(SHEET.org, HEAD.org);
  sheet_(SHEET.settings, HEAD.settings);
  if (getSetting_('제목', '') === '') setSetting_('제목', '근태계획');
  sheet_(SHEET.log, HEAD.log);
  if (!props.getProperty('MAIL_LOG_POS')) props.setProperty('MAIL_LOG_POS', String(logLength_()));
  installTriggers_();
  ensureMonths_();
  bumpStruct_();
}

function installTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (['dailyJob', 'sendDailySummary'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(1).inTimezone(TZ).create();
  ScriptApp.newTrigger('sendDailySummary').timeBased().everyDays(1).atHour(15).nearMinute(0).inTimezone(TZ).create();
}

function setAdminMenu() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('관리자 지정', '관리자로 지정할 사람의 성명을 입력하세요.', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var name = res.getResponseText().trim();
  var found = activePeople_().filter(function (p) { return p.name === name; });
  if (!found.length) { ui.alert('"' + name + '"님을 명단에서 찾지 못했습니다.'); return; }
  var p = found[0];
  if (found.length > 1) {
    var list = found.map(function (x, i) { return (i + 1) + ') ' + [x.org, x.g1, x.g2, x.g3].filter(String).join(' > '); }).join('\n');
    var pick = ui.prompt('같은 이름이 여러 명입니다', list + '\n\n번호를 입력하세요.', ui.ButtonSet.OK_CANCEL);
    if (pick.getSelectedButton() !== ui.Button.OK) return;
    p = found[+pick.getResponseText().trim() - 1];
    if (!p) { ui.alert('번호가 맞지 않습니다.'); return; }
  }
  var pw = ui.prompt(p.name + ' 관리자 비밀번호', '이 사람이 관리자로 접속할 때 쓸 비밀번호 (4자 이상)', ui.ButtonSet.OK_CANCEL);
  if (pw.getSelectedButton() !== ui.Button.OK) return;
  if (pw.getResponseText().length < 4) { ui.alert('4자 이상 입력하세요.'); return; }
  var em = ui.prompt(p.name + ' 이메일', '매일 15시 변경 요약 메일을 받을 주소 (비워 두면 받지 않음)', ui.ButtonSet.OK_CANCEL);
  if (em.getSelectedButton() !== ui.Button.OK) return;
  p.admin = true;
  p.pwHash = hashPw_(pw.getResponseText());
  if (em.getResponseText().trim()) p.email = em.getResponseText().trim();
  writePerson_(p);
  bumpStruct_();
  ui.alert(p.name + '님을 관리자로 지정했습니다.');
}

function ensureMonthsMenu() {
  var made = ensureMonths_();
  SpreadsheetApp.getUi().alert(made.length ? '새로 만든 일정: ' + made.join(', ') : '이미 모두 만들어져 있습니다.');
}

function sendSummaryMenu() {
  var n = sendDailySummary();
  SpreadsheetApp.getUi().alert(n ? n + '건의 변경을 메일로 보냈습니다.' : '보낼 변경이 없습니다.');
}

function showUrlMenu() {
  var url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(url ? '웹앱 주소 (팀원에게 공유하세요):\n\n' + url
    : '아직 웹앱으로 배포하지 않았습니다.\nApps Script 편집기에서 [배포 > 새 배포 > 웹 앱]을 실행하세요.');
}

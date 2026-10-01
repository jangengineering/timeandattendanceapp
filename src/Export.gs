/**
 * 첨부 엑셀 양식으로 내려받기.
 * 임시 구글시트에 양식을 그린 뒤 xlsx 로 변환하고, 임시 시트는 지운다.
 * 누구나 내려받을 수 있다 (명단에 있는 사람이면).
 */

var EXPORT_FOLDER = '근태계획_추출파일';

function exportMonth(actorId, token, key) {
  findActive_(allPeople_(), actorId);
  var view = monthView_(key);
  // 표 기간 중에 빠진 사람은 그 달 표에 남긴다
  var people = allPeople_().filter(function (p) { return p.active || (p.removedAt && p.removedAt > view.start); });
  var dates = viewDates_(view);
  var months = {};
  monthsInView_(view).forEach(function (k) { months[k] = readMonth_(k); });
  var values = people.map(function (p) {
    return dates.map(function (d) {
      var m = months[monthOfDate_(d)], r = m && m.rows[p.id];
      return r ? (r.values[parseYmd_(d).d - 1] || '') : '';
    });
  });
  var title = String(getSetting_('제목', '근태계획'));
  var lay = buildExportLayout_({ title: title, view: view, codes: getCodes_(), people: people, values: values, orgFormats: getOrgFormats_() });

  var yy = String(view.year).slice(2);
  var filename = title.replace(/[\\/:*?"<>|]/g, '_') + ' ' + yy + '년 ' + view.month + '월 근태계획_' +
    Utilities.formatDate(new Date(), TZ, 'yyMMdd_HHmm') + '.xlsx';

  var tmp = SpreadsheetApp.create('tmp_' + filename, lay.nrows, lay.ncols);
  try {
    var sh = tmp.getSheets()[0];
    sh.setName(view.month + '월 근무표');
    writeLayout_(sh, lay);
    SpreadsheetApp.flush();
    var resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx', {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true,
    });
    if (resp.getResponseCode() !== 200) throw new Error('엑셀 변환 실패 (' + resp.getResponseCode() + ')');
    var blob = resp.getBlob().setName(filename);
    return { filename: filename, base64: Utilities.base64Encode(blob.getBytes()), url: saveExportFile_(blob) };
  } finally {
    DriveApp.getFileById(tmp.getId()).setTrashed(true);
  }
}

/** 휴대폰 메신저 브라우저처럼 바로 다운로드가 안 되는 경우를 위한 링크. 하루 지난 파일은 지운다. */
function saveExportFile_(blob) {
  try {
    var it = DriveApp.getFoldersByName(EXPORT_FOLDER);
    var folder = it.hasNext() ? it.next() : DriveApp.createFolder(EXPORT_FOLDER);
    var old = folder.getFiles(), cutoff = Date.now() - 86400000;
    while (old.hasNext()) { var f = old.next(); if (f.getDateCreated().getTime() < cutoff) f.setTrashed(true); }
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/uc?export=download&id=' + file.getId();
  } catch (e) {
    return ''; // 회사 정책으로 링크 공유가 막힌 경우 등
  }
}

function pxWidth_(excelWidth) { return Math.round(excelWidth * 7 + 5); }

function writeLayout_(sh, lay) {
  var NR = lay.nrows, NC = lay.ncols;
  var all = sh.getRange(1, 1, NR, NC);
  all.setFontFamily('맑은 고딕').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  // 숫자처럼 보이는 문자열('05')은 앞의 ' 로 텍스트 유지
  all.setValues(lay.cells);
  all.setBackgrounds(lay.bg);
  all.setFontColors(lay.fg.map(function (r) { return r.map(function (c) { return c || '#000000'; }); }));
  all.setFontWeights(lay.bold.map(function (r) { return r.map(function (b) { return b ? 'bold' : 'normal'; }); }));
  sh.getRange(1, 1).setFontSize(35);
  [lay.tableRange, lay.countRange, lay.summaryRange, lay.timeRange].forEach(function (b) {
    sh.getRange(b[0], b[1], b[2], b[3]).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  });
  lay.merges.forEach(function (m) { sh.getRange(m[0], m[1], m[2], m[3]).merge(); });
  Object.keys(lay.colWidths).forEach(function (c) { sh.setColumnWidth(+c, pxWidth_(lay.colWidths[c])); });
  Object.keys(lay.rowHeights).forEach(function (r) { sh.setRowHeight(+r, Math.round(lay.rowHeights[r] * 4 / 3)); });
  // 구글시트는 병합된 제목(A1:마지막 날짜열)을 가르는 열 고정을 허용하지 않아 행만 고정한다.
  sh.setFrozenRows(lay.freeze[0]);
}

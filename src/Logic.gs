/**
 * 순수 로직 (시트/구글 서비스에 의존하지 않음) — Node 테스트(tools/test.mjs)에서도 그대로 불러 쓴다.
 * 날짜는 모두 'yyyy-MM-dd' 문자열, 월은 'yyyy-MM' 문자열로 다룬다.
 */

var WEEKDAYS = ['월', '화', '수', '목', '금', '토', '일'];

/** 엑셀 양식 오른쪽의 근태별 개수(COUNTIF) 열 머리글 — 첨부 양식 그대로. */
var COUNT_HEADERS = ['1Shift', '1Shift\n(OT)', '1Shift\n(전반)', '1Shift\n(후반)', '주간\n(OT)', '주간\n(연장)',
  '2Shift', '2Shift\n(OT)', '2Shift\n(조출)', '철야', 'OT'];

/** 하단 집계표의 근무형태 묶음 — 시간대별 근무인원 수식이 이 세 묶음을 쓴다. */
var SUMMARY_GROUPS = ['1Shift', '주간', '2Shift'];
var SUMMARY_GROUP_NOTE = { '1Shift': 'OT 포함', '주간': 'OT, 연장 포함', '2Shift': 'OT, 철야 포함' };
var SUMMARY_TOTAL_NOTE = { '1Shift': '', '주간': '주간 OT\n/연장 포함', '2Shift': '철야 포함' };

var DEFAULT_CODE = '주간';

/** 처음 설치할 때의 근태 목록. [표기, 배경, 글자, 근무형태, 기본형, 집계묶음, 자주씀] */
var DEFAULT_CODES = [
  ['주간', '#ffffff', '#000000', true, '', '주간', true],
  ['1Shift', '#ffccff', '#000000', true, '', '1Shift', true],
  ['2Shift', '#ffe699', '#000000', true, '', '2Shift', true],
  ['연차', '#ffffff', '#ff0000', false, '', '', true],
  ['반차', '#ffffff', '#ff0000', false, '', '', true],
  ['공가', '#ffffff', '#ff0000', false, '', '', true],
  ['휴무', '#fbe5d6', '#ff0000', false, '', '', true],
  ['교육', '#d0cece', '#000000', false, '', '', true],
  ['주간2', '#e2efda', '#000000', true, '', '주간', true],
  ['주간\n(OT)', '#ffffff', '#000000', true, '주간', '주간', false],
  ['주간\n(연장)', '#ffffff', '#000000', true, '주간', '주간', false],
  ['주간\n(반일교육)', '#ffffff', '#000000', true, '주간', '주간', false],
  ['1Shift\n(OT)', '#ffccff', '#000000', true, '1Shift', '1Shift', false],
  ['1Shift\n(전반)', '#ffccff', '#000000', true, '1Shift', '1Shift', false],
  ['1Shift\n(후반)', '#ffccff', '#000000', true, '1Shift', '1Shift', false],
  ['2Shift\n(OT)', '#ffe699', '#000000', true, '2Shift', '2Shift', false],
  ['2Shift\n(조출)', '#ffe699', '#000000', true, '2Shift', '2Shift', false],
  ['철야', '#ffe699', '#000000', true, '2Shift', '2Shift', false],
  ['검진', '#ffffff', '#ff0000', false, '', '', false],
  ['출장', '#d9d9d9', '#000000', false, '', '', false],
  ['힐링\n캠프', '#fbe5d6', '#ff0000', false, '', '', false],
].map(function (r) {
  return { code: r[0], bg: r[1], fg: r[2], work: r[3], base: r[4], group: r[5], primary: r[6], active: true };
});

var SUNDAY_BLANK_BG = '#bfbfbf';

/* ---------- 날짜 ---------- */
function pad2_(n) { return (n < 10 ? '0' : '') + n; }
function ymd_(y, m, d) { return y + '-' + pad2_(m) + '-' + pad2_(d); }
function parseYmd_(s) { var p = s.split('-'); return { y: +p[0], m: +p[1], d: +p[2] }; }
/** UTC 기준으로 계산해서 시간대/서머타임 영향을 받지 않게 한다. */
function toUtc_(s) { var p = parseYmd_(s); return Date.UTC(p.y, p.m - 1, p.d); }
function fromUtc_(t) { var d = new Date(t); return ymd_(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()); }
function addDays_(s, n) { return fromUtc_(toUtc_(s) + n * 86400000); }
/** 월=0 … 일=6 */
function weekdayIdx_(s) { return (new Date(toUtc_(s)).getUTCDay() + 6) % 7; }
function daysInMonth_(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }

function monthKey_(y, m) { return y + '-' + pad2_(m); }
function parseMonthKey_(k) { var p = k.split('-'); return { y: +p[0], m: +p[1] }; }
function addMonths_(k, n) {
  var p = parseMonthKey_(k), idx = p.y * 12 + (p.m - 1) + n;
  return monthKey_(Math.floor(idx / 12), idx % 12 + 1);
}
function monthOfDate_(s) { return s.slice(0, 7); }

/**
 * 한 달 근무표의 표시 범위: 1일이 속한 주의 월요일 ~ 말일이 속한 주의 일요일.
 * (첨부 양식의 10월표가 9/28(월) ~ 11/1(일)인 것과 같은 규칙)
 */
function monthView_(key) {
  var p = parseMonthKey_(key);
  var first = ymd_(p.y, p.m, 1), last = ymd_(p.y, p.m, daysInMonth_(p.y, p.m));
  var start = addDays_(first, -weekdayIdx_(first));
  var end = addDays_(last, 6 - weekdayIdx_(last));
  return { key: key, year: p.y, month: p.m, label: p.m + '월', start: start, end: end,
    days: Math.round((toUtc_(end) - toUtc_(start)) / 86400000) + 1 };
}
function viewDates_(view) {
  var out = [];
  for (var i = 0; i < view.days; i++) out.push(addDays_(view.start, i));
  return out;
}
/** 표시 범위에 걸친 달들 (앞 달, 그 달, 다음 달 중 해당되는 것) */
function monthsInView_(view) {
  var keys = [monthOfDate_(view.start), view.key, monthOfDate_(view.end)];
  return keys.filter(function (k, i) { return keys.indexOf(k) === i; });
}

/* ---------- 기본 근태 ---------- */
/**
 * 전 달 근무에서 가장 많았던 근무형태. 비근무(연차·휴무 등)는 빼고, 변형(OT 등)은 기본형으로 합산.
 * 동률이면 근태 목록 순서가 앞선 것. 알 수 없으면 '주간'.
 * @param {string[]} values 전 달 1일~말일 값
 * @param {string[]} dates  values 와 같은 길이의 날짜
 */
function defaultCodeFor_(values, dates, codes) {
  var byCode = {};
  codes.forEach(function (c) { byCode[c.code] = c; });
  var count = {};
  values.forEach(function (v, i) {
    if (!v || weekdayIdx_(dates[i]) === 6) return;
    var c = byCode[v];
    if (!c || !c.work) return;
    var key = c.base || c.code;
    count[key] = (count[key] || 0) + 1;
  });
  var best = '', bestN = 0;
  codes.forEach(function (c) {
    if ((count[c.code] || 0) > bestN) { best = c.code; bestN = count[c.code]; }
  });
  return best || DEFAULT_CODE;
}

/** 새 달 기본값: 월~토(공휴일 포함) = 기본 근태, 일요일 = 빈칸 */
function defaultMonthValues_(key, code) {
  var p = parseMonthKey_(key), out = [];
  for (var d = 1; d <= daysInMonth_(p.y, p.m); d++) out.push(weekdayIdx_(ymd_(p.y, p.m, d)) === 6 ? '' : code);
  return out;
}

/* ---------- 조직 표기 ---------- */
/** '건축팀\n(81)' → { name: '건축팀', sep: '\n' } / 'PC (3)' → { name: 'PC', sep: ' ' } / 'FAB' → { name: 'FAB', sep: null } */
function splitCount_(label) {
  var s = String(label == null ? '' : label);
  var m = s.match(/^([\s\S]*?)(\s*)\(\d+(?:\+\d+)*명?\)\s*$/);
  if (!m || !m[1].trim()) return { name: s.trim(), sep: null };
  return { name: m[1].trim(), sep: m[2] || '' };
}
function orgPathKey_(levels, upto) { return levels.slice(0, upto + 1).join('>'); }

/**
 * 소속·구분(4단계, B~E열) 병합 계산.
 * 빈 단계는 왼쪽 단계 칸이 가로로 덮고, 같은 경로가 연속된 행은 세로로 병합한다.
 * @param {string[][]} rows 사람별 [소속, 구분1, 구분2, 구분3]
 * @return {{row:number, col:number, nrows:number, ncols:number, path:string, name:string}[]} 0-기준 위치
 */
function orgBlocks_(rows) {
  var blocks = [];
  for (var L = 0; L < 4; L++) {
    var i = 0;
    while (i < rows.length) {
      var lv = rows[i];
      if (L > 0 && !lv[L]) { i++; continue; }
      var ext = 1;
      while (L + ext < 4 && !lv[L + ext]) ext++;
      var key = JSON.stringify(lv.slice(0, L + 1)) + '|' + ext;
      var j = i + 1;
      while (j < rows.length) {
        var lw = rows[j], ext2 = 1;
        while (L + ext2 < 4 && !lw[L + ext2]) ext2++;
        if (JSON.stringify(lw.slice(0, L + 1)) + '|' + ext2 !== key) break;
        j++;
      }
      blocks.push({ row: i, col: L, nrows: j - i, ncols: ext, path: orgPathKey_(lv, L), name: lv[L] });
      i = j;
    }
  }
  return blocks;
}

/** 표기 형식(sep)이 있으면 인원수를 붙인다. */
function orgLabel_(name, sep, n) {
  return sep == null ? name : name + sep + '(' + n + ')';
}

/* ---------- 엑셀 양식 레이아웃 ---------- */
function colName_(n) { // 1 → A
  var s = '';
  while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/**
 * 첨부 엑셀 양식과 같은 근무표를 만든다 (시트에 쓰기 전 단계의 순수 데이터).
 * @param {object} p
 *   title: 'P5 FAB 건축/안전팀', view: monthView_(), codes: [...],
 *   people: [{ team, org, g1, g2, g3, name }] (표시 순서대로),
 *   values: [[...]] 사람별 날짜값 (view 날짜 순), orgFormats: { path: sep }
 * @return {{ nrows, ncols, cells, bg, fg, bold, merges, colWidths, rowHeights, freeze, fontSize }}
 *   cells/bg/fg/bold 는 0-기준 2차원 배열, merges 는 [row, col, nrows, ncols] (1-기준)
 */
function buildExportLayout_(p) {
  var view = p.view, dates = viewDates_(view), D = dates.length, P = p.people.length;
  var FIRST = 7;                                // G열 = 첫 날짜
  var LASTD = FIRST + D - 1;                    // 마지막 날짜 열
  var GAP = LASTD + 1, CNT = GAP + 1;           // 빈 열, 개수 열 시작
  var NC = CNT + COUNT_HEADERS.length - 1;
  var PR = 5;                                   // 첫 사람 행
  var lastPR = PR + P - 1;
  var codeMap = {};
  p.codes.forEach(function (c) { codeMap[c.code] = c; });

  // 팀별 연속 구간 (집계 수식 범위)
  var teams = [], segs = {};
  p.people.forEach(function (x, i) {
    var t = x.team || '';
    if (teams.indexOf(t) < 0) { teams.push(t); segs[t] = []; }
    var s = segs[t], r = PR + i;
    if (s.length && s[s.length - 1][1] === r - 1) s[s.length - 1][1] = r; else s.push([r, r]);
  });

  // 집계 행 위치
  var groupsRows = SUMMARY_GROUPS.map(function (g) {
    return [g].concat(p.codes.filter(function (c) { return c.group === g && c.code !== g; }).map(function (c) { return c.code; }));
  });
  var perTeam = groupsRows.reduce(function (a, g) { return a + g.length; }, 0);
  var S0 = lastPR + 2;                          // 휴무/연차 행
  var T0 = S0 + Math.max(teams.length, 2);      // 팀별 집계 시작
  var TOT = T0 + perTeam * teams.length;        // 총원 집계 시작
  var TS = TOT + SUMMARY_GROUPS.length + 1;     // ■시간별 근무인원
  var NR = TS + 5;

  var cells = [], bg = [], fg = [], bold = [];
  for (var r = 0; r < NR; r++) {
    cells.push(new Array(NC).fill('')); bg.push(new Array(NC).fill(null));
    fg.push(new Array(NC).fill(null)); bold.push(new Array(NC).fill(false));
  }
  var merges = [];
  function set(r, c, v) { cells[r - 1][c - 1] = v; }
  function A(c, r, absC, absR) { return (absC ? '$' : '') + colName_(c) + (absR ? '$' : '') + r; }

  // 제목
  var yy = String(view.year).slice(2);
  set(1, 1, p.title + " '" + yy + '년 ' + view.month + '월 근태계획');
  merges.push([1, 1, 2, LASTD]);

  // 머리글
  set(3, 1, 'No.'); merges.push([3, 1, 2, 1]);
  set(3, 2, '소속'); merges.push([3, 2, 2, 1]);
  set(3, 3, '구분'); merges.push([3, 3, 2, 3]);
  set(3, 6, '성명'); merges.push([3, 6, 2, 1]);
  dates.forEach(function (d, i) {
    var wi = weekdayIdx_(d), c = FIRST + i;
    set(3, c, WEEKDAYS[wi]); set(4, c, parseYmd_(d).d);
    if (wi === 5) { fg[2][c - 1] = '#0000ff'; fg[3][c - 1] = '#0000ff'; }
    if (wi === 6) { fg[2][c - 1] = '#ff0000'; fg[3][c - 1] = '#ff0000'; }
  });
  COUNT_HEADERS.forEach(function (h, i) { set(3, CNT + i, h); merges.push([3, CNT + i, 2, 1]); });

  // 사람
  var orgRows = p.people.map(function (x) { return [x.org || '', x.g1 || '', x.g2 || '', x.g3 || '']; });
  orgBlocks_(orgRows).forEach(function (b) {
    var sep = p.orgFormats ? p.orgFormats[b.path] : null;
    set(PR + b.row, 2 + b.col, orgLabel_(b.name, sep === undefined ? null : sep, b.nrows));
    if (b.col === 0) bold[PR + b.row - 1][1] = true;
    if (b.nrows > 1 || b.ncols > 1) merges.push([PR + b.row, 2 + b.col, b.nrows, b.ncols]);
  });
  p.people.forEach(function (x, i) {
    var r = PR + i;
    set(r, 1, '=ROW()-4');
    set(r, 6, x.name);
    dates.forEach(function (d, j) {
      var v = (p.values[i] || [])[j] || '', c = FIRST + j, info = codeMap[v];
      set(r, c, v);
      bold[r - 1][c - 1] = true;
      if (!v && weekdayIdx_(d) === 6) bg[r - 1][c - 1] = SUNDAY_BLANK_BG;
      else if (info) { bg[r - 1][c - 1] = info.bg; fg[r - 1][c - 1] = info.fg; }
    });
    COUNT_HEADERS.forEach(function (h, k) {
      set(r, CNT + k, '=COUNTIF(' + A(FIRST, r, true, false) + ':' + A(LASTD, r, true, false) + ',' + A(CNT + k, 3, false, true) + ')');
    });
  });

  function countIf(c, team, crit) {
    return (segs[team] || []).map(function (s) {
      return 'COUNTIF(' + A(c, s[0], false, true) + ':' + A(c, s[1], false, true) + ',' + crit + ')';
    }).join('+') || '0';
  }

  // 휴무/연차 행
  set(S0, 4, '휴무'); set(S0 + 1, 4, '연차');
  teams.forEach(function (t, k) {
    var r = S0 + k;
    set(r, 6, t);
    for (var c = FIRST; c <= LASTD; c++) {
      set(r, c, '=' + countIf(c, t, A(4, S0, true, true)) + '+' + countIf(c, t, A(4, S0 + 1, true, true)));
    }
  });

  // 팀별 근태 인원
  var teamGroupRow = {}; // teamGroupRow[team][group] = 행
  var orgCounts = {};
  p.people.forEach(function (x) {
    var t = x.team || '', o = x.org || '';
    orgCounts[t] = orgCounts[t] || [];
    var last = orgCounts[t][orgCounts[t].length - 1];
    if (last && last.org === o) last.n++; else orgCounts[t].push({ org: o, n: 1 });
  });
  teams.forEach(function (t, k) {
    var r0 = T0 + k * perTeam;
    set(r0, 4, t + '\n(' + orgCounts[t].map(function (o) { return o.n; }).join('+') + '명)');
    merges.push([r0, 4, perTeam, 1]);
    teamGroupRow[t] = {};
    var r = r0;
    groupsRows.forEach(function (gr) {
      var gRow = r;
      teamGroupRow[t][gr[0]] = gRow;
      set(gRow, 5, SUMMARY_GROUP_NOTE[gr[0]] || '');
      gr.forEach(function (code, idx) {
        set(r, 6, code);
        for (var c = FIRST; c <= LASTD; c++) {
          var f = '=' + countIf(c, t, A(6, r, true, false));
          if (idx === 0) for (var v = 1; v < gr.length; v++) f += '+' + A(c, gRow + v);
          set(r, c, f);
        }
        r++;
      });
    });
  });

  // 총원
  set(TOT, 4, '총원\n(' + P + ')');
  merges.push([TOT, 4, SUMMARY_GROUPS.length, 1]);
  var totRow = {};
  SUMMARY_GROUPS.forEach(function (g, i) {
    var r = TOT + i;
    totRow[g] = r;
    set(r, 5, SUMMARY_TOTAL_NOTE[g] || '');
    set(r, 6, g);
    for (var c = FIRST; c <= LASTD; c++) {
      set(r, c, '=' + teams.map(function (t) { return A(c, teamGroupRow[t][g]); }).join('+'));
    }
  });

  // 시간별 근무인원 (첨부 양식 수식 그대로)
  set(TS, 3, '■시간별 근무인원');
  var slots = [["'05", "'07"], ["'07", '14'], ['14', '15'], ['15', '17'], ['17', '24']];
  slots.forEach(function (s, i) {
    var r = TS + 1 + i;
    set(r, 3, s[0]); set(r, 5, s[1]);
    for (var c = FIRST; c <= LASTD; c++) {
      var f;
      if (i === 0) f = '=' + A(c, totRow['1Shift']);
      else if (i === 1) f = '=' + A(c, r - 1) + '+' + A(c, totRow['주간']);
      else if (i === 2) f = '=' + A(c, r - 1) + '+' + A(c, totRow['2Shift']);
      else if (i === 3) f = '=' + A(c, r - 1) + '-' + A(c, TS + 1);
      else f = '=' + A(c, r - 1) + '-' + A(c, totRow['주간']);
      set(r, c, f);
    }
  });

  var colWidths = {};
  colWidths[1] = 4.375; colWidths[2] = 8.625; colWidths[3] = 8.625; colWidths[4] = 16.25; colWidths[5] = 8.625; colWidths[6] = 7.375;
  for (var c = FIRST; c <= LASTD; c++) colWidths[c] = 6.125;
  colWidths[GAP] = 3.625;
  for (c = CNT; c <= NC; c++) colWidths[c] = 7.5;
  var rowHeights = { 1: 30, 2: 30, 3: 25.15, 4: 25.15 };
  for (r = PR; r <= lastPR; r++) rowHeights[r] = 27;

  return {
    nrows: NR, ncols: NC, cells: cells, bg: bg, fg: fg, bold: bold, merges: merges,
    colWidths: colWidths, rowHeights: rowHeights, freeze: [4, 6],
    tableRange: [3, 1, lastPR - 2, LASTD], countRange: [3, CNT, lastPR - 2, COUNT_HEADERS.length],
    summaryRange: [S0, 4, TOT + SUMMARY_GROUPS.length - S0, LASTD - 3], timeRange: [TS + 1, 3, 5, LASTD - 2],
    firstDateCol: FIRST, lastDateCol: LASTD, firstPersonRow: PR, lastPersonRow: lastPR,
  };
}

/* ---------- 엑셀 양식 읽기 (가져오기) ---------- */
/**
 * 첨부 양식의 시트 값을 해석한다.
 * @param {any[][]} values 시트 전체 값 (병합 칸은 왼쪽 위만 값이 있음)
 * @param {{row:number,col:number,nrows:number,ncols:number}[]} merged 병합 범위 (1-기준)
 * @return {{ title, monthKey, dates, people:[{team,org,g1,g2,g3,name,values}], orgFormats }}
 */
function parseTemplate_(values, merged) {
  var get = function (r, c) { return (values[r - 1] || [])[c - 1]; };
  var str = function (v) { return v == null ? '' : String(v); };

  // 제목: "P5 FAB 건축/안전팀 '26년 10월 근태계획"
  var title = str(get(1, 1)).trim();
  var m = title.match(/^(.*?)\s*'?(\d{2,4})\s*년\s*(\d{1,2})\s*월/);
  if (!m) throw new Error('A1 제목에서 년/월을 찾지 못했습니다: ' + title);
  var year = +m[2] < 100 ? 2000 + +m[2] : +m[2];
  var key = monthKey_(year, +m[3]);
  var view = monthView_(key);

  // 날짜 열: G열부터 3행 요일이 있는 곳까지
  var dates = [];
  for (var c = 7; ; c++) {
    var wd = str(get(3, c)).trim();
    if (WEEKDAYS.indexOf(wd) < 0) break;
    dates.push(addDays_(view.start, c - 7));
  }
  if (!dates.length) throw new Error('3행에서 요일 머리글(월~일)을 찾지 못했습니다.');
  var d0 = +get(4, 7);
  if (d0 && d0 !== parseYmd_(view.start).d) throw new Error('G4 날짜(' + d0 + ')가 ' + key + ' 표의 첫 월요일(' + view.start + ')과 맞지 않습니다.');

  // 병합 칸: 세로로는 값을 채우고, 가로로 덮인 칸(빈 단계)은 빈칸으로 둔다.
  var filled = {};
  merged.forEach(function (g) {
    var v = get(g.row, g.col);
    for (var r = g.row; r < g.row + g.nrows; r++) {
      for (var cc = g.col; cc < g.col + g.ncols; cc++) filled[r + ',' + cc] = cc === g.col ? v : '';
    }
  });
  var cellOf = function (r, cc) { var k = r + ',' + cc; return k in filled ? filled[k] : get(r, cc); };

  var people = [], orgFormats = {};
  for (var r = 5; ; r++) {
    var name = str(get(r, 6)).trim();
    if (!name) break;
    var raw = [2, 3, 4, 5].map(function (cc) { return str(cellOf(r, cc)); });
    var lv = raw.map(function (x) { return splitCount_(x).name; });
    raw.forEach(function (x, i) {
      var sc = splitCount_(x);
      if (sc.name && sc.sep != null) orgFormats[orgPathKey_(lv, i)] = sc.sep;
    });
    var vals = dates.map(function (_, j) { return str(get(r, 7 + j)).trim(); });
    people.push({ team: /^안전/.test(lv[0]) ? '안전' : '건축', org: lv[0], g1: lv[1], g2: lv[2], g3: lv[3], name: name, values: vals });
  }
  if (!people.length) throw new Error('5행부터 성명(F열)을 찾지 못했습니다.');
  return { title: m[1].trim(), monthKey: key, dates: dates, people: people, orgFormats: orgFormats };
}

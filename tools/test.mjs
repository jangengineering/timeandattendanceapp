// 서버 코드 시나리오 테스트:  node tools/test.mjs [첨부양식.json]
// 첨부 양식 JSON(값+병합)이 있으면 가져오기/추출을 실제 양식으로 시험한다 (tools/sheet_to_json.py 로 생성).
import { readFileSync, readdirSync, existsSync } from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';
import { makeGas } from './fake-gas.mjs';

const SRC = new URL('../src/', import.meta.url);
const sampleJson = process.argv[2];

function boot(today) {
  const gas = makeGas({ today });
  const created = [];
  const ctx = {
    ...gas, Date: gas.Date, console,
    UrlFetchApp: { fetch: () => ({ getResponseCode: () => 200, getBlob: () => ({ setName() { return this; }, getBytes: () => [1, 2, 3] }) }) },
    DriveApp: {
      getFileById: () => ({ setTrashed() {} }), getFoldersByName: () => ({ hasNext: () => false }),
      createFolder: () => ({ getFiles: () => ({ hasNext: () => false }), createFile: () => ({ setSharing() {}, getId: () => 'F1' }) }),
      Access: {}, Permission: {},
    },
  };
  ctx.SpreadsheetApp = { ...gas.SpreadsheetApp, create: () => { const s = new gas.ss.constructor('TMP'); s.insertSheet('Sheet1'); created.push(s); return s; } };
  vm.createContext(ctx);
  for (const f of readdirSync(SRC).filter(f => f.endsWith('.gs')).sort()) vm.runInContext(readFileSync(new URL(f, SRC), 'utf8'), ctx, { filename: f });
  return { ctx, gas, created };
}

let passed = 0;
function test(name, fn) { fn(); passed++; console.log('  ✓ ' + name); }

/* ---------- 순수 로직 ---------- */
{
  const { ctx: L } = boot('2026-10-01');
  test('10월 표 범위: 9/28(월) ~ 11/1(일), 35일', () => {
    const v = L.monthView_('2026-10');
    assert.equal(v.start, '2026-09-28'); assert.equal(v.end, '2026-11-01'); assert.equal(v.days, 35);
  });
  test('11월 표 범위: 10/26 ~ 12/6, 42일 / 2월 경계', () => {
    const v = L.monthView_('2026-11');
    assert.equal(v.start, '2026-10-26'); assert.equal(v.end, '2026-12-06'); assert.equal(v.days, 42);
    assert.equal(L.addMonths_('2026-12', 1), '2027-01'); assert.equal(L.addMonths_('2027-01', -1), '2026-12');
    assert.equal(L.monthView_('2027-02').start, '2027-02-01');
  });
  test('기본 근태: 근무형태만, 변형은 기본형으로 합산, 일요일 제외', () => {
    const dates = Array.from({ length: 10 }, (_, i) => L.addDays_('2026-10-01', i));
    const vals = ['연차', '연차', '연차', '1Shift', '1Shift\n(OT)', '2Shift', '2Shift', '', '연차', '1Shift'];
    assert.equal(L.defaultCodeFor_(vals, dates, L.DEFAULT_CODES), '1Shift');
    assert.equal(L.defaultCodeFor_(['연차', '휴무', ''], dates.slice(0, 3), L.DEFAULT_CODES), '주간');
    assert.equal(L.defaultCodeFor_(['주간2', '주간2', '주간'], dates.slice(0, 3), L.DEFAULT_CODES), '주간2');
  });
  test('새 달 기본값: 월~토 채움, 일요일 빈칸', () => {
    const v = L.defaultMonthValues_('2026-11', '2Shift');
    assert.equal(v.length, 30); assert.equal(v[0], ''); /* 11/1 일 */ assert.equal(v[1], '2Shift'); assert.equal(v[7], '');
  });
  test('조직 표기 분리', () => {
    assert.deepEqual({ ...L.splitCount_('건축팀\n(81)') }, { name: '건축팀', sep: '\n' });
    assert.deepEqual({ ...L.splitCount_('PC (3)') }, { name: 'PC', sep: ' ' });
    assert.deepEqual({ ...L.splitCount_('공사팀장') }, { name: '공사팀장', sep: null });
  });
}

/* ---------- 시나리오 ---------- */
{
  const { ctx: S, gas, created } = boot('2026-10-01');
  S.setup_();
  test('처음 설정: 탭과 트리거 생성, 명단이 없으면 월 탭은 아직 없음', () => {
    for (const n of ['인원', '근태코드', '조직표기', '설정', '변경이력']) assert.ok(gas.ss.getSheetByName(n), n);
    assert.deepEqual([...gas.triggers].sort(), ['dailyJob', 'sendDailySummary']);
    assert.equal(S.existingMonths_().length, 0);
  });

  let parsed;
  if (sampleJson && existsSync(sampleJson)) {
    const { values, merged } = JSON.parse(readFileSync(sampleJson, 'utf8'));
    const sh = gas.ss.insertSheet('10월 근무표');
    values.forEach((row, i) => row.forEach((v, j) => { if (v != null) sh.set(i + 1, j + 1, v); }));
    merged.forEach(m => sh.merges.push([m.row, m.col, m.nrows, m.ncols]));
    parsed = S.readTemplateSheet_(sh);
    const res = S.importTemplate_(parsed, sh);
    test(`엑셀 가져오기: ${res.people}명, 9·10·11월 탭`, () => {
      assert.equal(res.people, parsed.people.length);
      assert.deepEqual([...S.existingMonths_()], ['2026-09', '2026-10', '2026-11']);
      assert.equal(S.getSetting_('제목'), parsed.title);
    });
  } else {
    // 첨부 양식이 없으면 작은 명단으로
    const sh = gas.ss.getSheetByName('인원');
    [['A', '건축'], ['B', '건축'], ['C', '안전']].forEach(([n, t], i) => sh.appendRow(['P' + i, t, t + '팀', '', '', '', n, i + 1, '', '', '', 'Y', '']));
    S.ensureMonths_();
  }

  const people = S.activePeople_();
  const a = people[0], b = people[1];
  const sched = S.getSchedule();
  test('일정 읽기: 두 달 화면 범위, 12월은 미생성', () => {
    if (parsed) assert.ok(sched.generated['2026-09-28']); assert.ok(sched.generated['2026-10-01']); assert.ok(sched.generated['2026-11-30']);
    assert.ok(!sched.generated['2026-12-01']);
    if (parsed) {
      const i = parsed.dates.indexOf('2026-10-05');
      assert.equal(sched.cells[a.id]['2026-10-05'], parsed.people[0].values[i]);
      assert.equal(sched.cells[a.id]['2026-09-28'] || '', parsed.people[0].values[0]);
      assert.equal(sched.cells[a.id]['2026-11-01'] || '', parsed.people[0].values[34]); // 11/1 은 엑셀 값으로 덮어씀
    }
  });
  test('11월 기본값: 10월 최다 근무형태, 일요일 빈칸', () => {
    const oct = S.readMonth_('2026-10').rows[b.id].values;
    const octDates = oct.map((_, i) => S.ymd_(2026, 10, i + 1));
    const expect = S.defaultCodeFor_(oct, octDates, S.getCodes_());
    assert.equal(sched.cells[b.id]['2026-11-02'], expect);
    assert.equal(sched.cells[b.id]['2026-11-08'] || '', '');
  });

  const v0 = sched.version;
  const prev = sched.cells[a.id]['2026-10-06'] || '';
  test('본인 일정 변경 → 변경이력 + 폴링으로 전달', () => {
    assert.deepEqual({ ...S.setCell({ actorId: a.id, personId: a.id, date: '2026-10-06', code: '연차', prev }) }, { ok: true });
    const ch = S.getChanges(v0, sched.sv);
    assert.equal(ch.changes.length, 1); assert.equal(ch.changes[0].code, '연차'); assert.equal(ch.version, v0 + 1);
  });
  test('동시 수정 충돌: 이전 값이 다르면 덮어쓰지 않음', () => {
    const r = S.setCell({ actorId: a.id, personId: a.id, date: '2026-10-06', code: '주간', prev });
    assert.equal(r.conflict, true); assert.equal(r.current, '연차');
  });
  test('팀원은 남의 일정 변경 불가, 없는 근태 거부', () => {
    assert.throws(() => S.setCell({ actorId: a.id, personId: b.id, date: '2026-10-06', code: '주간', prev: '' }), /관리자만/);
    assert.throws(() => S.setCell({ actorId: a.id, personId: a.id, date: '2026-10-07', code: '없는근태', prev: '' }), /없는 근태/);
    assert.throws(() => S.setCell({ actorId: a.id, personId: a.id, date: '2026-12-07', code: '주간', prev: '' }), /생성되지 않았습니다/);
  });

  // 관리자 지정 (메뉴와 같은 처리)
  Object.assign(a, { admin: true, pwHash: S.hashPw_('pw1234'), email: 'boss@example.com' }); S.writePerson_(a);
  let tok;
  test('관리자 로그인: 틀린 비밀번호 거부, 토큰 발급', () => {
    assert.throws(() => S.adminLogin(a.id, 'nope'), /비밀번호/);
    tok = S.adminLogin(a.id, 'pw1234');
    assert.ok(S.checkAdminToken(a.id, tok)); assert.ok(!S.checkAdminToken(b.id, tok));
  });
  test('관리자는 남의 일정 변경 가능 (대리 수정 기록)', () => {
    const cur = S.getSchedule().cells[b.id]['2026-10-07'] || '';
    assert.equal(S.setCell({ actorId: a.id, token: tok, personId: b.id, date: '2026-10-07', code: '교육', prev: cur }).ok, true);
    const h = S.getHistory(a.id, tok, 10);
    assert.equal(h[0].target, b.name); assert.equal(h[0].actor, a.name);
  });
  let newId;
  test('관리자: 인원 추가 → 같은 구분 뒤에 배치, 10·11월에 기본 근태 행', () => {
    const sv = S.structVersion_();
    newId = S.savePerson(a.id, tok, { name: '신입', team: b.team, org: b.org, g1: b.g1, g2: b.g2, g3: b.g3, admin: false });
    assert.ok(S.structVersion_() > sv);
    const list = S.activePeople_(), idx = list.findIndex(p => p.id === newId);
    assert.equal(list[idx - 1].org, b.org);
    const nov = S.readMonth_('2026-11').rows[newId].values;
    assert.equal(nov[1], '주간'); assert.equal(nov[0], '');
    assert.ok(S.readMonth_('2026-10').rows[newId]); assert.ok(!(S.readMonth_('2026-09') || { rows: {} }).rows[newId]);
    assert.ok(S.getChanges(0, sv).reload);
  });
  test('팀원은 인원 관리 불가', () => {
    assert.throws(() => S.savePerson(b.id, '', { name: 'x' }), /관리자만/);
    assert.throws(() => S.removePerson(b.id, tok, newId), /관리자만/);
  });
  test('관리자: 인원 빼기 → 명단에서 사라지고 기록은 유지', () => {
    S.removePerson(a.id, tok, newId);
    assert.ok(!S.activePeople_().some(p => p.id === newId));
    assert.ok(S.readMonth_('2026-10').rows[newId]);
    assert.throws(() => S.setCell({ actorId: newId, personId: newId, date: '2026-10-08', code: '주간', prev: '주간' }), /명단에 없는/);
  });
  test('관리자: 본인 권한 해제·본인 삭제 불가', () => {
    assert.throws(() => S.savePerson(a.id, tok, { id: a.id, name: a.name, admin: false }), /본인의 관리자/);
    assert.throws(() => S.removePerson(a.id, tok, a.id), /본인은/);
  });

  test('15시 요약 메일: 변경 건수만큼, 다시 보내면 0건', () => {
    const n = S.sendDailySummary();
    assert.equal(n, 2); assert.equal(gas.mails.length, 1); assert.equal(gas.mails[0].to, 'boss@example.com');
    assert.match(gas.mails[0].htmlBody, /연차/);
    assert.equal(S.sendDailySummary(), 0);
  });

  test('20일에는 12월 미생성, 21일에 12월 자동 생성 (11월 최다 근무형태)', () => {
    gas.setNow('2026-10-20T02:00:00+09:00'); S.dailyJob();
    assert.ok(!S.monthSheet_('2026-12'));
    gas.setNow('2026-10-21T01:30:00+09:00'); S.dailyJob();
    assert.ok(S.monthSheet_('2026-12'));
    const dec = S.readMonth_('2026-12').rows[b.id].values;
    assert.equal(dec[0], S.readMonth_('2026-11').rows[b.id].values[1]); // 12/1(화) = 11월 기본값
    assert.equal(dec[5], ''); // 12/6 일
    assert.ok(!S.readMonth_('2026-12').rows[newId]); // 뺀 사람은 생성 안 함
  });
  test('11월이 되면 화면은 11월·12월', () => {
    gas.setNow('2026-11-01T09:00:00+09:00');
    assert.deepEqual([...S.getBootstrap().months.map(m => m.key)], ['2026-11', '2026-12']);
  });

  test('엑셀 추출: 양식대로 그려지고 누구나 가능', () => {
    gas.setNow('2026-10-25T09:00:00+09:00');
    const r = S.exportMonth(b.id, '', '2026-10');
    assert.match(r.filename, /26년 10월 근태계획/);
    const sh = created.at(-1).getSheets()[0];
    assert.equal(sh.getName(), '10월 근무표');
    assert.match(String(sh.get(1, 1)), /'26년 10월 근태계획$/);
    assert.equal(sh.get(3, 7), '월'); assert.equal(sh.get(4, 7), 28);
    const names = []; for (let r2 = 5; sh.get(r2, 6); r2++) names.push(sh.get(r2, 6));
    assert.ok(names.includes('신입'), '10월에 있던(10/25 제외) 사람은 10월 표에 남음');
    if (parsed) assert.equal(names.length, parsed.people.length + 1);
  });
}

console.log(`\n${passed}개 통과`);

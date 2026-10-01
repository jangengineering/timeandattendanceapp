/* 미리보기용 가짜 서버. 실제 배포에는 포함되지 않는다. (tools/build-preview.mjs 가 주입) */
(function () {
  const CODES = [
    ['주간', '#ffffff', '#000000', 1, 1, ''], ['1Shift', '#ffccff', '#000000', 1, 1, ''], ['2Shift', '#ffe699', '#000000', 1, 1, ''],
    ['연차', '#ffffff', '#ff0000', 0, 1, ''], ['반차', '#ffffff', '#ff0000', 0, 1, ''], ['공가', '#ffffff', '#ff0000', 0, 1, ''],
    ['휴무', '#fbe5d6', '#ff0000', 0, 1, ''], ['교육', '#d0cece', '#000000', 0, 1, ''], ['주간2', '#e2efda', '#000000', 1, 1, ''],
    ['주간\n(OT)', '#ffffff', '#000000', 1, 0, '주간'], ['주간\n(연장)', '#ffffff', '#000000', 1, 0, '주간'], ['주간\n(반일교육)', '#ffffff', '#000000', 1, 0, '주간'],
    ['1Shift\n(OT)', '#ffccff', '#000000', 1, 0, '1Shift'], ['1Shift\n(전반)', '#ffccff', '#000000', 1, 0, '1Shift'], ['1Shift\n(후반)', '#ffccff', '#000000', 1, 0, '1Shift'],
    ['2Shift\n(OT)', '#ffe699', '#000000', 1, 0, '2Shift'], ['2Shift\n(조출)', '#ffe699', '#000000', 1, 0, '2Shift'], ['철야', '#ffe699', '#000000', 1, 0, '2Shift'],
    ['검진', '#ffffff', '#ff0000', 0, 0, ''], ['출장', '#d9d9d9', '#000000', 0, 0, ''], ['힐링\n캠프', '#fbe5d6', '#ff0000', 0, 0, ''],
  ].map(([code, bg, fg, work, primary, base]) => ({ code, bg, fg, work: !!work, primary: !!primary, base, active: true }));

  const ORG = [
    ['건축', 'SM', '', '', '', 1], ['건축', '건축팀', 'DSM', '', '', 1],
    ['건축', '건축팀', 'ST-1', 'FAB RC', 'FAB', 4], ['건축', '건축팀', 'ST-1', 'SUP 철골', 'SUP', 4],
    ['건축', '건축팀', 'ST-2', 'FAB 철골', 'FAB', 4], ['건축', '건축팀', '자재검사', '', '자재검사', 3],
    ['안전', '안전', 'ST2', '', 'FAB1', 4], ['안전', '안전', 'ST1', '', 'FAB2', 3],
  ];
  const NAMES = '김민준 이서연 박도윤 최서윤 정하준 강지우 조은우 윤하은 장시우 임지호 한수아 오예준 서지민 신유진 권도현 황채원 안주원 송민서 류건우 전다은 홍길동 고은서 문현우 양서진'.split(' ');
  const people = []; let n = 0;
  ORG.forEach(([team, org, g1, g2, g3, cnt]) => { for (let i = 0; i < cnt; i++) { const id = 'P' + (++n); people.push({ id, name: NAMES[(n - 1) % NAMES.length], team, org, g1, g2, g3, admin: id === 'P1' }); } });

  const fmt = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  const today = '2026-10-01';
  const months = [{ key: '2026-10', label: '10월', year: 2026, month: 10, start: '2026-09-28', days: 35 },
                  { key: '2026-11', label: '11월', year: 2026, month: 11, start: '2026-10-26', days: 42 }];
  const cells = {}, generated = {};
  const rot = ['1Shift', '2Shift', '주간'];
  for (let d = new Date(2026, 8, 1); d <= new Date(2026, 11, 6); d.setDate(d.getDate() + 1)) {
    const ds = fmt(d); if (d.getMonth() <= 10) generated[ds] = true;
    if (!generated[ds]) continue;
    const wd = (d.getDay() + 6) % 7, week = Math.floor((d - new Date(2026, 8, 28)) / 6048e5);
    people.forEach((p, i) => {
      let c = wd === 6 ? '' : p.team === '안전' ? (i % 2 ? '주간2' : '주간') : rot[(i + week + 30) % 3];
      if (wd === 5 && i % 4 === 0) c = '휴무';
      if ((i * 7 + d.getDate()) % 41 === 0 && wd < 5) c = '연차';
      (cells[p.id] = cells[p.id] || {})[ds] = c;
    });
  }
  let version = 1; const history = [];

  const api = {
    getBootstrap: () => ({ title: "P5 FAB 건축/안전팀 근태계획", codes: CODES, people, months, today }),
    checkAdminToken: (id, t) => t === 'tok',
    adminLogin: (id, pw) => { if (pw !== '1234') throw new Error('비밀번호가 맞지 않습니다. (미리보기 비밀번호: 1234)'); return 'tok'; },
    getSchedule: () => ({ cells: JSON.parse(JSON.stringify(cells)), generated, version, sv: 1 }),
    setCell: ({ actorId, personId, date, code, prev }) => {
      const cur = cells[personId][date] || '';
      if (cur !== prev) return { conflict: true, current: cur };
      cells[personId][date] = code; version++;
      history.unshift({ at: new Date().toLocaleString('ko-KR'), actor: people.find(p => p.id === actorId).name, target: people.find(p => p.id === personId).name, date, before: cur, after: code });
      return { ok: true, version };
    },
    getChanges: v => ({ changes: [], version }),
    getHistory: () => history,
    getPersonDetail: () => ({ email: 'admin@example.com' }),
    movePerson: (a, t, id, dir) => { const i = people.findIndex(p => p.id === id); const [p] = people.splice(i, 1); people.splice(i + dir, 0, p); },
    savePerson: (a, t, data) => {
      if (data.id) Object.assign(people.find(p => p.id === data.id), data);
      else { const id = 'P' + (++n); people.push({ ...data, id }); cells[id] = {}; Object.keys(generated).forEach(d => cells[id][d] = new Date(d).getDay() === 0 ? '' : '주간'); }
    },
    removePerson: (a, t, id) => { people.splice(people.findIndex(p => p.id === id), 1); },
    saveCode: (a, t, data) => { const c = CODES.find(x => x.code === data.code); if (c) Object.assign(c, data); else CODES.push(data); },
    exportMonth: () => { throw new Error('미리보기에서는 엑셀 추출을 할 수 없습니다.'); },
  };
  window.MockServer = {
    call: (fn, args) => new Promise((res, rej) => setTimeout(() => { try { res(api[fn](...args)); } catch (e) { rej(e); } }, 150)),
  };
})();

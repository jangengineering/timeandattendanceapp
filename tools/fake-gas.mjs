// Apps Script 서비스의 아주 작은 가짜 구현 — 서버 코드를 Node 에서 시험하기 위한 것.
import { createHash, randomUUID } from 'node:crypto';

class FakeRange {
  constructor(sheet, r, c, nr, nc) { Object.assign(this, { sheet, r, c, nr, nc }); }
  getRow() { return this.r; } getColumn() { return this.c; } getNumRows() { return this.nr; } getNumColumns() { return this.nc; }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.sheet.get(this.r + i, this.c + j)); out.push(row); }
    return out;
  }
  getValue() { return this.sheet.get(this.r, this.c); }
  setValues(v) {
    if (v.length !== this.nr || v.some(r => r.length !== this.nc)) throw new Error(`setValues 크기 불일치 ${v.length}x${v[0]?.length} vs ${this.nr}x${this.nc}`);
    v.forEach((row, i) => row.forEach((x, j) => this.sheet.set(this.r + i, this.c + j, x))); return this;
  }
  setValue(x) { this.sheet.set(this.r, this.c, x); return this; }
  clearContent() { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sheet.set(this.r + i, this.c + j, ''); return this; }
  getBackgrounds() { return this.getValues().map(r => r.map(() => '#ffffff')); }
  getFontColors() { return this.getValues().map(r => r.map(() => '#000000')); }
  getMergedRanges() { return this.sheet.merges.map(m => new FakeRange(this.sheet, ...m)); }
  merge() { this.sheet.merges.push([this.r, this.c, this.nr, this.nc]); return this; }
}
for (const m of ['setFontWeight', 'setBackground', 'setFontColor', 'setWrap', 'setHorizontalAlignment', 'setVerticalAlignment', 'setFontFamily',
  'setFontSize', 'setBackgrounds', 'setFontColors', 'setFontWeights', 'setBorder']) FakeRange.prototype[m] = function () { return this; };

class FakeSheet {
  constructor(ss, name) { this.ss = ss; this.name = name; this.data = []; this.merges = []; }
  getName() { return this.name; } setName(n) { this.name = n; return this; }
  get(r, c) { const v = (this.data[r - 1] || [])[c - 1]; return v === undefined || v === null ? '' : v; }
  set(r, c, v) { while (this.data.length < r) this.data.push([]); this.data[r - 1][c - 1] = v; }
  getLastRow() { for (let i = this.data.length; i > 0; i--) if ((this.data[i - 1] || []).some(v => v !== '' && v != null)) return i; return 0; }
  getLastColumn() { return Math.max(0, ...this.data.map(r => { for (let j = r.length; j > 0; j--) if (r[j - 1] !== '' && r[j - 1] != null) return j; return 0; })); }
  getRange(r, c, nr = 1, nc = 1) {
    if (typeof r === 'string') { const m = r.match(/^([A-Z]+)(\d+)$/); let col = 0; for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64; return new FakeRange(this, +m[2], col, 1, 1); }
    if (nr < 1 || nc < 1) throw new Error('범위 크기 0');
    return new FakeRange(this, r, c, nr, nc);
  }
  getDataRange() { return new FakeRange(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
  appendRow(row) { const r = this.getLastRow() + 1; row.forEach((v, j) => this.set(r, j + 1, v)); return this; }
  clear() { this.data = []; this.merges = []; return this; }
}
for (const m of ['setFrozenRows', 'setFrozenColumns', 'hideColumns', 'setColumnWidths', 'setColumnWidth', 'setRowHeight']) FakeSheet.prototype[m] = function () { return this; };

class FakeSpreadsheet {
  constructor(id) { this.id = id; this.sheets = []; }
  getId() { return this.id; }
  getSheets() { return this.sheets; }
  getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
  insertSheet(n) { const s = new FakeSheet(this, n); this.sheets.push(s); return s; }
}

export function makeGas({ today }) {
  const ss = new FakeSpreadsheet('SS1');
  const props = {}, cache = {}, mails = [], triggers = [];
  let now = new Date(today + 'T10:00:00+09:00');
  const gas = {
    ss, mails, triggers,
    setNow(s) { now = new Date(s); },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss, openById: () => ss, flush() {},
      BorderStyle: { SOLID: 'SOLID' },
    },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: k => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = String(v); },
      deleteProperty: k => { delete props[k]; }, getProperties: () => ({ ...props }) }) },
    CacheService: { getScriptCache: () => ({ get: k => (k in cache ? cache[k] : null), put: (k, v) => { cache[k] = v; } }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Utilities: {
      formatDate(d, tz, f) {
        const k = new Date(d.getTime() + 9 * 3600000); const p = n => String(n).padStart(2, '0');
        return f.replace('yyyy', k.getUTCFullYear()).replace('yy', String(k.getUTCFullYear()).slice(2)).replace('MM', p(k.getUTCMonth() + 1))
          .replace('dd', p(k.getUTCDate())).replace('HH', p(k.getUTCHours())).replace('mm', p(k.getUTCMinutes())).replace('ss', p(k.getUTCSeconds()))
          .replace(/\bM\b/, k.getUTCMonth() + 1).replace(/\bd\b/, k.getUTCDate());
      },
      getUuid: () => randomUUID(),
      computeDigest: (alg, s) => [...createHash('sha256').update(s).digest()],
      base64Encode: b => Buffer.from(b.map ? b.map(x => x & 255) : b).toString('base64'),
      sleep() {}, DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
    },
    Session: { getEffectiveUser: () => ({ getEmail: () => 'owner@example.com' }) },
    MailApp: { sendEmail: m => mails.push(m) },
    ScriptApp: { getOAuthToken: () => 'tok', getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/FAKE/exec' }), getProjectTriggers: () => triggers,
      deleteTrigger() {}, newTrigger: h => { const b = { timeBased: () => b, everyDays: () => b, atHour: () => b, nearMinute: () => b, inTimezone: () => b, create: () => triggers.push(h) }; return b; } },
  };
  // Date 를 고정 시각으로
  const RealDate = Date;
  gas.Date = class extends RealDate { constructor(...a) { super(...(a.length ? a : [now.getTime()])); } static now() { return now.getTime(); } };
  return gas;
}

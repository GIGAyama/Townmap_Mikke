/**
 * Tenant.gs — アクセス制御と、記録まわりの共通コア
 *
 * 認可の鉄則: すべての API はサーバー側で
 *   ① 本人確認（Session.getActiveUser）→ ② 名簿照合（status=active）
 *   → ③ 役割チェック → ④ 行所有者チェック
 * の順にガードする（入口は Bound.gs）。フロントの出し分けは防御とみなさない。
 *
 * 児童向けレスポンスでは email を露出せず、匿名 ID（uid = sha256(email) 先頭12桁）に
 * 置換する。児童の画面に他の子のメールアドレスを出さないため。
 */

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
  // 1文字の名前は普通の言葉と衝突する（本文の一般語まで置き換わり、
  // 実名も守れず本文も壊れる）ため対象外にする。
  if (target.length >= 2) { aliases[target] = '対象児童'; reverse['対象児童'] = target; }
  let seq = 0;
  (names || []).forEach(function (raw) {
    const name = String(raw === null || raw === undefined ? '' : raw).trim();
    if (name.length < 2 || aliases[name]) return; // 空欄・1文字・同名の重複は飛ばす
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
    if (compact.length >= 2 && compact !== name) {
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
  const pinId = vRecordId_(d.pin_id);
  // 児童が書く文字（なまえ・メモ）は safeCellText_ を通す。先生がシートを開いた
  // 瞬間に数式として実行されるのを防ぐため（Db.gs の safeCellText_ を参照）。
  appendRowLocked_(ss, TABLES.PINS, {
    pin_id: pinId,
    unit_id: vStr_(d.unit_id, 60, '単元'),
    map_id: vStr_(d.map_id, 60, '地図'),
    email: email,                            // 検証済み email（クライアント申告は無視）
    x: vNum_(d.x, 0, 100, '位置'),
    y: vNum_(d.y, 0, 100, '位置'),
    color: safeCellText_(vStr_(d.color, 20, 'アイコン')),
    title: safeCellText_(vStr_(d.title, 100, 'なまえ')),
    memo: safeCellText_(vStr_(d.memo, 1000, 'メモ')),
    image_url: vImageUrl_(d.image_url),
    created_at: new Date()
  });
  return { recordId: pinId };
}

function coreSaveChat_(ss, email, d) {
  const targetType = ['general', 'pin', 'chat'].indexOf(d.target_type) >= 0 ? d.target_type : 'general';
  const msg = vStr_(d.message, 500, 'メッセージ');
  if (!msg.trim()) throw new Error('BAD_INPUT: メッセージが空です');
  const chatId = vRecordId_(d.chat_id);
  appendRowLocked_(ss, TABLES.CHATS, {
    chat_id: chatId,
    unit_id: vStr_(d.unit_id, 60, '単元'),
    email: email,
    message: safeCellText_(msg),
    target_type: targetType,
    target_id: vStr_(d.target_id, 60, '返信先'),
    created_at: new Date()
  });
  return { recordId: chatId };
}

function coreToggleReaction_(ss, email, d) {
  const targetType = ['pin', 'chat'].indexOf(d.target_type) >= 0 ? d.target_type : 'pin';
  const targetId = vStr_(d.target_id, 60, '対象');
  const emoji = vStr_(d.emoji, 8, 'スタンプ');
  const unitId = vStr_(d.unit_id, 60, '単元');
  withScriptLock_(function () {
    const sheet = ensureSheet_(ss, TABLES.REACTIONS);
    const data = sheet.getDataRange().getValues();
    const H = headerMapFromRow_(data[0], TABLES.REACTIONS);
    requireCols_(H, TABLES.REACTIONS, ['email', 'target_type', 'target_id', 'emoji']);
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][H.email]).toLowerCase() === email && data[i][H.target_type] === targetType &&
          String(data[i][H.target_id]) === targetId && data[i][H.emoji] === emoji) {
        sheet.deleteRow(i + 1);
        return;
      }
    }
    sheet.appendRow(rowFor_(H, TABLES.REACTIONS, {
      reaction_id: Utilities.getUuid(),
      unit_id: unitId,
      email: email,
      target_type: targetType,
      target_id: targetId,
      emoji: safeCellText_(emoji),
      created_at: new Date()
    }, data[0].length));
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
    const PH = headerMapFromRow_(pinData[0], TABLES.PINS);
    // 行の持ち主が読めない状態で「自分のだけ」を判定させない（消せてしまうため）
    requireCols_(PH, TABLES.PINS, ['pin_id', 'email']);
    for (let i = 1; i < pinData.length; i++) {
      if (String(pinData[i][PH.pin_id]) === id) {
        if (ownOnly && String(pinData[i][PH.email]).toLowerCase() !== email) {
          throw new Error('FORBIDDEN: 自分の記録だけが削除できます');
        }
        pinSheet.deleteRow(i + 1);
        return { deleted: 'pin' };
      }
    }
    const chatSheet = ensureSheet_(ss, TABLES.CHATS);
    const chatData = chatSheet.getDataRange().getValues();
    const CH = headerMapFromRow_(chatData[0], TABLES.CHATS);
    requireCols_(CH, TABLES.CHATS, ['chat_id', 'email']);
    for (let i = 1; i < chatData.length; i++) {
      if (String(chatData[i][CH.chat_id]) === id) {
        if (ownOnly && String(chatData[i][CH.email]).toLowerCase() !== email) {
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
    const H = headerMapFromRow_(data[0], TABLES.PINS);
    requireCols_(H, TABLES.PINS, ['pin_id', 'email']);
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][H.pin_id]) === id) {
        if (String(data[i][H.email]).toLowerCase() !== email) {
          throw new Error('FORBIDDEN: 自分の記録だけが変更できます');
        }
        const patch = {};
        if (d.x !== undefined) patch.x = vNum_(d.x, 0, 100, '位置');
        if (d.y !== undefined) patch.y = vNum_(d.y, 0, 100, '位置');
        if (d.color !== undefined) patch.color = safeCellText_(vStr_(d.color, 20, 'アイコン'));
        if (d.title !== undefined) patch.title = safeCellText_(vStr_(d.title, 100, 'なまえ'));
        if (d.memo !== undefined) patch.memo = safeCellText_(vStr_(d.memo, 1000, 'メモ'));
        if (d.image_url !== undefined) patch.image_url = vImageUrl_(d.image_url);
        const row = setCells_(data[i].slice(), H, TABLES.PINS, patch);
        sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
        return { recordId: id };
      }
    }
    throw new Error('NOT_FOUND: 対象の記録が見つかりません');
  });
}

/**
 * 単元まわりの教員操作（単元の追加・地図の追加・チャット/スタンプの切替・スタンプの編集）。
 *
 * 以前は TeacherApi.gs と Legacy.gs に同じ処理が 2 本あった。コンテナバインド版
 * （Bound.gs）を足すと 3 本になり、片方だけ直したときに「先生の画面によって
 * 挙動が違う」が起きる。呼び出し口（誰が呼べるか）だけを各 API に残し、
 * シートを触る部分はここ 1 本にまとめる。
 *
 * **認可はここでは見ない。** 呼ぶ側が「先生であること」を確かめてから呼ぶこと。
 */
function coreUnitAction_(ss, p) {
  if (p.action === 'save_unit') {
    const unitId = vRecordId_(p.unit_id);
    const unitName = safeCellText_(vStr_(p.name, 60, '単元名'));
    const initMap = [{ id: 'm_' + Date.now(), name: vStr_(p.map_name, 40, '地図名') || '基本マップ', url: vImageUrl_(p.map_url) }];
    const initStamps = JSON.stringify(['📍', '🐛', '🌸', '🚗', '⚠️', '🏠', '❓', '💡']);
    withScriptLock_(function () {
      const unitSheet = ensureSheet_(ss, TABLES.UNITS);
      const data = unitSheet.getDataRange().getValues();
      const H = headerMapFromRow_(data[0], TABLES.UNITS);
      requireCols_(H, TABLES.UNITS, ['unit_id', 'name', 'maps_json', 'chat_enabled', 'stamp_enabled', 'custom_stamps', 'is_active', 'created_at']);
      for (let i = 1; i < data.length; i++) {
        if (data[i][H.is_active] === true) unitSheet.getRange(i + 1, H.is_active + 1).setValue(false);
      }
      unitSheet.appendRow(rowFor_(H, TABLES.UNITS, {
        unit_id: unitId, name: unitName, maps_json: JSON.stringify(initMap),
        chat_enabled: true, stamp_enabled: true, custom_stamps: initStamps,
        is_active: true, created_at: new Date()
      }, data[0].length));
    });
    return { unitId: unitId };
  }

  if (p.action === 'add_map') {
    const mapId = vRecordId_(p.map_id);
    const mapName = vStr_(p.name, 40, '地図名');
    const mapUrl = vImageUrl_(p.map_url);
    withScriptLock_(function () {
      const unitSheet = ensureSheet_(ss, TABLES.UNITS);
      const data = unitSheet.getDataRange().getValues();
      const H = headerMapFromRow_(data[0], TABLES.UNITS);
      requireCols_(H, TABLES.UNITS, ['unit_id', 'maps_json']);
      for (let i = 1; i < data.length; i++) {
        if (data[i][H.unit_id] === p.unit_id) {
          let maps = [];
          try { maps = JSON.parse(data[i][H.maps_json] || '[]'); } catch (e) { maps = []; }
          maps.push({ id: mapId, name: mapName, url: mapUrl });
          unitSheet.getRange(i + 1, H.maps_json + 1).setValue(JSON.stringify(maps));
          break;
        }
      }
      if (p.copy_from_map_id) {
        const pinSheet = ensureSheet_(ss, TABLES.PINS);
        const pinData = pinSheet.getDataRange().getValues();
        const PH = headerMapFromRow_(pinData[0], TABLES.PINS);
        requireCols_(PH, TABLES.PINS, ['pin_id', 'unit_id', 'map_id', 'email', 'x', 'y', 'color', 'title', 'memo', 'image_url', 'created_at']);
        const width = pinData[0].length;
        const newPins = [];
        for (let i = 1; i < pinData.length; i++) {
          if (pinData[i][PH.unit_id] === p.unit_id && pinData[i][PH.map_id] === p.copy_from_map_id) {
            newPins.push(rowFor_(PH, TABLES.PINS, {
              pin_id: Utilities.getUuid(), unit_id: p.unit_id, map_id: mapId,
              email: pinData[i][PH.email], x: pinData[i][PH.x], y: pinData[i][PH.y],
              color: pinData[i][PH.color], title: pinData[i][PH.title], memo: pinData[i][PH.memo],
              image_url: pinData[i][PH.image_url], created_at: new Date()
            }, width));
          }
        }
        if (newPins.length > 0) {
          pinSheet.getRange(pinSheet.getLastRow() + 1, 1, newPins.length, newPins[0].length).setValues(newPins);
        }
      }
    });
    return {};
  }

  if (p.action === 'toggle_chat' || p.action === 'toggle_stamp') {
    const col = p.action === 'toggle_chat' ? 'chat_enabled' : 'stamp_enabled';
    const val = p.action === 'toggle_chat' ? p.chat_enabled === true : p.stamp_enabled === true;
    withScriptLock_(function () {
      const unitSheet = ensureSheet_(ss, TABLES.UNITS);
      const data = unitSheet.getDataRange().getValues();
      const H = headerMapFromRow_(data[0], TABLES.UNITS);
      requireCols_(H, TABLES.UNITS, ['unit_id', col]);
      for (let i = 1; i < data.length; i++) {
        if (data[i][H.unit_id] === p.unit_id) { unitSheet.getRange(i + 1, H[col] + 1).setValue(val); break; }
      }
    });
    return {};
  }

  if (p.action === 'update_custom_stamps') {
    const stamps = (p.custom_stamps || []).slice(0, 24).map(function (s) { return vStr_(s, 8, 'スタンプ'); });
    withScriptLock_(function () {
      const unitSheet = ensureSheet_(ss, TABLES.UNITS);
      const data = unitSheet.getDataRange().getValues();
      const H = headerMapFromRow_(data[0], TABLES.UNITS);
      requireCols_(H, TABLES.UNITS, ['unit_id', 'custom_stamps']);
      for (let i = 1; i < data.length; i++) {
        if (data[i][H.unit_id] === p.unit_id) {
          unitSheet.getRange(i + 1, H.custom_stamps + 1).setValue(JSON.stringify(stamps));
          break;
        }
      }
    });
    return {};
  }

  throw new Error('BAD_INPUT: 不明な操作です');
}

/**
 * AI ポートフォリオのプロンプトを組み立てる（**実名は 1 文字も入れない**）。
 *
 * 以前は TeacherApi.gs と Legacy.gs に同じ文面が 2 本あり、片方だけ仮名化を直すと
 * もう片方から実名のまま Gemini へ流れる形だった。組み立てをここ 1 本にして、
 * 「仮名化を通していない文字列を prompt に足す」ができないようにしている。
 *
 * @param {string[]} allNames 名簿にある表示名（仮名の対応表を作るため全員ぶん渡す）
 * @param {string} targetName 対象児童の表示名
 * @return {{prompt: string, reverse: Object}} reverse は返答を実名に戻すための対応表
 */
function corePortfolioPrompt_(allNames, targetName, pins, chats, reactions) {
  const aliasMap = createNameAliases_(allNames, targetName);
  let prompt = 'あなたは小学校の先生です。児童「対象児童」' +
    'の「地図学習」での活動記録を分析し、温かいフィードバックを作成してください。\n' +
    '※児童名は「対象児童」「児童A」のような仮名にしてあります。返事でも仮名のまま書いてください。\n\n';
  prompt += '【ピンを刺した記録】\n';
  pins.forEach(function (pin) {
    prompt += '- 発見対象[' + redactForAi_(pin.title, aliasMap.aliases) + ']: メモ['
      + (redactForAi_(pin.memo, aliasMap.aliases) || 'なし') + '] アイコン[' + redactForAi_(pin.color, aliasMap.aliases) + ']\n';
  });
  prompt += '\n【発言記録】\n';
  chats.forEach(function (chat) { prompt += '- ' + redactForAi_(chat.message, aliasMap.aliases) + '\n'; });
  prompt += '\n【友達へのリアクション回数】: ' + reactions.length + '回\n';
  prompt += '\n以下の3項目で出力してください。\n1. 🔍 興味関心の傾向（どんなものに目を向けているか）\n' +
    '2. ✨ 素晴らしい点（表現や友達への関わりの良さ）\n3. 💌 先生からのメッセージ（小学生に向けて優しい言葉で）';
  return { prompt: prompt, reverse: aliasMap.reverse };
}

/** 組み立てたプロンプトを Gemini（正本 Gemini.gs / GigaGemini）へ送り、仮名を戻して返す */
function coreRunPortfolio_(apiKey, built) {
  // 通信・再試行・応答の取り出しは正本 Gemini.gs（GigaGemini）に任せる。
  // ここに直書きしていた頃は再試行が無く、混み合う時間帯（429）に
  // ポートフォリオ生成がそのまま失敗していた。API キーは正本側で
  // x-goog-api-key ヘッダに載る（URL クエリには入れない）。
  const aiText = GigaGemini.call({
    apiKey: apiKey,
    prompt: built.prompt,
    model: 'gemini-2.5-flash',
    systemInstruction: 'あなたは優しく、児童の良いところを見つけるのが得意な先生です。マークダウンを使用せず、プレーンテキストで見やすく出力してください。'
  });
  // 先生の画面では実名で読めるよう、仮名を表示名に戻してから返す。
  return rehydrateAliases_(aiText, built.reverse);
}

/** 単元系データの取得（読みは各シート getValues 1 回） */
function coreCollectUnitData_(ss, unitId) {
  return {
    pins: getTableData_(ss, TABLES.PINS).filter(function (p) { return p.unit_id === unitId; }),
    chats: getTableData_(ss, TABLES.CHATS).filter(function (c) { return c.unit_id === unitId; }),
    reactions: getTableData_(ss, TABLES.REACTIONS).filter(function (r) { return r.unit_id === unitId; })
  };
}

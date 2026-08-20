/**
 * Legacy.gs — 旧バインド型デプロイ互換 API（lg*）
 *
 * このコードをスプレッドシートにバインドして「アクセスしているユーザーとして実行」で
 * デプロイした場合（旧来の 1 クラス 1 デプロイ運用）に、従来どおり動作させるための層。
 * getSs_(null) は getActiveSpreadsheet() を優先するため（Tenant.gs）、バインド文脈では
 * 自動的にそのシートが DB になる。名簿は旧来の Users_名簿 シートをそのまま使う。
 *
 * 旧版との違い:
 *  - getDriveImages は廃止（フル Drive スコープ回避のため）。画像はアップロード
 *    （Images_画像シート保存）方式に統一。
 *  - 書き込み者 email はクライアント申告ではなく Session.getActiveUser() で強制。
 */

function lgEmail_() {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('AUTH_INVALID: Googleアカウントにログインしていません。');
  return String(email).toLowerCase();
}

function lgEnsureLegacySheets_(ss) {
  ensureSheet_(ss, TABLES.USERS);
  ensureSheet_(ss, TABLES.UNITS);
  ensureSheet_(ss, TABLES.PINS);
  ensureSheet_(ss, TABLES.CHATS);
  ensureSheet_(ss, TABLES.REACTIONS);
  ensureSheet_(ss, TABLES.IMAGES);
  ensureSheet_(ss, TABLES.SETTINGS);
}

function lgApiKey_(ss) {
  // 新: Settings シート → 旧: ScriptProperties(GEMINI_API_KEY) の順で読む（後方互換）
  return getSettingValue_(ss, 'geminiApiKey') ||
    PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
}

function lgGetInitData() {
  try {
    const email = lgEmail_();
    const ss = getSs_(null);
    lgEnsureLegacySheets_(ss);

    let users = getTableData_(ss, TABLES.USERS);
    let myUser = users.filter(function (u) { return String(u.email).toLowerCase() === email; })[0];

    if (users.length === 0) {
      // 最初にアクセスした人を先生として登録（旧版の挙動を維持）
      const newTeacher = { email: email, name: '先生', group_id: 'teacher', role: 'teacher', created_at: new Date().toLocaleString() };
      ensureSheet_(ss, TABLES.USERS).appendRow([newTeacher.email, newTeacher.name, newTeacher.group_id, newTeacher.role, newTeacher.created_at]);
      myUser = newTeacher;
      users = [newTeacher];
    }
    if (!myUser) return JSON.stringify({ success: false, status: 'unregistered', email: email });

    const units = getTableData_(ss, TABLES.UNITS);
    const activeUnit = formatUnit_(units.filter(function (u) { return u.is_active === true; })[0] || units[0] || null);
    const data = activeUnit ? coreCollectUnitData_(ss, activeUnit.unit_id) : { pins: [], chats: [], reactions: [] };

    return jsonOk_({
      user: myUser,
      users: users,
      activeUnit: activeUnit,
      units: units.map(function (u) { return formatUnit_(u); }),
      pins: data.pins,
      chats: data.chats,
      reactions: data.reactions,
      hasApiKey: !!lgApiKey_(ss)
    });
  } catch (e) { return jsonErr_(e); }
}

function lgSyncData(unitId) {
  try {
    const ss = getSs_(null);
    const units = getTableData_(ss, TABLES.UNITS);
    // 現在アクティブな単元を優先して返す（単元切替の自動追従用。StudentApi と同じ挙動）
    const activeUnit = formatUnit_(
      units.filter(function (u) { return u.is_active === true; })[0] ||
      units.filter(function (u) { return u.unit_id === unitId; })[0] || null);
    const data = coreCollectUnitData_(ss, unitId);
    return jsonOk_({
      pins: data.pins, chats: data.chats, reactions: data.reactions,
      activeUnit: activeUnit,
      users: getTableData_(ss, TABLES.USERS),
      hasApiKey: !!lgApiKey_(ss)
    });
  } catch (e) { return jsonErr_(e); }
}

function lgIsTeacher_(ss, email) {
  const u = getTableData_(ss, TABLES.USERS).filter(function (x) { return String(x.email).toLowerCase() === email; })[0];
  return !!u && u.role === 'teacher';
}

function lgExecuteAction(payloadJson) {
  try {
    const email = lgEmail_();
    const ss = getSs_(null);
    const p = JSON.parse(payloadJson);

    if (p.action === 'save_pin') return jsonOk_(coreSavePin_(ss, email, p));
    // 自分のピンのみ更新可（coreUpdateOwnPin_ が行所有者チェックを行う）
    if (p.action === 'update_pin') return jsonOk_(coreUpdateOwnPin_(ss, email, p.pin_id, p));
    if (p.action === 'save_chat') return jsonOk_(coreSaveChat_(ss, email, p));
    if (p.action === 'toggle_reaction') return jsonOk_(coreToggleReaction_(ss, email, p));
    if (p.action === 'delete_pin' || p.action === 'delete_chat') {
      return jsonOk_(coreDeleteRecord_(ss, email, p.pin_id || p.chat_id, !lgIsTeacher_(ss, email)));
    }

    // 以下は教員操作
    if (!lgIsTeacher_(ss, email)) throw new Error('FORBIDDEN: この操作は先生だけができます');

    if (p.action === 'save_users') {
      const userSheet = ensureSheet_(ss, TABLES.USERS);
      const existing = getTableData_(ss, TABLES.USERS).map(function (u) { return String(u.email).toLowerCase(); });
      (p.users || []).forEach(function (u) {
        if (!u || !u.email) return;
        const em = String(u.email).trim().toLowerCase();
        if (existing.indexOf(em) === -1) {
          userSheet.appendRow([em, String(u.name || '').trim(), String(u.group_id || '').trim(), 'student', new Date().toLocaleString()]);
        }
      });
      return jsonOk_({});
    }
    if (p.action === 'save_api_key') {
      setSettingValue_(ss, 'geminiApiKey', vStr_(p.api_key, 200, 'APIキー').trim());
      return jsonOk_({});
    }

    // 単元管理系は TeacherApi と同じ実装を経由させるため、疑似的に同処理を呼ぶ
    if (['save_unit', 'add_map', 'toggle_chat', 'toggle_stamp', 'update_custom_stamps'].indexOf(p.action) >= 0) {
      return lgUnitAction_(ss, p);
    }
    throw new Error('BAD_INPUT: 不明な操作です');
  } catch (e) { return jsonErr_(e); }
}

function lgUnitAction_(ss, p) {
  if (p.action === 'save_unit') {
    const unitId = vRecordId_(p.unit_id);
    const initMap = [{ id: 'm_' + Date.now(), name: vStr_(p.map_name, 40, '地図名') || '基本マップ', url: vImageUrl_(p.map_url) }];
    const initStamps = JSON.stringify(['📍', '🐛', '🌸', '🚗', '⚠️', '🏠', '❓', '💡']);
    withScriptLock_(function () {
      const unitSheet = ensureSheet_(ss, TABLES.UNITS);
      const data = unitSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][6] === true) unitSheet.getRange(i + 1, 7).setValue(false);
      }
      unitSheet.appendRow([unitId, vStr_(p.name, 60, '単元名'), JSON.stringify(initMap), true, true, initStamps, true, new Date()]);
    });
    return jsonOk_({ unitId: unitId });
  }
  if (p.action === 'add_map') {
    withScriptLock_(function () {
      const unitSheet = ensureSheet_(ss, TABLES.UNITS);
      const data = unitSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === p.unit_id) {
          let maps = [];
          try { maps = JSON.parse(data[i][2] || '[]'); } catch (e) { maps = []; }
          maps.push({ id: vRecordId_(p.map_id), name: vStr_(p.name, 40, '地図名'), url: vImageUrl_(p.map_url) });
          unitSheet.getRange(i + 1, 3).setValue(JSON.stringify(maps));
          break;
        }
      }
      if (p.copy_from_map_id) {
        const pinSheet = ensureSheet_(ss, TABLES.PINS);
        const pinData = pinSheet.getDataRange().getValues();
        const newPins = [];
        for (let i = 1; i < pinData.length; i++) {
          if (pinData[i][1] === p.unit_id && pinData[i][2] === p.copy_from_map_id) {
            newPins.push([Utilities.getUuid(), p.unit_id, p.map_id, pinData[i][3], pinData[i][4],
              pinData[i][5], pinData[i][6], pinData[i][7], pinData[i][8], pinData[i][9], new Date()]);
          }
        }
        if (newPins.length > 0) {
          pinSheet.getRange(pinSheet.getLastRow() + 1, 1, newPins.length, newPins[0].length).setValues(newPins);
        }
      }
    });
    return jsonOk_({});
  }
  if (p.action === 'toggle_chat' || p.action === 'toggle_stamp') {
    const col = p.action === 'toggle_chat' ? 4 : 5;
    const val = p.action === 'toggle_chat' ? p.chat_enabled === true : p.stamp_enabled === true;
    withScriptLock_(function () {
      const unitSheet = ensureSheet_(ss, TABLES.UNITS);
      const data = unitSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === p.unit_id) { unitSheet.getRange(i + 1, col).setValue(val); break; }
      }
    });
    return jsonOk_({});
  }
  if (p.action === 'update_custom_stamps') {
    const stamps = (p.custom_stamps || []).slice(0, 24).map(function (s) { return vStr_(s, 8, 'スタンプ'); });
    withScriptLock_(function () {
      const unitSheet = ensureSheet_(ss, TABLES.UNITS);
      const data = unitSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === p.unit_id) { unitSheet.getRange(i + 1, 6).setValue(JSON.stringify(stamps)); break; }
      }
    });
    return jsonOk_({});
  }
  throw new Error('BAD_INPUT: 不明な操作です');
}

function lgGenerateAIPortfolio(payloadJson) {
  try {
    const email = lgEmail_();
    const ss = getSs_(null);
    if (!lgIsTeacher_(ss, email)) throw new Error('FORBIDDEN: この操作は先生だけができます');
    const p = JSON.parse(payloadJson);
    const apiKey = lgApiKey_(ss);
    if (!apiKey) throw new Error('NO_API_KEY: AI分析を行うには、管理パネルの「AI設定」タブで Gemini API キーを設定してください。');

    const targetEmail = String(p.email || '').toLowerCase();
    const user = getTableData_(ss, TABLES.USERS).filter(function (u) { return String(u.email).toLowerCase() === targetEmail; })[0];
    const pins = getTableData_(ss, TABLES.PINS).filter(function (pin) { return pin.unit_id === p.unit_id && String(pin.email).toLowerCase() === targetEmail; });
    const chats = getTableData_(ss, TABLES.CHATS).filter(function (c) { return c.unit_id === p.unit_id && String(c.email).toLowerCase() === targetEmail; });
    const reactions = getTableData_(ss, TABLES.REACTIONS).filter(function (r) { return r.unit_id === p.unit_id && String(r.email).toLowerCase() === targetEmail; });

    if (pins.length === 0 && chats.length === 0 && reactions.length === 0) {
      return jsonOk_({ portfolio: 'まだ活動の記録（ピンやチャット）がありません。' });
    }

    // 新しい教員 API と同じく、AI には実名を渡さず仮名（対象児童・児童A…）で送る。
    const aliasMap = createNameAliases_(
      getTableData_(ss, TABLES.USERS).map(function (u) { return u.name; }),
      user ? user.name : ''
    );

    let prompt = 'あなたは小学校の先生です。児童「対象児童」' +
      'の「地図学習」での活動記録を分析し、温かいフィードバックを作成してください。\n' +
      '※児童名は「対象児童」「児童A」のような仮名にしてあります。返事でも仮名のまま書いてください。\n\n【ピンを刺した記録】\n';
    pins.forEach(function (pin) {
      prompt += '- 発見対象[' + redactForAi_(pin.title, aliasMap.aliases) + ']: メモ['
        + (redactForAi_(pin.memo, aliasMap.aliases) || 'なし') + '] アイコン[' + pin.color + ']\n';
    });
    prompt += '\n【発言記録】\n';
    chats.forEach(function (chat) { prompt += '- ' + redactForAi_(chat.message, aliasMap.aliases) + '\n'; });
    prompt += '\n【友達へのリアクション回数】: ' + reactions.length + '回\n';
    prompt += '\n以下の3項目で出力してください。\n1. 🔍 興味関心の傾向（どんなものに目を向けているか）\n2. ✨ 素晴らしい点（表現や友達への関わりの良さ）\n3. 💌 先生からのメッセージ（小学生に向けて優しい言葉で）';

    const payload = {
      contents: [{ parts: [{ text: prompt }] }],
      systemInstruction: { parts: [{ text: 'あなたは優しく、児童の良いところを見つけるのが得意な先生です。マークダウンを使用せず、プレーンテキストで見やすく出力してください。' }] }
    };
    const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
    const response = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey, options);
    const resData = JSON.parse(response.getContentText());
    if (resData.error) throw new Error('AI_ERROR: ' + resData.error.message);
    // 先生の画面では実名で読めるよう、仮名を名簿の名前に戻してから返す。
    return jsonOk_({ portfolio: rehydrateAliases_(resData.candidates[0].content.parts[0].text, aliasMap.reverse) });
  } catch (e) { return jsonErr_(e); }
}

function lgUploadImage(dataUrl) {
  try {
    const email = lgEmail_();
    const ss = getSs_(null);
    return jsonOk_({ imageRef: storeImage_(ss, email, dataUrl) });
  } catch (e) { return jsonErr_(e); }
}

function lgGetImage(imageRef) {
  try {
    lgEmail_();
    const ss = getSs_(null);
    return jsonOk_({ dataUrl: loadImage_(ss, imageRef) });
  } catch (e) { return jsonErr_(e); }
}

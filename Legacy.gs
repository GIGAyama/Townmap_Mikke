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

/**
 * lg* を通してよいのは「旧バインド型の学級」だけ。
 *
 * ⚠️ ここが無いと、コンテナバインド版（Bound.gs）の学級で穴になる。
 *    google.script.run は末尾 `_` の無い関数を誰でも直接呼べるので、児童が
 *    ブラウザのコンソールから lgGetInitData() を呼ぶと、
 *      1. Users_名簿 シートが作られ
 *      2. 「名簿が空なら最初の人を先生にする」で **その児童が先生として登録され**
 *      3. そのまま lgExecuteAction で単元の作成・削除ができる
 *    という経路が通る。Bound.gs の ownerEmail の判定を丸ごと迂回する。
 *    （実測: この関門を外すと、児童が学級の単元を作れることをテストで再現できる）
 *
 * 判定は Main.gs の boundModeFor_ に合わせる。中身がまだ何も無いファイルは
 * 'bound' 扱いなので、**空のシートから lg* で先生を作ることはできない**。
 * 旧バインド型として動いている学級（Users_名簿 に行がある）だけが通る。
 */
function assertLegacyContainer_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }
  if (!ss || boundModeFor_(ss) !== 'legacy') {
    throw new Error('FORBIDDEN: この入口はこの学級では使えません' +
      '（貼り付けで入れた旧版の学級のためのものです）');
  }
  return ss;
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
    const ss = assertLegacyContainer_();
    const email = lgEmail_();
    lgEnsureLegacySheets_(ss);

    let users = getTableData_(ss, TABLES.USERS);
    let myUser = users.filter(function (u) { return String(u.email).toLowerCase() === email; })[0];

    if (users.length === 0) {
      // 最初にアクセスした人を先生として登録（旧版の挙動を維持）
      const newTeacher = { email: email, name: '先生', group_id: 'teacher', role: 'teacher', created_at: new Date().toLocaleString() };
      appendRowLocked_(ss, TABLES.USERS, newTeacher);
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
    // 兄弟の lg* はすべて lgEmail_() で本人確認しているのに、ここだけ抜けていた。
    // google.script.run は末尾 `_` の無い関数を誰でも直接呼べるので、
    // 1 本の抜けで「名簿に載っていない人がクラスの記録を全部引ける」になる。
    const ss = assertLegacyContainer_();
    const email = lgEmail_();
    lgEnsureLegacySheets_(ss);
    if (!lgIsMember_(ss, email)) {
      throw new Error('NOT_MEMBER: このクラスの名簿に登録されていません。先生に確認してください');
    }
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

function lgUserRow_(ss, email) {
  return getTableData_(ss, TABLES.USERS)
    .filter(function (x) { return String(x.email).toLowerCase() === email; })[0] || null;
}

function lgIsTeacher_(ss, email) {
  const u = lgUserRow_(ss, email);
  return !!u && u.role === 'teacher';
}

/** 名簿に載っているか。名簿が空（＝初期化前）のときは lgGetInitData 側で先生として登録される */
function lgIsMember_(ss, email) {
  return !!lgUserRow_(ss, email);
}

function lgExecuteAction(payloadJson) {
  try {
    const ss = assertLegacyContainer_();
    const email = lgEmail_();
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
      const existing = getTableData_(ss, TABLES.USERS).map(function (u) { return String(u.email).toLowerCase(); });
      (p.users || []).slice(0, 200).forEach(function (u) {
        if (!u || !u.email) return;
        const em = String(u.email).trim().toLowerCase();
        if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(em)) return;
        if (existing.indexOf(em) >= 0) return;
        existing.push(em);
        appendRowLocked_(ss, TABLES.USERS, {
          email: em,
          name: safeCellText_(vStr_(u.name, 30, '氏名').trim()),
          group_id: safeCellText_(vStr_(u.group_id, 20, '班').trim()),
          role: 'student',
          created_at: new Date()
        });
      });
      return jsonOk_({});
    }
    if (p.action === 'save_api_key') {
      setSettingValue_(ss, 'geminiApiKey', vStr_(p.api_key, 200, 'APIキー').trim());
      return jsonOk_({});
    }

    // 単元管理系はすべて共有コア（Tenant.gs の coreUnitAction_）に任せる。
    // 同じ処理を 2 本持っていたころは、片方だけ直して「先生の画面によって
    // 挙動が違う」が起きていた。
    if (['save_unit', 'add_map', 'toggle_chat', 'toggle_stamp', 'update_custom_stamps'].indexOf(p.action) >= 0) {
      return jsonOk_(coreUnitAction_(ss, p));
    }
    throw new Error('BAD_INPUT: 不明な操作です');
  } catch (e) { return jsonErr_(e); }
}

function lgGenerateAIPortfolio(payloadJson) {
  try {
    const ss = assertLegacyContainer_();
    const email = lgEmail_();
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
    // 組み立ては corePortfolioPrompt_（Tenant.gs）に 1 本化してある。
    const built = corePortfolioPrompt_(
      getTableData_(ss, TABLES.USERS).map(function (u) { return u.name; }),
      user ? user.name : '', pins, chats, reactions);

    return jsonOk_({ portfolio: coreRunPortfolio_(apiKey, built) });
  } catch (e) { return jsonErr_(e); }
}

function lgUploadImage(dataUrl) {
  try {
    const ss = assertLegacyContainer_();
    const email = lgEmail_();
    return jsonOk_({ imageRef: storeImage_(ss, email, dataUrl) });
  } catch (e) { return jsonErr_(e); }
}

function lgGetImage(imageRef) {
  try {
    const ss = assertLegacyContainer_();
    lgEmail_();
    return jsonOk_({ dataUrl: loadImage_(ss, imageRef) });
  } catch (e) { return jsonErr_(e); }
}

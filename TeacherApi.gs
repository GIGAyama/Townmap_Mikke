/**
 * TeacherApi.gs — デプロイ T（教員ポータル）用 API
 *
 * T は「アクセスしているユーザーとして実行」なので、本人特定は Session.getActiveUser()。
 * すべての API は { success:true, ... } / { success:false, error } の JSON 文字列を返す。
 * クラスを操作する API は冒頭で「そのクラスの ownerEmail が自分か」を必ず検証する。
 */

/** クラス所有者チェック（tp* 共通ガード） */
function assertOwner_(classCode) {
  const email = teacherEmail_();
  const code = normalizeClassCode_(classCode);
  const rec = getClassRecord_(code);
  if (!rec || rec.revoked) {
    throw new Error('CLASS_NOT_FOUND: このクラスは存在しないか、すでに廃止されています');
  }
  if (String(rec.ownerEmail).toLowerCase() !== email) {
    throw new Error('FORBIDDEN: このクラスを管理できるのは作成した先生だけです');
  }
  return { email: email, code: code, rec: rec };
}

function studentUrlFor_(code) {
  const shell = getShellUrl_();
  return shell ? shell + '?c=' + code : '';
}

function activeStudentCount_(ss) {
  return getMembers_(ss).filter(function (m) {
    return m.status === 'active' && m.role !== 'teacher';
  }).length;
}

function upsertMember_(ss, m) {
  withScriptLock_(function () {
    const sheet = ensureSheet_(ss, TABLES.MEMBERS);
    const data = sheet.getDataRange().getValues();
    const H = headerMapFromRow_(data[0], TABLES.MEMBERS);
    requireCols_(H, TABLES.MEMBERS, ['email', 'displayName', 'role', 'status', 'number', 'groupId', 'joinedAt']);
    const target = String(m.email).toLowerCase();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][H.email]).toLowerCase() === target) {
        const patch = { email: target };
        if (m.displayName !== undefined) patch.displayName = safeCellText_(m.displayName);
        if (m.role) patch.role = m.role;
        if (m.status) patch.status = m.status;
        if (m.number !== undefined) patch.number = safeCellText_(m.number);
        if (m.groupId !== undefined) patch.groupId = safeCellText_(m.groupId);
        if (!data[i][H.joinedAt]) patch.joinedAt = new Date();
        const row = setCells_(data[i].slice(), H, TABLES.MEMBERS, patch);
        sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
        return;
      }
    }
    sheet.appendRow(rowFor_(H, TABLES.MEMBERS, {
      email: target,
      displayName: safeCellText_(m.displayName || ''),
      role: m.role || 'student',
      status: m.status || 'active',
      number: safeCellText_(m.number || ''),
      groupId: safeCellText_(m.groupId || ''),
      joinedAt: new Date()
    }, data[0].length));
  });
}

function setMemberStatus_(ss, emails, status) {
  const targets = emails.map(function (e) { return String(e).toLowerCase(); });
  withScriptLock_(function () {
    const sheet = ensureSheet_(ss, TABLES.MEMBERS);
    const data = sheet.getDataRange().getValues();
    const H = headerMapFromRow_(data[0], TABLES.MEMBERS);
    requireCols_(H, TABLES.MEMBERS, ['email', 'status']);
    for (let i = 1; i < data.length; i++) {
      if (targets.indexOf(String(data[i][H.email]).toLowerCase()) >= 0) {
        sheet.getRange(i + 1, H.status + 1).setValue(status);
      }
    }
  });
}

// ────────────────────────────────────────────────────────────────
// ポータル（クラス管理）
// ────────────────────────────────────────────────────────────────

/** 自分のクラス一覧。studentUrl は共通 URL（GitHub Pages）+ ?c=コード */
function tpGetMyPortal() {
  try {
    const email = teacherEmail_();
    const codes = listOwnedCodes_(email);
    const classes = [];
    codes.forEach(function (code) {
      const rec = getClassRecord_(code);
      if (!rec || rec.revoked) return;
      classes.push({
        classCode: code,
        className: rec.className,
        studentUrl: studentUrlFor_(code),
        spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/' + rec.spreadsheetId + '/edit',
        memberCount: rec.memberCount || 0,
        joinOpen: rec.joinOpen !== false,
        requireApproval: rec.requireApproval !== false,
        createdAt: rec.createdAt
      });
    });
    return jsonOk_({ classes: classes, teacherEmail: email, shellUrl: getShellUrl_() });
  } catch (e) { return jsonErr_(e); }
}

/**
 * クラス作成。教員本人の実行なのでシートは最初から教員所有になる。
 * DriveApp.makeCopy は使わない（フル Drive スコープ回避）。テンプレートは
 * SpreadsheetApp.openById(templateId).copy() で複製する（spreadsheets スコープで動く）。
 */
function tpCreateClass(className) {
  const userLock = LockService.getUserLock();
  try {
    const email = teacherEmail_();
    const name = vStr_(className, 50, 'クラス名').trim();
    if (!name) throw new Error('BAD_INPUT: クラス名を入力してください');
    const appAccount = String(getSetting_(PROP_KEYS.APP_ACCOUNT, true)).toLowerCase();

    if (!userLock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
      throw new Error('LOCK_BUSY: 処理が混み合っています。数秒待ってもう一度お試しください');
    }

    // 1. シート生成（テンプレートがあれば複製、無ければ新規作成 + 初期化）
    const templateId = getSetting_(PROP_KEYS.TEMPLATE, false);
    const title = CONFIG.APP_NAME + '_' + name;
    let ss;
    if (templateId) {
      ss = SpreadsheetApp.openById(templateId).copy(title);
      ensureClassSheets_(ss);
    } else {
      ss = SpreadsheetApp.create(title);
      initializeNewDatabase_(ss);
    }

    // 2. アプリアカウントを編集者に自動追加（この構成の心臓部）。
    //    これにより児童用デプロイ S（アプリアカウント実行）がこのシートに読み書きできる。
    //    児童自身には権限を一切与えない。
    try {
      ss.addEditor(appAccount);
    } catch (shareErr) {
      // 巻き戻し: Drive スコープを持たないためゴミ箱移動はできない。
      // 名前で失敗が分かるようにし、レジストリには登録しない（クラスとして成立させない）。
      try { ss.rename('【作成失敗・削除してください】' + title); } catch (e2) { /* ignore */ }
      throw new Error('SHARE_FAILED: クラス用シートは作成できましたが、アプリへの共有に失敗しました。' +
        '学校の Google Workspace で外部共有が制限されている可能性があります。' +
        '学校の管理者に「' + appAccount + ' への共有許可」を確認してください。' +
        '（ドライブに残った「【作成失敗・削除してください】」のシートは削除して構いません）');
    }

    // 3. コード発行 → レジストリ登録 → Members に自分を teacher/active で登録
    let code = null;
    withScriptLock_(function () {
      code = generateClassCode_();
      putClassRecord_(code, {
        spreadsheetId: ss.getId(),
        ownerEmail: email,
        className: name,
        createdAt: new Date().toISOString(),
        joinOpen: true,
        requireApproval: true,  // クラスコードは宛先であって認証ではない。既定は承認制
        revoked: false,
        memberCount: 0
      });
    });
    addOwnedCode_(email, code);
    writeMeta_(ss, {
      schemaVersion: CONFIG.SCHEMA_VERSION,
      classCode: code,
      className: name,
      ownerEmail: email,
      createdAt: new Date().toISOString()
    });
    upsertMember_(ss, { email: email, displayName: '先生', role: 'teacher', status: 'active' });

    return jsonOk_({
      classCode: code,
      studentUrl: studentUrlFor_(code),
      spreadsheetUrl: ss.getUrl()
    });
  } catch (e) {
    return jsonErr_(e);
  } finally {
    try { userLock.releaseLock(); } catch (e) { /* not held */ }
  }
}

/** URL / 生 ID のどちらからでもスプレッドシート ID を抽出 */
function extractSpreadsheetId_(input) {
  const s = String(input || '').trim();
  const m = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9\-_]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9\-_]{20,}$/.test(s)) return s;
  throw new Error('BAD_INPUT: スプレッドシートの URL または ID を入力してください');
}

/** 既存シートのクラス化 */
function tpRegisterExisting(input, className) {
  try {
    const email = teacherEmail_();
    const name = vStr_(className, 50, 'クラス名').trim();
    if (!name) throw new Error('BAD_INPUT: クラス名を入力してください');
    const appAccount = String(getSetting_(PROP_KEYS.APP_ACCOUNT, true)).toLowerCase();
    const ssId = extractSpreadsheetId_(input);

    let ss;
    try {
      ss = SpreadsheetApp.openById(ssId);
    } catch (e) {
      throw new Error('BAD_INPUT: スプレッドシートを開けません。自分が編集できるシートの URL を指定してください');
    }

    ensureClassSheets_(ss); // 不足シートの補完 + 旧名簿(Users_名簿)の移行

    try {
      ss.addEditor(appAccount);
    } catch (shareErr) {
      throw new Error('SHARE_FAILED: アプリへの共有に失敗しました。学校の管理者に「' +
        appAccount + ' への共有許可」を確認してください');
    }

    let code = null;
    withScriptLock_(function () {
      code = generateClassCode_();
      putClassRecord_(code, {
        spreadsheetId: ss.getId(),
        ownerEmail: email,
        className: name,
        createdAt: new Date().toISOString(),
        joinOpen: true,
        requireApproval: true,
        revoked: false,
        memberCount: 0
      });
    });
    // own_ とレジストリは読み→書きの2手なので、ロックの外で走らせると
    // 並行実行と混ざって片方の更新が消える。ここも必ずロックで包む。
    withScriptLock_(function () { addOwnedCode_(email, code); });
    writeMeta_(ss, {
      schemaVersion: CONFIG.SCHEMA_VERSION,
      classCode: code,
      className: name,
      ownerEmail: email,
      createdAt: new Date().toISOString()
    });
    upsertMember_(ss, { email: email, displayName: '先生', role: 'teacher', status: 'active' });
    withScriptLock_(function () { updateClassRecord_(code, { memberCount: activeStudentCount_(ss) }); });

    return jsonOk_({
      classCode: code,
      studentUrl: studentUrlFor_(code),
      spreadsheetUrl: ss.getUrl()
    });
  } catch (e) { return jsonErr_(e); }
}

/** クラス管理コンソール用: 名簿（メールアドレスを含む。教員のみが見る） */
function tpListMembers(classCode) {
  try {
    const g = assertOwner_(classCode);
    const ss = openClassSs_(g.code);
    const members = getMembers_(ss).map(function (m) {
      return {
        email: m.email,
        displayName: m.displayName,
        role: m.role,
        status: m.status,
        number: m.number,
        groupId: m.groupId,
        joinedAt: m.joinedAt
      };
    });
    return jsonOk_({
      members: members,
      classCode: g.code,
      className: g.rec.className,
      joinOpen: g.rec.joinOpen !== false,
      requireApproval: g.rec.requireApproval !== false,
      studentUrl: studentUrlFor_(g.code),
      spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/' + g.rec.spreadsheetId + '/edit'
    });
  } catch (e) { return jsonErr_(e); }
}

/** 参加承認。Members の status を active に変えるだけ（シート共有は不要。児童は権限を持たないため） */
function tpApprove(classCode, emails) {
  try {
    const g = assertOwner_(classCode);
    if (!emails || !emails.length) throw new Error('BAD_INPUT: 対象を選択してください');
    const ss = openClassSs_(g.code);
    setMemberStatus_(ss, emails, 'active');
    withScriptLock_(function () { updateClassRecord_(g.code, { memberCount: activeStudentCount_(ss) }); });
    return jsonOk_({});
  } catch (e) { return jsonErr_(e); }
}

function tpRemove(classCode, email) {
  try {
    const g = assertOwner_(classCode);
    const target = String(email || '').toLowerCase();
    if (!target) throw new Error('BAD_INPUT: 対象を指定してください');
    if (target === g.email) throw new Error('BAD_INPUT: 自分自身は削除できません');
    const ss = openClassSs_(g.code);
    setMemberStatus_(ss, [target], 'removed');
    withScriptLock_(function () { updateClassRecord_(g.code, { memberCount: activeStudentCount_(ss) }); });
    return jsonOk_({});
  } catch (e) { return jsonErr_(e); }
}

function tpSetJoinOpen(classCode, isOpen) {
  try {
    const g = assertOwner_(classCode);
    withScriptLock_(function () { updateClassRecord_(g.code, { joinOpen: isOpen === true }); });
    return jsonOk_({ joinOpen: isOpen === true });
  } catch (e) { return jsonErr_(e); }
}

function tpSetRequireApproval(classCode, required) {
  try {
    const g = assertOwner_(classCode);
    withScriptLock_(function () { updateClassRecord_(g.code, { requireApproval: required === true }); });
    return jsonOk_({ requireApproval: required === true });
  } catch (e) { return jsonErr_(e); }
}

/** コード再発行: 旧 cls_ を revoked:true にして新コードを発行し、own_ を差し替える */
function tpRotateCode(classCode) {
  try {
    const g = assertOwner_(classCode);
    let newCode = null;
    withScriptLock_(function () {
      newCode = generateClassCode_();
      const rec = getClassRecord_(g.code);
      // フィールドを1つずつ写すとレコードに項目が増えたとき取りこぼす。
      // 丸ごと引き継いで revoked だけ戻す（reflection 方式）。
      const next = JSON.parse(JSON.stringify(rec));
      next.revoked = false;
      putClassRecord_(newCode, next);
      updateClassRecord_(g.code, { revoked: true });
    });
    withScriptLock_(function () {
      removeOwnedCode_(g.email, g.code);
      addOwnedCode_(g.email, newCode);
    });
    try { setMeta_(openClassSs_(newCode), 'classCode', newCode); } catch (e) { /* meta は補助情報 */ }
    return jsonOk_({ classCode: newCode, studentUrl: studentUrlFor_(newCode) });
  } catch (e) { return jsonErr_(e); }
}

/** クラス廃止（レジストリから外すのみ。シートは教員の手元に残る） */
function tpRevokeClass(classCode) {
  try {
    const g = assertOwner_(classCode);
    withScriptLock_(function () {
      updateClassRecord_(g.code, { revoked: true });
      removeOwnedCode_(g.email, g.code);
    });
    return jsonOk_({});
  } catch (e) { return jsonErr_(e); }
}

// ────────────────────────────────────────────────────────────────
// アプリ本体（教員としての閲覧・管理）
// ────────────────────────────────────────────────────────────────

function tpBuildInitData_(g, ss) {
  const membersRaw = getMembers_(ss);
  const me = membersRaw.filter(function (m) {
    return String(m.email).toLowerCase() === g.email;
  })[0];
  const units = getTableData_(ss, TABLES.UNITS);
  const activeUnit = formatUnit_(units.filter(function (u) { return u.is_active === true; })[0] || units[0] || null);
  const data = activeUnit ? coreCollectUnitData_(ss, activeUnit.unit_id) : { pins: [], chats: [], reactions: [] };
  return {
    user: {
      email: uidOf_(g.email),
      name: (me && me.displayName) || '先生',
      group_id: (me && me.groupId) || 'teacher',
      role: 'teacher'
    },
    users: sanitizeMembers_(membersRaw),
    units: units.map(function (u) { return formatUnit_(u); }),
    activeUnit: activeUnit,
    pins: sanitizeRecords_(data.pins),
    chats: sanitizeRecords_(data.chats),
    reactions: sanitizeRecords_(data.reactions),
    hasApiKey: !!getSettingValue_(ss, 'geminiApiKey')
  };
}

function tpGetInitData(classCode) {
  try {
    const g = assertOwner_(classCode);
    const ss = openClassSs_(g.code);
    return jsonOk_(tpBuildInitData_(g, ss));
  } catch (e) { return jsonErr_(e); }
}

function tpSyncData(classCode, unitId) {
  try {
    const g = assertOwner_(classCode);
    const ss = openClassSs_(g.code);
    const units = getTableData_(ss, TABLES.UNITS);
    // 現在アクティブな単元を優先して返す（単元切替の自動追従用。StudentApi と同じ挙動）
    const activeUnit = formatUnit_(
      units.filter(function (u) { return u.is_active === true; })[0] ||
      units.filter(function (u) { return u.unit_id === unitId; })[0] || null);
    const data = coreCollectUnitData_(ss, unitId);
    return jsonOk_({
      pins: sanitizeRecords_(data.pins),
      chats: sanitizeRecords_(data.chats),
      reactions: sanitizeRecords_(data.reactions),
      activeUnit: activeUnit,
      users: sanitizeMembers_(getMembers_(ss)),
      hasApiKey: !!getSettingValue_(ss, 'geminiApiKey')
    });
  } catch (e) { return jsonErr_(e); }
}

/** 教員の操作（単元管理・名簿一括登録・チャット・削除など） */
function tpExecuteAction(classCode, payloadJson) {
  try {
    const g = assertOwner_(classCode);
    const ss = openClassSs_(g.code);
    const p = JSON.parse(payloadJson);
    const email = g.email;

    if (p.action === 'save_chat') {
      return jsonOk_(coreSaveChat_(ss, email, p));
    }
    if (p.action === 'toggle_reaction') {
      return jsonOk_(coreToggleReaction_(ss, email, p));
    }
    if (p.action === 'save_pin') {
      return jsonOk_(coreSavePin_(ss, email, p));
    }
    if (p.action === 'update_pin') {
      // 自分のピンのみ更新可（coreUpdateOwnPin_ が行所有者チェックを行う）
      return jsonOk_(coreUpdateOwnPin_(ss, email, p.pin_id, p));
    }
    if (p.action === 'delete_pin' || p.action === 'delete_chat') {
      // 教員は誰の記録でも削除できる（ownOnly=false）
      return jsonOk_(coreDeleteRecord_(ss, email, p.pin_id || p.chat_id, false));
    }
    if (['save_unit', 'add_map', 'toggle_chat', 'toggle_stamp', 'update_custom_stamps'].indexOf(p.action) >= 0) {
      return jsonOk_(coreUnitAction_(ss, p));
    }
    if (p.action === 'save_users') {
      // 名簿の一括事前登録（承認不要で active になる）
      (p.users || []).slice(0, 200).forEach(function (u) {
        if (!u || !u.email) return;
        const em = String(u.email).trim().toLowerCase();
        if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(em)) return;
        upsertMember_(ss, {
          email: em,
          displayName: vStr_(u.name, 30, '氏名').trim(),
          groupId: vStr_(u.group_id, 20, '班').trim(),
          role: 'student',
          status: 'active'
        });
      });
      withScriptLock_(function () { updateClassRecord_(g.code, { memberCount: activeStudentCount_(ss) }); });
      return jsonOk_({});
    }
    if (p.action === 'save_api_key') {
      setSettingValue_(ss, 'geminiApiKey', vStr_(p.api_key, 200, 'APIキー').trim());
      return jsonOk_({});
    }
    throw new Error('BAD_INPUT: 不明な操作です');
  } catch (e) { return jsonErr_(e); }
}

/** AI ポートフォリオ生成（教員のみ）。対象児童は uid で指定し、サーバー側で解決する */
function tpGenerateAIPortfolio(classCode, payloadJson) {
  try {
    const g = assertOwner_(classCode);
    const ss = openClassSs_(g.code);
    const p = JSON.parse(payloadJson);
    const apiKey = getSettingValue_(ss, 'geminiApiKey');
    if (!apiKey) {
      throw new Error('NO_API_KEY: AI分析を行うには、管理パネルの「AI設定」タブで Gemini API キーを設定してください。');
    }

    const member = getMembers_(ss).filter(function (m) { return uidOf_(m.email) === p.uid; })[0];
    if (!member) throw new Error('NOT_FOUND: 対象の児童が見つかりません');
    const targetEmail = String(member.email).toLowerCase();

    const pins = getTableData_(ss, TABLES.PINS).filter(function (pin) {
      return pin.unit_id === p.unit_id && String(pin.email).toLowerCase() === targetEmail;
    });
    const chats = getTableData_(ss, TABLES.CHATS).filter(function (chat) {
      return chat.unit_id === p.unit_id && String(chat.email).toLowerCase() === targetEmail;
    });
    const reactions = getTableData_(ss, TABLES.REACTIONS).filter(function (r) {
      return r.unit_id === p.unit_id && String(r.email).toLowerCase() === targetEmail;
    });

    if (pins.length === 0 && chats.length === 0 && reactions.length === 0) {
      return jsonOk_({ portfolio: 'まだ活動の記録（ピンやチャット）がありません。' });
    }

    // AI には実名（表示名）を渡さない。対象児童は「対象児童」、ほかの子は「児童A」…に置き換え、
    // ピンのメモやチャット本文に書かれた友達の名前・連絡先もまとめて消してから送る。
    // 組み立ては corePortfolioPrompt_（Tenant.gs）に 1 本化してある。
    const built = corePortfolioPrompt_(
      getMembers_(ss).map(function (m) { return m.displayName; }),
      member.displayName, pins, chats, reactions);

    return jsonOk_({ portfolio: coreRunPortfolio_(apiKey, built) });
  } catch (e) { return jsonErr_(e); }
}

/** 画像アップロード（教員: 地図背景など） */
function tpUploadImage(classCode, dataUrl) {
  try {
    const g = assertOwner_(classCode);
    const ss = openClassSs_(g.code);
    return jsonOk_({ imageRef: storeImage_(ss, g.email, dataUrl) });
  } catch (e) { return jsonErr_(e); }
}

function tpGetImage(classCode, imageRef) {
  try {
    const g = assertOwner_(classCode);
    const ss = openClassSs_(g.code);
    return jsonOk_({ dataUrl: loadImage_(ss, imageRef) });
  } catch (e) { return jsonErr_(e); }
}

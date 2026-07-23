/**
 * StudentApi.gs — デプロイ S（児童用アプリ）用 API
 *
 * すべての API は第 1 引数に ID トークンを受け取り、guardStudent_（Tenant.gs）で
 * ① トークン検証 → ② クラス解決 → ③ 名簿照合（status=active） を通してから処理する。
 * 書き込み者 email は常に「検証済みトークンの email」。クライアント申告の email は
 * どこにも使わない。レスポンスに他児童の email・スプレッドシート ID は含めない。
 */

// ────────────────────────────────────────────────────────────────
// 参加フロー
// ────────────────────────────────────────────────────────────────

/**
 * 自分の状態を確認する（名簿照合の前段なので guardStudent_ は使わない）。
 * state: 'active' | 'pending' | 'unregistered' | 'closed'
 */
function stGetStatus(idToken, classCode) {
  try {
    const user = verifyIdToken_(idToken);
    const code = normalizeClassCode_(classCode);
    const rec = requireClassRecord_(code);
    const ss = openClassSs_(code);
    const member = getMemberRow_(ss, user.email);

    let state;
    if (member && member.status === 'active') state = 'active';
    else if (member && member.status === 'pending') state = 'pending';
    else if (rec.joinOpen === false) state = 'closed';
    else state = 'unregistered';

    return jsonOk_({
      state: state,
      className: rec.className,
      displayName: member ? member.displayName : '',
      role: member && member.role === 'teacher' ? 'teacher' : 'student',
      requireApproval: rec.requireApproval !== false
    });
  } catch (e) { return jsonErr_(e); }
}

/**
 * クラス参加申請。表示名は氏名でなく出席番号・ニックネームでもよい。
 * requireApproval が true（既定）なら pending、false なら active。
 * 同一 email の重複登録は既存行の更新にする。
 */
function stJoin(idToken, classCode, displayName, number) {
  try {
    const user = verifyIdToken_(idToken);
    const code = normalizeClassCode_(classCode);
    const rec = requireClassRecord_(code);
    if (rec.joinOpen === false) {
      throw new Error('JOIN_CLOSED: いまは参加の受付が閉じられています。先生に確認してください');
    }
    const name = vStr_(displayName, 30, '表示名').trim();
    if (!name) throw new Error('BAD_INPUT: 表示名を入力してください');
    const num = vStr_(number, 10, '出席番号').trim();
    const ss = openClassSs_(code);
    const status = rec.requireApproval === false ? 'active' : 'pending';

    withScriptLock_(function () {
      const sheet = ensureSheet_(ss, TABLES.MEMBERS);
      const data = sheet.getDataRange().getValues();
      const target = user.email;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).toLowerCase() === target) {
          // 既存行の更新（active の行を pending に格下げしない）
          const nextStatus = data[i][3] === 'active' ? 'active' : status;
          sheet.getRange(i + 1, 1, 1, TABLES.MEMBERS.cols.length).setValues([[
            target, name, data[i][2] || 'student', nextStatus, num, data[i][5] || '', data[i][6] || new Date()
          ]]);
          return;
        }
      }
      sheet.appendRow([target, name, 'student', status, num, '', new Date()]);
    });

    if (status === 'active') {
      try { updateClassRecord_(code, { memberCount: activeStudentCount_(ss) }); } catch (e) { /* 補助情報 */ }
    }
    return jsonOk_({ state: status });
  } catch (e) { return jsonErr_(e); }
}

// ────────────────────────────────────────────────────────────────
// データ取得
// ────────────────────────────────────────────────────────────────

function stBuildInitData_(g) {
  const ss = g.ss;
  const membersRaw = getMembers_(ss);
  const units = getTableData_(ss, TABLES.UNITS);
  const activeUnit = formatUnit_(units.filter(function (u) { return u.is_active === true; })[0] || units[0] || null);
  const data = activeUnit ? coreCollectUnitData_(ss, activeUnit.unit_id) : { pins: [], chats: [], reactions: [] };
  return {
    user: {
      email: uidOf_(g.user.email),
      name: g.member.displayName || '',
      group_id: g.member.groupId || '',
      role: g.member.role === 'teacher' ? 'teacher' : 'student'
    },
    users: sanitizeMembers_(membersRaw),
    activeUnit: activeUnit,
    pins: sanitizeRecords_(data.pins),
    chats: sanitizeRecords_(data.chats),
    reactions: sanitizeRecords_(data.reactions)
  };
}

function stGetInitData(idToken, classCode) {
  try {
    const g = guardStudent_(idToken, classCode);
    return jsonOk_(stBuildInitData_(g));
  } catch (e) { return jsonErr_(e); }
}

function stSyncData(idToken, classCode, unitId) {
  try {
    const g = guardStudent_(idToken, classCode);
    const units = getTableData_(g.ss, TABLES.UNITS);
    const activeUnit = formatUnit_(units.filter(function (u) { return u.unit_id === unitId; })[0] || null);
    const data = coreCollectUnitData_(g.ss, unitId);
    return jsonOk_({
      pins: sanitizeRecords_(data.pins),
      chats: sanitizeRecords_(data.chats),
      reactions: sanitizeRecords_(data.reactions),
      activeUnit: activeUnit,
      users: sanitizeMembers_(getMembers_(g.ss))
    });
  } catch (e) { return jsonErr_(e); }
}

// ────────────────────────────────────────────────────────────────
// 書き込み（payload はホワイトリストしたキーのみ書き込む）
// ────────────────────────────────────────────────────────────────

/** 記録の追加。payload.type = 'pin' | 'chat' | 'reaction' */
function stSubmit(idToken, classCode, payload) {
  try {
    const g = guardStudent_(idToken, classCode);
    const p = payload || {};
    if (p.type === 'pin') return jsonOk_(coreSavePin_(g.ss, g.user.email, p));
    if (p.type === 'chat') {
      // チャット OFF の単元では一般メッセージを拒否（ピンへのコメントは可）
      if (p.target_type === 'general' || p.target_type === 'chat') {
        const unit = getTableData_(g.ss, TABLES.UNITS).filter(function (u) { return u.unit_id === p.unit_id; })[0];
        if (unit && unit.chat_enabled === false && g.member.role !== 'teacher') {
          throw new Error('FORBIDDEN: いまはチャットがオフになっています');
        }
      }
      return jsonOk_(coreSaveChat_(g.ss, g.user.email, p));
    }
    if (p.type === 'reaction') return jsonOk_(coreToggleReaction_(g.ss, g.user.email, p));
    throw new Error('BAD_INPUT: 不明な操作です');
  } catch (e) { return jsonErr_(e); }
}

/** 自分の記録一覧 */
function stListMine(idToken, classCode) {
  try {
    const g = guardStudent_(idToken, classCode);
    const me = g.user.email;
    const mine = function (r) { return String(r.email).toLowerCase() === me; };
    return jsonOk_({
      pins: sanitizeRecords_(getTableData_(g.ss, TABLES.PINS).filter(mine)),
      chats: sanitizeRecords_(getTableData_(g.ss, TABLES.CHATS).filter(mine)),
      reactions: sanitizeRecords_(getTableData_(g.ss, TABLES.REACTIONS).filter(mine))
    });
  } catch (e) { return jsonErr_(e); }
}

/** 自分のピンの更新（行所有者チェックはサーバー側で実施） */
function stUpdateMine(idToken, classCode, recordId, payload) {
  try {
    const g = guardStudent_(idToken, classCode);
    return jsonOk_(coreUpdateOwnPin_(g.ss, g.user.email, recordId, payload || {}));
  } catch (e) { return jsonErr_(e); }
}

/** 自分の記録の削除（行所有者チェックはサーバー側で実施） */
function stDeleteMine(idToken, classCode, recordId) {
  try {
    const g = guardStudent_(idToken, classCode);
    return jsonOk_(coreDeleteRecord_(g.ss, g.user.email, recordId, true));
  } catch (e) { return jsonErr_(e); }
}

// ────────────────────────────────────────────────────────────────
// 画像（Images_画像シート経由。Drive 権限もスプレッドシート ID も露出しない）
// ────────────────────────────────────────────────────────────────

function stUploadImage(idToken, classCode, dataUrl) {
  try {
    const g = guardStudent_(idToken, classCode);
    return jsonOk_({ imageRef: storeImage_(g.ss, g.user.email, dataUrl) });
  } catch (e) { return jsonErr_(e); }
}

function stGetImage(idToken, classCode, imageRef) {
  try {
    const g = guardStudent_(idToken, classCode);
    return jsonOk_({ dataUrl: loadImage_(g.ss, imageRef) });
  } catch (e) { return jsonErr_(e); }
}

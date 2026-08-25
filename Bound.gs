/**
 * Bound.gs — コンテナバインド版（先生ごとに 1 デプロイ）の API
 *
 * ── この形の配り方 ──────────────────────────────────────────────
 *   1. 先生が「スプレッドシートのコピー」を作る（スクリプトも一緒についてくる）
 *   2. 「拡張機能 ＞ Apps Script ＞ デプロイ」でウェブアプリを 1 本公開する
 *   3. 出てきた URL を学級に配る
 * データはそのスプレッドシートの中だけにあり、作者にも他の先生にも見えない。
 * クラスコードもレジストリも共通シェルも要らない。
 *
 * ── 前提にしているデプロイ設定（appsscript.json）────────────────────
 *   次のユーザーとして実行:     自分（USER_DEPLOYING）
 *   アクセスできるユーザー:     同一ドメインの全員（DOMAIN）
 *
 *   「自分として実行」にすると、シートの読み書きはすべて先生の権限で走るので、
 *   **児童はスプレッドシートへのアクセス権を 1 つも持たなくてよい**（＝児童が
 *   シートを直接開けない、onOpen を動かせない）。
 *
 *   代わりに Session.getActiveUser().getEmail()（＝開いている本人）は
 *   **アクセスできるユーザーが「同一ドメインの全員」のときしか取れない**。
 *   「全員」に開くと空文字になり、誰が誰だか分からなくなる。そのときは
 *   BOUND_NO_IDENTITY として、直し方を書いたエラーで止める（誰も通さない）。
 *
 * ── 先生の判定（ここを間違えると学級全員が先生になる）────────────────
 *   `Session.getEffectiveUser()` は「自分として実行」では**誰が開いてもデプロイした先生**を
 *   返す。これを本人確認に使うと、その瞬間に全員が先生として通る。だから使わない。
 *
 *   先生は「Settings シートの ownerEmail に記録された人」と「名簿の role=teacher」だけ。
 *   ownerEmail を書くのは次の 2 つの経路だけで、どちらも「最初に画面を開いた児童が
 *   先生になる」穴を作らない:
 *     (a) スプレッドシートのメニュー「初期設定」… メニューは**そのファイルの編集権を
 *         持つ人にしか出せない**ので、押せるのは先生だけ。
 *     (b) ウェブアプリ側の控えめな自動登録… `active !== effective` が観測できたとき
 *         **だけ**。これが観測できた時点で「自分として実行」が確定するので、
 *         effective はデプロイした先生だと言い切れる（boundAutoClaimOwner_）。
 */

/** Settings シートに置く設定のキー */
const BOUND_KEYS = {
  OWNER: 'ownerEmail',
  CLASS_NAME: 'className',
  JOIN_OPEN: 'joinOpen',
  REQUIRE_APPROVAL: 'requireApproval'
};

/** 'false' / false / 'いいえ' を偽として読む。空欄は既定値 */
function boundFlag_(v, dflt) {
  if (v === '' || v === null || v === undefined) return dflt;
  if (v === true || v === false) return v;
  const s = String(v).trim().toLowerCase();
  if (s === 'false' || s === 'no' || s === '0' || s === 'いいえ') return false;
  if (s === 'true' || s === 'yes' || s === '1' || s === 'はい') return true;
  return dflt;
}

/** バインドされたスプレッドシート。取れなければ「この形では動かない」ことを言って止める */
function boundSs_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }
  if (!ss) {
    throw new Error('NOT_BOUND: このスクリプトはスプレッドシートに束ねられていません。' +
      '配布用のスプレッドシートをコピーして、その中の Apps Script から公開してください');
  }
  return ss;
}

/**
 * 開いている本人のメールアドレス。
 * 「アクセスできるユーザー」が同一ドメインの全員でないと空になるので、
 * そのときは直し方つきで止める（推測で誰かにする、は絶対にしない）。
 */
function boundEmail_() {
  let email = '';
  try { email = String(Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) { email = ''; }
  if (!email) {
    throw new Error('BOUND_NO_IDENTITY: だれが開いているかを確認できません。' +
      '（先生へ: Apps Script の「デプロイを管理」で、アクセスできるユーザーを' +
      '「同一ドメインの全員」に、次のユーザーとして実行を「自分」にしてください）');
  }
  return email;
}

/** Settings に記録された先生。書き込みは boundSetOwner_ の 2 経路だけ */
function boundOwner_(ss) {
  return String(getSettingValue_(ss, BOUND_KEYS.OWNER) || '').trim().toLowerCase();
}

/**
 * 共通ガード。返す { ss, email, owner, member, isTeacher } を各 API が使う。
 * ロック外で呼ぶこと（読み取りだけなのでロックは要らない）。
 */
function guardBound_() {
  const ss = boundSs_();
  const email = boundEmail_();
  const owner = boundOwner_(ss) || boundAutoClaimOwner_(ss, email);
  const member = getMemberRow_(ss, email);
  const isTeacher = (!!owner && owner === email) || (!!member && member.role === 'teacher');
  return { ss: ss, email: email, owner: owner, member: member, isTeacher: isTeacher };
}

/** 先生だけが通れるガード */
function guardBoundTeacher_() {
  const g = guardBound_();
  if (!g.isTeacher) throw new Error('FORBIDDEN: この操作は先生だけができます');
  return g;
}

/** 名簿に載っていて active な人だけが通れるガード（先生は名簿が無くても通す） */
function guardBoundActive_() {
  const g = guardBound_();
  if (g.isTeacher) return g;
  if (!g.member || g.member.status !== 'active') {
    throw new Error('NOT_MEMBER: このクラスの名簿に登録されていません。先生に確認してください');
  }
  return g;
}

/** 先生として記録する（すでに記録があれば何もしない）。呼ぶ側が「その人でよい」ことを保証する */
function boundSetOwner_(ss, email) {
  const current = boundOwner_(ss);
  if (current) return current;
  const target = String(email || '').trim().toLowerCase();
  if (!target) return '';
  setSettingValue_(ss, BOUND_KEYS.OWNER, target);
  // 名簿にも先生として載せておく（管理画面が名簿越しに役割を見るため）
  try { upsertMember_(ss, { email: target, displayName: '先生', role: 'teacher', status: 'active' }); } catch (e) { /* 名簿は後からでも作れる */ }
  return target;
}

/**
 * メニュー「初期設定」から呼ぶ先生の登録。
 * スプレッドシートのメニューは**そのファイルの編集権を持つ人にしか出せない**ので、
 * ここで登録される人は必ず先生（コピーを作った本人）になる。
 */
function boundClaimOwnerFromContainer_(ss) {
  let email = '';
  try { email = String(Session.getEffectiveUser().getEmail() || '').toLowerCase(); } catch (e) { email = ''; }
  return boundSetOwner_(ss, email);
}

/**
 * ウェブアプリ側からの控えめな自動登録。
 *
 * `Session.getEffectiveUser()` は、
 *   ・「自分として実行（USER_DEPLOYING）」なら **誰が開いてもデプロイした先生**
 *   ・「アクセスしているユーザーとして実行（USER_ACCESSING）」なら **開いた本人**
 * を返す。後者でこれを先生とみなすと、**学級全員が先生になる**。
 *
 * そこで、`active !== effective` が観測できたときだけ登録する。これが観測できた時点で
 * 「自分として実行」であることが確定し、effective はデプロイした先生だと言い切れる。
 * 観測できない間（先生自身しか開いていない間）は登録せず、'setup' のまま待つ。
 *
 * @return {string} 登録した／すでに登録されている先生のメール。まだなら ''
 */
function boundAutoClaimOwner_(ss, activeEmail) {
  const current = boundOwner_(ss);
  if (current) return current;
  let effective = '';
  try { effective = String(Session.getEffectiveUser().getEmail() || '').toLowerCase(); } catch (e) { effective = ''; }
  if (!effective || !activeEmail || effective === activeEmail) return '';
  return boundSetOwner_(ss, effective);
}

// ────────────────────────────────────────────────────────────────
// 参加フロー
// ────────────────────────────────────────────────────────────────

/**
 * 自分の状態を確認する（名簿照合の前段なので guardBound_ の active 判定は使わない）。
 * state: 'teacher' | 'active' | 'pending' | 'unregistered' | 'closed' | 'setup'
 */
function bdGetStatus() {
  try {
    const ss = boundSs_();
    const email = boundEmail_();
    const owner = boundOwner_(ss) || boundAutoClaimOwner_(ss, email);

    if (!owner) {
      // まだ先生が確定していない。ここで開いた人を先生にはしない（それが例の穴）。
      return jsonOk_({
        state: 'setup',
        className: String(getSettingValue_(ss, BOUND_KEYS.CLASS_NAME) || ''),
        message: '準備がまだ終わっていません。（先生へ: いちどこのアプリのスプレッドシートを開いて、' +
          '上に「' + CONFIG.APP_NAME + '」メニューが出れば準備完了です）'
      });
    }

    const member = getMemberRow_(ss, email);
    const isTeacher = owner === email || (!!member && member.role === 'teacher');
    const joinOpen = boundFlag_(getSettingValue_(ss, BOUND_KEYS.JOIN_OPEN), true);
    const requireApproval = boundFlag_(getSettingValue_(ss, BOUND_KEYS.REQUIRE_APPROVAL), true);

    let state;
    if (isTeacher) state = 'teacher';
    else if (member && member.status === 'active') state = 'active';
    else if (member && member.status === 'pending') state = 'pending';
    else if (!joinOpen) state = 'closed';
    else state = 'unregistered';

    return jsonOk_({
      state: state,
      className: String(getSettingValue_(ss, BOUND_KEYS.CLASS_NAME) || ''),
      displayName: member ? member.displayName : '',
      role: isTeacher ? 'teacher' : 'student',
      requireApproval: requireApproval
    });
  } catch (e) { return jsonErr_(e); }
}

/** クラス参加申請。表示名は氏名でなく出席番号・ニックネームでもよい */
function bdJoin(displayName, number) {
  try {
    const ss = boundSs_();
    const email = boundEmail_();
    if (!(boundOwner_(ss) || boundAutoClaimOwner_(ss, email))) {
      throw new Error('NOT_READY: 先生の準備がまだ終わっていません。少し待ってからもう一度開いてください');
    }
    if (!boundFlag_(getSettingValue_(ss, BOUND_KEYS.JOIN_OPEN), true)) {
      throw new Error('JOIN_CLOSED: いまは参加の受付が閉じられています。先生に確認してください');
    }
    const name = vStr_(displayName, 30, '表示名').trim();
    if (!name) throw new Error('BAD_INPUT: 表示名を入力してください');
    const num = vStr_(number, 10, '出席番号').trim();

    const existing = getMemberRow_(ss, email);
    const status = existing && existing.status === 'active' ? 'active'
      : (boundFlag_(getSettingValue_(ss, BOUND_KEYS.REQUIRE_APPROVAL), true) ? 'pending' : 'active');

    upsertMember_(ss, {
      email: email, displayName: name, number: num,
      role: (existing && existing.role) || 'student', status: status
    });
    return jsonOk_({ state: status });
  } catch (e) { return jsonErr_(e); }
}

// ────────────────────────────────────────────────────────────────
// データ取得（児童にも先生にも同じ形で返す。email は uid に置き換える）
// ────────────────────────────────────────────────────────────────

function bdBuildInitData_(g) {
  const ss = g.ss;
  const units = getTableData_(ss, TABLES.UNITS);
  const activeUnit = formatUnit_(units.filter(function (u) { return u.is_active === true; })[0] || units[0] || null);
  const data = activeUnit ? coreCollectUnitData_(ss, activeUnit.unit_id) : { pins: [], chats: [], reactions: [] };
  return {
    user: {
      email: uidOf_(g.email),
      name: (g.member && g.member.displayName) || (g.isTeacher ? '先生' : ''),
      group_id: (g.member && g.member.groupId) || '',
      role: g.isTeacher ? 'teacher' : 'student'
    },
    users: sanitizeMembers_(getMembers_(ss)),
    activeUnit: activeUnit,
    units: units.map(function (u) { return formatUnit_(u); }),
    pins: sanitizeRecords_(data.pins),
    chats: sanitizeRecords_(data.chats),
    reactions: sanitizeRecords_(data.reactions),
    // API キーそのものは返さない（有無だけ）
    hasApiKey: g.isTeacher ? !!getSettingValue_(ss, 'geminiApiKey') : false
  };
}

function bdGetInitData() {
  try {
    return jsonOk_(bdBuildInitData_(guardBoundActive_()));
  } catch (e) { return jsonErr_(e); }
}

function bdSyncData(unitId) {
  try {
    const g = guardBoundActive_();
    const units = getTableData_(g.ss, TABLES.UNITS);
    // 現在アクティブな単元を優先して返す（先生が単元を切り替えたら児童側も追従する）
    const activeUnit = formatUnit_(
      units.filter(function (u) { return u.is_active === true; })[0] ||
      units.filter(function (u) { return u.unit_id === unitId; })[0] || null);
    const data = coreCollectUnitData_(g.ss, unitId);
    return jsonOk_({
      pins: sanitizeRecords_(data.pins),
      chats: sanitizeRecords_(data.chats),
      reactions: sanitizeRecords_(data.reactions),
      activeUnit: activeUnit,
      users: sanitizeMembers_(getMembers_(g.ss)),
      hasApiKey: g.isTeacher ? !!getSettingValue_(g.ss, 'geminiApiKey') : false
    });
  } catch (e) { return jsonErr_(e); }
}

// ────────────────────────────────────────────────────────────────
// 書き込み（入口は 1 つ。誰が何をしてよいかはここで決める）
// ────────────────────────────────────────────────────────────────

/** 先生でなければできない操作の一覧（フロントの出し分けは防御とみなさない） */
const BOUND_TEACHER_ACTIONS = [
  'save_unit', 'add_map', 'toggle_chat', 'toggle_stamp', 'update_custom_stamps',
  'save_users', 'save_api_key', 'approve_members', 'remove_member',
  'set_join_open', 'set_require_approval', 'set_class_name', 'repair_schema'
];

function bdExecuteAction(payloadJson) {
  try {
    const g = guardBoundActive_();
    const p = JSON.parse(payloadJson);
    const ss = g.ss;
    const email = g.email;

    if (BOUND_TEACHER_ACTIONS.indexOf(p.action) >= 0 && !g.isTeacher) {
      throw new Error('FORBIDDEN: この操作は先生だけができます');
    }

    // ── 記録系（児童も先生も。ただし他人の記録は先生だけが消せる）──
    if (p.action === 'save_pin') return jsonOk_(coreSavePin_(ss, email, p));
    if (p.action === 'update_pin') return jsonOk_(coreUpdateOwnPin_(ss, email, p.pin_id, p));
    if (p.action === 'save_chat') {
      // チャット OFF の単元では一般メッセージを拒否（ピンへのコメントは可）
      if (p.target_type === 'general' || p.target_type === 'chat') {
        const unit = getTableData_(ss, TABLES.UNITS).filter(function (u) { return u.unit_id === p.unit_id; })[0];
        if (unit && unit.chat_enabled === false && !g.isTeacher) {
          throw new Error('FORBIDDEN: いまはチャットがオフになっています');
        }
      }
      return jsonOk_(coreSaveChat_(ss, email, p));
    }
    if (p.action === 'toggle_reaction') return jsonOk_(coreToggleReaction_(ss, email, p));
    if (p.action === 'delete_pin' || p.action === 'delete_chat') {
      return jsonOk_(coreDeleteRecord_(ss, email, p.pin_id || p.chat_id, !g.isTeacher));
    }

    // ── 単元まわり（共有コア）──
    if (['save_unit', 'add_map', 'toggle_chat', 'toggle_stamp', 'update_custom_stamps'].indexOf(p.action) >= 0) {
      return jsonOk_(coreUnitAction_(ss, p));
    }

    // ── 名簿・設定 ──
    if (p.action === 'save_users') {
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
      return jsonOk_({});
    }
    if (p.action === 'approve_members') {
      const emails = (p.emails || []).slice(0, 200).map(function (e) { return String(e).toLowerCase(); });
      if (!emails.length) throw new Error('BAD_INPUT: 承認する人を選んでください');
      setMemberStatus_(ss, emails, 'active');
      return jsonOk_({});
    }
    if (p.action === 'remove_member') {
      const target = String(p.email || '').toLowerCase();
      if (!target) throw new Error('BAD_INPUT: 対象がありません');
      if (target === g.owner) throw new Error('FORBIDDEN: 先生ご自身は外せません');
      // 行は消さず status を removed にする（記録との対応が切れないように）
      setMemberStatus_(ss, [target], 'removed');
      return jsonOk_({});
    }
    if (p.action === 'set_join_open') {
      setSettingValue_(ss, BOUND_KEYS.JOIN_OPEN, p.value === true ? 'true' : 'false');
      return jsonOk_({});
    }
    if (p.action === 'set_require_approval') {
      setSettingValue_(ss, BOUND_KEYS.REQUIRE_APPROVAL, p.value === true ? 'true' : 'false');
      return jsonOk_({});
    }
    if (p.action === 'set_class_name') {
      setSettingValue_(ss, BOUND_KEYS.CLASS_NAME, vStr_(p.name, 50, 'クラス名').trim());
      return jsonOk_({});
    }
    if (p.action === 'save_api_key') {
      setSettingValue_(ss, 'geminiApiKey', vStr_(p.api_key, 200, 'APIキー').trim());
      return jsonOk_({});
    }
    if (p.action === 'repair_schema') {
      return jsonOk_(repairSchema_(ss));
    }
    throw new Error('BAD_INPUT: 不明な操作です');
  } catch (e) { return jsonErr_(e); }
}

// ────────────────────────────────────────────────────────────────
// 名簿の管理（先生のみ）
// ────────────────────────────────────────────────────────────────

/** 名簿一覧。email は先生の画面にだけ出す（承認の判断に要るため） */
function bdListMembers() {
  try {
    const g = guardBoundTeacher_();
    const members = getMembers_(g.ss).map(function (m) {
      return {
        email: String(m.email).toLowerCase(),
        displayName: m.displayName || '',
        number: m.number || '',
        groupId: m.groupId || '',
        role: m.role === 'teacher' ? 'teacher' : 'student',
        status: m.status || ''
      };
    });
    return jsonOk_({
      className: String(getSettingValue_(g.ss, BOUND_KEYS.CLASS_NAME) || ''),
      teacherEmail: g.owner,
      joinOpen: boundFlag_(getSettingValue_(g.ss, BOUND_KEYS.JOIN_OPEN), true),
      requireApproval: boundFlag_(getSettingValue_(g.ss, BOUND_KEYS.REQUIRE_APPROVAL), true),
      members: members
    });
  } catch (e) { return jsonErr_(e); }
}

// ────────────────────────────────────────────────────────────────
// シートの点検（先生のみ）。直すのは bdExecuteAction の repair_schema
// ────────────────────────────────────────────────────────────────

function bdCheckSchema() {
  try {
    const g = guardBoundTeacher_();
    const found = checkSchema_(g.ss);
    return jsonOk_({
      findings: found,
      report: formatSchemaReport_(found),
      ok: found.length === 0,
      blocking: found.filter(function (f) { return f.blocking; }).length,
      fixable: found.filter(function (f) { return f.fixable; }).length
    });
  } catch (e) { return jsonErr_(e); }
}

// ────────────────────────────────────────────────────────────────
// AI ポートフォリオ（先生のみ）
// ────────────────────────────────────────────────────────────────

function bdGenerateAIPortfolio(payloadJson) {
  try {
    const g = guardBoundTeacher_();
    const ss = g.ss;
    const p = JSON.parse(payloadJson);
    const apiKey = getSettingValue_(ss, 'geminiApiKey');
    if (!apiKey) {
      throw new Error('NO_API_KEY: AI分析を行うには、管理パネルの「AI設定」タブで Gemini API キーを設定してください。');
    }

    const member = getMembers_(ss).filter(function (m) { return uidOf_(m.email) === p.uid; })[0];
    if (!member) throw new Error('NOT_FOUND: 対象の児童が見つかりません');
    const targetEmail = String(member.email).toLowerCase();
    const sameStudent = function (r) {
      return r.unit_id === p.unit_id && String(r.email).toLowerCase() === targetEmail;
    };

    const pins = getTableData_(ss, TABLES.PINS).filter(sameStudent);
    const chats = getTableData_(ss, TABLES.CHATS).filter(sameStudent);
    const reactions = getTableData_(ss, TABLES.REACTIONS).filter(sameStudent);
    if (pins.length === 0 && chats.length === 0 && reactions.length === 0) {
      return jsonOk_({ portfolio: 'まだ活動の記録（ピンやチャット）がありません。' });
    }

    // AI には実名を渡さない。組み立ては corePortfolioPrompt_（Tenant.gs）に 1 本化してある。
    const built = corePortfolioPrompt_(
      getMembers_(ss).map(function (m) { return m.displayName; }),
      member.displayName, pins, chats, reactions);
    return jsonOk_({ portfolio: coreRunPortfolio_(apiKey, built) });
  } catch (e) { return jsonErr_(e); }
}

// ────────────────────────────────────────────────────────────────
// 画像
// ────────────────────────────────────────────────────────────────

function bdUploadImage(dataUrl) {
  try {
    const g = guardBoundActive_();
    return jsonOk_({ imageRef: storeImage_(g.ss, g.email, dataUrl) });
  } catch (e) { return jsonErr_(e); }
}

function bdGetImage(imageRef) {
  try {
    const g = guardBoundActive_();
    return jsonOk_({ dataUrl: loadImage_(g.ss, imageRef) });
  } catch (e) { return jsonErr_(e); }
}

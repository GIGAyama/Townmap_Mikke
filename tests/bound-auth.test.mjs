/**
 * bound-auth.test.mjs — コンテナバインド版の「誰が何をできるか」の検査
 *
 * この形は先生ごとに 1 デプロイなので、認可を間違えると**その学級の全員が先生**になる。
 * 手元では GAS を動かせないので、Session を偽物に差しかえて、次を見る。
 *
 *   1. `Session.getEffectiveUser()` を本人確認に使っていないこと
 *      （「自分として実行」では誰が開いてもデプロイした先生を返すため、
 *        これを使った瞬間に学級全員が先生になる）
 *   2. 「アクセスできるユーザー: 全員」で本人が分からないとき、誰も通さないこと
 *   3. 先生の操作を児童が呼んでも通らないこと（画面の出し分けは防御ではない）
 *   4. 承認前（pending）の児童がデータを読めないこと
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { loadGas, sheet, spreadsheet, parse, ROOT } from './helpers/gas-sandbox.mjs';

const MEMBERS = ['email', 'displayName', 'role', 'status', 'number', 'groupId', 'joinedAt'];
const PINS = ['pin_id', 'unit_id', 'map_id', 'email', 'x', 'y', 'color', 'title', 'memo', 'image_url', 'created_at'];
const UNITS = ['unit_id', 'name', 'maps_json', 'chat_enabled', 'stamp_enabled', 'custom_stamps', 'is_active', 'created_at'];
const CHATS = ['chat_id', 'unit_id', 'email', 'message', 'target_type', 'target_id', 'created_at'];
const REACTIONS = ['reaction_id', 'unit_id', 'email', 'target_type', 'target_id', 'emoji', 'created_at'];
const IMAGES = ['image_id', 'owner_email', 'created_at', 'chunk_count',
  'c1', 'c2', 'c3', 'c4', 'c5', 'c6', 'c7', 'c8', 'c9', 'c10'];

const TEACHER = 'sensei@school.ed.jp';
const STUDENT = 'ayumi@school.ed.jp';

/** 配ったスプレッドシートのコピー。settings / members は差しかえられる */
function container({ settings = [['ownerEmail', TEACHER]], members = [] } = {}) {
  return spreadsheet([
    sheet('Members', MEMBERS, members),
    sheet('Settings', ['key', 'value'], settings),
    sheet('_Meta', ['key', 'value'], []),
    sheet('Units_単元', UNITS, [['unt_000001', 'まちたんけん', '[]', true, true, '[]', true, '2026-01-01']]),
    sheet('Pins_ピン', PINS, []),
    sheet('Chats_チャット', CHATS, []),
    sheet('Reactions_反応', REACTIONS, []),
    sheet('Images_画像', IMAGES, [])
  ]);
}

const activeStudent = [[STUDENT, 'あゆみ', 'student', 'active', '1', '', '2026-01-01']];

test('「自分として実行」でも、児童は先生にならない（getEffectiveUser を本人確認に使っていない）', () => {
  const ss = container({ members: activeStudent });
  // ここが本番の形: 誰が開いても effectiveUser は先生のメールになる
  const { ctx } = loadGas({ ss, activeUser: STUDENT, effectiveUser: TEACHER });

  const status = parse(ctx.bdGetStatus());
  assert.equal(status.success, true);
  assert.equal(status.state, 'active');
  assert.equal(status.role, 'student', '児童が先生として通っている');

  const res = parse(ctx.bdExecuteAction(JSON.stringify({ action: 'save_api_key', api_key: 'AIza-dummy' })));
  assert.equal(res.success, false);
  assert.equal(res.code, 'FORBIDDEN');
});

test('デプロイした先生は先生として通る', () => {
  const ss = container();
  const { ctx } = loadGas({ ss, activeUser: TEACHER, effectiveUser: TEACHER });
  const status = parse(ctx.bdGetStatus());
  assert.equal(status.state, 'teacher');
  assert.equal(status.role, 'teacher');
});

test('名簿で role=teacher にした同僚も先生として通る', () => {
  const ss = container({ members: [['fuku@school.ed.jp', '副担任', 'teacher', 'active', '', '', '2026-01-01']] });
  const { ctx } = loadGas({ ss, activeUser: 'fuku@school.ed.jp', effectiveUser: TEACHER });
  assert.equal(parse(ctx.bdGetStatus()).role, 'teacher');
});

test('「アクセスできるユーザー: 全員」で本人が分からないときは、誰も通さない', () => {
  const ss = container({ members: activeStudent });
  const { ctx } = loadGas({ ss, activeUser: '', effectiveUser: TEACHER });

  const status = parse(ctx.bdGetStatus());
  assert.equal(status.success, false);
  assert.equal(status.code, 'BOUND_NO_IDENTITY');
  // 直し方が本文に入っていること（先生が読んで動けること）
  assert.match(status.error, /同一ドメインの全員/);

  const init = parse(ctx.bdGetInitData());
  assert.equal(init.success, false);
});

test('先生がまだ登録されていない間は、開いた人を先生にしない', () => {
  // ownerEmail が空。かつ active === effective（＝先生自身しか開いていない状態と区別できない）
  const ss = container({ settings: [] });
  const { ctx } = loadGas({ ss, activeUser: STUDENT, effectiveUser: STUDENT });

  const status = parse(ctx.bdGetStatus());
  assert.equal(status.state, 'setup');
  assert.notEqual(status.role, 'teacher');
  // Settings に ownerEmail が書かれていないこと
  assert.equal(ctx.getSettingValue_(ss, 'ownerEmail'), '');
});

test('active !== effective が観測できたときだけ、デプロイした先生を自動登録する', () => {
  const ss = container({ settings: [] });
  const { ctx } = loadGas({ ss, activeUser: STUDENT, effectiveUser: TEACHER });

  const status = parse(ctx.bdGetStatus());
  assert.notEqual(status.state, 'setup');
  assert.equal(ctx.getSettingValue_(ss, 'ownerEmail'), TEACHER);
  // 自動登録されたのは先生であって、いま開いている児童ではない
  assert.equal(status.role, 'student');
});

test('承認待ちの児童は、クラスのデータを 1 件も読めない', () => {
  const ss = container({ members: [[STUDENT, 'あゆみ', 'student', 'pending', '1', '', '2026-01-01']] });
  const { ctx } = loadGas({ ss, activeUser: STUDENT, effectiveUser: TEACHER });

  assert.equal(parse(ctx.bdGetStatus()).state, 'pending');
  const init = parse(ctx.bdGetInitData());
  assert.equal(init.success, false);
  assert.equal(init.code, 'NOT_MEMBER');
  assert.equal(parse(ctx.bdSyncData('unt_000001')).success, false);
  assert.equal(parse(ctx.bdListMembers()).success, false);
});

test('名簿に無い人は読めない（先生の名簿だけが入口）', () => {
  const ss = container({ members: activeStudent });
  const { ctx } = loadGas({ ss, activeUser: 'yoso@other.example', effectiveUser: TEACHER });
  const init = parse(ctx.bdGetInitData());
  assert.equal(init.success, false);
  assert.equal(init.code, 'NOT_MEMBER');
});

test('児童は自分の記録だけ消せる', () => {
  const ss = container({ members: activeStudent });
  const pins = ss.getSheetByName('Pins_ピン');
  pins.appendRow(['pin_000001', 'unt_000001', 'm1', STUDENT, 1, 1, '#000', 'じぶん', '', '', '2026-01-01']);
  pins.appendRow(['pin_000002', 'unt_000001', 'm1', 'hoka@school.ed.jp', 2, 2, '#000', 'ほかの子', '', '', '2026-01-01']);

  const { ctx } = loadGas({ ss, activeUser: STUDENT, effectiveUser: TEACHER });
  const ng = parse(ctx.bdExecuteAction(JSON.stringify({ action: 'delete_pin', pin_id: 'pin_000002' })));
  assert.equal(ng.success, false);
  assert.equal(ng.code, 'FORBIDDEN');

  const ok = parse(ctx.bdExecuteAction(JSON.stringify({ action: 'delete_pin', pin_id: 'pin_000001' })));
  assert.equal(ok.success, true);
  assert.equal(ctx.getTableData_(ss, ctx.TABLES.PINS).length, 1);
});

test('先生は誰の記録でも消せる', () => {
  const ss = container({ members: activeStudent });
  ss.getSheetByName('Pins_ピン')
    .appendRow(['pin_000002', 'unt_000001', 'm1', STUDENT, 2, 2, '#000', 'あゆみのピン', '', '', '2026-01-01']);

  const { ctx } = loadGas({ ss, activeUser: TEACHER, effectiveUser: TEACHER });
  const ok = parse(ctx.bdExecuteAction(JSON.stringify({ action: 'delete_pin', pin_id: 'pin_000002' })));
  assert.equal(ok.success, true);
});

test('児童むけの返事に、他の子のメールアドレスが入っていない', () => {
  const ss = container({
    members: activeStudent.concat([['hoka@school.ed.jp', 'ほか', 'student', 'active', '2', '', '2026-01-01']])
  });
  ss.getSheetByName('Pins_ピン')
    .appendRow(['pin_000003', 'unt_000001', 'm1', 'hoka@school.ed.jp', 3, 3, '#000', 'ほかの子のピン', '', '', '2026-01-01']);

  const { ctx } = loadGas({ ss, activeUser: STUDENT, effectiveUser: TEACHER });
  const init = parse(ctx.bdGetInitData());
  assert.equal(init.success, true);
  const text = JSON.stringify(init);
  assert.ok(!text.includes('hoka@school.ed.jp'), '他の子のメールが返っている');
  assert.ok(!text.includes(STUDENT), '自分のメールも uid に置き換えるはず');
});

test('先生の名簿一覧はメールを出す（承認の判断に要る）が、児童には出さない', () => {
  const ss = container({ members: activeStudent });
  const asTeacher = loadGas({ ss, activeUser: TEACHER, effectiveUser: TEACHER }).ctx;
  const list = parse(asTeacher.bdListMembers());
  assert.equal(list.success, true);
  assert.ok(JSON.stringify(list).includes(STUDENT));

  const asStudent = loadGas({ ss, activeUser: STUDENT, effectiveUser: TEACHER }).ctx;
  assert.equal(parse(asStudent.bdListMembers()).success, false);
});

test('参加の受付が閉じていれば、申請そのものを断る', () => {
  const ss = container({ settings: [['ownerEmail', TEACHER], ['joinOpen', 'false']] });
  const { ctx } = loadGas({ ss, activeUser: STUDENT, effectiveUser: TEACHER });
  assert.equal(parse(ctx.bdGetStatus()).state, 'closed');
  const res = parse(ctx.bdJoin('あゆみ', '1'));
  assert.equal(res.success, false);
  assert.equal(res.code, 'JOIN_CLOSED');
});

test('承認が要らない設定なら、申請がそのまま active になる', () => {
  const ss = container({ settings: [['ownerEmail', TEACHER], ['requireApproval', 'false']] });
  const { ctx } = loadGas({ ss, activeUser: STUDENT, effectiveUser: TEACHER });
  assert.equal(parse(ctx.bdJoin('あゆみ', '1')).state, 'active');
  assert.equal(parse(ctx.bdGetStatus()).state, 'active');
});

test('メニューの関数は、画面の無い文脈（ウェブアプリ）では 1 行も読まずに止まる', () => {
  const ss = container({ members: activeStudent });
  // hasUi を渡さない = getUi() が例外になる（本番のウェブアプリ文脈と同じ）
  const { ctx } = loadGas({ ss, activeUser: STUDENT, effectiveUser: TEACHER });
  assert.throws(() => ctx.showSheetCheck(), /getUi/);
  assert.throws(() => ctx.runSheetRepair(), /getUi/);
  assert.throws(() => ctx.runInitialSetup(), /getUi/);
  assert.throws(() => ctx.showClassInfo(), /getUi/);
});

test('メニューの「はじめの設定」は、押した本人（＝ファイルの編集者）を先生にする', () => {
  const ss = container({ settings: [] });
  const { ctx, ui } = loadGas({ ss, activeUser: TEACHER, effectiveUser: TEACHER, hasUi: true });
  ctx.runInitialSetup();
  assert.equal(ctx.getSettingValue_(ss, 'ownerEmail'), TEACHER);
  assert.ok(ui.alerts.length === 1);
  assert.match(ui.alerts[0].body, /先生として登録しました/);
});

test('AI に送る文には実名も他の子の名前も入らない', () => {
  const ss = container({
    members: activeStudent.concat([['hoka@school.ed.jp', 'ほかの子', 'student', 'active', '2', '', '2026-01-01']])
  });
  ss.getSheetByName('Pins_ピン')
    .appendRow(['pin_000001', 'unt_000001', 'm1', STUDENT, 1, 1, '#000',
      'あゆみの見つけたもの', 'ほかの子といっしょに見た。れんらくは 090-1234-5678', '', '2026-01-01']);

  let sent = null;
  const { ctx } = loadGas({
    ss, activeUser: TEACHER, effectiveUser: TEACHER,
    gemini: (opts) => { sent = opts.prompt; return '対象児童 はよく見ています'; }
  });
  ctx.bdExecuteAction(JSON.stringify({ action: 'save_api_key', api_key: 'AIza-dummy' }));

  const uid = ctx.uidOf_(STUDENT);
  const res = parse(ctx.bdGenerateAIPortfolio(JSON.stringify({ unit_id: 'unt_000001', uid: uid })));
  assert.equal(res.success, true);
  assert.ok(sent, 'Gemini が呼ばれていない');
  assert.ok(!sent.includes('あゆみ'), '実名が送られている: ' + sent);
  assert.ok(!sent.includes('ほかの子'), '友達の名前が送られている: ' + sent);
  assert.ok(!sent.includes('090-1234-5678'), '電話番号が送られている: ' + sent);
  // 返事は先生の画面で読めるよう実名に戻る
  assert.match(res.portfolio, /あゆみ/);
});

test('AI 生成は先生だけ', () => {
  const ss = container({ members: activeStudent, settings: [['ownerEmail', TEACHER], ['geminiApiKey', 'AIza-dummy']] });
  const { ctx } = loadGas({ ss, activeUser: STUDENT, effectiveUser: TEACHER, gemini: () => 'ng' });
  const res = parse(ctx.bdGenerateAIPortfolio(JSON.stringify({ unit_id: 'unt_000001', uid: 'u0' })));
  assert.equal(res.success, false);
  assert.equal(res.code, 'FORBIDDEN');
});

test('bound の学級では、旧版の入口（lg*）から先生になれない', () => {
  // google.script.run は末尾 `_` の無い関数を誰でも呼べる。児童がコンソールから
  // lgGetInitData() を呼ぶと、以前は
  //   Users_名簿 が作られる → 「名簿が空なら最初の人を先生」で児童が先生になる
  //   → lgExecuteAction で単元を作れる
  // という経路で、Bound.gs の ownerEmail の判定を丸ごと迂回できた。
  const ss = container({ members: activeStudent });
  const { ctx } = loadGas({ ss, activeUser: STUDENT, effectiveUser: TEACHER });

  const init = parse(ctx.lgGetInitData());
  assert.equal(init.success, false, '旧版の入口が通ってしまう');
  assert.equal(init.code, 'FORBIDDEN');

  const act = parse(ctx.lgExecuteAction(JSON.stringify({
    action: 'save_unit', unit_id: 'unt_000099', name: 'のっとり', map_name: 'x', map_url: ''
  })));
  assert.equal(act.success, false);
  assert.equal(act.code, 'FORBIDDEN');

  // 単元は増えていない（学級のデータに 1 行も触れていない）
  assert.equal(ctx.getTableData_(ss, ctx.TABLES.UNITS).length, 1);
  // Users_名簿 が作られてもいない
  assert.equal(ss.getSheetByName('Users_名簿'), null);
});

test('旧版で動いている学級（Users_名簿 に行がある）は、これまでどおり通る', () => {
  const ss = spreadsheet([
    sheet('Users_名簿', ['email', 'name', 'group_id', 'role', 'created_at'],
      [[TEACHER, '先生', 'teacher', 'teacher', '2026-01-01'],
       [STUDENT, 'あゆみ', '1班', 'student', '2026-01-01']]),
    sheet('Units_単元', UNITS, [['unt_000001', 'まちたんけん', '[]', true, true, '[]', true, '2026-01-01']]),
    sheet('Pins_ピン', PINS, []), sheet('Chats_チャット', CHATS, []),
    sheet('Reactions_反応', REACTIONS, []), sheet('Images_画像', IMAGES, []),
    sheet('Settings', ['key', 'value'], [])
  ]);
  const { ctx } = loadGas({ ss, activeUser: STUDENT, effectiveUser: STUDENT });
  assert.equal(ctx.boundModeFor_(ss), 'legacy');
  const init = parse(ctx.lgGetInitData());
  assert.equal(init.success, true, JSON.stringify(init));
  assert.equal(init.user.role, 'student');
});

test('空のファイルでは、旧版の入口から「最初の人が先生」にならない', () => {
  // Users_名簿 が空なら boundModeFor_ は 'bound' を返すので、lg* はそもそも通らない。
  const ss = container({ settings: [] });
  const { ctx } = loadGas({ ss, activeUser: STUDENT, effectiveUser: STUDENT });
  const init = parse(ctx.lgGetInitData());
  assert.equal(init.success, false);
  assert.equal(init.code, 'FORBIDDEN');
});

test('authorizeApp は、エディタ以外から呼ぶと先生のメールを返さない', () => {
  const ss = container({ members: activeStudent });
  // 児童のブラウザから: active(児童) !== effective(先生)
  const asStudent = loadGas({ ss, activeUser: STUDENT, effectiveUser: TEACHER }).ctx;
  assert.throws(() => asStudent.authorizeApp(), /FORBIDDEN/);

  // GAS エディタから: active === effective
  const asEditor = loadGas({ ss, activeUser: TEACHER, effectiveUser: TEACHER }).ctx;
  assert.match(asEditor.authorizeApp(), /実行者: /);
});

test('公開エンドポイントに、認可の無いものが増えていない', () => {
  // 末尾 `_` の無いトップレベル関数は google.script.run から誰でも呼べる。
  // 増やしたときに「認可を書いたか」を必ず考えるよう、一覧を固定しておく。
  const known = new Set([
    // 入口・診断
    'doGet', 'authorizeApp', 'onOpen',
    // スプレッドシートのメニュー（getUi() を先に取るので画面が無い文脈では止まる）
    'showSheetCheck', 'runSheetRepair', 'showClassInfo', 'runInitialSetup',
    // コンテナバインド版
    'bdGetStatus', 'bdJoin', 'bdGetInitData', 'bdSyncData', 'bdExecuteAction',
    'bdListMembers', 'bdCheckSchema', 'bdGenerateAIPortfolio', 'bdUploadImage', 'bdGetImage',
    // 共通デプロイ版（児童）
    'stGetStatus', 'stJoin', 'stGetInitData', 'stSyncData', 'stSubmit', 'stListMine',
    'stUpdateMine', 'stDeleteMine', 'stUploadImage', 'stGetImage',
    // 共通デプロイ版（教員ポータル）
    'tpGetMyPortal', 'tpCreateClass', 'tpRegisterExisting', 'tpListMembers', 'tpApprove',
    'tpRemove', 'tpSetJoinOpen', 'tpSetRequireApproval', 'tpRotateCode', 'tpRevokeClass',
    'tpGetInitData', 'tpSyncData', 'tpExecuteAction', 'tpGenerateAIPortfolio',
    'tpUploadImage', 'tpGetImage',
    // 旧バインド型
    'lgGetInitData', 'lgSyncData', 'lgExecuteAction', 'lgGenerateAIPortfolio',
    'lgUploadImage', 'lgGetImage'
  ]);

  const found = new Set();
  for (const f of fs.readdirSync(ROOT).filter((n) => n.endsWith('.gs'))) {
    const src = fs.readFileSync(ROOT + '/' + f, 'utf8');
    const re = /^function\s+([A-Za-z][A-Za-z0-9]*)\s*\(/gm;
    let m;
    while ((m = re.exec(src)) !== null) if (!m[1].endsWith('_')) found.add(m[1]);
  }

  const added = [...found].filter((n) => !known.has(n));
  assert.deepEqual(added, [],
    '公開エンドポイントが増えています。認可を書いたうえで、このテストの一覧にも足してください: ' + added.join(', '));
  const removed = [...known].filter((n) => !found.has(n));
  assert.deepEqual(removed, [], '一覧にあるのに実体が無い関数: ' + removed.join(', '));
});

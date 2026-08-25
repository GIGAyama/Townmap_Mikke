/**
 * みっけ！ — コンテナバインド版（先生ごとに 1 デプロイ）
 *
 * ── 配り方 ─────────────────────────────────────────────────────
 *   1. 先生が「スプレッドシートのコピー」を作る（このスクリプトも一緒についてくる）
 *   2. スプレッドシートのメニュー「みっけ！」＞「はじめの設定」を 1 回押す
 *   3. 「拡張機能 ＞ Apps Script ＞ デプロイ」でウェブアプリを 1 本公開し、URL を配る
 *
 * 記録は、束ねられたそのスプレッドシートの中だけにあります。作者にも、
 * ほかの先生にも見えません。共通の入口も、配布元の設定も、クラスコードもありません。
 *
 * ── デプロイの設定（appsscript.json）────────────────────────────
 *   次のユーザーとして実行:  自分（USER_DEPLOYING）
 *   アクセスできるユーザー:  同一ドメインの全員（DOMAIN）
 *
 *   「自分として実行」にすると読み書きが先生の権限で走るので、**児童は
 *   スプレッドシートへのアクセス権を 1 つも持たなくて済みます**（＝児童が
 *   シートを直接開けない、onOpen を動かせない）。
 *   本人確認は `Session.getActiveUser()` だけで行い、これは「同一ドメインの全員」の
 *   ときしか取れません。「全員」に開くと空になるので、Bound.gs が誰も通しません。
 *
 * ── OAuth スコープが 4 つだけである理由（appsscript.json）──────────
 *   - spreadsheets:            束ねられたシートの読み書き
 *   - script.external_request: Gemini API 呼び出し
 *   - userinfo.email:          開いている本人の特定（認可の土台）
 *   - script.container.ui:     スプレッドシートのメニュー（点検・修整・はじめの設定）
 *   DriveApp を一切使わないことでフル Drive スコープを回避し、先生の同意画面に出る
 *   許可を最小にしています。画像は Images シートに圧縮 Data URL をチャンク保存する
 *   方式（Db.gs）にして Drive 依存を無くしました。
 */

const CONFIG = {
  APP_NAME: 'みっけ！',
  SCHEMA_VERSION: 2,
  LOCK_TIMEOUT_MS: 10000,
  // 画像は Data URL を 40,000 文字ずつセルに分割保存（1セル上限50,000文字）
  IMAGE_CHUNK_CHARS: 40000,
  MAX_IMAGE_DATAURL_CHARS: 400000
};

/**
 * 承認の確認（GAS エディタから手動で実行する）。
 *
 * 「自分として実行」では、実行時の権限は「先生がこのスクリプトに与えた承認」で
 * 決まります。承認は実行時ではなく事前に一度だけ行うもので、
 * appsscript.json の oauthScopes を変えた場合や初回承認を飛ばした場合、
 * 児童側で「〜を呼び出す権限がありません」等のエラーになります。
 *
 * 対処: GAS エディタでこの関数を選んで「実行」→ 表示される承認画面で許可する。
 * 再デプロイは不要（既存デプロイに即反映）。
 */
function authorizeApp() {
  // google.script.run は末尾 `_` の無い関数を誰でも呼べる。この関数は
  // GAS エディタから人が動かすためのものなので、**開いている本人と実行者が
  // 同じとき（＝エディタ文脈）だけ**通す。これが無いと、児童がブラウザから
  // 呼ぶだけで先生（デプロイした人）のメールアドレスが返る。
  const active = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  const effective = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (!active || active !== effective) {
    throw new Error('FORBIDDEN: この関数は Apps Script エディタから実行してください');
  }

  const results = [];
  results.push('実行者: ' + Session.getEffectiveUser().getEmail());  // userinfo.email
  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=check',
    { muteHttpExceptions: true });                                    // script.external_request
  results.push('UrlFetch: OK (HTTP ' + res.getResponseCode() + ' は正常です)');
  results.push('spreadsheets スコープ: ' + (ScriptApp.getOAuthToken() ? '承認済み' : '不明'));
  const summary = results.join(' / ');
  Logger.log(summary);
  return summary;
}

function sha256Hex_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(function (b) { return ((b + 256) % 256).toString(16).padStart(2, '0'); })
    .join('');
}

function jsonOk_(obj) {
  const out = obj || {};
  out.success = true;
  out.status = 'success';
  return JSON.stringify(out);
}

function jsonErr_(e) {
  const msg = (e && e.message) ? e.message : String(e);
  let code = '';
  let error = msg;
  const m = msg.match(/^([A-Z_]+):\s*(.*)$/);
  if (m) { code = m[1]; error = m[2] || msg; }
  return JSON.stringify({ success: false, status: 'error', code: code, error: error, message: error });
}

/**
 * 入口。束ねられたスプレッドシートがそのまま学級なので、URL パラメータは要らない。
 * 誰が先生で誰が児童かはサーバー側（Bound.gs）が決める。画面の出し分けは案内であって、
 * 防御ではない。
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.diag === '1') return doGetDiag_();
  const t = HtmlService.createTemplateFromFile('App');
  return t.evaluate()
    .setTitle(CONFIG.APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * 接続診断エンドポイント（?diag=1）。先生が「設定が合っているか」を 1 回で見るためのもの。
 * 秘密情報（メールアドレス・ID・トークン・児童の記録）は一切返さない。
 *
 * 判定の仕組み:
 *   - identity: 開いている本人が特定できているか。false なら「アクセスできるユーザー」が
 *     「同一ドメインの全員」になっていない（Bound.gs は誰も通さない）。
 *   - executeAs: 実行者と閲覧者が違えば「自分として実行」が確定する。同じ場合は
 *     先生自身が開いているだけかもしれないので 'unknown' とし、**緑と言い切らない**。
 *   - setup: 先生が「はじめの設定」を済ませたか。
 */
function doGetDiag_() {
  let effective = '';
  let active = '';
  try { effective = String(Session.getEffectiveUser().getEmail() || '').toLowerCase(); } catch (err) {}
  try { active = String(Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (err) {}

  let bound = false;
  let setup = false;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    bound = !!ss;
    if (ss) setup = !!getSettingValue_(ss, BOUND_KEYS.OWNER);
  } catch (err) { /* 束ねられていない・読めない */ }

  const out = {
    ok: true,
    app: CONFIG.APP_NAME,
    schemaVersion: CONFIG.SCHEMA_VERSION,
    boundToSpreadsheet: bound,
    identity: !!active,
    executeAs: (active && effective && active !== effective) ? 'USER_DEPLOYING' : 'unknown',
    setupDone: setup
  };
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

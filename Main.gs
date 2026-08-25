/**
 * みっけ！ クラス単位マルチテナント版
 *
 * ── アーキテクチャ概要 ─────────────────────────────────────────────
 * 同一プロジェクトから 2 本の Web アプリデプロイを発行して運用する。
 *
 *   デプロイ T（教員ポータル）: 実行するユーザー = ウェブアプリケーションにアクセスしているユーザー
 *     → クラス作成が教員本人の権限で走り、スプレッドシートは最初から教員所有になる。
 *       同じ実行内で addEditor(アプリアカウント) を呼び、共有まで自動化する。
 *   デプロイ S（児童用アプリ）  : 実行するユーザー = 自分（アプリアカウント）
 *     → すべての読み書きがアプリアカウント権限で走る。児童はシートへの権限を一切持たない。
 *       Session.getActiveUser() が使えないため、本人確認は ID トークン検証（Auth.gs）で行う。
 *
 * 入口は GitHub Pages のシェル（共通 URL 1 つ）。シェルが Google サインイン（GIS）を担当し、
 * ID トークンを iframe 内の本アプリへ postMessage で渡す。
 *
 * ── OAuth スコープが 3 つだけである理由（appsscript.json）──────────
 *   - spreadsheets:          シート読み書き・SpreadsheetApp.create・addEditor はこのスコープで動く
 *   - script.external_request: ID トークン検証(tokeninfo) と Gemini API 呼び出し
 *   - userinfo.email:        デプロイ T で教員本人のメールを特定する
 *   DriveApp を一切使わない（makeCopy や画像検索をしない）ことでフル Drive スコープを回避し、
 *   教員の初回同意画面に出る許可を最小にしている。画像はクラス DB 内の Images シートに
 *   圧縮 Data URL をチャンク保存する方式（Db.gs）にして Drive 依存を無くした。
 */

const CONFIG = {
  APP_NAME: 'みっけ！',
  SCHEMA_VERSION: 2,
  LOCK_TIMEOUT_MS: 10000,
  REGISTRY_CACHE_SEC: 600,   // ScriptProperties の日次読み書き上限対策（§レジストリはCache前置）
  TOKEN_CACHE_SEC: 300,
  CODE_ALPHABET: 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', // 紛らわしい文字(I,1,O,0)を除外
  CODE_LENGTH: 8,
  // 画像は Data URL を 40,000 文字ずつセルに分割保存（1セル上限50,000文字）
  IMAGE_CHUNK_CHARS: 40000,
  MAX_IMAGE_DATAURL_CHARS: 400000
};

/**
 * 配布元設定（ScriptProperties）。レジストリ(cls_/own_)とこの sp_* 以外を
 * ScriptProperties に置いてはならない（1値9KB/全体500KB制限のため）。
 *   sp_appAccountEmail : アプリ運営用 Google アカウントのメールアドレス（必須）
 *   sp_googleClientId  : GIS 用 OAuth クライアント ID（必須・aud 検証に使用）
 *   sp_shellUrl        : GitHub Pages シェルの URL（必須・児童用URL/QRの生成に使用）
 *   sp_dbTemplateId    : （任意）DB テンプレートのスプレッドシート ID
 */
const PROP_KEYS = {
  APP_ACCOUNT: 'sp_appAccountEmail',
  CLIENT_ID: 'sp_googleClientId',
  SHELL_URL: 'sp_shellUrl',
  TEMPLATE: 'sp_dbTemplateId'
};

/**
 * 承認トリガー（GAS エディタから手動で実行する）。
 *
 * デプロイ S は「自分（アプリアカウント）として実行」のため、実行時の権限は
 * 「アプリアカウントがこのスクリプトに与えた承認」で決まる。承認は実行時ではなく
 * 事前に一度だけ行うもので、appsscript.json の oauthScopes を変更した場合や
 * 初回承認を飛ばした場合、児童側で
 * 「UrlFetchApp.fetch を呼び出す権限がありません」等のエラーになる。
 *
 * 対処: アプリアカウントで GAS エディタを開き、この関数を選択して「実行」→
 * 表示される承認画面ですべて許可する。再デプロイは不要（既存デプロイに即反映）。
 * 戻り値で 3 スコープが実際に機能しているかを確認できる。
 */
function authorizeApp() {
  const results = [];
  results.push('実行者: ' + Session.getEffectiveUser().getEmail());  // userinfo.email
  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=check',
    { muteHttpExceptions: true });                                    // script.external_request
  results.push('UrlFetch(トークン検証): OK (HTTP ' + res.getResponseCode() + ' は正常です)');
  results.push('spreadsheets スコープ: ' + (ScriptApp.getOAuthToken() ? '承認済み' : '不明'));
  const summary = results.join(' / ');
  Logger.log(summary);
  return summary;
}

function getSetting_(key, required) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  if (!v && required) {
    throw new Error('アプリの初期設定（' + key + '）がまだ行われていません。管理者に連絡してください。');
  }
  return v || '';
}

function getShellUrl_() {
  let url = getSetting_(PROP_KEYS.SHELL_URL, false);
  if (url && url.slice(-1) !== '/') url += '/';
  return url;
}

function getShellOrigin_() {
  const url = getShellUrl_();
  const m = url.match(/^(https?:\/\/[^\/]+)/);
  return m ? m[1] : '';
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
 * ルーティング。
 *
 * ■ コンテナバインド（いまの配り方）
 *   スプレッドシートのコピーを先生に配り、そのファイルの中の Apps Script から
 *   先生ご自身がウェブアプリを 1 本公開する。URL パラメータは何も要らない。
 *   束ねられたスプレッドシートがそのままクラス DB になる（Bound.gs）。
 *
 * ■ 共通デプロイ（前の配り方。すでに公開してある学級のために残す）
 *   1 つのプロジェクトから T（教員ポータル）と S（児童用）の 2 本を公開し、
 *   GitHub Pages のシェルがクラスコードを付けて開く。
 *     - 教員ポータル(T) はシェルが必ず ?portal=1 を付けて開く
 *     - 児童用(S)     はシェルが必ず ?mode=student&c=<code> を付けて開く
 *
 * ■ 旧バインド型（さらに前。Users_名簿 シートで動いていた学級）
 *   Members シートが無く Users_名簿 にデータがあるファイルは legacy で開く。
 *   ここを消すと、貼り付けで入れた古い学級の記録が見えなくなる。
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.diag === '1') return doGetDiag_();
  let bound = null;
  try { bound = SpreadsheetApp.getActiveSpreadsheet(); } catch (err) { bound = null; }
  const mode = p.portal === '1' ? 'teacher'
             : p.mode === 'student' ? 'student'
             : bound ? boundModeFor_(bound)
             : 'landing';
  const t = HtmlService.createTemplateFromFile('App');
  t.bootMode = mode;
  t.bootClassCode = String(p.c || '').replace(/[^A-Z2-9]/gi, '').toUpperCase().slice(0, 16);
  t.bootShellUrl = getShellUrl_();
  t.bootShellOrigin = getShellOrigin_();
  return t.evaluate()
    .setTitle(CONFIG.APP_NAME)
    // GitHub Pages シェルが iframe 埋め込みするため必須。
    // GAS は frame-ancestors の特定オリジン限定ができないため ALLOWALL 一択。
    // その代償として、ID トークン検証（Auth.gs）とシェル側 origin 検証を必須の防御線とする。
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * 束ねられたスプレッドシートを、新しい形（bound）で開くか旧バインド型（legacy）で開くか。
 *
 * 判定は「そのファイルに何が入っているか」だけで行う。URL のパラメータや
 * ScriptProperties は使わない（先生がどちらの手順で入れたかを覚えていなくても、
 * 開いたときに正しいほうが出るようにするため）。
 *
 *   Members シートに 1 行でもある            → bound（新しい形）
 *   Members が空で Users_名簿 にデータがある  → legacy（貼り付けで入れた古い学級）
 *   どちらも空（＝配ったテンプレートのコピー直後） → bound
 */
function boundModeFor_(ss) {
  try {
    const members = ss.getSheetByName(TABLES.MEMBERS.name);
    if (members && members.getLastRow() >= 2) return 'bound';
    const users = ss.getSheetByName(TABLES.USERS.name);
    if (users && users.getLastRow() >= 2) return 'legacy';
  } catch (err) { /* 読めないときは新しい形で開き、Bound.gs 側が案内を出す */ }
  return 'bound';
}

/**
 * 接続診断エンドポイント（?diag=1）。docs/diag.html とシェルのフォールバック画面が
 * fetch して「このデプロイに匿名で届くか・S/T どちらの設定か・初期設定が済んでいるか」を
 * 判定する。秘密情報（メールアドレス・ID・トークン）は一切返さない。
 *
 * 判定の仕組み:
 *   - この JSON が Cookie なしの fetch で読めた時点で「アクセスできるユーザー: 全員
 *     （＝匿名アクセス可）」が確定する。ログイン必須設定だと Google が
 *     accounts.google.com へリダイレクトするため fetch 自体が失敗する。
 *   - deployKind: 実効ユーザーがアプリアカウントなら 'S'（自分として実行）、
 *     それ以外のログインユーザーなら 'T'。匿名到達時に実効ユーザーが空なら、
 *     「全員に公開されているのに実行者がアプリアカウントでない」= S の設定ミス。
 *     （注: アプリアカウント本人のブラウザから T を開いた場合も 'S' と出る）
 */
function doGetDiag_() {
  let effective = '';
  let active = '';
  try { effective = String(Session.getEffectiveUser().getEmail() || '').toLowerCase(); } catch (e) {}
  try { active = String(Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) {}
  const appAccount = String(getSetting_(PROP_KEYS.APP_ACCOUNT, false) || '').toLowerCase();
  const out = {
    ok: true,
    app: CONFIG.APP_NAME,
    schemaVersion: CONFIG.SCHEMA_VERSION,
    deployKind: !effective ? 'unknown'
              : (appAccount && effective === appAccount) ? 'S' : 'T',
    anonymousAccess: !active,
    config: {
      appAccount: !!appAccount,
      clientId: !!getSetting_(PROP_KEYS.CLIENT_ID, false),
      shellUrl: getShellUrl_()  // 児童用 URL に含まれる公開情報のため露出可
    }
  };
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

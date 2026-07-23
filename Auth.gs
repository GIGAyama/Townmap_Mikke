/**
 * Auth.gs — 児童の本人確認: ID トークン検証（デプロイ S の生命線）
 *
 * デプロイ S は「自分（アプリアカウント）として実行」のため Session.getActiveUser() が
 * 使えない。そこで、すべての児童向け API は第 1 引数に GIS（Google Identity Services）の
 * ID トークン（JWT）を受け取り、ここで検証してから処理する。
 *
 * 信頼境界: 信用できる入力は「T では Session.getActiveUser()、S では検証済み ID トークン」のみ。
 * e.parameter・postMessage・google.script.run の引数はすべて改ざん可能として扱う。
 * 検証済み email 以外を書き込み者として記録しない（フロントから email を渡させない）。
 */

/**
 * ID トークンを検証し { email, sub, name } を返す。失敗時は日本語メッセージで throw。
 *
 * クォータについて: tokeninfo 呼び出しは UrlFetch を 1 回消費する。
 * トークンの sha256 をキーにした CacheService（TTL 300 秒）を前置しているため、
 * 同一トークンでの連続操作はキャッシュヒットになり UrlFetch を消費しない。
 * ID トークンの寿命は約 1 時間 → 児童 1 人あたり最大 12 回/時間 × 40 人 ≒ 480 回/時間。
 * 1 日 6 時間使い続けても 3,000 回弱で、UrlFetch 日次上限
 * （個人 20,000 回 / Workspace 100,000 回）に十分収まる。
 */
function verifyIdToken_(idToken) {
  if (!idToken || typeof idToken !== 'string') {
    throw new Error('TOKEN_EXPIRED: サインイン情報がありません。ページを再読み込みしてください');
  }
  const cache = CacheService.getScriptCache();
  const key = 'tok_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken)).slice(0, 40);
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  const res = UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
    { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    throw new Error('TOKEN_EXPIRED: サインインの有効期限が切れました。再読み込みしてください');
  }
  const info = JSON.parse(res.getContentText());

  // aud 検証: 必須。他アプリ向けに発行されたトークンの流用を防ぐ
  const clientId = getSetting_(PROP_KEYS.CLIENT_ID, true);
  if (info.aud !== clientId) throw new Error('AUTH_INVALID: 認証情報が不正です');
  // iss 検証: 必須。Google 発行であることの確認
  if (info.iss !== 'https://accounts.google.com' && info.iss !== 'accounts.google.com') {
    throw new Error('AUTH_INVALID: 認証情報が不正です');
  }
  if (Number(info.exp) * 1000 < Date.now()) {
    throw new Error('TOKEN_EXPIRED: サインインの有効期限が切れました');
  }
  if (info.email_verified !== 'true' && info.email_verified !== true) {
    throw new Error('AUTH_INVALID: メールアドレスが確認されていないアカウントです');
  }

  const user = { email: String(info.email).toLowerCase(), sub: info.sub, name: info.name || '' };
  cache.put(key, JSON.stringify(user), CONFIG.TOKEN_CACHE_SEC);
  return user;
}

/**
 * デプロイ T 用: アクセス中の教員本人のメールアドレス。
 * （T は「アクセスしているユーザーとして実行」なので Session が使える）
 */
function teacherEmail_() {
  const email = Session.getActiveUser().getEmail();
  if (!email) {
    throw new Error('AUTH_INVALID: Google アカウントでログインした状態で、教員ポータルから開いてください');
  }
  return String(email).toLowerCase();
}

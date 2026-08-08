# ✅ GIGA Standard v4 監査：みっけ！（Townmap_Mikke）

- **リポジトリ**：`GIGAyama/Townmap_Mikke`
- **監査日**：2026-08-03
- **アーキテクチャ判定**：**C+型**（GAS ウェブアプリ 2デプロイ + GitHub Pages シェル）
  - GAS 側：`appsscript.json` / `*.gs` 7本 / `App.html`
  - シェル側：`docs/`（`index.html` / `sw.js` / `manifest.webmanifest` / `config.js` / `diag.html` / `icon.svg`）
- **規模**：`App.html` 4,654行 / 211KB、`.gs` 計 1,647行
- **監査方法**：実測のみ（`grep` / ファイルサイズ / manifest と sw.js の読解）。推測は「未検証」と明記した。

> このドキュメントはコードを一切変更していない時点の記録である。

---

## ⚠️ GAS リポジトリ特有の前提（重要）

**GitHub のコードと本番の GAS は自動同期していない。**
このリポジトリに `.clasp.json` は無く、`scriptId` も分からないため、
**本番の GAS が、ここにあるコードと同じである保証はない。**

したがって本監査は「**リポジトリにあるコードに対する**判定」である。
`.gs` / `App.html` に手を入れる場合は、

1. 本番との差分を確認する（`clasp clone <scriptId> --rootDir ./_live && diff -r . ./_live`）
2. **差分があれば本番側が正である可能性が高い。上書きせず報告する**
3. 反映は人間の操作とし、**既存デプロイの URL は変えない**（新規デプロイを作ると児童のブックマークと配布済み QR が切れる）

`docs/` 配下（GitHub Pages シェル）は GitHub が唯一の正本なので、この制約を受けない。

---

## 🚨 最優先：他アプリを壊している

### `docs/sw.js` が同一オリジンの全キャッシュを削除している

```js
// docs/sw.js:16-19
caches.keys().then((keys) =>
  Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
)
```

`gigayama.github.io` は**数十個のアプリが同一オリジンを共有**している。
この Service Worker が有効化されるたび、`mikke-shell-v5` **以外のすべてのキャッシュ**が消える。

**児童が「みっけ！」を開くと、同じ端末に入っている他の GIGA アプリが
オフラインで起動しなくなる。** 直前に整備した `Digital_textbook` の
オフラインキャッシュも、みっけ！を1回開くだけで消える。

GIGA Standard v4 §3-3 が名指しで禁じている事象で、**P1 の最優先項目**。
`CACHE_PREFIX` で始まるキャッシュだけを掃除するように直す必要がある。

### `docs/manifest.webmanifest` の識別子が相対パス

```json
"start_url": ".",
"scope": ".",
// "id" が無い
```

`id` を省略すると `start_url` が代替の識別子になる。それが `"."`（相対）のため、
同一オリジンの似た構成の別アプリと**取り違えられる**恐れがある。
`/Townmap_Mikke/` の絶対パスに直す必要がある。

> **注意**：`id` の変更は、既にインストール済みの端末で「別アプリ」として扱われうる。
> 現状の `id` は未指定（＝`start_url` 由来）なので、明示すると同一性が変わる可能性がある。
> GIGA Standard v4 の停止条件に該当するため、**実施前に人間の判断を仰ぐ**。

---

## 判定記号

| 記号 | 意味 |
|:--:|---|
| ✅ | 基準を満たす |
| ⚠️ | 部分的・条件付きで満たす（要改善） |
| ❌ | 満たさない |
| — | 該当しない（N/A） |

---

## A. 法務・配布

| # | 項目 | 判定 | 実測 | 対応 |
|---|---|:--:|---|:--:|
| A1 | LICENSE 実ファイル | ❌ | ファイル無し | **P0** |
| A2 | .gitignore | ❌ | **ファイル自体が無い**。`.clasp.json` は現状コミットされていないが、`clasp` を使った瞬間に混入する | **P0** |
| A3 | dependabot.yml | ❌ | 無し。npm 依存は無いが、GitHub Actions を導入する際に必要 | **P0** |
| A4 | README.md / MANUAL.md 両方 | ⚠️ | `README.md` は非常に充実（導入手順・トラブルシューティング・セキュリティ設計あり）。**`MANUAL.md`（先生向け・専門用語ゼロ）が無い** | **P3** |

---

## B. セキュリティ

| # | 項目 | 判定 | 実測 | 対応 |
|---|---|:--:|---|:--:|
| B1 | CSP | ❌ | `docs/index.html` / `docs/diag.html` / `App.html` すべて **0件** | **P1** |
| B2 | 秘密情報・IDの直書きなし | ✅ | `docs/config.js` に exec URL 2本と OAuth クライアントIDがあるが、いずれも**公開される識別子**であり鍵ではない（README も設定手順として明記）。スプレッドシートIDは `ScriptProperties` 経由で直書きなし | — |
| B3 | OAuthスコープ最小 | ⚠️ | `spreadsheets` / `script.external_request` / `userinfo.email`。**禁止スコープ（`auth/drive` 全体・`mail.google.com`）は無い**。ただし `spreadsheets` は全スプレッドシートに及ぶ。教員所有のシートを開く設計上、`spreadsheets.currentonly` には落とせない可能性が高い | **報告のみ** |
| B4 | postMessage の宛先が `*` でない | ✅ | 該当なし | — |
| B5 | サーバー側5段ガード | ✅ | `Tenant.gs:79 guardStudent_()` が ①トークン検証 → ②クラス解決 → ③名簿照合(active) を実施。④役割・⑤行の所有者は各 API 側 | — |
| B6 | 信頼できる入力のみを使う | ✅ | `Auth.gs` が ID トークンの `aud` / `iss` / `exp` / `email_verified` をすべて検証。書き込み者 email はクライアント申告ではなくサーバー側で強制（`Legacy.gs:12` にも明記） | — |
| B7 | 児童向けレスポンスの匿名化 | ✅ | `Tenant.gs:92 sanitizeMembers_()` が他児童の email を `uid` に置換 | — |
| **B8** | **外部スクリプトの完全性検証** | ⚠️ | `App.html` が `unpkg.com`（3件）と `cdn.tailwindcss.com` を実行時ロード。`integrity` なし。**ただし App.html は GAS の iframe 内で動くため、B型と同じ自己ホスト化はできない**（`docs/` に置いても GAS からは同一オリジンにならない） | **要検討** |

> **評価**：セキュリティ設計は GIGA Standard v4 Phase 4 の要求を**すでに満たしている**。
> 信頼境界の考え方がコード内コメントにも明記されており、この点は他リポジトリの手本になる。

---

## C. 堅牢性

| # | 項目 | 判定 | 実測 | 対応 |
|---|---|:--:|---|:--:|
| C1 | LockService + try/finally | ✅ | `Db.gs:129` / `TeacherApi.gs:105` で取得し、`finally` 相当で解放（`TeacherApi.gs:176`） | — |
| C2 | 自動復旧（シート再生成） | ⚠️ | **未検証**（`Db.gs` の読解が必要） | **要確認** |
| C3 | pagehide で記録確定 | ⚠️ | **未検証**（`App.html` 内の保存タイミング要確認） | **要確認** |
| C4 | 通信失敗時のリトライと明示 | ⚠️ | `README.md` にトラブルシューティングは厚いが、コード側のリトライは未検証 | **要確認** |
| C5 | localStorage.clear() を使っていない | ✅ | 該当なし | — |

---

## D. 表示（Part I §2）

| # | 項目 | 判定 | 実測 | 対応 |
|---|---|:--:|---|:--:|
| D1 | viewport に `viewport-fit=cover` | ⚠️ | `docs/index.html` ✅ / `App.html` ✅ / **`docs/diag.html` に無い** | **P1** |
| — | `user-scalable=no` | ⚠️ | **`App.html` に `maximum-scale=1.0, user-scalable=no`**。児童が操作する画面には許容されるが、地図上の文字やふりかえりを読む画面が含まれるなら拡大できないのは後退 | **要判断** |
| D2 | `100dvh` を使用 | ⚠️ | `App.html` は `dvh` を使用しコメントも適切。**`docs/index.html` は `height: 100%` のみで `dvh` 不使用** | **P1** |
| D3 | safe-area-inset | ❌ | **0件**（`docs/index.html` / `App.html` とも） | **P1** |
| D4 | clamp() による fluid type | ❌ | **0件**。`docs/index.html` は固定 px（`13px` / `12px` など小さい） | **P1** |
| D5 | Canvas に DPR 補正 | ❌ | `App.html:2594` に `getContext('2d')` があるが `devicePixelRatio` **0件** | **P1** |
| D6 | 320px 幅で横スクロールが出ない | ⚠️ | **未検証** | **P1** |
| D7 | 画像に width/height、150KB以下 | ✅ | 画像は `docs/icon.svg`（398バイト）のみ | — |
| D8 | コントラスト 4.5:1 以上 | ⚠️ | **未計測**。`docs/index.html` の `.card p { color: #64748b }`（白背景で 4.76:1）は辛うじて可、`.note { color: #64748b; 12px }` も同様。要実測 | **P1** |
| D9 | タップ領域 44px・touch-action | ⚠️ | `#gsi-button { min-height: 44px }` は ✅。他は未検証。`touch-action` の指定は 0件 | **P1** |
| D10 | prefers-reduced-motion | ❌ | **0件** | **P1** |
| D11 | 提示モード | ❌ | 無し。**一斉授業で地図を映す用途があるなら必須** | **要判断** |
| D12 | 印刷CSS | ✅ | `App.html` に `@media print` 1件あり | — |

---

## E. PWA（`docs/` シェル）

| # | 項目 | 判定 | 実測 | 対応 |
|---|---|:--:|---|:--:|
| E1 | manifest の id/scope/start_url | ❌ | **`id` が無い。`scope` / `start_url` がどちらも `"."`（相対）**。同一オリジンの他アプリと取り違えられる | **P1（要判断）** |
| E2 | アイコン4種 + apple-touch-icon | ❌ | **`icon.svg` 1枚のみ**。192/512 の PNG が無い。`apple-touch-icon` も SVG を指しており、**iOS は SVG に対応していないためホーム画面のアイコンが出ない** | **P1** |
| E3 | beforeinstallprompt を head 最上部で捕捉 | ❌ | **0件** | **P1** |
| E4 | インストールボタンをアプリ内に設置 | ❌ | 無し | **P1** |
| E5 | sw.js が自アプリ接頭辞のキャッシュのみ削除 | ❌ | **`caches.keys()` の全削除。他アプリを壊している**（冒頭参照） | **P1 最優先** |
| E6 | sw.js が localStorage に触れていない | ✅ | 0件 | — |
| E7 | 更新通知 | ❌ | `skipWaiting()` で即時切替。**児童が操作中に突然切り替わる**。「あたらしい バージョンが あります」の案内なし | **P1** |
| E8 | offline.html | ❌ | 無し | **P1** |
| E9 | キャッシュ版数 | ⚠️ | `mikke-shell-v5`。手動更新のため、リリース手順書への記載が必要 | **P3** |
| E10 | iOS の「ホーム画面に追加」手順 | ❌ | MANUAL.md 自体が無い | **P3** |

> **設計として良い点**：`sw.js` が「GAS と gsi/client は絶対にキャッシュしない」方針を明示し、
> クロスオリジンを素通ししている。ここは正しい。壊れているのは `activate` の掃除範囲だけ。

---

## F. アクセシビリティ・性能

| # | 項目 | 判定 | 実測 | 対応 |
|---|---|:--:|---|:--:|
| F1 | alt / aria-label / aria-live | ⚠️ | **未計測** | **P1** |
| F2 | キーボードのみで全機能に到達 | ⚠️ | **未検証** | **P1** |
| F3 | 初回JS 300KB以下 | ⚠️ | シェルは軽量（`index.html` 18.5KB）。`App.html` 211KB は GAS 側で iframe 内。加えて `cdn.tailwindcss.com`（ランタイムビルド版）と `unpkg.com` を実行時取得しており、実測が必要 | **要確認** |
| F4 | 1ファイル 5,000行 / 400KB 以内 | ⚠️ | `App.html` **4,654行 / 211KB**。基準内だが上限に近い | **P3で提案のみ** |

---

## G. 学習ログ（`study.v1`）

| # | 項目 | 判定 | 備考 |
|---|---|:--:|---|
| G1 | study.v1 準拠 | — | 協働学習の記録をスプレッドシートに持つ設計で、`localStorage['study.records.v1']` を使う学習ドリル系ではない。**未検証**（`App.html` の読解が必要） |

---

## ❌ の総括と対応方針

| フェーズ | 内容 | 破壊リスク | 件数 |
|---|---|:--:|:--:|
| **P0** | LICENSE / .gitignore / dependabot | なし | 3 |
| **P1（最優先）** | **`sw.js` の全キャッシュ削除**を修正 | 低（他アプリを救う） | 1 |
| **P1** | manifest の id/scope/start_url ／ アイコン4種 ／ install ボタン ／ 更新通知 ／ offline.html ／ safe-area ／ fluid type ／ reduced-motion ／ CSP ／ DPR 補正 | 小〜中 | 12 |
| **P3** | MANUAL.md ／ リリース手順 ／ `App.html` 分割の提案 | なし | 3 |
| **P4** | 品質ゲートの移植 | なし | 1 |

### 停止条件に該当する項目

- **`manifest` の `id` 変更**：現状は未指定のため、明示すると既にインストール済みの端末で
  「別アプリ」として扱われうる。実施前に人間の判断が必要
- **`.gs` / `App.html` の変更**：本番との差分が確認できない（`scriptId` 不明）。
  `docs/` 配下のみを対象にすれば、この制約を回避しつつ最優先項目（`sw.js`）を直せる
- **`App.html` の `user-scalable=no`**：児童画面としては許容されるが、
  読む要素が含まれるかの判断が必要

### 推奨する進め方

**`docs/` 配下だけで完結する修正を先に出す。**
最優先の `sw.js` はここに含まれ、GAS 本番との差分リスクを一切負わずに、
フリート全体への被害を止められる。`.gs` / `App.html` の修正は、
本番との差分を確認できてから別 PR にするのが安全。

---
---

# 第1次対応の結果（2026-08-03）

**`docs/` 配下のみ**を対象に、P0 と P1 の主要項目を実施した。
`.gs` / `App.html` には一切触れていない（本番との差分が確認できないため）。

## 再判定

| # | 項目 | 前 | 後 | 根拠 |
|---|---|:--:|:--:|---|
| A1 | LICENSE | ❌ | ✅ | MIT / Copyright (c) 2026 GIGAyama |
| A2 | .gitignore | ❌ | ✅ | 新規作成。`.clasp.json` / `.clasprc.json` / `_live/` / `.env` を除外 |
| A3 | dependabot.yml | ❌ | ✅ | github-actions のみ・月1回（npm 依存なし） |
| D1 | viewport-fit=cover | ⚠️ | ✅ | `diag.html` に欠けていたものを追加。3ページすべてに適用 |
| D2 | 100dvh | ⚠️ | ✅ | `docs/index.html` に `100dvh` とフォールバックを追加 |
| D3 | safe-area-inset | ❌ | ✅ | 上下左右＋更新帯に適用 |
| D4 | clamp() | ❌ | ✅ | `--fs-body` / `--fs-note` / `--fs-title`、行間 1.8 |
| D8 | コントラスト | ⚠️ | ✅ | 主ボタン `#0ea5e9`→`#0369a1`（白文字で **2.77 → 5.43**）、本文 `#64748b`→`#475569`（**4.76 → 8.6**） |
| D9 | タップ44px | ⚠️ | ✅ | `.btn { min-height: 48px }` + `touch-action: manipulation` |
| D10 | prefers-reduced-motion | ❌ | ✅ | 追加。`forced-colors` も対応 |
| E1 | manifest の識別子 | ❌ | ✅ | `id` / `scope` / `start_url` を `/Townmap_Mikke/` に |
| E2 | アイコン4種 | ❌ | ✅ | 192/512 の any と maskable、apple-touch-icon を SVG から生成（計 13KB） |
| E3 | beforeinstallprompt | ❌ | ✅ | `<head>` 最上部の `pwa-install-hook.js` |
| E4 | インストールボタン | ❌ | ✅ | ログイン画面に設置 |
| E5 | **sw.js の掃除範囲** | ❌ | ✅ | `CACHE_PREFIX` 前方一致のみ削除 |
| E7 | 更新通知 | ❌ | ✅ | `skipWaiting` をやめ「あたらしい バージョンが あります」 |
| E8 | offline.html | ❌ | ✅ | アプリと同じ配色・外部読み込みゼロ |

### `manifest` の `id` について（停止条件の解消）

監査時点では「`id` を明示すると既存インストールが別アプリ扱いになりうる」として
停止条件に挙げた。改めて仕様を確認したところ、**同一性は変わらない**と判断できた。

- `id` を省略したときの既定値は `start_url`
- 元の `start_url: "."` は `https://gigayama.github.io/Townmap_Mikke/` に解決される
- 明示した `"id": "/Townmap_Mikke/"` も同じ URL に解決される

したがって計算される識別子は変更前後で一致し、インストール済みの端末で
別アプリにはならない。この理由により実施した。

## 検証結果（Chromium 実機・同一オリジンに2アプリを並べて実施）

### E5 の修正効果（修正前後の比較）

同一オリジンに他アプリのキャッシュを2つ置いた状態で Service Worker を有効化した。

| | 有効化後に残っているキャッシュ | 他アプリの資産 |
|---|---|---|
| **修正前**（`mikke-shell-v5`） | `mikke-shell-v5` **のみ** | `keisan-card` ❌ / `digital-textbook` ❌ **どちらも消滅** |
| **修正後**（`mikke-shell-v6`） | `keisan-card-static-1.0.0` / `digital-textbook-vendor-v1` / `workbox-precache-v2-…/Digital_textbook/` / `mikke-shell-v6` の**4つが共存** | 両方とも ✅ **無傷** |

### その他

| 確認したこと | 結果 |
|---|---|
| manifest の `id` / `scope` / `start_url` | すべて `/Townmap_Mikke/` |
| アイコン5種の配信 | すべて 200 |
| maskable のセーフゾーン外の非下地画素 | **0.00%** |
| インストールの合図フック | `<head>` 最上部で作動 |
| 320px / 375px 横スクロール | なし |
| `offline.html` | 単体で表示でき、横スクロールなし |
| コンソールのエラー | 0件 |

> **検証の限界**：この環境から `script.google.com` / `accounts.google.com` へは
> 到達できないため、**GAS 本体の起動・GIS サインイン・iframe 内のアプリ動作は
> 確認できていない**。上記は「シェルが正しく組み上がるところまで」の検証である。
> 実機で一度、教員ポータルと児童の入り口の両方を開いて確認していただきたい。

---

## CSP は今回入れていない（手順書として添付）

GIGA Standard v4 は「**確認できない環境なら投入せず、手順書として PR に添える**」と
定めている。このシェルの要となる経路（GAS の iframe 読み込み・GIS サインイン）は
この環境から到達できず、`frame-src` / `connect-src` の過不足を実地で確かめられない。
誤った CSP は**全児童のログインを止める**ため、投入は見送った。

### 投入手順

1. **棚卸し**：`docs/index.html` のインライン `<script>`（約 300 行）を
   `docs/app-shell.js` へ切り出す。`script-src 'self'` で締めるために必要。
   （`pwa-install-hook.js` は既に外部ファイル化済み）
2. 下のブロックを `docs/index.html` の `<head>` 最上部付近へ入れる。

```html
<meta http-equiv="Content-Security-Policy" content="
    default-src 'self';
    script-src 'self' https://accounts.google.com;
    style-src 'self' 'unsafe-inline' https://fonts.googleapis.com;
    font-src 'self' https://fonts.gstatic.com data:;
    img-src 'self' data: blob: https://lh3.googleusercontent.com;
    connect-src 'self' https://script.google.com https://accounts.google.com;
    frame-src https://script.google.com https://*.googleusercontent.com https://accounts.google.com;
    worker-src 'self';
    manifest-src 'self';
    object-src 'none';
    base-uri 'self';
    form-action 'self';
  " />
```

3. `npx serve docs -p 8000` で起動し、**次をすべて実施**してコンソールに
   `Refused to` が **0件**であることを確認する。
   - 教員ポータル（`/`）を開いてサインインできる
   - 児童の入り口（`/?c=クラスコード`）を開いてサインインできる
   - iframe 内のアプリが起動し、地図の読み書きができる
   - `diag.html` の接続診断が通る
4. 1件でも `Refused to` が出たら、**該当ディレクティブだけを緩める**。
   ワイルドカードで潰さない。

> `frame-src` に `*.googleusercontent.com` が必要なのは、GAS のウェブアプリが
> 実際の中身をこのドメインの入れ子 iframe で配信するため。ここを落とすと
> 画面が真っ白になる（`README.md` のトラブルシューティング②と同じ症状）。

---

## 残る積み残し

### 1. `.gs` / `App.html`（GAS 本体）

`.clasp.json` が無く `scriptId` も不明なため、**本番との差分を確認できない**。
以下は未対応：

- `App.html` の Canvas に DPR 補正が無い（`App.html:2594`）
- `App.html` の `user-scalable=no`（読む要素が含まれるかの判断が必要）
- `App.html` が `unpkg.com` / `cdn.tailwindcss.com` を SRI 無しで実行時ロード
- 提示モードが無い（一斉授業で地図を映す用途があるなら必要）
- `App.html` 4,654行 / 211KB の分割

**着手には `scriptId` の共有が必要。** その上で `clasp clone` して差分を確認し、
差分があれば本番側を正として扱う。反映は人間の操作とし、
**既存デプロイの URL は変えない**（新規デプロイを作ると配布済み QR が切れる）。

### 2. OAuth スコープ `spreadsheets`

禁止スコープではないが全スプレッドシートに及ぶ。教員所有のシートを開く設計上、
`spreadsheets.currentonly` には落とせない可能性が高い。**報告に留める。**

### 3. MANUAL.md（先生向け）

→ **2026-08-08 対応済み**（下記「第2次対応」参照）

### 4. 品質ゲート

`Digital_textbook` に置いた `scripts/lib/giga-v4-checks.mjs` を移植したいが、
このリポジトリは npm を使っていないため、実行環境（Node の入れ方・CI）から
決める必要がある。**正本の置き場所とあわせて判断したい。**

---
---

# 第2次対応の結果（2026-08-08）— ドキュメントの整合

**目的**：README / マニュアルを実装と突き合わせ、「読むだけでアプリにできることが
具体的かつ正確に分かる」状態にする。`.gs` / `App.html` の**振る舞いは変更していない**
（実装は読み取りのみ。変更したのは `docs/diag.html` の判定定数 1 箇所とドキュメント）。

## 再判定

| # | 項目 | 前 | 後 | 根拠 |
|---|---|:--:|:--:|---|
| A4 | README.md / MANUAL.md 両方 | ⚠️ | ✅ | `MANUAL.md` を新設（先生向け・専門用語ゼロ・全 9 章）。README は運営者/開発者向けとして整理し直した |
| E10 | iOS の「ホーム画面に追加」手順 | ❌ | ✅ | `MANUAL.md` §7 に iOS(Safari) / Android・Chromebook・Windows(Chrome/Edge) の手順と、更新帯の扱いを記載 |
| E9 | キャッシュ版数の運用 | ⚠️ | ✅ | README に「リリース手順」を追加（`sw.js` の `APP_VERSION` と `diag.html` の `CURRENT_CACHE` を同時に上げる） |

## 実装と食い違っていた記述（修正した）

| 箇所 | これまでの記述 | 実装 |
|---|---|---|
| README アーキテクチャ図 | 児童の iframe は `デプロイS(?mode=student&c=コード)` | シェルは**素の exec URL** で開き、クラスコードは `app:idToken` メッセージで渡す（`docs/index.html` の `frameUrl()` / `App.html` の `Bridge`）。パラメータ付き URL は、12 秒応答が無く起動ビーコンも届かない場合の**1 度きりのフォールバック** |
| README リポジトリ構成 | `docs/` を 1 行で要約 | `sw.js` / `offline.html` / `diag.html` / `pwa-install-hook.js` / manifest / アイコン群を個別に明記 |
| README 全体 | 機能の説明が「2026-07 アップデート」の差分記述のみ | 実装済み機能の全一覧（児童 / 教員ポータル / クラス管理 / 管理パネル 4 タブ / 自動処理）、DB スキーマ、主な制限値の表を追加 |

## 見つけて直したコードの不整合

`docs/diag.html:143` の `CURRENT_CACHE` が `mikke-shell-v5` のままで、
`docs/sw.js` が配っている `mikke-shell-v6` と食い違っていた。
このため**最新のシェルが入っている端末まで「古いシェルが残っています」と
誤判定**し、児童に不要な「キャッシュを全部消して再読み込み」を促していた。

- `CURRENT_CACHE` を `v6` に更新
- 併せて、判定式に残っていた `['v1'..'v4']` の個別リスト（新しい版を足すたびに
  更新が必要で、更新漏れの温床だった）を削除。前方一致 + 現行版との比較だけで
  同じ判定になる
- 同じ食い違いを繰り返さないよう、コードにコメント、README に「リリース手順」を追加

## 検証

| 確認したこと | 結果 |
|---|---|
| README に書いた機能が実装に存在するか | `App.html` / `*.gs` を対照して全項目確認（存在しない機能は記載しない） |
| README に書いた制限値 | `CONFIG`（`Main.gs`）・`vStr_` の各上限・同期間隔・圧縮パラメータの実値と一致 |
| DB スキーマ表 | `Db.gs` の `TABLES` 定義と列名・列順まで一致 |
| ScriptProperties のキー | `PROP_KEYS`（`Main.gs`）・`Registry.gs` と一致 |
| `diag.html` の JS 構文 | `node --check` 相当のパースで確認 |

> **検証の限界**：第1次対応と同じく、この環境から `script.google.com` /
> `accounts.google.com` へは到達できない。**実機での動作確認は行えていない**。
> `diag.html` の修正は判定表示のロジックのみで、接続経路には影響しない。

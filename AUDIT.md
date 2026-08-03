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

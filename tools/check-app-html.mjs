#!/usr/bin/env node
/**
 * check-app-html.mjs — JSX（src/app.jsx）が構文として通るかを見る
 *
 * なぜ要るか:
 *   以前 App.html は 1,856 行の JSX を <script type="text/babel"> に入れて持っていた。
 *   Babel はブラウザの中で動くので、**構文を壊しても push は通り、GAS への反映も通り、
 *   壊れていることが分かるのは児童が開いた瞬間**（真っ白な画面）になった。
 *   いまは npm run build がビルド時に翻訳するので、壊れていればそこで落ちる。
 *   この検査は、ビルドを走らせる前に原本だけを素早く見るためのもの。
 *   .gs は ci.yml が node --check で見ているのに、いちばん大きいこのファイルだけが
 *   何にも見られていなかった。
 *
 * やっていること:
 *   1. <script type="text/babel"> の中身を取り出す
 *   2. GAS のテンプレート記法（<?= ?> / <?!= ?>）を、値が入ったあとの形に置き換える
 *      （置き換えないと Babel から見ればただの構文エラーになる）
 *   3. @babel/core で preset-react を通してパースする
 *
 * 実行: node tools/check-app-html.mjs
 */

import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import path from 'node:path';
import process from 'node:process';

const require = createRequire(import.meta.url);
// 2026-08-28 まで App.html の <script type="text/babel"> を見ていた。
// JSX は src/app.jsx に移り、ビルド時に 1 回だけ翻訳するようになったので、
// 見る先も原本に移す。App.html を見たままにすると、babel の script が
// もう無いので**何も検査しないまま緑になる**。
const file = process.argv[2] || 'src/app.jsx';

let babel;
try {
  babel = require('@babel/core');
  require.resolve('@babel/preset-react');
} catch (e) {
  console.error('[check-app-html] @babel/core / @babel/preset-react が入っていません。');
  console.error('[check-app-html] `npm install` を先に実行してください（この検査は飛ばしません）。');
  process.exit(1);
}

const html = readFileSync(file, 'utf8');

// .jsx ならファイルまるごとが JSX。.html なら <script type="text/babel"> を拾う。
const blocks = [];
if (file.endsWith('.jsx')) {
  blocks.push({ code: html, line: 1 });
} else {
  const re = /<script\b[^>]*type=["']text\/babel["'][^>]*>([\s\S]*?)<\/script>/gi;
  let m;
  while ((m = re.exec(html)) !== null) {
    // 中身の開始位置を控えておく（エラーを元ファイルの行番号で言うため）
    const contentStart = m.index + m[0].length - m[1].length - '</script>'.length;
    const line = html.slice(0, contentStart).split('\n').length;
    blocks.push({ code: m[1], line });
  }
}

if (blocks.length === 0) {
  console.error(`[check-app-html] ${file} に見るべき JSX が 1 つも見つかりませんでした。`);
  console.error('[check-app-html] 作りが変わったか、検査側の想定が古いかのどちらかです。');
  process.exit(1);
}

// GAS のテンプレート記法を、評価後の形に寄せる。
//   <?= bootMode ?>   → 文字列（doGet が入れる値）
//   <?!= x ?>         → 同上（エスケープなし）
//   <? ... ?>         → 制御構文。App.html では使っていないので、見つけたら知らせる
const substitute = (code) => code
  .replace(/<\?!?=\s*[\s\S]*?\?>/g, 'TEMPLATE_VALUE')
  .replace(/<\?[\s\S]*?\?>/g, '');

let failed = 0;
for (const block of blocks) {
  try {
    babel.transformSync(substitute(block.code), {
      filename: path.basename(file) + '.jsx',
      babelrc: false,
      configFile: false,
      presets: [require.resolve('@babel/preset-react')],
      compact: true,
      comments: false,
      sourceType: 'script'
    });
  } catch (err) {
    failed++;
    const at = err.loc ? ` (${file} の ${block.line + err.loc.line - 1} 行目あたり)` : '';
    console.error(`[check-app-html] 構文エラー${at}: ${err.message.split('\n')[0]}`);
  }
}

if (failed) {
  console.error(`[check-app-html] ❌ ${failed} / ${blocks.length} ブロックが構文として通りませんでした。`);
  process.exit(1);
}
console.log(`[check-app-html] ✅ JSX は構文として通りました（${blocks.length} ブロック）`);

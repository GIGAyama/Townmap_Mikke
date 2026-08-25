#!/usr/bin/env node
/**
 * check-syntax.mjs — .gs と docs 配下の .js が構文として壊れていないかを見る
 *
 * GAS は貼り付けて保存するまで構文誤りを教えてくれない。反映は main への push で
 * 自動なので、構文を壊したまま押すと**授業中に画面が出なくなる**。
 * node は .gs 拡張子を知らないので、いったん .js として写してから見る。
 *
 * 実行: node tools/check-syntax.mjs
 */
import { execFileSync } from 'node:child_process';
import { readdirSync, statSync, mkdtempSync, copyFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import process from 'node:process';

const files = [];
for (const name of readdirSync('.')) {
  if (name.endsWith('.gs')) files.push(name);
}
const walk = (dir) => {
  let entries;
  try { entries = readdirSync(dir); } catch (e) { return; }
  for (const name of entries) {
    const full = path.join(dir, name);
    if (statSync(full).isDirectory()) walk(full);
    else if (name.endsWith('.js') || name.endsWith('.mjs')) files.push(full);
  }
};
walk('docs');
walk('tools');
walk('scripts');

if (!files.length) {
  console.error('[check-syntax] 見るファイルが 1 つもありません。想定が古いか、置き場所が変わっています。');
  process.exit(1);
}

const work = mkdtempSync(path.join(tmpdir(), 'mikke-syntax-'));
let failed = 0;
try {
  for (const file of files) {
    // .mjs はモジュールとして、それ以外はスクリプトとして見る
    const ext = file.endsWith('.mjs') ? '.mjs' : '.js';
    const copy = path.join(work, file.replace(/[\\/]/g, '_') + ext);
    copyFileSync(file, copy);
    try {
      execFileSync(process.execPath, ['--check', copy], { stdio: 'pipe' });
    } catch (err) {
      failed++;
      console.error(`[check-syntax] ❌ ${file}`);
      console.error(String(err.stderr || err.message).trim().split('\n').slice(0, 4).join('\n'));
    }
  }
} finally {
  rmSync(work, { recursive: true, force: true });
}

if (failed) {
  console.error(`[check-syntax] ❌ ${failed} / ${files.length} ファイルが構文として通りませんでした。`);
  process.exit(1);
}
console.log(`[check-syntax] ✅ 構文は通りました（${files.length} ファイル）`);

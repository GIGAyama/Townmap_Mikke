#!/usr/bin/env node
/**
 * 生成物が原本と食い違っていないか確かめる。
 *
 * 原本（src/app.jsx・tools/extra.css・tailwind.config.js）を直したのに
 * npm run build を忘れて push すると、リポジトリの中身だけが古いまま残る。
 * 構文検査も通ってしまい、GAS へ貼って開くまで気づけない。
 */
import { readFileSync, existsSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const FILES = ['vendor.html', 'css.html', 'app.html'];
const hash = (f) => createHash('sha256').update(readFileSync(f)).digest('hex').slice(0, 16);

const before = {};
for (const f of FILES) {
  const p = join(ROOT, f);
  if (!existsSync(p)) {
    console.error(`❌ ${f} がありません。\`npm run build\` を実行してください。`);
    process.exit(1);
  }
  before[f] = hash(p);
}
execFileSync(process.execPath, [join(ROOT, 'tools', 'build.mjs')], { stdio: 'pipe' });
let bad = 0;
for (const f of FILES) {
  const after = hash(join(ROOT, f));
  if (before[f] !== after) {
    console.error(`❌ ${f} が原本と食い違っています（${before[f]} → ${after}）。`);
    bad++;
  } else {
    console.log(`✅ ${f} は最新です（${after}）。`);
  }
}
if (bad) {
  console.error('\n原本を直したら `npm run build` を走らせてから push してください。');
  process.exit(1);
}

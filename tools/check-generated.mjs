#!/usr/bin/env node
/**
 * 生成物が原本と食い違っていないか確かめる。
 *
 * 原本（src/app.jsx・tools/extra.css・tailwind.config.js）を直したのに
 * npm run build を忘れて push すると、リポジトリの中身だけが古いまま残る。
 * 構文検査も通ってしまい、GAS へ貼って開くまで気づけない。
 *
 * ⚠️ 比べる相手は「コミットされている中身」であって、作業ツリーのファイルではない。
 *    前の版は作業ツリーを控えてから作り直して比べていたので、先に
 *    `npm run build` が走っているだけで、何を壊しても必ず緑になった
 *    （2026-08-28 に実測）。比べ方は正本 tools/vendor/verify-generated.mjs に
 *    1 つだけ置き、順番で結果が変わらないようにしてある。
 */
import { execFileSync } from 'node:child_process';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

import { verifyAgainstCommitted } from './vendor/verify-generated.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const FILES = ["vendor.html","css.html","app.html"];

const code = await verifyAgainstCommitted(ROOT, FILES, () =>
  execFileSync(process.execPath, [join(ROOT, 'tools', 'build.mjs')], { stdio: 'pipe' }),
);
process.exit(code);

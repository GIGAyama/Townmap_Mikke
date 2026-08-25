/**
 * Schema.gs — コンテナ（配ったスプレッドシート）の作りを点検し、安全な範囲だけ直す
 *
 * ■ なぜ要るか
 *   コンテナバインドで配る形では、**先生がシートを直接開いて触るのが普通の操作**になる。
 *   誤字を直す、いらない行を消す、メモ用の列を足す、シートの名前を変える。どれも起こる。
 *   このアプリは列を見出し名で読み書きするようになった（Db.gs）ので、列を足したり
 *   並べ替えたりしても壊れないが、**見出しごと消された・列が丸ごと無くなった**場合は
 *   読み書きができない。そのとき「何がどうなっているか」を先生の言葉で言うのがここ。
 *
 * ■ 直してよいこと / いけないこと
 *   直す（足す・そろえる）:
 *     - 無いシートを作る
 *     - 1 行も無いシートに見出し行を書く
 *     - 足りない列を**いちばん右に足す**（既存の列は 1 本も動かさない）
 *     - 見出しの表記ゆれ（前後の空白・全角空白・大文字小文字）を正規の書き方にそろえる
 *   直さない（人がやる）:
 *     - 列を消す・動かす・並べ替える
 *     - **見出しの行ごと消えている（1 行目がデータになっている）ときに見出しを書き戻す**
 *       … 間違った列に正しいラベルが付き、そこから先は誰も間違いに気づけなくなる。
 *     - 同じ見出しが 2 つあるときにどちらかを消す
 *
 * ■ 呼び出し口
 *   - スプレッドシートのメニュー「みっけ！」＞「シートを点検する」/「シートを直す」
 *   - 先生の管理パネル（Bound.gs の bdCheckSchema / bdRepairSchema）
 */

/** 点検の対象。クラス DB として使うシート（旧 Users_名簿 は対象外） */
function schemaTables_() {
  return CLASS_TABLE_KEYS.map(function (k) { return TABLES[k]; });
}

/**
 * シートの作りを点検する。**1 セルも書き換えない。**
 *
 * @return {{sheet:string, kind:string, detail:string, fixable:boolean, blocking:boolean}[]}
 *   正常なら空配列。blocking=true は「人が直すまでアプリが正しく読み書きできない」もの。
 */
function checkSchema_(ss) {
  const found = [];
  schemaTables_().forEach(function (table) {
    const sheet = ss.getSheetByName(table.name);
    if (!sheet) {
      found.push({
        sheet: table.name, kind: 'シートが無い',
        detail: '「' + table.name + '」シートがありません。作り直せます（中のデータは戻りません）',
        fixable: true, blocking: true
      });
      return;
    }

    const lastRow = sheet.getLastRow();
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    if (lastRow === 0) {
      found.push({
        sheet: table.name, kind: '見出しが無い',
        detail: '1 行も入っていません。見出し行を書けます',
        fixable: true, blocking: true
      });
      return;
    }

    const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const H = headerMapFromRow_(header, table);
    const foundCols = Object.keys(H);

    // 見出しの行ごと消えている疑い。ここで見出しを書き戻すのがいちばん危ない。
    if (foundCols.length === 0) {
      found.push({
        sheet: table.name, kind: '見出しの行が見あたらない',
        detail: '1 行目に「' + table.cols.join('」「') + '」が 1 つもありません。' +
          '見出しの行ごと消えて、1 行目がデータになっている可能性があります。' +
          '**ここは自動で直しません**（間違った列に正しい名前を付けてしまうため）。' +
          '1 行目に空の行を挿入して、消える前の見出しを手で書き戻してください',
        fixable: false, blocking: true
      });
      return;
    }

    // 同じ見出しが 2 つ以上。左のほうを使うので、どちらが本物か機械には決められない。
    const normHeader = header.map(normalizeHeader_);
    table.cols.forEach(function (name) {
      const key = normalizeHeader_(name);
      let count = 0;
      normHeader.forEach(function (v) { if (v === key) count++; });
      if (count > 1) {
        found.push({
          sheet: table.name, kind: '同じ見出しが 2 つある',
          detail: '「' + name + '」が ' + count + ' か所にあります。左のほうを使うので、' +
            '右のほうに入っているデータは読まれません。どちらか一方の名前を変えてください',
          fixable: false, blocking: false
        });
      }
    });

    // 足りない列。いちばん右に足せる（既存の列は動かさない）。
    const missing = table.cols.filter(function (name) { return H[name] === undefined; });
    if (missing.length) {
      found.push({
        sheet: table.name, kind: '列が足りない',
        detail: '「' + missing.join('」「') + '」の列がありません。いちばん右に足せます',
        fixable: true, blocking: true
      });
    }

    // 表記ゆれ（読めてはいるが、書き方が正規と違う）。
    const typos = table.cols.filter(function (name) {
      return H[name] !== undefined && String(header[H[name]]) !== name;
    });
    if (typos.length) {
      found.push({
        sheet: table.name, kind: '見出しの書き方がちがう',
        detail: '「' + typos.map(function (n) { return String(header[H[n]]) + '」→「' + n; }).join('」/「') +
          '」。読めてはいるので急ぎではありませんが、そろえられます',
        fixable: true, blocking: false
      });
    }

    // 先生が足した列（これは異常ではない。消さないことを伝えるために出す）。
    const known = {};
    Object.keys(H).forEach(function (n) { known[H[n]] = true; });
    const extra = [];
    for (let i = 0; i < header.length; i++) {
      if (!known[i] && normalizeHeader_(header[i]) !== '') extra.push(String(header[i]));
    }
    if (extra.length) {
      found.push({
        sheet: table.name, kind: 'アプリが使わない列がある',
        detail: '「' + extra.join('」「') + '」。アプリは触りません。そのまま置いておけます',
        fixable: false, blocking: false
      });
    }
  });
  return found;
}

/**
 * 安全な範囲だけ直す。**足すことと、書き方をそろえることしかしない。**
 * 消す・動かす・並べ替えるは一切しない。見出しの行ごと消えている場合は何もしない。
 *
 * @return {{done:string[], skipped:string[]}}
 */
function repairSchema_(ss) {
  const done = [];
  const skipped = [];

  withScriptLock_(function () {
    schemaTables_().forEach(function (table) {
      let sheet = ss.getSheetByName(table.name);

      if (!sheet) {
        sheet = ss.insertSheet(table.name);
        writeHeaderRow_(sheet, table);
        done.push('「' + table.name + '」シートを作りました');
        return;
      }
      if (sheet.getLastRow() === 0) {
        writeHeaderRow_(sheet, table);
        done.push('「' + table.name + '」に見出し行を書きました');
        return;
      }

      const lastCol = Math.max(sheet.getLastColumn(), 1);
      const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      const H = headerMapFromRow_(header, table);

      if (Object.keys(H).length === 0) {
        skipped.push('「' + table.name + '」は見出しの行が見あたりません。' +
          '自動では直しません（間違った列に名前を付けてしまうため）。手で戻してください');
        return;
      }

      // (1) 表記ゆれをそろえる（位置は動かさない）
      const fixes = [];
      table.cols.forEach(function (name) {
        if (H[name] !== undefined && String(header[H[name]]) !== name) {
          sheet.getRange(1, H[name] + 1).setValue(name);
          fixes.push(name);
        }
      });
      if (fixes.length) done.push('「' + table.name + '」の見出し「' + fixes.join('」「') + '」の書き方をそろえました');

      // (2) 足りない列をいちばん右に足す（既存の列は 1 本も動かさない）
      const missing = table.cols.filter(function (name) { return H[name] === undefined; });
      if (missing.length) {
        const maxCols = sheet.getMaxColumns();
        const need = lastCol + missing.length;
        if (need > maxCols) sheet.insertColumnsAfter(maxCols, need - maxCols);
        sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing])
          .setBackground('#41B3A3').setFontColor('white').setFontWeight('bold');
        done.push('「' + table.name + '」の右に列「' + missing.join('」「') + '」を足しました' +
          '（すでに入っているデータは動かしていません）');
      }

      if (sheet.getFrozenRows() < 1) sheet.setFrozenRows(1);
    });
  });

  return { done: done, skipped: skipped };
}

/** 点検結果を、先生が読む 1 つの文字列にする */
function formatSchemaReport_(found) {
  if (!found.length) return 'シートの作りは想定どおりです。';
  const lines = found.map(function (f) {
    return (f.blocking ? '⚠ ' : '・') + '「' + f.sheet + '」' + f.kind + '：' + f.detail;
  });
  const fixable = found.filter(function (f) { return f.fixable; }).length;
  lines.push('');
  lines.push(fixable
    ? 'このうち ' + fixable + ' 件は「シートを直す」で足せます（列を消したり動かしたりはしません）。'
    : '自動で直せるものはありません。上の内容にそって手で直してください。');
  return lines.join('\n');
}

// ────────────────────────────────────────────────────────────────
// スプレッドシートのメニュー（コンテナバインドのときだけ意味がある）
// ────────────────────────────────────────────────────────────────

/**
 * スプレッドシートを開いたときのメニュー。
 * 独立スクリプト（ウェブアプリだけ）の文脈では呼ばれない。
 */
function onOpen(e) {
  try {
    SpreadsheetApp.getUi()
      .createMenu(CONFIG.APP_NAME)
      .addItem('はじめの設定（先生として登録）', 'runInitialSetup')
      .addSeparator()
      .addItem('シートを点検する', 'showSheetCheck')
      .addItem('シートを直す（足すだけ）', 'runSheetRepair')
      .addSeparator()
      .addItem('この学級の設定を見る', 'showClassInfo')
      .addToUi();
  } catch (err) {
    // ウェブアプリとして動いているときは画面が無い。何もしない。
  }
}

/**
 * メニュー「はじめの設定（先生として登録）」。コピーを作ったあと、1 回だけ押す。
 *
 * ここで先生を記録するのは、**メニューはそのファイルの編集権を持つ人にしか出せない**
 * ため。ウェブアプリ側で「最初に開いた人を先生にする」形にすると、先生が URL を配った
 * あとで最初に開いた児童が恒久的に先生になってしまう。
 *
 * ⚠️ getUi() を先に取るのは showSheetCheck と同じ理由（ウェブアプリ文脈ではここで止まる）。
 *
 * なお `onOpen` の中でこれを自動で走らせることはできない。単純トリガーは
 * 承認の要るサービス（`Session.getEffectiveUser().getEmail()` は userinfo.email が要る）を
 * 呼べないため、黙って失敗する。先生に 1 回押してもらう形にしてある。
 */
function runInitialSetup() {
  const ui = SpreadsheetApp.getUi();          // 画面が無ければ、ここで止まる
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 足りないシートを作る。**先生が自分で足したシートは消さない**
  //（initializeNewDatabase_ は既定シートの掃除もするので、ここでは使わない）
  CLASS_TABLE_KEYS.forEach(function (k) { ensureSheet_(ss, TABLES[k]); });

  const before = boundOwner_(ss);
  const owner = boundClaimOwnerFromContainer_(ss);
  const found = checkSchema_(ss);

  const lines = [];
  if (!owner) {
    lines.push('先生の登録ができませんでした。もう一度お試しください。');
  } else if (before) {
    lines.push('すでに ' + before + ' が先生として登録されています。');
    lines.push('（別の先生に引き継ぐときは Settings シートの ownerEmail を書き換えてください）');
  } else {
    lines.push('先生として登録しました: ' + owner);
  }
  lines.push('');
  lines.push(formatSchemaReport_(found));
  lines.push('');
  lines.push('つぎは「拡張機能 ＞ Apps Script ＞ デプロイ ＞ 新しいデプロイ」で、');
  lines.push('次のユーザーとして実行＝「自分」、アクセスできるユーザー＝「同一ドメインの全員」');
  lines.push('で公開し、出てきた URL を学級に配ってください。');
  ui.alert('はじめの設定', lines.join('\n'), ui.ButtonSet.OK);
}

/**
 * メニュー「シートを点検する」。
 *
 * ⚠️ google.script.run は末尾 `_` の無い関数を誰でも呼べるので、この関数も
 *    児童のブラウザから呼べてしまう。**先に getUi() を取る**のはそのため。
 *    画面が無い文脈（ウェブアプリ）ではここで例外になり、シートを 1 枚も読まずに終わる。
 *    返すのは見出しの並びだけなので、通ったとしても児童の記録や名前は出ない。
 */
function showSheetCheck() {
  const ui = SpreadsheetApp.getUi();          // 画面が無ければ、ここで止まる
  const found = checkSchema_(SpreadsheetApp.getActiveSpreadsheet());
  ui.alert('シートの点検', formatSchemaReport_(found), ui.ButtonSet.OK);
}

/**
 * メニュー「シートを直す（足すだけ）」。
 * getUi() を先に取るのは showSheetCheck と同じ理由。さらに、実際に書き換えるので
 * 先生に 1 回確認してから走らせる。
 */
function runSheetRepair() {
  const ui = SpreadsheetApp.getUi();          // 画面が無ければ、ここで止まる
  const answer = ui.alert('シートを直す',
    '足りないシートと列を足し、見出しの書き方をそろえます。\n' +
    '列を消したり、並べ替えたりはしません。よろしいですか？',
    ui.ButtonSet.OK_CANCEL);
  if (answer !== ui.Button.OK) return;

  const result = repairSchema_(SpreadsheetApp.getActiveSpreadsheet());
  const lines = [];
  if (result.done.length) lines.push('直したもの:\n' + result.done.map(function (t) { return '・' + t; }).join('\n'));
  if (result.skipped.length) lines.push('直さなかったもの:\n' + result.skipped.map(function (t) { return '・' + t; }).join('\n'));
  if (!lines.length) lines.push('直すところはありませんでした。');
  ui.alert('シートを直しました', lines.join('\n\n'), ui.ButtonSet.OK);
}

/**
 * メニュー「この学級の設定を見る」。
 * 参加の受付・承認・学級名だけを出す。API キーや児童の記録は出さない。
 * getUi() を先に取るのは showSheetCheck と同じ理由。
 */
function showClassInfo() {
  const ui = SpreadsheetApp.getUi();          // 画面が無ければ、ここで止まる
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const members = getTableData_(ss, TABLES.MEMBERS);
  const active = members.filter(function (m) { return m.status === 'active' && m.role !== 'teacher'; }).length;
  const pending = members.filter(function (m) { return m.status === 'pending'; }).length;
  ui.alert('この学級の設定', [
    '学級名: ' + (getSettingValue_(ss, 'className') || '（未設定）'),
    '参加の受付: ' + (getSettingValue_(ss, 'joinOpen') === 'false' ? '閉じている' : '開いている'),
    '参加に承認が必要: ' + (getSettingValue_(ss, 'requireApproval') === 'false' ? 'いいえ' : 'はい'),
    'いま使える児童: ' + active + ' 人（承認待ち ' + pending + ' 人）',
    '',
    '児童に配る URL は、Apps Script の「デプロイ」＞「デプロイを管理」で確かめられます。'
  ].join('\n'), ui.ButtonSet.OK);
}

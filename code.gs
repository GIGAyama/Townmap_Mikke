/**
 * まちたんけんマップ - Server Side Script
 * v2.0 Refactored for Robustness and Maintainability
 */

// ==========================================
//  設定・定数定義
// ==========================================
const CONFIG = {
  SHEET_NAME: 'pin_data',
  IMAGE_FOLDER_NAME: 'まちたんけん_画像データ',
  ADMIN_PASSWORD: 'sensei', // 必要に応じて変更してください
  DEFAULT_MAP_URL: "https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEj7b2aHhD8U0x0W3v2qVzQJjNfLwX6yZ8lK1oE4rT5yU9iO0pA3sD4fG7hJ2kL5nB8mV0cW3eR6tY9uI1oP4aS7dF0gH2jK5lZ8xC3vB6n/s1600/map_town_illustration.png",
  LOCK_TIMEOUT: 10000 // 排他制御の待機時間(ms)
};

// スプレッドシートの列定義（列の追加・入替時はここを修正）
const COLUMNS = {
  ID: 0,
  TIMESTAMP: 1,
  PIN_COLOR: 2,
  AUTHOR_NAME: 3,
  MAP_X: 4,
  MAP_Y: 5,
  SHOP_NAME: 6,
  DESCRIPTION: 7,
  IMAGE_URL: 8,
  GROUP_ID: 9
};

// ==========================================
//  Webアプリのエントリーポイント
// ==========================================
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('みんなのまちたんけんマップ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://drive.google.com/uc?id=1ffTLalkZVzwAIQtCDFN5OzR3CePkzDjh&.png');
  return htmlOutput;;
}

// ==========================================
//  API: データ取得 (Read)
// ==========================================
function getPinData() {
  const result = { status: 'success', data: [], map_url: '' };
  
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    
    // プロパティから現在の地図URLを取得
    const props = PropertiesService.getScriptProperties();
    result.map_url = props.getProperty('MAP_URL') || CONFIG.DEFAULT_MAP_URL;

    if (lastRow > 1) {
      // ヘッダー行を除いてデータを取得
      const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
      
      // 配列データをオブジェクトにマッピング
      result.data = data.map(row => ({
        id:          String(row[COLUMNS.ID]),
        timestamp:   row[COLUMNS.TIMESTAMP], // 日付型としてクライアントに渡す
        pin_color:   String(row[COLUMNS.PIN_COLOR]),
        author_name: String(row[COLUMNS.AUTHOR_NAME]),
        map_x:       Number(row[COLUMNS.MAP_X]),
        map_y:       Number(row[COLUMNS.MAP_Y]),
        shop_name:   String(row[COLUMNS.SHOP_NAME]),
        description: String(row[COLUMNS.DESCRIPTION]),
        image_url:   String(row[COLUMNS.IMAGE_URL]),
        group_id:    String(row[COLUMNS.GROUP_ID] || 'all')
      }));
    }
  } catch (error) {
    console.error('getPinData Error:', error);
    result.status = 'error';
    result.message = error.toString();
  }

  return JSON.stringify(result);
}

// ==========================================
//  API: データ保存・更新・削除 (Write)
// ==========================================
function savePinData(payloadJson) {
  const lock = LockService.getScriptLock();
  
  // 排他制御：同時に書き込みが来た場合、順番待ちさせる
  if (lock.tryLock(CONFIG.LOCK_TIMEOUT)) {
    const result = { status: 'success', message: '' };
    
    try {
      const params = JSON.parse(payloadJson);
      const action = params.action;
      const sheet = getSheet();

      if (action === 'save_pin') {
        _handleSavePin(sheet, params, result);
      } else if (action === 'delete_pin') {
        _handleDeletePin(sheet, params, result);
      } else {
        throw new Error('不明なアクションです');
      }

    } catch (error) {
      console.error('savePinData Error:', error);
      result.status = 'error';
      result.message = error.toString();
    } finally {
      lock.releaseLock();
    }
    return JSON.stringify(result);
  } else {
    return JSON.stringify({
      status: 'error',
      message: 'サーバーが混み合っています。少し待ってからもう一度押してください。'
    });
  }
}

// ==========================================
//  API: 設定変更 (Map URL)
// ==========================================
function saveMapSetting(payloadJson) {
  const result = { status: 'success', message: '' };
  try {
    const params = JSON.parse(payloadJson);
    if (params.password !== CONFIG.ADMIN_PASSWORD) {
      throw new Error('パスワードがちがいます');
    }
    PropertiesService.getScriptProperties().setProperty('MAP_URL', params.map_url);
    result.message = '地図を変更しました';
  } catch (error) {
    result.status = 'error';
    result.message = error.toString();
  }
  return JSON.stringify(result);
}

// ==========================================
//  内部ロジック (Private Helpers)
// ==========================================

/**
 * ピンの新規保存処理
 */
function _handleSavePin(sheet, params, result) {
  let imageUrl = '';
  if (params.image_data) {
    imageUrl = saveImageToDrive(params.image_data, params.id);
  }

  const now = new Date();
  const formattedDate = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

  // appendRowのために配列を作成（順序はCOLUMNS定義に従う必要はないが、appendRowは左から埋めるため注意）
  // 堅牢にするため、全ての要素数を持つ配列を初期化
  const newRow = new Array(Object.keys(COLUMNS).length).fill('');
  
  newRow[COLUMNS.ID] = params.id;
  newRow[COLUMNS.TIMESTAMP] = formattedDate;
  newRow[COLUMNS.PIN_COLOR] = params.pin_color;
  newRow[COLUMNS.AUTHOR_NAME] = params.author_name;
  newRow[COLUMNS.MAP_X] = params.map_x;
  newRow[COLUMNS.MAP_Y] = params.map_y;
  newRow[COLUMNS.SHOP_NAME] = params.shop_name;
  newRow[COLUMNS.DESCRIPTION] = params.description;
  newRow[COLUMNS.IMAGE_URL] = imageUrl;
  newRow[COLUMNS.GROUP_ID] = params.group_id;

  sheet.appendRow(newRow);
  result.message = 'ピンをさしました！';
}

/**
 * ピンの削除処理
 */
function _handleDeletePin(sheet, params, result) {
  if (params.password !== CONFIG.ADMIN_PASSWORD) {
    throw new Error('パスワードがちがいます');
  }

  const data = sheet.getDataRange().getValues();
  // ヘッダー行(0)を除く、末尾から検索して削除（行ずれ防止のため後ろから）
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][COLUMNS.ID]) === String(params.target_id)) {
      sheet.deleteRow(i + 1); // deleteRowは1-based index
      result.message = '削除しました';
      return;
    }
  }
  throw new Error('指定されたIDが見つかりませんでした');
}

/**
 * シート取得・初期化
 */
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    // ヘッダー生成
    const headers = Object.keys(COLUMNS).sort((a, b) => COLUMNS[a] - COLUMNS[b]);
    sheet.appendRow(headers);
    // 1行目を固定
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 画像をDriveに保存し、thumbnailリンクを返す
 */
function saveImageToDrive(base64Data, fileName) {
  try {
    const folderId = getOrCreateImageFolderId();
    const folder = DriveApp.getFolderById(folderId);
    
    // "data:image/jpeg;base64,..." の形式からデータ部のみ抽出
    const match = base64Data.match(/^data:([a-zA-Z0-9]+\/[a-zA-Z0-9-.+]+);base64,(.+)$/);
    if (!match) return '';

    const contentType = match[1];
    const decodedData = Utilities.base64Decode(match[2]);
    const blob = Utilities.newBlob(decodedData, contentType, fileName + '.jpg');
    
    const file = folder.createFile(blob);
    
    // 権限設定: リンクを知っている全員が閲覧可
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Google Driveのサムネイル生成APIを利用 (sz=w1000 で幅1000px指定)
    return `https://drive.google.com/thumbnail?sz=w1000&id=${file.getId()}`;
    
  } catch (e) {
    console.error('Image Save Error:', e);
    return ''; // 画像保存に失敗してもデータ登録は進めるため空文字を返す
  }
}

/**
 * 画像保存フォルダの取得または作成
 */
function getOrCreateImageFolderId() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('IMAGE_FOLDER_ID');
  
  if (savedId) {
    try {
      DriveApp.getFolderById(savedId);
      return savedId;
    } catch (e) {
      // 保存されていたIDが無効な場合は再検索へ進む
    }
  }

  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const ssFile = DriveApp.getFileById(ssId);
  // スプレッドシートと同じフォルダに画像を保存
  const parentParents = ssFile.getParents();
  const parentFolder = parentParents.hasNext() ? parentParents.next() : DriveApp.getRootFolder();
  
  const folders = parentFolder.getFoldersByName(CONFIG.IMAGE_FOLDER_NAME);
  let targetFolder;
  
  if (folders.hasNext()) {
    targetFolder = folders.next();
  } else {
    targetFolder = parentFolder.createFolder(CONFIG.IMAGE_FOLDER_NAME);
  }
  
  const newId = targetFolder.getId();
  props.setProperty('IMAGE_FOLDER_ID', newId);
  return newId;
}

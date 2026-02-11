/**
 * まちたんけんマップ - Server Side Script (GIGA Standard v2)
 * v3.1 Added Teacher Dashboard & Dynamic Group Management
 */

// ==========================================
//  設定・定数定義
// ==========================================
const CONFIG = {
  SHEET_NAME: 'pin_data',
  GROUP_SHEET_NAME: 'group_config',
  IMAGE_FOLDER_NAME: 'まちたんけん_画像データ',
  ADMIN_PASSWORD: 'sensei', // 必要に応じて変更してください
  DEFAULT_MAP_URL: "https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEj7b2aHhD8U0x0W3v2qVzQJjNfLwX6yZ8lK1oE4rT5yU9iO0pA3sD4fG7hJ2kL5nB8mV0cW3eR6tY9uI1oP4aS7dF0gH2jK5lZ8xC3vB6n/s1600/map_town_illustration.png",
  LOCK_TIMEOUT: 10000, // 排他制御の待機時間(ms)
  GEMINI_MODEL: 'gemini-1.5-flash'
};

// スプレッドシートの列定義
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
  GROUP_ID: 9,
  DELETED_AT: 10
};

// ==========================================
//  Webアプリのエントリーポイント
// ==========================================
function doGet(e) {
  setupEnvironment();
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('みんなのまちたんけんマップ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://drive.google.com/uc?id=1ffTLalkZVzwAIQtCDFN5OzR3CePkzDjh&.png');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
//  API: データ取得 (Read)
// ==========================================
function getPinData() {
  const result = { status: 'success', data: [], map_url: '', groups: [] };
  
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    
    // プロパティから現在の地図URLを取得
    const props = PropertiesService.getScriptProperties();
    result.map_url = props.getProperty('MAP_URL') || CONFIG.DEFAULT_MAP_URL;

    // グループ情報の取得
    result.groups = _getGroups();

    if (lastRow > 1) {
      const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
      result.data = data
        .filter(row => !row[COLUMNS.DELETED_AT])
        .map(row => ({
          id:          String(row[COLUMNS.ID]),
          timestamp:   formatDate(row[COLUMNS.TIMESTAMP]),
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
      message: '他の人が使っています。少し待ってからもう一度押してください。'
    });
  }
}

// ==========================================
//  API: 設定変更 (Map URL & Groups)
// ==========================================
function saveMapSetting(payloadJson) {
  const result = { status: 'success', message: '' };
  try {
    const params = JSON.parse(payloadJson);
    if (params.password !== CONFIG.ADMIN_PASSWORD) {
      throw new Error('パスワードがちがいます');
    }
    
    // 地図URLの保存
    if (params.map_url !== undefined) {
      PropertiesService.getScriptProperties().setProperty('MAP_URL', params.map_url);
    }

    result.message = '設定を変更しました';
  } catch (error) {
    result.status = 'error';
    result.message = error.toString();
  }
  return JSON.stringify(result);
}

function saveGroupData(payloadJson) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(CONFIG.LOCK_TIMEOUT)) {
    const result = { status: 'success', message: '' };
    try {
      const params = JSON.parse(payloadJson);
      if (params.password !== CONFIG.ADMIN_PASSWORD) {
        throw new Error('パスワードがちがいます');
      }

      // グループシートの更新 (完全洗い替え)
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName(CONFIG.GROUP_SHEET_NAME);
      if (!sheet) sheet = ss.insertSheet(CONFIG.GROUP_SHEET_NAME);
      
      sheet.clear();
      sheet.appendRow(['id', 'name']); // Header
      
      const rows = params.groups.map(g => [g.id, g.name]);
      if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, 2).setValues(rows);
      }
      
      result.message = 'グループ情報を更新しました';
    } catch (error) {
      result.status = 'error';
      result.message = error.toString();
    } finally {
      lock.releaseLock();
    }
    return JSON.stringify(result);
  } else {
    return JSON.stringify({ status: 'error', message: '混み合っています' });
  }
}

// ==========================================
//  内部ロジック (Private Helpers)
// ==========================================

function setupEnvironment() {
  const props = PropertiesService.getScriptProperties();
  const isSetup = props.getProperty('IS_SETUP');

  if (!isSetup) {
    getSheet(); // ピンデータシート作成
    _getGroups(); // グループシート作成（なければデフォルト作成）
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const defaultSheet = ss.getSheetByName('シート1');
    if (defaultSheet && ss.getSheets().length > 1) {
      ss.deleteSheet(defaultSheet);
    }
    props.setProperty('IS_SETUP', 'true');
  }
}

function _getGroups() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.GROUP_SHEET_NAME);
  
  if (!sheet) {
    // デフォルトグループの作成
    sheet = ss.insertSheet(CONFIG.GROUP_SHEET_NAME);
    sheet.appendRow(['id', 'name']);
    const defaultGroups = [];
    for(let i=1; i<=10; i++) {
        defaultGroups.push([String(i), `${i}ぱん`]);
    }
    sheet.getRange(2, 1, defaultGroups.length, 2).setValues(defaultGroups);
    return defaultGroups.map(r => ({ id: r[0], name: r[1] }));
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  return data.map(row => ({ id: String(row[0]), name: String(row[1]) }));
}

function _handleSavePin(sheet, params, result) {
  let imageUrl = '';
  if (params.image_data) {
    imageUrl = saveImageToDrive(params.image_data, params.id);
  }

  const now = new Date();
  const newRow = new Array(Object.keys(COLUMNS).length).fill('');
  
  newRow[COLUMNS.ID] = params.id;
  newRow[COLUMNS.TIMESTAMP] = now;
  newRow[COLUMNS.PIN_COLOR] = params.pin_color;
  newRow[COLUMNS.AUTHOR_NAME] = params.author_name;
  newRow[COLUMNS.MAP_X] = params.map_x;
  newRow[COLUMNS.MAP_Y] = params.map_y;
  newRow[COLUMNS.SHOP_NAME] = params.shop_name;
  newRow[COLUMNS.DESCRIPTION] = params.description;
  newRow[COLUMNS.IMAGE_URL] = imageUrl;
  newRow[COLUMNS.GROUP_ID] = params.group_id;
  newRow[COLUMNS.DELETED_AT] = '';

  sheet.appendRow(newRow);
  result.message = 'ピンをさしました！';
}

function _handleDeletePin(sheet, params, result) {
  if (params.password !== CONFIG.ADMIN_PASSWORD) {
    throw new Error('パスワードがちがいます');
  }
  const data = sheet.getDataRange().getValues();
  let targetRowIndex = -1;
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][COLUMNS.ID]) === String(params.target_id)) {
      targetRowIndex = i + 1;
      break;
    }
  }
  if (targetRowIndex > 0) {
    const now = new Date();
    sheet.getRange(targetRowIndex, COLUMNS.DELETED_AT + 1).setValue(now);
    result.message = '削除しました';
  } else {
    throw new Error('指定されたIDが見つかりませんでした');
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    const headers = Object.keys(COLUMNS).sort((a, b) => COLUMNS[a] - COLUMNS[b]);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#E8F0FE').setFontWeight('bold');
  }
  return sheet;
}

function saveImageToDrive(base64Data, fileName) {
  try {
    const folderId = getOrCreateImageFolderId();
    const folder = DriveApp.getFolderById(folderId);
    const match = base64Data.match(/^data:([a-zA-Z0-9]+\/[a-zA-Z0-9-.+]+);base64,(.+)$/);
    if (!match) return '';
    const contentType = match[1];
    const decodedData = Utilities.base64Decode(match[2]);
    const blob = Utilities.newBlob(decodedData, contentType, fileName + '.jpg');
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return `https://drive.google.com/thumbnail?sz=w1000&id=${file.getId()}`;
  } catch (e) {
    console.error('Image Save Error:', e);
    return '';
  }
}

function getOrCreateImageFolderId() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('IMAGE_FOLDER_ID');
  if (savedId) {
    try {
      DriveApp.getFolderById(savedId);
      return savedId;
    } catch (e) {}
  }
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const ssFile = DriveApp.getFileById(ssId);
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

function formatDate(date) {
  if (!date) return '';
  try {
    return Utilities.formatDate(new Date(date), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  } catch (e) {
    return String(date);
  }
}


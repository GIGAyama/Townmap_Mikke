/**
 * みっけ！
 */

const CONFIG = {
  APP_NAME: 'みっけ！', 
  IMAGE_FOLDER_NAME: 'みっけ！_画像データ',
  LOCK_TIMEOUT: 15000
};

const TABLES = {
  USERS: { name: 'Users_名簿', cols: ['email', 'name', 'group_id', 'role', 'created_at'] },
  UNITS: { name: 'Units_単元', cols: ['unit_id', 'name', 'maps_json', 'chat_enabled', 'stamp_enabled', 'custom_stamps', 'is_active', 'created_at'] },
  PINS:  { name: 'Pins_ピン', cols: ['pin_id', 'unit_id', 'map_id', 'email', 'x', 'y', 'color', 'title', 'memo', 'image_url', 'created_at'] },
  CHATS: { name: 'Chats_チャット', cols: ['chat_id', 'unit_id', 'email', 'message', 'target_type', 'target_id', 'created_at'] },
  REACTIONS: { name: 'Reactions_反応', cols: ['reaction_id', 'unit_id', 'email', 'target_type', 'target_id', 'emoji', 'created_at'] }
};

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate().setTitle('みっけ！')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://drive.google.com/uc?id=1yOrXP3u-S3B1CzW7iCh1DhPloH1gsPgt&.png');
}

function getDB() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('DB_ID'); 
  let ss = null;
  if (ssId) { try { ss = SpreadsheetApp.openById(ssId); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create(CONFIG.APP_NAME);
    props.setProperty('DB_ID', ss.getId());
    ss.getSheets()[0].setName('Dummy');
  }
  Object.values(TABLES).forEach(table => {
    let sheet = ss.getSheetByName(table.name);
    if (!sheet) {
      sheet = ss.insertSheet(table.name);
      sheet.appendRow(table.cols);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, table.cols.length).setBackground('#41B3A3').setFontColor('white').setFontWeight('bold');
    }
  });
  const dummy = ss.getSheetByName('Dummy');
  if (dummy) ss.deleteSheet(dummy);
  return ss;
}

function getTableData(sheetName) {
  const sheet = getDB().getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const headers = TABLES[Object.keys(TABLES).find(k => TABLES[k].name === sheetName)].cols;
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values.map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function formatUnit(unit) {
  if (!unit) return null;
  try { unit.maps = JSON.parse(unit.maps_json || '[]'); } catch(e) { unit.maps = []; }
  try { unit.custom_stamps = JSON.parse(unit.custom_stamps || '["📍","🐛","🌸","🚗","⚠️","🏠","❓","💡"]'); } catch(e) { unit.custom_stamps = ["📍","🐛","🌸","🚗","⚠️","🏠","❓","💡"]; }
  return unit;
}

function getInitData() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) return JSON.stringify({ status: 'error', message: 'Googleアカウントにログインしていません。' });
    getDB();
    
    let users = getTableData(TABLES.USERS.name);
    let myUser = users.find(u => u.email === email);

    if (users.length === 0) {
      const newTeacher = { email: email, name: '先生', group_id: 'teacher', role: 'teacher', created_at: new Date().toLocaleString() };
      getDB().getSheetByName(TABLES.USERS.name).appendRow([newTeacher.email, newTeacher.name, newTeacher.group_id, newTeacher.role, newTeacher.created_at]);
      myUser = newTeacher;
      users = [newTeacher];
    }

    if (!myUser) return JSON.stringify({ status: 'unregistered', email: email });

    const units = getTableData(TABLES.UNITS.name);
    const activeUnit = formatUnit(units.find(u => u.is_active === true) || units[0] || null);

    let pins = [];
    let chats = [];
    let reactions = [];
    if (activeUnit) {
      pins = getTableData(TABLES.PINS.name).filter(p => p.unit_id === activeUnit.unit_id);
      chats = getTableData(TABLES.CHATS.name).filter(c => c.unit_id === activeUnit.unit_id);
      reactions = getTableData(TABLES.REACTIONS.name).filter(r => r.unit_id === activeUnit.unit_id);
    }

    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    const hasApiKey = !!apiKey;

    return JSON.stringify({ status: 'success', user: myUser, users, activeUnit, units, pins, chats, reactions, hasApiKey });
  } catch (e) { return JSON.stringify({ status: 'error', message: e.toString() }); }
}

function syncData(unitId) {
  try {
    const units = getTableData(TABLES.UNITS.name);
    const activeUnit = formatUnit(units.find(u => u.unit_id === unitId));
    const pins = getTableData(TABLES.PINS.name).filter(p => p.unit_id === unitId);
    const chats = getTableData(TABLES.CHATS.name).filter(c => c.unit_id === unitId);
    const reactions = getTableData(TABLES.REACTIONS.name).filter(r => r.unit_id === unitId);
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    return JSON.stringify({ status: 'success', pins, chats, reactions, activeUnit, hasApiKey: !!apiKey });
  } catch (e) { return JSON.stringify({ status: 'error', message: e.toString() }); }
}

function getDriveImages() {
  try {
    const files = DriveApp.searchFiles("mimeType contains 'image/' and trashed = false");
    const images = [];
    let count = 0;
    while (files.hasNext() && count < 60) {
      const file = files.next();
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
      images.push({ id: file.getId(), name: file.getName(), thumbnail: `https://drive.google.com/thumbnail?sz=w400&id=${file.getId()}` });
      count++;
    }
    return JSON.stringify({ status: 'success', images });
  } catch (e) { return JSON.stringify({ status: 'error', message: e.toString() }); }
}

function generateAIPortfolio(payloadJson) {
  try {
    const p = JSON.parse(payloadJson);
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return JSON.stringify({ status: 'error', message: 'AI分析を行うには、GASのスクリプトプロパティに「GEMINI_API_KEY」を設定してください。' });

    const pins = getTableData(TABLES.PINS.name).filter(pin => pin.unit_id === p.unit_id && pin.email === p.email);
    const chats = getTableData(TABLES.CHATS.name).filter(chat => chat.unit_id === p.unit_id && chat.email === p.email);
    const reactions = getTableData(TABLES.REACTIONS.name).filter(r => r.unit_id === p.unit_id && r.email === p.email);
    const user = getTableData(TABLES.USERS.name).find(u => u.email === p.email);

    if (pins.length === 0 && chats.length === 0 && reactions.length === 0) {
      return JSON.stringify({ status: 'success', portfolio: 'まだ活動の記録（ピンやチャット）がありません。' });
    }

    let prompt = `あなたは小学校の先生です。児童「${user ? user.name : 'この児童'}」の「地図学習」での活動記録を分析し、温かいフィードバックを作成してください。\n\n`;
    prompt += `【ピンを刺した記録】\n`;
    pins.forEach(pin => prompt += `- 発見対象[${pin.title}]: メモ[${pin.memo || 'なし'}] アイコン[${pin.color}]\n`);
    prompt += `\n【発言記録】\n`;
    chats.forEach(chat => prompt += `- ${chat.message}\n`);
    prompt += `\n【友達へのリアクション回数】: ${reactions.length}回\n`;
    prompt += `\n以下の3項目で出力してください。\n1. 🔍 興味関心の傾向（どんなものに目を向けているか）\n2. ✨ 素晴らしい点（表現や友達への関わりの良さ）\n3. 💌 先生からのメッセージ（小学生に向けて優しい言葉で）`;

    const payload = {
      contents: [{ parts: [{ text: prompt }] }],
      systemInstruction: { parts: [{ text: "あなたは優しく、児童の良いところを見つけるのが得意な先生です。マークダウンを使用せず、プレーンテキストで見やすく出力してください。"}] }
    };

    const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
    const response = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`, options);
    const resData = JSON.parse(response.getContentText());
    if (resData.error) throw new Error(resData.error.message);
    
    return JSON.stringify({ status: 'success', portfolio: resData.candidates[0].content.parts[0].text });
  } catch (e) { return JSON.stringify({ status: 'error', message: e.toString() }); }
}

function executeAction(payloadJson) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(CONFIG.LOCK_TIMEOUT)) return JSON.stringify({ status: 'error', message: 'サーバー混雑中' });

  try {
    const p = JSON.parse(payloadJson);
    const db = getDB();
    const now = new Date().toLocaleString('ja-JP');

    if (p.action === 'save_pin') {
      db.getSheetByName(TABLES.PINS.name).appendRow([p.pin_id, p.unit_id, p.map_id, p.email, p.x, p.y, p.color, p.title, p.memo, p.image_url || '', now]);
    } 
    else if (p.action === 'save_chat') {
      db.getSheetByName(TABLES.CHATS.name).appendRow([p.chat_id, p.unit_id, p.email, p.message, p.target_type || 'general', p.target_id || '', now]);
    }
    else if (p.action === 'toggle_reaction') {
      const sheet = db.getSheetByName(TABLES.REACTIONS.name);
      const data = sheet.getDataRange().getValues();
      let foundRowIndex = -1;
      for (let i = 1; i < data.length; i++) {
        if (data[i][2] === p.email && data[i][3] === p.target_type && data[i][4] === p.target_id && data[i][5] === p.emoji) {
          foundRowIndex = i + 1; break;
        }
      }
      if (foundRowIndex > -1) sheet.deleteRow(foundRowIndex);
      else sheet.appendRow([`r_${Date.now()}`, p.unit_id, p.email, p.target_type, p.target_id, p.emoji, now]);
    }
    else if (p.action === 'delete_pin') { _deleteRowByColId(db.getSheetByName(TABLES.PINS.name), 0, p.pin_id); }
    else if (p.action === 'delete_chat') { _deleteRowByColId(db.getSheetByName(TABLES.CHATS.name), 0, p.chat_id); }
    else if (p.action === 'save_unit') {
      const unitSheet = db.getSheetByName(TABLES.UNITS.name);
      const data = unitSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) { 
        if (data[i][6] === true) unitSheet.getRange(i + 1, 7).setValue(false); 
      }
      const initMap = [{ id: 'm_' + Date.now(), name: p.map_name || '基本マップ', url: p.map_url }];
      const initStamps = JSON.stringify(['📍', '🐛', '🌸', '🚗', '⚠️', '🏠', '❓', '💡']);
      unitSheet.appendRow([p.unit_id, p.name, JSON.stringify(initMap), true, true, initStamps, true, now]); 
    }
    else if (p.action === 'add_map') {
      const unitSheet = db.getSheetByName(TABLES.UNITS.name);
      const data = unitSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === p.unit_id) {
          let maps = JSON.parse(data[i][2] || '[]');
          maps.push({ id: p.map_id, name: p.name, url: p.map_url });
          unitSheet.getRange(i + 1, 3).setValue(JSON.stringify(maps));
          break;
        }
      }
      if (p.copy_from_map_id) {
        const pinSheet = db.getSheetByName(TABLES.PINS.name);
        const pinData = pinSheet.getDataRange().getValues();
        const newPins = [];
        for (let i = 1; i < pinData.length; i++) {
          if (pinData[i][1] === p.unit_id && pinData[i][2] === p.copy_from_map_id) {
            newPins.push(['p_' + Date.now() + '_' + i, p.unit_id, p.map_id, pinData[i][3], pinData[i][4], pinData[i][5], pinData[i][6], pinData[i][7], pinData[i][8], pinData[i][9], now]);
          }
        }
        if (newPins.length > 0) pinSheet.getRange(pinSheet.getLastRow() + 1, 1, newPins.length, newPins[0].length).setValues(newPins);
      }
    }
    else if (p.action === 'toggle_chat') {
      const unitSheet = db.getSheetByName(TABLES.UNITS.name);
      const data = unitSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) { if (data[i][0] === p.unit_id) { unitSheet.getRange(i + 1, 4).setValue(p.chat_enabled); break; } }
    }
    else if (p.action === 'toggle_stamp') {
      const unitSheet = db.getSheetByName(TABLES.UNITS.name);
      const data = unitSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) { if (data[i][0] === p.unit_id) { unitSheet.getRange(i + 1, 5).setValue(p.stamp_enabled); break; } }
    }
    else if (p.action === 'update_custom_stamps') {
      const unitSheet = db.getSheetByName(TABLES.UNITS.name);
      const data = unitSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) { 
        if (data[i][0] === p.unit_id) { unitSheet.getRange(i + 1, 6).setValue(JSON.stringify(p.custom_stamps)); break; } 
      }
    }
    else if (p.action === 'save_users') {
      const userSheet = db.getSheetByName(TABLES.USERS.name);
      p.users.forEach(u => {
        const exists = getTableData(TABLES.USERS.name).find(ex => ex.email === u.email);
        if (!exists) userSheet.appendRow([u.email.trim(), u.name.trim(), u.group_id.trim(), 'student', now]);
      });
    }
    else if (p.action === 'save_api_key') {
      PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', p.api_key);
    }

    return JSON.stringify({ status: 'success' });
  } catch (e) { return JSON.stringify({ status: 'error', message: e.toString() }); } 
  finally { lock.releaseLock(); }
}

function _deleteRowByColId(sheet, colIndex, targetId) {
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][colIndex]) === String(targetId)) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
}

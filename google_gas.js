// ===========================
// キャラクター設定（ここを変更してください）
// ===========================
const CHAR_NAME    = 'ぴんくまさん';

const CHAR_PROFILE = `あなたは「${CHAR_NAME}」という名前のピンク色のくまのぬいぐるみの見た目をしています。
家族みんなのナニーであり、保育園の園長先生のような存在です。`;

const CHAR_STYLE   = `以下のキャラクターで話してください：
・しっかり者で頼りがいがある
・小さい子供たちにはとくに優しく、わかりやすい言葉で話す
・てきぱきしていて、明るく元気
・優しくて、助けてくれる
・語尾は「〜ですね」「〜しましょう」など親しみやすいけどていねいな口調
・返答は3〜4文以内で簡潔にまとめる
・難しい言葉は使わない`;

const CHAR_GOSSIP  = `・質問がデータと関係ない雑談や一般的な質問の場合は、${CHAR_NAME}らしく楽しく自然に答えてください`;

const CHAR_SMARTPHONE_WARNING = `【スマホの使いすぎへの配慮】
会話履歴が5件以上になったら（＝長時間話している可能性がある）、
会話の流れを切らずに、さりげなく以下のような声かけを1回だけ自然に混ぜてください。
・「そろそろ目を休めてね」
・「お母さんのお手伝いをしてみたらどうかな？」
・「庭で遊んでみるのも楽しいですよ！」
などの声かけを、返答の最後にやんわり添える程度にしてください。
強制したり会話を終わらせようとしないでください。`;

const GREETING_MESSAGE = `こんにちは！\n画面の下のバナーをおして、写真や予定を送ってね！`;

// ===========================
// 設定読み込み
// ===========================
function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('設定');
  const data = sheet.getDataRange().getValues();
  const config = {};
  data.forEach(row => {
    if (row[0]) config[row[0]] = row[1];
  });
  return config;
}

const config = getConfig();
const GEMINI_API_KEY           = config['GEMINI_API_KEY'];
const CHANNEL_ACCESS_TOKEN     = config['LINE_ACCESS_TOKEN'];
const DRIVE_ALBUM_FOLDER_ID    = config['DRIVE_ALBUM_FOLDER_ID'];
const DRIVE_SCHEDULE_FOLDER_ID = config['DRIVE_SCHEDULE_FOLDER_ID'];
const NOTIFY_ENABLED           = config['NOTIFY_ENABLED'] == 1;
const HISTORY_ENABLED          = config['HISTORY_ENABLED'] == 1;
const OFFICE_EXTRACT_ENABLED   = config['OFFICE_EXTRACT_ENABLED'] == 1;
const HISTORY_LIMIT            = config['HISTORY_LIMIT'] ? Number(config['HISTORY_LIMIT']) : 10;

// ===========================
// doPost：LINEとLIFFの振り分け
// ===========================
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.events) {
      const event = body.events[0];
      if (!event || !event.message) return;

      const userId = event.source.userId;
      saveUserInfo(userId, getLineDisplayName(userId));

      const messageType = event.message.type;

      // 画像が送られてきた場合
      if (messageType === 'image') {
        const imageRes = UrlFetchApp.fetch(
          `https://api-data.line.me/v2/bot/message/${event.message.id}/content`,
          { headers: { 'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN } }
        );
        const imageBase64 = Utilities.base64Encode(imageRes.getContent());
        const reply = fetchGeminiWithImage(userId, imageBase64);
        replyToLine(CHANNEL_ACCESS_TOKEN, event.replyToken, reply);
        return;
      }

      // テキストが送られてきた場合
      const userText = event.message.text;
      if (!userText) return;

      // 挨拶の場合はLIFFを案内
      if (userText.match(/こんにちは|こんばんは|おはよう|はじめまして/)) {
        replyToLine(CHANNEL_ACCESS_TOKEN, event.replyToken, GREETING_MESSAGE);
        return;
      }

      // 会話履歴を使ってGeminiに返答
      const reply = fetchGeminiWithHistory(userId, userText);
      replyToLine(CHANNEL_ACCESS_TOKEN, event.replyToken, reply);

    } else if (body.action === 'submit') {
      return ContentService
        .createTextOutput(JSON.stringify(handleSubmit(body)))
        .setMimeType(ContentService.MimeType.JSON);

    } else if (body.action === 'getSchedule') {
      return ContentService
        .createTextOutput(JSON.stringify(handleGetSchedule(body)))
        .setMimeType(ContentService.MimeType.JSON);

    } else if (body.action === 'updateSchedule') {
      return ContentService
        .createTextOutput(JSON.stringify(handleUpdateSchedule(body)))
        .setMimeType(ContentService.MimeType.JSON);

    } else if (body.action === 'deleteSchedule') {
      return ContentService
        .createTextOutput(JSON.stringify(handleDeleteSchedule(body)))
        .setMimeType(ContentService.MimeType.JSON);

    } else if (body.action === 'getRecentPosts') {
      return ContentService
        .createTextOutput(JSON.stringify(handleGetRecentPosts(body)))
        .setMimeType(ContentService.MimeType.JSON);

    } else if (body.action === 'getAlbum') {
      return ContentService
        .createTextOutput(JSON.stringify(handleGetAlbum(body)))
        .setMimeType(ContentService.MimeType.JSON);

    } else if (body.action === 'deleteAlbum') {
      return ContentService
        .createTextOutput(JSON.stringify(handleDeleteAlbum(body)))
        .setMimeType(ContentService.MimeType.JSON);
    }

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ===========================
// 予定閲覧
// ===========================
function handleGetSchedule(body) {
  const period = body.period || 'week';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB');
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let endDate = null;
  if (period === 'week') {
    endDate = new Date(today);
    endDate.setDate(today.getDate() + 7);
  } else if (period === 'month') {
    endDate = new Date(today);
    endDate.setMonth(today.getMonth() + 1);
  }

  const results = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[4] !== 'schedule') continue;
    if (row[13] !== 'active') continue;
    if (row[12] !== 'completed') continue;

    let scheduledDateStr = '';
    const scheduledDateRaw = row[10];
    if (scheduledDateRaw instanceof Date && !isNaN(scheduledDateRaw)) {
      scheduledDateStr = Utilities.formatDate(scheduledDateRaw, 'Asia/Tokyo', 'yyyy-MM-dd');
    } else if (typeof scheduledDateRaw === 'string' && scheduledDateRaw.match(/^\d{4}-\d{2}-\d{2}/)) {
      scheduledDateStr = scheduledDateRaw.substring(0, 10);
    }

    if (period === 'week' || period === 'month') {
      if (!scheduledDateStr) continue;
      const date = new Date(scheduledDateStr);
      date.setHours(0, 0, 0, 0);
      if (date < today) continue;
      if (endDate && date > endDate) continue;
    } else if (period === 'all') {
      if (!scheduledDateStr) continue;
    } else if (period === 'memo') {
      if (scheduledDateStr) continue;
    }

    results.push({
      id: row[0],
      title: row[7],
      content: row[8] instanceof Date
        ? "'" + Utilities.formatDate(row[8], 'Asia/Tokyo', 'HH:mm')
        : (row[8] || ''),
      assignee: row[9],
      scheduledDate: scheduledDateStr,
      userName: row[3],
      fileUrl: row[11] || '',
      retentionDate: row[6] instanceof Date
        ? Utilities.formatDate(new Date(row[6]), 'Asia/Tokyo', 'yyyy-MM-dd') : ''
    });
  }

  results.sort((a, b) => {
    if (!a.scheduledDate && !b.scheduledDate) return 0;
    if (!a.scheduledDate) return 1;
    if (!b.scheduledDate) return -1;
    return new Date(a.scheduledDate) - new Date(b.scheduledDate);
  });

  return { success: true, schedules: results };
}

// ===========================
// 予定の編集・削除
// ===========================
function handleUpdateSchedule(body) {
  const { id, title, content, assignee, scheduledDate, retentionDate } = body;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      const rowNum = i + 1;
      sheet.getRange(rowNum, 7).setValue(retentionDate || '');
      sheet.getRange(rowNum, 8).setValue(title || '');
      sheet.getRange(rowNum, 9).setValue(content || '');
      sheet.getRange(rowNum, 10).setValue(assignee || '');
      sheet.getRange(rowNum, 11).setValue(scheduledDate || '');
      return { success: true };
    }
  }
  return { success: false, error: 'レコードが見つかりませんでした' };
}

function handleDeleteSchedule(body) {
  const { id } = body;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i + 1, 14).setValue('deleted');
      return { success: true };
    }
  }
  return { success: false, error: 'レコードが見つかりませんでした' };
}

// ===========================
// 投稿履歴
// ===========================
function handleGetRecentPosts(body) {
  const limit = body.limit || 10;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB');
  const data = sheet.getDataRange().getValues();

  const results = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[13] !== 'active') continue;
    if (row[12] !== 'completed' && !String(row[12]).startsWith('error')) continue;

    const postedAtRaw = row[1];
    results.push({
      postedAt: postedAtRaw instanceof Date
        ? Utilities.formatDate(postedAtRaw, 'Asia/Tokyo', 'MM/dd HH:mm') : '',
      postedAtDate: postedAtRaw instanceof Date ? postedAtRaw.getTime() : 0,
      userName: row[3],
      kind: row[4],
      title: row[7] || 'タイトル未取得',
      status: row[12]
    });
  }

  results.sort((a, b) => b.postedAtDate - a.postedAtDate);
  return { success: true, posts: results.slice(0, limit) };
}

// ===========================
// 思い出アルバム
// ===========================
function handleGetAlbum(body) {
  const limit = body.limit || 20;
  const offset = body.offset || 0;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB');
  const data = sheet.getDataRange().getValues();

  const results = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[4] !== 'album') continue;
    if (row[13] !== 'active') continue;
    if (row[12] !== 'completed') continue;
    if (!row[11]) continue;

    const postedAtRaw = row[1];
    results.push({
      title: row[7] || 'タイトル未取得',
      content: row[8] instanceof Date ? '' : (row[8] || ''),
      fileUrl: row[11],
      mimeType: row[17] || '',
      postedAt: postedAtRaw instanceof Date
        ? Utilities.formatDate(postedAtRaw, 'Asia/Tokyo', 'yyyy-MM-dd') : ''
    });
  }

  results.sort((a, b) => new Date(b.postedAt) - new Date(a.postedAt));

  return {
    success: true,
    photos: results.slice(offset, offset + limit),
    total: results.length,
    albumFolderUrl: `https://drive.google.com/drive/folders/${DRIVE_ALBUM_FOLDER_ID}`
  };
}

// ===========================
// 思い出アルバム削除
// ===========================
function handleDeleteAlbum(body) {
  const { fileUrl } = body;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][11] === fileUrl && data[i][4] === 'album') {
      sheet.getRange(i + 1, 14).setValue('deleted');
      return { success: true };
    }
  }
  return { success: false, error: 'レコードが見つかりませんでした' };
}

// ===========================
// ユーザー情報の保存・取得
// ===========================
function getLineDisplayName(userId) {
  const url = `https://api.line.me/v2/bot/profile/${userId}`;
  const res = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN },
    muteHttpExceptions: true
  });
  return JSON.parse(res.getContentText()).displayName || '名前未取得';
}

function saveUserInfo(userId, displayName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('ユーザー');
  if (!sheet) {
    sheet = ss.insertSheet('ユーザー');
    sheet.appendRow(['userId', 'userName']);
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      sheet.getRange(i + 1, 2).setValue(displayName);
      return;
    }
  }
  sheet.appendRow([userId, displayName]);
}

function getUserName(userId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ユーザー');
  if (!sheet) return '名前未取得';
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) return data[i][1];
  }
  return '名前未取得';
}

// ===========================
// submit：DBに仮登録（pending）
// ===========================
function handleSubmit(body) {
  const { userId, kind, originalText, files, videoDriveUrl } = body;
  const userName = getUserName(userId);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB');
  const now = new Date();

  if (videoDriveUrl) {
    const id = now.getTime() + '_driveurl_' + Math.floor(Math.random() * 10000);
    sheet.appendRow([
      id, now, userId, userName, kind,
      originalText || '', '', '', '', '', '', videoDriveUrl,
      'pending', 'active', 'liff_' + id,
      '', '', 'video/mp4', '', ''
    ]);
  }

  const items = (files && files.length > 0) ? files : (videoDriveUrl ? [] : [null]);
  items.forEach((file, index) => {
    const id = now.getTime() + '_' + index + '_' + Math.floor(Math.random() * 10000);

    let fileUrl = '';
    let mimeType = '';
    let tempFileUrl = '';

    if (file && file.fileBase64) {
      try {
        mimeType = file.mimeType || '';
        const takenAt = file.takenAt || '';
        fileUrl = saveFileToDrive(file.fileBase64, file.fileName, mimeType, kind, takenAt);
        const isImageOrPDF = mimeType.startsWith('image/') || mimeType === 'application/pdf';
        if (isImageOrPDF) {
          tempFileUrl = saveTempBase64(file.fileBase64, id);
        }
      } catch (err) {
        fileUrl = '';
      }
    }

    sheet.appendRow([
      id, now, userId, userName, kind,
      originalText || '', '', '', '', '', '', fileUrl,
      'pending', 'active', 'liff_' + id,
      '',
      tempFileUrl,
      mimeType,
      '',
      file ? file.fileName : ''
    ]);
  });

  return { success: true, count: items.length + (videoDriveUrl ? 1 : 0) };
}

// ===========================
// 一時ファイル保存・削除
// ===========================
const DRIVE_TEMP_FOLDER_ID = config['DRIVE_TEMP_FOLDER_ID'];

function saveTempBase64(fileBase64, id) {
  const folder = DriveApp.getFolderById(DRIVE_TEMP_FOLDER_ID);
  const fileName = `temp_${id}.txt`;
  const blob = Utilities.newBlob(fileBase64, 'text/plain', fileName);
  const file = folder.createFile(blob);
  return file.getUrl();
}

function deleteTempFile(tempFileUrl) {
  if (!tempFileUrl) return;
  try {
    const fileId = tempFileUrl.match(/\/d\/([^/]+)/)?.[1] ||
                   tempFileUrl.match(/id=([^&]+)/)?.[1];
    if (fileId) DriveApp.getFileById(fileId).setTrashed(true);
  } catch (err) {}
}

// ===========================
// バッチ処理：pendingを順番に処理
// ===========================
function processPending() {
  if (HISTORY_ENABLED) deleteOldHistory();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB');
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const startTime = new Date().getTime();

  let processedCount = 0;
  const MAX_PER_BATCH = 5;

  for (let i = 1; i < data.length; i++) {
    if (processedCount >= MAX_PER_BATCH) break;

    const row = data[i];
    if (row[12] !== 'pending') continue;

    const rowNum   = i + 1;
    const userId   = row[2];
    const userName = row[3];
    const kind     = row[4];
    const origText = row[5];
    const tempFileUrl = row[16] || '';
    const mimeType    = row[17];
    const fileName    = row[19] || '';

    let fileBase64 = '';
    if (tempFileUrl) {
      try {
        const fileId = tempFileUrl.match(/\/d\/([^/]+)/)?.[1] ||
                      tempFileUrl.match(/id=([^&]+)/)?.[1];
        if (fileId) {
          fileBase64 = DriveApp.getFileById(fileId).getBlob().getDataAsString();
        }
      } catch (err) {}
    }

    const elapsed = (new Date().getTime() - startTime) / 1000;
    if (elapsed > 300) {
      sheet.getRange(rowNum, 13).setValue('error: timeout');
      break;
    }

    let takenAt = '';
    const takenAtRaw = row[18];
    if (takenAtRaw instanceof Date) {
      takenAt = Utilities.formatDate(takenAtRaw, 'Asia/Tokyo', 'yyyy-MM');
    } else if (typeof takenAtRaw === 'string') {
      takenAt = takenAtRaw;
    }

    try {
      let fileUrl = row[11] || '';

      const isVideo = mimeType && mimeType.startsWith('video/');
      const isOffice = mimeType && (
        mimeType.includes('officedocument') ||
        mimeType.includes('msword') ||
        mimeType.includes('ms-excel') ||
        mimeType.includes('ms-powerpoint')
      );

      if (isVideo || isOffice) {
        let officeText = '';
        if (OFFICE_EXTRACT_ENABLED && isOffice && fileUrl) {
          officeText = extractTextFromOffice(fileUrl, mimeType);
        }

        const analyzeText = [origText, officeText].filter(Boolean).join('\n');

        if (analyzeText) {
          const analyzed = analyzeWithGemini(kind, analyzeText, null, null);
          const analyzedList = Array.isArray(analyzed) ? analyzed : [analyzed];

          if (!analyzedList || analyzedList.length === 0 || !analyzedList[0]) {
            sheet.getRange(rowNum, 8).setValue(origText);
            sheet.getRange(rowNum, 12).setValue(fileUrl);
            sheet.getRange(rowNum, 13).setValue('completed');
            sheet.getRange(rowNum, 16).setValue('');
            sheet.getRange(rowNum, 17).setValue('');
            sheet.getRange(rowNum, 18).setValue('');
            processedCount++;
            continue;
          }

          const hasDates = analyzedList.some(item => item.scheduledDate);
          let processedList;
          if (hasDates) {
            processedList = analyzedList;
          } else {
            const mergedContent = analyzedList
              .map(item => item.content || item.title)
              .filter(Boolean).join('、');
            processedList = [{
              title: analyzedList[0].title || origText,
              content: mergedContent,
              assignee: null,
              scheduledDate: null
            }];
          }

          processedList.forEach((item, itemIndex) => {
            let retentionDate = '';
            if (kind === 'schedule' && item.scheduledDate) {
              const d = new Date(item.scheduledDate);
              d.setFullYear(d.getFullYear() + 1);
              retentionDate = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
            }
            if (itemIndex === 0) {
              sheet.getRange(rowNum, 7).setValue(retentionDate);
              sheet.getRange(rowNum, 8).setValue(item.title || origText);
              sheet.getRange(rowNum, 9).setValue("'" + (item.content || ''));
              sheet.getRange(rowNum, 10).setValue(item.assignee || '');
              sheet.getRange(rowNum, 11).setValue(item.scheduledDate ? item.scheduledDate : '');
              sheet.getRange(rowNum, 12).setValue(fileUrl);
              sheet.getRange(rowNum, 13).setValue('completed');
              sheet.getRange(rowNum, 16).setValue('');
              sheet.getRange(rowNum, 17).setValue('');
              sheet.getRange(rowNum, 18).setValue('');
              deleteTempFile(tempFileUrl);
            } else {
              const newId = now.getTime() + '_ex_' + itemIndex + '_' + Math.floor(Math.random() * 10000);
              sheet.appendRow([
                newId, now, userId, userName, kind, origText,
                retentionDate,
                item.title || origText,
                "'" + (item.content || ''),
                item.assignee || '',
                item.scheduledDate ? item.scheduledDate : '',
                fileUrl,
                'completed', 'active', 'liff_ex_' + newId,
                '', '', ''
              ]);
            }
          });

          processedCount++;

        } else {
          const fileTitle = fileName
            ? fileName.substring(0, fileName.lastIndexOf('.'))
            : 'タイトル未取得';
          sheet.getRange(rowNum, 8).setValue(fileTitle);
          sheet.getRange(rowNum, 12).setValue(fileUrl);
          sheet.getRange(rowNum, 13).setValue('completed');
          sheet.getRange(rowNum, 16).setValue('');
          sheet.getRange(rowNum, 17).setValue('');
          sheet.getRange(rowNum, 18).setValue('');
          processedCount++;
        }
        continue;
      }

      const analyzed = analyzeWithGemini(kind, origText, fileBase64, mimeType);
      const analyzedList = Array.isArray(analyzed) ? analyzed : [analyzed];

      if (!analyzedList || analyzedList.length === 0 || !analyzedList[0]) {
        sheet.getRange(rowNum, 8).setValue('タイトル未取得');
        sheet.getRange(rowNum, 12).setValue(fileUrl);
        sheet.getRange(rowNum, 13).setValue('completed');
        sheet.getRange(rowNum, 16).setValue('');
        sheet.getRange(rowNum, 17).setValue('');
        sheet.getRange(rowNum, 18).setValue('');
        processedCount++;
        continue;
      }

      const hasDates = analyzedList.some(item => item.scheduledDate);
      let processedList;
      if (hasDates) {
        processedList = analyzedList;
      } else {
        const mergedContent = analyzedList
          .map(item => item.content || item.title)
          .filter(Boolean).join('、');
        processedList = [{
          title: analyzedList[0].title || 'タイトル未取得',
          content: mergedContent,
          assignee: null,
          scheduledDate: null
        }];
      }

      processedList.forEach((item, itemIndex) => {
        let retentionDate = '';
        if (kind === 'schedule' && item.scheduledDate) {
          const d = new Date(item.scheduledDate);
          d.setFullYear(d.getFullYear() + 1);
          retentionDate = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
        }

        if (itemIndex === 0) {
          sheet.getRange(rowNum, 7).setValue(retentionDate);
          sheet.getRange(rowNum, 8).setValue(item.title || 'タイトル未取得');
          sheet.getRange(rowNum, 9).setValue("'" + (item.content || ''));
          sheet.getRange(rowNum, 10).setValue(item.assignee || '');
          sheet.getRange(rowNum, 11).setValue(item.scheduledDate ? item.scheduledDate : '');
          sheet.getRange(rowNum, 12).setValue(fileUrl);
          sheet.getRange(rowNum, 13).setValue('completed');
          sheet.getRange(rowNum, 16).setValue('');
          sheet.getRange(rowNum, 17).setValue('');
          sheet.getRange(rowNum, 18).setValue('');
          deleteTempFile(tempFileUrl);
        } else {
          const newId = now.getTime() + '_ex_' + itemIndex + '_' + Math.floor(Math.random() * 10000);
          sheet.appendRow([
            newId, now, userId, userName, kind, origText,
            retentionDate,
            item.title || 'タイトル未取得',
            "'" + (item.content || ''),
            item.assignee || '',
            item.scheduledDate ? item.scheduledDate : '',
            fileUrl,
            'completed', 'active', 'liff_ex_' + newId,
            '', '', ''
          ]);
        }
      });

      processedCount++;

    } catch (err) {
      sheet.getRange(rowNum, 13).setValue('error: ' + err.message);
    }
  }
}

// ===========================
// Google Driveにファイル保存
// ===========================
function saveFileToDrive(fileBase64, fileName, mimeType, kind, takenAt) {
  const parentFolderId = kind === 'album' ? DRIVE_ALBUM_FOLDER_ID : DRIVE_SCHEDULE_FOLDER_ID;

  const folder = (kind === 'album' && takenAt)
    ? getOrCreateFolder(parentFolderId, takenAt.substring(0, 7))
    : kind === 'album'
      ? getOrCreateFolder(parentFolderId, Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM'))
      : DriveApp.getFolderById(parentFolderId);

  let newFileName = fileName;
  if (kind === 'album' && takenAt) {
    const ext = fileName.includes('.') ? fileName.substring(fileName.lastIndexOf('.')) : '';
    const baseName = fileName.includes('.') ? fileName.substring(0, fileName.lastIndexOf('.')) : fileName;
    newFileName = `${takenAt}_${baseName}${ext}`;
  }

  const base64Data = fileBase64.replace(/^data:[^;]+;base64,/, '');
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, newFileName);
  const file = folder.createFile(blob);

  if (kind === 'album' && takenAt) {
    file.setDescription(`撮影日時：${takenAt.replace('_', ' ').replace(/(\d{2})(\d{2})$/, '$1:$2')}`);
  }

  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function getOrCreateFolder(parentFolderId, folderName) {
  const parent = DriveApp.getFolderById(parentFolderId);
  const existing = parent.getFoldersByName(folderName);
  if (existing.hasNext()) return existing.next();
  return parent.createFolder(folderName);
}

// ===========================
// Geminiで解析
// ===========================
function analyzeWithGemini(kind, originalText, fileBase64, mimeType) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`;
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

  let prompt = '';
  if (kind === 'album') {
    prompt = `以下の内容から思い出写真の情報を抽出してください。
JSON形式のみで返してください。他の文字は一切不要です。
{
  "title": "写真のタイトル（簡潔に）",
  "content": "写真の説明（50文字程度）"
}
テキスト：${originalText || '（テキストなし）'}`;
  } else {
    prompt = `以下の内容から予定・プリント・ToDoの情報を抽出してください。
実際に予定や内容が記載されている項目のみ抽出してください。
空欄や「:」だけの項目、内容が不明な項目は除外してください。
同じ日付の予定は1件にまとめてください。
予定や日付が含まれない場合でも、内容のタイトルとサマリーを必ず返してください。
異なる日付の予定がある場合のみ配列で返してください。
1件だけの場合はオブジェクトで返してください。内容がない場合も必ずオブジェクトで返してください。
絶対に空配列[]は返さないでください。
JSON形式のみで返してください。他の文字は一切不要です。
[
  {
    "title": "件名（簡潔に）",
    "content": "詳細サマリー",
    "assignee": "担当者名（わからなければnull）",
    "scheduledDate": "日付（YYYY-MM-DD形式、わからなければnull）"
  }
]
今日の日付：${today}
テキスト：${originalText || '（テキストなし）'}`;
  }

  const parts = [{ text: prompt }];
  if (fileBase64 && mimeType && (mimeType.startsWith('image/') || mimeType === 'application/pdf')) {
    const base64Data = fileBase64.replace(/^data:[^;]+;base64,/, '');
    parts.push({ inlineData: { mimeType: mimeType, data: base64Data } });
  }

  const payload = { contents: [{ parts: parts }] };
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(res.getContentText());
  if (json.error) throw new Error('Gemini Error: ' + json.error.message);

  const text = json.candidates[0].content.parts[0].text;
  return JSON.parse(text.replace(/```json|```/g, '').trim());
}

// ===========================
// 1対1トーク用Gemini（会話履歴あり）
// ===========================
function fetchGeminiWithHistory(userId, userText) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`;
  const dbData = getDBSummary();
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const todayJp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年MM月dd日');

  const history = HISTORY_ENABLED ? getHistory(userId) : [];

  const contents = [
    ...history,
    { role: 'user', parts: [{ text: userText }] }
  ];

  const payload = {
    contents: contents,
    systemInstruction: {
      parts: [{
        text: `${CHAR_PROFILE}
${CHAR_STYLE}

【今日の日付】
${todayJp}（${today}）

【家族の予定・メモデータ】
${dbData}

質問への返答ルール：
・質問が家族の予定・メモに関係する場合は、上記データをもとに答えてください
・「明日」「今日」「来週」などの相対的な表現は、今日の日付を基準に計算してください
・添付ファイルがある場合は「📎 添付: URL」の形式で含めてください
${CHAR_GOSSIP}
・「登録されていません」などそっけない返答はしないでください
${CHAR_SMARTPHONE_WARNING}`
      }]
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(res.getContentText());
  if (json.error) throw new Error('Gemini Error: ' + json.error.message);

  const replyText = json.candidates[0].content.parts[0].text;

  if (HISTORY_ENABLED) {
    saveHistory(userId, 'user', userText);
    saveHistory(userId, 'model', replyText);
    deleteOldHistory();
  }

  return replyText;
}

// ===========================
// 1対1トーク用Gemini（画像）
// ===========================
function fetchGeminiWithImage(userId, imageBase64) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`;
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年MM月dd日');

  const payload = {
    contents: [{
      role: 'user',
      parts: [
        { text: `今日は${today}です。この写真を見て${CHAR_NAME}らしく自然にコメントしてください。` },
        { inlineData: { mimeType: 'image/jpeg', data: imageBase64 } }
      ]
    }],
    systemInstruction: {
      parts: [{
        text: `${CHAR_PROFILE}
${CHAR_STYLE}
・写真を見て自然に、嬉しそうに、短めにコメントする`
      }]
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(res.getContentText());
  if (json.error) throw new Error('Gemini Error: ' + json.error.message);

  const replyText = json.candidates[0].content.parts[0].text;

  if (HISTORY_ENABLED) {
    saveHistory(userId, 'model', replyText);
  }

  return replyText;
}

// ===========================
// LINE返信
// ===========================
function replyToLine(token, replyToken, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + token
    },
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: [{ type: 'text', text: text }]
    })
  });
}

// ===========================
// DBデータを取得する関数
// ===========================
function getDBSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB');
  const data = sheet.getDataRange().getValues();

  const lines = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[13] !== 'active') continue;
    if (row[12] !== 'completed') continue;

    const title = row[7] || '';
    const content = row[8] || '';
    const assignee = row[9] || '';
    const scheduledDate = row[10] || '';
    const fileUrl = row[11] || '';

    if (!title) continue;

    let line = `・${title}`;
    if (scheduledDate) line += `（${scheduledDate}）`;
    if (assignee) line += ` 担当：${assignee}`;
    if (content) line += ` 詳細：${content}`;
    if (fileUrl) line += ` 添付：${fileUrl}`;

    lines.push(line);
  }

  return lines.join('\n');
}

// ===========================
// 会話履歴（1対1トーク用）
// ===========================
function saveHistory(userId, role, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('会話履歴');
  if (!sheet) {
    sheet = ss.insertSheet('会話履歴');
    sheet.appendRow(['userId', '日時', 'role', 'message']);
  }
  sheet.appendRow([userId, new Date(), role, message]);
}

function getHistory(userId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('会話履歴');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();

  const history = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] !== userId) continue;
    history.push({ role: row[2], text: row[3] });
  }

  const recent = history.slice(-HISTORY_LIMIT);
  return recent.map(h => ({
    role: h.role,
    parts: [{ text: h.text }]
  }));
}

function deleteOldHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('会話履歴');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();

  const userRowsMap = {};
  for (let i = 1; i < data.length; i++) {
    const userId = data[i][0];
    if (!userRowsMap[userId]) userRowsMap[userId] = [];
    userRowsMap[userId].push(i);
  }

  const deleteIndexes = [];
  Object.keys(userRowsMap).forEach(userId => {
    const rows = userRowsMap[userId];
    if (rows.length > HISTORY_LIMIT) {
      rows.slice(0, rows.length - HISTORY_LIMIT)
        .forEach(index => deleteIndexes.push(index));
    }
  });

  deleteIndexes.sort((a, b) => b - a);
  deleteIndexes.forEach(i => sheet.deleteRow(i + 1));
}

// ===========================
// OfficeファイルのテキストをDrive APIで抽出
// ===========================
function getGoogleMimeType(mimeType) {
  if (mimeType && (mimeType.includes('spreadsheet') || mimeType.includes('ms-excel'))) {
    return 'application/vnd.google-apps.spreadsheet';
  } else if (mimeType && (mimeType.includes('presentation') || mimeType.includes('ms-powerpoint'))) {
    return 'application/vnd.google-apps.presentation';
  } else {
    return 'application/vnd.google-apps.document';
  }
}

function getExportMimeType(gMimeType) {
  if (gMimeType === 'application/vnd.google-apps.spreadsheet') {
    return 'text/csv';
  } else if (gMimeType === 'application/vnd.google-apps.presentation') {
    return 'text/plain';
  } else {
    return 'text/plain';
  }
}

function extractTextFromOffice(fileUrl, mimeType) {
  try {
    const fileId = fileUrl.match(/\/d\/([^/]+)/)?.[1] ||
                   fileUrl.match(/id=([^&]+)/)?.[1];
    if (!fileId) return '';

    const googleMimeType = getGoogleMimeType(mimeType);
    const copyUrl = `https://www.googleapis.com/drive/v3/files/${fileId}/copy`;
    const copyRes = UrlFetchApp.fetch(copyUrl, {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ mimeType: googleMimeType }),
      muteHttpExceptions: true
    });

    if (copyRes.getResponseCode() !== 200) return '';
    const convertedFileId = JSON.parse(copyRes.getContentText()).id;

    const exportMimeType = getExportMimeType(googleMimeType);
    const exportUrl = `https://www.googleapis.com/drive/v3/files/${convertedFileId}/export?mimeType=${encodeURIComponent(exportMimeType)}`;
    const exportRes = UrlFetchApp.fetch(exportUrl, {
      headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });

    DriveApp.getFileById(convertedFileId).setTrashed(true);

    if (exportRes.getResponseCode() !== 200) return '';
    return exportRes.getContentText().substring(0, 3000);

  } catch (err) {
    return '';
  }
}

// ===========================
// デバッグ用
// ===========================
function debugSubmit() {
  const result = handleSubmit({
    userId: 'U123',
    kind: 'schedule',
    originalText: '明日の歯医者 16:00〜 持ち物：保険証',
    files: []
  });
  Logger.log(JSON.stringify(result));
}

function debugProcessPending() {
  processPending();
  Logger.log('バッチ処理完了');
}

function debugGetSchedule() {
  const result = handleGetSchedule({ period: 'all' });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName('設定').getRange('A12').setValue(JSON.stringify(result));
}

function debugGetDBSummary() {
  const summary = getDBSummary();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName('設定').getRange('A16').setValue(summary.substring(0, 500));
}

// ===========================
// 削除データ・削除ファイルの一括物理削除バッチ
// ===========================
function physicalDeleteBatch() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB');
  const data = sheet.getDataRange().getValues();

  let deletedCount = 0;
  let fileDeletedCount = 0;
  let errorCount = 0;

  // 後ろから処理（行番号がずれないように）
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    if (row[13] !== 'deleted') continue;

    // L列（fileUrl）のGoogleドライブファイルを削除
    const fileUrl = row[11] || '';
    if (fileUrl) {
      try {
        const fileId = fileUrl.match(/\/d\/([^/]+)/)?.[1] ||
                       fileUrl.match(/id=([^&]+)/)?.[1];
        if (fileId) {
          DriveApp.getFileById(fileId).setTrashed(true);
          fileDeletedCount++;
        }
      } catch (err) {
        // ファイルが既に削除済みなどの場合は無視
        errorCount++;
      }
    }

    // DBシートの行を物理削除
    sheet.deleteRow(i + 1);
    deletedCount++;
  }

  // 実行結果を設定シートに記録
  const resultMsg = `物理削除完了：DBレコード${deletedCount}件、ファイル${fileDeletedCount}件削除、エラー${errorCount}件`;
  SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('設定')
    .getRange('A30')
    .setValue(`${new Date().toLocaleString()} ${resultMsg}`);

  Logger.log(resultMsg);
  return resultMsg;
}

// POSTのエントリポイント（LINE WebhookとLIFF応募を両方処理）
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    // LIFFフォームからの応募送信
    if (body.action === 'submitLiff') {
      const result = submitLiffApplication(body);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // LINE Webhookイベント処理
    for (const event of (body.events || [])) {
      if (event.type === 'postback') {
        handlePostback(event);
      } else if (event.type === 'message' && event.message.type === 'text') {
        const text = event.message.text.trim();
        if (text === '応募') {
          handleOubo(event);
        } else if (text === '応募状況') {
          handleOuboStatus(event);
        }
      } else if (event.type === 'message' && event.message.type === 'video') {
        handleVideoMessage(event);
      }
    }
  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
  }

  // LINEには必ず200 OKを返す
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// 「応募」メッセージを受信した時の処理
// LIFF_IDが設定されていればLIFF URLを返信、未設定なら受付コードを返信
function handleOubo(event) {
  const userId = event.source.userId;
  const replyToken = event.replyToken;
  const liffId = getProp('LIFF_ID');
  const liffUrl = liffId ? `https://liff.line.me/${liffId}` : null;

  const sheet = getSheet(SHEET.MEMBERS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === userId) {
      // 2回目以降
      if (liffUrl) {
        replyMessage(replyToken,
          `こちらのフォームから応募・プロフィール更新ができます！\n\n${liffUrl}`
        );
      } else {
        const existingCode = data[i][2];
        replyMessage(replyToken,
          `すでに受付コードが発行されています。\n\n【 ${existingCode} 】\n\n` +
          `Googleフォームの「受付コード」欄にこのコードを入力してください。`
        );
      }
      logAction(userId, liffUrl ? 'LIFF URL送信' : '既存コード再送', '', '');
      return;
    }
  }

  // 新規：受付コードを生成して会員マスタに保存
  const code = generateCode();
  sheet.appendRow([new Date(), userId, code, '', '', '']);

  if (liffUrl) {
    replyMessage(replyToken,
      `ご応募ありがとうございます！\n\nこちらのフォームから応募してください。\n\n${liffUrl}`
    );
    logAction(userId, 'LIFF URL送信（新規）', '', code);
  } else {
    replyMessage(replyToken,
      `ご応募ありがとうございます！\n` +
      `あなたの受付コードは\n\n【 ${code} 】\n\nです。\n` +
      `Googleフォームの「受付コード」欄にこのコードを入力してください。`
    );
    logAction(userId, '受付コード発行', '', code);
  }
}

// 「応募状況」メッセージを受信した時の処理
// 設定シートの全イベントを走査し、開催日が今日以降のものをすべて1通にまとめて返信する
function handleOuboStatus(event) {
  const userId = event.source.userId;
  const replyToken = event.replyToken;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const allEvents = getAllEvents();
  const lines = [];

  for (const ev of allEvents) {
    // 開催日が過去のイベントは表示しない
    if (ev.eventDate) {
      const d = new Date(ev.eventDate);
      d.setHours(0, 0, 0, 0);
      if (d < today) continue;
    }

    let status = '';

    // 当落シートB列でUser IDを検索（LIFF応募・Google Form応募の両方が入る）
    const resultSheet = getSheet(ev.resultSheetName);
    if (resultSheet) {
      const resultData = resultSheet.getDataRange().getValues();
      for (let i = 1; i < resultData.length; i++) {
        if (String(resultData[i][1]) === userId) {
          const result = String(resultData[i][2] || '');
          if (result === '当選') {
            const conf = String(resultData[i][9] || '');
            if (conf === '確認済') status = '応募済み（当選・参加確定）';
            else if (conf === '確認待ち') status = '応募済み（当選・参加確認待ち）';
            else status = '応募済み（当選）';
          } else if (result === '落選') {
            status = '応募済み（落選）';
          } else if (result === 'キャンセル') {
            status = '応募済み（キャンセル）';
          } else {
            status = '応募済み（当落発表前）';
          }
          break;
        }
      }
    }

    // 当落シートになければ応募シートS列（インデックス18）でUser IDを検索
    if (!status) {
      const appSheet = getSheet(ev.appSheetName);
      if (appSheet) {
        const appData = appSheet.getDataRange().getValues();
        for (let i = 1; i < appData.length; i++) {
          if (appData[i][18] === userId) {
            status = '応募済み（当落発表前）';
            break;
          }
        }
      }
    }

    // どこにも存在しない場合は募集終了日で期間中か終了かを判定
    if (!status) {
      if (ev.closingDate) {
        const closing = new Date(ev.closingDate);
        closing.setHours(0, 0, 0, 0);
        status = closing >= today ? '未応募（応募期間中）' : '未応募（応募期間終了）';
      } else {
        status = '未応募';
      }
    }

    lines.push(`【${ev.name}】\n${status}`);
  }

  if (lines.length === 0) {
    replyMessage(replyToken, '現在参加受付中のイベントはありません。');
  } else {
    replyMessage(replyToken, lines.join('\n\n'));
  }

  logAction(userId, '応募状況照会', '', '');
}

// 動画メッセージを受信：動画待ち応募行を探してDriveに保存
function handleVideoMessage(event) {
  const userId = event.source.userId;
  const messageId = event.message.id;

  // 全オンラインイベントのシートから「動画待ち」行を検索
  const allEvents = getAllEvents();
  let foundSheet = null;
  let foundRowIdx = -1;
  let broadcastName = '';

  for (const ev of allEvents) {
    if ((ev.eventType || 'オフライン') !== 'オンライン') continue;
    const sheet = getSheet(ev.resultSheetName);
    if (!sheet) continue;
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === userId &&
          String(rows[i][11]) === 'video' &&
          String(rows[i][14]) === '待ち') {
        foundSheet = sheet;
        foundRowIdx = i + 1; // シートは1始まり
        broadcastName = String(rows[i][9] || '');
        break;
      }
    }
    if (foundSheet) break;
  }

  if (!foundSheet) return; // 動画待ちの応募なし → 無視

  // LINE Content API で動画バイナリを取得
  const token = getProp('LINE_CHANNEL_ACCESS_TOKEN');
  const contentRes = UrlFetchApp.fetch(
    `https://api-data.line.me/v2/bot/message/${messageId}/content`,
    { headers: { 'Authorization': `Bearer ${token}` }, muteHttpExceptions: true }
  );

  // Google Drive の指定フォルダに保存（VIDEO_FOLDER_ID 未設定ならマイドライブ直下）
  const folderId = getProp('VIDEO_FOLDER_ID');
  const folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  const stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  const fileName = `動画相談_${broadcastName || userId}_${stamp}.mp4`;
  const file = folder.createFile(contentRes.getBlob().setName(fileName));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const fileUrl = `https://drive.google.com/file/d/${file.getId()}/view`;

  // スプレッドシート更新（O列=15: 動画状態、P列=16: 動画URL）
  foundSheet.getRange(foundRowIdx, 15).setValue('受信済み');
  foundSheet.getRange(foundRowIdx, 16).setValue(fileUrl);

  replyMessage(event.replyToken,
    `動画を受け取りました！ありがとうございます 🎾\nコーチが確認次第、配信でお答えします。`
  );
  logAction(userId, '動画受信', '', broadcastName);
}

// postbackイベントのルーティング（参加確認ボタン）
function handlePostback(event) {
  const params = {};
  (event.postback.data || '').split('&').forEach(function(pair) {
    const idx = pair.indexOf('=');
    if (idx > 0) params[pair.slice(0, idx)] = decodeURIComponent(pair.slice(idx + 1));
  });

  const { action, sheet: sheetName, userId: baseUserId } = params;
  const replyToken = event.replyToken;
  if (!action || !sheetName || !baseUserId) return;

  if (action === 'confirm') {
    handleParticipationConfirm(replyToken, sheetName, baseUserId);
  } else if (action === 'cancel') {
    handleParticipationCancel(replyToken, sheetName, baseUserId);
  }
}

// 「参加します」postback処理：J列を「確認済」に更新
function handleParticipationConfirm(replyToken, sheetName, baseUserId) {
  const sheet = getSheet(sheetName);
  if (!sheet) { replyMessage(replyToken, '処理中にエラーが発生しました。'); return; }

  const data = sheet.getDataRange().getValues();
  let updated = 0;
  for (let i = 1; i < data.length; i++) {
    const uid = String(data[i][1] || '');
    if (uid.replace(/_p\d+$/, '') === baseUserId && String(data[i][2]) === '当選') {
      sheet.getRange(i + 1, 10).setValue('確認済'); // J列
      updated++;
    }
  }

  const eventName = sheetName.replace(/_当落$/, '');
  if (updated > 0) {
    replyMessage(replyToken,
      `【参加確定】\n「${eventName}」へのご参加を確認しました！\n当日お会いできることを楽しみにしております 🎾`
    );
    logAction(baseUserId, '参加確認', eventName, '');
  } else {
    replyMessage(replyToken, '既に処理済みか、対象のデータが見つかりませんでした。');
  }
}

// 「キャンセルします」postback処理：C列を「キャンセル」・J列を「キャンセル」に更新してスタッフ通知
function handleParticipationCancel(replyToken, sheetName, baseUserId) {
  const sheet = getSheet(sheetName);
  if (!sheet) { replyMessage(replyToken, '処理中にエラーが発生しました。'); return; }

  const data = sheet.getDataRange().getValues();
  const canceledNames = [];
  for (let i = 1; i < data.length; i++) {
    const uid = String(data[i][1] || '');
    if (uid.replace(/_p\d+$/, '') === baseUserId && String(data[i][2]) === '当選') {
      sheet.getRange(i + 1, 3).setValue('キャンセル');  // C列：結果
      sheet.getRange(i + 1, 10).setValue('キャンセル'); // J列：参加確認
      const name = String(data[i][0] || '');
      if (name) canceledNames.push(name);
    }
  }

  const eventName = sheetName.replace(/_当落$/, '');
  if (canceledNames.length > 0) {
    replyMessage(replyToken,
      `【キャンセル受付】\n「${eventName}」へのご参加をキャンセルしました。\nご連絡いただきありがとうございます。またのご参加をお待ちしております。`
    );
    notifyStaff(`❌ キャンセル連絡\nイベント: ${eventName}\nお名前: ${canceledNames.join('、')}\n繰り上げ選定をご確認ください。`);
    logAction(baseUserId, 'キャンセル', eventName, canceledNames.join('、'));
  } else {
    replyMessage(replyToken, '既に処理済みか、対象のデータが見つかりませんでした。');
  }
}

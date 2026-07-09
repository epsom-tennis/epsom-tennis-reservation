// POSTのエントリポイント（LINE WebhookとLIFF応募を両方処理）
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    // LIFFフロントエンドからのイベントログ
    if (body.action === 'logEvent') {
      logAction(body.userId || '', body.actionType || '', body.eventId || '', body.detail || '');
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

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
        const urlMatch = text.match(/https?:\/\/\S+/i);
        if (text === '応募') {
          handleOubo(event);
        } else if (text === '応募状況') {
          handleOuboStatus(event);
        } else if (urlMatch) {
          // ギガファイル便等の動画アップロードURLが文中のどこにあっても拾う（前後に説明文があってもOK）
          handleVideoUrlMessage(event, urlMatch[0]);
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

// 応募状況デバッグ：GASエディタから直接実行して当落シートにSTAFF_USER_IDがあるかチェック
function debugApplicationStatus() {
  const userId = getProp('STAFF_USER_ID');
  Logger.log('確認userId: ' + userId);
  const allEvents = getAllEvents();
  Logger.log('イベント数: ' + allEvents.length);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  for (const ev of allEvents) {
    const sheet = getSheet(ev.resultSheetName);
    if (!sheet) { Logger.log('[' + ev.name + '] シートなし: ' + ev.resultSheetName); continue; }
    const data = sheet.getDataRange().getValues();
    Logger.log('[' + ev.name + '] シート行数=' + data.length + ' eventDate=' + ev.eventDate);
    let found = false;
    for (let i = 0; i < data.length; i++) {
      const cellVal = String(data[i][1] || '');
      Logger.log('  row' + i + ' B列=' + cellVal + ' 一致=' + (cellVal === userId));
      if (cellVal === userId) { found = true; break; }
    }
    Logger.log('[' + ev.name + '] 結果: ' + (found ? '応募済み' : '未発見'));
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
  const offlineLines = [];
  const tournamentLines = [];
  const onlineLines = [];

  for (const ev of allEvents) {
    // 停止中・応募状況非表示・応募開始前・開催日が過去のイベントは表示しない
    if (ev.status === '停止') continue;
    if (ev.ouboStatusHidden) continue;
    if (ev.openingDate && ev.openingDate > today) continue;
    if (ev.eventDate) {
      const d = new Date(ev.eventDate);
      d.setHours(0, 0, 0, 0);
      if (d < today) continue;
    }

    const isOnline = (ev.eventType || 'オフライン') === 'オンライン';
    const isTournament = (ev.eventType || 'オフライン') === '大会';
    let status = '';

    // 当落シートB列でUser IDを検索（LIFF応募・Google Form応募の両方が入る）
    const resultSheet = getSheet(ev.resultSheetName);
    if (resultSheet) {
      const resultData = resultSheet.getDataRange().getValues();
      for (let i = 1; i < resultData.length; i++) {
        if (String(resultData[i][1]) === userId) {
          const result = String(resultData[i][2] || '');
          if (result === '当選') {
            if (isTournament) {
              // 大会は先着順で即確定のため「参加確定」と表示する
              status = '応募済み（参加確定）✅';
            } else {
              const conf = String(resultData[i][9] || '');
              if (conf === '確認済') status = '応募済み（当選・参加確定）';
              else if (conf === '確認待ち') status = '応募済み（当選・参加確認待ち）';
              else status = '応募済み（当選）';
            }
          } else if (result === '落選') {
            status = '応募済み（落選）';
          } else if (result === 'キャンセル') {
            status = '応募済み（キャンセル）';
          } else if (isOnline) {
            // オンラインは当落なし。締め切りなしなら複数回応募可能を案内
            const hasDeadline = !!(ev.closingDate || ev.closingDateTimeAt);
            status = hasDeadline ? '応募済み ✅' : '応募済み ✅\nこちらのイベントは複数回応募可能です！';
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
            if (isOnline) {
              const hasDeadline = !!(ev.closingDate || ev.closingDateTimeAt);
              status = hasDeadline ? '応募済み ✅' : '応募済み ✅\nこちらのイベントは複数回応募可能です！';
            } else {
              status = '応募済み（当落発表前）';
            }
            break;
          }
        }
      }
    }

    // 大会で未応募の場合：定員チェックを行い満員・残り枠少を先行表示する
    if (!status && isTournament && ev.capacity > 0) {
      const currentCount = countTournamentParticipants_(ev.resultSheetName);
      const remaining = ev.capacity - currentCount;
      if (remaining <= 0) {
        status = '満員（応募終了）';
      } else if (remaining === 1) {
        status = '残り1名（1人のみ応募可・ペアでの応募は終了しました）';
      } else if (remaining / ev.capacity <= 0.50) {
        // 残り枠少の場合は枠数は表示せず警告のみ
        if (ev.closingDateTimeAt) {
          const isOpen = ev.closingDateTimeAt >= new Date();
          const fmt = Utilities.formatDate(ev.closingDateTimeAt, 'Asia/Tokyo', 'M月d日 H:mm');
          status = isOpen
            ? `⚠️ 残り枠が少なくなっています（${fmt}まで受付中）`
            : '⚠️ 残り枠が少なくなっています（応募期間終了）';
        } else if (ev.closingDate) {
          const closing = new Date(ev.closingDate);
          closing.setHours(0, 0, 0, 0);
          status = closing >= today
            ? '⚠️ 残り枠が少なくなっています（応募期間中）'
            : '⚠️ 残り枠が少なくなっています（応募期間終了）';
        } else {
          status = '⚠️ 残り枠が少なくなっています';
        }
      }
    }

    // どこにも存在しない場合は募集終了日時で期間中か終了かを判定
    // closingDateTimeAt（日時）があればそちらを優先し時刻も表示する
    // オンラインイベントで締め切りなしの場合は「常時募集中」と表示する
    if (!status) {
      if (ev.closingDateTimeAt) {
        const isOpen = ev.closingDateTimeAt >= new Date();
        const fmt = Utilities.formatDate(ev.closingDateTimeAt, 'Asia/Tokyo', 'M月d日 H:mm');
        status = isOpen ? `未応募（${fmt}まで受付中）` : '未応募（応募期間終了）';
      } else if (ev.closingDate) {
        const closing = new Date(ev.closingDate);
        closing.setHours(0, 0, 0, 0);
        status = closing >= today ? '未応募（応募期間中）' : '未応募（応募期間終了）';
      } else {
        status = isOnline ? '常時募集中' : '未応募';
      }
    }

    // オフラインのみ：応募済み（当落発表前）の場合、当落通知予定日があれば追記する
    if (!isOnline && !isTournament && status === '応募済み（当落発表前）' && ev.resultAnnouncementDate) {
      const fmt = Utilities.formatDate(ev.resultAnnouncementDate, 'Asia/Tokyo', 'M月d日');
      status += `\n（当落は${fmt}頃にお知らせします）`;
    }

    const line = `【${ev.name}】\n${status}`;
    if (isOnline) onlineLines.push(line);
    else if (isTournament) tournamentLines.push(line);
    else offlineLines.push(line);
  }

  const sections = [];
  if (offlineLines.length > 0) sections.push(`『オフラインイベント』\n\n` + offlineLines.join('\n\n'));
  if (tournamentLines.length > 0) sections.push(`『大会』\n\n` + tournamentLines.join('\n\n'));
  if (onlineLines.length > 0) sections.push(`『オンラインイベント』\n\n` + onlineLines.join('\n\n'));

  const footer = '\n\n──────────\n都合が悪くなってキャンセルを希望される場合や、情報を間違えて入力していた場合は、その旨をこのLINEにご連絡ください。担当者が確認いたします。';

  if (sections.length === 0) {
    replyMessage(replyToken, '現在参加受付中のイベントはありません。' + footer);
  } else {
    replyMessage(replyToken, sections.join('\n\n──────────\n\n') + footer);
  }

  logAction(userId, '応募状況照会', '', '');
}

// 全オンラインイベントのシートから、指定ユーザーの動画相談応募行を検索する（動画状態は問わない＝応募直後に限らずいつでも紐付けられる）
// 同一ユーザーが複数の動画相談に応募している場合は最後に見つかった行（シート上もっとも下＝最新）を採用する
// （動画メッセージ・動画URLテキストの両方の受信処理から共通で使う）
function findVideoApplicationRow_(userId) {
  const allEvents = getAllEvents();
  let found = null;
  for (const ev of allEvents) {
    if ((ev.eventType || 'オフライン') !== 'オンライン') continue;
    const sheet = getSheet(ev.resultSheetName);
    if (!sheet) continue;
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === userId && String(rows[i][12]) === 'video') {
        found = { sheet, rowIdx: i + 1, broadcastName: String(rows[i][10] || '') }; // シートは1始まり
      }
    }
  }
  return found;
}

// 動画メッセージを受信：内容の保存は行わず、応募行が見つかれば「受信済み」を記録して自動返信するだけのシンプルな処理
// 動画本体はLINEのトーク上に残るため、担当者はLINEアプリで直接確認する運用とする
function handleVideoMessage(event) {
  const userId = event.source.userId;
  const found = findVideoApplicationRow_(userId);

  if (found) {
    found.sheet.getRange(found.rowIdx, 16).setValue('受信済み'); // P列：動画状態
    logAction(userId, '動画受信', '', found.broadcastName);
  } else {
    // 応募記録に紐付けできない動画（応募前・応募と無関係に送られたもの）→ スタッフに通知してLINE上で内容を確認してもらう
    notifyStaff(`🎥 未紐付けの動画を受信しました\nUserID: ${userId}\n\n応募記録が見つからないため、LINEのトークで内容をご確認ください。`);
    logAction(userId, '動画受信_未紐付け', '', '');
  }

  replyMessage(event.replyToken, getMsgTemplate_('video_received'));
}

// 動画URL（ギガファイル便等のアップロードサービスのリンク）をテキストで受信
// 動画相談の応募行が見つかれば直接保存して紐付け、見つからない場合もスタッフに通知する
// LINEから動画ファイルを直接受け取る場合と異なりサイズ制限が無いため、ファイルサイズチェック・Drive保存は行わない
function handleVideoUrlMessage(event, url) {
  const userId = event.source.userId;
  const found = findVideoApplicationRow_(userId);

  if (found) {
    // スプレッドシート更新（P列=16: 動画状態、Q列=17: 動画URL）
    found.sheet.getRange(found.rowIdx, 16).setValue('受信済み');
    found.sheet.getRange(found.rowIdx, 17).setValue(url);
    logAction(userId, '動画URL受信', '', found.broadcastName);
  } else {
    // 応募記録に紐付けできないURL（応募前・関係ないリンク等）→ スタッフに通知して手動確認してもらう
    notifyStaff(`🔗 未紐付けのURLを受信しました\nUserID: ${userId}\n${url}\n\n応募記録が見つからないため、内容をご確認ください。`);
    logAction(userId, '動画URL受信_未紐付け', '', url);
  }

  replyMessage(event.replyToken, getMsgTemplate_('video_url_received'));
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
// 参加確認期限（設定シートV列）を過ぎている場合は「期限切れキャンセル」として扱い、期限切れメッセージを返す
function handleParticipationConfirm(replyToken, sheetName, baseUserId) {
  const sheet = getSheet(sheetName);
  if (!sheet) { replyMessage(replyToken, '処理中にエラーが発生しました。'); return; }

  const ev = getAllEvents().find(e => e.resultSheetName === sheetName);
  const isExpired = !!(ev && ev.confirmDeadlineAt && new Date() > ev.confirmDeadlineAt);
  const statusValue = isExpired ? '期限切れキャンセル' : '確認済';

  const data = sheet.getDataRange().getValues();
  let updated = 0;
  for (let i = 1; i < data.length; i++) {
    const uid = String(data[i][1] || '');
    if (uid.replace(/_p\d+$/, '') === baseUserId && String(data[i][2]) === '当選') {
      sheet.getRange(i + 1, 10).setValue(statusValue); // J列
      updated++;
    }
  }

  // シート名は内部識別子（イベント名+日付サフィックス）のため、設定シートに登録された本来のイベント名を使う
  const eventName = ev ? ev.name : sheetName.replace(/_当落$/, '');
  if (updated > 0) {
    if (isExpired) {
      replyMessage(replyToken, renderTemplate_(getMsgTemplate_('participation_expired'), { eventName }));
      logAction(baseUserId, '参加確認_期限切れ', eventName, '');
    } else {
      replyMessage(replyToken, renderTemplate_(getMsgTemplate_('participation_confirmed'), { eventName }));
      logAction(baseUserId, '参加確認', eventName, '');
    }
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

  // シート名は内部識別子（イベント名+日付サフィックス）のため、設定シートに登録された本来のイベント名を使う
  const ev = getAllEvents().find(e => e.resultSheetName === sheetName);
  const eventName = ev ? ev.name : sheetName.replace(/_当落$/, '');
  if (canceledNames.length > 0) {
    replyMessage(replyToken, renderTemplate_(getMsgTemplate_('participation_canceled'), { eventName }));
    notifyStaff(`❌ キャンセル連絡\nイベント: ${eventName}\nお名前: ${canceledNames.join('、')}\n繰り上げ選定をご確認ください。`);
    logAction(baseUserId, 'キャンセル', eventName, canceledNames.join('、'));
  } else {
    replyMessage(replyToken, '既に処理済みか、対象のデータが見つかりませんでした。');
  }
}

// 日付に曜日を付けてフォーマット（例: 2025/06/15 (日)）
function formatDateWithDay(date) {
  if (!date) return '';
  const days = ['日', '月', '火', '水', '木', '金', '土'];
  const dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
  // ISO weekday: 1=月 ... 7=日 → %7 で配列インデックスに変換
  const dayIdx = parseInt(Utilities.formatDate(date, 'Asia/Tokyo', 'u')) % 7;
  return dateStr + ' (' + days[dayIdx] + ')';
}

// 紛らわしい文字（O・0・I・1）を除いた英数字8桁の受付コードを生成
function generateCode() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let code = '';
  for (let i = 0; i < 8; i++) {
    code += chars[Math.floor(Math.random() * chars.length)];
  }
  return code;
}

// スプレッドシートのシートを名前で取得
function getSheet(name) {
  const ss = SpreadsheetApp.openById(getProp('SPREADSHEET_ID'));
  return ss.getSheetByName(name);
}

// ひらがなを全角カタカナに変換する（フリガナ入力の表記ゆれを統一するため。U+3041-3096のひらがな範囲を+0x60シフト）
function toFullWidthKatakana_(str) {
  if (!str) return '';
  return String(str).replace(/[ぁ-ゖ]/g, ch => String.fromCharCode(ch.charCodeAt(0) + 0x60));
}

// 電話番号を「全角→半角」「ハイフン等の除去」で半角数字のみに統一する
function normalizePhone_(str) {
  if (!str) return '';
  const halfWidth = String(str).replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0));
  return halfWidth.replace(/[^0-9]/g, '');
}

// 設定シートO〜V列ヘッダーを補充する（当落メッセージの差し込み用イベント詳細項目）
function ensureEventDetailColumns_(sheet) {
  if (!sheet.getRange(1, 15).getValue()) {
    sheet.getRange(1, 15).setValue('集合時間');
    sheet.getRange(1, 16).setValue('コートについて');
    sheet.getRange(1, 17).setValue('持ち物');
    sheet.getRange(1, 18).setValue('参加費');
    sheet.getRange(1, 19).setValue('更衣室について');
    sheet.getRange(1, 20).setValue('施設URL');
    sheet.getRange(1, 21).setValue('参加確認期限');
  }
  // V列：参加確認期限の実日時（期限切れ自動キャンセル判定用）。表示用テキスト（U列）とは別管理
  if (!sheet.getRange(1, 22).getValue()) {
    sheet.getRange(1, 22).setValue('参加確認期限（日時・自動キャンセル用）');
  }
  // W列：募集締め切り日時（応募状況返信での締め切り時刻表示用）。日付のみのC列とは別管理
  if (!sheet.getRange(1, 23).getValue()) {
    sheet.getRange(1, 23).setValue('募集締め切り日時');
  }
  // X列：当落通知予定日（応募状況返信で「〇月〇日頃に当落をお知らせします」と表示する用）
  if (!sheet.getRange(1, 24).getValue()) {
    sheet.getRange(1, 24).setValue('当落通知予定日');
  }
  // Y列：無料イベントフラグ（TRUE=無料、FALSE/空=有料）
  if (!sheet.getRange(1, 25).getValue()) {
    sheet.getRange(1, 25).setValue('無料イベント');
  }
  // Z列：定員（大会イベント専用。先着順の最大参加人数）
  if (!sheet.getRange(1, 26).getValue()) {
    sheet.getRange(1, 26).setValue('定員');
  }
  // AA列：応募状況非表示（TRUEにすると「応募状況」送信時にこのイベントを表示しない）
  if (!sheet.getRange(1, 27).getValue()) {
    sheet.getRange(1, 27).setValue('応募状況非表示');
  }
  // AB列：限定公開フラグ（TRUEにするとLIFFの通常一覧から非表示になり、専用URL（?invite=）でのみ表示・応募できる）
  if (!sheet.getRange(1, 28).getValue()) {
    sheet.getRange(1, 28).setValue('限定公開');
  }
  // AC列：限定公開コード（専用URLの?invite=パラメータと一致した場合のみイベントを表示・応募可能にする。紹介枠の合言葉としても使う）
  if (!sheet.getRange(1, 29).getValue()) {
    sheet.getRange(1, 29).setValue('限定公開コード');
  }
  // AD列：紹介枠予約人数（大会専用。定員のうちこの人数分は、AC列のコードを持つ人にのみ確保する）
  if (!sheet.getRange(1, 30).getValue()) {
    sheet.getRange(1, 30).setValue('紹介枠予約人数');
  }
  // AE列：先着受付終了日時（大会専用。これ以降の応募は自動当選にせず抽選待ちとして保留し、人数の上限も設けない）
  if (!sheet.getRange(1, 31).getValue()) {
    sheet.getRange(1, 31).setValue('先着受付終了日時');
  }
}

// 設定シートの全イベント行を返す（1行目はヘッダーのためスキップ）
// 戻り値: [{ name, eventDate, closingDate, appSheetName, resultSheetName, winMsg, loseMsg, eventTime, venue, coachName, description, eventType, meetingTime, courtType, items, fee, lockerInfo, facilityUrl, confirmDeadline }, ...]
// L列(eventType): "オフライン" または "オンライン"（空欄の場合は"オフライン"扱い）
function getAllEvents() {
  const sheet = getSheet(SHEET.CONFIG);
  if (!sheet) return [];
  ensureEventDetailColumns_(sheet);
  const data = sheet.getDataRange().getValues();
  const events = [];
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][0]).trim();
    if (!name) continue;
    const eventDate    = data[i][1] ? new Date(data[i][1]) : null;
    const closingDate  = data[i][2] ? new Date(data[i][2]) : null;
    const appSheetName = String(data[i][3] || '').trim();
    const winMsg       = String(data[i][4] || '').trim();  // E列：当選メッセージ（旧・個別上書き用）
    const loseMsg      = String(data[i][5] || '').trim();  // F列：落選メッセージ（旧・個別上書き用）
    const eventTime    = String(data[i][6] || '').trim();  // G列：開催時間（レッスン時間として使用）
    const venue        = String(data[i][7] || '').trim();  // H列：開催場所
    const coachName    = String(data[i][8] || '').trim();  // I列：コーチ名
    const description  = String(data[i][9] || '').trim();  // J列：イベント内容
    const openingDate  = data[i][10] ? new Date(data[i][10]) : null;  // K列：応募開始日
    const eventType    = String(data[i][11] || 'オフライン').trim();    // L列：イベント種別
    const channelUrl   = String(data[i][12] || '').trim();             // M列：チャンネルURL
    const status       = String(data[i][13] || '').trim();             // N列：状態（停止/空=公開）
    const meetingTime     = String(data[i][14] || '').trim(); // O列：集合時間
    const courtType       = String(data[i][15] || '').trim(); // P列：コートについて
    const items           = String(data[i][16] || '').trim(); // Q列：持ち物
    const fee              = String(data[i][17] || '').trim(); // R列：参加費
    const lockerInfo       = String(data[i][18] || '').trim(); // S列：更衣室について
    const facilityUrl      = String(data[i][19] || '').trim(); // T列：施設URL
    const confirmDeadline  = String(data[i][20] || '').trim(); // U列：参加確認期限（表示用テキスト）
    const confirmDeadlineAt   = data[i][21] ? new Date(data[i][21]) : null; // V列：参加確認期限の実日時（期限切れ自動キャンセル判定用）
    const closingDateTimeAt   = data[i][22] ? new Date(data[i][22]) : null; // W列：募集締め切り日時（時刻込み）
    const resultAnnouncementDate = data[i][23] ? new Date(data[i][23]) : null; // X列：当落通知予定日
    const isFreeEvent         = data[i][24] === true || String(data[i][24]).toUpperCase() === 'TRUE'; // Y列：無料イベントフラグ
    const capacity            = parseInt(data[i][25]) || 0; // Z列：定員（大会専用。先着順の最大参加人数）
    const ouboStatusHidden    = data[i][26] === true || String(data[i][26]).toUpperCase() === 'TRUE'; // AA列：応募状況非表示フラグ
    const isRestricted       = data[i][27] === true || String(data[i][27]).toUpperCase() === 'TRUE'; // AB列：限定公開フラグ
    const restrictedCode     = String(data[i][28] || '').trim(); // AC列：限定公開コード（紹介枠の合言葉も兼ねる）
    const referralReserved   = parseInt(data[i][29]) || 0; // AD列：紹介枠予約人数（大会専用）
    const firstComeDeadlineAt = data[i][30] ? new Date(data[i][30]) : null; // AE列：先着受付終了日時（大会専用）
    const resultSheetName = appSheetName
      ? appSheetName.replace('_応募', '_当落')
      : name.replace(/[/?\*[\]:\\]/g, '').replace(/\s/g, '') + '_当落';
    events.push({
      name, eventDate, closingDate, openingDate, appSheetName, resultSheetName, winMsg, loseMsg,
      eventTime, venue, coachName, description, eventType, channelUrl, status,
      meetingTime, courtType, items, fee, lockerInfo, facilityUrl, confirmDeadline, confirmDeadlineAt,
      closingDateTimeAt, resultAnnouncementDate, isFreeEvent, capacity, ouboStatusHidden,
      isRestricted, restrictedCode, referralReserved, firstComeDeadlineAt,
    });
  }
  return events;
}

// アクション履歴シートに1行追加する
function logAction(userId, actionType, eventId, detail) {
  try {
    const sheet = getSheet(SHEET.ACTION_LOG);
    if (!sheet) return;
    sheet.appendRow([new Date(), userId || '', actionType, eventId || '', detail || '']);
  } catch (err) {
    Logger.log('logAction error: ' + err.toString());
  }
}

// LINE APIへのPOSTリクエスト共通処理
function linePost(endpoint, payload) {
  const token = getProp('LINE_CHANNEL_ACCESS_TOKEN');
  const res = UrlFetchApp.fetch(`https://api.line.me/v2/bot/message/${endpoint}`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${token}`,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  // ステータス・本文をエラー時に必ずログへ残す（応答がJSONでない異常時もここで捕捉する）
  const status = res.getResponseCode();
  const bodyText = res.getContentText();
  let result = {};
  try {
    result = bodyText ? JSON.parse(bodyText) : {};
  } catch (err) {
    Logger.log(`LINE API応答のJSON解析失敗 [${endpoint}] status=${status} body=${bodyText}`);
    return { message: 'JSON解析失敗', raw: bodyText };
  }
  if (status !== 200 || result.message) {
    Logger.log(`LINE API error [${endpoint}] status=${status}: ${result.message || bodyText}`);
  }
  return result;
}

// リプライ送信（Webhookで受信したメッセージへの返信）
function replyMessage(replyToken, text) {
  return linePost('reply', {
    replyToken,
    messages: [{ type: 'text', text }],
  });
}

// プッシュ送信（特定のUser IDまたはグループIDへ送信）
function pushMessage(to, text) {
  return linePost('push', {
    to,
    messages: [{ type: 'text', text }],
  });
}

// Quick Reply付きプッシュ送信（当選通知の参加確認ボタンに使用）
function pushMessageWithQuickReply(to, text, quickReply) {
  const msg = { type: 'text', text };
  if (quickReply) msg.quickReply = quickReply;
  return linePost('push', { to, messages: [msg] });
}

// 参加確認用Quick Replyオブジェクトを生成（postbackにシート名とuserIdを埋め込む）
function buildParticipationQuickReply(sheetName, userId) {
  const enc = encodeURIComponent;
  return {
    items: [
      {
        type: 'action',
        action: {
          type: 'postback',
          label: '参加します',
          data: 'action=confirm&sheet=' + enc(sheetName) + '&userId=' + enc(userId),
          displayText: '参加します',
        },
      },
      {
        type: 'action',
        action: {
          type: 'postback',
          label: 'キャンセルします',
          data: 'action=cancel&sheet=' + enc(sheetName) + '&userId=' + enc(userId),
          displayText: 'キャンセルします',
        },
      },
    ],
  };
}

// スタッフグループへ通知（STAFF_GROUP_IDが未設定の場合はスキップ）
function notifyStaff(text) {
  const groupId = getProp('STAFF_GROUP_ID');
  if (!groupId) {
    Logger.log('STAFF_GROUP_ID未設定のためスタッフ通知をスキップ: ' + text);
    return;
  }
  return pushMessage(groupId, text);
}

// 大会シートの現在参加人数をカウント（ペア=2名、1人=1名として集計）
function countTournamentParticipants_(resultSheetName) {
  try {
    const sheet = getSheet(resultSheetName);
    if (!sheet || sheet.getLastRow() <= 1) return 0;
    const data = sheet.getDataRange().getValues();
    let count = 0;
    for (let i = 1; i < data.length; i++) {
      const form = String(data[i][10] || ''); // K列：参加形式
      count += (form === 'ペア') ? 2 : 1;
    }
    return count;
  } catch (err) {
    Logger.log('countTournamentParticipants_ error: ' + err.toString());
    return 0;
  }
}

// アラートメールを ALERT_EMAIL 宛に送信（attachmentsを渡すとBlobを添付する。動画相談の動画添付などで使用）
function sendAlertEmail(subject, body, attachments) {
  const email = getProp('ALERT_EMAIL');
  if (!email) {
    Logger.log('ALERT_EMAIL未設定のためメール送信をスキップ: ' + subject);
    return;
  }
  const options = attachments && attachments.length ? { attachments } : {};
  GmailApp.sendEmail(email, subject, body, options);
}

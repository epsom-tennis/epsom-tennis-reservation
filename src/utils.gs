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

// 設定シートの全イベント行を返す（1行目はヘッダーのためスキップ）
// 戻り値: [{ name, eventDate, closingDate, appSheetName, resultSheetName, winMsg, loseMsg, eventTime, venue, coachName, description, eventType }, ...]
// L列(eventType): "オフライン" または "オンライン"（空欄の場合は"オフライン"扱い）
function getAllEvents() {
  const sheet = getSheet(SHEET.CONFIG);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const events = [];
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][0]).trim();
    if (!name) continue;
    const eventDate    = data[i][1] ? new Date(data[i][1]) : null;
    const closingDate  = data[i][2] ? new Date(data[i][2]) : null;
    const appSheetName = String(data[i][3] || '').trim();
    const winMsg       = String(data[i][4] || '').trim();  // E列：当選メッセージ
    const loseMsg      = String(data[i][5] || '').trim();  // F列：落選メッセージ
    const eventTime    = String(data[i][6] || '').trim();  // G列：開催時間
    const venue        = String(data[i][7] || '').trim();  // H列：開催場所
    const coachName    = String(data[i][8] || '').trim();  // I列：コーチ名
    const description  = String(data[i][9] || '').trim();  // J列：イベント内容
    const openingDate  = data[i][10] ? new Date(data[i][10]) : null;  // K列：応募開始日
    const eventType    = String(data[i][11] || 'オフライン').trim();    // L列：イベント種別
    const channelUrl   = String(data[i][12] || '').trim();             // M列：チャンネルURL
    const status       = String(data[i][13] || '').trim();             // N列：状態（停止/空=公開）
    const resultSheetName = appSheetName
      ? appSheetName.replace('_応募', '_当落')
      : name.replace(/[/?\*[\]:\\]/g, '').replace(/\s/g, '') + '_当落';
    events.push({ name, eventDate, closingDate, openingDate, appSheetName, resultSheetName, winMsg, loseMsg, eventTime, venue, coachName, description, eventType, channelUrl, status });
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

  const result = JSON.parse(res.getContentText());
  if (result.message) {
    Logger.log(`LINE API error [${endpoint}]: ${result.message}`);
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

// アラートメールを ALERT_EMAIL 宛に送信
function sendAlertEmail(subject, body) {
  const email = getProp('ALERT_EMAIL');
  if (!email) {
    Logger.log('ALERT_EMAIL未設定のためメール送信をスキップ: ' + subject);
    return;
  }
  GmailApp.sendEmail(email, subject, body);
}

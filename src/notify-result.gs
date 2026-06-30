// 対象イベントの当落シートを開いた状態でメニューから呼び出す。
// 未送信の行に当落通知を一括送信する。
function sendResults() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = sheet.getName();

  if (!sheetName.endsWith('_当落')) {
    SpreadsheetApp.getUi().alert(
      '対象の当落シートを開いた状態でこの機能を実行してください。\n' +
      '（シート名が「_当落」で終わるシートを選択してください）'
    );
    return;
  }

  const { winCount, loseCount } = sendResultsCore(sheet);

  if (winCount + loseCount === 0) {
    SpreadsheetApp.getUi().alert(
      '送信対象がありません。\n' +
      '「当選」または「落選」が入力されていて、未送信の行があるか確認してください。'
    );
    return;
  }

  // シート名は内部識別子（イベント名+日付サフィックス）のため、設定シートに登録された本来のイベント名を使う
  const ev = getAllEvents().find(e => e.resultSheetName === sheetName);
  const eventName = ev ? ev.name : sheetName.replace('_当落', '');
  notifyStaff(`📨 当落通知 送信完了\n${eventName}\n当選: ${winCount}名 / 落選: ${loseCount}名`);
  SpreadsheetApp.getUi().alert(`送信完了\n当選: ${winCount}名 / 落選: ${loseCount}名`);
}

// メニューからもダッシュボードからも呼び出せる送信処理の共通実装
// 同一LINEアカウントからの複数人応募は1通にまとめて送信し、各参加者の名前を付ける
// オンラインイベントは「当選者には配信をもってご連絡」が基本のため、落選者にはLINEを送らない（シートの送信済み記録のみ行う）
function sendResultsCore(sheet) {
  const sheetName = sheet.getName();
  const ev = getAllEvents().find(e => e.resultSheetName === sheetName);
  const isOnline = !!ev && ev.eventType === 'オンライン';
  const messages = getResultMessages(sheetName);
  const data = sheet.getDataRange().getValues();

  // 送信対象行を基底UserIDでグループ化（_p2, _p3 サフィックスを除いた実際のLINE User ID単位）
  const groups = {};
  for (let i = 1; i < data.length; i++) {
    const userId = String(data[i][1] || '');
    const result = String(data[i][2] || '');
    const sent   = String(data[i][3] || '');
    if (!userId || sent === '済') continue;
    if (result !== '当選' && result !== '落選') continue;

    const baseUserId = userId.replace(/_p\d+$/, '');
    if (!groups[baseUserId]) groups[baseUserId] = [];
    groups[baseUserId].push({ rowIdx: i, name: String(data[i][0] || ''), result });
  }

  let winCount = 0, loseCount = 0, pushCount = 0;

  for (const baseUserId of Object.keys(groups)) {
    const participants = groups[baseUserId];

    // オンラインの落選者はメッセージ送信対象から除外する（オフラインは全員対象）
    const messageParticipants = isOnline ? participants.filter(p => p.result === '当選') : participants;

    // 参加者ごとにメッセージブロックを生成し、複数人なら区切り線でつなぐ
    const blocks = messageParticipants.map(p => {
      const body = p.result === '当選' ? messages.win : messages.lose;
      return (p.name ? p.name + ' 様\n' : '') + body;
    });

    // 当選者がいる場合は参加確認ボタン（Quick Reply）付きで送信。送る内容が無い場合（オンラインで全員落選）はスキップ
    // 送信に失敗した場合（ブロック等でLINE APIがエラーを返した場合）も他のユーザーの処理を止めずに続行し、シートに「送信エラー」を記録する
    let sendFailed = false;
    if (blocks.length > 0) {
      const finalMessage = blocks.join('\n\n──────────\n\n');
      const hasWinners = participants.some(p => p.result === '当選');
      try {
        const pushResult = hasWinners
          ? pushMessageWithQuickReply(baseUserId, finalMessage, buildParticipationQuickReply(sheetName, baseUserId))
          : pushMessage(baseUserId, finalMessage);
        if (pushResult && pushResult.message) sendFailed = true; // LINE APIがエラーレスポンスを返した場合
      } catch (err) {
        Logger.log('sendResultsCore push error [' + baseUserId + ']: ' + err.toString());
        sendFailed = true;
      }
      pushCount++;
      if (pushCount % 10 === 0) Utilities.sleep(1000);
    }

    for (const p of participants) {
      sheet.getRange(p.rowIdx + 1, 4).setValue(sendFailed ? '送信エラー' : '済');
      sheet.getRange(p.rowIdx + 1, 5).setValue(new Date());
      if (p.result === '当選') {
        if (!sendFailed) sheet.getRange(p.rowIdx + 1, 10).setValue('確認待ち'); // J列：参加確認（送信エラー時は確認待ちにしない）
        winCount++;
      } else {
        loseCount++;
      }
      logAction(baseUserId, (sendFailed ? '送信エラー_' : '') + (p.result === '当選' ? '当落通知_当選' : '当落通知_落選'), sheetName.replace('_当落', ''), p.name);
    }
  }

  return { winCount, loseCount };
}

// イベント情報（開催日・コーチ・場所・持ち物など）を当選・落選の基本形テンプレートに差し込んでメッセージを生成する。
// イベント種別（オンライン/オフライン）ごとにダッシュボードの「文章管理」タブで編集できる。
// 該当イベントが設定シートに見つからない場合のみ汎用デフォルト文を返す。
function getResultMessages(resultSheetName) {
  const ev = getAllEvents().find(e => e.resultSheetName === resultSheetName);
  if (!ev) return { win: defaultWinMessage(), lose: defaultLoseMessage() };

  return {
    win:  buildResultMessage_(ev, 'win'),
    lose: buildResultMessage_(ev, 'lose'),
  };
}

function defaultWinMessage() {
  return (
    `【当選のお知らせ】\n` +
    `このたびはイベントへの参加が確定しました！\n` +
    `詳細は別途ご連絡します。\n` +
    `ご参加をお待ちしております。`
  );
}

function defaultLoseMessage() {
  return (
    `【落選のお知らせ】\n` +
    `今回は定員に達したため、ご参加いただけませんでした。\n` +
    `ご応募いただきありがとうございました。\n` +
    `またのご応募をお待ちしています。`
  );
}

// 「参加します／キャンセルします」ボタン付きの当落通知を、STAFF_USER_ID宛に実際に送って動作確認するためのテスト機能。
// 専用の「テスト_当落」シートに自分（STAFF_USER_ID）を当選者として1行登録し、本番と同じsendResultsCoreで送信する。
// ボタンを押すとこのテスト行が本当に更新されるので、参加確定・期限切れ・キャンセルの返信文まで実際の挙動として確認できる。
// 何度実行してもテスト行は1行だけになるよう、既存のテスト行は送信前にリセットする。
function testSendParticipationFlow() {
  const ui = SpreadsheetApp.getUi();
  const staffUserId = getProp('STAFF_USER_ID');
  if (!staffUserId) { ui.alert('STAFF_USER_IDが未設定です。スクリプトプロパティを確認してください。'); return; }

  const sheetName = 'テスト_当落';
  const ss = SpreadsheetApp.openById(getProp('SPREADSHEET_ID'));
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['お名前', 'User ID', '結果', '送信済み', '送信日時', 'コーチについて', '流入経路', '応募きっかけ', '応募日時', '参加確認']);
    sheet.setFrozenRows(1);
  } else {
    // 既存のテスト行をクリアして1行だけに保つ（ヘッダー行は残す）
    if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
  }
  sheet.appendRow(['テスト太郎', staffUserId, '当選', '', '', '', '', '', new Date(), '']);

  const { winCount } = sendResultsCore(sheet);
  if (winCount > 0) {
    ui.alert(
      'テスト送信完了\n\n' +
      'STAFF_USER_IDのLINEに「参加します／キャンセルします」ボタン付きの当落通知を送りました。\n' +
      '実際にボタンを押すと「テスト_当落」シートのJ・C列が本当に更新され、参加確定・期限切れ・キャンセルの返信文も確認できます。\n\n' +
      '確認が終わったら「テスト_当落」シートは削除して問題ありません（次回また自動作成されます）。'
    );
  } else {
    ui.alert('送信に失敗しました。実行ログを確認してください。');
  }
}

// スプレッドシートを開いたときにメニューを追加
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('イベント管理')
    .addItem('当落通知を送信', 'sendResults')
    .addItem('新しいイベントをセットアップ', 'setupNewEvent')
    .addSeparator()
    .addItem('🧪 参加確認フローを自分でテスト送信', 'testSendParticipationFlow')
    .addSeparator()
    .addItem('［一度だけ実行］オンライン列ズレを修正', 'migrateOnlineColumnShift')
    .addItem('［一度だけ実行］電話番号の頭の0を復元', 'migratePhoneLeadingZero')
    .addSeparator()
    .addItem('🗑️ 全データをリセット（本番開始前専用）', 'resetAllData')
    .addToUi();
}

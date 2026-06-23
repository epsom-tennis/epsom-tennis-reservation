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

  const eventName = sheetName.replace('_当落', '');
  notifyStaff(`📨 当落通知 送信完了\n${eventName}\n当選: ${winCount}名 / 落選: ${loseCount}名`);
  SpreadsheetApp.getUi().alert(`送信完了\n当選: ${winCount}名 / 落選: ${loseCount}名`);
}

// メニューからもダッシュボードからも呼び出せる送信処理の共通実装
// 同一LINEアカウントからの複数人応募は1通にまとめて送信し、各参加者の名前を付ける
function sendResultsCore(sheet) {
  const sheetName = sheet.getName();
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

    // 参加者ごとにメッセージブロックを生成し、複数人なら区切り線でつなぐ
    const blocks = participants.map(p => {
      const body = p.result === '当選' ? messages.win : messages.lose;
      return (p.name ? p.name + ' 様\n' : '') + body;
    });
    const finalMessage = blocks.join('\n\n──────────\n\n');

    // 当選者がいる場合は参加確認ボタン（Quick Reply）付きで送信
    const hasWinners = participants.some(p => p.result === '当選');
    if (hasWinners) {
      pushMessageWithQuickReply(baseUserId, finalMessage, buildParticipationQuickReply(sheetName, baseUserId));
    } else {
      pushMessage(baseUserId, finalMessage);
    }
    pushCount++;
    if (pushCount % 10 === 0) Utilities.sleep(1000);

    for (const p of participants) {
      sheet.getRange(p.rowIdx + 1, 4).setValue('済');
      sheet.getRange(p.rowIdx + 1, 5).setValue(new Date());
      if (p.result === '当選') {
        sheet.getRange(p.rowIdx + 1, 10).setValue('確認待ち'); // J列：参加確認
        winCount++;
      } else {
        loseCount++;
      }
      logAction(baseUserId, p.result === '当選' ? '当落通知_当選' : '当落通知_落選', sheetName.replace('_当落', ''), p.name);
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

// スプレッドシートを開いたときにメニューを追加
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('イベント管理')
    .addItem('当落通知を送信', 'sendResults')
    .addItem('新しいイベントをセットアップ', 'setupNewEvent')
    .addToUi();
}

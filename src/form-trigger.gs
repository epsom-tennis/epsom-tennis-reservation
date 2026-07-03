// Googleフォーム送信時に自動実行されるトリガー関数
// ※GASエディタの「トリガー」からスプレッドシートの「フォーム送信時」として設定する（初回1回のみ）
// ※複数イベントのフォームをすべて同一スプレッドシートに連携すれば、このトリガー1つで全イベントに対応できる
function onFormSubmit(e) {
  try {
    // 受付コードを取得・正規化（大文字統一）
    const code = String((e.namedValues['受付コード'] || [''])[0]).trim().toUpperCase();
    const name = String((e.namedValues['お名前'] || ['（名前不明）'])[0]).trim();

    if (!code) {
      notifyStaff('⚠️ フォーム送信：受付コードが空欄でした（お名前: ' + name + '）');
      return;
    }

    // フォームの回答が記録されたシートをイベントの応募シートとして取得
    const appSheet = e.range.getSheet();
    const appSheetName = appSheet.getName();
    const resultSheetName = appSheetName.replace('_応募', '_当落');
    const appData = appSheet.getDataRange().getValues();
    const currentRow = e.range.getRow();
    const eventId = appSheetName.replace('_応募', '');

    // 重複チェック：同じ受付コードで既に応募済み（T列「済」）の行があれば重複とみなす
    for (let i = 1; i < appData.length; i++) {
      const rowNum = i + 1;
      if (rowNum === currentRow) continue;
      const existingCode = String(appData[i][1]).trim().toUpperCase(); // B列（受付コード）
      const sent = appData[i][19]; // T列（応募完了通知済み）
      if (existingCode === code && sent === '済') {
        notifyStaff(
          `⚠️ 重複応募を検知しました\n` +
          `イベント: ${appSheetName}\n` +
          `お名前: ${name}\n受付コード: ${code}\n` +
          `同じ受付コードで既に応募済みの行があります。確認してください。`
        );
        return;
      }
    }

    // 会員マスタC列を走査してUser IDを取得（大文字比較）
    const membersSheet = getSheet(SHEET.MEMBERS);
    const membersData = membersSheet.getDataRange().getValues();

    let userId = null;
    let memberRow = -1;
    for (let i = 1; i < membersData.length; i++) {
      const storedCode = String(membersData[i][2]).trim().toUpperCase(); // C列（受付コード）
      if (storedCode === code) {
        userId = membersData[i][1]; // B列（User ID）
        memberRow = i + 1; // 1-indexed（getRange用）
        break;
      }
    }

    if (!userId) {
      // 受付コード突合失敗：スタッフにアラート＋メール送信
      const alertMsg =
        `受付コード突合失敗\n` +
        `イベント: ${appSheetName}\n` +
        `お名前: ${name}\n入力されたコード: ${code}\n` +
        `会員マスタに一致するコードがありませんでした。\n\n` +
        `【対応方法】\n` +
        `この方にLINEで「応募」と再送するよう伝えてください。\n` +
        `再送すると正しい受付コードが届くので、フォームをもう一度送り直してもらえば自動で処理されます。`;
      notifyStaff('⚠️ ' + alertMsg);
      sendAlertEmail('[テニスイベント] 受付コード突合失敗', alertMsg);
      logAction('', '突合失敗', eventId, `code=${code} name=${name}`);
      return;
    }

    // 応募シートの今回の行にS列（User ID）・T列（済）を記録
    appSheet.getRange(currentRow, 19).setValue(userId); // S列
    appSheet.getRange(currentRow, 20).setValue('済');   // T列

    // 当落シートにお名前・User IDを転記（C〜E列は空欄のまま）
    const ss = SpreadsheetApp.openById(getProp('SPREADSHEET_ID'));
    let resultSheet = ss.getSheetByName(resultSheetName);
    if (!resultSheet) {
      resultSheet = ss.insertSheet(resultSheetName);
      resultSheet.appendRow(['お名前', 'User ID', '結果', '送信済み', '送信日時', 'コーチについて', '流入経路', '応募きっかけ', '応募日時', '参加確認']);
      resultSheet.setFrozenRows(1);
    }
    resultSheet.appendRow([name, userId, '', '', '', '', '', '']);

    // 会員マスタのE〜K列を更新
    if (memberRow > 0) {
      membersSheet.getRange(memberRow, 5).setValue(name);       // E列：名前
      membersSheet.getRange(memberRow, 6).setValue(new Date()); // F列：最終更新日時
      // G〜K列：フォームに対応する質問がある場合のみ更新（空の場合は既存値を保持）
      const extraFields = [
        { key: '年齢',           col: 7  }, // G列
        { key: '性別',           col: 8  }, // H列
        { key: 'テニスレベル',   col: 9  }, // I列
        { key: 'メールアドレス', col: 10 }, // J列
        { key: 'お電話番号',     col: 11 }, // K列
      ];
      for (const f of extraFields) {
        const val = String((e.namedValues[f.key] || [''])[0]).trim();
        if (!val) continue;
        const range = membersSheet.getRange(memberRow, f.col);
        // 電話番号は数値扱いされると頭の0が消えるため、書式をテキスト固定してから書き込む
        if (f.col === 11) range.setNumberFormat('@');
        range.setValue(val);
      }
    }

    // 参加者に応募完了通知を送信
    const ev = getAllEvents().find(function(e) { return e.resultSheetName === resultSheetName; });
    const resultDateStr = (ev && ev.resultAnnouncementDate instanceof Date)
      ? Utilities.formatDate(ev.resultAnnouncementDate, 'Asia/Tokyo', 'M月d日') + '頃に'
      : '後日';
    const evDateLine = (ev && ev.eventDate instanceof Date)
      ? Utilities.formatDate(ev.eventDate, 'Asia/Tokyo', 'M月d日') + '開催\n'
      : '';
    const resultLine = (ev && ev.resultAnnouncementDate instanceof Date)
      ? '\n（当落結果' + Utilities.formatDate(ev.resultAnnouncementDate, 'Asia/Tokyo', 'M月d日') + '頃お知らせ予定）'
      : '';
    const eventBlock = evDateLine + '・' + (ev ? ev.name : 'イベント') + resultLine;
    pushMessage(userId,
      `${eventBlock}\nの応募を受け付けました！\n\nしばらくお待ちください。`
    );

    // スタッフグループに受付完了ログを投稿
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MM/dd HH:mm');
    notifyStaff(`✅ 応募受付完了 ${now}\nイベント: ${appSheetName}\nお名前: ${name}`);

    // アクション履歴に記録
    logAction(userId, '応募完了', eventId, name);

  } catch (err) {
    Logger.log('onFormSubmit error: ' + err.toString());
    notifyStaff('⚠️ フォームトリガーでエラーが発生しました：' + err.toString());
  }
}

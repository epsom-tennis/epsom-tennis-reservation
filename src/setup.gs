// 初回セットアップ用。固定シートを自動作成する。GASエディタから一度だけ手動実行する。
// ※旧「LINE登録者シート」が存在する場合は「会員マスタ」に自動改名する。
function setupSpreadsheet() {
  const ss = SpreadsheetApp.openById(getProp('SPREADSHEET_ID'));

  // 設定シート（1行目はヘッダー、2行目以降にイベントを1行ずつ追加する）
  let configSheet = ss.getSheetByName(SHEET.CONFIG);
  if (!configSheet) {
    configSheet = ss.insertSheet(SHEET.CONFIG);
  }
  if (configSheet.getLastRow() === 0) {
    configSheet.appendRow(['イベント名', '開催日', '募集終了日', '応募シート名', '当選メッセージ', '落選メッセージ']);
    configSheet.setFrozenRows(1);
  }

  // 会員マスタ（旧: LINE登録者シートから自動移行）
  let membersSheet = ss.getSheetByName(SHEET.MEMBERS);
  if (!membersSheet) {
    const oldSheet = ss.getSheetByName('LINE登録者シート');
    if (oldSheet) {
      // 旧シートを改名して列を追加
      oldSheet.setName(SHEET.MEMBERS);
      membersSheet = oldSheet;
      if (membersSheet.getLastRow() > 0) {
        membersSheet.getRange(1, 5).setValue('名前');
        membersSheet.getRange(1, 6).setValue('最終更新日時');
      }
    } else {
      membersSheet = ss.insertSheet(SHEET.MEMBERS);
      membersSheet.appendRow(['登録日時', 'User ID', '受付コード', '備考', '名前', '最終更新日時', '年齢', '性別', 'テニスレベル', 'メールアドレス', '電話番号', 'フリガナ', '緊急連絡先', 'テニス頻度', 'テニス歴', 'テニス地域', 'テニス環境']);
      membersSheet.setFrozenRows(1);
    }
  }

  // 会員マスタのG〜K列ヘッダーを追加（G列が空の場合のみ・既存シートへの追記）
  if (membersSheet.getLastRow() > 0 && !membersSheet.getRange(1, 7).getValue()) {
    membersSheet.getRange(1, 7).setValue('年齢');
    membersSheet.getRange(1, 8).setValue('性別');
    membersSheet.getRange(1, 9).setValue('テニスレベル');
    membersSheet.getRange(1, 10).setValue('メールアドレス');
    membersSheet.getRange(1, 11).setValue('電話番号');
  }
  // L〜Q列ヘッダー（LIFFフォーム対応で追加）
  if (membersSheet.getLastRow() > 0 && !membersSheet.getRange(1, 12).getValue()) {
    membersSheet.getRange(1, 12).setValue('フリガナ');
    membersSheet.getRange(1, 13).setValue('緊急連絡先');
    membersSheet.getRange(1, 14).setValue('テニス頻度');
    membersSheet.getRange(1, 15).setValue('テニス歴');
    membersSheet.getRange(1, 16).setValue('テニス地域');
    membersSheet.getRange(1, 17).setValue('テニス環境');
  }

  // 設定シートのG〜J列ヘッダーを追加（イベント詳細フィールド）
  if (configSheet.getLastRow() > 0 && !configSheet.getRange(1, 7).getValue()) {
    configSheet.getRange(1, 7).setValue('開催時間');
    configSheet.getRange(1, 8).setValue('開催場所');
    configSheet.getRange(1, 9).setValue('コーチ名');
    configSheet.getRange(1, 10).setValue('イベント内容');
  }
  // 設定シートのK〜L列ヘッダーを追加（応募開始日・イベント種別）
  if (configSheet.getLastRow() > 0 && !configSheet.getRange(1, 12).getValue()) {
    if (!configSheet.getRange(1, 11).getValue()) configSheet.getRange(1, 11).setValue('応募開始日');
    configSheet.getRange(1, 12).setValue('イベント種別');
  }

  // アクション履歴シート
  let actionSheet = ss.getSheetByName(SHEET.ACTION_LOG);
  if (!actionSheet) {
    actionSheet = ss.insertSheet(SHEET.ACTION_LOG);
    actionSheet.appendRow(['タイムスタンプ', 'User ID', 'アクション種別', 'イベント識別名', '詳細']);
    actionSheet.setFrozenRows(1);
  }

  SpreadsheetApp.getUi().alert(
    'セットアップ完了！\n\n' +
    '作成・確認したシート:\n' +
    '・設定シート\n' +
    '・会員マスタ\n' +
    '・アクション履歴\n\n' +
    '【次にやること】\n' +
    '「イベント管理 > 新しいイベントをセットアップ」からイベントを追加してください。'
  );
}

// ダッシュボードからイベントを新規作成する（google.script.runから呼び出す）
function createNewEvent(data) {
  try {
    const { name, eventDate, closingDate, openingDate, eventTime, venue, coachName, description, eventType } = data;
    if (!name || !eventDate || !closingDate) {
      return { success: false, error: 'イベント名・開催日・募集終了日は必須です。' };
    }
    const ss = SpreadsheetApp.openById(getProp('SPREADSHEET_ID'));
    const configSheet = ss.getSheetByName(SHEET.CONFIG);
    if (!configSheet) return { success: false, error: '設定シートが見つかりません。' };

    const evDateObj = new Date(eventDate);
    if (isNaN(evDateObj.getTime())) return { success: false, error: '開催日の形式が正しくありません（YYYY/MM/DD）。' };
    const closingDateObj = new Date(closingDate);
    if (isNaN(closingDateObj.getTime())) return { success: false, error: '募集終了日の形式が正しくありません（YYYY/MM/DD）。' };
    let openingDateObj = null;
    if (openingDate) {
      openingDateObj = new Date(openingDate);
      if (isNaN(openingDateObj.getTime())) return { success: false, error: '応募開始日の形式が正しくありません（YYYY/MM/DD）。' };
    }

    // 同名イベントの重複チェック
    const configData = configSheet.getDataRange().getValues();
    for (let i = 1; i < configData.length; i++) {
      if (String(configData[i][0]).trim() === name.trim()) {
        return { success: false, error: `同名のイベント「${name}」が既に存在します。` };
      }
    }

    // シート名用識別子を生成（シート名に使えない文字を除去）
    const identifier = name.replace(/[/?\*[\]:\\]/g, '').replace(/\s/g, '').substring(0, 20)
      + '_' + Utilities.formatDate(evDateObj, 'Asia/Tokyo', 'MMdd');
    const appSheetName    = identifier + '_応募';
    const resultSheetName = identifier + '_当落';

    // 当落シートを作成
    let resultSheet = ss.getSheetByName(resultSheetName);
    if (!resultSheet) {
      resultSheet = ss.insertSheet(resultSheetName);
      resultSheet.appendRow(['お名前', 'User ID', '結果', '送信済み', '送信日時', 'コーチについて', '流入経路', '応募きっかけ', '応募日時', '参加確認']);
      resultSheet.setFrozenRows(1);
    }

    // 設定シートに行を追加（G〜K列に詳細情報）
    configSheet.appendRow([
      name, evDateObj, closingDateObj, appSheetName, '', '',
      eventTime || '', venue || '', coachName || '', description || '', openingDateObj || '', eventType || 'オフライン',
    ]);

    return {
      success: true,
      appSheetName,
      resultSheetName,
      eventDate:   Utilities.formatDate(evDateObj,      'Asia/Tokyo', 'yyyy/MM/dd'),
      closingDate: Utilities.formatDate(closingDateObj, 'Asia/Tokyo', 'yyyy/MM/dd'),
    };
  } catch (err) {
    Logger.log('createNewEvent error: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

// 新しいイベントをセットアップする。
// イベントごとのGoogleフォームをスプレッドシートに連携した後に実行する。
function setupNewEvent() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(getProp('SPREADSHEET_ID'));

  // 識別名の入力
  const idRes = ui.prompt(
    '新しいイベントのセットアップ (1/4)',
    '識別名を入力してください（例: コーチA6月）\n' +
    '※シート名「{識別名}_応募」「{識別名}_当落」に使用されます',
    ui.ButtonSet.OK_CANCEL
  );
  if (idRes.getSelectedButton() !== ui.Button.OK) return;
  const identifier = idRes.getResponseText().trim();
  if (!identifier) { ui.alert('識別名が入力されていません。'); return; }

  // イベント名の入力
  const nameRes = ui.prompt(
    '新しいイベントのセットアップ (2/4)',
    'イベント名を入力してください（例: コーチAレッスン 6月15日）',
    ui.ButtonSet.OK_CANCEL
  );
  if (nameRes.getSelectedButton() !== ui.Button.OK) return;
  const eventName = nameRes.getResponseText().trim();

  // 開催日の入力
  const dateRes = ui.prompt(
    '新しいイベントのセットアップ (3/4)',
    '開催日を入力してください（例: 2026/06/15）',
    ui.ButtonSet.OK_CANCEL
  );
  if (dateRes.getSelectedButton() !== ui.Button.OK) return;
  const eventDate = new Date(dateRes.getResponseText().trim());
  if (isNaN(eventDate.getTime())) {
    ui.alert('開催日の形式が正しくありません。YYYY/MM/DDで入力してください。');
    return;
  }

  // 募集終了日の入力
  const closingRes = ui.prompt(
    '新しいイベントのセットアップ (4/4)',
    '募集終了日を入力してください（例: 2026/06/10）',
    ui.ButtonSet.OK_CANCEL
  );
  if (closingRes.getSelectedButton() !== ui.Button.OK) return;
  const closingDate = new Date(closingRes.getResponseText().trim());
  if (isNaN(closingDate.getTime())) {
    ui.alert('募集終了日の形式が正しくありません。YYYY/MM/DDで入力してください。');
    return;
  }

  const appSheetName = identifier + '_応募';
  const resultSheetName = identifier + '_当落';

  // 応募シートの存在確認（Googleフォームとの連携が先に必要）
  const appSheet = ss.getSheetByName(appSheetName);
  if (!appSheet) {
    ui.alert(
      `シート「${appSheetName}」が見つかりません。\n\n` +
      `先にGoogleフォームの回答先として「${appSheetName}」という名前のシートを作成・連携してから実行してください。`
    );
    return;
  }

  // 当落シートの作成
  let resultSheet = ss.getSheetByName(resultSheetName);
  if (!resultSheet) {
    resultSheet = ss.insertSheet(resultSheetName);
    resultSheet.appendRow(['お名前', 'User ID', '結果', '送信済み', '送信日時', 'コーチについて', '流入経路', '応募きっかけ', '応募日時']);
    resultSheet.setFrozenRows(1);
  }

  // 設定シートにイベント行を追加（E・F列は空欄。後から設定シートで直接入力する）
  const configSheet = ss.getSheetByName(SHEET.CONFIG);
  configSheet.appendRow([eventName, eventDate, closingDate, appSheetName, '', '']);

  const eventDateStr = Utilities.formatDate(eventDate, 'Asia/Tokyo', 'yyyy/MM/dd');
  const closingDateStr = Utilities.formatDate(closingDate, 'Asia/Tokyo', 'yyyy/MM/dd');

  ui.alert(
    `セットアップ完了！\n\n` +
    `イベント名: ${eventName}\n` +
    `開催日: ${eventDateStr}\n` +
    `募集終了日: ${closingDateStr}\n` +
    `応募シート: ${appSheetName}\n` +
    `当落シート: ${resultSheetName}（新規作成）\n\n` +
    `【当選・落選メッセージについて】\n` +
    `設定シートのE列・F列に文面を入力してください。\n` +
    `空欄の場合はデフォルト文が使用されます。\n\n` +
    `【onFormSubmitトリガーについて】\n` +
    `初めてのイベントの場合のみ、GASエディタの「トリガー」から\n` +
    `onFormSubmit をスプレッドシートの「フォーム送信時」として設定してください。\n` +
    `2つ目以降のイベントは既存のトリガーが自動で対応します。`
  );
}

// 設定シートにイベント種別列（L列）のヘッダーを追加する。GASエディタから一度だけ手動実行。
function addEventTypeHeader() {
  const ss = SpreadsheetApp.openById(getProp('SPREADSHEET_ID'));
  const configSheet = ss.getSheetByName(SHEET.CONFIG);
  if (!configSheet) { SpreadsheetApp.getUi().alert('設定シートが見つかりません。'); return; }
  if (!configSheet.getRange(1, 11).getValue()) configSheet.getRange(1, 11).setValue('応募開始日');
  configSheet.getRange(1, 12).setValue('イベント種別');
  SpreadsheetApp.getUi().alert('完了！設定シートのL列に「イベント種別」を追加しました。\n\nオンラインイベントの行に「オンライン」と入力してください。\n（空欄はオフライン扱い）');
}

// システム診断（テスト用）：シート存在確認・スクリプトプロパティ確認・実データ確認
function runDiagnose() {
  const result = { props: {}, sheets: {}, membersLastRows: [], actionLogLastRows: [], errors: [] };
  try {
    result.props.spreadsheetId  = !!getProp('SPREADSHEET_ID');
    result.props.lineToken      = !!getProp('LINE_CHANNEL_ACCESS_TOKEN');
    result.props.liffId         = !!getProp('LIFF_ID');
    result.props.dashboardToken = !!getProp('DASHBOARD_TOKEN');
  } catch (e) { result.errors.push('props: ' + e.toString()); }

  try {
    const ss = SpreadsheetApp.openById(getProp('SPREADSHEET_ID'));
    const names = ss.getSheets().map(s => s.getName());
    result.sheets.found     = names;
    result.sheets.config    = names.includes(SHEET.CONFIG);
    result.sheets.members   = names.includes(SHEET.MEMBERS);
    result.sheets.actionLog = names.includes(SHEET.ACTION_LOG);

    // 会員マスタの末尾3行を取得
    const membersSheet = ss.getSheetByName(SHEET.MEMBERS);
    if (membersSheet) {
      result.sheets.membersLastRow = membersSheet.getLastRow();
      const lastRows = membersSheet.getDataRange().getValues().slice(-3);
      result.membersLastRows = lastRows.map(r => ({ col_B_userId: String(r[1]), col_C_code: String(r[2]), col_E_name: String(r[4]) }));
    }

    // アクション履歴の末尾3行を取得
    const logSheet = ss.getSheetByName(SHEET.ACTION_LOG);
    if (logSheet && logSheet.getLastRow() > 1) {
      result.actionLogLastRows = logSheet.getDataRange().getValues().slice(-3)
        .map(r => ({ time: String(r[0]), userId: String(r[1]), action: String(r[2]) }));
    }
  } catch (e) { result.errors.push('sheets: ' + e.toString()); }

  return result;
}

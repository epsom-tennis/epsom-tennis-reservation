// 管理ダッシュボード + LIFF向けJSON API
// GASのWebアプリとしてデプロイ（実行ユーザー：自分、アクセス権：全員）
function doGet(e) {
  const page   = e && e.parameter && e.parameter.page;
  const action = e && e.parameter && e.parameter.action;
  const token  = e && e.parameter && e.parameter.token;

  // LIFF向けJSON API（GitHub PagesのHTMLからfetch()で呼び出す）
  if (action === 'getEvents') {
    const userId = e && e.parameter && e.parameter.userId;
    return liffApiResponse(getLiffEventsJson(userId));
  }
  if (action === 'getMember') {
    const userId = e && e.parameter && e.parameter.userId;
    return liffApiResponse(getMemberData(userId));
  }
  // 診断エンドポイント（テスト用）
  if (action === 'diagnose' && token === getProp('DASHBOARD_TOKEN')) {
    return liffApiResponse(runDiagnose());
  }
  // LIFFエンドポイントURL更新（テスト用）
  if (action === 'updateLiffEndpoint' && token === getProp('DASHBOARD_TOKEN')) {
    const newUrl = e.parameter.url;
    if (!newUrl) return liffApiResponse({ ok: false, error: 'url required' });
    try {
      const liffId = getProp('LIFF_ID');
      const lineToken = getProp('LINE_CHANNEL_ACCESS_TOKEN');
      const res = UrlFetchApp.fetch(`https://api.line.me/liff/v1/apps/${liffId}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${lineToken}` },
        payload: JSON.stringify({ view: { type: 'full', url: newUrl } }),
        muteHttpExceptions: true,
      });
      const body = res.getContentText();
      const status = res.getResponseCode();
      return liffApiResponse({ ok: status === 200, status, body });
    } catch (err) {
      return liffApiResponse({ ok: false, error: err.toString() });
    }
  }
  // LINE push送信テスト（テスト用）
  if (action === 'testPush' && token === getProp('DASHBOARD_TOKEN')) {
    const targetUserId = e.parameter.userId;
    if (!targetUserId) return liffApiResponse({ ok: false, error: 'userId required' });
    try {
      const result = pushMessage(targetUserId, '[テスト送信] GASからのLINE送信テストです。このメッセージが届いていれば正常です。');
      return liffApiResponse({ ok: true, lineResult: result });
    } catch (err) {
      return liffApiResponse({ ok: false, error: err.toString() });
    }
  }
  // Webhookの応募処理を直接テスト（テスト用）
  if (action === 'testOubo' && token === getProp('DASHBOARD_TOKEN')) {
    const testUserId = (e.parameter.userId || 'Utest_direct_001');
    const fakeEvent = {
      source: { userId: testUserId },
      replyToken: 'test_token_skip_reply'
    };
    try {
      handleOubo(fakeEvent);
      return liffApiResponse({ ok: true, userId: testUserId });
    } catch (err) {
      return liffApiResponse({ ok: false, error: err.toString() });
    }
  }

  // 旧LIFFページ（後方互換のため残す）
  if (page === 'liff') {
    return getLiffPage();
  }

  // ダッシュボードはDASHBOARD_TOKENで保護
  const dashboardToken = getProp('DASHBOARD_TOKEN');
  if (!dashboardToken || token !== dashboardToken) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;text-align:center;padding:80px 20px">' +
      '<h2>⛔ アクセス権限がありません</h2>' +
      '<p style="color:#888">正しいURLでアクセスしてください。</p></div>'
    ).setTitle('Access Denied');
  }

  return HtmlService.createHtmlOutput(getDashboardHtml())
    .setTitle('イベント管理ダッシュボード')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// JSON APIレスポンスを生成する
function liffApiResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// LIFF向けイベント一覧を返す（応募開始日・締切日でフィルタ済み）
// userIdが渡された場合はalreadyAppliedフラグも付与する
function getLiffEventsJson(userId) {
  try {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const filtered = getAllEvents()
      .filter(ev => ev.status !== '停止' && (!ev.openingDate || ev.openingDate <= today) && (!ev.closingDate || ev.closingDate >= today));

    // 応募済みシート名のセットを構築
    const appliedSheets = new Set();
    if (userId) {
      const baseUserId = userId.replace(/_p\d+$/, '');
      for (const ev of filtered) {
        const sheet = getSheet(ev.resultSheetName);
        if (!sheet || sheet.getLastRow() <= 1) continue;
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          const uid = String(data[i][1] || '');
          if (uid === userId || uid.replace(/_p\d+$/, '') === baseUserId) {
            appliedSheets.add(ev.resultSheetName);
            break;
          }
        }
      }
    }

    return filtered.map(ev => ({
      name:            ev.name,
      resultSheetName: ev.resultSheetName,
      eventDate:       formatDateWithDay(ev.eventDate),
      closingDate:     ev.closingDate ? Utilities.formatDate(ev.closingDate, 'Asia/Tokyo', 'yyyy/MM/dd') : '',
      eventTime:       ev.eventTime   || '',
      venue:           ev.venue       || '',
      coachName:       ev.coachName   || '',
      description:     ev.description || '',
      eventType:       ev.eventType   || 'オフライン',
      channelUrl:      ev.channelUrl  || '',
      alreadyApplied:  appliedSheets.has(ev.resultSheetName),
    }));
  } catch (err) {
    Logger.log('getLiffEventsJson error: ' + err.toString());
    return [];
  }
}

// ===== クライアントから呼び出すサーバー関数 =====

// 全イベントの一覧と統計情報を返す
function getEventsData() {
  const events = getAllEvents();
  return events.map(ev => {
    const appSheet = getSheet(ev.appSheetName);
    const resultSheet = getSheet(ev.resultSheetName);

    // 応募数は当落シートを基準にカウント（Google FormとLIFF両方を含む）
    const appCount = resultSheet && resultSheet.getLastRow() > 1 ? resultSheet.getLastRow() - 1 : 0;

    let winCount = 0, loseCount = 0, sentCount = 0, pendingCount = 0;
    if (resultSheet && resultSheet.getLastRow() > 1) {
      const data = resultSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const result = data[i][2];
        const sent = data[i][3];
        if (result === '当選') winCount++;
        if (result === '落選') loseCount++;
        if (sent === '済') sentCount++;
        if ((result === '当選' || result === '落選') && sent !== '済') pendingCount++;
      }
    }

    return {
      name: ev.name,
      eventDate:    formatDateWithDay(ev.eventDate),
      closingDate:  ev.closingDate  ? Utilities.formatDate(ev.closingDate,  'Asia/Tokyo', 'yyyy/MM/dd') : '',
      openingDate:  ev.openingDate  ? Utilities.formatDate(ev.openingDate,  'Asia/Tokyo', 'yyyy/MM/dd') : '',
      appSheetName: ev.appSheetName,
      resultSheetName: ev.resultSheetName,
      eventTime:   ev.eventTime   || '',
      venue:       ev.venue       || '',
      coachName:   ev.coachName   || '',
      description: ev.description || '',
      appCount, winCount, loseCount, sentCount, pendingCount,
      status: ev.status || '',
    };
  });
}

// 指定イベントの応募者一覧を返す（当落シートを基準に、Google FormとLIFF両方の応募を含む）
function getApplicants(appSheetName, resultSheetName) {
  const resultSheet = getSheet(resultSheetName);
  if (!resultSheet || resultSheet.getLastRow() <= 1) return [];

  // 会員マスタからテニス情報マップを作成（User ID → 年齢・性別・レベル等）
  const memberMap = {};
  const membersSheet = getSheet(SHEET.MEMBERS);
  if (membersSheet && membersSheet.getLastRow() > 1) {
    const mData = membersSheet.getDataRange().getValues();
    for (let i = 1; i < mData.length; i++) {
      const uid = String(mData[i][1] || '');
      if (uid) {
        memberMap[uid] = {
          age:           String(mData[i][6]  || ''),
          gender:        String(mData[i][7]  || ''),
          tennisLevel:   String(mData[i][8]  || ''),
          tennisFreq:    String(mData[i][13] || ''),
          tennisHistory: String(mData[i][14] || ''),
        };
      }
    }
  }

  // 応募シートから応募日時マップを作成（Google Form経由の応募）
  const appDateMap = {};
  const appSheet = getSheet(appSheetName);
  if (appSheet && appSheet.getLastRow() > 1) {
    const appData = appSheet.getDataRange().getValues();
    for (let i = 1; i < appData.length; i++) {
      const uid = String(appData[i][18] || ''); // S列
      if (uid && !appDateMap[uid]) {
        appDateMap[uid] = appData[i][0]
          ? Utilities.formatDate(new Date(appData[i][0]), 'Asia/Tokyo', 'MM/dd HH:mm')
          : '';
      }
    }
  }

  const winCountMap = buildWinCountMap();

  const data = resultSheet.getDataRange().getValues();
  const applicants = [];
  for (let i = 1; i < data.length; i++) {
    const userId = String(data[i][1] || '');
    if (!userId) continue;
    // 応募日時: Google Form応募シートに記録があればそれを優先、なければ当落シートI列（LIFF応募日時）を使用
    let appliedAt = appDateMap[userId] || '';
    if (!appliedAt && data[i][8]) {
      try { appliedAt = Utilities.formatDate(new Date(data[i][8]), 'Asia/Tokyo', 'MM/dd HH:mm'); } catch(e) {}
    }
    const mInfo = memberMap[userId] || {};
    const baseUserId = userId.replace(/_p\d+$/, '');
    applicants.push({
      name:           String(data[i][0] || ''),
      userId,
      appliedAt,
      result:         String(data[i][2] || ''),
      sent:           String(data[i][3] || ''),
      coachKnowledge: String(data[i][5] || ''),
      confirmation:   String(data[i][9] || ''),
      winCount:       winCountMap[baseUserId] || 0,
      age:            mInfo.age           || '',
      gender:         mInfo.gender        || '',
      tennisLevel:    mInfo.tennisLevel   || '',
      tennisFreq:     mInfo.tennisFreq    || '',
      tennisHistory:  mInfo.tennisHistory || '',
    });
  }
  return applicants;
}

// 全当落シートを横断して各UserIDの当選回数を集計する（_p2/_p3は基底IDに統合）
function buildWinCountMap() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const map = {};
  for (const sheet of sheets) {
    if (!sheet.getName().endsWith('_当落')) continue;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const uid = String(data[i][1] || '');
      if (uid && String(data[i][2] || '') === '当選') {
        const base = uid.replace(/_p\d+$/, '');
        map[base] = (map[base] || 0) + 1;
      }
    }
  }
  return map;
}

// 複数の当落結果を一括で書き込む（results: [{userId, result}, ...]）
function setResultsBatch(resultSheetName, results) {
  if (!results || results.length === 0) return { success: false, error: '結果が指定されていません。' };
  const sheet = getSheet(resultSheetName);
  if (!sheet) return { success: false, error: 'シートが見つかりません。' };
  const data = sheet.getDataRange().getValues();
  const resultMap = {};
  results.forEach(r => { resultMap[r.userId] = r.result; });
  for (let i = 1; i < data.length; i++) {
    const uid = String(data[i][1]);
    if (resultMap[uid] !== undefined) {
      sheet.getRange(i + 1, 3).setValue(resultMap[uid]);
    }
  }
  return { success: true, count: results.length };
}

// 当落シートの指定User IDの行のC列に当落を書き込む
function setResult(resultSheetName, userId, result) {
  if (result !== '当選' && result !== '落選') {
    return { success: false, error: '結果は「当選」または「落選」のみ指定できます。' };
  }
  const sheet = getSheet(resultSheetName);
  if (!sheet) return { success: false, error: 'シートが見つかりません。' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === userId) {
      sheet.getRange(i + 1, 3).setValue(result);
      return { success: true };
    }
  }
  return { success: false, error: '指定されたUser IDが見つかりません。' };
}

// ダッシュボードから当落通知を一括送信する（sendResultsCoreを呼び出す）
function sendResultsFromDashboard(resultSheetName) {
  const sheet = getSheet(resultSheetName);
  if (!sheet) return { success: false, error: 'シートが見つかりません。' };

  const { winCount, loseCount } = sendResultsCore(sheet);

  if (winCount + loseCount > 0) {
    const eventName = resultSheetName.replace('_当落', '');
    notifyStaff(`📨 当落通知 送信完了\n${eventName}\n当選: ${winCount}名 / 落選: ${loseCount}名`);
  }

  return { success: true, winCount, loseCount };
}

// ステータスで絞り込んだUser IDリストを返す
function getFilteredUsers(appSheetName, resultSheetName, status) {
  const membersSheet = getSheet(SHEET.MEMBERS);
  if (!membersSheet || membersSheet.getLastRow() <= 1) return [];

  // 応募済みUser IDのセットを構築
  const submittedSet = new Set();
  const appSheet = getSheet(appSheetName);
  if (appSheet && appSheet.getLastRow() > 1) {
    const appData = appSheet.getDataRange().getValues();
    for (let i = 1; i < appData.length; i++) {
      if (appData[i][18]) submittedSet.add(String(appData[i][18]));
    }
  }

  // 当落ステータスのマップを構築
  const resultMap = {};
  const resultSheet = getSheet(resultSheetName);
  if (resultSheet && resultSheet.getLastRow() > 1) {
    const resultData = resultSheet.getDataRange().getValues();
    for (let i = 1; i < resultData.length; i++) {
      resultMap[String(resultData[i][1])] = String(resultData[i][2] || '');
    }
  }

  const membersData = membersSheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < membersData.length; i++) {
    const userId = String(membersData[i][1]);
    const name = String(membersData[i][4] || membersData[i][2] || '（名前未取得）'); // E列（名前）優先、なければ受付コード
    const userResult = resultMap[userId] || '';
    const isSubmitted = submittedSet.has(userId);

    let match = false;
    if (status === '当選' && userResult === '当選') match = true;
    else if (status === '落選' && userResult === '落選') match = true;
    else if (status === '応募済み' && isSubmitted && !userResult) match = true;
    else if (status === '未応募' && !isSubmitted) match = true;

    if (match) result.push({ name, userId });
  }
  return result;
}

// 指定User IDリストに一括でメッセージを送信する
function sendBroadcast(userIds, message) {
  if (!message || !userIds || userIds.length === 0) {
    return { success: false, error: '送信先またはメッセージが指定されていません。' };
  }

  let count = 0;
  for (const userId of userIds) {
    pushMessage(userId, message);
    logAction(userId, '絞り込み送信', '', message.substring(0, 50));
    count++;
    if (count % 10 === 0) Utilities.sleep(1000);
  }

  notifyStaff(`📢 絞り込み送信完了\n${count}名に送信しました`);
  return { success: true, count };
}

// 会員マスタの全会員データを返す
function getMembersData() {
  const sheet = getSheet(SHEET.MEMBERS);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data = sheet.getDataRange().getValues();
  const members = [];
  for (let i = 1; i < data.length; i++) {
    const userId = String(data[i][1] || '');
    if (!userId) continue;
    members.push({
      userId,
      name:          String(data[i][4]  || ''),
      furigana:      String(data[i][11] || ''),
      age:           String(data[i][6]  || ''),
      gender:        String(data[i][7]  || ''),
      tennisLevel:   String(data[i][8]  || ''),
      tennisFreq:    String(data[i][13] || ''),
      tennisHistory: String(data[i][14] || ''),
      tennisArea:    String(data[i][15] || ''),
      tennisEnv:     String(data[i][16] || ''),
      email:         String(data[i][9]  || ''),
      phone:         String(data[i][10] || '').replace(/^(\d{9,10})$/, '0$1'),
      registeredAt: data[i][0] ? Utilities.formatDate(new Date(data[i][0]), 'Asia/Tokyo', 'yyyy/MM/dd') : '',
    });
  }
  return members;
}

// 指定User IDの全イベント応募履歴を返す
function getMemberHistory(userId) {
  try {
    const events = getAllEvents();
    const history = [];
    for (const ev of events) {
      const resultSheet = getSheet(ev.resultSheetName);
      if (!resultSheet || resultSheet.getLastRow() <= 1) continue;
      const data = resultSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][1]) === userId) {
          let appliedAt = '';
          if (data[i][8]) {
            try { appliedAt = Utilities.formatDate(new Date(data[i][8]), 'Asia/Tokyo', 'MM/dd HH:mm'); } catch(e) {}
          }
          history.push({
            eventName: ev.name,
            eventDate: formatDateWithDay(ev.eventDate),
            result:    String(data[i][2] || '未処理'),
            appliedAt,
          });
          break;
        }
      }
    }
    return { success: true, history };
  } catch (err) {
    Logger.log('getMemberHistory error: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

// イベントの公開状態を切り替える（停止 ↔ 公開）
function toggleEventStatus(appSheetName) {
  try {
    const ss = SpreadsheetApp.openById(getProp('SPREADSHEET_ID'));
    const configSheet = ss.getSheetByName(SHEET.CONFIG);
    if (!configSheet) return { success: false, error: '設定シートが見つかりません。' };
    const data = configSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][3]).trim() === appSheetName) {
        const isStopped = String(data[i][13] || '').trim() === '停止';
        configSheet.getRange(i + 1, 14).setValue(isStopped ? '' : '停止');
        return { success: true };
      }
    }
    return { success: false, error: 'イベントが見つかりません。' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ダッシュボード設定（LIFF IDなど）を返す
function getDashboardConfig() {
  return { liffId: getProp('LIFF_ID') || '' };
}

// ===== HTMLの生成 =====

function getDashboardHtml() {
  return '<!DOCTYPE html>' +
'<html lang="ja">' +
'<head>' +
'<meta charset="UTF-8">' +
'<meta name="viewport" content="width=device-width, initial-scale=1">' +
'<title>イベント管理ダッシュボード</title>' +
'<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">' +
'<style>' +
'body{font-size:14px;background:#f8f9fa}' +
'.event-card{cursor:pointer;transition:box-shadow .15s}' +
'.event-card:hover{box-shadow:0 2px 8px rgba(0,0,0,.15)}' +
'.event-card.active{border-color:#0d6efd!important;background:#f0f4ff}' +
'.badge-pending{background:#ffc107;color:#000}' +
'.status-当選{color:#198754;font-weight:bold}' +
'.status-落選{color:#dc3545}' +
'.status-未処理{color:#6c757d}' +
'.status-キャンセル{color:#6f42c1}' +
'#spinner{display:none}' +
'</style>' +
'</head>' +
'<body>' +
'<div class="container-fluid py-3 px-4">' +
'<div class="d-flex align-items-center mb-3 gap-2">' +
'<h5 class="mb-0 me-3">📋 イベント管理ダッシュボード</h5>' +
'<button class="btn btn-sm btn-outline-secondary ms-auto" onclick="showLiffLinks()">🔗 LIFFリンク一覧</button>' +
'<div id="spinner" class="spinner-border spinner-border-sm text-primary"></div>' +
'</div>' +

'<ul class="nav nav-tabs mb-3" id="mainTab" role="tablist">' +
'<li class="nav-item"><a class="nav-link active" href="#" id="tab-btn-events" onclick="showTab(\'events\');return false">イベント一覧</a></li>' +
'<li class="nav-item"><a class="nav-link" href="#" id="tab-btn-broadcast" onclick="showTab(\'broadcast\');return false">絞り込み送信</a></li>' +
'<li class="nav-item"><a class="nav-link" href="#" id="tab-btn-members" onclick="showTab(\'members\');return false">会員一覧</a></li>' +
'</ul>' +

'<!-- イベント一覧タブ -->' +
'<div id="tab-events">' +
'<div class="d-flex align-items-center gap-2 mb-2">' +
'<button class="btn btn-success btn-sm" onclick="showNewEventModal()">＋ イベントを新規登録</button>' +
'</div>' +
'<!-- イベント新規作成モーダル -->' +
'<div id="newEventModal" class="card p-3 mb-3" style="display:none;border:2px solid #198754">' +
'<h6 class="mb-3">📋 新しいイベントを登録</h6>' +
'<div class="row g-2">' +
'<div class="col-12"><label class="form-label fw-bold">イベント種別</label>' +
'<select class="form-select" id="ne_type" onchange="onEventTypeChange()">' +
'<option value="オフライン">📍 オフライン</option>' +
'<option value="オンライン">💻 オンライン</option>' +
'</select></div>' +
'<div class="col-12"><label class="form-label fw-bold">イベント名<span class="text-danger">*</span></label>' +
'<input type="text" class="form-control" id="ne_name" placeholder="コーチAレッスン 7月15日"></div>' +
'<!-- オフライン専用 -->' +
'<div id="ne_offline_fields" class="col-12">' +
'<div class="row g-2">' +
'<div class="col-6"><label class="form-label fw-bold">応募開始日</label>' +
'<input type="date" class="form-control" id="ne_opening"><div class="form-text">空欄にするとすぐ表示</div></div>' +
'<div class="col-6"><label class="form-label fw-bold">募集終了日<span class="text-danger" id="ne_closing_req">*</span><span class="text-muted small fw-normal ms-1" id="ne_closing_opt" style="display:none">（任意）</span></label>' +
'<input type="date" class="form-control" id="ne_closing"></div>' +
'<div class="col-6"><label class="form-label fw-bold">開催日<span class="text-danger">*</span></label>' +
'<input type="date" class="form-control" id="ne_date"></div>' +
'<div class="col-6"><label class="form-label fw-bold">開催時間</label>' +
'<div class="d-flex align-items-center gap-1">' +
'<input type="time" class="form-control" id="ne_time_start">' +
'<span class="px-1">〜</span>' +
'<input type="time" class="form-control" id="ne_time_end">' +
'</div></div>' +
'<div class="col-6"><label class="form-label fw-bold">開催場所</label>' +
'<input type="text" class="form-control" id="ne_venue" placeholder="渋谷テニスコート"></div>' +
'</div></div>' +
'<!-- オンライン専用 -->' +
'<div id="ne_online_fields" class="col-12" style="display:none">' +
'<div class="row g-2">' +
'<div class="col-12"><label class="form-label fw-bold">チャンネルURL</label>' +
'<input type="url" class="form-control" id="ne_channel_url" placeholder="https://..."></div>' +
'<div class="col-6"><label class="form-label fw-bold">応募開始日<span class="text-muted small fw-normal ms-1">（任意）</span></label>' +
'<input type="date" class="form-control" id="ne_opening_online"></div>' +
'<div class="col-6"><label class="form-label fw-bold">募集終了日<span class="text-muted small fw-normal ms-1">（任意）</span></label>' +
'<input type="date" class="form-control" id="ne_closing_online"></div>' +
'</div></div>' +
'<!-- 共通 -->' +
'<div class="col-12"><label class="form-label fw-bold">コーチ名</label>' +
'<input type="text" class="form-control" id="ne_coach" placeholder="山田 コーチ"></div>' +
'<div class="col-12"><label class="form-label fw-bold">イベント内容</label>' +
'<textarea class="form-control" id="ne_desc" rows="3" placeholder="イベントの説明・内容を入力"></textarea></div>' +
'</div>' +
'<div class="d-flex gap-2 mt-3 align-items-center">' +
'<button class="btn btn-success" onclick="submitNewEvent()">登録する</button>' +
'<button class="btn btn-outline-secondary" onclick="hideNewEventModal()">キャンセル</button>' +
'<span id="ne_result" class="text-muted small"></span>' +
'</div></div>' +
'<div id="eventList" class="row g-2 mb-3"></div>' +
'<div id="applicantSection" style="display:none">' +
'<div class="d-flex align-items-center gap-2 mb-2 flex-wrap">' +
'<h6 class="mb-0" id="applicantTitle"></h6>' +
'<button class="btn btn-success btn-sm ms-auto" onclick="sendNotifications()">📨 当落通知を送信</button>' +
'<button class="btn btn-outline-secondary btn-sm" onclick="closeApplicants()">✕ 閉じる</button>' +
'</div>' +
'<div class="d-flex gap-2 flex-wrap mb-2 p-2 bg-light rounded">' +
'<button class="btn btn-primary btn-sm" onclick="batchWinLose()">✅ チェックを当選・残りを落選</button>' +
'<button class="btn btn-outline-success btn-sm" onclick="batchSet(\'当選\')">チェックした人を当選</button>' +
'<button class="btn btn-outline-danger btn-sm" onclick="batchSet(\'落選\')">チェックした人を落選</button>' +
'<div class="ms-auto d-flex gap-1">' +
'<button class="btn btn-outline-secondary btn-sm" onclick="toggleAllChecks(true)">全選択</button>' +
'<button class="btn btn-outline-secondary btn-sm" onclick="toggleAllChecks(false)">全解除</button>' +
'</div></div>' +
'<div class="d-flex flex-wrap gap-2 align-items-end mb-2 p-2 bg-white border rounded">' +
'<div><div class="small fw-bold mb-1">レベル</div><select class="form-select form-select-sm" id="afLevel" onchange="filterApplicants()" style="min-width:90px"><option value="">全員</option></select></div>' +
'<div><div class="small fw-bold mb-1">テニス歴</div><select class="form-select form-select-sm" id="afHistory" onchange="filterApplicants()" style="min-width:90px"><option value="">全員</option></select></div>' +
'<div><div class="small fw-bold mb-1">頻度</div><select class="form-select form-select-sm" id="afFreq" onchange="filterApplicants()" style="min-width:110px"><option value="">全員</option></select></div>' +
'<div><div class="small fw-bold mb-1">性別</div><select class="form-select form-select-sm" id="afGender" onchange="filterApplicants()" style="min-width:80px"><option value="">全員</option><option value="男性">男性</option><option value="女性">女性</option></select></div>' +
'<div><div class="small fw-bold mb-1">年齢</div><div class="d-flex align-items-center gap-1"><input type="number" class="form-control form-control-sm" id="afAgeMin" placeholder="下" style="width:56px" onchange="filterApplicants()"><span class="small">〜</span><input type="number" class="form-control form-control-sm" id="afAgeMax" placeholder="上" style="width:56px" onchange="filterApplicants()"></div></div>' +
'<div><div class="small fw-bold mb-1">コーチ認知</div><select class="form-select form-select-sm" id="afCoach" onchange="filterApplicants()" style="min-width:130px"><option value="">全員</option></select></div>' +
'<div class="align-self-end"><button class="btn btn-outline-secondary btn-sm" onclick="resetApplicantFilters()">リセット</button></div>' +
'<div class="align-self-end text-muted small ms-1" id="afCount"></div>' +
'</div>' +
'<div class="table-responsive">' +
'<table class="table table-sm table-hover bg-white">' +
'<thead class="table-light">' +
'<tr><th style="width:36px"><input type="checkbox" id="chkAll" onchange="toggleAllChecks(this.checked)"></th>' +
'<th>名前</th><th>年齢</th><th>性別</th><th>レベル</th><th>テニス歴</th><th>頻度</th><th>応募日時</th><th>当落</th><th>当選回数</th><th>コーチ認知</th><th>通知</th><th>操作</th></tr>' +
'</thead>' +
'<tbody id="applicantBody"></tbody>' +
'</table>' +
'</div>' +
'</div>' +
'</div>' +

'<!-- 絞り込み送信タブ -->' +
'<div id="tab-broadcast" style="display:none">' +
'<div class="row g-3">' +
'<div class="col-md-4">' +
'<div class="card p-3">' +
'<div class="mb-2">' +
'<label class="form-label fw-bold">対象イベント</label>' +
'<select class="form-select" id="bcastEvent"><option value="">（選択してください）</option></select>' +
'</div>' +
'<div class="mb-2">' +
'<label class="form-label fw-bold">ステータスで絞り込み</label>' +
'<select class="form-select" id="bcastStatus">' +
'<option value="未応募">未応募</option>' +
'<option value="応募済み">応募済み</option>' +
'<option value="当選">当選</option>' +
'<option value="落選">落選</option>' +
'</select>' +
'</div>' +
'<button class="btn btn-outline-primary w-100 mb-2" onclick="loadBcastUsers()">対象者を確認</button>' +
'<div id="bcastUserList" class="border rounded p-2 bg-white" style="min-height:60px;max-height:200px;overflow-y:auto;font-size:12px;"></div>' +
'</div>' +
'</div>' +
'<div class="col-md-8">' +
'<div class="card p-3">' +
'<label class="form-label fw-bold">メッセージ</label>' +
'<textarea class="form-control mb-2" id="bcastMessage" rows="10" placeholder="送信するメッセージを入力してください"></textarea>' +
'<div class="d-flex align-items-center gap-2">' +
'<button class="btn btn-primary" onclick="execBroadcast()">📢 送信する</button>' +
'<span id="bcastResult" class="text-muted small"></span>' +
'</div>' +
'</div>' +
'</div>' +
'</div>' +
'</div>' +

'<!-- 会員一覧タブ -->' +
'<div id="tab-members" style="display:none">' +
'<div class="row g-3">' +
'<div class="col-md-4">' +
'<div class="card p-3">' +
'<div class="mb-2"><label class="form-label fw-bold">性別</label>' +
'<select class="form-select" id="mFilterGender"><option value="">全員</option><option value="男性">男性</option><option value="女性">女性</option></select></div>' +
'<div class="mb-2"><label class="form-label fw-bold">年齢</label>' +
'<div class="d-flex align-items-center gap-1">' +
'<input type="number" class="form-control" id="mFilterAgeMin" placeholder="下限" min="0">' +
'<span class="px-1">〜</span>' +
'<input type="number" class="form-control" id="mFilterAgeMax" placeholder="上限" min="0">' +
'</div></div>' +
'<div class="mb-2"><label class="form-label fw-bold">テニスレベル</label>' +
'<select class="form-select" id="mFilterLevel"><option value="">全員</option></select></div>' +
'<button class="btn btn-outline-primary w-100 mb-2" onclick="filterMembers()">対象者を確認</button>' +
'<div id="mTargetList" class="border rounded p-2 bg-white" style="min-height:60px;max-height:200px;overflow-y:auto;font-size:12px;"></div>' +
'</div></div>' +
'<div class="col-md-8">' +
'<div class="card p-3 mb-2">' +
'<label class="form-label fw-bold">メッセージ</label>' +
'<textarea class="form-control mb-2" id="mMessage" rows="6" placeholder="送信するメッセージを入力してください"></textarea>' +
'<div class="d-flex align-items-center gap-2">' +
'<button class="btn btn-primary" onclick="execMemberBroadcast()">📢 送信する</button>' +
'<span id="mResult" class="text-muted small"></span>' +
'</div></div>' +
'<div class="table-responsive">' +
'<table class="table table-sm table-hover bg-white">' +
'<thead class="table-light"><tr><th>名前</th><th>フリガナ</th><th>年齢</th><th>性別</th><th>電話番号</th><th>メール</th><th>レベル</th><th>テニス歴</th><th>頻度</th><th>地域</th><th>環境</th><th>登録日</th><th></th></tr></thead>' +
'<tbody id="membersBody"></tbody>' +
'</table></div></div></div></div>' +

'<!-- 会員応募履歴モーダル -->' +
'<div id="memberHistoryOverlay" style="display:none;position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.5);z-index:9999" onclick="if(event.target===this)closeMemberHistory()">' +
'<div style="background:#fff;max-width:580px;margin:60px auto;border-radius:8px;padding:20px;max-height:80vh;overflow-y:auto">' +
'<div class="d-flex justify-content-between align-items-center mb-3">' +
'<h6 class="mb-0" id="mhTitle"></h6>' +
'<button class="btn btn-sm btn-outline-secondary" onclick="closeMemberHistory()">✕ 閉じる</button>' +
'</div>' +
'<div id="mhBody"></div>' +
'</div></div>' +

'</div>' +

'<script>' +
'var eventsData=[];' +
'var currentEvent=null;' +
'var bcastUserIds=[];' +
'var membersData=[];' +
'var mTargetIds=[];' +
'var membersLoaded=false;' +
'var allApplicantsData=[];' +
'var dashConfig={};' +

'window.onload=function(){' +
'google.script.run' +
'.withSuccessHandler(function(c){dashConfig=c;loadEvents();})' +
'.withFailureHandler(function(){loadEvents();})' +
'.getDashboardConfig();' +
'};' +

'function showNewEventModal(){document.getElementById("newEventModal").style.display="";}' +
'function hideNewEventModal(){' +
'document.getElementById("newEventModal").style.display="none";' +
'document.getElementById("ne_result").textContent="";' +
'document.getElementById("ne_type").value="オフライン";' +
'onEventTypeChange();' +
'["ne_name","ne_opening","ne_closing","ne_date","ne_venue","ne_coach","ne_desc","ne_channel_url","ne_opening_online","ne_closing_online"].forEach(function(id){var el=document.getElementById(id);if(el)el.value="";});' +
'["ne_time_start","ne_time_end"].forEach(function(id){var el=document.getElementById(id);if(el)el.value="";});' +
'}' +
'function onEventTypeChange(){' +
'var isOnline=document.getElementById("ne_type").value==="オンライン";' +
'document.getElementById("ne_offline_fields").style.display=isOnline?"none":"";' +
'document.getElementById("ne_online_fields").style.display=isOnline?"":"none";' +
'}' +
'function submitNewEvent(){' +
'var evType=document.getElementById("ne_type").value;' +
'var isOnline=evType==="オンライン";' +
'var name=document.getElementById("ne_name").value.trim();' +
'var coach=document.getElementById("ne_coach").value.trim();' +
'var desc=document.getElementById("ne_desc").value.trim();' +
'var date="",closing="",opening="",time="",venue="",channelUrl="";' +
'if(isOnline){' +
'channelUrl=(document.getElementById("ne_channel_url")||{}).value||"";' +
'opening=((document.getElementById("ne_opening_online")||{}).value||"").replace(/-/g,"/");' +
'closing=((document.getElementById("ne_closing_online")||{}).value||"").replace(/-/g,"/");' +
'}else{' +
'date=document.getElementById("ne_date").value.replace(/-/g,"/");' +
'closing=document.getElementById("ne_closing").value.replace(/-/g,"/");' +
'opening=document.getElementById("ne_opening").value.replace(/-/g,"/");' +
'var tS=document.getElementById("ne_time_start").value;' +
'var tE=document.getElementById("ne_time_end").value;' +
'time=tS&&tE?tS+"〜"+tE:(tS||"");' +
'venue=document.getElementById("ne_venue").value.trim();' +
'}' +
'if(!name){alert("イベント名は必須です。");return;}' +
'if(!isOnline&&(!date||!closing)){alert("オフラインイベントには開催日・募集終了日が必須です。");return;}' +
'var res=document.getElementById("ne_result");res.textContent="登録中...";' +
'google.script.run' +
'.withSuccessHandler(function(r){' +
'if(r.success){res.textContent="✅ 登録完了";setTimeout(function(){hideNewEventModal();loadEvents();},1500);}' +
'else{res.textContent="❌ "+r.error;}' +
'})' +
'.withFailureHandler(function(e){res.textContent="❌ "+e.message;})' +
'.createNewEvent({name:name,eventDate:date,closingDate:closing,openingDate:opening,eventTime:time,venue:venue,coachName:coach,description:desc,channelUrl:channelUrl,eventType:evType});' +
'}' +

'function showTab(t){' +
'document.getElementById("tab-events").style.display=t==="events"?"":"none";' +
'document.getElementById("tab-broadcast").style.display=t==="broadcast"?"":"none";' +
'document.getElementById("tab-members").style.display=t==="members"?"":"none";' +
'document.getElementById("tab-btn-events").classList.toggle("active",t==="events");' +
'document.getElementById("tab-btn-broadcast").classList.toggle("active",t==="broadcast");' +
'document.getElementById("tab-btn-members").classList.toggle("active",t==="members");' +
'if(t==="members"&&!membersLoaded){membersLoaded=true;loadMembers();}' +
'}' +

'function spin(on){document.getElementById("spinner").style.display=on?"":"none";}' +

'function loadEvents(){' +
'spin(true);' +
'google.script.run.withSuccessHandler(function(ev){spin(false);renderEvents(ev);}).withFailureHandler(function(e){spin(false);alert("エラー: "+e.message);}).getEventsData();' +
'}' +

'function renderEvents(events){' +
'eventsData=events;' +
'var el=document.getElementById("eventList");' +
'el.innerHTML="";' +
'var sel=document.getElementById("bcastEvent");' +
'sel.innerHTML="<option value=\'\'>（選択してください）</option>";' +
'if(!events||events.length===0){el.innerHTML="<div class=\'col\'><p class=\'text-muted\'>設定シートにイベントがありません。</p></div>";return;}' +
'events.forEach(function(ev,idx){' +
'var badge=ev.pendingCount>0?"<span class=\'badge badge-pending ms-1\'>"+ev.pendingCount+"件未送信</span>":"";' +
'var div=document.createElement("div");' +
'div.className="col-md-4 col-lg-3";' +
'var detail=(ev.coachName?"<div class=\'text-muted small\'>👤 "+ev.coachName+"</div>":"")+' +
'(ev.eventTime?"<div class=\'text-muted small\'>🕐 "+ev.eventTime+"</div>":"")+' +
'(ev.venue?"<div class=\'text-muted small\'>📍 "+ev.venue+"</div>":"");' +
'var isStopped=ev.status==="停止";' +
'var statusBadge=isStopped?"<span class=\'badge bg-danger ms-1 align-middle\' style=\'font-size:10px\'>停止中</span>":"";' +
'var toggleBtn="<button class=\'btn btn-sm "+(isStopped?"btn-outline-success":"btn-outline-danger")+" py-0 px-2\' onclick=\'doToggleStatus("+idx+",event)\'>"+(isStopped?"▶ 再開":"⏸ 停止")+"</button>";' +
'var copyBtn=dashConfig.liffId?"<button class=\'btn btn-sm btn-outline-success py-0 px-2\' onclick=\'copyLiffUrl("+idx+",event)\'>🔗 リンクコピー</button>":"";' +
'var btns="<div class=\'mt-2 d-flex gap-1\'>"+toggleBtn+(copyBtn?copyBtn:"")+"</div>";' +
'div.innerHTML="<div class=\'card event-card h-100 border\' style=\'"+(isStopped?"opacity:0.55":"")+"\'onclick=\'selectEvent("+idx+")\'>"+' +
'"<div class=\'card-body py-2\'>"+' +
'"<div class=\'fw-bold mb-1\'>"+ev.name+badge+statusBadge+"</div>"+' +
'"<div class=\'text-muted small\'>"+(ev.openingDate?"応募開始: "+ev.openingDate+" / ":"")+(ev.eventDate?"開催: "+ev.eventDate+" / ":"")+"締切: "+(ev.closingDate||"常時")+"</div>"+' +
'detail+' +
'"<div class=\'small mt-1\'>応募: "+ev.appCount+"名 ／ 当選: "+ev.winCount+"名 ／ 落選: "+ev.loseCount+"名</div>"+' +
'btns+' +
'"</div></div>";' +
'el.appendChild(div);' +
'var opt=document.createElement("option");' +
'opt.value=idx;opt.textContent=ev.name;sel.appendChild(opt);' +
'});' +
'}' +

'function showLiffLinks(){' +
'if(!dashConfig.liffId){alert("LIFF_IDが未設定です。スクリプトプロパティを確認してください。");return;}' +
'var base="https://liff.line.me/"+dashConfig.liffId;' +
'var links=[' +
'{label:"📝 プロフィール登録",desc:"初回登録専用。既存会員が開いても登録フォームが表示されます。",url:base+"?page=register"},' +
'{label:"✏️ プロフィール修正",desc:"既存会員向け。プロフィール編集欄が開いた状態で表示されます。",url:base+"?page=profile"},' +
'{label:"📍 オフライン応募",desc:"オフラインイベントのみ表示されます。",url:base+"?type=offline"},' +
'{label:"💻 オンライン応募",desc:"オンラインイベントのみ表示されます。",url:base+"?type=online"},' +
'];' +
'var html=links.map(function(l){' +
'return "<div class=\'mb-3\'><div class=\'fw-bold mb-1\'>"+l.label+"</div><div class=\'text-muted small mb-1\'>"+l.desc+"</div><div class=\'d-flex gap-2\'><input type=\'text\' class=\'form-control form-control-sm\' value=\'"+l.url+"\' readonly onclick=\'this.select()\'><button class=\'btn btn-sm btn-outline-primary flex-shrink-0\' onclick=\'cpLink(this,\""+encodeURIComponent(l.url)+"\")\'>コピー</button></div></div>";' +
'}).join("");' +
'document.getElementById("liffLinksBody").innerHTML=html;' +
'document.getElementById("liffLinksModal").style.display="flex";' +
'}' +
'function hideLiffLinks(){document.getElementById("liffLinksModal").style.display="none";}' +
'function cpLink(btn,encodedUrl){' +
'var url=decodeURIComponent(encodedUrl);' +
'if(navigator.clipboard){navigator.clipboard.writeText(url).then(function(){var t=btn.textContent;btn.textContent="✅ コピー済";setTimeout(function(){btn.textContent=t;},1500);});}' +
'else{prompt("コピーしてください:",url);}' +
'}' +

'function doToggleStatus(idx,e){' +
'e.stopPropagation();' +
'var ev=eventsData[idx];' +
'var isStopped=ev.status==="停止";' +
'var msg=isStopped?"「"+ev.name+"」を再開しますか？\n\nLIFFフォームに再表示されます。":"「"+ev.name+"」を停止しますか？\n\nLIFFフォームから非表示になります。";' +
'if(!confirm(msg))return;' +
'google.script.run' +
'.withSuccessHandler(function(r){if(r.success){loadEvents();}else{alert("エラー: "+r.error);}})' +
'.withFailureHandler(function(e){alert("エラー: "+e.message);})' +
'.toggleEventStatus(ev.appSheetName);' +
'}' +

'function copyLiffUrl(idx,e){' +
'e.stopPropagation();' +
'if(!dashConfig.liffId){alert("LIFF_IDが未設定です。スクリプトプロパティを確認してください。");return;}' +
'var url="https://liff.line.me/"+dashConfig.liffId+"?event="+encodeURIComponent(eventsData[idx].name);' +
'if(navigator.clipboard){navigator.clipboard.writeText(url).then(function(){alert("コピーしました！\\n\\n"+url);});}' +
'else{prompt("このURLをコピーしてください:",url);}' +
'}' +

'function selectEvent(idx){' +
'currentEvent=eventsData[idx];' +
'document.getElementById("applicantSection").style.display="";' +
'document.getElementById("applicantTitle").textContent=currentEvent.name+" — 応募者一覧";' +
'loadApplicants(currentEvent.appSheetName,currentEvent.resultSheetName);' +
'}' +

'function closeApplicants(){' +
'document.getElementById("applicantSection").style.display="none";' +
'currentEvent=null;' +
'}' +

'function loadApplicants(appSheetName,resultSheetName){' +
'spin(true);' +
'google.script.run.withSuccessHandler(function(ap){spin(false);renderApplicants(ap);}).withFailureHandler(function(e){spin(false);alert("エラー: "+e.message);}).getApplicants(appSheetName,resultSheetName);' +
'}' +

'function renderApplicants(applicants){' +
'allApplicantsData=applicants;' +
'var lvls=[...new Set(applicants.map(function(a){return a.tennisLevel;}).filter(Boolean))].sort();' +
'var hists=[...new Set(applicants.map(function(a){return a.tennisHistory;}).filter(Boolean))];' +
'var freqs=[...new Set(applicants.map(function(a){return a.tennisFreq;}).filter(Boolean))];' +
'var coaches=[...new Set(applicants.map(function(a){return a.coachKnowledge;}).filter(Boolean))];' +
'var fillSel=function(id,vals){var s=document.getElementById(id);s.innerHTML="<option value=\'\'>全員</option>"+vals.map(function(v){return"<option value=\'"+v+"\'>"+v+"</option>";}).join("");};' +
'fillSel("afLevel",lvls);fillSel("afHistory",hists);fillSel("afFreq",freqs);fillSel("afCoach",coaches);' +
'["afGender","afAgeMin","afAgeMax"].forEach(function(id){document.getElementById(id).value="";});' +
'document.getElementById("afCount").textContent=applicants.length+"名";' +
'renderApplicantsTable(applicants);' +
'}' +

'function filterApplicants(){' +
'var lv=document.getElementById("afLevel").value;' +
'var hi=document.getElementById("afHistory").value;' +
'var fr=document.getElementById("afFreq").value;' +
'var gn=document.getElementById("afGender").value;' +
'var co=document.getElementById("afCoach").value;' +
'var amin=parseInt(document.getElementById("afAgeMin").value)||0;' +
'var amax=parseInt(document.getElementById("afAgeMax").value)||999;' +
'var filtered=allApplicantsData.filter(function(a){' +
'if(lv&&a.tennisLevel!==lv)return false;' +
'if(hi&&a.tennisHistory!==hi)return false;' +
'if(fr&&a.tennisFreq!==fr)return false;' +
'if(gn&&a.gender!==gn)return false;' +
'if(co&&a.coachKnowledge!==co)return false;' +
'if(a.age){var ag=parseInt(a.age);if(!isNaN(ag)&&(ag<amin||ag>amax))return false;}' +
'return true;' +
'});' +
'document.getElementById("afCount").textContent=filtered.length+"/"+allApplicantsData.length+"名";' +
'renderApplicantsTable(filtered);' +
'}' +

'function resetApplicantFilters(){' +
'["afLevel","afHistory","afFreq","afGender","afCoach","afAgeMin","afAgeMax"].forEach(function(id){document.getElementById(id).value="";});' +
'document.getElementById("afCount").textContent=allApplicantsData.length+"名";' +
'renderApplicantsTable(allApplicantsData);' +
'}' +

'function renderApplicantsTable(applicants){' +
'var tbody=document.getElementById("applicantBody");' +
'tbody.innerHTML="";' +
'if(!applicants||applicants.length===0){' +
'tbody.innerHTML="<tr><td colspan=\'13\' class=\'text-center text-muted\'>応募者がいません。</td></tr>";return;' +
'}' +
'applicants.forEach(function(ap){' +
'var cls=ap.result?"status-"+ap.result:"status-未処理";' +
'var sentBadge=ap.sent==="済"?"<span class=\'badge bg-success\'>送信済</span>":"<span class=\'badge bg-secondary\'>未送信</span>";' +
'var confBadge=ap.result==="当選"&&ap.confirmation?' +
  '(ap.confirmation==="確認済"?"<br><small class=\'text-success\'>✓ 確認済</small>"' +
  ':ap.confirmation==="確認待ち"?"<br><small class=\'text-warning\'>⏳ 確認待ち</small>":""):"";' +
'var tr=document.createElement("tr");' +
'tr.innerHTML="<td><input type=\'checkbox\' class=\'row-check\' data-userid=\'"+ap.userId+"\'></td>"+' +
'"<td>"+ap.name+"</td>"+' +
'"<td class=\'text-center small\'>"+ap.age+"</td>"+' +
'"<td class=\'text-center small\'>"+ap.gender+"</td>"+' +
'"<td class=\'text-muted small\'>"+ap.tennisLevel+"</td>"+' +
'"<td class=\'text-muted small\'>"+ap.tennisHistory+"</td>"+' +
'"<td class=\'text-muted small\'>"+ap.tennisFreq+"</td>"+' +
'"<td class=\'text-muted small\'>"+ap.appliedAt+"</td>"+' +
'"<td class=\'"+cls+"\'>"+( ap.result||"未処理")+confBadge+"</td>"+' +
'"<td class=\'text-center small\'>"+(ap.winCount||0)+"</td>"+' +
'"<td class=\'text-muted small\'>"+ap.coachKnowledge+"</td>"+' +
'"<td>"+sentBadge+"</td>"+' +
'"<td>"+' +
'"<button class=\'btn "+(ap.result==="当選"?"btn-success":"btn-outline-success")+" btn-sm py-0 me-1\' onclick=\'setResult(\\\""+ap.userId+"\\\",\\\"当選\\\",this)\'>当選</button>"+' +
'"<button class=\'btn "+(ap.result==="落選"?"btn-danger":"btn-outline-danger")+" btn-sm py-0\' onclick=\'setResult(\\\""+ap.userId+"\\\",\\\"落選\\\",this)\'>落選</button>"+' +
'"</td>";' +
'tbody.appendChild(tr);' +
'});' +
'}' +

'function toggleAllChecks(on){' +
'document.querySelectorAll("#applicantBody .row-check").forEach(function(cb){cb.checked=on;});' +
'var hdr=document.getElementById("chkAll");if(hdr)hdr.checked=on;' +
'}' +

'function batchWinLose(){' +
'if(!currentEvent)return;' +
'var rows=Array.from(document.querySelectorAll("#applicantBody tr"));' +
'var results=rows.map(function(tr){var cb=tr.querySelector(".row-check");return cb?{userId:cb.dataset.userid,result:cb.checked?"当選":"落選"}:null;}).filter(Boolean);' +
'if(results.length===0){alert("応募者がいません。");return;}' +
'var wc=results.filter(function(r){return r.result==="当選";}).length;' +
'var lc=results.filter(function(r){return r.result==="落選";}).length;' +
'if(!confirm("当選: "+wc+"名 / 落選: "+lc+"名\\nこの内容で確定しますか？"))return;' +
'spin(true);' +
'google.script.run' +
'.withSuccessHandler(function(res){spin(false);if(res.success){loadApplicants(currentEvent.appSheetName,currentEvent.resultSheetName);loadEvents();}else{alert("エラー: "+res.error);}})' +
'.withFailureHandler(function(e){spin(false);alert("エラー: "+e.message);})' +
'.setResultsBatch(currentEvent.resultSheetName,results);' +
'}' +

'function batchSet(result){' +
'if(!currentEvent)return;' +
'var checked=Array.from(document.querySelectorAll("#applicantBody .row-check:checked"));' +
'if(checked.length===0){alert("対象者を選択してください。");return;}' +
'var results=checked.map(function(cb){return{userId:cb.dataset.userid,result:result};});' +
'if(!confirm(checked.length+"名を"+result+"にしますか？"))return;' +
'spin(true);' +
'google.script.run' +
'.withSuccessHandler(function(res){spin(false);if(res.success){loadApplicants(currentEvent.appSheetName,currentEvent.resultSheetName);loadEvents();}else{alert("エラー: "+res.error);}})' +
'.withFailureHandler(function(e){spin(false);alert("エラー: "+e.message);})' +
'.setResultsBatch(currentEvent.resultSheetName,results);' +
'}' +

'function setResult(userId,result,btn){' +
'if(!currentEvent)return;' +
'btn.disabled=true;' +
'var td=btn.closest("td");' +
'td.querySelectorAll("button").forEach(function(b){b.classList.remove("btn-success","btn-danger");b.classList.add(b.textContent==="当選"?"btn-outline-success":"btn-outline-danger");});' +
'btn.classList.remove(result==="当選"?"btn-outline-success":"btn-outline-danger");' +
'btn.classList.add(result==="当選"?"btn-success":"btn-danger");' +
'google.script.run' +
'.withSuccessHandler(function(res){btn.disabled=false;if(res.success){loadApplicants(currentEvent.appSheetName,currentEvent.resultSheetName);}else{btn.disabled=false;alert("エラー: "+res.error);}})' +
'.withFailureHandler(function(e){btn.disabled=false;alert("エラー: "+e.message);})' +
'.setResult(currentEvent.resultSheetName,userId,result);' +
'}' +

'function sendNotifications(){' +
'if(!currentEvent)return;' +
'if(!confirm("未送信の当落通知を一括送信します。よろしいですか？"))return;' +
'spin(true);' +
'google.script.run' +
'.withSuccessHandler(function(res){spin(false);if(res.success){alert("送信完了\\n当選: "+res.winCount+"名 / 落選: "+res.loseCount+"名");loadApplicants(currentEvent.appSheetName,currentEvent.resultSheetName);loadEvents();}else{alert("エラー: "+res.error);}})' +
'.withFailureHandler(function(e){spin(false);alert("エラー: "+e.message);})' +
'.sendResultsFromDashboard(currentEvent.resultSheetName);' +
'}' +

'function loadBcastUsers(){' +
'var idxVal=document.getElementById("bcastEvent").value;' +
'var status=document.getElementById("bcastStatus").value;' +
'if(!idxVal){alert("イベントを選択してください。");return;}' +
'var ev=eventsData[parseInt(idxVal)];' +
'spin(true);' +
'google.script.run' +
'.withSuccessHandler(function(users){' +
'spin(false);bcastUserIds=users.map(function(u){return u.userId;});' +
'var listEl=document.getElementById("bcastUserList");' +
'if(!users||users.length===0){listEl.textContent="対象者がいません。";}' +
'else{listEl.innerHTML="<div class=\'fw-bold mb-1\'>"+users.length+"名</div>"+users.map(function(u){return "<div>"+u.name+"</div>";}).join("");}' +
'})' +
'.withFailureHandler(function(e){spin(false);alert("エラー: "+e.message);})' +
'.getFilteredUsers(ev.appSheetName,ev.resultSheetName,status);' +
'}' +

'function execBroadcast(){' +
'var message=document.getElementById("bcastMessage").value.trim();' +
'if(!message){alert("メッセージを入力してください。");return;}' +
'if(bcastUserIds.length===0){alert("先に「対象者を確認」ボタンを押してください。");return;}' +
'if(!confirm(bcastUserIds.length+"名にメッセージを送信します。よろしいですか？"))return;' +
'spin(true);' +
'var resultEl=document.getElementById("bcastResult");resultEl.textContent="";' +
'google.script.run' +
'.withSuccessHandler(function(res){spin(false);if(res.success){resultEl.textContent="✅ "+res.count+"名に送信しました";document.getElementById("bcastMessage").value="";bcastUserIds=[];document.getElementById("bcastUserList").innerHTML="";}else{alert("エラー: "+res.error);}})' +
'.withFailureHandler(function(e){spin(false);alert("エラー: "+e.message);})' +
'.sendBroadcast(bcastUserIds,message);' +
'}' +

'function loadMembers(){' +
'spin(true);' +
'google.script.run' +
'.withSuccessHandler(function(data){spin(false);initMembersTab(data);})' +
'.withFailureHandler(function(e){spin(false);alert("エラー: "+e.message);})' +
'.getMembersData();' +
'}' +

'function initMembersTab(data){' +
'membersData=data;' +
'var levels=[...new Set(data.map(function(m){return m.tennisLevel;}).filter(function(v){return v;}))].sort();' +
'var sel=document.getElementById("mFilterLevel");' +
'sel.innerHTML="<option value=\'\'>全員</option>"+levels.map(function(l){return"<option value=\'"+l+"\'>"+l+"</option>";}).join("");' +
'renderMembersTable(data);' +
'}' +

'function renderMembersTable(members){' +
'var tbody=document.getElementById("membersBody");' +
'if(!members||members.length===0){tbody.innerHTML="<tr><td colspan=\'13\' class=\'text-center text-muted\'>会員データがありません。</td></tr>";return;}' +
'tbody.innerHTML=members.map(function(m){' +
'var uid=m.userId.replace(/"/g,"&quot;");var nm=m.name.replace(/"/g,"&quot;");' +
'return"<tr>"+' +
'"<td>"+m.name+"</td>"+' +
'"<td class=\'text-muted small\'>"+m.furigana+"</td>"+' +
'"<td class=\'text-center\'>"+m.age+"</td>"+' +
'"<td>"+m.gender+"</td>"+' +
'"<td class=\'small\'>"+m.phone+"</td>"+' +
'"<td class=\'text-muted small\'>"+m.email+"</td>"+' +
'"<td class=\'small\'>"+m.tennisLevel+"</td>"+' +
'"<td class=\'text-muted small\'>"+m.tennisHistory+"</td>"+' +
'"<td class=\'text-muted small\'>"+m.tennisFreq+"</td>"+' +
'"<td class=\'text-muted small\'>"+m.tennisArea+"</td>"+' +
'"<td class=\'text-muted small\'>"+m.tennisEnv+"</td>"+' +
'"<td class=\'text-muted small\'>"+m.registeredAt+"</td>"+' +
'"<td><button class=\'btn btn-outline-primary btn-sm py-0\' onclick=\'showMemberHistory(\\\""+uid+"\\\",\\\""+nm+"\\\")\'>履歴</button></td>"+' +
'"</tr>";' +
'}).join("");' +
'}' +

'function filterMembers(){' +
'var gender=document.getElementById("mFilterGender").value;' +
'var ageMin=parseInt(document.getElementById("mFilterAgeMin").value)||0;' +
'var ageMax=parseInt(document.getElementById("mFilterAgeMax").value)||999;' +
'var level=document.getElementById("mFilterLevel").value;' +
'var filtered=membersData.filter(function(m){' +
'if(gender&&m.gender!==gender)return false;' +
'if(m.age){var a=parseInt(m.age);if(!isNaN(a)&&(a<ageMin||a>ageMax))return false;}' +
'if(level&&m.tennisLevel!==level)return false;' +
'return true;' +
'});' +
'mTargetIds=filtered.map(function(m){return m.userId;});' +
'var listEl=document.getElementById("mTargetList");' +
'if(filtered.length===0){listEl.textContent="対象者がいません。";}' +
'else{listEl.innerHTML="<div class=\'fw-bold mb-1\'>"+filtered.length+"名</div>"+filtered.map(function(m){return"<div>"+m.name+"</div>";}).join("");}' +
'renderMembersTable(filtered);' +
'}' +

'function execMemberBroadcast(){' +
'var message=document.getElementById("mMessage").value.trim();' +
'if(!message){alert("メッセージを入力してください。");return;}' +
'if(mTargetIds.length===0){alert("先に「対象者を確認」ボタンを押してください。");return;}' +
'if(!confirm(mTargetIds.length+"名にメッセージを送信します。よろしいですか？"))return;' +
'spin(true);' +
'var resultEl=document.getElementById("mResult");resultEl.textContent="";' +
'google.script.run' +
'.withSuccessHandler(function(res){spin(false);if(res.success){resultEl.textContent="✅ "+res.count+"名に送信しました";document.getElementById("mMessage").value="";mTargetIds=[];document.getElementById("mTargetList").innerHTML="";}else{alert("エラー: "+res.error);}})' +
'.withFailureHandler(function(e){spin(false);alert("エラー: "+e.message);})' +
'.sendBroadcast(mTargetIds,message);' +
'}' +
'function showMemberHistory(userId,name){' +
'document.getElementById("mhTitle").textContent=name+" の応募履歴";' +
'document.getElementById("mhBody").innerHTML="<div class=\'text-muted small\'>読み込み中...</div>";' +
'document.getElementById("memberHistoryOverlay").style.display="";' +
'google.script.run' +
'.withSuccessHandler(function(res){' +
'if(!res.success){document.getElementById("mhBody").innerHTML="<p class=\'text-danger\'>エラー: "+res.error+"</p>";return;}' +
'if(res.history.length===0){document.getElementById("mhBody").innerHTML="<p class=\'text-muted\'>応募履歴はありません。</p>";return;}' +
'var trs=res.history.map(function(h){' +
'var cls=h.result==="当選"?"text-success fw-bold":h.result==="落選"?"text-danger":"text-muted";' +
'return"<tr><td>"+h.eventName+"</td><td class=\'text-muted small\'>"+h.eventDate+"</td><td class=\'text-muted small\'>"+h.appliedAt+"</td><td class=\'"+cls+"\'>"+h.result+"</td></tr>";' +
'}).join("");' +
'document.getElementById("mhBody").innerHTML=' +
'"<table class=\'table table-sm\'>"+' +
'"<thead class=\'table-light\'><tr><th>イベント名</th><th>開催日</th><th>応募日</th><th>結果</th></tr></thead>"+' +
'"<tbody>"+trs+"</tbody></table>";' +
'})' +
'.withFailureHandler(function(e){document.getElementById("mhBody").innerHTML="<p class=\'text-danger\'>エラー: "+e.message+"</p>";})' +
'.getMemberHistory(userId);' +
'}' +
'function closeMemberHistory(){document.getElementById("memberHistoryOverlay").style.display="none";}' +

'</script>' +
'<!-- LIFFリンク一覧モーダル -->' +
'<div id="liffLinksModal" style="display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.5);z-index:2000;align-items:center;justify-content:center;padding:16px">' +
'<div class="card p-4" style="max-width:600px;width:100%">' +
'<div class="d-flex justify-content-between align-items-center mb-3">' +
'<h6 class="mb-0">🔗 LIFFリンク一覧</h6>' +
'<button class="btn btn-sm btn-outline-secondary" onclick="hideLiffLinks()">✕ 閉じる</button>' +
'</div>' +
'<p class="text-muted small mb-3">各リンクをコピーしてLINEで配布してください。</p>' +
'<div id="liffLinksBody"></div>' +
'</div></div>' +
'</body></html>';
}

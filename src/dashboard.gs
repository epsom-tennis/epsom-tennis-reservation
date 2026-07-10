// 管理ダッシュボード + LIFF向けJSON API
// GASのWebアプリとしてデプロイ（実行ユーザー：自分、アクセス権：全員）
function doGet(e) {
  const page   = e && e.parameter && e.parameter.page;
  const action = e && e.parameter && e.parameter.action;
  const token  = e && e.parameter && e.parameter.token;

  // LIFF向けJSON API（GitHub PagesのHTMLからfetch()で呼び出す）
  if (action === 'getEvents') {
    const userId = e && e.parameter && e.parameter.userId;
    const invite = e && e.parameter && e.parameter.invite;
    return liffApiResponse(getLiffEventsJson(userId, invite));
  }
  if (action === 'getMember') {
    const userId = e && e.parameter && e.parameter.userId;
    return liffApiResponse(getMemberData(userId));
  }
  if (action === 'getTerms') {
    return liffApiResponse(getTermsContent());
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
// accessCodeが渡された場合、限定公開イベントのうちコードが一致するものだけ通常一覧に含める
function getLiffEventsJson(userId, accessCode) {
  try {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const filtered = getAllEvents()
      .filter(ev => ev.status !== '停止' && (!ev.openingDate || ev.openingDate <= today) && (!ev.closingDate || ev.closingDate >= today))
      .filter(ev => !ev.isRestricted || (accessCode && findReferralCode_(ev.name, accessCode)));

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

    return filtered.map(ev => {
      const isTournament = ev.eventType === '大会';
      // 先着受付終了日時を過ぎている場合は抽選待ち期間のため、満員系のバッジは出さない（人数上限を設けないため）
      const isLotteryPhase = !!(ev.firstComeDeadlineAt && new Date() > ev.firstComeDeadlineAt);
      let capacityStatus = '';
      if (isTournament && ev.capacity > 0 && !isLotteryPhase) {
        const referralMatch = accessCode ? findReferralCode_(ev.name, accessCode) : null;
        if (referralMatch && referralMatch.maxCount > 0) {
          // 上限件数付きの紹介コード：人数ではなく「1人でもペアでも1応募＝1件」の残り件数で判定する（ペア枠終了の概念は無い）
          const remaining = referralMatch.maxCount - countReferralCodeUsage_(ev.resultSheetName, referralMatch.code);
          if (remaining <= 0) capacityStatus = 'full';
          else if (remaining / referralMatch.maxCount <= 0.50) capacityStatus = 'low';
          else capacityStatus = 'normal';
        } else {
          // 上限なしの紹介コードを持つ人には大会全体の残り枠を、一般の人には紹介コードの上限件数合計×2（全件ペアの場合）を除いた枠を見せる
          const totalReferralReserved = referralMatch ? 0 : getReferralCodesForEvent_(ev.name).reduce((sum, c) => sum + c.maxCount * 2, 0);
          const effectiveCapacity = ev.capacity - totalReferralReserved;
          const current = countTournamentParticipants_(ev.resultSheetName);
          const remaining = effectiveCapacity - current;
          if (remaining <= 0) capacityStatus = 'full';
          else if (remaining === 1) capacityStatus = 'pair_closed';
          else if (remaining / effectiveCapacity <= 0.50) capacityStatus = 'low';
          else capacityStatus = 'normal';
        }
      }
      return {
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
        fee:             ev.fee         || '',
        isFreeEvent:     ev.isFreeEvent === true,
        alreadyApplied:  appliedSheets.has(ev.resultSheetName),
        capacity:        ev.capacity    || 0,
        capacityStatus,
      };
    });
  } catch (err) {
    Logger.log('getLiffEventsJson error: ' + err.toString());
    return [];
  }
}

// ===== クライアントから呼び出すサーバー関数 =====

// 指定イベントに登録された紹介コード一覧と、それぞれの使用人数を返す（編集モーダルの紹介コード管理用）
function getReferralCodesForEvent(eventName) {
  try {
    const codes = getReferralCodesForEvent_(eventName);
    const ev = getAllEvents().find(e => e.name === eventName);
    const usedCounts = codes.map(c => ev ? countReferralCodeUsage_(ev.resultSheetName, c.code) : 0);
    return {
      success: true,
      codes: codes.map((c, i) => ({ code: c.code, referrerName: c.referrerName, maxCount: c.maxCount, usedCount: usedCounts[i] })),
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// 紹介コードを1件追加する（同じ大会に同じコードは重複登録できない）
function addReferralCode(eventName, code, referrerName, maxCount) {
  try {
    if (!eventName || !String(code || '').trim()) return { success: false, error: 'コードは必須です。' };
    const trimmedCode = String(code).trim();
    const sheet = ensureReferralCodesSheet_();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === eventName && String(data[i][1]).trim() === trimmedCode) {
        return { success: false, error: `コード「${trimmedCode}」は既にこの大会に登録されています。` };
      }
    }
    sheet.appendRow([eventName, trimmedCode, String(referrerName || '').trim(), parseInt(maxCount) || '']);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// 紹介コードを1件削除する
function deleteReferralCode(eventName, code) {
  try {
    const sheet = ensureReferralCodesSheet_();
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).trim() === eventName && String(data[i][1]).trim() === code) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: '対象のコードが見つかりません。' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

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
      eventDateISO:    ev.eventDate    ? Utilities.formatDate(ev.eventDate,    'Asia/Tokyo', 'yyyy-MM-dd') : '',
      closingDateISO:  ev.closingDate  ? Utilities.formatDate(ev.closingDate,  'Asia/Tokyo', 'yyyy-MM-dd') : '',
      openingDateISO:  ev.openingDate  ? Utilities.formatDate(ev.openingDate,  'Asia/Tokyo', 'yyyy-MM-dd') : '',
      appSheetName: ev.appSheetName,
      resultSheetName: ev.resultSheetName,
      eventTime:   ev.eventTime   || '',
      venue:       ev.venue       || '',
      coachName:   ev.coachName   || '',
      description: ev.description || '',
      eventType:   ev.eventType   || 'オフライン',
      channelUrl:  ev.channelUrl  || '',
      meetingTime:     ev.meetingTime     || '',
      courtType:       ev.courtType       || '',
      items:           ev.items           || '',
      fee:             ev.fee             || '',
      lockerInfo:      ev.lockerInfo      || '',
      facilityUrl:     ev.facilityUrl     || '',
      confirmDeadline: ev.confirmDeadline || '',
      confirmDeadlineAtISO: ev.confirmDeadlineAt ? Utilities.formatDate(ev.confirmDeadlineAt, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm") : '',
      closingDateTimeAtISO: ev.closingDateTimeAt ? Utilities.formatDate(ev.closingDateTimeAt, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm") : '',
      resultAnnouncementDateISO: ev.resultAnnouncementDate ? Utilities.formatDate(ev.resultAnnouncementDate, 'Asia/Tokyo', 'yyyy-MM-dd') : '',
      isFreeEvent:      ev.isFreeEvent      === true,
      capacity:         ev.capacity         || 0,
      ouboStatusHidden: ev.ouboStatusHidden === true,
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
          email:         String(mData[i][9]  || ''),
          phone:         String(mData[i][10] || '').replace(/^(\d{9,10})$/, '0$1'),
          furigana:      String(mData[i][11] || ''),
          tennisFreq:    String(mData[i][13] || ''),
          tennisHistory: String(mData[i][14] || ''),
          tennisArea:    String(mData[i][15] || ''),
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
      coachKnowledge:  String(data[i][5]  || ''),
      confirmation:    String(data[i][9]  || ''),
      shootingConsent: String(data[i][10] || ''),
      winCount:        winCountMap[baseUserId] || 0,
      age:            mInfo.age           || '',
      gender:         mInfo.gender        || '',
      tennisLevel:    mInfo.tennisLevel   || '',
      email:          mInfo.email         || '',
      phone:          mInfo.phone         || '',
      furigana:       mInfo.furigana      || '',
      tennisFreq:     mInfo.tennisFreq    || '',
      tennisHistory:  mInfo.tennisHistory || '',
      tennisArea:     mInfo.tennisArea    || '',
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
    // シート名は内部識別子（イベント名+日付サフィックス）のため、設定シートに登録された本来のイベント名を使う
    const ev = getAllEvents().find(e => e.resultSheetName === resultSheetName);
    const eventName = ev ? ev.name : resultSheetName.replace('_当落', '');
    notifyStaff(`📨 当落通知 送信完了\n${eventName}\n当選: ${winCount}名 / 落選: ${loseCount}名`);
  }

  return { success: true, winCount, loseCount };
}

// ステータスで絞り込んだUser IDリストを返す
function getFilteredUsers(appSheetName, resultSheetName, status) {
  const membersSheet = getSheet(SHEET.MEMBERS);
  if (!membersSheet || membersSheet.getLastRow() <= 1) return [];

  // 応募済みUser IDのセット・当落ステータスのマップを構築
  // 当落シートを基準にする（Google FormとLIFF両方の応募がここに集約される）
  const submittedSet = new Set();
  const resultMap = {};
  const resultSheet = getSheet(resultSheetName);
  if (resultSheet && resultSheet.getLastRow() > 1) {
    const resultData = resultSheet.getDataRange().getValues();
    for (let i = 1; i < resultData.length; i++) {
      const userId = String(resultData[i][1] || '');
      if (!userId) continue;
      submittedSet.add(userId);
      resultMap[userId] = String(resultData[i][2] || '');
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

// LIFFリンク一覧をSTAFF_USER_IDのLINEにpush送信する
function sendLiffLinksToStaff() {
  const staffUserId = getProp('STAFF_USER_ID');
  if (!staffUserId) return { success: false, error: 'STAFF_USER_IDが未設定です。スクリプトプロパティを確認してください。' };
  const liffId = getProp('LIFF_ID');
  if (!liffId) return { success: false, error: 'LIFF_IDが未設定です。' };
  const base = 'https://liff.line.me/' + liffId;
  const text = [
    '🔗 LIFFリンク一覧',
    '',
    '📝 プロフィール登録',
    base + '?page=register',
    '',
    '✏️ プロフィール修正',
    base + '?page=profile',
    '',
    '📍 オフライン応募',
    base + '?type=offline',
    '',
    '💻 オンライン応募',
    base + '?type=online',
    '',
    '🎥 動画相談（ビデオ相談を自動選択）',
    base + '?type=online&consult=video',
    '',
    '📝 文章相談（文章相談を自動選択）',
    base + '?type=online&consult=text',
    '',
    '🏆 大会応募',
    base + '?type=tournament',
  ].join('\n');
  pushMessage(staffUserId, text);
  return { success: true };
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

// 規約テキストをシートから読み込む。未設定のキーはデフォルト文で補完する
function getTermsContent() {
  const defaults = getDefaultTermsContent();
  try {
    const ss = SpreadsheetApp.openById(getProp('SPREADSHEET_ID'));
    const sheet = ss.getSheetByName(SHEET.TERMS);
    if (!sheet || sheet.getLastRow() === 0) return defaults;
    const data = sheet.getDataRange().getValues();
    const result = Object.assign({}, defaults);
    for (const row of data) {
      const key = String(row[0] || '').trim();
      const val = String(row[1] || '').trim();
      if (key && val) result[key] = val;
    }
    return result;
  } catch (e) {
    Logger.log('getTermsContent error: ' + e);
    return defaults;
  }
}

// デフォルト規約文を返す（初回・リセット用）
function getDefaultTermsContent() {
  return {
    tos_register: 'Epsom & Co. TENNIS SUPPORT PROJECT\n会員登録 利用規約・プライバシーポリシー\n\n株式会社サウスポー（以下「当社」）が運営する「Epsom & Co. TENNIS SUPPORT PROJECT」の会員サービスにご登録いただくにあたり、以下の規約・プライバシーポリシーを定めます。本サービスにご登録された時点で、本規約に同意いただいたものとみなします。\n\n第1条（会員登録）\n本サービスへの会員登録は、所定のフォームへの情報入力をもって成立します。登録いただいた情報は今後のイベント参加申し込みに使用され、再入力の手間を省くことができます。登録内容に誤りや虚偽があった場合、当社は登録を無効とする場合があります。\n\n第2条（未成年の登録）\n未成年の方が登録する場合は、保護者の同意を得たうえでご登録ください。\n\n第3条（禁止事項）\n虚偽情報による登録・第三者へのアカウント情報の提供等の不適切な行為を禁じます。\n\n第4条（登録情報の変更・退会）\n登録情報に変更が生じた場合は、フォームより速やかにご変更ください。退会をご希望の場合は当社までご連絡ください。\n\n第5条（規約の変更）\n当社は必要に応じて本規約を変更することができます。\n\nプライバシーポリシー\n\n取得する個人情報：氏名・フリガナ・年齢・性別・メールアドレス・電話番号・緊急連絡先（18歳以下）・テニスレベル・テニス歴等\n利用目的：会員情報の管理・イベント参加申し込み・当落連絡・運営連絡\n第三者提供：法令に基づく場合を除き、本人の同意なく第三者へ提供しません。',
    tos_offline: 'Epsom & Co. TENNIS SUPPORT\nイベント応募 利用規約\n\n株式会社サウスポーが主催する「Epsom & Co. TENNIS SUPPORT PROJECT」のイベントへの応募および参加条件を定めます。応募された時点で本規約に同意したものとみなします。\n\n第1条（応募方法）\n応募は、指定の応募フォームへの入力をもって成立します。応募者多数の場合は抽選を行い、当選・落選の双方にLINEでご連絡します。\n\n第2条（参加資格）\n当選連絡を受けた本人のみ参加可能です。当選権利の譲渡・転売はできません。\n\n第3条（未成年の参加）\n保護者の同意を得たうえで応募してください。\n\n第4条（キャンセル）\n参加できなくなった場合は速やかにご連絡ください。有料イベントのキャンセル料等は各イベントの募集要項に準じます。\n\n第5条（イベントの変更・中止）\nやむを得ない事情によりイベント内容を変更・中止する場合があります。\n\n第6条（禁止事項）\n他参加者・スタッフへの迷惑行為、虚偽応募、当選権利の譲渡等を禁じます。\n\n第7条（撮影・広報利用）\nイベント当日の写真・動画を公式SNS等で使用する場合があります。\n\n第8条（健康・安全）\n体調不良や怪我がある場合は参加を控えてください。\n\n第9条（免責）\nイベント参加に伴う事故・怪我・盗難等について、主催者は故意または重大な過失がない限り責任を負いません。各自でスポーツ傷害保険等へのご加入をお勧めします。',
    tos_online: 'Epsom & Co. TENNIS SUPPORT\nビデオ相談企画 応募規約\n\n株式会社サウスポーが実施する「ビデオ相談企画」への応募条件を定めます。応募された時点で本規約に同意したものとみなします。\n\n第1条（応募方法）\n指定フォームへの入力と動画の提出をもって応募が成立します。\n\n第2条（動画利用）\n応募データはYouTube配信・SNS投稿・広報活動等で使用される場合があります。配信後の削除・修正はお受けできません。\n\n第3条（著作者人格権）\n応募コンテンツの編集・改変について著作者人格権を行使しないものとします。\n\n第4条（禁止事項）\n虚偽応募・第三者の権利を侵害する動画の投稿等を禁じます。\n\n第5条（免責）\n主催者の故意または重大な過失がある場合を除き、責任を負いません。',
    media_offline_free: '株式会社サウスポーでは、イベントの様子（コート全体の風景、練習風景など）を写真・動画で記録しており、公式SNSをはじめとする各種媒体での情報発信や告知等に使用します。\n\nイベント中の撮影および参加者の皆様へのインタビュー（顔出しあり）へのご協力を前提として運営しております。ご応募の際は、撮影・インタビューへのご協力に同意いただいた上でお申し込みください。',
    media_offline_paid: '株式会社サウスポーでは、イベントの様子（コート全体の風景、練習風景など）を写真・動画で記録しており、公式SNSをはじめとする各種媒体での情報発信や告知等に使用します。\n\nイベント風景（遠景等による映り込みを含む）の撮影・公開への同意が必要となります。ただし、個別のお写真撮影やインタビュー（顔出しあり）へのご協力については任意（拒否可能）となります。お断りされる場合は、当日スタッフまでお気軽にお申し出ください。',
    media_online: 'ご応募いただいた動画およびご相談内容は、YouTube生配信「テニスポットラジオ」、アーカイブ配信、SNS投稿、切り抜き動画、Webサイト等で使用させていただく場合があります。\n\nまた、配信・投稿の際に動画の一部を編集・加工して使用する場合があります。\n\n応募後および配信後は、原則として動画の削除依頼や公開停止のご要望にはお応えできません。',
  };
}

// 規約テキストを規約管理シートに保存する
function saveTermsContent(data) {
  try {
    const ss = SpreadsheetApp.openById(getProp('SPREADSHEET_ID'));
    let sheet = ss.getSheetByName(SHEET.TERMS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET.TERMS);
    }
    sheet.clearContents();
    const keys = ['tos_register', 'tos_offline', 'tos_online', 'media_offline_free', 'media_offline_paid', 'media_online'];
    for (const key of keys) {
      if (data[key] !== undefined) sheet.appendRow([key, data[key]]);
    }
    return { success: true };
  } catch (e) {
    Logger.log('saveTermsContent error: ' + e);
    return { success: false, error: e.toString() };
  }
}

// ファネル集計：アクション履歴からステージ別のユニーク人数を返す
function getFunnelStats(days) {
  try {
    const sheet = getSheet(SHEET.ACTION_LOG);
    if (!sheet || sheet.getLastRow() <= 1) return { stages: [], days: days };
    const data = sheet.getDataRange().getValues();
    const cutoff = days > 0 ? new Date(Date.now() - days * 86400 * 1000) : null;
    const stages = ['LIFF起動', '応募ボタン押下', 'LIFF応募', '動画受信'];
    const sets = {};
    stages.forEach(s => sets[s] = new Set());
    for (let i = 1; i < data.length; i++) {
      if (cutoff && data[i][0] < cutoff) continue;
      const userId = String(data[i][1] || '');
      const action = String(data[i][2] || '');
      if (sets[action]) sets[action].add(userId);
    }
    const base = sets['LIFF起動'].size || 1;
    return {
      days: days,
      stages: stages.map((s, i) => ({
        name: s,
        count: sets[s].size,
        rate: i === 0 ? null : Math.round(sets[s].size / base * 100),
      })),
    };
  } catch (err) {
    Logger.log('getFunnelStats error: ' + err.toString());
    throw err;
  }
}

// 種別・相談方法別ファネル（LIFF起動時のURLパラメータでセグメント分け）
function getFunnelByType(days) {
  try {
    const sheet = getSheet(SHEET.ACTION_LOG);
    if (!sheet || sheet.getLastRow() <= 1) return { segments: [], days };
    const data = sheet.getDataRange().getValues();
    const cutoff = days > 0 ? new Date(Date.now() - days * 86400 * 1000) : null;
    const typeMap = {};
    const userTypeMap = {}; // userId -> segment（初回起動時の種別を記録）
    const getOrCreate = seg => {
      if (!typeMap[seg]) typeMap[seg] = { open: new Set(), press: new Set(), complete: new Set() };
      return typeMap[seg];
    };
    for (let i = 1; i < data.length; i++) {
      if (cutoff && data[i][0] < cutoff) continue;
      const userId = String(data[i][1] || '');
      const action = String(data[i][2] || '');
      const eventId = String(data[i][3] || '');
      const detail  = String(data[i][4] || '');
      if (action === 'LIFF起動') {
        let seg = eventId || '通常';
        if (detail) seg += '/' + detail;
        if (!userTypeMap[userId]) userTypeMap[userId] = seg;
        getOrCreate(seg).open.add(userId);
      } else if (action === '応募ボタン押下') {
        getOrCreate(userTypeMap[userId] || '通常').press.add(userId);
      } else if (action === 'LIFF応募') {
        getOrCreate(userTypeMap[userId] || '通常').complete.add(userId);
      }
    }
    const segments = Object.entries(typeMap).map(([seg, sets]) => ({
      segment: seg,
      open: sets.open.size, press: sets.press.size, complete: sets.complete.size,
      pressRate:    sets.open.size > 0 ? Math.round(sets.press.size    / sets.open.size * 100) : 0,
      completeRate: sets.open.size > 0 ? Math.round(sets.complete.size / sets.open.size * 100) : 0,
    })).sort((a, b) => b.open - a.open);
    return { segments, days };
  } catch (err) { Logger.log('getFunnelByType error: ' + err); throw err; }
}

// 日別ユニーク人数の推移
function getFunnelByDay(days) {
  try {
    const sheet = getSheet(SHEET.ACTION_LOG);
    if (!sheet || sheet.getLastRow() <= 1) return { stages: [], days: [], totalDays: days };
    const data = sheet.getDataRange().getValues();
    const cutoff = days > 0 ? new Date(Date.now() - days * 86400 * 1000) : null;
    const stages = ['LIFF起動', '応募ボタン押下', 'LIFF応募', '動画受信'];
    const dayMap = {};
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (cutoff && data[i][0] < cutoff) continue;
      const userId = String(data[i][1] || '');
      const action = String(data[i][2] || '');
      if (!stages.includes(action)) continue;
      const dateStr = Utilities.formatDate(new Date(data[i][0]), 'Asia/Tokyo', 'MM/dd');
      if (!dayMap[dateStr]) { dayMap[dateStr] = {}; stages.forEach(s => dayMap[dateStr][s] = new Set()); }
      dayMap[dateStr][action].add(userId);
    }
    const result = Object.entries(dayMap)
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([date, s]) => ({ date, counts: stages.map(st => s[st].size) }));
    return { stages, days: result, totalDays: days };
  } catch (err) { Logger.log('getFunnelByDay error: ' + err); throw err; }
}

// 離脱分析：応募ボタンを押したが完了しなかったユーザーと選択イベント
function getDropoffDetail(days) {
  try {
    const sheet = getSheet(SHEET.ACTION_LOG);
    if (!sheet || sheet.getLastRow() <= 1) return { pressed: 0, completed: 0, total: 0, byEvent: [], dropoffs: [], days };
    const data = sheet.getDataRange().getValues();
    const cutoff = days > 0 ? new Date(Date.now() - days * 86400 * 1000) : null;
    const pressed = {}, completed = new Set();
    for (let i = 1; i < data.length; i++) {
      if (cutoff && data[i][0] < cutoff) continue;
      const userId = String(data[i][1] || '');
      const action = String(data[i][2] || '');
      if (action === '応募ボタン押下') {
        pressed[userId] = {
          ts: Utilities.formatDate(new Date(data[i][0]), 'Asia/Tokyo', 'MM/dd HH:mm'),
          events: String(data[i][3] || ''), consult: String(data[i][4] || ''),
        };
      } else if (action === 'LIFF応募') {
        completed.add(userId);
      }
    }
    const dropoffs = Object.entries(pressed)
      .filter(([uid]) => !completed.has(uid))
      .map(([, d]) => d)
      .sort((a, b) => b.ts.localeCompare(a.ts));
    const eventCounts = {};
    dropoffs.forEach(d => (d.events || '').split('、').filter(Boolean).forEach(ev => { eventCounts[ev] = (eventCounts[ev] || 0) + 1; }));
    return {
      days, pressed: Object.keys(pressed).length, completed: completed.size, total: dropoffs.length,
      byEvent: Object.entries(eventCounts).sort((a,b)=>b[1]-a[1]).map(([ev,count])=>({ev,count})),
      dropoffs: dropoffs.slice(0, 50),
    };
  } catch (err) { Logger.log('getDropoffDetail error: ' + err); throw err; }
}

// 個人タイムライン：指定UserIDのアクション履歴を時系列で返す
function getUserTimeline(userId) {
  try {
    const sheet = getSheet(SHEET.ACTION_LOG);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    const data = sheet.getDataRange().getValues();
    return data.slice(1)
      .filter(r => String(r[1] || '') === userId)
      .map(r => ({
        ts:      Utilities.formatDate(new Date(r[0]), 'Asia/Tokyo', 'MM/dd HH:mm'),
        action:  String(r[2] || ''),
        eventId: String(r[3] || ''),
        detail:  String(r[4] || ''),
      }))
      .sort((a, b) => a.ts.localeCompare(b.ts));
  } catch (err) { Logger.log('getUserTimeline error: ' + err); throw err; }
}

// スプレッドシートに「ファネル集計」シートを作成（ROWS+UNIQUE+FILTER数式で自動集計）
function setupFunnelSheet() {
  const ss = SpreadsheetApp.openById(getProp('SPREADSHEET_ID'));
  let sheet = ss.getSheetByName('ファネル集計');
  if (sheet) { sheet.clearContents(); } else { sheet = ss.insertSheet('ファネル集計'); }
  const log = SHEET.ACTION_LOG;
  sheet.getRange('A1').setValue('📊 ファネル集計').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A2').setValue('最終出力: ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));
  sheet.getRange('A4:E4').setValues([['ステップ', 'ユニーク人数', '起動からの率', '前ステップからの率', '離脱数']]).setFontWeight('bold');
  const stages = ['LIFF起動', '応募ボタン押下', 'LIFF応募', '動画受信'];
  stages.forEach((stage, idx) => {
    const row = idx + 5;
    sheet.getRange(row, 1).setValue(stage);
    sheet.getRange(row, 2).setFormula(
      `=IFERROR(ROWS(UNIQUE(FILTER(${log}!B:B,${log}!C:C="${stage}",${log}!B:B<>""))),0)`
    );
    sheet.getRange(row, 3).setFormula(idx === 0 ? '="100%"' : `=IFERROR(TEXT(B${row}/B5,"0%"),"-")`);
    sheet.getRange(row, 4).setFormula(idx === 0 ? '="-"' : `=IFERROR(TEXT(B${row}/B${row-1},"0%"),"-")`);
    sheet.getRange(row, 5).setFormula(idx === stages.length-1 ? '="-"' : `=IFERROR(B${row}-B${row+1},"-")`);
  });
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 80);
  return { success: true };
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
'<li class="nav-item"><a class="nav-link" href="#" id="tab-btn-funnel" onclick="showTab(\'funnel\');return false">📊 ファネル分析</a></li>' +
'<li class="nav-item"><a class="nav-link" href="#" id="tab-btn-messages" onclick="showTab(\'messages\');return false">📝 文章管理</a></li>' +
'<li class="nav-item"><a class="nav-link" href="#" id="tab-btn-terms" onclick="showTab(\'terms\');return false">📋 規約管理</a></li>' +
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
'<div class="col-8"><label class="form-label fw-bold">イベント種別</label>' +
'<select class="form-select" id="ne_type" onchange="onEventTypeChange()">' +
'<option value="オフライン">📍 オフライン</option>' +
'<option value="オンライン">💻 オンライン</option>' +
'<option value="選手交流">🤝 選手交流</option>' +
'<option value="大会">🏆 大会</option>' +
'</select></div>' +
'<div class="col-4 d-flex align-items-end pb-1">' +
'<div class="form-check"><input type="checkbox" class="form-check-input" id="ne_is_free"><label class="form-check-label fw-bold" for="ne_is_free">無料イベント</label></div>' +
'<div class="form-check ms-3"><input type="checkbox" class="form-check-input" id="ne_oubo_hidden"><label class="form-check-label fw-bold text-muted" for="ne_oubo_hidden">応募状況に非表示</label></div>' +
'</div>' +
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
'<div class="col-6"><label class="form-label fw-bold">集合時間</label>' +
'<input type="text" class="form-control" id="ne_meeting" placeholder="18:50"></div>' +
'<div class="col-6"><label class="form-label fw-bold">コートについて</label>' +
'<input type="text" class="form-control" id="ne_court" placeholder="カーペットコートになります。"></div>' +
'<div class="col-12"><label class="form-label fw-bold">持ち物</label>' +
'<textarea class="form-control" id="ne_items" rows="2" placeholder="・ラケット\n・テニスウェア\n・テニスシューズ"></textarea></div>' +
'<div class="col-6"><label class="form-label fw-bold">参加費</label>' +
'<input type="text" class="form-control" id="ne_fee" placeholder="EPSOM&CO.様のサポートにより、無料でご参加いただけます！"></div>' +
'<div class="col-6"><label class="form-label fw-bold">募集締め切り日時<span class="text-muted small fw-normal ms-1">（任意）</span></label>' +
'<input type="datetime-local" class="form-control" id="ne_closing_at"><div class="form-text">設定すると「応募状況」に「〇月〇日 〇:〇〇まで受付中」と表示されます</div></div>' +
'<div class="col-6"><label class="form-label fw-bold">当落通知予定日<span class="text-muted small fw-normal ms-1">（任意）</span></label>' +
'<input type="date" class="form-control" id="ne_result_announcement"><div class="form-text">設定すると「応募状況」に「当落は〇月〇日頃にお知らせします」と表示されます</div></div>' +
'<div class="col-6"><label class="form-label fw-bold">参加確認期限</label>' +
'<input type="text" class="form-control" id="ne_deadline" placeholder="5月27日（水）12:00"></div>' +
'<div class="col-6"><label class="form-label fw-bold">参加確認期限（自動キャンセル日時）<span class="text-muted small fw-normal ms-1">（任意）</span></label>' +
'<input type="datetime-local" class="form-control" id="ne_deadline_at"><div class="form-text">この日時を過ぎて「参加します」が押された場合、自動的に期限切れとして扱います</div></div>' +
'<div class="col-12"><label class="form-label fw-bold">更衣室について</label>' +
'<textarea class="form-control" id="ne_locker" rows="2" placeholder="受付でのお声がけは不要です。..."></textarea></div>' +
'<div class="col-12"><label class="form-label fw-bold">施設URL</label>' +
'<input type="url" class="form-control" id="ne_facility_url" placeholder="https://..."></div>' +
'<!-- 大会専用 -->' +
'<div id="ne_capacity_field" class="col-6" style="display:none">' +
'<label class="form-label fw-bold">定員<span class="text-danger">*</span></label>' +
'<input type="number" class="form-control" id="ne_capacity" placeholder="32" min="2">' +
'<div class="form-text">先着順の最大参加人数（ペア=2名カウント）</div>' +
'</div>' +
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
'<div class="col-12" id="ne_coach_field"><label class="form-label fw-bold">コーチ名</label>' +
'<input type="text" class="form-control" id="ne_coach" placeholder="山田 コーチ"></div>' +
'<div class="col-12"><label class="form-label fw-bold">イベント内容</label>' +
'<textarea class="form-control" id="ne_desc" rows="3" placeholder="イベントの説明・内容を入力"></textarea></div>' +
'</div>' +
'<div class="d-flex gap-2 mt-3 align-items-center">' +
'<button class="btn btn-success" onclick="submitNewEvent()">登録する</button>' +
'<button class="btn btn-outline-secondary" onclick="hideNewEventModal()">キャンセル</button>' +
'<span id="ne_result" class="text-muted small"></span>' +
'</div></div>' +
'<!-- イベント詳細編集モーダル -->' +
'<div id="editEventModal" class="card p-3 mb-3" style="display:none;border:2px solid #0d6efd">' +
'<h6 class="mb-3">✏️ イベント詳細を編集</h6>' +
'<div class="row g-2">' +
'<div class="col-12"><label class="form-label fw-bold">イベント名<span class="text-danger">*</span></label><input type="text" class="form-control fw-bold" id="ee_name"></div>' +
'<div class="col-12 d-flex gap-3 justify-content-end">' +
'<div class="form-check"><input type="checkbox" class="form-check-input" id="ee_is_free"><label class="form-check-label fw-bold" for="ee_is_free">無料イベント</label></div>' +
'<div class="form-check"><input type="checkbox" class="form-check-input" id="ee_oubo_hidden"><label class="form-check-label fw-bold text-muted" for="ee_oubo_hidden">応募状況に非表示</label></div>' +
'</div>' +
'<div id="ee_offline_fields" class="col-12">' +
'<div class="row g-2">' +
'<div class="col-6"><label class="form-label fw-bold">応募開始日</label><input type="date" class="form-control" id="ee_opening"></div>' +
'<div class="col-6"><label class="form-label fw-bold">募集終了日</label><input type="date" class="form-control" id="ee_closing"></div>' +
'<div class="col-6"><label class="form-label fw-bold">開催日</label><input type="date" class="form-control" id="ee_date"></div>' +
'<div class="col-6"><label class="form-label fw-bold">開催時間（レッスン時間）</label><input type="text" class="form-control" id="ee_time" placeholder="19:00〜21:00"></div>' +
'<div class="col-6"><label class="form-label fw-bold">開催場所</label><input type="text" class="form-control" id="ee_venue"></div>' +
'<div class="col-6"><label class="form-label fw-bold">集合時間</label><input type="text" class="form-control" id="ee_meeting"></div>' +
'<div class="col-6"><label class="form-label fw-bold">コートについて</label><input type="text" class="form-control" id="ee_court"></div>' +
'<div class="col-6"><label class="form-label fw-bold">参加費</label><input type="text" class="form-control" id="ee_fee"></div>' +
'<div class="col-12"><label class="form-label fw-bold">持ち物</label><textarea class="form-control" id="ee_items" rows="2"></textarea></div>' +
'<div class="col-6"><label class="form-label fw-bold">募集締め切り日時<span class="text-muted small fw-normal ms-1">（任意）</span></label><input type="datetime-local" class="form-control" id="ee_closing_at"></div>' +
'<div class="col-6"><label class="form-label fw-bold">当落通知予定日<span class="text-muted small fw-normal ms-1">（任意）</span></label><input type="date" class="form-control" id="ee_result_announcement"></div>' +
'<div class="col-6"><label class="form-label fw-bold">参加確認期限</label><input type="text" class="form-control" id="ee_deadline"></div>' +
'<div class="col-6"><label class="form-label fw-bold">参加確認期限（自動キャンセル日時）</label><input type="datetime-local" class="form-control" id="ee_deadline_at"></div>' +
'<div class="col-12"><label class="form-label fw-bold">更衣室について</label><textarea class="form-control" id="ee_locker" rows="2"></textarea></div>' +
'<div class="col-12"><label class="form-label fw-bold">施設URL</label><input type="url" class="form-control" id="ee_facility_url"></div>' +
'<div id="ee_capacity_field" class="col-6" style="display:none">' +
'<label class="form-label fw-bold">定員</label>' +
'<input type="number" class="form-control" id="ee_capacity" placeholder="32" min="2">' +
'<div class="form-text">先着順の最大参加人数（ペア=2名カウント）</div>' +
'</div>' +
'<div id="ee_referral_field" class="col-12" style="display:none">' +
'<label class="form-label fw-bold">紹介コード</label>' +
'<div id="ee_referral_list" class="mb-2"></div>' +
'<div class="d-flex gap-2 flex-wrap align-items-end">' +
'<div><label class="form-label small mb-0">コード</label><input type="text" class="form-control form-control-sm" id="ee_referral_code_input" style="width:130px" placeholder="SHOKAI001"></div>' +
'<div><label class="form-label small mb-0">紹介者名</label><input type="text" class="form-control form-control-sm" id="ee_referral_name_input" style="width:130px" placeholder="山田花子"></div>' +
'<div><label class="form-label small mb-0">上限件数<span class="text-muted">（1人でもペアでも1件。空欄=無制限）</span></label><input type="number" class="form-control form-control-sm" id="ee_referral_max_input" style="width:90px" min="1"></div>' +
'<button class="btn btn-outline-primary btn-sm" onclick="addReferralCodeFromUI()">＋ 追加</button>' +
'</div>' +
'<div class="form-text">紹介コードで応募した人の備考欄に、ここで設定した紹介者名が自動で記録されます。</div>' +
'</div>' +
'</div></div>' +
'<div id="ee_online_fields" class="col-12" style="display:none">' +
'<div class="row g-2">' +
'<div class="col-12"><label class="form-label fw-bold">チャンネルURL</label><input type="url" class="form-control" id="ee_channel_url"></div>' +
'<div class="col-6"><label class="form-label fw-bold">応募開始日</label><input type="date" class="form-control" id="ee_opening_online"></div>' +
'<div class="col-6"><label class="form-label fw-bold">募集終了日</label><input type="date" class="form-control" id="ee_closing_online"></div>' +
'</div></div>' +
'<div class="col-12" id="ee_coach_field"><label class="form-label fw-bold">コーチ名</label><input type="text" class="form-control" id="ee_coach"></div>' +
'<div class="col-12"><label class="form-label fw-bold">イベント内容</label><textarea class="form-control" id="ee_desc" rows="3"></textarea></div>' +
'</div>' +
'<div class="d-flex gap-2 mt-3 align-items-center">' +
'<button class="btn btn-primary" onclick="submitEditEvent()">保存する</button>' +
'<button class="btn btn-outline-secondary" onclick="hideEditEventModal()">キャンセル</button>' +
'<span id="ee_result" class="text-muted small"></span>' +
'</div></div>' +
'<div id="eventList" class="row g-2 mb-3"></div>' +
'<div id="applicantSection" style="display:none">' +
'<div class="d-flex align-items-center gap-2 mb-2 flex-wrap">' +
'<h6 class="mb-0" id="applicantTitle"></h6>' +
'<button class="btn btn-success btn-sm ms-auto" onclick="sendNotifications()">📨 当落通知を送信</button>' +
'<button class="btn btn-outline-primary btn-sm" onclick="openPrintModal()">📄 表を作成</button>' +
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
'<th>名前</th><th>年齢</th><th>性別</th><th>レベル</th><th>テニス歴</th><th>頻度</th><th>応募日時</th><th>当落</th><th>当選回数</th><th>撮影可否</th><th>コーチ認知</th><th>通知</th><th>操作</th></tr>' +
'</thead>' +
'<tbody id="applicantBody"></tbody>' +
'</table>' +
'</div>' +
'</div>' +

'<div id="printModal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;overflow-y:auto">' +
'<div style="background:#fff;margin:40px auto;padding:24px;max-width:500px;border-radius:8px;box-shadow:0 4px 24px rgba(0,0,0,.2)">' +
'<h6 class="mb-1">📄 表を作成</h6>' +
'<p class="text-muted small mb-3">チェックした人の表が別タブで開きます。ブラウザの印刷機能でPDFに保存できます。</p>' +
'<div id="printColChecks" class="mb-3" style="columns:2;gap:12px"></div>' +
'<div class="d-flex gap-2 flex-wrap">' +
'<button class="btn btn-primary" onclick="buildPrintTable()">表を作成 →</button>' +
'<button class="btn btn-outline-secondary" onclick="document.getElementById(\'printModal\').style.display=\'none\'">キャンセル</button>' +
'</div>' +
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

'<!-- ファネル分析タブ -->' +
'<div id="tab-funnel" style="display:none">' +
'<div class="d-flex align-items-center gap-2 mb-3 flex-wrap">' +
'<select class="form-select form-select-sm" id="funnelDays" style="width:auto">' +
'<option value="7">直近7日</option>' +
'<option value="30" selected>直近30日</option>' +
'<option value="90">直近90日</option>' +
'<option value="0">全期間</option>' +
'</select>' +
'<button class="btn btn-sm btn-primary" onclick="loadFunnel()">📊 集計する</button>' +
'<button class="btn btn-sm btn-outline-secondary" onclick="runSetupFunnelSheet()">📋 スプレッドシートに出力</button>' +
'<button class="btn btn-sm btn-outline-info" onclick="toggleLookerInfo()">📈 Looker Studio</button>' +
'</div>' +
'<div id="lookerInfo" style="display:none" class="alert alert-info mb-3 small">' +
'<strong>Looker Studio でグラフを作る手順</strong><br>' +
'① <a href="https://lookerstudio.google.com" target="_blank">lookerstudio.google.com</a> を開き「レポートを作成」<br>' +
'② データソース → 「Googleスプレッドシート」 → このスプレッドシートの「アクション履歴」シートを選択<br>' +
'③ C列(actionType)をディメンション、B列(userId)をCOUNT_DISTINCTで集計するとファネルグラフが作れます' +
'</div>' +
'<ul class="nav nav-pills mb-3 small">' +
'<li class="nav-item"><a class="nav-link py-1 active" href="#" id="ftab-btn-overview" onclick="showFunnelTab(\'overview\');return false">📊 概要</a></li>' +
'<li class="nav-item"><a class="nav-link py-1" href="#" id="ftab-btn-bytype" onclick="showFunnelTab(\'bytype\');return false">🗂 種別別</a></li>' +
'<li class="nav-item"><a class="nav-link py-1" href="#" id="ftab-btn-daily" onclick="showFunnelTab(\'daily\');return false">📅 日別推移</a></li>' +
'<li class="nav-item"><a class="nav-link py-1" href="#" id="ftab-btn-dropoff" onclick="showFunnelTab(\'dropoff\');return false">⚠️ 離脱分析</a></li>' +
'<li class="nav-item"><a class="nav-link py-1" href="#" id="ftab-btn-user" onclick="showFunnelTab(\'user\');return false">👤 個人検索</a></li>' +
'</ul>' +
'<div id="ftab-overview"><p class="text-muted text-center mt-4">「集計する」を押してください</p></div>' +
'<div id="ftab-bytype" style="display:none"><p class="text-muted text-center mt-4">「集計する」を押してください</p></div>' +
'<div id="ftab-daily" style="display:none"><p class="text-muted text-center mt-4">「集計する」を押してください</p></div>' +
'<div id="ftab-dropoff" style="display:none"><p class="text-muted text-center mt-4">「集計する」を押してください</p></div>' +
'<div id="ftab-user" style="display:none">' +
'<div class="input-group mb-3" style="max-width:440px">' +
'<input type="text" class="form-control form-control-sm" id="userSearchId" placeholder="LINE User ID（U で始まる文字列）">' +
'<button class="btn btn-sm btn-outline-secondary" onclick="loadUserTimeline()">検索</button>' +
'</div>' +
'<div id="userTimelineResult"><p class="text-muted small">会員一覧タブで名前をクリックすると User ID が取得できます</p></div>' +
'</div>' +
'</div>' +

'<!-- 文章管理タブ -->' +
'<div id="tab-messages" style="display:none">' +
'<div class="alert alert-light border small mb-3">LINEに自動送信される文章をここで編集できます。<code>{xxx}</code>の部分は実際の内容（イベント名や開催日・コーチ名など）に自動で置き換わるので、そのまま残してください。当選・落選メッセージはオンライン/オフラインそれぞれの基本形1つに対し、イベントごとの詳細（開催日・場所・持ち物など）は「イベント一覧」タブの「✏️ 詳細編集」から入力してください。<br><code>{{#if xxx}}...{{/if}}</code>で囲まれた範囲は、その項目（イベント詳細編集で入力するもの）が未入力の場合に見出しごと非表示になります。誤って削除しないようご注意ください。<br>改行はそのままEnterで入力してください（LINEはプレーンテキストなので<code>&lt;br&gt;</code>と書いても改行されず、文字としてそのまま表示されてしまいます）。<br>「🧪 自分に試し送信」を押すと、<code>{xxx}</code>部分をサンプルの値に置き換えて、STAFF_USER_IDに設定されているLINEへテスト送信できます（保存していない編集中の内容でも送信されます）。</div>' +
'<div id="msgTemplatesList"><p class="text-muted text-center mt-4">読み込み中...</p></div>' +
'</div>' +

'<!-- 規約管理タブ -->' +
'<div id="tab-terms" style="display:none">' +
'<div class="alert alert-light border small mb-3">LIFFフォームに表示される利用規約・プライバシーポリシー・撮影同意文を編集できます。HTMLタグ（&lt;br&gt;&lt;strong&gt;など）はそのまま使用できます。「💾 保存」を押すと即座に反映されます。</div>' +
'<div class="d-flex justify-content-end mb-2 gap-2"><button class="btn btn-outline-secondary" onclick="initTerms()">🔄 デフォルト文を読み込む</button><button class="btn btn-primary" onclick="saveAllTerms()">💾 すべて保存</button><span id="termsResult" class="text-muted small ms-3 align-self-center"></span></div>' +
'<div class="row g-3">' +
'<div class="col-12"><label class="form-label fw-bold">会員登録 利用規約・プライバシーポリシー</label><textarea class="form-control font-monospace" id="tos_register" rows="10"></textarea></div>' +
'<div class="col-12"><label class="form-label fw-bold">オフラインイベント応募 利用規約</label><textarea class="form-control font-monospace" id="tos_offline" rows="10"></textarea></div>' +
'<div class="col-12"><label class="form-label fw-bold">オンライン（ビデオ相談）応募規約</label><textarea class="form-control font-monospace" id="tos_online" rows="10"></textarea></div>' +
'<div class="col-12"><label class="form-label fw-bold">撮影・広報利用同意（オフライン・無料イベント用）</label><textarea class="form-control font-monospace" id="media_offline_free" rows="6"></textarea></div>' +
'<div class="col-12"><label class="form-label fw-bold">撮影・広報利用同意（オフライン・有料イベント用）</label><textarea class="form-control font-monospace" id="media_offline_paid" rows="6"></textarea></div>' +
'<div class="col-12"><label class="form-label fw-bold">撮影・広報利用同意（オンライン）</label><textarea class="form-control font-monospace" id="media_online" rows="6"></textarea></div>' +
'</div></div>' +

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
'var messagesLoaded=false;' +
'var termsLoaded=false;' +
'var msgTemplatesData={};' +
'var allApplicantsData=[];' +
'var dashConfig={};' +

'window.onload=function(){' +
'google.script.run' +
'.withSuccessHandler(function(c){dashConfig=c;loadEvents();})' +
'.withFailureHandler(function(){loadEvents();})' +
'.getDashboardConfig();' +
'};' +

'function showNewEventModal(){' +
'document.getElementById("newEventModal").style.display="";' +
'document.getElementById("ne_items").value="・ラケット\\n・テニスウェア\\n・テニスシューズ";' +
'document.getElementById("ne_fee").value="EPSOM&CO.様のサポートにより、無料でご参加いただけます！";' +
'}' +
'function hideNewEventModal(){' +
'document.getElementById("newEventModal").style.display="none";' +
'document.getElementById("ne_result").textContent="";' +
'document.getElementById("ne_type").value="オフライン";' +
'onEventTypeChange();' +
'["ne_name","ne_opening","ne_closing","ne_date","ne_venue","ne_coach","ne_desc","ne_channel_url","ne_opening_online","ne_closing_online",' +
'"ne_meeting","ne_court","ne_items","ne_fee","ne_deadline","ne_deadline_at","ne_locker","ne_facility_url"].forEach(function(id){var el=document.getElementById(id);if(el)el.value="";});' +
'["ne_time_start","ne_time_end"].forEach(function(id){var el=document.getElementById(id);if(el)el.value="";});' +
'}' +
'function onEventTypeChange(){' +
'var evType=document.getElementById("ne_type").value;' +
'var isOnline=evType==="オンライン";' +
'var isTournament=evType==="大会";' +
'document.getElementById("ne_offline_fields").style.display=isOnline?"none":"";' +
'document.getElementById("ne_online_fields").style.display=isOnline?"":"none";' +
'var cf=document.getElementById("ne_capacity_field");if(cf)cf.style.display=isTournament?"":"none";' +
'var ncf=document.getElementById("ne_coach_field");if(ncf)ncf.style.display=isTournament?"none":"";' +
'}' +
'function submitNewEvent(){' +
'var evType=document.getElementById("ne_type").value;' +
'var isOnline=evType==="オンライン";' +
'var name=document.getElementById("ne_name").value.trim();' +
'var coach=document.getElementById("ne_coach").value.trim();' +
'var desc=document.getElementById("ne_desc").value.trim();' +
'var date="",closing="",opening="",time="",venue="",channelUrl="";' +
'var meeting="",court="",items="",fee="",deadline="",deadlineAt="",locker="",facilityUrl="";' +
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
'meeting=document.getElementById("ne_meeting").value.trim();' +
'court=document.getElementById("ne_court").value.trim();' +
'items=document.getElementById("ne_items").value.trim();' +
'fee=document.getElementById("ne_fee").value.trim();' +
'deadline=document.getElementById("ne_deadline").value.trim();' +
'deadlineAt=document.getElementById("ne_deadline_at").value;' +
'closingAt=document.getElementById("ne_closing_at").value;' +
'resultAnnouncement=document.getElementById("ne_result_announcement").value;' +
'isFreeEvent=document.getElementById("ne_is_free").checked;' +
'ouboStatusHidden=document.getElementById("ne_oubo_hidden").checked;' +
'locker=document.getElementById("ne_locker").value.trim();' +
'facilityUrl=document.getElementById("ne_facility_url").value.trim();' +
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
'.createNewEvent({name:name,eventDate:date,closingDate:closing,openingDate:opening,eventTime:time,venue:venue,coachName:coach,description:desc,channelUrl:channelUrl,eventType:evType,' +
'meetingTime:meeting,courtType:court,items:items,fee:fee,confirmDeadline:deadline,confirmDeadlineAt:deadlineAt,closingDateTimeAt:closingAt,resultAnnouncementDate:resultAnnouncement,isFreeEvent:isFreeEvent,lockerInfo:locker,facilityUrl:facilityUrl,' +
'capacity:evType==="大会"?parseInt((document.getElementById("ne_capacity")||{}).value||"0")||0:0,' +
'ouboStatusHidden:ouboStatusHidden});' +
'}' +

'function openEditEventModal(idx,e){' +
'if(e)e.stopPropagation();' +
'var ev=eventsData[idx];' +
'document.getElementById("editEventModal").dataset.appSheetName=ev.appSheetName;' +
'document.getElementById("editEventModal").dataset.eventName=ev.name;' +
'document.getElementById("ee_name").value=ev.name;' +
'var isOnline=ev.eventType==="オンライン";' +
'var isTournamentEv=ev.eventType==="大会";' +
'document.getElementById("ee_offline_fields").style.display=isOnline?"none":"";' +
'document.getElementById("ee_online_fields").style.display=isOnline?"":"none";' +
'var eecf=document.getElementById("ee_capacity_field");if(eecf)eecf.style.display=isTournamentEv?"":"none";' +
'var eerf=document.getElementById("ee_referral_field");if(eerf)eerf.style.display=isTournamentEv?"":"none";' +
'if(isTournamentEv)loadReferralCodes(ev.name);' +
'var eecoacf=document.getElementById("ee_coach_field");if(eecoacf)eecoacf.style.display=isTournamentEv?"none":"";' +
'var eecap=document.getElementById("ee_capacity");if(eecap)eecap.value=ev.capacity||"";' +
'document.getElementById("ee_date").value=ev.eventDateISO||"";' +
'document.getElementById("ee_closing").value=ev.closingDateISO||"";' +
'document.getElementById("ee_opening").value=ev.openingDateISO||"";' +
'document.getElementById("ee_time").value=ev.eventTime||"";' +
'document.getElementById("ee_venue").value=ev.venue||"";' +
'document.getElementById("ee_meeting").value=ev.meetingTime||"";' +
'document.getElementById("ee_court").value=ev.courtType||"";' +
'document.getElementById("ee_fee").value=ev.fee||"";' +
'document.getElementById("ee_items").value=ev.items||"";' +
'document.getElementById("ee_is_free").checked=!!ev.isFreeEvent;' +
'document.getElementById("ee_oubo_hidden").checked=!!ev.ouboStatusHidden;' +
'document.getElementById("ee_closing_at").value=ev.closingDateTimeAtISO||"";' +
'document.getElementById("ee_result_announcement").value=ev.resultAnnouncementDateISO||"";' +
'document.getElementById("ee_deadline").value=ev.confirmDeadline||"";' +
'document.getElementById("ee_deadline_at").value=ev.confirmDeadlineAtISO||"";' +
'document.getElementById("ee_locker").value=ev.lockerInfo||"";' +
'document.getElementById("ee_facility_url").value=ev.facilityUrl||"";' +
'document.getElementById("ee_channel_url").value=ev.channelUrl||"";' +
'document.getElementById("ee_opening_online").value=ev.openingDateISO||"";' +
'document.getElementById("ee_closing_online").value=ev.closingDateISO||"";' +
'document.getElementById("ee_coach").value=ev.coachName||"";' +
'document.getElementById("ee_desc").value=ev.description||"";' +
'document.getElementById("ee_result").textContent="";' +
'document.getElementById("editEventModal").style.display="";' +
'document.getElementById("editEventModal").scrollIntoView({behavior:"smooth"});' +
'}' +
'function hideEditEventModal(){document.getElementById("editEventModal").style.display="none";}' +
'function loadReferralCodes(eventName){' +
'var box=document.getElementById("ee_referral_list");' +
'box.textContent="読み込み中...";' +
'google.script.run' +
'.withSuccessHandler(function(res){' +
'if(!res.success){box.textContent="読み込み失敗: "+res.error;return;}' +
'if(res.codes.length===0){box.innerHTML="<div class=\\"text-muted small\\">まだ紹介コードは登録されていません。</div>";return;}' +
'box.innerHTML=res.codes.map(function(c){' +
'var limitText=c.maxCount>0?(c.usedCount+"/"+c.maxCount+"件"):(c.usedCount+"件（上限なし）");' +
'return "<div class=\\"d-flex justify-content-between align-items-center border rounded p-2 mb-1\\">"+' +
'"<div><b>"+escHtml(c.code)+"</b>"+(c.referrerName?"（"+escHtml(c.referrerName)+"）":"")+" — "+limitText+"</div>"+' +
'"<button class=\\"btn btn-sm btn-outline-danger\\" onclick=\\"deleteReferralCodeFromUI(\'"+escHtml(c.code)+"\')\\">削除</button>"+' +
'"</div>";' +
'}).join("");' +
'})' +
'.withFailureHandler(function(e){box.textContent="読み込み失敗: "+e.message;})' +
'.getReferralCodesForEvent(eventName);' +
'}' +
'function addReferralCodeFromUI(){' +
'var eventName=document.getElementById("editEventModal").dataset.eventName;' +
'var code=document.getElementById("ee_referral_code_input").value.trim();' +
'var name=document.getElementById("ee_referral_name_input").value.trim();' +
'var max=document.getElementById("ee_referral_max_input").value;' +
'if(!code){alert("コードを入力してください。");return;}' +
'google.script.run' +
'.withSuccessHandler(function(res){' +
'if(res.success){' +
'document.getElementById("ee_referral_code_input").value="";' +
'document.getElementById("ee_referral_name_input").value="";' +
'document.getElementById("ee_referral_max_input").value="";' +
'loadReferralCodes(eventName);' +
'}else{alert(res.error);}' +
'})' +
'.withFailureHandler(function(e){alert(e.message);})' +
'.addReferralCode(eventName,code,name,max);' +
'}' +
'function deleteReferralCodeFromUI(code){' +
'var eventName=document.getElementById("editEventModal").dataset.eventName;' +
'if(!confirm("このコードを削除しますか？"))return;' +
'google.script.run' +
'.withSuccessHandler(function(res){' +
'if(res.success)loadReferralCodes(eventName);' +
'else alert(res.error);' +
'})' +
'.withFailureHandler(function(e){alert(e.message);})' +
'.deleteReferralCode(eventName,code);' +
'}' +
'function submitEditEvent(){' +
'var modal=document.getElementById("editEventModal");' +
'var appSheetName=modal.dataset.appSheetName;' +
'var isOnline=document.getElementById("ee_online_fields").style.display!=="none";' +
'var payload={appSheetName:appSheetName,' +
'name:document.getElementById("ee_name").value.trim(),' +
'coachName:document.getElementById("ee_coach").value.trim(),' +
'description:document.getElementById("ee_desc").value.trim()};' +
'if(isOnline){' +
'payload.channelUrl=document.getElementById("ee_channel_url").value.trim();' +
'payload.openingDate=document.getElementById("ee_opening_online").value.replace(/-/g,"/");' +
'payload.closingDate=document.getElementById("ee_closing_online").value.replace(/-/g,"/");' +
'}else{' +
'payload.eventDate=document.getElementById("ee_date").value.replace(/-/g,"/");' +
'payload.closingDate=document.getElementById("ee_closing").value.replace(/-/g,"/");' +
'payload.openingDate=document.getElementById("ee_opening").value.replace(/-/g,"/");' +
'payload.eventTime=document.getElementById("ee_time").value.trim();' +
'payload.venue=document.getElementById("ee_venue").value.trim();' +
'payload.meetingTime=document.getElementById("ee_meeting").value.trim();' +
'payload.courtType=document.getElementById("ee_court").value.trim();' +
'payload.fee=document.getElementById("ee_fee").value.trim();' +
'payload.items=document.getElementById("ee_items").value.trim();' +
'payload.isFreeEvent=document.getElementById("ee_is_free").checked;' +
'payload.closingDateTimeAt=document.getElementById("ee_closing_at").value;' +
'payload.resultAnnouncementDate=document.getElementById("ee_result_announcement").value;' +
'payload.confirmDeadline=document.getElementById("ee_deadline").value.trim();' +
'payload.confirmDeadlineAt=document.getElementById("ee_deadline_at").value;' +
'payload.lockerInfo=document.getElementById("ee_locker").value.trim();' +
'payload.facilityUrl=document.getElementById("ee_facility_url").value.trim();' +
'var eecapEl=document.getElementById("ee_capacity");payload.capacity=eecapEl?parseInt(eecapEl.value)||0:0;' +
'var eohEl=document.getElementById("ee_oubo_hidden");payload.ouboStatusHidden=!!(eohEl&&eohEl.checked);' +
'}' +
'var res=document.getElementById("ee_result");res.textContent="保存中...";' +
'google.script.run' +
'.withSuccessHandler(function(r){' +
'if(r.success){res.textContent="✅ 保存完了";setTimeout(function(){hideEditEventModal();loadEvents();},1000);}' +
'else{res.textContent="❌ "+r.error;}' +
'})' +
'.withFailureHandler(function(e){res.textContent="❌ "+e.message;})' +
'.updateEventDetails(payload);' +
'}' +

'function showTab(t){' +
'["events","broadcast","members","funnel","messages","terms"].forEach(function(n){' +
'document.getElementById("tab-"+n).style.display=t===n?"":"none";' +
'document.getElementById("tab-btn-"+n).classList.toggle("active",t===n);' +
'});' +
'if(t==="members"&&!membersLoaded){membersLoaded=true;loadMembers();}' +
'if(t==="messages"&&!messagesLoaded){messagesLoaded=true;loadMessageTemplates();}' +
'if(t==="terms"&&!termsLoaded){termsLoaded=true;loadTerms();}' +
'}' +

'function spin(on){document.getElementById("spinner").style.display=on?"":"none";}' +

'function loadEvents(){' +
'spin(true);' +
'google.script.run' +
'.withSuccessHandler(function(ev){spin(false);renderEvents(ev);})' +
'.withFailureHandler(function(e){spin(false);document.getElementById("eventList").innerHTML="<div class=\'col\'><div class=\'alert alert-danger\'>⚠ GASエラー: "+(e&&e.message?e.message:JSON.stringify(e))+"</div></div>";})' +
'.getEventsData();' +
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
'var editBtn="<button class=\'btn btn-sm btn-outline-primary py-0 px-2\' onclick=\'openEditEventModal("+idx+",event)\'>✏️ 詳細編集</button>";' +
'var btns="<div class=\'mt-2 d-flex gap-1\'>"+toggleBtn+(copyBtn?copyBtn:"")+editBtn+"</div>";' +
'div.innerHTML="<div class=\'card event-card h-100 border\' style=\'"+(isStopped?"opacity:0.55":"")+"\'  onclick=\'selectEvent("+idx+")\'>"+' +
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

'function showFunnelTab(t){' +
'["overview","bytype","daily","dropoff","user"].forEach(function(n){' +
'document.getElementById("ftab-"+n).style.display=t===n?"":"none";' +
'document.getElementById("ftab-btn-"+n).classList.toggle("active",t===n);' +
'});' +
'}' +

'function loadFunnel(){' +
'var days=parseInt(document.getElementById("funnelDays").value)||0;' +
'var active="overview";' +
'["overview","bytype","daily","dropoff"].forEach(function(n){' +
'if(document.getElementById("ftab-btn-"+n).classList.contains("active"))active=n;' +
'});' +
'var el=document.getElementById("ftab-"+active);' +
'el.innerHTML="<p class=\'text-center text-muted\'>集計中...</p>";' +
'var ok=function(d){' +
'if(active==="overview")el.innerHTML=renderFunnel(d);' +
'else if(active==="bytype")el.innerHTML=renderFunnelByType(d);' +
'else if(active==="daily")el.innerHTML=renderFunnelByDay(d);' +
'else if(active==="dropoff")el.innerHTML=renderDropoff(d);' +
'};' +
'var ng=function(e){el.innerHTML="<div class=\'alert alert-danger\'>エラー: "+(e&&e.message?e.message:String(e))+"</div>";};' +
'if(active==="overview")google.script.run.withSuccessHandler(ok).withFailureHandler(ng).getFunnelStats(days);' +
'else if(active==="bytype")google.script.run.withSuccessHandler(ok).withFailureHandler(ng).getFunnelByType(days);' +
'else if(active==="daily")google.script.run.withSuccessHandler(ok).withFailureHandler(ng).getFunnelByDay(days);' +
'else if(active==="dropoff")google.script.run.withSuccessHandler(ok).withFailureHandler(ng).getDropoffDetail(days);' +
'}' +

'function renderFunnel(stats){' +
'if(!stats||!stats.stages||!stats.stages.length)return"<p class=\'text-muted text-center\'>データがありません</p>";' +
'var label=stats.days>0?"直近"+stats.days+"日":"全期間";' +
'var max=stats.stages[0].count||1;' +
'var html="<div class=\'card p-3\'><p class=\'text-muted small mb-3\'>集計期間: "+label+"</p>";' +
'stats.stages.forEach(function(s,i){' +
'var p=max>0?Math.round(s.count/max*100):0;' +
'html+="<div class=\'mb-4\'><div class=\'d-flex justify-content-between mb-1\'><span class=\'fw-bold\'>"+s.name+"</span><span>"+s.count+"人";' +
'if(s.rate!==null)html+=" <span class=\'text-muted small\'>(起動の"+s.rate+"%)</span>";' +
'html+="</span></div><div class=\'progress mb-1\' style=\'height:24px\'><div class=\'progress-bar\' style=\'width:"+p+"%\'>"+s.count+"人</div></div>";' +
'if(i<stats.stages.length-1){var drop=s.count-stats.stages[i+1].count;var dp=s.count>0?Math.round(drop/s.count*100):0;html+="<div class=\'text-muted small\'>↓ 離脱: "+drop+"人 ("+dp+"%)</div>";}' +
'html+="</div>";});' +
'html+="</div>";return html;' +
'}' +

'function renderFunnelByType(stats){' +
'if(!stats||!stats.segments||!stats.segments.length)return"<p class=\'text-muted text-center\'>データがありません</p>";' +
'var label=stats.days>0?"直近"+stats.days+"日":"全期間";' +
'var names={"online":"💻 オンライン","offline":"📍 オフライン","online/video":"🎥 動画相談","online/text":"📝 文章相談","通常":"📱 通常アクセス"};' +
'var html="<p class=\'text-muted small mb-2\'>集計期間: "+label+"</p><div class=\'table-responsive\'><table class=\'table table-sm table-hover\'><thead class=\'table-light\'><tr><th>アクセス種別</th><th class=\'text-end\'>起動</th><th class=\'text-end\'>ボタン押下</th><th class=\'text-end\'>応募完了</th><th class=\'text-end\'>完了率</th></tr></thead><tbody>";' +
'stats.segments.forEach(function(s){' +
'var n=names[s.segment]||s.segment;' +
'var badge=s.completeRate>=50?"bg-success":s.completeRate>=30?"bg-warning text-dark":"bg-danger";' +
'html+="<tr><td>"+n+"</td><td class=\'text-end\'>"+s.open+"人</td><td class=\'text-end\'>"+s.press+"人 <span class=\'text-muted small\'>("+s.pressRate+"%)</span></td><td class=\'text-end\'>"+s.complete+"人</td><td class=\'text-end\'><span class=\'badge "+badge+"\'>"+s.completeRate+"%</span></td></tr>";' +
'});' +
'html+="</tbody></table></div>";return html;' +
'}' +

'function renderFunnelByDay(stats){' +
'if(!stats||!stats.days||!stats.days.length)return"<p class=\'text-muted text-center\'>データがありません</p>";' +
'var colors=["#0d6efd","#6c757d","#198754","#ffc107"];' +
'var label=stats.totalDays>0?"直近"+stats.totalDays+"日":"全期間";' +
'var maxes=stats.stages.map(function(_,si){return Math.max.apply(null,stats.days.map(function(d){return d.counts[si]||0;}))||1;});' +
'var html="<p class=\'text-muted small mb-2\'>集計期間: "+label+"</p><div class=\'table-responsive\'><table class=\'table table-sm\'><thead class=\'table-light\'><tr><th>日付</th>";' +
'stats.stages.forEach(function(s){html+="<th class=\'text-end\'>"+s+"</th>";});' +
'html+="</tr></thead><tbody>";' +
'stats.days.forEach(function(d){' +
'html+="<tr><td>"+d.date+"</td>";' +
'd.counts.forEach(function(c,i){' +
'var p=Math.round(c/maxes[i]*60);' +
'html+="<td class=\'text-end\' style=\'min-width:80px\'><div style=\'display:inline-flex;align-items:center;gap:4px;width:100%;justify-content:flex-end\'><div style=\'background:"+colors[i]+"33;border-radius:3px;height:12px;width:"+p+"px;min-width:2px\'></div>"+c+"</div></td>";' +
'});' +
'html+="</tr>";' +
'});' +
'html+="</tbody></table></div>";return html;' +
'}' +

'function renderDropoff(stats){' +
'var label=stats.days>0?"直近"+stats.days+"日":"全期間";' +
'var html="<p class=\'text-muted small mb-3\'>集計期間: "+label+"</p>";' +
'html+="<div class=\'row g-3 mb-3\'>";' +
'html+="<div class=\'col-sm-4\'><div class=\'card p-3 text-center\'><div class=\'h3 text-primary mb-0\'>"+stats.pressed+"</div><div class=\'small text-muted\'>ボタンを押した人</div></div></div>";' +
'html+="<div class=\'col-sm-4\'><div class=\'card p-3 text-center\'><div class=\'h3 text-success mb-0\'>"+stats.completed+"</div><div class=\'small text-muted\'>応募完了した人</div></div></div>";' +
'html+="<div class=\'col-sm-4\'><div class=\'card p-3 text-center\'><div class=\'h3 text-danger mb-0\'>"+stats.total+"</div><div class=\'small text-muted\'>途中で止まった人</div></div></div>";' +
'html+="</div>";' +
'if(stats.byEvent&&stats.byEvent.length){' +
'html+="<div class=\'card p-3 mb-3\'><strong class=\'d-block mb-2\'>止まった人が選んでいたイベント</strong><table class=\'table table-sm mb-0\'><thead><tr><th>イベント</th><th class=\'text-end\'>人数</th></tr></thead><tbody>";' +
'stats.byEvent.forEach(function(r){html+="<tr><td>"+r.ev+"</td><td class=\'text-end\'>"+r.count+"人</td></tr>";});' +
'html+="</tbody></table></div>";' +
'}' +
'if(stats.dropoffs&&stats.dropoffs.length){' +
'html+="<div class=\'card p-3\'><strong class=\'d-block mb-2\'>対象ユーザー詳細（最大50件）</strong><div class=\'table-responsive\'><table class=\'table table-sm mb-0\'><thead><tr><th>日時</th><th>選択イベント</th><th>相談種別</th></tr></thead><tbody>";' +
'stats.dropoffs.forEach(function(d){' +
'var c=d.consult==="video"?"🎥 動画":d.consult==="text"?"📝 文章":"-";' +
'html+="<tr><td>"+d.ts+"</td><td>"+(d.events||"-")+"</td><td>"+c+"</td></tr>";' +
'});' +
'html+="</tbody></table></div></div>";' +
'}' +
'return html;' +
'}' +

'function loadUserTimeline(){' +
'var uid=document.getElementById("userSearchId").value.trim();' +
'if(!uid){alert("User IDを入力してください");return;}' +
'var el=document.getElementById("userTimelineResult");' +
'el.innerHTML="<p class=\'text-muted\'>検索中...</p>";' +
'google.script.run' +
'.withSuccessHandler(function(rows){el.innerHTML=renderUserTimeline(uid,rows);})' +
'.withFailureHandler(function(e){el.innerHTML="<div class=\'alert alert-danger\'>エラー: "+(e&&e.message?e.message:String(e))+"</div>";})' +
'.getUserTimeline(uid);' +
'}' +

'function renderUserTimeline(uid,rows){' +
'if(!rows||!rows.length)return"<p class=\'text-muted\'>このUser IDのデータが見つかりません</p>";' +
'var icons={"LIFF起動":"🔵","応募ボタン押下":"🟡","LIFF応募":"✅","動画受信":"🎥"};' +
'var html="<p class=\'small text-muted mb-3\'>User ID: "+uid+" ("+rows.length+"件)</p>";' +
'html+="<div style=\'padding-left:24px;border-left:2px solid #dee2e6\'>";' +
'rows.forEach(function(r){' +
'html+="<div class=\'mb-3\' style=\'position:relative\'>";' +
'html+="<div style=\'position:absolute;left:-32px;font-size:16px\'>"+(icons[r.action]||"⚪")+"</div>";' +
'html+="<div class=\'small text-muted\'>"+r.ts+"</div>";' +
'html+="<div class=\'fw-bold\'>"+r.action+"</div>";' +
'if(r.eventId)html+="<div class=\'small\'>"+r.eventId+"</div>";' +
'if(r.detail)html+="<div class=\'small text-muted\'>"+r.detail+"</div>";' +
'html+="</div>";' +
'});' +
'html+="</div>";return html;' +
'}' +

'function runSetupFunnelSheet(){' +
'google.script.run' +
'.withSuccessHandler(function(){alert("スプレッドシートに「ファネル集計」シートを作成しました。");})' +
'.withFailureHandler(function(e){alert("エラー: "+(e&&e.message?e.message:String(e)));})' +
'.setupFunnelSheet();' +
'}' +

'function toggleLookerInfo(){' +
'var el=document.getElementById("lookerInfo");' +
'el.style.display=el.style.display==="none"?"":"none";' +
'}' +

'function loadMessageTemplates(){' +
'document.getElementById("msgTemplatesList").innerHTML="<p class=\'text-muted text-center mt-4\'>読み込み中...</p>";' +
'google.script.run' +
'.withSuccessHandler(function(data){msgTemplatesData=data;renderMessageTemplates(data);})' +
'.withFailureHandler(function(e){document.getElementById("msgTemplatesList").innerHTML="<div class=\'alert alert-danger\'>エラー: "+(e&&e.message?e.message:String(e))+"</div>";})' +
'.getMessageTemplates();' +
'}' +

'function escHtml(s){return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");}' +

'function renderMessageTemplates(data){' +
'var html="<div class=\'d-flex justify-content-end mb-3\'><button class=\'btn btn-success\' onclick=\'saveAllTemplates()\'>💾 すべて保存</button> <span class=\'small text-muted align-self-center ms-2\' id=\'tmplAllResult\'></span></div>";' +
'(data.templates||[]).forEach(function(t,i){' +
'html+="<div class=\'card p-3 mb-3\'>";' +
'html+="<div class=\'d-flex justify-content-between align-items-start mb-2 flex-wrap gap-1\'>";' +
'html+="<label class=\'fw-bold mb-0\'>"+t.label+"</label>";' +
'if(t.vars&&t.vars.length)html+="<span class=\'text-muted small\'>変数: "+t.vars.map(function(v){return "{"+v+"}";}).join(" ")+"</span>";' +
'html+="</div>";' +
'html+="<textarea class=\'form-control mb-2\' rows=\'4\' id=\'tmpl_"+i+"\' oninput=\'markTemplateDirty("+i+")\'>"+escHtml(t.value)+"</textarea>";' +
'html+="<button class=\'btn btn-sm btn-primary\' id=\'tmplBtn_"+i+"\' onclick=\'saveTemplate("+i+")\'>保存</button> ";' +
'html+="<button class=\'btn btn-sm btn-outline-secondary\' onclick=\'testSendTemplate("+i+")\'>🧪 自分に試し送信</button> ";' +
'html+="<span class=\'small text-muted\' id=\'tmplResult_"+i+"\'></span>";' +
'html+="</div>";' +
'});' +
'document.getElementById("msgTemplatesList").innerHTML=html;' +
'}' +

'function markTemplateDirty(i){' +
'var btn=document.getElementById("tmplBtn_"+i);' +
'btn.className="btn btn-sm btn-primary";btn.textContent="保存";' +
'document.getElementById("tmplResult_"+i).textContent="";' +
'}' +

'function saveTemplate(i){' +
'var t=msgTemplatesData.templates[i];' +
'var val=document.getElementById("tmpl_"+i).value;' +
'var btn=document.getElementById("tmplBtn_"+i);' +
'var resEl=document.getElementById("tmplResult_"+i);' +
'resEl.textContent="保存中...";' +
'var payload={};payload[t.key]=val;' +
'google.script.run' +
'.withSuccessHandler(function(r){' +
'if(r.success){btn.className="btn btn-sm btn-success";btn.textContent="✅ 保存済み";resEl.textContent="";}' +
'else{resEl.textContent="エラー: "+r.error;}' +
'})' +
'.withFailureHandler(function(e){resEl.textContent="エラー: "+(e&&e.message?e.message:String(e));})' +
'.saveMessageTemplates(payload);' +
'}' +

'function loadTerms(){' +
'var keys=["tos_register","tos_offline","tos_online","media_offline_free","media_offline_paid","media_online"];' +
'keys.forEach(function(k){var el=document.getElementById(k);if(el)el.value="読み込み中...";});' +
'google.script.run' +
'.withSuccessHandler(function(data){' +
'keys.forEach(function(k){var el=document.getElementById(k);if(el)el.value=data[k]||"";});' +
'})' +
'.withFailureHandler(function(e){console.error("loadTerms error",e);})' +
'.getTermsContent();' +
'}' +

'function initTerms(){' +
'if(!confirm("現在の入力内容をデフォルト文で上書きします。よろしいですか？"))return;' +
'var keys=["tos_register","tos_offline","tos_online","media_offline_free","media_offline_paid","media_online"];' +
'keys.forEach(function(k){var el=document.getElementById(k);if(el)el.value="読み込み中...";});' +
'google.script.run' +
'.withSuccessHandler(function(data){' +
'keys.forEach(function(k){var el=document.getElementById(k);if(el)el.value=data[k]||"";});' +
'document.getElementById("termsResult").textContent="デフォルト文を読み込みました。「💾 すべて保存」で確定してください。";' +
'})' +
'.withFailureHandler(function(e){console.error("initTerms error",e);})' +
'.getDefaultTermsContent();' +
'}' +

'function saveAllTerms(){' +
'var keys=["tos_register","tos_offline","tos_online","media_offline_free","media_offline_paid","media_online"];' +
'var payload={};' +
'keys.forEach(function(k){var el=document.getElementById(k);if(el)payload[k]=el.value;});' +
'var res=document.getElementById("termsResult");' +
'res.textContent="保存中...";' +
'google.script.run' +
'.withSuccessHandler(function(r){res.textContent=r.success?"✅ 保存しました":"❌ "+r.error;})' +
'.withFailureHandler(function(e){res.textContent="❌ "+(e&&e.message?e.message:String(e));})' +
'.saveTermsContent(payload);' +
'}' +

'function saveAllTemplates(){' +
'var payload={};' +
'(msgTemplatesData.templates||[]).forEach(function(t,i){' +
'var el=document.getElementById("tmpl_"+i);' +
'if(el)payload[t.key]=el.value;' +
'});' +
'var resEl=document.getElementById("tmplAllResult");' +
'resEl.textContent="保存中...";' +
'google.script.run' +
'.withSuccessHandler(function(r){' +
'if(r.success){' +
'resEl.textContent="✅ すべて保存しました";' +
'(msgTemplatesData.templates||[]).forEach(function(t,i){' +
'var btn=document.getElementById("tmplBtn_"+i);' +
'if(btn){btn.className="btn btn-sm btn-success";btn.textContent="✅ 保存済み";}' +
'var r2=document.getElementById("tmplResult_"+i);if(r2)r2.textContent="";' +
'});' +
'}else{resEl.textContent="エラー: "+r.error;}' +
'})' +
'.withFailureHandler(function(e){resEl.textContent="エラー: "+(e&&e.message?e.message:String(e));})' +
'.saveMessageTemplates(payload);' +
'}' +

'function testSendTemplate(i){' +
'var val=document.getElementById("tmpl_"+i).value;' +
'var resEl=document.getElementById("tmplResult_"+i);' +
'resEl.textContent="送信中...";' +
'google.script.run' +
'.withSuccessHandler(function(r){resEl.textContent=r.success?"✅ 試し送信しました（自分のLINEを確認してください）":"エラー: "+r.error;})' +
'.withFailureHandler(function(e){resEl.textContent="エラー: "+(e&&e.message?e.message:String(e));})' +
'.testSendMessageTemplate(val);' +
'}' +

'function showLiffLinks(){' +
'if(!dashConfig.liffId){alert("LIFF_IDが未設定です。スクリプトプロパティを確認してください。");return;}' +
'var base="https://liff.line.me/"+dashConfig.liffId;' +
'var links=[' +
'{label:"📝 プロフィール登録",desc:"初回登録専用。既存会員が開いても登録フォームが表示されます。",url:base+"?page=register"},' +
'{label:"✏️ プロフィール修正",desc:"既存会員向け。プロフィール編集欄が開いた状態で表示されます。",url:base+"?page=profile"},' +
'{label:"📍 オフライン応募",desc:"オフラインイベントのみ表示されます。",url:base+"?type=offline"},' +
'{label:"💻 オンライン応募",desc:"オンラインイベントのみ表示されます。",url:base+"?type=online"},' +
'{label:"🎥 動画相談",desc:"オンラインのみ・動画相談を自動選択。",url:base+"?type=online&consult=video"},' +
'{label:"📝 文章相談",desc:"オンラインのみ・文章相談を自動選択。",url:base+"?type=online&consult=text"},' +
'{label:"🏆 大会応募",desc:"大会イベントのみ表示されます。",url:base+"?type=tournament"},' +
'];' +
'var html=links.map(function(l){' +
'return "<div class=\'mb-3\'><div class=\'fw-bold mb-1\'>"+l.label+"</div><div class=\'text-muted small mb-1\'>"+l.desc+"</div><div class=\'d-flex gap-2\'><input type=\'text\' class=\'form-control form-control-sm\' value=\'"+l.url+"\' readonly onclick=\'this.select()\'><button class=\'btn btn-sm btn-outline-primary flex-shrink-0\' data-url=\'"+l.url+"\' onclick=\'cpLink(this)\'>コピー</button></div></div>";' +
'}).join("");' +
'document.getElementById("liffLinksBody").innerHTML=html;' +
'document.getElementById("liffLinksModal").style.display="flex";' +
'}' +
'function sendLiffLinksLine(){' +
'var btn=document.getElementById("sendLiffLinksBtn");' +
'btn.disabled=true;btn.textContent="送信中...";' +
'google.script.run' +
'.withSuccessHandler(function(r){' +
'if(r.success){btn.textContent="✅ 送信しました！";setTimeout(function(){btn.textContent="📨 LINEに送る";btn.disabled=false;},3000);}' +
'else{alert("エラー: "+r.error);btn.textContent="📨 LINEに送る";btn.disabled=false;}' +
'})' +
'.withFailureHandler(function(e){alert("エラー: "+e.message);btn.textContent="📨 LINEに送る";btn.disabled=false;})' +
'.sendLiffLinksToStaff();' +
'}' +
'function hideLiffLinks(){document.getElementById("liffLinksModal").style.display="none";}' +
'function cpLink(btn){' +
'var url=btn.getAttribute("data-url");' +
'if(navigator.clipboard){navigator.clipboard.writeText(url).then(function(){var t=btn.textContent;btn.textContent="✅ コピー済";setTimeout(function(){btn.textContent=t;},1500);});}' +
'else{prompt("コピーしてください:",url);}' +
'}' +

'function doToggleStatus(idx,e){' +
'e.stopPropagation();' +
'var ev=eventsData[idx];' +
'var isStopped=ev.status==="停止";' +
'var msg=isStopped?"「"+ev.name+"」を再開しますか？\\n\\nLIFFフォームに再表示されます。":"「"+ev.name+"」を停止しますか？\\n\\nLIFFフォームから非表示になります。";' +
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
'"<td class=\'text-center small\'>"+(ap.shootingConsent||"")+"</td>"+' +
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

'var PRINT_COLS=[' +
'{key:"name",label:"名前",def:true},' +
'{key:"furigana",label:"フリガナ",def:true},' +
'{key:"age",label:"年齢",def:true},' +
'{key:"gender",label:"性別",def:true},' +
'{key:"tennisLevel",label:"テニスレベル",def:true},' +
'{key:"tennisHistory",label:"テニス歴",def:false},' +
'{key:"tennisFreq",label:"プレー頻度",def:false},' +
'{key:"tennisArea",label:"テニスエリア",def:false},' +
'{key:"coachKnowledge",label:"コーチ認知",def:false},' +
'{key:"result",label:"当落結果",def:false},' +
'{key:"confirmation",label:"参加確認",def:false},' +
'{key:"appliedAt",label:"応募日時",def:false},' +
'{key:"email",label:"メールアドレス",def:false},' +
'{key:"phone",label:"電話番号",def:false},' +
'];' +

'function openPrintModal(){' +
'var checked=Array.from(document.querySelectorAll("#applicantBody .row-check:checked"));' +
'if(checked.length===0){alert("表に含める人をチェックしてください。");return;}' +
'var html="";' +
'PRINT_COLS.forEach(function(c,i){' +
'html+="<label style=\'display:block;margin-bottom:6px\'><input type=\'checkbox\' id=\'pc_"+i+"\' "+(c.def?"checked":"")+"> "+c.label+"</label>";' +
'});' +
'document.getElementById("printColChecks").innerHTML=html;' +
'document.getElementById("printModal").style.display="block";' +
'}' +

'function buildPrintTable(){' +
'var checkedIds=Array.from(document.querySelectorAll("#applicantBody .row-check:checked")).map(function(cb){return cb.dataset.userid;});' +
'var cols=PRINT_COLS.filter(function(c,i){return document.getElementById("pc_"+i)&&document.getElementById("pc_"+i).checked;});' +
'if(cols.length===0){alert("表示する列を1つ以上選んでください。");return;}' +
'var rows=allApplicantsData.filter(function(a){return checkedIds.indexOf(a.userId)!==-1;});' +
'var eventName=currentEvent?currentEvent.name:"応募者";' +
'var thead="<tr>"+cols.map(function(c){return"<th>"+c.label+"</th>";}).join("")+"</tr>";' +
'var tbody=rows.map(function(a,i){' +
'return"<tr>"' +
'+cols.map(function(c){return"<td>"+(a[c.key]||"")+"</td>";}).join("")' +
'+"</tr>";' +
'}).join("");' +
'var html="<!DOCTYPE html><html lang=\'ja\'><head><meta charset=\'utf-8\'><title>"+eventName+" 応募者一覧</title>"' +
'+"<style>body{font-family:sans-serif;font-size:13px;padding:20px}h2{font-size:16px;margin-bottom:12px}"' +
'+"table{border-collapse:collapse;width:100%}th,td{border:1px solid #ccc;padding:6px 8px;text-align:left}"' +
'+"th{background:#f0f0f0;font-weight:bold}tr:nth-child(even){background:#f9f9f9}"' +
'+"#printBtn{margin-bottom:16px;padding:8px 20px;background:#0d6efd;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:14px}"' +
'+"@media print{#printBtn{display:none}}"' +
'+"</style></head><body>"' +
'+"<button id=\'printBtn\' onclick=\'window.print()\'>🖨 印刷 / PDFで保存</button>"' +
'+"<h2>"+eventName+" — 応募者一覧（"+rows.length+"名）</h2>"' +
'+"<table><thead>"+thead+"</thead><tbody>"+tbody+"</tbody></table>"' +
'+"</body></html>";' +
'var w=window.open("","_blank");' +
'if(w){w.document.write(html);w.document.close();}else{alert("ポップアップがブロックされています。ブラウザの設定でこのサイトのポップアップを許可してください。");}' +
'document.getElementById("printModal").style.display="none";' +
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
'<p class="text-muted small mb-3">各リンクをコピーして配布するか、LINEに一括送信できます。</p>' +
'<div id="liffLinksBody"></div>' +
'<div class="border-top pt-3 mt-2">' +
'<button id="sendLiffLinksBtn" class="btn btn-success w-100" onclick="sendLiffLinksLine()">📨 LINEに送る（STAFF_USER_ID宛）</button>' +
'</div>' +
'</div></div>' +
'</body></html>';
}

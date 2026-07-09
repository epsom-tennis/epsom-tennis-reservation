// LIFF応募フォーム
// doGet(?page=liff) からルーティングされる。
// スクリプトプロパティ LIFF_ID を設定してから使用する。

// LIFFのIDトークンをLINEのAPIで検証し、userIdが一致するか確認する
function verifyLiffToken_(idToken, expectedUserId) {
  if (!idToken) return false;
  try {
    const channelId = getProp('LIFF_ID').split('-')[0]; // "2010231562-h4nq4P3s" → "2010231562"
    const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
      method: 'POST',
      contentType: 'application/x-www-form-urlencoded',
      payload: 'id_token=' + encodeURIComponent(idToken) + '&client_id=' + encodeURIComponent(channelId),
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) return false;
    const body = JSON.parse(res.getContentText());
    return body.sub === expectedUserId;
  } catch (err) {
    Logger.log('verifyLiffToken_ error: ' + err.toString());
    return false;
  }
}

// GitHub Pages版LIFFにリダイレクト（GAS経由では LINE webview で動作しないため）
function getLiffPage() {
  const url = 'https://epsom-tennis.github.io/epsom-tennis-reservation/liff/';
  return HtmlService.createHtmlOutput(
    '<html><head><meta http-equiv="refresh" content="0; url=' + url + '"></head>' +
    '<body>リダイレクト中... <a href="' + url + '">こちらをクリック</a></body></html>'
  ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 会員データを返す（追加参加者情報も含む）
function getMemberData(userId) {
  try {
    const membersSheet = getSheet(SHEET.MEMBERS);
    if (!membersSheet || membersSheet.getLastRow() <= 1) return null;
    const data = membersSheet.getDataRange().getValues();
    let primary = null;
    const additional = [];
    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][1]);
      if (rowId === userId) {
        primary = extractMemberRow_(data[i]);
      } else if (rowId.startsWith(userId + '_p')) {
        additional.push(extractMemberRow_(data[i]));
      }
    }
    if (!primary) return null;
    primary.additionalParticipants = additional;
    return primary;
  } catch (err) {
    Logger.log('getMemberData error: ' + err.toString());
    return null;
  }
}

// Sheetsが日付型に自動変換した値をYYYY-MM-DD文字列に戻す
function formatDateValue_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  return String(v);
}

// 複数オフラインイベントの当落発表日からテンプレート用の文字列を生成する
// 例: [Date(7/15), Date(7/15)] → "7月15日頃に"  /  [null] → "後日"
function buildOfflineResultDateStr_(dates) {
  const formatted = dates
    .filter(function(d) { return d instanceof Date; })
    .map(function(d) { return Utilities.formatDate(d, 'Asia/Tokyo', 'M月d日'); })
    .filter(function(v, i, a) { return a.indexOf(v) === i; });
  if (formatted.length === 0) return '後日';
  return formatted.join('・') + '頃に';
}

function extractMemberRow_(row) {
  return {
    name:          String(row[4]  || ''),
    age:           String(row[6]  || ''),
    gender:        String(row[7]  || ''),
    tennisLevel:   String(row[8]  || ''),
    email:         String(row[9]  || ''),
    phone:         String(row[10] || ''),
    furigana:      String(row[11] || ''),
    emergency:     String(row[12] || ''),
    tennisFreq:    String(row[13] || ''),
    tennisHistory: String(row[14] || ''),
    tennisArea:    String(row[15] || ''),
    tennisEnv:     String(row[16] || ''),
    birthDate:     formatDateValue_(row[17]),
    prefecture:    String(row[18] || ''),
  };
}

// LIFF応募フォームの送信処理（クライアントから呼び出す）
function submitLiffApplication(data) {
  try {
    // LIFFトークンでuserIdの正当性を検証
    if (!verifyLiffToken_(data.idToken, data.userId)) {
      return { success: false, error: '認証に失敗しました。LINEアプリから再度開き直してください。' };
    }

    const membersSheet = getSheet(SHEET.MEMBERS);
    const ss = SpreadsheetApp.openById(getProp('SPREADSHEET_ID'));
    const allServerEvents = getAllEvents(); // 開催日・当落発表日はサーバー側から取得する
    const appliedNames = [];
    const appliedOnlineNames = [];
    const appliedOfflineNames = [];
    const appliedOfflineResultDates = []; // appliedOfflineNamesと並行：当落発表予定日(Date or null)
    const appliedOfflineEventDates  = []; // appliedOfflineNamesと並行：開催日(Date or null)
    const appliedTournamentNames      = []; // 大会：先着順で確定したイベント名
    const appliedTournamentEventDates = []; // 大会：開催日(Date or null)
    const appliedParticipantNames = []; // 新規応募した参加者のフルネーム

    // 主参加者 + 追加参加者をまとめて処理
    const participants = [{ ...data, _suffix: '', _participantNum: 1 }];
    (data.additionalParticipants || []).forEach(function(p, idx) {
      participants.push({ ...p, userId: data.userId + '_p' + (idx + 2), _suffix: '（追加' + (idx + 2) + '人目）', _participantNum: idx + 2 });
    });

    for (const p of participants) {
      const fullName = ((p.familyName || '') + ' ' + (p.givenName || '')).trim();
      // フリガナはひらがな入力でも全角カタカナに統一、電話番号はハイフン等を除去して半角数字のみに統一する
      const furigana = toFullWidthKatakana_(((p.familyNameKana || '') + ' ' + (p.givenNameKana || '')).trim());
      const areaStr  = (p.tennisArea || []).join('・');
      const envStr   = (p.tennisEnv  || []).join('・');
      const pUserId  = p.userId;
      const phone     = normalizePhone_(p.phone || '');
      const emergency = normalizePhone_(p.emergency || '');

      // 会員マスタを更新または新規追加
      const membersData = membersSheet.getDataRange().getValues();
      let memberRow = -1;
      for (let i = 1; i < membersData.length; i++) {
        if (String(membersData[i][1]) === pUserId) { memberRow = i + 1; break; }
      }
      if (memberRow > 0) {
        membersSheet.getRange(memberRow, 5).setValue(fullName);
        membersSheet.getRange(memberRow, 6).setValue(new Date());
        membersSheet.getRange(memberRow, 7).setValue(p.age || '');
        membersSheet.getRange(memberRow, 8).setValue(p.gender || '');
        membersSheet.getRange(memberRow, 9).setValue(p.tennisLevel || '');
        membersSheet.getRange(memberRow, 10).setValue(p.email || '');
        // 電話番号は数値扱いされると頭の0が消えるため、書式をテキスト固定してから書き込む
        membersSheet.getRange(memberRow, 11).setNumberFormat('@').setValue(phone);
        membersSheet.getRange(memberRow, 12).setValue(furigana);
        membersSheet.getRange(memberRow, 13).setNumberFormat('@').setValue(emergency);
        membersSheet.getRange(memberRow, 14).setValue(p.tennisFreq || '');
        membersSheet.getRange(memberRow, 15).setValue(p.tennisHistory || '');
        membersSheet.getRange(memberRow, 16).setValue(areaStr);
        membersSheet.getRange(memberRow, 17).setValue(envStr);
        membersSheet.getRange(memberRow, 18).setNumberFormat('@').setValue(p.birthDate || '');
        membersSheet.getRange(memberRow, 19).setValue(p.prefecture || '');
      } else {
        membersSheet.appendRow([
          new Date(), pUserId, '', 'LIFF登録', fullName, new Date(),
          p.age || '', p.gender || '', p.tennisLevel || '', p.email || '', phone, furigana,
          emergency, p.tennisFreq || '', p.tennisHistory || '', areaStr, envStr,
          p.birthDate || '', p.prefecture || '',
        ]);
        // appendRowは数値扱いで書き込まれるため、電話番号と生年月日はテキスト書式で再書き込みする
        const newMemberRow = membersSheet.getLastRow();
        membersSheet.getRange(newMemberRow, 11).setNumberFormat('@').setValue(phone);
        membersSheet.getRange(newMemberRow, 13).setNumberFormat('@').setValue(emergency);
        if (p.birthDate) membersSheet.getRange(newMemberRow, 18).setNumberFormat('@').setValue(p.birthDate);
      }

      // 各選択イベントの当落シートに応募行を追加
      let pHasNewApply = false;
      for (const ev of (data.selectedEvents || [])) {
        const isTournament = (ev.eventType || 'オフライン') === '大会';

        // 大会は追加参加者（_p2以降）をスキップ。ペアは1行で2名分として扱う
        if (isTournament && p._participantNum > 1) continue;

        // participantNums が指定されている場合はその参加者のみ応募
        if (ev.participantNums && ev.participantNums.length > 0 && !ev.participantNums.includes(p._participantNum)) {
          continue;
        }
        const resultSheet = ss.getSheetByName(ev.resultSheetName);
        if (!resultSheet) continue;
        const resultData = resultSheet.getDataRange().getValues();
        const isOnline = (ev.eventType || 'オフライン') === 'オンライン';
        let already = false;
        for (let i = 1; i < resultData.length; i++) {
          if (String(resultData[i][1]) === pUserId) { already = true; break; }
        }
        if (already && !isOnline) continue; // オンラインは常時募集のため重複応募を許可

        // 大会：先着順チェック（定員はサーバー側のデータを使い、クライアント偽装を防ぐ）
        if (isTournament) {
          const serverEvForCap = allServerEvents.find(function(e) { return e.resultSheetName === ev.resultSheetName; });
          const capacity = (serverEvForCap && serverEvForCap.capacity) || 0;
          if (capacity > 0) {
            let currentCount = 0;
            for (let j = 1; j < resultData.length; j++) {
              const form = String(resultData[j][10] || ''); // K列：参加形式
              currentCount += (form === 'ペア') ? 2 : 1;
            }
            const needed = (ev.participantForm === 'ペア') ? 2 : 1;
            if (currentCount + needed > capacity) {
              return { success: false, error: `「${ev.name}」は定員に達しているため応募できません。` };
            }
          }
        }

        const evCoach  = ev.coachKnowledge || data.coachKnowledge || '';
        const evSrcStr = (ev.eventSource   || data.eventSource   || []).join('・');
        // 大会は「参加応募のきっかけ」質問なし
        const evRsnStr = isTournament ? '' : (ev.applyReason || data.applyReason || []).join('・');
        const isVideo = isOnline && (data.onlineConsultType || '') === 'video';
        const onlineConsultPhoneNorm = normalizePhone_(data.onlineConsultPhone || '');
        // J列「参加確認」はオンライン・オフライン共通の基本列（応募時点では空欄、当落確定後に運用される）
        resultSheet.appendRow(
          [fullName, pUserId, '', '', '', evCoach, evSrcStr, evRsnStr, new Date(), '']
          .concat(isOnline ? [
            data.onlineBroadcastName || '',  // K: 配信名
            data.onlineConcern       || '',  // L: お悩み内容
            data.onlineConsultType   || '',  // M: 相談方法
            data.onlinePhoneConsult  || '',  // N: 電話相談希望
            onlineConsultPhoneNorm,          // O: 電話番号
            isVideo ? '待ち' : '',           // P: 動画状態
            '',                              // Q: 動画URL
            '未確認',                        // R: 対応状況（スタッフ管理用・未確認/確認中/回答済）
          ] : isTournament ? [
            ev.participantForm  || '',       // K: 参加形式（1人/ペア）
            ev.pairPartnerName  || '',       // L: ペア相手名
          ] : [data.shootingConsent || ''])  // K: 撮影可否（オフライン有料イベントのみ入力）
        );
        // 大会：先着順のため応募時点で当選確定・通知済みとして記録する
        if (isTournament) {
          const newRow = resultSheet.getLastRow();
          resultSheet.getRange(newRow, 3).setValue('当選');   // C: 結果
          resultSheet.getRange(newRow, 4).setValue('済');     // D: 送信済み
          resultSheet.getRange(newRow, 5).setValue(new Date()); // E: 送信日時
        }
        if (isOnline && onlineConsultPhoneNorm) {
          // appendRowは数値扱いで書き込まれるため、電話番号列（O列=15）だけ書式をテキスト固定して再書き込みする
          resultSheet.getRange(resultSheet.getLastRow(), 15).setNumberFormat('@').setValue(onlineConsultPhoneNorm);
        }
        logAction(pUserId, 'LIFF応募', ev.resultSheetName.replace('_当落', ''), fullName);
        pHasNewApply = true;
        if (!appliedNames.includes(ev.name)) {
          appliedNames.push(ev.name);
          if (isOnline) {
            appliedOnlineNames.push(ev.name);
          } else if (isTournament) {
            const serverEvT = allServerEvents.find(function(e) { return e.resultSheetName === ev.resultSheetName; });
            appliedTournamentNames.push(ev.name);
            appliedTournamentEventDates.push(serverEvT ? (serverEvT.eventDate || null) : null);
          } else {
            // 開催日・当落発表日はクライアントデータには含まれないためサーバー側から取得する
            const serverEv = allServerEvents.find(function(e) { return e.resultSheetName === ev.resultSheetName; });
            appliedOfflineNames.push(ev.name);
            appliedOfflineResultDates.push(serverEv ? (serverEv.resultAnnouncementDate || null) : null);
            appliedOfflineEventDates.push(serverEv ? (serverEv.eventDate || null) : null);
          }
        }
      }
      // 1件でも新規応募があれば参加者名を記録
      if (pHasNewApply && fullName && !appliedParticipantNames.includes(fullName)) {
        appliedParticipantNames.push(fullName);
      }
    }

    const selectedCount = (data.selectedEvents || []).length;
    if (selectedCount > 0 && appliedNames.length === 0) {
      return { success: false, error: '選択したイベントはすでに応募済みです。' };
    }

    let _pushResult = null;

    if (appliedNames.length > 0) {
      const namesPart = appliedParticipantNames.length > 0
        ? appliedParticipantNames.map(n => n + ' 様').join('、')
        : '';
      const msgParts = [renderTemplate_(getMsgTemplate_('header_apply'), { names: namesPart })];

      // 大会：先着順で参加確定
      if (appliedTournamentNames.length > 0) {
        msgParts.push(renderTemplate_(getMsgTemplate_('tournament_apply'), {
          events: appliedTournamentNames.map(function(name, idx) {
            const d = appliedTournamentEventDates[idx];
            const dateLine = d instanceof Date ? Utilities.formatDate(d, 'Asia/Tokyo', 'M月d日') + '開催\n' : '';
            return dateLine + '・' + name;
          }).join('\n\n'),
        }));
      }

      // オフラインイベント：当落通知あり
      if (appliedOfflineNames.length > 0) {
        const resultDate = buildOfflineResultDateStr_(appliedOfflineResultDates);
        msgParts.push(renderTemplate_(getMsgTemplate_('offline_apply'), {
          events: appliedOfflineNames.map(function(name, idx) {
            const d = appliedOfflineEventDates[idx];
            const r = appliedOfflineResultDates[idx];
            const dateLine   = d instanceof Date ? Utilities.formatDate(d, 'Asia/Tokyo', 'M月d日') + '開催\n' : '';
            const resultLine = r instanceof Date ? '\n（当落結果' + Utilities.formatDate(r, 'Asia/Tokyo', 'M月d日') + '頃お知らせ予定）' : '';
            return dateLine + '・' + name + resultLine;
          }).join('\n\n'),
          resultDate,
        }));
      }

      // オンライン相談：全員対応・相談詳細を確認
      if (appliedOnlineNames.length > 0) {
        const isVideo = (data.onlineConsultType || '') === 'video';
        const eventsStr = appliedOnlineNames.map(n => '・' + n).join('\n');
        if (isVideo) {
          const phoneConsult = data.onlinePhoneConsult || '';
          let phoneInfo = '\n電話相談：' + (phoneConsult || '未選択');
          if (phoneConsult === '希望する' && data.onlineConsultPhone) {
            phoneInfo += '（' + normalizePhone_(data.onlineConsultPhone) + '）';
          }
          msgParts.push(renderTemplate_(getMsgTemplate_('online_video_apply'), { events: eventsStr, phoneInfo }));
        } else {
          msgParts.push(renderTemplate_(getMsgTemplate_('online_text_apply'), { events: eventsStr }));
        }
      }

      const pushMsg = msgParts.join('\n\n');
      _pushResult = pushMessage(data.userId, pushMsg);
      Logger.log('PUSH userId=' + data.userId + ' result=' + JSON.stringify(_pushResult));

      // 動画相談の場合はLINEへの動画送信を依頼
      if ((data.onlineConsultType || '') === 'video') {
        pushMessage(data.userId, getMsgTemplate_('video_request'));
      }
      const nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MM/dd HH:mm');
      notifyStaff(`✅ LIFF応募 ${nowStr}\n${appliedParticipantNames.join('、')}（計${appliedParticipantNames.length}名）\n${appliedNames.join('、')}`);
    } else {
      _pushResult = pushMessage(data.userId, getMsgTemplate_(data.isNewRegistration ? 'registration_done' : 'profile_done'));
      Logger.log('PUSH(profile) userId=' + data.userId + ' result=' + JSON.stringify(_pushResult));
    }

    return { success: true, appliedEvents: appliedNames };

  } catch (err) {
    Logger.log('submitLiffApplication error: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

// GASエディタから直接実行してイベント応募完了メッセージをテスト送信する
// STAFF_USER_IDに送信し、ログでLINE APIの結果を確認できる
function testEventApplyMessage() {
  const userId = getProp('STAFF_USER_ID');
  if (!userId) {
    Logger.log('ERROR: STAFF_USER_IDが未設定です。スクリプトプロパティを確認してください。');
    return;
  }
  Logger.log('送信先 userId: ' + userId);

  const namesPart = 'テストユーザー 様';
  const template1 = getMsgTemplate_('header_apply');
  Logger.log('header_applyテンプレート: ' + template1);
  const msg1 = renderTemplate_(template1, { names: namesPart });
  Logger.log('header_apply描画結果: ' + msg1);

  const template2 = getMsgTemplate_('offline_apply');
  Logger.log('offline_applyテンプレート: ' + template2);
  const msg2 = renderTemplate_(template2, { events: '・テストイベント' });
  Logger.log('offline_apply描画結果: ' + msg2);

  const fullMsg = [msg1, msg2].join('\n\n');
  Logger.log('送信メッセージ全文:\n' + fullMsg);

  const result = pushMessage(userId, fullMsg);
  Logger.log('LINE API結果: ' + JSON.stringify(result));
}

function getLiffHtml(liffId, eventsJson) {
  return `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no">
<title>イベント参加登録</title>
<script charset="utf-8" src="https://static.line-scdn.net/liff/edge/2/sdk.js"><\/script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,'Hiragino Kaku Gothic ProN','Noto Sans JP',sans-serif;font-size:15px;background:#f2f2f7;color:#222;-webkit-tap-highlight-color:transparent}
.wrap{max-width:520px;margin:0 auto;padding:12px 12px 48px}
h1{font-size:17px;text-align:center;padding:14px 0 10px;color:#06c755;font-weight:bold}
.card{background:#fff;border-radius:10px;padding:16px;margin-bottom:10px}
.sec-label{font-size:13px;font-weight:bold;color:#555;border-bottom:1px solid #eee;padding-bottom:8px;margin-bottom:14px;letter-spacing:.3px}
.sec-note{font-size:12px;color:#888;font-weight:normal;margin-left:6px}
.fg{margin-bottom:14px}
.fg:last-child{margin-bottom:0}
.frow{display:flex;gap:8px}
.frow .fg{flex:1;min-width:0}
.fl{display:block;font-size:13px;color:#666;margin-bottom:4px}
.req{color:#e00;font-size:11px;margin-left:2px}
.sub-note{font-size:11px;color:#999;margin-top:3px}
input[type=text],input[type=email],input[type=tel],input[type=number]{width:100%;padding:9px 10px;border:1px solid #ddd;border-radius:7px;font-size:15px;background:#fff;-webkit-appearance:none;appearance:none}
input[type=number]{max-width:120px}
input:focus{border-color:#06c755;outline:none}
.rlist,.clist{display:flex;flex-direction:column;gap:6px}
.opt{display:flex;align-items:center;gap:9px;padding:9px 10px;border:1px solid #eee;border-radius:7px;cursor:pointer}
.opt input{width:18px;height:18px;flex-shrink:0;cursor:pointer;accent-color:#06c755}
.opt span{line-height:1.4;font-size:14px}
.ev-opt{display:flex;align-items:flex-start;gap:9px;padding:10px;border:1px solid #eee;border-radius:7px;cursor:pointer;margin-bottom:7px}
.ev-opt:last-child{margin-bottom:0}
.ev-opt input{width:18px;height:18px;flex-shrink:0;margin-top:2px;cursor:pointer;accent-color:#06c755}
.ev-name{font-weight:bold;font-size:14px;margin-bottom:2px}
.ev-date{font-size:12px;color:#888}
.btn{width:100%;padding:14px;background:#06c755;color:#fff;border:none;border-radius:9px;font-size:16px;font-weight:bold;cursor:pointer;margin-top:4px;letter-spacing:.5px}
.btn:disabled{background:#bbb;cursor:default}
.gerr{color:#e00;font-size:13px;text-align:center;padding:8px 10px;background:#fff0f0;border-radius:6px;margin-bottom:8px;display:none;line-height:1.6}
.ret-note{font-size:12px;color:#059a48;background:#f0fff6;border:1px solid #b8f0cc;border-radius:7px;padding:9px 11px;margin-bottom:10px}
.tos-box{font-size:12px;color:#555;background:#f9f9f9;border:1px solid #eee;border-radius:7px;padding:10px;max-height:120px;overflow-y:auto;margin-bottom:10px;line-height:1.7}
.loading{text-align:center;padding:70px 20px;color:#888}
.spin{width:36px;height:36px;border:3px solid #eee;border-top-color:#06c755;border-radius:50%;animation:spin .8s linear infinite;margin:0 auto 14px}
@keyframes spin{to{transform:rotate(360deg)}}
.success{text-align:center;padding:40px 20px}
.si{font-size:54px;margin-bottom:14px}
.st{font-size:20px;font-weight:bold;margin-bottom:10px}
.sb{color:#555;line-height:1.7;white-space:pre-line;font-size:14px}
.errscr{text-align:center;padding:50px 20px}
.ei{font-size:48px;margin-bottom:14px}
.em{color:#555;line-height:1.6;white-space:pre-line;font-size:14px}
</style>
</head>
<body>
<div class="wrap">
  <h1>🎾 EPSOM&amp;CO テニスイベント参加登録</h1>
  <div id="loading" class="loading"><div class="spin"></div>読み込み中...</div>
  <div id="fw" style="display:none"></div>
  <div id="success" class="success" style="display:none">
    <div class="si">✅</div>
    <div class="st">応募を受け付けました！</div>
    <div class="sb" id="sbody"></div>
  </div>
  <div id="errscr" class="errscr" style="display:none">
    <div class="ei">⚠️</div>
    <div class="em" id="emsg"></div>
  </div>
</div>
<script>
var LID = '${liffId}';
var eventsData = ${eventsJson};
var userId = null, memberData = null;

window.addEventListener('load', function() {
  var currentUrl = window.location.href;
  document.getElementById('loading').innerHTML = '<div class="spin"></div>ステップ1: LIFF初期化中...<br><span style="font-size:10px;color:#aaa;word-break:break-all">' + currentUrl + '</span>';

  var timer = setTimeout(function() {
    document.getElementById('loading').innerHTML = '<div style="color:#e00;font-size:13px;padding:20px">⚠️ タイムアウトしました<br><br>URL: ' + currentUrl + '<br><br>isInClient: ' + (typeof liff !== "undefined" ? liff.isInClient() : "liff未定義") + '</div>';
  }, 15000);

  liff.init({ liffId: LID, withLoginOnExternalBrowser: true })
    .then(function() {
      document.getElementById('loading').innerHTML = '<div class="spin"></div>ステップ2: プロフィール取得中...';
      return liff.getProfile();
    })
    .then(function(profile) {
      clearTimeout(timer);
      if (!profile) { showErr('プロフィールの取得に失敗しました。'); return; }
      userId = profile.userId;
      renderForm();
      google.script.run
        .withSuccessHandler(function(m) { if (m) { memberData = m; prefillForm(m); } })
        .withFailureHandler(function() {})
        .getMemberData(userId);
    })
    .catch(function(e) {
      clearTimeout(timer);
      showErr('エラーが発生しました。\\n' + (e && e.message ? e.message : JSON.stringify(e)));
    });
});

function chk(nm) { return Array.from(document.querySelectorAll('input[name="' + nm + '"]:checked')).map(function(el){return el.value;}); }
function rv(nm)  { var el = document.querySelector('input[name="' + nm + '"]:checked'); return el ? el.value : ''; }
function val(id) { return document.getElementById(id).value.trim(); }
function ro(nm, arr) {
  return arr.map(function(o) {
    return '<label class="opt"><input type="radio" name="' + nm + '" value="' + o + '"><span>' + o + '</span></label>';
  }).join('');
}
function roP(nm, arr, sel) {
  return arr.map(function(o) {
    return '<label class="opt"><input type="radio" name="' + nm + '" value="' + o + '"' + (sel === o ? ' checked' : '') + '><span>' + o + '</span></label>';
  }).join('');
}
function co(nm, arr, sels) {
  return arr.map(function(o) {
    return '<label class="opt"><input type="checkbox" name="' + nm + '" value="' + o + '"' + (sels.indexOf(o) !== -1 ? ' checked' : '') + '><span>' + o + '</span></label>';
  }).join('');
}
function h(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

function renderForm(m) {
  document.getElementById('loading').style.display = 'none';
  var fw = document.getElementById('fw');

  var evHtml = eventsData.length === 0
    ? '<p style="color:#888;font-size:14px;text-align:center;padding:8px 0">現在、募集中のイベントはありません。</p>'
    : eventsData.map(function(ev, i) {
        return '<label class="ev-opt"><input type="checkbox" name="ev" value="' + i + '"><div><div class="ev-name">' + h(ev.name) + '</div><div class="ev-date">開催：' + ev.eventDate + '　締切：' + ev.closingDate + '</div></div></label>';
      }).join('');
  var cOpts = ro('coach', ['知っていて、レッスンも受けたことがある','知っているがレッスンを受けたことはない','前回のイベントで知った','今回のイベントで初めて知った','その他']);
  var sOpts = co('src',  ['公式SNS','コーチSNS','LINEでのお知らせ','知人・友人の紹介','Google検索','その他'], []);
  var rOpts = co('rsn',  ['コーチのレッスンを受けてみたかった','無料イベントだった','テニス仲間に誘われた','テニスイベントに興味があった','その他'], []);

  fw.innerHTML =
    '<div id="ret-note" style="display:none" class="ret-note">以前の登録情報が入力されています。変更があれば修正してから応募してください。</div>' +

    '<div class="card">' +
      '<div class="sec-label">👤 プロフィール<span class="sec-note">初回のみ入力・以降は自動入力</span></div>' +
      '<div class="frow"><div class="fg"><label class="fl">苗字<span class="req">*</span></label><input type="text" id="fN" placeholder="山田"></div>' +
      '<div class="fg"><label class="fl">名前<span class="req">*</span></label><input type="text" id="gN" placeholder="太郎"></div></div>' +
      '<div class="frow"><div class="fg"><label class="fl">フリガナ（苗字）<span class="req">*</span></label><input type="text" id="fK" placeholder="ヤマダ"></div>' +
      '<div class="fg"><label class="fl">フリガナ（名前）<span class="req">*</span></label><input type="text" id="gK" placeholder="タロウ"></div></div>' +
      '<div class="fg"><label class="fl">年齢<span class="req">*</span></label><input type="number" id="age" placeholder="30" min="10" max="99"></div>' +
      '<div class="fg"><label class="fl">性別<span class="req">*</span></label><div class="rlist" id="gndList">' + roP('gnd', ['男性','女性','その他'], '') + '</div></div>' +
      '<div class="fg"><label class="fl">メールアドレス<span class="req">*</span></label><input type="email" id="email" placeholder="example@email.com"></div>' +
      '<div class="fg"><label class="fl">電話番号<span class="req">*</span></label><input type="tel" id="phone" placeholder="09012345678"></div>' +
      '<div class="fg"><label class="fl">緊急連絡先</label>' +
        '<input type="tel" id="emergency" placeholder="09012345678">' +
        '<div class="sub-note">※未成年でご参加の方は必ずご記入ください</div></div>' +
      '<div class="fg"><label class="fl">テニスレベル<span class="req">*</span></label><div class="rlist" id="lvList">' + roP('lv', ['初心者','初級','初中級','中級','中上級','上級','超上級','プロ'], '') + '</div></div>' +
      '<div class="fg"><label class="fl">テニスをプレイする頻度<span class="req">*</span></label><div class="rlist" id="freqList">' + roP('freq', ['ほぼ毎日','週2〜3回（平日が多い）','週2〜3回（土日祝が多い）','週1回（平日が多い）','週1回（土日祝が多い）','月2〜3回','月1回','年2〜3回','年1回','その他'], '') + '</div></div>' +
      '<div class="fg"><label class="fl">テニス歴<span class="req">*</span></label><div class="rlist" id="histList">' + roP('hist', ['1年未満','1〜3年','3〜5年','5年以上','その他'], '') + '</div></div>' +
      '<div class="fg"><label class="fl">普段テニスをしている地域<span class="req">*</span><span class="sec-note">複数選択可</span></label><div class="clist" id="areaList">' + co('area', ['東京','神奈川','千葉','埼玉','その他関東圏','関東圏外','その他'], []) + '</div></div>' +
      '<div class="fg"><label class="fl">テニスをプレイしている環境<span class="req">*</span><span class="sec-note">複数選択可</span></label><div class="clist" id="envList">' + co('env', ['スクールに通っている','会社のサークル','学校の部活・サークル','テニス仲間とコートを借りる','その他'], []) + '</div></div>' +
    '</div>' +

    '<div class="card">' +
      '<div class="sec-label">📅 今回の応募内容<span class="sec-note">毎回入力</span></div>' +
      '<div class="fg"><label class="fl">応募するイベント<span class="sec-note">複数選択可・任意</span></label>' + evHtml + '</div>' +
      '<div class="fg"><label class="fl">コーチのレッスンについて<span class="req">*</span></label><div class="rlist">' + cOpts + '</div></div>' +
      '<div class="fg"><label class="fl">イベントを知ったきっかけ<span class="req">*</span><span class="sec-note">複数選択可</span></label><div class="clist">' + sOpts + '</div></div>' +
      '<div class="fg"><label class="fl">参加応募のきっかけ<span class="req">*</span><span class="sec-note">複数選択可</span></label><div class="clist">' + rOpts + '</div></div>' +
    '</div>' +

    '<div class="card">' +
      '<div class="sec-label">📋 利用規約・プライバシーポリシー<span class="sec-note">毎回確認</span></div>' +
      '<div class="tos-box">Epsom&amp;Co. TENNIS SUPPORT イベント応募 利用規約<br><br>' +
      '本イベントに応募された時点で、本規約およびプライバシーポリシーに同意したものとみなします。応募者多数の場合は抽選を行い、当選・落選の双方にLINEで結果をご連絡します。未成年の方は保護者の同意を得たうえで応募してください。当選後にキャンセルが発生した場合は速やかにご連絡ください。イベント当日は写真・動画の撮影を行う場合があり、公式SNS等で使用される場合があります。取得した個人情報はイベント運営の目的のみに使用し、本人の同意なく第三者へ提供しません。</div>' +
      '<label class="opt" style="border:none;padding:0;align-items:flex-start">' +
        '<input type="checkbox" id="tos" style="margin-top:2px;accent-color:#06c755"><span style="font-size:14px">利用規約・プライバシーポリシーに同意する<span class="req">*</span></span>' +
      '</label>' +
    '</div>' +

    '<div id="gerr" class="gerr"></div>' +
    '<button class="btn" id="sbtn" onclick="doSubmit()">送信する</button>';

  fw.style.display = '';
}

// 会員データが届いたらフォームに埋める
function prefillForm(m) {
  if (!m) return;
  document.getElementById('ret-note').style.display = '';
  var np = (m.name || '').split(' '), fp = (m.furigana || '').split(' ');
  document.getElementById('fN').value = np[0] || '';
  document.getElementById('gN').value = np.slice(1).join(' ') || '';
  document.getElementById('fK').value = fp[0] || '';
  document.getElementById('gK').value = fp.slice(1).join(' ') || '';
  document.getElementById('age').value = m.age || '';
  document.getElementById('email').value = m.email || '';
  document.getElementById('phone').value = m.phone || '';
  document.getElementById('emergency').value = m.emergency || '';
  if (m.gender) { var el = document.querySelector('input[name="gnd"][value="' + m.gender + '"]'); if (el) el.checked = true; }
  if (m.tennisLevel) { var el = document.querySelector('input[name="lv"][value="' + m.tennisLevel + '"]'); if (el) el.checked = true; }
  if (m.tennisFreq) { var el = document.querySelector('input[name="freq"][value="' + m.tennisFreq + '"]'); if (el) el.checked = true; }
  if (m.tennisHistory) { var el = document.querySelector('input[name="hist"][value="' + m.tennisHistory + '"]'); if (el) el.checked = true; }
  (m.tennisArea || '').split('・').filter(Boolean).forEach(function(v) { var el = document.querySelector('input[name="area"][value="' + v + '"]'); if (el) el.checked = true; });
  (m.tennisEnv  || '').split('・').filter(Boolean).forEach(function(v) { var el = document.querySelector('input[name="env"][value="' + v + '"]');  if (el) el.checked = true; });
}

function doSubmit() {
  var btn = document.getElementById('sbtn'), gerr = document.getElementById('gerr');
  gerr.style.display = 'none';
  var fN = val('fN'), gN = val('gN'), fK = val('fK'), gK = val('gK');
  var age = val('age'), gnd = rv('gnd'), lv = rv('lv');
  var email = val('email'), phone = val('phone'), emergency = val('emergency');
  var freq = rv('freq'), hist = rv('hist');
  var areas = chk('area'), envs = chk('env');
  var evIdxs = chk('ev'), coach = rv('coach'), srcs = chk('src'), rsns = chk('rsn');
  var tos = document.getElementById('tos').checked;

  var errs = [];
  if (!fN || !gN)     errs.push('お名前を入力してください。');
  if (!fK || !gK)     errs.push('フリガナを入力してください。');
  if (!age)           errs.push('年齢を入力してください。');
  if (!gnd)           errs.push('性別を選択してください。');
  if (!email)         errs.push('メールアドレスを入力してください。');
  if (!phone)         errs.push('電話番号を入力してください。');
  if (!lv)            errs.push('テニスレベルを選択してください。');
  if (!freq)          errs.push('テニスをプレイする頻度を選択してください。');
  if (!hist)          errs.push('テニス歴を選択してください。');
  if (areas.length === 0) errs.push('テニスをしている地域を選択してください。');
  if (envs.length  === 0) errs.push('テニスをプレイしている環境を選択してください。');
  if (evIdxs.length > 0) {
    if (!coach)            errs.push('コーチのレッスンについて回答してください。');
    if (srcs.length === 0) errs.push('イベントを知ったきっかけを選択してください。');
    if (rsns.length === 0) errs.push('参加応募のきっかけを選択してください。');
  }
  if (!tos)           errs.push('利用規約への同意が必要です。');

  if (errs.length) {
    gerr.innerHTML = errs.join('<br>');
    gerr.style.display = 'block';
    gerr.scrollIntoView({ behavior: 'smooth', block: 'center' });
    return;
  }

  btn.disabled = true; btn.textContent = '送信中...';
  var selEvs = evIdxs.map(function(i) { return eventsData[parseInt(i)]; });

  google.script.run
    .withSuccessHandler(function(res) {
      if (res.success) {
        document.getElementById('fw').style.display = 'none';
        var applied = res.appliedEvents || [];
        document.querySelector('.st').textContent = applied.length > 0 ? '応募を受け付けました！' : 'プロフィールを更新しました！';
        document.getElementById('sbody').textContent = applied.length > 0
          ? '応募イベント:\\n' + applied.map(function(n){return '・' + n;}).join('\\n') + '\\n\\n当落結果はLINEでお知らせします。\\nしばらくお待ちください。'
          : 'プロフィール情報を更新しました。\\nありがとうございます！';
        document.getElementById('success').style.display = '';
        setTimeout(function() { try { if (liff.isInClient()) liff.closeWindow(); } catch(e){} }, 3000);
      } else {
        btn.disabled = false; btn.textContent = '送信する';
        gerr.textContent = res.error || '送信に失敗しました。再度お試しください。';
        gerr.style.display = 'block';
      }
    })
    .withFailureHandler(function() {
      btn.disabled = false; btn.textContent = '送信する';
      gerr.textContent = '通信エラーが発生しました。再度お試しください。';
      gerr.style.display = 'block';
    })
    .submitLiffApplication({
      userId: userId,
      familyName: fN, givenName: gN,
      familyNameKana: fK, givenNameKana: gK,
      age: age, gender: gnd, tennisLevel: lv,
      email: email, phone: phone, emergency: emergency,
      tennisFreq: freq, tennisHistory: hist,
      tennisArea: areas, tennisEnv: envs,
      selectedEvents: selEvs,
      coachKnowledge: coach, eventSource: srcs, applyReason: rsns,
    });
}

function showErr(msg) {
  document.getElementById('loading').style.display = 'none';
  document.getElementById('emsg').textContent = msg;
  document.getElementById('errscr').style.display = '';
}
<\/script>
</body>
</html>`;
}

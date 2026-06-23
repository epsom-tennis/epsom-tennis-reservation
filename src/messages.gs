// LIFFから自動送信されるLINEメッセージのテンプレート管理
// 「メッセージ設定」シートに保存し、ダッシュボードから編集できるようにする。
// シートに値が無い/シート自体が無い場合はdefaultにフォールバックするため、未セットアップでも動作する。

const MESSAGE_SHEET_NAME = 'メッセージ設定';

const MESSAGE_TEMPLATE_DEFS = [
  {
    key: 'header_apply',
    label: '応募受付ヘッダー（オフライン・オンライン共通の先頭文）',
    vars: ['names'],
    default: '{names}ご応募を受け付けました！',
  },
  {
    key: 'offline_apply',
    label: 'オフラインイベント 応募受付',
    vars: ['events'],
    default: '【オフラインイベント】\n{events}\n\n当落結果は後日このLINEでお知らせします。\nしばらくお待ちください。',
  },
  {
    key: 'online_text_apply',
    label: 'オンライン相談（文章）受付',
    vars: ['events'],
    default: '【オンライン相談】\n{events}\n\nご相談方法：文章\n\nご相談内容を受け付けました。\n配信にてお答えしますので、お楽しみに！',
  },
  {
    key: 'online_video_apply',
    label: 'オンライン相談（動画）受付',
    vars: ['events', 'phoneInfo'],
    default: '【オンライン相談】\n{events}\n\nご相談方法：動画{phoneInfo}\n\nご相談内容を受け付けました。\n配信にてお答えしますので、お楽しみに！',
  },
  {
    key: 'video_request',
    label: '動画送信のお願い（動画相談応募の直後に追加送信）',
    vars: [],
    default: '動画をこのLINEに直接送ってください 🎥\n受け取り次第、コーチが確認します。',
  },
  {
    key: 'video_received',
    label: '動画受信確認（動画を送ってもらった直後の自動返信）',
    vars: [],
    default: '動画を受け取りました！ありがとうございます 🎾\nコーチが確認次第、配信でお答えします。',
  },
  {
    key: 'registration_done',
    label: '初回プロフィール登録 完了',
    vars: [],
    default: 'プロフィール情報を更新しました。ありがとうございます！',
  },
  {
    key: 'profile_done',
    label: 'プロフィール更新 完了（2回目以降の修正）',
    vars: [],
    default: 'プロフィール情報を更新しました。ありがとうございます！',
  },
  {
    key: 'participation_confirmed',
    label: '参加確定（当選者が「参加します」を押した直後の返信）',
    vars: ['eventName'],
    default: '【参加確定】\n「{eventName}」へのご参加を確認しました！\n当日お会いできることを楽しみにしております 🎾',
  },
  {
    key: 'participation_canceled',
    label: 'キャンセル受付（当選者が「キャンセルします」を押した直後の返信）',
    vars: ['eventName'],
    default: '【キャンセル受付】\n「{eventName}」へのご参加をキャンセルしました。\nご連絡いただきありがとうございます。またのご参加をお待ちしております。',
  },
];

// 「メッセージ設定」シートを取得（無ければ作成し、未登録キーのデフォルト行を補充）
function ensureMessageTemplatesSheet_() {
  const ss = SpreadsheetApp.openById(getProp('SPREADSHEET_ID'));
  let sheet = ss.getSheetByName(MESSAGE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MESSAGE_SHEET_NAME);
    sheet.appendRow(['キー', '表示名', '本文', '使える変数（メモ）']);
    sheet.setFrozenRows(1);
  }
  const data = sheet.getDataRange().getValues();
  const existingKeys = new Set(data.slice(1).map(r => String(r[0])));
  MESSAGE_TEMPLATE_DEFS.forEach(def => {
    if (!existingKeys.has(def.key)) {
      sheet.appendRow([def.key, def.label, def.default, def.vars.map(v => '{' + v + '}').join(' ')]);
    }
  });
  return sheet;
}

// 指定キーのテンプレート本文を返す（シートに値が無ければdefaultを返す）
function getMsgTemplate_(key) {
  const def = MESSAGE_TEMPLATE_DEFS.find(d => d.key === key);
  try {
    const sheet = ensureMessageTemplatesSheet_();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === key) {
        const val = String(data[i][2] || '').trim();
        return val || (def ? def.default : '');
      }
    }
  } catch (err) {
    Logger.log('getMsgTemplate_ error: ' + err.toString());
  }
  return def ? def.default : '';
}

// {変数名} を実際の値に置き換える
function renderTemplate_(text, vars) {
  let result = text || '';
  Object.keys(vars || {}).forEach(k => {
    result = result.split('{' + k + '}').join(vars[k] != null ? vars[k] : '');
  });
  return result;
}

// ===== ダッシュボード用 =====

// 全テンプレートの現在値 + イベント別の当選/落選文を返す
function getMessageTemplates() {
  ensureMessageTemplatesSheet_();
  const templates = MESSAGE_TEMPLATE_DEFS.map(def => ({
    key: def.key,
    label: def.label,
    vars: def.vars,
    value: getMsgTemplate_(def.key),
  }));

  const events = getAllEvents().map(ev => {
    const msgs = getResultMessages(ev.resultSheetName);
    return { resultSheetName: ev.resultSheetName, name: ev.name, winMsg: msgs.win, loseMsg: msgs.lose };
  });

  return { templates, events };
}

// 共通テンプレートを保存する（values: {key: 本文, ...}）
function saveMessageTemplates(values) {
  try {
    const sheet = ensureMessageTemplatesSheet_();
    const data = sheet.getDataRange().getValues();
    Object.keys(values || {}).forEach(key => {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === key) {
          sheet.getRange(i + 1, 3).setValue(values[key]);
          break;
        }
      }
    });
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// イベント別の当選/落選メッセージを設定シートE/F列に保存する
function saveEventResultMessage(resultSheetName, type, text) {
  try {
    const appSheetName = resultSheetName.replace('_当落', '_応募');
    const configSheet = getSheet(SHEET.CONFIG);
    if (!configSheet) return { success: false, error: '設定シートが見つかりません。' };
    const data = configSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][3]).trim() === appSheetName) {
        const col = type === 'win' ? 5 : 6; // E列=5, F列=6（1-indexed）
        configSheet.getRange(i + 1, col).setValue(text);
        return { success: true };
      }
    }
    return { success: false, error: 'イベントが見つかりません。' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

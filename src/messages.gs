// LIFFから自動送信されるLINEメッセージのテンプレート管理
// 「メッセージ設定」シートに保存し、ダッシュボードから編集できるようにする。
// シートに値が無い/シート自体が無い場合はdefaultにフォールバックするため、未セットアップでも動作する。

const MESSAGE_SHEET_NAME = 'メッセージ設定';

const MESSAGE_TEMPLATE_DEFS = [
  {
    key: 'win_offline',
    label: '【オフライン】当選メッセージ（基本形）',
    vars: ['eventDate', 'coachName', 'venue', 'meetingTime', 'lessonTime', 'courtType', 'items', 'fee', 'lockerInfo', 'facilityUrl', 'confirmDeadline'],
    default:
      'この度はテニスイベントにご応募いただきありがとうございました！\n\n' +
      '抽選の結果、ご当選となりましたのでご連絡いたします🎾\n\n' +
      '当日は下記内容をご確認のうえ、ご参加をお願いいたします。\n\n' +
      '【開催日程】\n{eventDate}\n\n' +
      '【コーチ】\n{coachName}\n\n' +
      '【開催場所】\n{venue}\n\n' +
      '{{#if meetingTime}}【集合時間】\n{meetingTime}\n※お時間になりましたら、直接コート周辺までお越しください。\n\n{{/if}}' +
      '【レッスン時間】\n{lessonTime}\n\n' +
      '{{#if courtType}}【コートについて】\n{courtType}\n\n{{/if}}' +
      '{{#if items}}【持ち物】\n{items}\n\n※レンタル用品のご用意はございません。\n※ボールは事務局にてご用意いたします。\n\n{{/if}}' +
      '{{#if fee}}【参加費】\n{fee}\n\n{{/if}}' +
      '{{#if lockerInfo}}【更衣室について】\n{lockerInfo}\n\n{{/if}}' +
      '{{#if facilityUrl}}【施設について】\n施設に関する詳細は、下記公式HPをご確認ください。\n{facilityUrl}\n\n※施設ルールに沿ってご利用をお願いいたします。\n\n{{/if}}' +
      '【保険・怪我について】\n練習中の事故・怪我につきましては、事務局では責任を負いかねます。\n必要に応じて、スポーツ障害保険等へのご加入をお願いいたします。\n\n' +
      '【雨天時について】\nインドアコートでの開催となるため、雨天時も原則開催予定となっております。\nただし、荒天や交通状況等により開催可否の判断が必要となった場合は、開始1時間前までにLINEにてご連絡いたします。\nご連絡がない場合は、予定通り開催となります。\n\n' +
      '{{#if confirmDeadline}}【参加確認のお願い】\n確実にご参加いただける方を優先させていただくため、\n【{confirmDeadline}まで】に「参加します」とご返信をお願いいたします。\n\n期限までにご返信がない場合は、キャンセル扱いとさせていただく場合がございます。\n\n{{/if}}' +
      'その他ご質問等ございましたら、こちらのLINEよりお気軽にご連絡ください。\n当日お会いできるのを楽しみにしております🎾',
  },
  {
    key: 'lose_offline',
    label: '【オフライン】落選メッセージ（基本形）',
    vars: ['eventDate', 'eventName'],
    default:
      'この度はテニスイベントにご応募いただきありがとうございました！\n\n' +
      '抽選の結果、{eventDate}開催「{eventName}」につきましては、今回残念ながら落選となりました。\n\n' +
      '定員に達したため、ご参加いただくことができませんでした。\nまたのご応募を心よりお待ちしております。\n\n' +
      'その他ご質問等ございましたら、こちらのLINEよりお気軽にご連絡ください。',
  },
  {
    key: 'win_online',
    label: '【オンライン】当選メッセージ（基本形）',
    vars: ['eventDate'],
    default:
      'この度はオンライン相談にご応募いただきありがとうございました！\n\n' +
      '抽選の結果、ご当選となりましたのでご連絡いたします🎾\n\n' +
      '【配信予定日】\n{eventDate}\n\n' +
      'ご相談内容には配信にて回答させていただきます。\n当日をお楽しみにお待ちください。\n\n' +
      'その他ご質問等ございましたら、こちらのLINEよりお気軽にご連絡ください。',
  },
  {
    key: 'lose_online',
    label: '【オンライン】落選メッセージ（基本形）',
    vars: ['eventName'],
    default:
      'この度はオンライン相談にご応募いただきありがとうございました！\n\n' +
      '抽選の結果、「{eventName}」につきましては、今回残念ながら落選となりました。\n\n' +
      'またのご応募を心よりお待ちしております。\n\n' +
      'その他ご質問等ございましたら、こちらのLINEよりお気軽にご連絡ください。',
  },
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
    default: '動画をこのLINEに直接送ってください 🎥\nファイルサイズが大きい場合は、ギガファイル便などのアップロードサービスのURLを送っていただいても構いません。\n受け取り次第、コーチが確認します。',
  },
  {
    key: 'video_received',
    label: '動画受信確認（動画を送ってもらった直後の自動返信）',
    vars: [],
    default: '動画の送信ありがとうございます！🎾\n担当者が確認のうえ、ご連絡いたします。',
  },
  {
    key: 'video_url_received',
    label: 'URL受信確認（ギガファイル便等のURLを送ってもらった直後の自動返信。動画と無関係なURLが届く可能性もあるため汎用的な文言にしている）',
    vars: [],
    default: 'URLの送付を確認しました！ありがとうございます 🎾\n担当者が確認のうえ、ご連絡いたします。',
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
    key: 'participation_expired',
    label: '参加確認期限切れ（期限を過ぎて「参加します」が押された場合の返信）',
    vars: ['eventName'],
    default: '【期限切れ】\n「{eventName}」の参加確認期限を過ぎているため、キャンセル扱いとなりました。\nご連絡が必要な場合は、こちらのLINEよりお問い合わせください。',
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
        // 誤って<br>タグが入力されていても改行として機能するように変換する（LINEはプレーンテキストのため<br>は表示されない）
        const val = String(data[i][2] || '').trim().replace(/<br\s*\/?>/gi, '\n');
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

// {{#if 変数名}}〜{{/if}} で囲んだ範囲を、その変数が空ならまるごと取り除く（見出しごと非表示にするため）
function renderConditionalBlocks_(text, vars) {
  return (text || '').replace(/\{\{#if (\w+)\}\}([\s\S]*?)\{\{\/if\}\}/g, (_, key, inner) => {
    return (vars && vars[key]) ? inner : '';
  });
}

// URLをLINEがリンク化・プレビュー表示しないように、スキーム直後にゼロ幅スペースを挟んで見た目を保ったまま無効化する
function breakUrlPreview_(url) {
  if (!url) return url;
  return url.replace(/^(https?:\/\/)/i, '$1' + '​');
}

// 当落メッセージ用の日付フォーマット（例: 5月30日（土））
function formatDateJp_(date) {
  if (!date) return '';
  const days = ['日', '月', '火', '水', '木', '金', '土'];
  const dayIdx = parseInt(Utilities.formatDate(date, 'Asia/Tokyo', 'u')) % 7;
  return Utilities.formatDate(date, 'Asia/Tokyo', 'M月d日') + '（' + days[dayIdx] + '）';
}

// イベント情報を当落メッセージの基本形テンプレートに差し込んで本文を生成する
// type: 'win' または 'lose'
function buildResultMessage_(ev, type) {
  const isOnline = (ev.eventType || 'オフライン') === 'オンライン';
  const tmpl = getMsgTemplate_(type + '_' + (isOnline ? 'online' : 'offline'));
  const vars = {
    eventName:       ev.name || '',
    eventDate:       formatDateJp_(ev.eventDate),
    coachName:       ev.coachName || '',
    venue:           ev.venue || '',
    meetingTime:     ev.meetingTime || '',
    lessonTime:      ev.eventTime || '',
    courtType:       ev.courtType || '',
    items:           ev.items || '',
    fee:             ev.fee || '',
    lockerInfo:      ev.lockerInfo || '',
    facilityUrl:     breakUrlPreview_(ev.facilityUrl || ''),
    confirmDeadline: ev.confirmDeadline || '',
  };
  return renderTemplate_(renderConditionalBlocks_(tmpl, vars), vars);
}

// ===== ダッシュボード用 =====

// 全テンプレートの現在値を返す
function getMessageTemplates() {
  ensureMessageTemplatesSheet_();
  const templates = MESSAGE_TEMPLATE_DEFS.map(def => ({
    key: def.key,
    label: def.label,
    vars: def.vars,
    value: getMsgTemplate_(def.key),
  }));

  return { templates };
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

// 文章管理タブの「試し送信」用：{xxx}をサンプル値で埋め、STAFF_USER_ID宛にテスト送信する
function testSendMessageTemplate(text) {
  try {
    const staffUserId = getProp('STAFF_USER_ID');
    if (!staffUserId) return { success: false, error: 'STAFF_USER_IDが未設定です。スクリプトプロパティを確認してください。' };

    const sampleVars = {
      eventDate:       '6月15日（日）',
      eventName:       'サンプルイベント',
      coachName:       '山田コーチ',
      venue:           '渋谷テニスコート',
      meetingTime:     '18:50',
      lessonTime:      '19:00〜21:00',
      courtType:       'カーペットコートになります。',
      items:           '・ラケット\n・テニスウェア\n・テニスシューズ',
      fee:             '無料でご参加いただけます！',
      lockerInfo:      '受付でのお声がけは不要です。',
      facilityUrl:     breakUrlPreview_('https://example.com'),
      confirmDeadline: '6月10日（火）12:00',
      names:           '山田 様',
      events:          '・サンプルイベントA\n・サンプルイベントB',
      phoneInfo:       '（090-1234-5678）',
    };
    const rendered = renderTemplate_(renderConditionalBlocks_(text, sampleVars), sampleVars);
    pushMessage(staffUserId, '【テスト送信（サンプル値で差し込み）】\n\n' + rendered);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

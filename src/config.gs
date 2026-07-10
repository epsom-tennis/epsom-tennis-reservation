// 固定シート名（スプレッドシートのタブ名と完全一致）
// 応募シート・当落シートはイベントごとに {識別名}_応募 / {識別名}_当落 の命名規則で動的に参照する
const SHEET = {
  CONFIG: '設定シート',
  MEMBERS: '会員マスタ',
  ACTION_LOG: 'アクション履歴',
  TERMS: '規約管理',
  REFERRAL_CODES: '紹介コード',
};

// スクリプトプロパティを取得するヘルパー
function getProp(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

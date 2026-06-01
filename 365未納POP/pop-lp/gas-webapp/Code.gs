/**
 * FIT365 多言語POP — ウェブアプリとして公開
 * デプロイ: デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 * 次のユーザーとして実行: 自分
 * アクセスできるユーザー: 全員（または組織内）
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('FIT365 月会費未納のお知らせ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

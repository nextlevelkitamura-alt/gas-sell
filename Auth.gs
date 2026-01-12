// 【追記】認証門番システム（Webアプリ用）
// ==========================================

/**
 * 外部からの認証リクエストを受け取る（▶実行ボタンは押しちゃダメ！）
 * デプロイ手順:
 * 1. 右上の「デプロイ」 > 「新しいデプロイ」
 * 2. 種類の選択: 「ウェブアプリ」
 * 3. 説明: "Auth" など適当に
 * 4. 次のユーザーとして実行: 「自分 (your_email@...)」
 * 5. アクセスできるユーザー: 「全員」
 * 6. 「デプロイ」ボタンを押して、発行されたURL (Web App URL) をコピーする
 */
function doGet(e) {
  // パラメータ取得
  const userEmail = e.parameter.email;
  
  if (!userEmail) {
    return ContentService.createTextOutput("Error: No email provided");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("認証リスト"); // シート名が画像と一致していること
  
  if (!sheet) {
    return ContentService.createTextOutput("Error: '認証リスト' sheet not found");
  }
  
  // 1行目（見出し）を飛ばし、2行目以降の「C列（3列目）」をすべて取得
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return ContentService.createTextOutput("NG");
  
  // getRange(row, col, numRows, numCols)
  // C列は3番目
  const authEmails = sheet.getRange(2, 3, lastRow - 1, 1).getValues().flat();

  // 実際にツールを使っている人のアドレスが含まれているか判定
  const isAllowed = authEmails.includes(userEmail);
  
  // ログに残す（デバッグ用）
  console.log(`Auth Request: ${userEmail} -> ${isAllowed ? "OK" : "NG"}`);
  
  return ContentService.createTextOutput(isAllowed ? "OK" : "NG");
}

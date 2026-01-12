/**
 * ■ セットアップツールの使い方
 * 
 * 1. まず、チェックボックスなどが正しい状態の「見本となるシート」を開きます。
 * 2. 上部メニューの「⭐ Setup」から「開発者用: 設定をスキャンしてログ表示」を実行します。
 * 3. 実行ログ（表示→ログ）を確認し、出力された [... ] から始まるコード全体をコピーします。
 * 4. このファイルの下記 `const SETUP_CONFIG = []` の中身（[]の部分）を、コピーした内容で書き換えて保存します。
 *    例: const SETUP_CONFIG = [{"sheetName": "...", ...}, ...];
 * 
 * 5. これで準備完了です。
 *    おかしくなったシートを開き、「⭐ Setup」から「シートを修復/セットアップ」を実行すると、
 *    プルダウンの形やチェックボックスが元通りになります。
 */

// -------------------------------------------------------------
// ここにスキャンした設定を貼り付けてください
// -------------------------------------------------------------
const SETUP_CONFIG = []; 
// ↑ スキャン結果をここにペーストします。
// 例: const SETUP_CONFIG = [ { "sheetName": "...", "ranges": [ ... ] } ];


// -------------------------------------------------------------
// メニュー追加
// -------------------------------------------------------------
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⭐ Setup')
    .addItem('シートを修復/セットアップ', 'runSetup')
    .addSeparator()
    .addItem('開発者用: 設定をスキャンしてログ表示', 'getValidationConfig')
    .addToUi();
}

/**
 * 1. 【修復実行】
 * SETUP_CONFIGの設定に基づいて、現在のスプレッドシートの入力規則（プルダウン・チェックボックス）を復元します。
 * Checkboxが "FALSE" などの文字列になってしまっている場合も修正します。
 */
function runSetup() {
  if (!SETUP_CONFIG || SETUP_CONFIG.length === 0) {
    SpreadsheetApp.getUi().alert('設定(SETUP_CONFIG)が空です。\nまずは「設定をスキャン」をして、コードに貼り付けてください。');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let fixedCount = 0;

  SETUP_CONFIG.forEach(sheetData => {
    const sheet = ss.getSheetByName(sheetData.sheetName);
    if (!sheet) return;

    sheetData.ranges.forEach(rule => {
      // rule: { row, col, numRows, numCols, type: 'CHECKBOX' | 'LIST', values: [...] }
      const range = sheet.getRange(rule.row, rule.col, rule.numRows, rule.numCols);
      
      // 入力規則の再設定
      if (rule.type === 'CHECKBOX') {
        // チェックボックスの作成
        const ruleBuilder = SpreadsheetApp.newDataValidation().requireCheckbox();
        range.setDataValidation(ruleBuilder.build());
        
        // 既存の値が "FALSE" / "TRUE" という「文字列」になっている場合、ブール値に直す
        // これをやらないとチェックボックスが機能しないことがある
        const values = range.getValues();
        let valChanged = false;
        const newValues = values.map(row => row.map(cell => {
          if (cell === 'FALSE') return false;
          if (cell === 'TRUE') return true;
          return cell;
        }));
        
        // 値が変わった箇所があれば更新
        if (JSON.stringify(values) !== JSON.stringify(newValues)) {
          range.setValues(newValues);
        }
        
      } else if (rule.type === 'LIST') {
        // プルダウン(リスト)の作成
        // リストの選択肢が存在する場合のみ
        if (rule.values && rule.values.length > 0) {
          const ruleBuilder = SpreadsheetApp.newDataValidation()
            .requireValueInList(rule.values)
            .setAllowInvalid(true) // 無効な入力を許可するかはお好みで（通常は警告のみにする等）
            .build();
          range.setDataValidation(ruleBuilder.build());
        }
      }
      
      fixedCount++;
    });
  });

  SpreadsheetApp.getUi().alert(`セットアップ完了\n${fixedCount} 箇所の設定を適用しました。`);
}

/**
 * 2. 【スキャン実行】
 * 現在のアクティブなシート（または全シート）の状態を読み取り、
 * runSetupで使える形式のJSONログを出力します。
 */
function getValidationConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const exportData = [];

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    // データがある範囲を取得（負荷軽減のため、最大行まで見ない方が良いが、今回はgetDataRangeで）
    const dataRange = sheet.getDataRange();
    const validations = dataRange.getDataValidations();
    const numRows = validations.length;
    const numCols = numRows > 0 ? validations[0].length : 0;
    
    const sheetRules = [];

    // セル走査してValidationを探す
    // 連続した同じルールはまとめるロジックを入れるときれいですが、
    // 実装を簡単にするため、1セルずつ or 行単位で見る簡易版にします。
    // 今回は「連続領域の結合」簡易ロジックを入れます。
    
    for (let r = 0; r < numRows; r++) {
      for (let c = 0; c < numCols; c++) {
        const rule = validations[r][c];
        if (!rule) continue;

        const criteriaType = rule.getCriteriaType();
        let type = null;
        let values = null;

        if (criteriaType === SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
          type = 'CHECKBOX';
        } else if (criteriaType === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
          type = 'LIST';
          values = rule.getCriteriaValues()[0]; // リストの配列
        }
        
        if (type) {
            sheetRules.push({
              row: r + 1, // 1-indexed
              col: c + 1,
              numRows: 1,
              numCols: 1,
              type: type,
              values: values
            });
        }
      }
    }
    
    // 単純なリストだと膨大になるので、ここでの最適化（結合）は省略し、
    // まずは動作優先でそのまま出力します。
    // ※実運用で重すぎる場合は、隣接セル結合ロジックを追加します。
    
    if (sheetRules.length > 0) {
      exportData.push({
        sheetName: sheetName,
        ranges: sheetRules
      });
    }
  });
  
  // ログ出力
  const jsonString = JSON.stringify(exportData);
  Logger.log('▼▼▼ 下記のコードをコピーしてください ▼▼▼');
  Logger.log(jsonString);
  Logger.log('▲▲▲ コピー範囲終了 ▲▲▲');
  
  SpreadsheetApp.getUi().alert('スキャン完了。ログを確認してください (表示 > ログ)');
}

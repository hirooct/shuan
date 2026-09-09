function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // ==========================================
  // 【追加】「はじめに」シートのチェック確認
  // ==========================================
  const introSheet = ss.getSheetByName("はじめに");
  const CHECKBOX_CELL = "B14"; // ★チェックボックスのセルをB14に設定しました
  
  if (introSheet) {
    const isChecked = introSheet.getRange(CHECKBOX_CELL).getValue();
    
    // チェックが入っていない（false）場合は、「はじめに」シートを表示して処理を終了
    if (!isChecked) {
      ss.setActiveSheet(introSheet);
      return;
    }
  }
  
  // ==========================================
  // 【既存】チェックがある場合の「週案」移動処理
  // ==========================================
  const sheet = ss.getSheetByName("週案"); 
  
  if (!sheet) return;
  
  const today = new Date();
  today.setHours(0, 0, 0, 0); // 時間をリセット
  
  // 10行目のデータを取得（B列から最後の列まで）
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 2) return;
  
  const dateRange = sheet.getRange(10, 2, 1, lastColumn - 1);
  const values = dateRange.getValues()[0];
  
  // 今日と同じ日付の列を探す
  for (let i = 0; i < values.length; i++) {
    let cellValue = values[i];
    
    if (cellValue instanceof Date) {
      cellValue.setHours(0, 0, 0, 0);
      
      if (cellValue.getTime() === today.getTime()) {
        const targetColumn = i + 2; // B列（2列目）のズレ調整
        
        // 1. まず、最終列のさらに右（ダミー）を選択して、画面を右いっぱいにスクロールさせます
        sheet.setActiveSelection(sheet.getRange(10, lastColumn));
        SpreadsheetApp.flush(); // 画面の動きを確定
        
        // 2. その後、本来の「今日の日付」を選択します。
        // 右側から戻る動きになるため、確実に「左端」に表示されるようになります。
        sheet.setActiveSelection(sheet.getRange(10, targetColumn));
        break;
      }
    }
  }
}

/**
 * 週案・設定・通信データシートの特定のデータ範囲を初期化（クリア）する関数
 */
function initializeWeeklyPlan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheetName = "週案";
  const configSheetName = "設定";
  const dataSheetName = "通信データ";
  
  const sheet = ss.getSheetByName(targetSheetName);
  const configSheet = ss.getSheetByName(configSheetName);
  const dataSheet = ss.getSheetByName(dataSheetName);
  
  if (!sheet) {
    Browser.msgBox("エラー", `「${targetSheetName}」シートが見つかりません。`, Browser.Buttons.OK);
    return;
  }
  if (!configSheet) {
    Browser.msgBox("エラー", `「${configSheetName}」シートが見つかりません。`, Browser.Buttons.OK);
    return;
  }
  if (!dataSheet) {
    Browser.msgBox("エラー", `「${dataSheetName}」シートが見つかりません。`, Browser.Buttons.OK);
    return;
  }
  
  // ユーザーへの確認アラート（誤操作防止）
  const response = Browser.msgBox("確認", "指定されたすべてのシートのデータを初期化してもよろしいですか？", Browser.Buttons.YES_NO);
  if (response !== "yes") {
    return; // 「いいえ」が押された場合は処理を中断
  }

  // ==========================================
  // 1. 【週案シート】のクリア処理
  // ==========================================
  const lastColumn = sheet.getLastColumn();
  if (lastColumn >= 2) { // B列（2列目）以降が存在する場合
    // ① 9行目のB列から最終列までをクリア
    sheet.getRange(9, 2, 1, lastColumn - 1).clearContent();
    
    // ② B12行目から最終列・32行目までのクリア（12行目から21行分）
    sheet.getRange(12, 2, 21, lastColumn - 1).clearContent();
  }
  
  // ==========================================
  // 2. 【設定シート】のクリア処理
  // ==========================================
  const configLastRow = configSheet.getLastRow();
  
  // ① F列, G列, H列, I列の2行目から下をクリア
  if (configLastRow >= 2) {
    const numRowsToClear = configLastRow - 2 + 1; // 2行目から最終行までの行数
    configSheet.getRange(2, 6, numRowsToClear, 1).clearContent(); // F列 (6列目)
    configSheet.getRange(2, 7, numRowsToClear, 1).clearContent(); // G列 (7列目)
    configSheet.getRange(2, 8, numRowsToClear, 1).clearContent(); // H列 (8列目)
    configSheet.getRange(2, 9, numRowsToClear, 1).clearContent(); // I列 (9列目)
  }
  
  // ② AB列の3行目から下をクリア
  if (configLastRow >= 3) {
    const numRowsToClearAB = configLastRow - 3 + 1; // 3行目から最終行までの行数
    configSheet.getRange(3, 28, numRowsToClearAB, 1).clearContent(); // AB列 (28列目)
  }
  
  // ③ 固定範囲のクリア
  configSheet.getRange("M3:P4").clearContent(); // M3:P4
  configSheet.getRange("V3:Z8").clearContent(); // V3:Z8
  configSheet.getRange("U13").clearContent();   // U13
  
 // ==========================================
  // 3. 【通信データシート】のクリア処理
  // ==========================================
  const dataLastRow = dataSheet.getLastRow();
  // B列の1行目から下をクリア（データが存在する場合のみ）
  if (dataLastRow >= 1) {
    const numRowsToClearData = dataLastRow - 1 + 1; // 1行目から最終行までの行数（つまり dataLastRow と同じ）
    dataSheet.getRange(1, 2, numRowsToClearData, 1).clearContent(); // B列 (2列目)の1行目からクリア
  }
  
  // 完了メッセージ
  Browser.msgBox("完了", "すべての指定範囲の初期化が完了しました。", Browser.Buttons.OK);
}
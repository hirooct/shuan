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
 * スプレッドシート側から実行する場合も、安全な初期化処理へ統一する
 */
function initializeWeeklyPlan() {
  const first = Browser.msgBox(
    '配付用初期化',
    '入力データを初期化します。シート構造と数式は残りますが、入力内容は元に戻せません。続けますか？',
    Browser.Buttons.YES_NO
  );
  if (first !== 'yes') return;

  const confirmText = Browser.inputBox(
    '確認文字の入力',
    '本当に初期化する場合だけ「初期化する」と入力してください。',
    Browser.Buttons.OK_CANCEL
  );
  if (confirmText === 'cancel') return;

  const finalConfirm = Browser.msgBox(
    '最終確認',
    '本当に初期化しても大丈夫ですか？この操作は元に戻せません。',
    Browser.Buttons.YES_NO
  );
  if (finalConfirm !== 'yes') return;

  const result = initializeForDistribution(confirmText);
  Browser.msgBox(
    result.success ? '完了' : '初期化しませんでした',
    result.message || result.error || '処理結果を確認できませんでした。',
    Browser.Buttons.OK
  );
}

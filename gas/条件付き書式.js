function generateDotArtWithConditionalFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetId = sheet.getSheetId();
  const ui = SpreadsheetApp.getUi();
  
  // プロパティサービス（進捗の記録用ストレージ）
  const props = PropertiesService.getScriptProperties();
  
  // =================================================================
  // 1. 保存された進捗（何行目まで終わったか）を読み込む
  // =================================================================
  let startRow = parseInt(props.getProperty("LAST_PROCESSED_ROW") || "0", 10);
  
  if (startRow === 0) {
    ss.toast("新規処理として開始します...", "スタート", -1);
  } else {
    ss.toast(`前回の続き（${startRow + 1}行目）から再開します...`, "レジューム再開", -1);
  }

  const folderName = "mZXjn09dp2zurWN49atK";
  const fileName = "a.txt";

  let compressedText;
  try {
    const folders = DriveApp.getFoldersByName(folderName);
    if (!folders.hasNext()) return;
    const folder = folders.next();
    const files = folder.getFilesByName(fileName);
    if (!files.hasNext()) return;
    compressedText = files.next().getBlob().getDataAsString("UTF-8").trim();
  } catch (e) {
    ui.alert("読み込みエラー", e.message, ui.ButtonSet.OK);
    return;
  }

  // =================================================================
  // 2. zstdの解凍処理（GAS特有の型変換に対応）
  // =================================================================
  let jsonStr;
  try {
    const gasBytes = Utilities.base64Decode(compressedText);
    const uint8Array = new Uint8Array(gasBytes.length);
    for (let i = 0; i < gasBytes.length; i++) {
      let b = gasBytes[i];
      uint8Array[i] = b < 0 ? b + 256 : b;
    }
    
    const cdnUrl = "https://cdn.jsdelivr.net/npm/fzstd@0.1.1/umd/index.js";
    const response = UrlFetchApp.fetch(cdnUrl);
    const script = response.getContentText();
    eval(script); 
    
    if (typeof fzstd === 'undefined' || !fzstd.decompress) {
      throw new Error("fzstd ライブラリの読み込みに失敗しました。");
    }
    
    const decompressedUint8 = fzstd.decompress(uint8Array);
    jsonStr = Utilities.newBlob(decompressedUint8).getDataAsString("UTF-8");
  } catch (e) {
    ui.alert("zstd解凍エラー", "データの解凍に失敗しました: " + e.message, ui.ButtonSet.OK);
    return;
  }

  const data = JSON.parse(jsonStr);

  const listData = data.list; 
  const imgData = data.img;   
  const numRows = imgData.length;      
  const numCols = imgData[0].length;   

  // 全て処理し終わっている場合の安全弁
  if (startRow >= numRows) {
    ui.alert("完了済み", "このデータはすでに最後まで描画されています。リセットする場合は専用関数を実行してください。", ui.ButtonSet.OK);
    return;
  }

  const targetRange = sheet.getRange(1, 1, numRows, numCols);

  // 初回実行時のみシートの初期化、サイズ拡張、条件付き書式ルールの設定を行う
  if (startRow === 0) {
    const currentMaxRows = sheet.getMaxRows();
    const currentMaxColumns = sheet.getMaxColumns();
    if (numRows > currentMaxRows) sheet.insertRowsAfter(currentMaxRows, numRows - currentMaxRows);
    if (numCols > currentMaxColumns) sheet.insertColumnsAfter(currentMaxColumns, numCols - currentMaxColumns);
    
    // シートの初期化 (書式・テキストのクリア、文字サイズを最小に)
    sheet.clearConditionalFormatRules();
    targetRange.setBackground(null)
               .setFontColor(null)
               .setFontSize(1)
               .setHorizontalAlignment("center")
               .setVerticalAlignment("middle")
               .clearContent();
    
    for (let col = 1; col <= numCols; col++) sheet.setColumnWidth(col, 20);
    for (let row = 1; row <= numRows; row++) sheet.setRowHeight(row, 20);

    // =================================================================
    // 3. 【追加】条件付き書式ルールの作成と適用 (初回のみ一括実行)
    // =================================================================
    ss.toast("条件付き書式ルールを設定中...", "準備", -1);
    const rules = [];
    for (let i = 0; i < listData.length; i++) {
      const item = listData[i];
      const colorId = parseInt(item.id, 10);
      
      const toHex = (c) => {
        const hex = Math.max(0, Math.min(255, c)).toString(16);
        return hex.length === 1 ? "0" + hex : hex;
      };
      const hexColor = "#" + toHex(item.r) + toHex(item.g) + toHex(item.b);

      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberEqualTo(colorId)
        .setBackground(hexColor)
        .setFontColor(hexColor) // 文字色も同色にすることで、セル内の数値を隠します
        .setRanges([targetRange]) 
        .build();
      rules.push(rule);
    }
    sheet.setConditionalFormatRules(rules);
  }

  // =================================================================
  // 4. chunkに分割して「数値（カラーID）」を書き込むループ
  // =================================================================
  const CHUNK_SIZE = 80; 

  for (let r = startRow; r < numRows; r += CHUNK_SIZE) {
    const endRow = Math.min(r + CHUNK_SIZE, numRows);
    const valueGridChunk = [];

    // このchunkの数値データを構築
    for (let current_r = r; current_r < endRow; current_r++) {
      const valueRow = [];
      for (let c = 0; c < numCols; c++) {
        const val = parseInt(imgData[current_r][c], 10);
        valueRow.push(isNaN(val) ? "" : val);
      }
      valueGridChunk.push(valueRow);
    }

    // 該当する行の範囲に、一撃で数値を流し込む (条件付き書式によって自動で色が変わる)
    const chunkRange = sheet.getRange(r + 1, 1, valueGridChunk.length, numCols);
    chunkRange.setValues(valueGridChunk);

    // --- チェックポイントを記録 ---
    props.setProperty("LAST_PROCESSED_ROW", String(endRow));
    
    // 進捗パーセントの計算と表示
    const percent = Math.floor((endRow / numRows) * 100);
    ss.toast(`${endRow}行目まで数値を書き込みました（自動着色中）`, `進捗: ${percent}%`, -1);
  }

  // 完全に終わったら記録を消去（次回また1行目からスタートできるようにする）
  props.deleteProperty("LAST_PROCESSED_ROW");
  
  ss.toast("条件付き書式によるドット絵描画が完了しました！", "完了", 5);
}

// 途中で最初から強制的にやり直したいときに実行する救済関数
function resetResumeStatus() {
  PropertiesService.getScriptProperties().deleteProperty("LAST_PROCESSED_ROW");
  SpreadsheetApp.getActiveSpreadsheet().toast("進捗記録をリセットしました。次回は1行目から始まります。", "リセット完了", 5);
}

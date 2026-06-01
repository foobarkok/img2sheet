function generateDotArtWithResumeBackgroundOnly() {
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
  // 【確定版】zstdの解凍処理（GAS特有の型変換に対応）
  // =================================================================
  let jsonStr;
  try {
    // 1. Base64デコード（GAS特有の、符号付き-128〜127のByte配列が返る）
    const gasBytes = Utilities.base64Decode(compressedText);
    
    // 2. GASの特殊配列を、JavaScript標準の「Uint8Array(0〜255)」に厳密に変換
    const uint8Array = new Uint8Array(gasBytes.length);
    for (let i = 0; i < gasBytes.length; i++) {
      let b = gasBytes[i];
      uint8Array[i] = b < 0 ? b + 256 : b;
    }
    
    // 3. fzstdライブラリをCDNから読み込んで実行
    const cdnUrl = "https://cdn.jsdelivr.net/npm/fzstd@0.1.1/umd/index.js";
    const response = UrlFetchApp.fetch(cdnUrl);
    const script = response.getContentText();
    eval(script); 
    
    if (typeof fzstd === 'undefined' || !fzstd.decompress) {
      throw new Error("fzstd ライブラリの読み込みに失敗しました。");
    }
    
    // 4. 正しい型（Uint8Array）でzstd解凍を実行
    const decompressedUint8 = fzstd.decompress(uint8Array);
    
    // 5. 解凍されたUint8Array（バイナリ）をUTF-8文字列（JSON）に復元
    // Utilities.newBlob() は Uint8Array をそのまま受け取れるので、これで一発変換できます
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

  // カラーパレットのマッピング（RGB割合）
  const colorMap = {};
  for (let i = 0; i < listData.length; i++) {
    const item = listData[i];
    colorMap[item.id] = { red: item.r / 255, green: item.g / 255, blue: item.b / 255 };
  }

  // 初回実行時のみシートの初期化と拡張を行う
  if (startRow === 0) {
    const currentMaxRows = sheet.getMaxRows();
    const currentMaxColumns = sheet.getMaxColumns();
    if (numRows > currentMaxRows) sheet.insertRowsAfter(currentMaxRows, numRows - currentMaxRows);
    if (numCols > currentMaxColumns) sheet.insertColumnsAfter(currentMaxColumns, numCols - currentMaxColumns);
    sheet.clearConditionalFormatRules();
    
    // 【背景色のみの最適化】初回の全体の初期化段階で、全セルの値（テキスト）をクリア
    sheet.getRange(1, 1, numRows, numCols).clearContent();
    
    for (let col = 1; col <= numCols; col++) sheet.setColumnWidth(col, 20);
    for (let row = 1; row <= numRows; row++) sheet.setRowHeight(row, 20);
  }

  // =================================================================
  // 2. 100行ずつ分割して背景色のみを書き込むループ
  // =================================================================
  const CHUNK_SIZE = 100; // 100行ずつ処理

  for (let r = startRow; r < numRows; r += CHUNK_SIZE) {
    const endRow = Math.min(r + CHUNK_SIZE, numRows);
    const apiRows = [];

    // 100行分のデータを構築
    for (let current_r = r; current_r < endRow; current_r++) {
      const rowValues = [];
      for (let c = 0; c < numCols; c++) {
        const val = parseInt(imgData[current_r][c], 10);
        const colorId = isNaN(val) ? "" : val;
        const rgb = colorMap[colorId] || { red: 1, green: 1, blue: 1 }; // 無ければ白

        rowValues.push({
          userEnteredFormat: {
            backgroundColor: rgb,
            wrapStrategy: "CLIP"
          }
        });
      }
      apiRows.push({ values: rowValues });
    }

    // Sheets API（updateCells）で100行分を一撃書き込み
    const requests = [{
      updateCells: {
        rows: apiRows,
        fields: "userEnteredFormat.backgroundColor,userEnteredFormat.wrapStrategy",
        range: {
          sheetId: sheetId,
          startRowIndex: r, // 開始行（0始まり）
          endRowIndex: endRow,
          startColumnIndex: 0,
          endColumnIndex: numCols
        }
      }
    }];

    // Googleサーバーへの即時反映
    Sheets.Spreadsheets.batchUpdate({ requests: requests }, ss.getId());

    // --- チェックポイントを記録 ---
    props.setProperty("LAST_PROCESSED_ROW", String(endRow));
    
    // 進捗パーセントの計算と表示
    const percent = Math.floor((endRow / numRows) * 100);
    ss.toast(`${endRow}行目まで背景色を塗り終えました`, `進捗: ${percent}%`, -1);
  }

  // 完全に終わったら記録を消去（次回また1行目からスタートできるようにする）
  props.deleteProperty("LAST_PROCESSED_ROW");
  
  ss.toast("背景色のみでのドット絵描画が完了しました！", "完了", 5);
}

// 途中で最初から強制的にやり直したいときに実行する救済関数
function resetResumeStatus() {
  PropertiesService.getScriptProperties().deleteProperty("LAST_PROCESSED_ROW");
  SpreadsheetApp.getActiveSpreadsheet().toast("進捗記録をリセットしました。次回は1行目から始まります。", "リセット完了", 5);
}

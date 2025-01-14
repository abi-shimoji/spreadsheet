const TEST_NO_SHEET_ROW_GAP = 5;

function createEvidenceSheet() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName("テスト項目");

  const latestRowNo = sheet.getLastRow();
  const testFinalNo = latestRowNo - TEST_NO_SHEET_ROW_GAP;

  for (let i = 1; i <= testFinalNo; i++) {
    const newSheetName = `エビデンス${i}`;

    if (spreadSheet.getSheetByName(newSheetName)) {
      continue;
    }

    // テスト項目詳細取得
    let cat1 = sheet.getRange(`C${i + TEST_NO_SHEET_ROW_GAP}`).getValue(); // 分類1
    let cat2 = sheet.getRange(`D${i + TEST_NO_SHEET_ROW_GAP}`).getValue(); // 分類2
    let cat3 = sheet.getRange(`E${i + TEST_NO_SHEET_ROW_GAP}`).getValue(); // 分類3
    let prerequisites = sheet.getRange(`F${i + TEST_NO_SHEET_ROW_GAP}`).getValue(); // 前提手順
    let operatigIns = sheet.getRange(`G${i + TEST_NO_SHEET_ROW_GAP}`).getValue(); // 操作手順
    let expectedResults = sheet.getRange(`H${i + TEST_NO_SHEET_ROW_GAP}`).getValue(); // 期待結果

    if (!cat1) {
      for (let j = i - 1; TEST_NO_SHEET_ROW_GAP !== j + TEST_NO_SHEET_ROW_GAP; j--) {
        cat1 = sheet.getRange(`C${j + TEST_NO_SHEET_ROW_GAP}`).getValue();
        if (cat1) break;
      }
    }
    if (!cat2) {
      for (let j = i - 1; TEST_NO_SHEET_ROW_GAP !== j + TEST_NO_SHEET_ROW_GAP; j--) {
        cat2 = sheet.getRange(`D${j + TEST_NO_SHEET_ROW_GAP}`).getValue();
        if (cat2) break;
      }
    }
    if (!cat3) {
      for (let j = i - 1; TEST_NO_SHEET_ROW_GAP !== j + TEST_NO_SHEET_ROW_GAP; j--) {
        cat3 = sheet.getRange(`E${j + TEST_NO_SHEET_ROW_GAP}`).getValue();
        if (cat3) break;
      }
    }
    if (!prerequisites) {
      for (let j = i - 1; TEST_NO_SHEET_ROW_GAP !== j + TEST_NO_SHEET_ROW_GAP; j--) {
        prerequisites = sheet.getRange(`F${j + TEST_NO_SHEET_ROW_GAP}`).getValue();
        if (prerequisites) break;
      }
    }
    if (!operatigIns) {
      for (let j = i - 1; TEST_NO_SHEET_ROW_GAP !== j + TEST_NO_SHEET_ROW_GAP; j--) {
        operatigIns = sheet.getRange(`G${j + TEST_NO_SHEET_ROW_GAP}`).getValue();
        if (operatigIns) break;
      }
    }
    if (!expectedResults) {
      for (let j = i - 1; TEST_NO_SHEET_ROW_GAP !== j + TEST_NO_SHEET_ROW_GAP; j--) {
        expectedResults = sheet.getRange(`H${j + TEST_NO_SHEET_ROW_GAP}`).getValue();
        if (expectedResults) break;
      }
    }

    // シート追加
    const newSheet = spreadSheet.insertSheet(newSheetName);
    
    // 追加したシートを最後尾に移動
    const sheets = spreadSheet.getSheets();
    spreadSheet.setActiveSheet(newSheet);
    spreadSheet.moveActiveSheet(sheets.length);

    // 新しいシートテスト項目詳細を設定
    newSheet.getRange(`B2`).setValue(`分類1`);
    newSheet.getRange(`C2`).setValue(cat1);
    newSheet.getRange(`B3`).setValue(`分類2`);
    newSheet.getRange(`C3`).setValue(cat2);
    newSheet.getRange(`B4`).setValue(`分類3`);
    newSheet.getRange(`C4`).setValue(cat3);
    newSheet.getRange(`B5`).setValue(`前提手順`);
    newSheet.getRange(`C5`).setValue(prerequisites);
    newSheet.getRange(`B6`).setValue(`操作手順`);
    newSheet.getRange(`C6`).setValue(operatigIns);
    newSheet.getRange(`B7`).setValue(`期待結果`);
    newSheet.getRange(`C7`).setValue(expectedResults);

    // 書式設定
    const testItemTitleRange = newSheet.getRange("B2:B7");
    testItemTitleRange.setBackground('#00af98');
    testItemTitleRange.setBorder(true, true, true, true, true, true);
    const testItemValueRange = newSheet.getRange("C2:C7");
    testItemValueRange.setBorder(true, true, true, true, true, true);
    testItemValueRange.setWrap(true);

    // 各列の横幅を設定
    newSheet.setColumnWidth(1, 20);
    newSheet.setColumnWidth(2, 100);
    newSheet.setColumnWidth(3, 600);

    // ハイパーリンク情報
    const newSheetId = newSheet.getSheetId();
    const newSheetUrl = `${spreadSheet.getUrl()}#gid=${newSheetId}`;
    // ハイパーリンク挿入処理
    const linkCell = sheet.getRange(i + TEST_NO_SHEET_ROW_GAP, 9);
    linkCell.setFormula(`=HYPERLINK("${newSheetUrl}", "エビデンス${i}")`);
  }

  // テスト項目書シートに移動
  let sheets = spreadSheet.getSheets();
  spreadSheet.setActiveSheet(sheets[sheet.getIndex() - 1]);
}

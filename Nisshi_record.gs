//入力シートの入力値を記録する
function Nisshi_record(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetN = ss.getSheetByName('入力');
  const sheetR = ss.getSheetByName('記録');
  const kyusyoku = sheetN.getRange("AJ6").getValue();  //給食の有無を取得

  //記録シートの最終記録日の列を取得
  //行の先頭列から右方向に取得するコード
  let lastCol = sheetR.getRange(3, 1).getNextDataCell(SpreadsheetApp.Direction.NEXT).getColumn();
  console.log("最終列=",lastCol);
  if (lastCol != 1){
    if (lastCol == sheetR.getMaxColumns() ){
      sheetR.insertColumnAfter(lastCol);
      console.log("1行追加しました");
    }
  }
  //入力データの最終行を取得
  let lastRow = sheetN.getRange(sheetN.getMaxRows(), 49).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  //日付の取得
  let today = sheetN.getRange("AW3").getValue().setHours(0,0,0,0);  //今日の日付
  if (sheetR.getRange(3,lastCol).getValue()!="月日"){
    var past = sheetR.getRange(3,lastCol).getValue().setHours(0,0,0,0);   //記録シートの最終日
  }else{var past = "";}
  if (past == today){
    let copyRange = sheetN.getRange(1,49,lastRow,1);
    let pasteRange = sheetR.getRange(1,lastCol);
    copyRange.copyTo(pasteRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    sheetR.getRange(1,lastCol).setFormulaR1C1("RC[-1]+1");
    if (kyusyoku ==""){ //給食なしの日は給食のデータを記録しない
      sheetR.getRange(267,lastCol , 90 ,1).clear();
    };
  }else {
    let copyRange = sheetN.getRange(1,49,lastRow,1);
    let pasteRange = sheetR.getRange(1,lastCol +1 );
    copyRange.copyTo(pasteRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    sheetR.getRange(1,lastCol+1).setFormulaR1C1("RC[-1]+1");
    if (kyusyoku ==""){sheetR.getRange(267,lastCol+1 , 90 ,1).clear()}; //給食なしの日は給食のデータを記録しない
  }
}

//入力シートの入力値の記録Backupをする
function Nisshi_record_backup(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetN = ss.getSheetByName('入力');
  const sheetR = ss.getSheetByName('記録backup');

  //記録シートの最終記録日の列を取得
  //行の先頭列から右方向に取得するコード
  let lastCol = sheetR.getRange(3, 1).getNextDataCell(SpreadsheetApp.Direction.NEXT).getColumn();
  if (lastCol != 1){
    if (lastCol == sheetR.getMaxColumns() ){
      sheetR.insertColumnAfter(lastCol);
      console.log("1行追加しました");
    }
  }
  //入力データの最終行を取得
  let lastRow = sheetN.getRange(sheetN.getMaxRows(), 49).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  //日付の取得
  let today = sheetN.getRange("AW3").getValue().setHours(0,0,0,0);
  let past = sheetR.getRange(3,lastCol).getValue().setHours(0,0,0,0);

  if (past == today){
    let copyRange = sheetN.getRange(1,49,lastRow,1);
    let pasteRange = sheetR.getRange(1,lastCol);
    copyRange.copyTo(pasteRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    sheetR.getRange(1,lastCol).setFormulaR1C1("RC[-1]+1");
  }else {
    let copyRange = sheetN.getRange(1,49,lastRow,1);
    let pasteRange = sheetR.getRange(1,lastCol +1 );
    copyRange.copyTo(pasteRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    sheetR.getRange(1,lastCol+1).setFormulaR1C1("RC[-1]+1");
  }
}

//読込シートの情報を上書き
function Gakko_Overwrite() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNY = ss.getSheetByName('学校読込');
  const sheetR = ss.getSheetByName('記録');
  const write_column = sheetNY.getRange('Y2').getValue();
  //学校日誌読込の修正データをコピー、記録に貼り付け
  let copyRange = sheetNY.getRange("Z6:Z368");
  let pasteRange = sheetR.getRange(6,write_column);
  copyRange.copyTo(pasteRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  sheetNY.getRange("Y3").activate()
}
function Hoken_Overwrite() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNY = ss.getSheetByName('保健読込');
  const sheetR = ss.getSheetByName('記録');
  const write_row = sheetNY.getRange('Y5').getValue();
  const write_column = sheetNY.getRange('AA2').getValue();
  //学校日誌読込の修正データをコピー、記録に貼り付け
  let copyRange = sheetNY.getRange("AB6:AB142");
  let pasteRange = sheetR.getRange(write_row+1,write_column);
  copyRange.copyTo(pasteRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  sheetNY.getRange("AA3").activate()
}
function Kyusyoku_Overwrite() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNY = ss.getSheetByName('給食読込');
  const sheetR = ss.getSheetByName('記録');
  const write_row = sheetNY.getRange('AD5').getValue();
  const write_column = sheetNY.getRange('AF2').getValue();
  //学校日誌読込の修正データをコピー、記録に貼り付け
  let copyRange = sheetNY.getRange("AG6:AG96");
  let pasteRange = sheetR.getRange(write_row+1,write_column);
  copyRange.copyTo(pasteRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  sheetNY.getRange("AF3").activate()
}

// ************通常シートの関数を復元************
function Gakko_restore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNY = ss.getSheetByName('学校');

  //学校日誌記録読込の枠データをコピー、貼り付け
  let copyRange = sheetNY.getRange("AB2:AV38");
  let pasteRange = sheetNY.getRange("B2");
  copyRange.copyTo(pasteRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  sheetNY.getRange("Y3").activate()
}
function Kyusyoku_restore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNY = ss.getSheetByName('給食');

  //給食日誌記録読込の枠データをコピー、貼り付け
  let copyRange = sheetNY.getRange("AI2:BJ38");
  let pasteRange = sheetNY.getRange("B2");
  copyRange.copyTo(pasteRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  sheetNY.getRange("AF3").activate()
}
function Hoken_restore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNY = ss.getSheetByName('保健');

  //日誌記録読込の枠データをコピー、貼り付け
  let copyRange = sheetNY.getRange("AC2:AW65");
  let pasteRange = sheetNY.getRange("B2");
  copyRange.copyTo(pasteRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  //備考のフォントサイズを元に戻す
  cell1 = sheetNY.getRange(29,3,4,20);
  cell1.setFontSize(9);
  cell2 = sheetNY.getRange(33,3,31,20);
  cell2.setFontSize(8);

  //表示設定を解除
  sheetNY.showRows(29,34);
  // 29行から34行を25ピクセルの高さに設定します
  sheetNY.setRowHeights(29, 35, 25);
  sheetNY.getRange("C3").activate()
}

// ************読み込みシートの関数復元****************
function Gakko_Read_restore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNY = ss.getSheetByName('学校読込');

  //学校日誌記録読込の枠データをコピー、貼り付け
  let copyRange = sheetNY.getRange("AA2:AU38");
  let pasteRange = sheetNY.getRange("B2");
  copyRange.copyTo(pasteRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  sheetNY.getRange("Y3").activate()
}

function Kyusyoku_Read_restore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNY = ss.getSheetByName('給食読込');

  //給食日誌記録読込の枠データをコピー、貼り付け
  let copyRange = sheetNY.getRange("AH2:BI38");
  let pasteRange = sheetNY.getRange("B2");
  copyRange.copyTo(pasteRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  sheetNY.getRange("AF3").activate()
}

function Hoken_Read_restore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNY = ss.getSheetByName('保健読込');

  //日誌記録読込の枠データをコピー、貼り付け
  let copyRange = sheetNY.getRange("AC2:AW65");
  let pasteRange = sheetNY.getRange("B2");
  copyRange.copyTo(pasteRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  //備考のフォントサイズを元に戻す
  cell1 = sheetNY.getRange(29,3,4,20);
  cell1.setFontSize(9);
  cell2 = sheetNY.getRange(33,3,31,20);
  cell2.setFontSize(8);

  //表示設定を解除
  sheetNY.showRows(29,34);
  // 29行から34行を25ピクセルの高さに設定します
  sheetNY.setRowHeights(29, 35, 25);
  sheetNY.getRange("C3").activate()
}

//***********入力シートの関数復元****************
function Nyuryoku_restore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheet = ss.getSheetByName('入力');
  spreadsheet.getRange('B6').activate();
  spreadsheet.getRange('B55:AS97').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  
  spreadsheet.getRangeList(['G3','C51:E51']).activate()
  .clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('G3').activate();
};

//*********すべての入力シートの関数を復元*****************
function all_restore(){
  Gakko_restore();
  Kyusyoku_restore();
  Hoken_restore();
  Gakko_Read_restore();
  Kyusyoku_Read_restore();
  Hoken_Read_restore();
  Nyuryoku_restore();
}

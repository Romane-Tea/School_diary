//欠席情報を記録する
function kesseki_record(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetN = ss.getSheetByName('欠席入力');
  const sheetR = ss.getSheetByName('欠席記録');

  //欠席記録の最終記録日の列を取得
  var lastCol = sheetR.getRange(8, sheetR.getMaxColumns()).getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).getColumn();
  //欠席入力データの最終行を取得
  var lastRow = sheetN.getRange(sheetN.getMaxRows(), 1).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

  //日付の取得
  var today = sheetN.getRange("J3").getValue().setHours(0,0,0,0);
  var todayH = sheetN.getRange("J3").getValue();
  var past = sheetR.getRange(8,lastCol).getValue().setHours(0,0,0,0);

  if (past == today){
      paste(lastCol-2,lastRow,todayH,'欠席記録');
  }else{ paste(lastCol +1,lastRow,todayH,'欠席記録'); }
}

//欠席情報を記録をBackupする
function kesseki_record_backup(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetN = ss.getSheetByName('欠席入力');
  const sheetRB = ss.getSheetByName('欠席記録backup');

  //欠席記録の最終記録日の列を取得
  var lastCol = sheetRB.getRange(8, sheetRB.getMaxColumns()).getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).getColumn();
  //欠席入力データの最終行を取得
  var lastRow = sheetN.getRange(sheetN.getMaxRows(), 1).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

  //日付の取得
  var today = sheetN.getRange("J3").getValue().setHours(0,0,0,0);
  var todayH = sheetN.getRange("J3").getValue();
  var past = sheetRB.getRange(8,lastCol).getValue().setHours(0,0,0,0);

  if (past == today){
      paste(lastCol-2,lastRow,todayH,'欠席記録backup');
  }else{ paste(lastCol +1,lastRow,todayH,'欠席記録backup'); }
}

//データ貼り付け関数
function paste(lastCol,lastRow,todayH,sheet_name){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetN = ss.getSheetByName('欠席入力');
  const sheetR = ss.getSheetByName(sheet_name);

  //授業日「〇」の取得
  var study_day = sheetN.getRange("G8").getValue();
  var stady_past = sheetR.getRange(7,lastCol,1,3);   //記録シートに〇を貼り付けるセルを指定

  if (lastCol < 7){
    lastCol = 7;
  }

  //日付を貼付け
  sheetR.getRange(8,lastCol,1,3).setValue(todayH);

  //欠席入力のデータをコピー、貼り付け
  let copyRange = sheetN.getRange(9,6,lastRow -8 ,3);
  let pasteRange = sheetR.getRange(9,lastCol );
  copyRange.copyTo(pasteRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  //集計表をコピー、貼り付け
  if (lastCol !=7){
    let copyRangeH = sheetR.getRange(1,7,6,3);
    let pasteRangeH = sheetR.getRange(1,lastCol,6,3);
    copyRangeH.copyTo(pasteRangeH);
  }
  itiranPaste(lastCol,sheet_name);
  stady_past.setValue(study_day);  //記録シートに〇を貼り付ける

}

//欠席記録に一覧を作成
function itiranPaste(lastCol,sheet_name){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetR = ss.getSheetByName(sheet_name);
  const sheetI = ss.getSheetByName('一覧表');

  //欠席一覧表のデータをコピー、貼り付け
  let copyRangeI = sheetI.getRange(3,8,122 ,3);
  let pasteRangeI = sheetR.getRange(515,lastCol );
  copyRangeI.copyTo(pasteRangeI, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
}

//欠席入力の入力値をクリア
function Kesseki_clear(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('欠席入力');
  let lastRow = sheet.getRange(sheet.getMaxRows(), 1).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  sheet.getRange(10,6,lastRow,4).clearContent().setBackground(null);
}

function Kesseki_clear_hand(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('欠席入力');

  var result=Browser.msgBox("欠席情報をすべて消しますか？この作業は元に戻せません。",Browser.Buttons.YES_NO);
  if(result=="no"){
    Browser.msgBox("中止しました。");
  }else{
  //欠席入力データの最終行を取得
    Kesseki_clear()
    Browser.msgBox("すべて消去しました。");
  }
}
/******* 欠席修正 ********/
// 欠席数式復元
function Re_kesseki_syusei_funk() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheet = ss.getSheetByName('欠席修正');
  const seito_su = spreadsheet.getRange('N1').getValue();
  // 集計表と1人目の数式をコピー
  spreadsheet.getRange('F1').activate();
  spreadsheet.getRange('AD1:AF10').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  //数式を人数分コピー
  spreadsheet.getRange(11,6,seito_su,3).activate();
  spreadsheet.getRange('F10:H10').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  spreadsheet.getRange(10,6,seito_su,2).setBackground(null);
  spreadsheet.getRange('F10').activate();
};

// 欠席修正を実行
function Re_kesseki_paste() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetS = ss.getSheetByName('欠席修正');
  const sheetR = ss.getSheetByName('欠席記録');

  //欠席記録の記録列を取得
  var targetCol = sheetS.getRange('M1').getValue();
  //欠席入力データの最終行を取得
  var lastRow = sheetS.getRange(sheetS.getMaxRows(), 1).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

  //欠席修正のデータをコピー、貼り付け
  let copyRange = sheetS.getRange(9,6,lastRow -8 ,3);
  let pasteRange = sheetR.getRange(9,targetCol );
  copyRange.copyTo(pasteRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  const sheetI = ss.getSheetByName('一覧表');

  //欠席一覧表のデータをコピー、貼り付け
  let copyRangeI = sheetI.getRange(3,13,122 ,3);
  let pasteRangeI = sheetR.getRange(515,targetCol );
  copyRangeI.copyTo(pasteRangeI, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

}
/*****   読み込み  ***** */
function Clear_kesseki_kiroku() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('欠席記録読込');
  sheet.getRange('F13').activate();
  sheet.getRange('R13:T513').copyTo(sheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  sheet.getRange('F13').activate();
}

function kesseki_reRecord(){
  //欠席修正情報を上書き
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetY = ss.getSheetByName('欠席記録読込');
  const sheetR = ss.getSheetByName('欠席記録');
  const sheetI = ss.getSheetByName('一覧表');

  //記録する日付を取得
  var record_day = sheetY.getRange('M2').getDisplayValue();
  var record_column = sheetY.getRange('M3').getValue();

  //欠席読込データの最終行を取得
  var lastRow = sheetY.getRange(sheetY.getMaxRows(), 1).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

  //実行するか確認    返り値は"ok"と"cancel"
  var quession = Browser.msgBox("記録を上書きしますか？", Browser.Buttons.OK_CANCEL);

  if (quession == "cancel"){
    console.log(quession +"=cancel");
    return;
  }else if (quession == "ok"){
    console.log(quession+"=OK");
    //処理を継続
  }

  //欠席記録読込のデータをコピー、貼り付け
  let copyRange = sheetY.getRange(13,6,lastRow -12 ,3);
  let pasteRange = sheetR.getRange(11,record_column );
  copyRange.copyTo(pasteRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  //欠席一覧表のデータをコピー、貼り付け
  let copyRangeI = sheetI.getRange(3,13,122 ,3);
  let pasteRangeI = sheetR.getRange(515,record_column );
  copyRangeI.copyTo(pasteRangeI, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  Clear_kesseki_kiroku()
  sheetY.getRange("F13").activate();
}

//　欠席集計の数式を削除
function syukei_delete() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheet = ss.getSheetByName('欠席集計');
  spreadsheet.getRange('11:510').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: false});
  spreadsheet.getRange('B3').activate();
};

//　欠席集計の数式を人数分貼り付け
function syukei_add() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheet = ss.getSheetByName('欠席集計');
  //シートの最終行を取得
  var lastColumn = spreadsheet.getLastColumn();
  //生徒数を取得
  var lastRaw = spreadsheet.getRange('B7').getValue()-1;
  spreadsheet.getRange(11,1,lastRaw,lastColumn).activate();
  spreadsheet.getRange('10:10').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
};

//欠席記録と一覧のズレ日を出力
function zure(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('欠席記録');
  dataList = GetSpreadsheet("欠席記録");  //欠席記録シート全体のデータを取得
  //シートの最終行を取得
  var lastColumn = sheet.getLastColumn();
  for (var i=6 ; i<lastColumn; i +=3 ){
    var one = Number(dataList[1][i])+Number(dataList[1][i+1])+Number(dataList[1][i+2]);
    for (var one_k = 0; one_k < 50; one_k++){
      if (dataList[one_k+514][i] != ""){}else{break;}
    }
    if (one != one_k){console.log(dataList[7][i]+"1年",one,one_k)}

    var two = Number(dataList[2][i])+Number(dataList[2][i+1])+Number(dataList[2][i+2]);
    for (var two_k = 0; two_k < 50; two_k++){
      if (dataList[two_k+514][i+1] != ""){}else{break;}
    }
    if (two != two_k){console.log(dataList[7][i]+"2年",two,two_k)}

    var three = Number(dataList[3][i])+Number(dataList[3][i+1])+Number(dataList[3][i+2]);
    for (var three_k = 0; three_k < 50; three_k++){
      if (dataList[three_k+514][i+2] != ""){}else{break;}
    }
    if (three != three_k){console.log(dataList[7][i]+"3年",three,three_k)}
  }
}

//欠席記録の欠席者一覧を更新する（時間がかかる）
function kesseki_kiroku_update() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetR = ss.getSheetByName('欠席記録');
  let range = sheetR.getDataRange();
  let datas = range.getValues();

  //出席番号格納変数
  var no_0 = 0;

  for (let i=0; i<Math.floor(datas[8].length/3); i++){
    //行格納変数
    let row1 = 514;
    let row2 = 514;
    let row3 = 514;
    for (let j=9; j<datas.length; j++){
      if (datas[j][2]==""){
        break;
      } else if (datas[j][i*3+6]=="欠席" || datas[j][i*3+6]=="忌引" || datas[j][i*3+6]=="出停"){
        let grade = Number(String(datas[j][1])[0]);  //学年
        // console.log(datas[j][1],grade);
        let n_class = String(datas[j][1])[1];  //学級
        if (n_class=="-"){
          n_class = datas[j][1].split("-")[1];
        }else{
          n_class = Number(n_class);
        }
        let name =  datas[j][2];     //名前
        let kesseki = datas[j][i*3+6];  //欠席理由
        let ziyuu = datas[j][i*3+6+1];  //事由
        let hoka = datas[j][i*3+6+2];   //その他
        
        //忌引きや出停により分別
        let teisi = "";
        if (kesseki=="忌引" || kesseki=="出停"){
          teisi = kesseki;
        }
        let no = grade + "-" + n_class;
        if (no == no_0 ){
          no = " 〃 ";
        }else{
          no_0 = no;
        }
        let temp_data ="";
        if (ziyuu=="その他"){
          temp_data = no + " " + name + "(" + hoka + ")" + teisi;
        } else {
          temp_data = no + " " + name + "(" + ziyuu + ")" + teisi;
        }

        //書き込むセル
        let w_row =0;
        let w_col =0;
        if (grade==1){
          w_row = row1;
          w_col = 1;
        }else if(grade==2){
          w_row = row2;
          w_col = 2;
        }else{
          w_row = row3;
          w_col = 3;
        }
        datas[w_row][i*3+w_col+5]=temp_data;
        if (grade==1){row1++}else if(grade==2){row2++}else{row3++};
      }
    }
  }
  let past_datas =[]
  for (let i = 514; i < datas.length; i++) {
    var row = datas[i];
    past_datas.push(row);
  }
  sheetR.getRange(515,1,past_datas.length,past_datas[0].length).setValues(past_datas);
}

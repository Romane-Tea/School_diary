//入力シートの入力値をクリア
function Nisshi_clear() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheet = ss.getSheetByName('入力');
  spreadsheet.getRangeList(['G3:H3', 'C28:J28', 'C30:J30', 'C46:J47','C51:E51', 'R6:S6', 'U6:V6', 'X6:Y6', 'V19:Y23', 'V25:Y32', 'V34:Y35', 'AB19:AC36', 'S23:U23', 'S30:U32', 'S35:U35', 'R37:AA40', 'AF12:AH16', 'AH22:AJ22', 'AG21:AG22', 'AH21:AJ22', 'AH24:AI26', 'AN27']).activate()
  .clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('G3').activate();
}

//隠れた行を再表示する
function Redisplay(sheet){
  if (!sheet ){
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
  }

  //表示設定を解除
  sheet.showRows(29,34);
  // 29行から63行を25ピクセルの高さに設定します
  sheet.setRowHeights(29, 34, 25);
  cell1 = sheet.getRange(29,3,4,20);
  cell1.setFontSize(9);
  cell1.setFontWeight('nomal');
  cell2 = sheet.getRange(33,3,31,20);
  cell2.setFontSize(8);
  cell2.setFontWeight('nomal');
}

//列の調節、空白列を非表示、行の高さを調整
function Re_arrange(sheet){
  if (!sheet ){
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // const sheet = ss.getActiveSheet();
    const sheet = ss.getSheetByName('保健');
  }
  //シートデータの読み込み
  hoken_date = sheet.getDataRange().getValues();

  try{  //シート内にエラーがあるかどうか判断
    // 1行の高さ25×10行(表示可能行)=250px+12px　全部で34行
    // 自由記述欄の欠席以外の最初の4行の表示数
    let free_row = hoken_date[26][23]; //sheet.getRange("X27").getValue();
    // 最初の4行＋欠席者一覧の表示行数を取得
    let show_row = hoken_date[27][23] ; //sheet.getRange("X28").getValue();

    // 欠席以外の記述欄の非表示設定
    if (free_row < 4){ sheet.hideRows(29+free_row,4-free_row); }

    // 12行まで表示、それ以上は高さを縮小する
    if ( show_row < 13 ){
      //欠席者の空白列を,全部で12行になるようにして、残り非表示　欠席表示は33行から62行まで30行
      // 29 + 4 - free_row + show_row +12 -show_row = 45 -free_row
      sheet.hideRows(45 - free_row , 18 - free_row );
    }else{
      //空白行を非表示に 29行目から全体35行の中から、空白となる行を選ぶ。35 -(free_row+4-free_row+show_row-free_row) =row
      sheet.hideRows(29+show_row + 4 - free_row , 30 - show_row + free_row);

      //行数が多い時、高さを調整する
      let row_h = Math.round( 262 / show_row )+1;
      sheet.setRowHeights(31, show_row + 4 - free_row , row_h);

      //行の高さが小さい時、フォントサイズを7にする
      if (row_h < 14){
        sheet.getRange(29,3,34,16).setFontSize(7);
      }
    }

    //出停、忌引を太字にする
    let lists = [3, 11, 17];
    for (let r=33 ; r<33+show_row ;r++){
      for (let col of lists){
        mozi = hoken_date[r-1][col-1]; // mozi = sheet.getRange(r,col).getDisplayValue();
        if ((mozi.match('出停') != null) || (mozi.match('忌引') != null)){
          sheet.getRange(r,col).setFontWeight('bold');
        }else{
          sheet.getRange(r,col).setFontWeight('nomal');
        }
      }
    }
  }catch(e){
    Browser.msgBox( "保健シート内「X27」付近でエラーがあります。" );
    return;
  }
}

function alart_ok(){
  Browser.msgBox( "許可完了" );
}

function Refresh(){
  SpreadsheetApp.flush();   //シート再計算
}
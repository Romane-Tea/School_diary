//時間ベースで自動実行する
function auto() {
  syukei_delete();          //欠席集計の計算式を削除
  SpreadsheetApp.flush();   //シート再計算
  kesseki_record();         //欠席入力情報を記録する
  kesseki_record_backup();  //欠席入力情報をbackupシートに記録する
  Nisshi_record();          //各種日誌情報を記録する
  Nisshi_record_backup();   //各種日誌情報をbackupシートに記録する
  save_pdf_auto();          //各種日誌を保存していなければ自動保存する
  Kesseki_clear();          //欠席入力をすべて消す
  
  // Nisshi_clear();           //各種日誌情報を全て消す
  // hoken_font_size();        //保健日誌のフォントサイズを標準に戻す
  all_restore();            //日誌関係をすべて初期化
};

//各種日誌をすべてPDFで保存する
function save_all(){
  let mailAddress = Session.getActiveUser().getEmail();
  console.log('実行者：'+mailAddress);

  const ss = SpreadsheetApp.getActiveSpreadsheet(); //スプレッドシート
  ssId = ss.getId(); // スプレッドシートIDを取得
  SpreadsheetApp.flush();
  kesseki_record();
  Nisshi_record();
  const sheet1 = ss.getSheetByName('学校');
  const sheet2 = ss.getSheetByName('保健');
  const sheet3 = ss.getSheetByName('給食');
  const sheet4 = ss.getSheetByName('入力');
  save_pdf(ssId, sheet1, "all");
  save_pdf(ssId, sheet2, "all");
  save_pdf(ssId, sheet3, "all");
  sheet4.activate();
};

//自動実行時にPDFを条件に応じて保存する
function save_pdf_auto(){
  const ss = SpreadsheetApp.getActiveSpreadsheet(); //スプレッドシート
  ssId = ss.getId(); // スプレッドシートIDを取得
  const sheet1 = ss.getSheetByName('学校');
  const sheet2 = ss.getSheetByName('保健');
  const sheet3 = ss.getSheetByName('給食');
  const hoken_write = ss.getSheetByName('入力').getRange('G6').getValue();    //学校日を取得
  const kyusyoku_write = ss.getSheetByName('入力').getRange('AJ6').getValue();//給食日を取得
  save_pdf(ssId, sheet1, "auto");
  if (hoken_write == '〇'){ save_pdf(ssId, sheet2, "auto");}    //学校日なら保健日誌をPDF保存
  if (kyusyoku_write == '〇'){ save_pdf(ssId, sheet3, "auto");} //給食日なら給食日誌をPDF保存
};

//保健日誌の備考のフォントサイズを元に戻す
function hoken_font_size() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); //スプレッドシート
  const sheet = ss.getSheetByName('保健');
  sheet.getRange('M11:V28').activate();
  sheet.getActiveRangeList().setFontSize(10);
};

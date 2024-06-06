// グローバル変数としてキャッシュを定義する
const user_data = CacheService.getScriptCache();
const cacheKey = "userInfo"; // キャッシュのキー
const listSheet_data = CacheService.getScriptCache();
const cacheKey2 = "listSheetInfo"; // キャッシュのキー

function doGet(e) {
  //デプロイアドレスを取得
  var deploy = ScriptApp.getService().getUrl();

  // キャッシュから基本情報を取得
  var userInfo = user_data.get(cacheKey);
  if (!userInfo) {
    // キャッシュが存在しない場合はスプレッドシートから読み込む
    console.log("キャッシュ作成");
    userInfo = access_ok();
    user_data.put(cacheKey, JSON.stringify(userInfo)); // 基本情報をキャッシュに保存
  } else {
    userInfo = JSON.parse(userInfo); // キャッシュから取得した文字列を多次元配列に戻す
    console.log("キャッシュ読み込み");
  }

  //ファイルにアクセスしているGoogleユーザーのメールアドレスを取得する
  var activeUser = Session.getActiveUser().getEmail();
  console.log(activeUser);

  // 下の8行のコメントアウトを外すと、メールアドレス登録者のみ実行可能になる
  // //GAS実行を許可するGoogleユーザーのメールアドレスの一覧をTOPから取得
  // var acceptedUsers = userInfo();
  
  // //許可されていない場合、GASを終了する
  // if(acceptedUsers.indexOf(activeUser) == -1) {
  //   Browser.msgBox("あなたには実行権限がありません。");
  //   return;
  // }

  let page = e.parameter.p;

  if (!page) {
    page = 'page';
  }
  const template = HtmlService.createTemplateFromFile(page)
  template.btn= deploy;     //デプロイアドレスを付加
  
  if (page != 'index') {
    for (const key in e.parameter) {
      template[key] = e.parameter[key];
    }
  }
  return template
    .evaluate()
    .setTitle("出欠管理アプリ")
    .addMetaTag('viewport', 'width=device-width,initial-scale=1');
    // .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 許可ユーザー一覧を取得
function access_ok() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('リスト');
  // 最終行を取得する
  var last_row = sheet.getLastRow();     // 最終行取得
  var acceptedUsers = [];

  // 一度にデータを取得
  var range = sheet.getRange(3, 3, last_row - 2, 1);
  var values = range.getValues();

  for (var n = 0; n < values.length; n++) {
    acceptedUsers.push(values[n][0]);
  }
  console.log("access_ok=" + acceptedUsers);
  return acceptedUsers;
}


//シートネームを呼び出す関数
function SHEET_NAME() {
   return SpreadsheetApp.getActiveSheet().getName(); 
}

//シートナンバーからシートネームを取り出す関数
function GET_SHEET_NAME(sheet_no) {
  try{
    var sheet_name = SpreadsheetApp.getActive().getSheets()[sheet_no - 1].getSheetName();
  }catch(e){}
  return sheet_name
}

// シートIDを割り出す関数
function sheetid() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetId();
}

//欠席入力シートに、入力のあった行に自動で時刻を入れる関数
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== "欠席入力" || e.range.columnStart !== 6) return; // シート名とF列に値が入力されたかを確認
  var row = e.range.getRow();
  var value = e.range.getValue();
  if (value === "") { // H列から値が削除された場合、I列から時刻を削除
    sheet.getRange(row, 9).clearContent();
  } else { // F列に値が入力された場合、I列に現在時刻を8:10形式で自動入力
    var time = Utilities.formatDate(new Date(), "JST", "H:mm");
    sheet.getRange(row, 9).setValue(time);
  }
  // セルが正しく更新されるようにダミー パラメーターを追加
  // SpreadsheetApp.getActive().getSheetByName("dummy").getRange("A1").setValue(new Date());
}

//************  以下html用の関数 ****************//

// スプレッドシートのデータをシート名で読み込む
function GetSpreadsheet(sheet_name) {
  console.log(sheet_name);
  
  // キャッシュを利用するかどうかのフラグ
  var useCache = (sheet_name === 'リスト');
  
  // キャッシュを利用する場合
  if (useCache) {
    // キャッシュからデータを読み込む
    var cachedData = listSheet_data.get(cacheKey2);
    if (cachedData != null) {
      console.log("リストキャッシュ使用");
      return JSON.parse(cachedData);
    }
  }
  
  // 操作するシート名を指定して開く
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  
  // 全データを取得するので、最終列と最終行を取得する
  var last_col = sheet.getLastColumn();  // 最終列取得
  var last_row = sheet.getLastRow();     // 最終行取得
  var data = sheet.getRange(1, 1, last_row, last_col).getDisplayValues();
  
  // キャッシュを利用する場合はキャッシュに保存
  if (useCache) {
    listSheet_data.put(cacheKey2, JSON.stringify(data));
    console.log("リストキャッシュ保存");
  }
  return data;
}


// スプレッドシートのデータをシート名で読み込む
function GetSpreadsheet_TEST(sheet_name) {
  console.time("処理Aの時間計測");
  sheet_name="欠席記録";
  
  // 操作するシート名を指定して開く
  // var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).;
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getDataRange().getValues();
  // 全データを取得するので、最終列と最終行を取得する
  // var last_col = sheet.getLastColumn();  // 最終列取得
  // var last_row = sheet.getLastRow();     // 最終行取得
  // var data = sheet.getRange(1, 1, last_row, last_col).getDisplayValues();
  // console.log(last_row,last_col,data[0][0]);
  console.log(data[0].length);
  console.timeEnd("処理Aの時間計測");

  //   console.time("処理Aの時間計測");
  // sheet_name="欠席記録";
  
  // // 操作するシート名を指定して開く
  // var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  
  // // 全データを取得するので、最終列と最終行を取得する
  // var last_col = sheet.getLastColumn();  // 最終列取得
  // var last_row = sheet.getLastRow();     // 最終行取得
  // const range = sheet.getRange(1, 1, last_row, last_col);
  // const values = range.getDisplayValues();
  // var newValues = [];
  // for (let i = 0; i < values.length; i++) {
  //   newValues[i] = [];
  //   for (let j = 0; j < values[i].length; j++) {
  //     newValues[i][j] = values[i][j] + 1;
  //   }
  // }

  // var data = sheet.getRange(1, 1, newValues.length, newValues[0].length).getDisplayValues();
  // console.log(data[0][0]);
  // console.timeEnd("処理Aの時間計測");
}
//スプレッドシートにシート名でデータを書き込む
function SetSpreadSheet(sheet_name,text,row,column){
  //ファイルをIDで指定、シートをシート名で指定
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  
  //データを書き;込む範囲を指定して、書き込みます
  console.log('sheet_name='+ sheet_name +" : "+ row+" "+column);
  sheet.getRange(Number(row)+1, Number(column)+1).setValue(text);
}


////////////////////////////////////////////////////////////////////////////////////
//データが日付かどうかを判定する　日付ならtrue
function isDate(d) {
  if ( Object.prototype.toString.call(d) == "[object Date]" ){
    return true;
  } 
  return false;
}

// 日付オブジェクトデータを文字列にして返す
function date_to_value(data){
  for(var r = 0; r < data.length; r++) {
    for(var c = 0; c < data[r].length; c++) {
    if (isDate(data[r][c])){
      data[r][c] = Utilities.formatDate(data[r][c], "JST", "hh:mm:ss");
    }else{
      data[r][c] = data[r][c]
      }
    }
  }
  return data;
}
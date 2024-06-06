// pdf作成条件判定関数
function save_pdf(ssId,sheet,auto) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); //スプレッドシート

  if (!sheet){      //シート上のマクロ起動時の設定
    var ssId = ss.getId(); // スプレッドシートIDを取得
    // var sheet = ss.getSheetByName('日誌入力');
    var sheet = ss.getActiveSheet(); //スプレッドシート内のターゲットシート
    var dialog = true
  }
  var sheet_name = sheet.getRange("A1").getValue();
  var shId = sheet.getSheetId(); //ターゲットシートのID
  var parentFolder = DriveApp.getFileById(ssId).getParents(); // IDからスプレッドシートのファイルを取得⇒親フォルダを取得
  var folderId = parentFolder.next().getId(); // 親フォルダIDを取得

  // PDFファイル存在時の動作確認用
  var start_sheet = ss.getSheetByName('リスト');
  var pdfBool = start_sheet.getRange("AM22").getValue();    //上書き保存の設定値
  var auto_pdfBool = start_sheet.getRange("AM25").getValue(); //自動実行時の上書き保存設定値
  if (auto == 'all'){
    dialog == false;
  } else if (auto == 'auto'){
    dialog == false;
    pdfBool = auto_pdfBool; // auto実行時に、Bool値を書き換える
  }

  //シート名で印刷範囲を決める
  if (sheet_name =='学校'){
    var today = sheet.getRange("Y3").getValue();
  }else if (sheet_name =='保健'){
    var today = sheet.getRange("AA3").getValue();
    // 体裁を整える
    Re_arrange(sheet);
  } else if (sheet_name =='給食'){
    var today = sheet.getRange("AF3").getValue();
  } else {  // 完了メッセージポップアップ
    SpreadsheetApp.getActiveSpreadsheet().toast('不明なシートか、A1セルの値が消えています。', 'エラー', 5);
    return;
  }
  //日付データを取得,加工
  // var today = new Date( today );
  var fileName = 'R' + (today.getFullYear() - 2018) + '.' + (today.getMonth() + 1) + '.' + (today.getDate()) + sheet_name ;

  // PDFフォルダの有無を確認
  const targetFolderName = "PDF";
  const folderIterator = DriveApp.getFolderById(folderId).getFoldersByName(targetFolderName);  // 指定した名前のフォルダを取得
  
  let targetFolder;
  if (folderIterator.hasNext()) {
    // 存在する場合
    targetFolder = folderIterator.next();
    console.log("PDFフォルダが既にあります"+'('+sheet_name+')');
    } else {
    // 存在しない場合
    targetFolder = DriveApp.getFolderById(folderId).createFolder(targetFolderName);
    console.log("PDFフォルダを作成");
  }
  targetFolderId = targetFolder.getId()
  console.log("PDFフォルダのID=" + targetFolderId)

  // PDF/BackupPDFフォルダの有無を確認
  if (pdfBool == 0 ){
    const targetFolderName2 = "BackupPDF";
    const folderIterator2 = DriveApp.getFolderById(targetFolderId).getFoldersByName(targetFolderName2);  // 指定した名前のフォルダを取得
    var targetFolder2;
    if (folderIterator2.hasNext()) {
      // 存在する場合
      targetFolder2 = folderIterator2.next();
      console.log("BackupPDFフォルダが既にあります");
      } else {
      // 存在しない場合
      targetFolder2 = DriveApp.getFolderById(targetFolderId).createFolder(targetFolderName2);
      console.log("BackupPDFフォルダを作成");
    }
    targetFolderId2 = targetFolder2.getId()
    console.log("BackupPDFフォルダのID=" + targetFolderId2 )
  }else {
    var targetFolder2 ="";
    var targetFolderId2 ="";
  }
  
  // 作成するPDFファイルの存在の有無　　IDを取得
  const fileNamePDF = fileName + '.pdf';
  var files = DriveApp.getFolderById(targetFolderId).getFiles();
  while (files.hasNext()) {
    var file = files.next();
    var tempfileName = file.getName() ;
    if(tempfileName == fileNamePDF){
      break;
      }
  }
  if (!file || file.getName() != fileNamePDF){    //ファイルが存在しない、又はファイルネームが違うとき
    var pdfId = ""
    Logger.log("既存のPDFはありません");
  }else if (file.getName() == fileNamePDF){
    var pdfId = file.getId()
    Logger.log("既存のPDF_ID="+pdfId);
    if (auto == "auto" && auto_pdfBool == 2) {
      return;
      }
  }
  SpreadsheetApp.flush();
  //関数createPdfを実行し、PDFを作成して保存する
  createPdf(targetFolder, targetFolderId, targetFolder2, targetFolderId2, pdfBool, ssId, shId, sheet_name, fileName , pdfId, dialog);
}

//必要事項が入力してるかチェック ※天気、未定者
function input_check(){
  var ss = SpreadsheetApp.getActiveSpreadsheet(); //スプレッドシート
  var input_sheet = ss.getSheetByName('入力');
  var weather = input_sheet.getRange("G3").getValue(); //天気が入っている
  var mitei = input_sheet.getRange("AC10").getValue(); //未定者がいないか
  if (weather =="" && mitei !=0 ){
    Browser.msgBox("天気が入力されていません。未定者がいます。入力後、もう一度実行してください。")
  } else if (weather =="" ){
    Browser.msgBox("天気が入力されていません。入力後、もう一度実行してください。")
  } else if (mitei !=0 ){
    Browser.msgBox("未定者がいます。入力後、もう一度実行してください。")
  }
}

//必要事項が入力してるかチェック ※給食関係
function input_check_kyusyoku(){
  var ss = SpreadsheetApp.getActiveSpreadsheet(); //スプレッドシート
  var input_sheet = ss.getSheetByName('入力');
  var pre_ab = input_sheet.getRange("AJ6");  //給食日かどうか
  var menu = input_sheet.getRange("AF12").getValue(); //メニューが入っている
  var temperature = input_sheet.getRange("AC10").getValue(); //温度が入っているか
  if (pre_ab == ""){
    return;
  } else if (menu =="" && temperature == "" ){
    Browser.msgBox("メニュー、配膳室の温度等が入力されていません。入力後、もう一度実行してください。")
  } else if (menu =="" ){
    Browser.msgBox("メニューが入力されていません。入力後、もう一度実行してください。")
  } else if (temperature = "" ){
    Browser.msgBox("配膳室の温度等が未入力がいます。入力後、もう一度実行してください。")
  }
}

//個別PDF保存関数　学校、保健、給食
function save_gakko(){
  let mailAddress = Session.getActiveUser().getEmail();
  console.log('実行者：'+mailAddress);

  var ss = SpreadsheetApp.getActiveSpreadsheet(); //スプレッドシート
  ssId = ss.getId(); // スプレッドシートIDを取得
  SpreadsheetApp.flush();
  kesseki_record();
  Nisshi_record();
  var sheet1 = ss.getSheetByName('学校');
  var sheet4 = ss.getSheetByName('入力');
  sheet1.activate();
  save_pdf();
};

function save_hoken(){
  let mailAddress = Session.getActiveUser().getEmail();
  console.log('実行者：'+mailAddress);

  var ss = SpreadsheetApp.getActiveSpreadsheet(); //スプレッドシート
  ssId = ss.getId(); // スプレッドシートIDを取得
  SpreadsheetApp.flush();
  kesseki_record();
  Nisshi_record();
  var sheet2 = ss.getSheetByName('保健');
  var sheet4 = ss.getSheetByName('入力');
  sheet2.activate();
  save_pdf();
};

function save_kyusyoku(){
  let mailAddress = Session.getActiveUser().getEmail();
  console.log('実行者：'+mailAddress);

  var ss = SpreadsheetApp.getActiveSpreadsheet(); //スプレッドシート
  ssId = ss.getId(); // スプレッドシートIDを取得
  SpreadsheetApp.flush();
  kesseki_record();
  Nisshi_record();
  var sheet3 = ss.getSheetByName('給食');
  var sheet4 = ss.getSheetByName('入力');
  sheet3.activate();
  save_pdf();
};

//PDF作成関数 
function createPdf(targetFolder, targetFolderId, targetFolder2, targetFolderId2, pdfBool, ssId, shId, sheet_name, fileName, pdfId, dialog){
  //シート名で印刷範囲を決める
  if (sheet_name =='学校'){
    var print_range = "&range=B2:V38";
  }else if (sheet_name =='保健'){
    var print_range = "&range=B2:V63"; 
  } else if (sheet_name =='給食'){
    var print_range = "&range=B2:AC38"; 
  }
  let baseUrl = "https://docs.google.com/spreadsheets/d/"   //PDFを作成するためのベースとなるURL
          +  ssId
          + "/export?&gid="
          + shId

  let pdfOptions = "&exportFormat=pdf&format=pdf"   //PDFのオプションを指定
              + "&size=A4" //用紙サイズ (A4)
              + "&portrait=true"  //用紙の向き true: 縦向き / false: 横向き
              + "&gridlines=false" //グリッドラインの表示有無
              + "&fitw=true"  //ページ幅を用紙にフィットさせるか true: フィットさせる / false: 原寸大
              + "&top_margin=0.748" //上の余白
              + "&right_margin=0.50" //右の余白
              + "&bottom_margin=0.748" //下の余白
              + "&left_margin=0.90" //左の余白
              + print_range ;//"&range=B4%3AR65"  //セル範囲を指定 %3A はコロン(:)を表す

              // + "&printtitle=false" //スプレッドシート名の表示有無
              // + "&sheetnames=false" //シート名の表示有無
              // + "&horizontal_alignment=CENTER" //水平方向の位置
              // + "&vertical_alignment=TOP" //垂直方向の位置
              // + "&fzr=false" //固定行の表示有無
              // + "&fzc=false" //固定列の表示有無;

  let url = baseUrl + pdfOptions;    //PDFを作成するためのURL
  //アクセストークンを取得する
  let token = ScriptApp.getOAuthToken();
  //headersにアクセストークンを格納する
  let options = {
    headers: {
        'Authorization': 'Bearer ' +  token
    }
  };

  //PDFを作成する
  let blob = UrlFetchApp.fetch(url, options).getBlob().setName(fileName + '.pdf');

  if (pdfId != ""){   // PDFが存在するなら
    if (pdfBool == 1){
      //ファイルの新バージョンをアップロードするコード
      let resource = {
        uploadType: "resumable",  // "media"でも可
      }
      Drive.Files.update(resource, pdfId, blob);   // 拡張サービス[Drive]使用　https://www.ka-net.org/blog/?p=4372
    }else if (pdfBool == 0){
      //Backupフォルダに保存し、新しく作成
      let file = DriveApp.getFileById(pdfId);   //移動するpdfファイル
      // let BackupFolder = DriveApp.getFolderById(targetFolderId2);　//Backupフォルダ
      file.moveTo(targetFolder2);
      targetFolder.createFile(blob);  // ファイルをフォルダ内に作成
    }
  }else{          // PDFが存在しないなら
    targetFolder.createFile(blob);  // ファイルをフォルダ内に作成
  }
  
  // 作成したPDFファイルのIDを取得
  const fileNamePDF = fileName + '.pdf';
  var files = DriveApp.getFolderById(targetFolderId).getFiles();
  while (files.hasNext()) {
    var file = files.next();
    if(file.getName() == fileNamePDF){break;}
  }
  var pdfId = file.getId()
  Logger.log(file.getId());
  make_link(sheet_name,pdfId);
  
  if (dialog == true){
    pdf_link(pdfId)
  }
}

//入力タブ内にリンクを作成する
function make_link(sheet_name,pdfId){
  let ss = SpreadsheetApp.getActiveSpreadsheet(); //スプレッドシート
  let input_sheet = ss.getSheetByName("入力");
  let link = "https://drive.google.com/file/d/" + pdfId;
  if (sheet_name == "学校"){
    input_sheet.getRange("C51").setValue(link);
    Logger.log("pdf作成　sheet_name学校="+sheet_name);
  } else if (sheet_name == "保健"){
    input_sheet.getRange("D51").setValue(link);
    Logger.log("pdf作成　sheet_name保健="+sheet_name);
  } else if (sheet_name == "給食"){
    input_sheet.getRange("E51").setValue(link);
    Logger.log("pdf作成　sheet_name給食="+sheet_name);
  }
}

function pdf_link(pdfId){
  var link = "https://drive.google.com/file/d/"
  var linkId = pdfId
  var htmlOutput = HtmlService
      .createHtmlOutput('<p><a href=' +link + linkId + ' target="blank">PDFを表示する</a></p>')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(350)
      .setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '作成したPDFを表示します。');
}

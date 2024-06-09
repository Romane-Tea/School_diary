// 置換用連想配列
var byouketu_label_list = { '発熱': '発熱', '頭痛': '頭痛', 'かぜ症状': 'かぜ症状', '下痢・腹痛': '下痢・腹痛', '嘔気・嘔吐': '嘔気・嘔吐', '発疹': '発疹', 'インフルエンザ様症状': 'インフルエンザ' };
// それぞれの配列
var kibiki_list = ["祖父死去", "祖母死去", "曾祖父死去", "曾祖母死去", "葬儀"];
var syuttei_list = ["感染症", "インフルエンザ", "帯状疱疹", "学級閉鎖"];
var byouketu_list = ["発熱","頭痛","かぜ症状","下痢・腹痛","吐嘔気・嘔吐","発疹","咽頭痛","咳","家事都合","体調不良","けが","喘息"]
// 6:欠席、遅刻、早退、遅刻・早退
// 9:病気、通院、家庭の都合、忌引き、体調不良、その他
// 9:病気理由:発熱、頭痛、かぜ症状、下痢・腹痛、嘔気・嘔吐、発疹、インフルエンザ様症状、その他


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('tetoru読み込み')
    .addItem('Select CSV File', 'importCSV')
    .addToUi();
}

function importCSV() {
  main();
}

function main() {
  var html = HtmlService.createHtmlOutputFromFile("dialog");
  SpreadsheetApp.getUi().showModalDialog(html, 'ローカルファイル読込');
}

function writeSheet(formObject) {
  // フォームで指定したテキストファイルを読み込む
  var fileBlob = formObject.myFile;

  // テキストとして取得（Windowsの場合、文字コードに UTF-8 を指定）
  let text = fileBlob.getDataAsString("UTF-8");

  // 改行文字を正規表現で置換して、一時的な文字列に変換
  let tempNewline = "@NEWLINE@";
  let csvData = text.replace(/(\r\n|\r|\n)(?=(?:(?:[^"]*"){2})*[^"]*$)/g, tempNewline);

  // 一時的な改行文字で分割し、それぞれの行をコンマで分割して配列に格納する
  csvData = csvData.split(tempNewline).map(function(row) {
    return row.split(',');
  });
  csvData[csvData.length-1] =[,,,,,,,,,,,,,,,,,,,];

  // 書き込むシートを取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const t_sheet = ss.getSheetByName("tetoruデータ");
  t_sheet.clear();

  // テキストファイルをシートに展開する
  try {
    t_sheet.getRange(1, 1,csvData.length, csvData[0].length ).setValues(csvData);
  }finally {
    // 処理終了したら、データ処理して欠席入力に記録
    let new_data = extraction(csvData);
    write_kesseki(new_data);
  }
}

function write_kesseki(new_data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("欠席入力");

  // シートから番号の配列を取得
  let s_data = sheet.getRange("A:F").getValues();
  let numbers = s_data.map((subArray) => subArray[0]);
  // 配列の数だけ処理を行う
  let not_write = 0;  //すでに書き込んであるデータ数の変数
  for (let i = 0; i < new_data.length; i++) {
    let number = new_data[i][1];
    let row = numbers.indexOf(number) + 1; // 行番号

    if (row > 0) { // 番号が見つかった場合
      console.log("欠席データ"+row,s_data[row-1][5]);

      // F列、G列、H列にデータを書き込み
      sheet.getRange(row, 6, 1, 3).setValues([new_data[i].slice(3, 6)]);
      // セルの色を変更
      sheet.getRange(row, 6).setBackground("yellow");

      // if (s_data[row-1][5] ==""){
      //   // F列、G列、H列にデータを書き込み
      //   sheet.getRange(row, 6, 1, 3).setValues([new_data[i].slice(3, 6)]);
      //   // セルの色を変更
      //   sheet.getRange(row, 6).setBackground("yellow");
      // } else { not_write++}
    } else { not_write++} //番号が見つからなかった場合
  }
  // 処理終了のメッセージボックスを出力
  let msg = (not_write === 0) ? i + "件のデータが該当しました。\\n(取り込んだデータは黄色く表示されています。)" : i + "件のデータが該当しました。\\n" + not_write + "件のデータが該当生徒なし(要確認)。\\n(取り込んだデータは黄色く表示されています。)";
  // let msg = (not_write === 0) ? i + "件のデータが該当しました。\\n(取り込んだデータは黄色く表示されています。)" : i + "件のデータが該当しました。\\n" + not_write + "件のデータが既に入力済み、もしくは該当生徒なし。\\n(取り込んだデータは黄色く表示されています。)";
  Browser.msgBox(msg);
}

//csvdataを欠席入力用に加工
function extraction(csvData) {
  new_data = [];  // 空の配列
  let now_day = new Date();
  let day = now_day = new Date().toLocaleDateString("ja-JP", {year: "numeric",month: "2-digit",day: "2-digit"});	// 現在の日付をISO形式で取得し、文字列として取得

  for (let i = 1; i < csvData.length - 1; i++) {
    let temp_data = csvData[i];
    //日付を検査
    if (temp_data[0].split("(")[0] === String(day)) {
      let number = "";
      //出席番号を取得
      if (temp_data.indexOf("特別支援学級") !== -1) {
        let toku_num = temp_data.indexOf("特別支援学級");
        let class_name = temp_data[toku_num + 1].replace('学級','');
        number = temp_data[1][0] + "-" + class_name;
      } else if (temp_data.indexOf("教職員") !== -1) {
        continue;
      } else {
        number = Number(temp_data[1][0])*1000 + Number(temp_data[2][2])*100 + Number(temp_data[3]);
      }
      //欠席理由を取得
      let kesseki_kind = temp_data[6].replace(/"/g, '');  // 連絡種別  replaceは"削除
      let reason = temp_data[9].replace(/"/g, '');    // 理由
      let biko = temp_data[10].replace(/"/g, '');   // 備考
      let kesseki_result = "";    //確定した欠席種別用空の変数
      let reason_result = "";     //確定した理由用空の変数
      let biko_result = "";        //確定したその他用空の変数

      //理由の内容の処理
      if (reason.includes("病")) {    //理由内容が病気の時の処理
        kesseki_result = kesseki_kind;
        reason_result = reason.split("（")[1].slice(0, -1);  //理由内容の（）内の文字を取得
        reason_result = searchByKey(byouketu_label_list, reason_result) || reason_result; // 理由を置換
        biko_result = biko;

      } else if (reason.includes("忌")) {  //忌引きの時の処理
        kesseki_result = "忌引き";
        reason_result = kibiki_list.includes(biko) ? biko : "その他"; //備考に該当する理由があるとき、それを理由に、なければその他にする
        biko_result = kibiki_list.includes(biko) ? "" : biko; //該当時空白にする

      //その他の内容処理
      } else {
        kesseki_result = kesseki_kind;
        if (searchByKey(syuttei_list, biko) != null) {  //出停リスト該当時
          kesseki_result = "出停";
          reason_result = biko;

        } else if (searchByKey(byouketu_label_list, biko) != null){ //病欠ラベルリスト該当時
          let value = searchByKey(byouketu_label_list, biko);
          reason_result = value == null ? "その他" : value;
          biko_result = value == null ? biko : "";

        } else if(byouketu_list.includes(biko)){  //病欠リスト該当時
          reason_result = biko;
          biko_result ="";

        } else {
          reason_result = reason.split("\r\n")[1];
          biko_result = biko ;
        }
      }

      if (temp_data[7] != "") biko_result += `(${temp_data[7]}予定)`;
      if (temp_data[8] != "") biko_result += `(${temp_data[8]}予定)`;

      new_data.push([day,number,temp_data[4],kesseki_result, reason_result, biko_result]);
    }
  }
  return new_data;
}

function searchByKey(obj, key) {
  if (key in obj) {
    return obj[key];
  } else {
    return null;
  }
}

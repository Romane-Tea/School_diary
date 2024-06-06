<script>
// グローバル関数群
grade = 0;   //学年
tclass = 0;   //学級
start_row = 9;  //生徒読み込み行（1行目が０)
data1 = [];    //欠席入力シート全体のデータ
data_backup = [];  //読み込んだデータのバックアップ用
reason_list = []; //事由、理由一覧
top_data = []; //リストシートのデータ
hyouziUI = 1; // 1:表形式　2:座席表形式
kind = "grade"; //grade:クラスごと　all_gr:全校　ziyu:欠席等入力あり
yoko = 6;
seki_No = 1;  //席順  1:座席順　2:出席番号順
hyouzi = [1, 1, 1, 1, 1, 1, 1, 1]; //[欠席,出停,忌引,遅刻,早退,転出,未定] 7つ

//最初に実行
document.addEventListener('DOMContentLoaded', function () {
  start_html();
}, false);

async function start_html() {
  data1 = await get_sheet_data("欠席入力");  //欠席入力シート全体のデータを取得
  data_backup = JSON.parse(JSON.stringify(data1));  //読み込んだデータのバックアップを保存
  top_data = await get_sheet_data("リスト");  //リストシート全体のデータを取得
  hyouziUI = top_data[11][38];     //表示初期値を取得

  // for (var i = 0; i<4 ; i++){  //34列行目から37行目 -1
  //   const reason_item =[];
  //   for (var j=0 ; j<32 ;j++){  //34列目から65列目 -1
  //     if (top_data[i+33][j+33] ==""){break;}else{
  //     reason_item[j] = top_data[i+33][j+33];
  //     }
  //   }
  //   console.log(reason_item);
  //   reason_list[i] = reason_item;
  // }
  
  for (var i = 0; i<4 ; i++){  //33列行目から36列目 -1
    const reason_item =[];
    for (var j=0 ; j<32 ;j++){  //3行目から33列目 -1
      if (top_data[j+1][i+32] ==""){break;}else{
      reason_item[j] = top_data[j+1][i+32];
      }
    }
    reason_list[i] = reason_item;
    console.log("i="+i,reason_item);
  }
  // 表示設定を読み込む
  for (var i = 0; i < 8; i++) {
    hyouzi[i] = top_data[i + 13][38];
  }
  // // 事由リストを作成
  // make_list();
  // トップバー作成
  make_first(hyouziUI);
  // メインの表を作成
  make_table(hyouziUI);
  // クラス一覧を取得、選択リストに反映
  class_list()
}

// シート名からデータを読み込む関数（コードスクリプトの読み込み）
function get_sheet_data(sheet_name) {
  return new Promise((resolve, reject) => {
    google.script.run
      .withSuccessHandler((result) => resolve(result))
      .GetSpreadsheet(sheet_name);
  });
}

// トップバーを作成
function make_first(hyouziUI) {
  var first_bar = document.getElementById('first_bar');
  var child_n = first_bar.childElementCount;
  //ボタン等を一旦削除
  for (var i = 0; i < child_n - 2; i++) {
    first_bar.removeChild(first_bar.lastChild);
  }

  if (hyouziUI == 1) {  //一覧表示時のトップバーの設定
    var td1 = document.createElement('td');
    var td2 = document.createElement('td');
    var td3 = document.createElement('td');
    var td4 = document.createElement('td');
    var td5 = document.createElement('td');
    var btn = document.createElement('button');
    var input1 = document.createElement('input');
    var input2 = document.createElement('input');
    var input3 = document.createElement('input');

    td1.setAttribute('onclick', "getElementById('hurigana').click();");
    td1.setAttribute('class','btn btn_check');
    td1.textContent ='ふりがな';
    input1.setAttribute('type', 'checkbox');
    input1.setAttribute('id', 'hurigana');
    input1.setAttribute('onclick', "checkbox_cell(this,'col2')");
    input1.setAttribute('name', 'ふりがな');
    td1.appendChild(input1);

    td2.setAttribute('onclick', "getElementById('birthday').click();");
    td2.setAttribute('class','btn btn_check');
    td2.innerText = '生年月日';
    input2.setAttribute('type', 'checkbox');
    input2.setAttribute('id', 'birthday');
    input2.setAttribute('onclick', "checkbox_cell(this,'col3')");
    input2.setAttribute('name', '生年月日');
    td2.appendChild(input2);

    td3.setAttribute('onclick', "getElementById('fm').click();");
    td3.setAttribute('class','btn btn_check');
    td3.innerText = '性別';
    input3.setAttribute('type', 'checkbox');
    input3.setAttribute('id', 'fm');
    input3.setAttribute('onclick', "checkbox_cell(this,'col4')");
    input3.setAttribute('name', '性別');
    td3.appendChild(input3);
    
    btn.setAttribute('class', "btn btn_change");
    btn.setAttribute('onclick', 'change_hyouziUI()');
    btn.innerText = '座席形式に';
    td4.appendChild(btn);

    td5.setAttribute('id', 'apli_name');
    td5.setAttribute('class', 'dai');
    td5.innerText = '出欠管理アプリ 表形式 ' + top_data[9][38];

    first_bar.appendChild(td1);
    first_bar.appendChild(td2);
    first_bar.appendChild(td3);
    first_bar.appendChild(td4);
    first_bar.appendChild(td5);

    // checkboxの初期値を設定
    document.getElementById("hurigana").checked = true;
    document.getElementById("birthday").checked = true;
    document.getElementById("fm").checked = true;

  } else if (hyouziUI == 2) {   //座席表示時のトップバーの設定
    var td1 = document.createElement('td');
    var td2 = document.createElement('td');
    var btn1 = document.createElement('button');
    var btn2 = document.createElement('button');

    btn1.setAttribute('id', 'chage_seki');
    btn1.setAttribute('class', 'btn btn_refresh');
    btn1.setAttribute('onclick', 'chage_seki()');
    btn1.innerText = '出席番号順に';

    btn2.setAttribute('id', 'chage_seki');
    btn2.setAttribute('class', 'btn btn_change');
    btn2.setAttribute('onclick', 'change_hyouziUI()');
    btn2.innerText = '表形式に';

    td1.appendChild(btn1);
    td1.appendChild(btn2);

    td2.setAttribute('id', 'apli_name');
    td2.setAttribute('class', 'dai');
    td2.innerText = '出欠管理アプリ 座席形式 ' + top_data[9][38];

    first_bar.appendChild(td1);
    first_bar.appendChild(td2);
  }
}

//表示形式の切り替えボタン関数
function change_hyouziUI() {
  var scrollEle = document.getElementById('scroll_table');
  var tableEle1 = document.getElementById('data-table');
  var tableEle2 = document.getElementById('data-table2');
  if (hyouziUI == 1) { 
    tableEle1.innerText ="";
    scrollEle.style.height = '0vh';
    tableEle2.style.height = '85vh';
    hyouziUI = 2;
    // tableEle.setAttribute('class','data-table2');
    }
  else if (hyouziUI == 2) {
    tableEle2.innerText ="";
    scrollEle.style.height = '85vh';
    tableEle2.style.height = '0vh';
    hyouziUI = 1;
    // tableEle.setAttribute('class','data-table');
    }
  console.log("hyouziUI="+ hyouziUI);
  make_first(hyouziUI);
  // メインの表を作成
  make_table(hyouziUI);
}

//座席の表示順を交互に変更
function chage_seki(){
  var chage_seki_text = document.getElementById('chage_seki');
  if(seki_No==1){
    chage_seki_text.innerText = '座席順に'
    seki_No = 2;
  }else{
    chage_seki_text.innerText = '出席番号順に'
    seki_No = 1;
  }
  var tableEle = document.getElementById('data-table2');
  tableEle.innerText = "";
  make_table(hyouziUI);
}

//シートの学級一覧からリストとして返す関数
function class_list() {
  var obj = document.getElementById('select_class');
  while (obj.lastChild) {
    obj.removeChild(obj.lastChild);
  }
  let op = document.createElement("option");
  op.value = "";
  op.text = "学級を選択";
  obj.appendChild(op);

  for (var i = 1; i < 7; i++) {     //行番号3行目から8行目まで
    for (var j = 1; j < 9; j++) {   //38列目から45列目まで
      if ((top_data[i+1][j+38] != "") && (top_data[i+1][38] != "特支")) {
        let op = document.createElement("option");
        op.value = i + "_" + j;
        op.text = top_data[i+1][j+38];
        obj.appendChild(op);
      } else if ((top_data[i+1][j+38] != "") && (top_data[i+1][38] == "特支")) {
        let op = document.createElement("option");
        op.value = "特支_" + top_data[i+1][j+38];
        op.text = top_data[i+1][j+38];
        obj.appendChild(op);
      }
    }
  }
}

// 一覧表示分岐
function make_table(hyouziUI) {
  //表をタイトル以外削除
  var tableEle1 = document.getElementById('data-table');  //表のオブジェクトを取得
  tableEle1.innerText = "";                               //一旦一覧表を削除
  var tableEle2 = document.getElementById('data-table2');  //表のオブジェクトを取得
  tableEle2.innerText = "";                               //一旦座席表を削除
  if (hyouziUI == 1) { make_hyou_table(kind); }
  if (hyouziUI == 2) { make_zaseki_table(grade, tclass); }
}

//表テーブルを作成
function make_hyou_table(kind) {
  var tableEle = document.getElementById('data-table');
  var th = document.createElement('thead');
  var tr = document.createElement('tr');
  var th1 = document.createElement('th');
  var th2 = document.createElement('th');
  var th3 = document.createElement('th');
  var th4 = document.createElement('th');
  var th5 = document.createElement('th');
  var th6 = document.createElement('th');
  tr.setAttribute("class", "td_line");

  th1.innerText = "学籍番号";
  th2.innerText = "氏名";
  th3.innerText = "ふりがな";
  th4.innerText = "生年月日";
  th5.innerText = "性別";
  th6.innerText = "出欠等";

  th3.setAttribute("id", "col2");
  th4.setAttribute("id", "col3");
  th5.setAttribute("id", "col4");

  tr.appendChild(th1);
  tr.appendChild(th2);
  tr.appendChild(th3);
  tr.appendChild(th4);
  tr.appendChild(th5);
  tr.appendChild(th6);
  th.appendChild(tr);
  tableEle.appendChild(th);

  for (var i = start_row ; i < data1.length; i++) {
    id_num = data1[i][0];     // 4桁の番号を参照
    id_gr = Number(data1[i][0][0]);   // 4桁の千の位を取得
    id_cl = Number(data1[i][0][1]);   // 4桁の百の位を取得

    var hantei = Boolean("true");    //boolean型の設定
    if (kind == 'grade') {           //選択の種類が学級毎なら
      if (grade == '特支') {
        hantei = ((data1[i][0][1] == '-') );// && (data1[i][0].split('-')[1] == tclass));
      } else {
        hantei = ((id_gr == grade) && (id_cl == tclass));
      }
    } else if (kind == 'all_gr') {    //種類が全校なら
      hantei = (true && (data1[i][0] != ""));
    } else if (kind == 'ziyu') {      //種類が欠席事由なら
      hantei = (data1[i][5] != "");
    }

    if (hantei == true) {
      // テーブルの行を追加する
      var tr = document.createElement('tr');
      var td1 = document.createElement('td');
      var td2 = document.createElement('td');
      var td3 = document.createElement('td');
      var td4 = document.createElement('td');
      var td5 = document.createElement('td');
      var td6 = document.createElement('td');

      tr.setAttribute("class", "td_line");
      // td1 学籍番号
      td1.setAttribute("id", data1[i][0]);
      td1.innerText = data1[i][0];
      // td2 名前
      td2.innerText = data1[i][1];
      // td3 ふりがな
      td3.innerText = data1[i][2];
      // td4 生年月日
      td4.innerText = data1[i][3];
      // td5 性別
      td5.setAttribute("id", i);
      td5.innerText = data1[i][4];
      // td6 出欠
      if (data1[i][5] == "") {
        make_buttons(td6, i, 0);
      } else {
        var btn = document.createElement('button');
        td6.innerText = '[' + data1[i][8] + ']' + data1[i][5] + '　:　' + data1[i][6] + '　:　' + data1[i][7] + '　　';
        td6.setAttribute("id", i + '_syukketu');
        td6.setAttribute('align', 'left');
        td6.style.color = "red";
        td6.style.fontWeight = "bold";
        btn.setAttribute("id", i + '_clear');
        btn.setAttribute('class', 'kbtn');
        btn.innerText = 'クリア';
        btn.setAttribute('onclick', "clear_click(this)");
        td6.appendChild(btn);
      }
      tr.appendChild(td1);
      tr.appendChild(td2);
      tr.appendChild(td3);
      tr.appendChild(td4);
      tr.appendChild(td5);
      tr.appendChild(td6);
      tableEle.appendChild(tr);
    }
  }

}

//座席テーブルを作成
function make_zaseki_table(grade, tclass) {
  var tableEle = document.getElementById('data-table2');
  var bango = Number(grade) * 1000 + Number(tclass) * 100;

  // クラス人数を数える
  var ninzu = 0;
  for (var c = 1; c < data1.length; c++) {
    if (grade == '特支') {
      if (data1[c][0].substr(2) == tclass) {
        ninzu += 1;
      }
    } else if (Number(data1[c][0]) > bango && Number(data1[c][0]) < bango + 100) {
      ninzu += 1;
    }
  }

  // 縦の行数を決める
  if ((ninzu % 6) == 0) {
    tate = ninzu / 6;
  } else {
    tate = Math.floor(ninzu / 6) + 1;
  }

  ninzu = yoko * tate;

  var n = 1;        //通し番号
  for (var i = 1; i <= tate; i++) {
    // テーブルの行を追加する
    var tr = document.createElement('tr');
    for (var j = 0; j < yoko; j++) {
      if (seki_No == 1) {
        for (let s_row = start_row; s_row < data1.length; s_row++) {  //10行目からプログラム番号を検索する
          var td = document.createElement('td');
          var boolean_s = false;
          if (grade == '特支') {
            boolean_s = data1[s_row][0].substr(2) == tclass;
            start_row = s_row + 1;
          } else {
            boolean_s = (data1[s_row][9] == n && data1[s_row][0] > bango && data1[s_row][0] < bango + 100);
          }

          //プログラム番号が通し番号と同じ,かつ学級が条件にあえば
          if (boolean_s) {
            var text = data1[s_row][0] + " " + data1[s_row][1] + "  " + data1[s_row][4] + "\n" + data1[s_row][2];
            td.setAttribute("id", s_row + '_syukketu');
            td.setAttribute("class", "btn btn--orange");
            td.innerText = text + '\n';
            td.setAttribute("key", s_row)

            // td 出欠ボタン作成or欠席理由等を表示
            if (data1[s_row][5] == "") {
              make_buttons(td, s_row, 0);
            } else {
              var btn = document.createElement('button');
              td.innerText = text + '\n[' + data1[s_row][8] + ']' + data1[s_row][5] + ':' + data1[s_row][6] + ':' + data1[s_row][7] + '　';
              td.setAttribute("class", "btn btn--orange_black");
              td.style.fontWeight = "bold";
              btn.setAttribute("id", s_row + '_clear');
              btn.setAttribute('class', 'kbtn');
              btn.innerText = 'クリア';
              btn.setAttribute('onclick', "clear_click(this)");
              td.appendChild(btn);
            }
            tr.appendChild(td);
            break;

          } else if (s_row == data1.length - 1) {
            td.setAttribute("value", "kuseki");
            td.setAttribute("key", "0")
            td.innerText = "";
            tr.appendChild(td);
          }
        }
      } else if (seki_No == 2) {
        for (let s_row = start_row; s_row < data1.length; s_row++) {  //13行目からプログラム番号を検索する

          var td = document.createElement('td');
          var boolean_s = false;
          if (grade == '特支') {
            boolean_s = data1[s_row][0].substr(2) == tclass;
            start_row = s_row + 1;
          } else {
            boolean_s = (data1[s_row][0] == (bango + 1 + n));
          }

          //出席番号を通し番号順に
          if (boolean_s) {
            var text = data1[s_row][0] + " " + data1[s_row][1] + "  " + data1[s_row][4] + "\n" + data1[s_row][2];
            td.setAttribute("id", s_row + '_syukketu');
            td.setAttribute("class", "btn btn--orange");
            td.innerText = text + '\n';
            td.setAttribute("key", s_row)

            // td 出欠ボタン作成or欠席理由等を表示
            if (data1[s_row][5] == "") {
              make_buttons(td, s_row, 0);
            } else {
              var btn = document.createElement('button');
              td.innerText = text + '\n[' + data1[s_row][8] + ']' + data1[s_row][5] + ':' + data1[s_row][6] + ':' + data1[s_row][7] + '　';
              td.setAttribute("class", "btn btn--orange_black");
              td.style.fontWeight = "bold";
              btn.setAttribute("id", s_row + '_clear');
              btn.setAttribute('class', 'kbtn');
              btn.innerText = 'クリア';
              btn.setAttribute('onclick', "clear_click(this)");
              td.appendChild(btn);
            }
            tr.appendChild(td);
            break;

          } else if (s_row == data1.length - 1) {
            td.setAttribute("value", "kuseki");
            td.setAttribute("key", "0")
            td.innerText = "";
            tr.appendChild(td);
          }
        }
      }
      tableEle.appendChild(tr);
      n++;
      if (n == ninzu + 1) {
        break;
      }
    }
  }
}

//学級選択時の動作
function select_class() {
  var select_class = document.getElementById('select_class');
  var class_id = select_class.value;
  kind = "grade"; 
  grade = class_id.split('_')[0];
  tclass = class_id.split('_')[1];
  console.log("grade="+grade + " tclass="+tclass);
  make_table(hyouziUI);
}

//復元ボタンの動作
function rev_click(btn) {
  var btn_id = btn.id;
  var i = btn_id.split('_')[0];          //列番号を取得
  var td = document.getElementById(i + '_syukketu');
  data1[i][5] = data_backup[i][5];      //データに欠席等を復元
  data1[i][6] = data_backup[i][6];      //データに事由を復元
  data1[i][7] = data_backup[i][7];      //データにその他を復元
  data1[i][8] = data_backup[i][8];      //データにその他を復元
  google.script.run.SetSpreadSheet("欠席入力", data_backup[i][5], i, 5); //スプレットシートの理由を消去
  google.script.run.SetSpreadSheet("欠席入力", data_backup[i][6], i, 6); //スプレットシートの事由を消去
  google.script.run.SetSpreadSheet("欠席入力", data_backup[i][7], i, 7); //スプレットシートのその他を消去
  google.script.run.SetSpreadSheet("欠席入力", data_backup[i][8], i, 8); //スプレットシートの時間を消去
  //ボタン等を一旦削除
  while (td.lastChild) {
    td.removeChild(td.lastChild);
  }
  if (hyouziUI ==1 ){
    td.textContent = '[' + data_backup[i][8] + ']' + data_backup[i][5] + '　:　' + data_backup[i][6] + '　:　' + data_backup[i][7];
  } else {
    var text = data1[i][0] + " " + data1[i][1] + "  " + data1[i][4] + "\n" + data1[i][2];
    td.innerText = text + '\n[' + data_backup[i][8] + ']' + data_backup[i][5] + ':' + data_backup[i][6] + ':' + data_backup[i][7];
  }
  var btn = document.createElement('button');
  btn.setAttribute("id", i + '_clear');
  btn.setAttribute('class', 'kbtn');
  btn.innerText = 'クリア';
  btn.setAttribute('onclick', "clear_click(this)");
  td.appendChild(btn);
}

//出欠等のボタンクリック時の動作
function btclick(btn) {
  var btn_id = btn.id;          //列番号を取得
  var td = document.getElementById(btn_id + '_syukketu');
  var btn_text = btn.textContent;   //欠席等の文字列を取得
  time_now = get_time_now();    //ボタンを押した時刻を取得
  data1[btn_id][5] = btn_text;      //データに欠席等を入力
  data1[btn_id][8] = time_now;      //データに時刻を入力
  google.script.run.SetSpreadSheet("欠席入力", btn_text, btn_id, 5); //スプレットシートに出欠等を記録
  google.script.run.SetSpreadSheet("欠席入力", time_now, btn_id, 8); //スプレットシートに時間を記録

  //ボタン等を一旦削除
  while (td.lastChild) {
    td.removeChild(td.lastChild);
  }
  //押したボタンによってリスト内容変更
  if (btn_text == "出停") {
    select_item = reason_list[2];
  } else if (btn_text == "忌引") {
    select_item = reason_list[3];
  } else {
    select_item = reason_list[1];
  }
  var selecter = document.createElement('select');
  selecter.setAttribute('id', btn_id + '_select_reason');
  selecter.setAttribute('name', btn_id + '_' + btn_text);

  let op = document.createElement("option");
  op.value = "";
  op.text = "事由を選択";
  selecter.appendChild(op);
  for (var i = 1; i < select_item.length; i++) {
    if (select_item[i] == "") { break }
    else {
      let op = document.createElement("option");
      op.value = btn_id + '_' + btn_text + '_' + select_item[i];
      op.text = select_item[i];
      selecter.appendChild(op);
    }
  }
  let op1 = document.createElement("option");

  op1.value = btn_id + '_' + btn_text + '_その他';
  op1.text = 'その他';
  selecter.appendChild(op1);

  if (hyouziUI ==1 ){
    td.textContent = '[' + time_now + ']' + btn_text + '　:';
    td.setAttribute('align', 'left');
    td.style.color = "red";
  } else {
    var text = data1[btn_id][0] + " " + data1[btn_id][1] + "  " + data1[btn_id][4] + "\n" + data1[btn_id][2];
    td.textContent = text + '\n' + '[' + time_now + ']' + btn_text + '　:';
    td.setAttribute("class", "btn btn--orange_black");
  }
  td.style.fontWeight = "bold";
  td.appendChild(selecter);

  var btn = document.createElement('button');
  btn.setAttribute("id", btn_id + '_clear');
  btn.setAttribute('class', 'kbtn');
  btn.innerText = 'クリア';
  btn.setAttribute('onclick', "clear_click(this)");
  td.appendChild(btn);
}

//欠席等の理由をリストから選択した時の動作
document.addEventListener('change', function (e) {
  var elem = e.target.id;     //idを取得

  if (elem == null || elem == 'hurigana' || elem == 'birthday' || elem == 'fm' || elem.indexOf('_text') == 0 ){ }
  else if (elem == 'select_class') { }    //クラス選択なら何もしない
  else if (elem == "slider"){
    changeFontSize();
  }
  else {
    var result = document.getElementById(elem).value;
    var id_s = result.split('_');

    var row_n = id_s[0];  //欠席等のrow番号を取得
    var ziyu = id_s[1];   //事由を取得
    var reason = id_s[2]; //理由を取得
    var time_e = data1[row_n][8]; //入力時間を取得

    console.log(row_n + ":" + ziyu + ":" + reason);
    data1[row_n][6] = reason;      //データに欠席等を入力
    google.script.run.SetSpreadSheet("欠席入力", reason, row_n, 6); //スプレットシートの事由を消去

    var td = document.getElementById(row_n + '_syukketu');
    //ボタン等を一旦削除
    while (td.lastChild) {
      td.removeChild(td.lastChild);
    }

    //理由選択を追加、そしてその他をの理由を追加
    if (hyouziUI ==1 ){
      td.textContent = '[' + time_e + ']' + ziyu + '　:　' + reason + '　:　';
    } else {
      var text = data1[row_n][0] + " " + data1[row_n][1] + "  " + data1[row_n][4] + "\n" + data1[row_n][2];
      td.innerText = text + '\n' + '[' + time_e + ']' + ziyu + '　:　' + reason + '　:　';
    }
    var input = document.createElement('input');
    input.setAttribute('id', row_n + '_text');
    input.setAttribute('type', 'text');
    input.setAttribute('name', ziyu + '　:　' + reason + '　:　');
    input.setAttribute('size', '5px');

    var add_btn = document.createElement('button');
    add_btn.setAttribute("id", row_n + '_add');
    add_btn.setAttribute('class', 'kbtn');
    add_btn.innerText = '追加';
    add_btn.setAttribute('onclick', "add_click(this)");

    var btn = document.createElement('button');
    btn.setAttribute("id", row_n + '_clear');
    btn.setAttribute('class', 'kbtn');
    btn.innerText = 'クリア';
    btn.setAttribute('onclick', "clear_click(this)");
    td.appendChild(input);
    td.appendChild(add_btn);
    td.appendChild(btn);
  }
});

//追加ボタンの動作
function add_click(btn) {
  var btn_id = btn.id;
  var i = btn_id.split('_')[0];   //列番号を取得
  var element = document.getElementById(i + '_text');
  var td = document.getElementById(i + '_syukketu');
  btn_text = element.name;
  input_text = element.value;
  data1[i][7] = input_text;      //データにその他の理由を入力
  google.script.run.SetSpreadSheet("欠席入力", input_text, i, 7); //スプレットシートにその他の理由を入力

  while (td.lastChild) {
    td.removeChild(td.lastChild);
  }
  if (hyouziUI ==1 ){
    td.textContent = '[' + data1[i][8] + ']' + btn_text + input_text + '　　';
  } else {
    var text = data1[i][0] + " " + data1[i][1] + "  " + data1[i][4] + "\n" + data1[i][2];
    td.innerText = text + '\n' + '[' + data1[i][8] + ']' + btn_text + input_text + '　　';
  }
  var btn = document.createElement('button');
  btn.setAttribute("id", i + '_clear');
  btn.setAttribute('class', 'kbtn');
  btn.innerText = 'クリア';
  btn.setAttribute('onclick', "clear_click(this)");
  td.appendChild(btn);
}

function checkbox_cell(obj, id) {
  var CELL = document.getElementById(id);
  var TABLE = CELL.parentNode.parentNode.parentNode;
  for (var i = 0; TABLE.rows[i]; i++) {
    TABLE.rows[i].cells[CELL.cellIndex].style.display = (obj.checked) ? '' : 'none';
  }
}

//クリアボタンの動作
function clear_click(btn) {
  var btn_id = btn.id;
  var i = btn_id.split('_')[0];          //列番号を取得
  // console.log(btn_id);
  var td = document.getElementById(i + '_syukketu');
  data1[i][5] = '';      //データに欠席等を入力
  data1[i][6] = '';      //データに事由を入力
  data1[i][7] = '';      //データにその他を入力
  google.script.run.SetSpreadSheet("欠席入力", '', i, 5); //スプレットシートの理由を消去
  google.script.run.SetSpreadSheet("欠席入力", '', i, 6); //スプレットシートの事由を消去
  google.script.run.SetSpreadSheet("欠席入力", '', i, 7); //スプレットシートのその他を消去
  google.script.run.SetSpreadSheet("欠席入力", '', i, 8); //スプレットシートの時間を消去
  //ボタン等を一旦削除
  while (td.lastChild) {
    td.removeChild(td.lastChild);
  }
  if (hyouziUI==1){}else{
    var text = data1[i][0] + " " + data1[i][1] + "  " + data1[i][4] + "\n" + data1[i][2];
    td.innerText = text + '\n'
  }
  var j = data_backup[i][8];
  if (j == "") {
    make_buttons(td, i, 0);
  } else {
    make_buttons(td, i, 1);
  }
}

//欠席者数一覧の表を作成
function make_table_kesseki() {
  var tableEle = document.getElementById('data-table');
  tableEle.innerText = "";
  var kesseki_data = [];
  //欠席入力のE1からI6までを読み込み表示
  for (var i = 0; i < 6; i++) {
    var kesseki_data_row = [];
    for (var j = 4; j < 9; j++) {
      kesseki_data_row[j - 4] = data1[i][j];
    }
    kesseki_data[i] = kesseki_data_row;
  }
  for (var i = 0; i < kesseki_data.length; i++) {
    if (kesseki_data[i][0] == "") {
    } else {
      var tr = document.createElement('tr');
      for (var j = 0; j < kesseki_data[i].length; j++) {
        var td = document.createElement('td');
        var text = "　　" + kesseki_data[i][j] + "　　";
        td.innerText = text;
        tr.appendChild(td);
      }
      tableEle.appendChild(tr);
    }
  }
}

//欠席等入力無しのときの、ボタン一覧を作成する関数  rev=1で、復元ボタン作成
function make_buttons(td, i, rev) {
  var btn1 = document.createElement('button');
  var btn2 = document.createElement('button');
  var btn3 = document.createElement('button');
  var btn4 = document.createElement('button');
  var btn5 = document.createElement('button');
  var btn6 = document.createElement('button');
  var btn7 = document.createElement('button');
  var btn8 = document.createElement('button');

  btn1.setAttribute("id", i);
  btn1.innerText = "欠席";
  btn1.setAttribute("onclick", "btclick(this)");
  btn1.setAttribute("class", "kbtn kbtn-kesseki");

  btn2.setAttribute("id", i);
  btn2.innerText = "出停";
  btn2.setAttribute("onclick", "btclick(this)");
  btn2.setAttribute("class", "kbtn kbtn-syuttei");

  btn3.setAttribute("id", i);
  btn3.innerText = "忌引";
  btn3.setAttribute("onclick", "btclick(this)");
  btn3.setAttribute("class", "kbtn kbtn-kibiki");

  btn4.setAttribute("id", i);
  btn4.innerText = "遅刻";
  btn4.setAttribute("onclick", "btclick(this)");
  btn4.setAttribute("class", "kbtn kbtn-tikoku");

  btn5.setAttribute("id", i);
  btn5.innerText = "早退";
  btn5.setAttribute("onclick", "btclick(this)");
  btn5.setAttribute("class", "kbtn kbtn-soutai");

  btn6.setAttribute("id", i);
  btn6.innerText = "転出";
  btn6.setAttribute("onclick", "btclick(this)");
  btn6.setAttribute("class", "kbtn kbtn-tensyutu");

  btn7.setAttribute("id", i);
  btn7.innerText = "未定";
  btn7.setAttribute("onclick", "btclick(this)");
  btn7.setAttribute("class", "kbtn kbtn-tikoku");

  td.setAttribute("id", i + '_syukketu');
  if (hyouzi[0] == 1) { td.appendChild(btn1); }
  if (hyouzi[1] == 1) { td.appendChild(btn2); }
  if (hyouzi[2] == 1) { td.appendChild(btn3); }
  if (hyouzi[3] == 1) { td.appendChild(btn4); }
  if (hyouzi[4] == 1) { td.appendChild(btn5); }
  if (hyouzi[5] == 1) { td.appendChild(btn6); }
  if (hyouzi[6] == 1) { td.appendChild(btn7); }
  if (rev == 1) {
    btn8.setAttribute("id", i);
    btn8.innerText = "復元";
    btn8.setAttribute("onclick", "rev_click(this)");
    btn8.setAttribute("class", "kbtn kbtn-kesseki");
    td.appendChild(btn8);
  }
}

//再読み込み
function refresh() {
  var obj = document.getElementById('select_class');
  while (obj.lastChild) {
    obj.removeChild(obj.lastChild);
  }
  let op = document.createElement("option");
  op.value = "";
  op.text = "読み込み中…少々お待ちください";
  obj.appendChild(op);
  hyouziUI = 2;
  change_hyouziUI();
  var tableEle = document.getElementById('data-table');  //表のオブジェクトを取得
  tableEle.innerText = "";
  start_html();
}

//全校一覧を作成
function itiran() {
  hyouziUI = 2;
  kind = 'all_gr' ;
  change_hyouziUI();
  make_table(hyouziUI);
}

//欠席情報入力者のみの一覧を作成
function k_itiran() {
  hyouziUI = 2;
  kind = 'ziyu' ;
  change_hyouziUI();
  make_table(hyouziUI);
}

// スライダーの値が変わったらフォントサイズを変更する関数
function changeFontSize() {
  // スライダーとデータ要素を取得
  var slider = document.getElementById("slider");
  var root = document.documentElement;
  var data = document.getElementById("data-table");
  var font_size = document.getElementById("fontsize");
  // スライダーの値に応じてフォントサイズを計算
  var fontSize = slider.value + "px";
  // データ要素のフォントサイズを設定
  // ルート要素のカスタムプロパティの値を設定
  root.style.setProperty("--font-size", fontSize);
  data.style.fontSize = fontSize;
  font_size.innerText = slider.value;
}

function get_time_now() {
  let now = new Date();
  let Month = now.getMonth() + 1;
  let DateA = now.getDate().toString().padStart(2, '0');
  let Hour = now.getHours().toString().padStart(2, '0');
  let Min = now.getMinutes().toString().padStart(2, '0');
  let Sec = now.getSeconds().toString().padStart(2, '0');
  let time = Hour + ':' + Min; // + ':' +Sec
  return time;
}
</script >
<script>
const week = ["日", "月", "火", "水", "木", "金", "土"];
var k_ziyu = []; //["欠席","出停","忌引","遅刻","早退","学級閉鎖","学年閉鎖","臨時休業","未定"];
var kesseki_riyu = []; //["頭痛","咽頭痛","咳","腹痛","吐気","家事都合","体調不良","けが","喘息"]
var teisi_riyu = []; //["感染症","家族の風邪","濃厚接触者","インフルエンザ","帯状疱疹","頭痛","発熱","風邪症状"];
var kibiki_riyu = []; //["祖父死去","祖母死去","曾祖父死去","曾祖母死去","葬儀"]
const today = new Date();
// 月末だとずれる可能性があるため、1日固定で取得
var showDate = new Date(today.getFullYear(), today.getMonth(), 1);
var today_org = new Date(today.getFullYear(), today.getMonth(), today.getDate());
var grade = 0;   //学年
var tclass = 0;   //学級
var student_id = 0; //生徒の行番号
var dataList = []; // データを入れるための配列
var dataList_backup = [];
var reason_list =[];  //事由、理由一覧
var hyouziUI = 1; // 1:個人　2:一覧

//最初に実行
document.addEventListener('DOMContentLoaded', function () {
  start_html();
}, false);

async function start_html(){
  console.log("start",new Date().getSeconds());
  dataList = await get_sheet_data("欠席記録");  //欠席記録シート全体のデータを取得
  console.log("dataList OK",new Date().getSeconds());
  dataList_backup = JSON.parse(JSON.stringify(dataList));  //読み込んだデータのバックアップを保存
  console.log("dataList_backup OK",new Date().getSeconds());
  top_data = await get_sheet_data("リスト");  //リストシート全体のデータを取得
  // hyouziUI = top_data[11][38];     //表示初期値を取得

  //欠席等の事由一覧を取得
  for (let i = 0; i<4 ; i++){  //33列行目から36列目 -1
    const reason_item =[];
    for (let j=0 ; j<32 ;j++){  //3行目から33列目 -1
      if (top_data[j+2][i+32] ==""){break;}else{
      reason_item[j] = top_data[j+2][i+32];
      }
    }
    reason_list[i] = reason_item;
    console.log("i="+i,reason_item);
  }
  k_ziyu= reason_list[0];
  kesseki_riyu = reason_list[1];
  teisi_riyu =reason_list[2];
  kibiki_riyu =reason_list[3];
  console.log("top_data OK",new Date().getSeconds());
  class_list();               //クラスリストを作成
  console.log("class_list OK",new Date().getSeconds());
  make_data(hyouziUI);        //表示データを作成
}

//表示データを作成
function make_data(hyouziUI){
  let calendar = document.getElementById('calendar');
  let tableEle = document.getElementById('data_table');
  if (hyouziUI ==1){          //個人表示なら
    tableEle.innerText ="";   //一旦一覧表を削除
    showProcess(showDate, calendar, student_id);
  } else if (hyouziUI ==2){   //一覧表示なら
    calendar.innerText ="";   //一旦カレンダーを削除
    showProcess_itiran(showDate);
  }
}

//表示形式の切り替えボタン関数
function change_hyouziUI() {
  let calendar = document.getElementById('calendar');
  let tableEle = document.getElementById('data_table');
  let change_hyouziUI_text = document.getElementById('change_hyouziUI');
  if (hyouziUI == 1) {           //個人表示なら
    calendar.style.height = '0vh';
    hyouziUI = 2;
    change_hyouziUI_text.innerText = "個人表示へ";
    select_class();
    }
  else if (hyouziUI == 2) {     //一覧表示なら
    calendar.style.height = 'auto';
    tableEle.style.height = '0vh';
    hyouziUI = 1;
    change_hyouziUI_text.innerText = "一覧表示へ";
    student_id = 0;             //選択している生徒の行番号を初期化
    select_class();
    }
  // メインの表を作成
  make_data(hyouziUI);
}

// 前の月表示
function prev() {
	if (showDate.getMonth() == 3){  //4月より前は戻らない
		return
	}else{
	showDate.setMonth(showDate.getMonth() - 1);
  make_data(hyouziUI);
	}
}

// 次の月表示
function next() {
	if (showDate.getMonth() == 2){  //3月より先は進まない
		return
	}else{
	showDate.setMonth(showDate.getMonth() + 1);
	make_data(hyouziUI);
	}
}

// カレンダー表示
function showProcess(date, calendar, student_id) {
	let year = date.getFullYear();
	let month = date.getMonth(); // 0始まり
	document.querySelector('#header').innerHTML = "欠席記録確認カレンダー　" + year + "年 " + (month + 1) + "月";

	let calendar = createProcess(year, month, student_id);
	document.querySelector('#calendar').innerHTML = calendar;
}

// 一覧表示
function showProcess_itiran(date) {
	let year = date.getFullYear();
	let month = date.getMonth(); // 0始まり
  document.querySelector('#header').innerHTML = "欠席記録確認カレンダー　" + year + "年 " + (month + 1) + "月";
  let data_table = make_itiran(date);
  document.querySelector('#data_table').innerHTML = data_table;
  let op = document.createElement("option");
  op.value = "";
  op.text = "学級を選択";
}

// カレンダー作成
function createProcess(year, month, student_id) {
	// 曜日
	let calendar = "<table class='calendar_table'><tr class='dayOfWeek'>";
	for (let i = 0; i < week.length; i++) {
		calendar += "<th>" + week[i] + "</th>";
	}
	calendar += "</tr>";

	let count = 0;
	let startDayOfWeek = new Date(year, month, 1).getDay();	//1日の開始曜日 0:日曜日～6:土曜日
	let endDate = new Date(year, month + 1, 0).getDate();		//月の最終日
	let lastMonthEndDate = new Date(year, month, 0).getDate();	//前の月の最終日
	let row = Math.ceil((startDayOfWeek + endDate) / week.length);	//カレンダーの行数

 	let temp_date = (month + 1) + "月" + "1日";
	for (let start_col = 6; start_col < dataList[7].length; start_col++) {
		if (dataList[7][start_col] == temp_date) {
			break;
		}
	}
	// 1行ずつ設定
	for (let i = 0; i < row; i++) {
		calendar += "<tr>";
		// 1colum単位で設定
		for (let j = 0; j < week.length; j++) {
			if (i == 0 && j < startDayOfWeek) {
				// 1行目で1日まで先月の日付を設定
				calendar += "<td class='disabled'>" + (lastMonthEndDate - startDayOfWeek + j + 1) + "<br>" + "</td>";
			} else if (count >= endDate) {
				// 最終行で最終日以降、翌月の日付を設定
				count++;
				calendar += "<td class='disabled'>" + (count - endDate) + "<br>" + "</td>";
			} else {
				//欠席データを取得
        const kesseki = kesseki_data(student_id, start_col, year, month,  count);

        // 当月の日付を曜日に照らし合わせて設定
				count++;
				let dateInfo = checkDate(year, month, count, start_col);
				let btn = document.createElement('button');
				
				//当日なら
				if (dateInfo.isToday) {
					if (student_id == 0) {
						calendar += "<td class='today'>" + count + "<br>" + kesseki;
					} else {
						calendar += "<td id=" +count+"_"+ student_id + "_" + (start_col + 3 * count-3) + " class='today'>" + count + "<br>" + kesseki + "</td>";
					}
				//祝日なら
				} else if (dateInfo.isHoliday) {
					calendar += "<td class='holiday' title='" + dateInfo.holidayName + "'>" + count + "<br>" + kesseki + "</td>";
				//休業日なら
				} else if (dateInfo.kyugyoday) {
					calendar += "<td class='kyugyo'>" + count + "<br><span>休日等" + kesseki + "</span></td>";
				//平日
				} else {
					if (student_id == 0) {
						calendar += "<td id='" + count + "'>" + count + "<br><font size='1.2'>" + kesseki + "</font></td>";
					} else {
						btn = "<button id=" +count+"_"+ student_id + "_" + (start_col + 3 * count-3) + "_btn onclick='syusei_click(this)' class='btn btn_kesseki'>修正</button>"
						calendar += "<td id='" + count +"_"+ student_id + "_" + (start_col + 3 * count-3) +"'>" + count + "<br><span>" + kesseki + btn + "</span></td>";
					}
				}
			}
		}
		calendar += "</tr>";
	}
	return calendar;
}

//一覧表示作成
function make_itiran(date){
	let year = date.getFullYear();
	let month = date.getMonth(); // 0始まり
	let data_table = "" ;//document.getElementById(data_table);
	let count = 0;
	let startDay = new Date(year, month, 1).getDate();	//月の開始日
	let endDate = new Date(year, month + 1, 0).getDate();		//月の最終日

	let temp_date = (month + 1) + "月" + "1日";
	//開始日の列番号を取得
	for (let start_col = 6; start_col < dataList[7].length; start_col++) {
		if (dataList[7][start_col] == temp_date) {
			break;
		}
	}
	//１行目の日にち一覧を作成
	if (grade == 0){
	} else {
		data_table += "<tr><th class='fixed02'>番号　名前</th>";

		for (let i = 1; i < (endDate-startDay+2); i++) {
      const week_day = new Date(year, month, i).getDay(); //曜日の数字を取得
      data_table += "<th class='fixed02'>" + i +"("+ week[week_day] +")</th>";  //日付＋曜日
		}

		//2行目以降の生徒一覧を作成
		for (student_id = 9; student_id < 510; student_id++) {    //行番号10行目から510行目まで
			id_num = dataList[student_id][1];                       //4桁の番号を取得

			if (grade=="特支"){                 //特支の場合の処理
				id_gr = dataList[student_id][0];
				if (id_gr == grade) {
					data_table += "<tr><th class='fixed01'>"+dataList[student_id][1]+"<br>"+dataList[student_id][2]+"</th>";
					for (let count=1; count <(endDate-startDay+2); count++){
						//欠席データを取得
						const kesseki = kesseki_data(student_id, start_col-3, year, month,  count);
						const day_data = make_day(kesseki, year, month, count, student_id, start_col);
						data_table += day_data;
					}
					data_table += "</tr>";
				}
			} else {      //通常学級の場合の処理
				id_gr = dataList[student_id][0];
				id_cl = dataList[student_id][1][1];
				if ((id_gr == grade)&&(id_cl==tclass)) {
					data_table += "<tr><th class='fixed02'>"+dataList[student_id][1]+"<br>"+dataList[student_id][2]+"</th>";
					for (let count=1; count <(endDate-startDay+2); count++){
						//欠席データを取得
						const kesseki = kesseki_data(student_id, start_col-3, year, month,  count);
						const day_data = make_day(kesseki, year, month, count, student_id, start_col);
						data_table += day_data;
					}
					data_table += "</tr>";
				}
			}
		}
		data_table += "</th></tr>";
	}
	return data_table;
}

//欠席データを取得
function kesseki_data(student_id, start_col, year, month,  count){
  if ((student_id == 0) || ((start_col + 3 * count) >= dataList[7].length)) { //初期状態は何も表示しない
    kesseki = "";
  } else if (today_org >= new Date(year, month, count + 2)) {
    //欠席記録を取得する
    if (dataList[student_id][start_col + 3 * count] != "") {      //欠席事由
      t_kesseki = dataList[student_id][start_col + 3 * count];
    } else { t_kesseki = ""; }
    if (dataList[student_id][start_col + 3 * count + 1] != "") {  //理由
      t_ziyu = ":" + dataList[student_id][start_col + 3 * count + 1]
    } else { t_ziyu = ""; }
    if (dataList[student_id][start_col + 3 * count + 2] != "") {  //その他詳細
      t_syosai = ":" + dataList[student_id][start_col + 3 * count + 2]
    } else { t_syosai = ""; }
    kesseki = t_kesseki + t_ziyu + t_syosai + "<br>";
  } else {
    kesseki = ""
  }
  return kesseki;
}

// 当月の日付を曜日に照らし合わせて設定　一覧用
function make_day(kesseki, year, month, count, student_id, start_col){
  let dateInfo = checkDate(year, month, count, start_col);
  let btn = document.createElement('button');
  
  //当日なら
  if (dateInfo.isToday) {
    if (hyouziUI == 2) {
      return "<td class='today'>" + kesseki;
    } else {
      return "<td id=" +count+"_"+ student_id + "_" + (start_col + 3 * count-3) + " class='today'><br>" + kesseki + "</td>";
    }
  //祝日なら
  } else if (dateInfo.isHoliday) {
    return "<td class='holiday' title='" + dateInfo.holidayName + "'>" + count + "<br>" + kesseki + "</td>";
  //休業日なら
  } else if (dateInfo.kyugyoday) {
    return "<td class='kyugyo'><span>休日等" + kesseki + "</span></td>";
  //平日
  } else {
    if (student_id == 0) {
      return "<td id='" + count + "'><font size='1.2'>" + kesseki + "</font></td>";
    } else {
      btn = "<button id=" +count+"_"+ student_id + "_" + (start_col + 3 * count-3) + "_btn onclick='syusei_click(this)' class='btn btn_kesseki'>修正</button>"
      return "<td id='" + count +"_"+ student_id + "_" + (start_col + 3 * count-3) +"'><span>" + kesseki + btn + "</span></td>";
    }
  }
}

//　ボタンクリック動作　//

//欠席等の理由をリストから選択した時の動作
document.addEventListener('change', function (e) {
  const elem = e.target.id;         //idを取得 　日にち_生徒行番号_日付列_事由の種類
	const e_day = elem.split('_')[0];	//日にちを取得
	const kind = elem.split('_')[3];	//事由の種類を取得

  if (kind == 'ziyu' || kind == 'riyu' ){ //欠席等事由を選択、もしくは理由を選択したとき
		const result = document.getElementById(elem).value;
		const student_id = Number(result.split('_')[1]);
		const col = Number(result.split('_')[2]);
		const zi = result.split('_')[3];		//事由を取得
		const td_id = result.split('_')[0]+"_"+result.split('_')[1]+"_"+result.split('_')[2];
		const td = document.getElementById(td_id);

		if (kind == 'ziyu') {
			if (zi == "欠席" || zi =="遅刻" || zi == "早退"){
				list = kesseki_riyu;
			} else if (zi == "出停" || zi =="学級閉鎖" || zi =="学年閉鎖" || zi =="臨時休業"){
				list = teisi_riyu;
			} else if (zi == "忌引"){
				list = kibiki_riyu;
			}
			rev_btn ="<button id=" +td_id+ " onclick='rev_click(this)' class='btn btn_kesseki'>復元</button>";
      if (hyouziUI ==1 ){
        td.innerHTML = e_day+"<br><span>"+zi+":";
      } else{
        td.innerHTML = "<span>"+zi+":";
      }
			td.append(make_riyu(td_id,list,zi));
			td.append(make_rev(td_id));
			dataList[student_id][col]=zi; 	//データの欠席等を入力
      google.script.run.SetSpreadSheet("欠席記録", dataList[student_id][col], student_id, col);  //シートに欠席等を入力
		}
		else if (kind == "riyu"){
			const zi = result.split('_')[4];		//事由を取得
			const re = result.split('_')[3];		//理由を取得
			input ="<input id="+td_id+"_text"+" size='2' type=text name="+zi+":"+re+" /input>"
			add_btn ="<button id=" +td_id+ " onclick='add_click(this)' class='btn btn_kesseki'>追加</button>"
      if (hyouziUI ==1 ){
        td.innerHTML = e_day+"<br><span>"+zi+":"+re+":"+input+add_btn;
      } else{
        td.innerHTML = "<span>"+zi+":"+re+":"+input+add_btn;
      }
			td.append(make_rev(td_id));
			dataList[student_id][col+1]=re; //データに理由を入力
      google.script.run.SetSpreadSheet("欠席記録", dataList[student_id][col+1], student_id, col+1);  //シートに理由を入力
		}
	}  else if (elem == "slider"){  //スライダー変更時
    changeFontSize();
  }
});

//修正ボタン
function syusei_click(btn) {
  const btn_id = btn.id;										        //ボタンのIDを取得
	const day_id = btn_id.split('_')[0]; 			        //日付を取得
  const student_id = Number(btn_id.split('_')[1]);  //行番号を取得
	const col = Number(btn_id.split('_')[2]); 				//列番号を取得
	const td_id = day_id+"_"+student_id+"_"+col;
  const td = document.getElementById(td_id);
	dataList[student_id][col]=""; 	//データの欠席等を削除
	dataList[student_id][col+1]=""; //データの理由を削除
	dataList[student_id][col+2]=""; //データのその他を削除
  google.script.run.SetSpreadSheet("欠席記録", '', student_id, col); //スプレットシートの欠席等を消去
  google.script.run.SetSpreadSheet("欠席記録", '', student_id, col+1); //スプレットシートの理由を消去
  google.script.run.SetSpreadSheet("欠席記録", '', student_id, col+2); //スプレットシートのその他を消去
  if (hyouziUI ==1 ){
    td.innerHTML = day_id+"<br>";
  }else{
    td.innerHTML = "";
  }
	td.append(make_k_ziyu(td_id));
	td.append(make_rev(td_id));

}

//復元ボタンの動作
function rev_click(btn) {
  const btn_id = btn.id;
	const day_id = btn_id.split('_')[0]; 			        //日付を取得
  const student_id = Number(btn_id.split('_')[1]);  //行番号を取得
	const col = Number(btn_id.split('_')[2]); 				//列番号を取得
	const td_id = day_id+"_"+student_id+"_"+col;
  const td = document.getElementById(td_id);
  dataList[student_id][col] = dataList_backup[student_id][col];       //データに欠席等を復元
  dataList[student_id][col+1] = dataList_backup[student_id][col+1];   //データに理由を復元
  dataList[student_id][col+2] = dataList_backup[student_id][col+2];   //データにその他を復元
  google.script.run.SetSpreadSheet("欠席記録", dataList_backup[student_id][col], student_id, col);      //スプレットシートの欠席等を復元
  google.script.run.SetSpreadSheet("欠席記録", dataList_backup[student_id][col+1], student_id, col+1);  //スプレットシートの理由を復元
  google.script.run.SetSpreadSheet("欠席記録", dataList_backup[student_id][col+2], student_id, col+2);  //スプレットシートのその他を復元

	//欠席記録をバックアップデータから取得する
	if (dataList_backup[student_id][col] != "") {
		t_kesseki = dataList_backup[student_id][col];
	} else { t_kesseki = "" }
	if (dataList_backup[student_id][col+1] != "") {
		t_ziyu = ":" + dataList_backup[student_id][col+1]
	} else { t_ziyu = "" }
	if (dataList_backup[student_id][col+2] != "") {
		t_syosai = ":" + dataList_backup[student_id][col+2]
	} else { t_syosai = "" }
	const kesseki = t_kesseki + t_ziyu + t_syosai + "<br>";

	btn = "<button id=" +day_id+"_"+ student_id + "_" + (col) + "_btn onclick='syusei_click(this)' class='btn btn_kesseki'>修正</button>";
  if (hyouziUI ==1 ){
    td.innerHTML = "<td id='" + day_id +"_"+ student_id + "_" + (col) +"'>" + day_id + "<br><span>" + kesseki + btn + "</span></td>";
  }else{
    td.innerHTML = "<td id='" + day_id +"_"+ student_id + "_" + (col) +"'><span>" + kesseki + btn + "</span></td>";
  }
}

//その他の理由　追加ボタンの動作
function add_click(btn) {
  const td_id = btn.id;
	const day_id = td_id.split('_')[0]; 			      //日付を取得
  const student_id = Number(td_id.split('_')[1]); //行番号を取得
	const col = Number(td_id.split('_')[2]); 				//列番号を取得
  const element = document.getElementById(td_id + '_text');
  const td = document.getElementById(td_id);
  btn_text = element.name;
  input_text = element.value;
  dataList[student_id][col+2] = input_text;   //データにその他の理由を入力
  google.script.run.SetSpreadSheet("欠席記録", input_text, student_id, col+2); //スプレットシートにその他の理由を入力

	btn = "<button id=" +day_id+"_"+ student_id + "_" + (col) + "_btn onclick='syusei_click(this)' class='btn btn_kesseki'>修正</button>";
  if (hyouziUI ==1 ){
  	td.innerHTML= day_id +"<br><span>"+btn_text+":"+input_text+btn;
  } else {
    td.innerHTML= "<span>"+btn_text+":"+input_text+btn;
  }
	td.append(make_rev(td_id));
}

//復元ボタン作成
function make_rev(td_id){
	const rev_btn = document.createElement('button');
	rev_btn.innerText='復元';
	rev_btn.setAttribute('id',td_id);
	rev_btn.setAttribute('class','btn btn_hukugen');
	rev_btn.setAttribute('onclick','rev_click(this)');
	return rev_btn;
}

//欠席事由一覧リストを作成
function make_k_ziyu(btn_id){
	const selecter_ziyu = document.createElement('select');
  selecter_ziyu.setAttribute('id', btn_id + '_ziyu');

  let op = document.createElement("option");
  op.value = "";
  op.text = "事由を選択";
  selecter_ziyu.appendChild(op);
  for (let i = 0; i < k_ziyu.length; i++) {
		let op = document.createElement("option");
		op.value = btn_id + '_' + k_ziyu[i] +'_'+ (i+1);
		op.text = k_ziyu[i];
		selecter_ziyu.appendChild(op);
  }
	let op1 = document.createElement("option");
  op1.value = btn_id + '_' + 'その他'+'_'+ (i+1);
  op1.text = 'その他';
  selecter_ziyu.appendChild(op1);
	return selecter_ziyu;
}
//理由一覧リストを作成
function make_riyu(td_id,list,re){
	const selecter_riyu = document.createElement('select');
  selecter_riyu.setAttribute('id', td_id + '_riyu');

  let op = document.createElement("option");
  op.value = "";
  op.text = "理由を選択";
  selecter_riyu.appendChild(op);
  for (let i = 0; i < list.length; i++) {
		let op = document.createElement("option");
		op.value = td_id + '_' + list[i] +'_'+ re;
		op.text = list[i];
		selecter_riyu.appendChild(op);
  }
	let op1 = document.createElement("option");
	op1.value = td_id + '_' + 'その他'+'_'+ re;
  op1.text = 'その他';
  selecter_riyu.appendChild(op1);
	return selecter_riyu;
}

// 日付チェック
function checkDate(year, month, day, start_col) {
	if (isToday(year, month, day)) {
		return {
			isToday: true,
			isHoliday: false,
			holidayName: "",
			kyugyoday: false,
		};
	}

	let checkHoliday = isHoliday(year, month, day, start_col);
	return {
		isToday: false,
		isHoliday: checkHoliday[0],
		holidayName: checkHoliday[1],
		kyugyoday: checkHoliday[2],
	};
}

// 当日かどうか
function isToday(year, month, day) {
	return (year == today.getFullYear()
		&& month == (today.getMonth())
		&& day == today.getDate());
}

// 祝日かどうか
function isHoliday(year, month, day, start_col) {
	let checkDate = year + '/' + (month + 1) + '/' + day;
	// 1行目はヘッダーのため、初期値1で開始
	for (let i = 1; i < dataList.length; i++) {
		if (dataList[i][0] === checkDate) {
			return [true, dataList[i][1], false];
		} else if (dataList[6][start_col + day * 3 - 3] != "〇") {
			return [false, "", true]
		}
	}
	return [false, "", false];
}

//学級選択時の動作 生徒一覧を作成
function select_class() {
  let select_class = document.getElementById('select_class');
  let select_student = document.getElementById('select_student');
  let class_id = select_class.value;
  grade = class_id.split('_')[0];
  tclass = class_id.split('_')[1];

  //生徒一覧の選択肢を一度削除
  while (select_student.lastChild) {
    select_student.removeChild(select_student.lastChild);
  }
  let op = document.createElement("option");
  op.value = "";

  //hyouziUIによって動作を変える
  if (hyouziUI==2){               //一覧表示のときの動作
    op.text = "←学級を選択";
    select_student.appendChild(op);
    make_data(hyouziUI);
  }else{                           //個人表示のときの動作
    op.text = "生徒を選択";
    select_student.appendChild(op);
    for (let i = 9; i < 510; i++) {     //行番号10行目から510行目まで
      id_num = dataList[i][1];
      if (grade=="特支"){
        id_gr = dataList[i][0];
        if (id_gr == grade) {
          let op = document.createElement("option");
          op.value = i;
          op.text = dataList[i][1] + dataList[i][2];
          select_student.appendChild(op);
        }
      } else {
        id_gr = dataList[i][0];
        id_cl = dataList[i][1][1];
        if ((id_gr == grade)&&(id_cl==tclass)) {
          let op = document.createElement("option");
          op.value = i;
          op.text = dataList[i][1] + dataList[i][2];
          select_student.appendChild(op);
        }
      }
    }
  }
}

//生徒選択時の動作
function select_student() {
	let select_student = document.getElementById('select_student');
	student_id = select_student.value;
	showProcess(showDate, calendar, student_id);
}

// シート名からデータを読み込む関数（コードスクリプトの読み込み）
function get_sheet_data(sheet_name) {
  return new Promise((resolve, reject) => {
    google.script.run
      .withSuccessHandler((result) => resolve(result))
      .GetSpreadsheet(sheet_name);
  });
}
//シートの学級一覧からリストとして返す関数
function class_list() {
  let obj = document.getElementById('select_class');
  while (obj.lastChild) {
    obj.removeChild(obj.lastChild);
  }
  let op = document.createElement("option");
  op.value = "";
  op.text = "学級を選択";
  obj.appendChild(op);

  for (let i = 1; i < 7; i++) {     //行番号3行目から8行目まで
    for (let j = 1; j < 9; j++) {   //38列目から45列目まで
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

// スライダーの値が変わったらフォントサイズを変更する関数
function changeFontSize() {
  // スライダーとデータ要素を取得
  var slider = document.getElementById("slider");
  let root = document.documentElement;
  let data = document.getElementById("data_table");
  let font_size = document.getElementById("fontsize");
  // スライダーの値に応じてフォントサイズを計算
  let fontSize = slider.value + "px";
  // データ要素のフォントサイズを設定
  // ルート要素のカスタムプロパティの値を設定
  root.style.setProperty("--font-size", fontSize);
  data.style.fontSize = fontSize;
  font_size.innerText = slider.value;
}

</script>
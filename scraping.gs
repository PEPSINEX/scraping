// 対象シート名
var targetSheetName = 'スクレイピング'; 

// URL入力欄の最初のセルを定義
var urlFirstRow = 2;
var urlFirstcol = 'B';

// スクレイピングで取得するデータと出力先を定義
var targetInfo = [
  {
    'name'      : 'title',
    'regexp'    : '<title>(.*?)<\/title>',
    'col'       : 3, // 出力先：C列
  },
  {
    'name'      : 'description',
    'regexp'    : '<meta name="description" content=(.*?)>',
    'col'       : 4, // 出力先：D列
  },
  {
    'name'      : 'keywords',
    'regexp'    : '<meta name="keywords" content=(.*?)>',
    'col'       : 5, // 出力先：E列
  },
]

var sheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName);
var urls = sheet.getRange(urlFirstcol + urlFirstRow + ':' + urlFirstcol).getValues();

// ファイルを開いたとき、スクリプト実行ボタンをメニューに追加
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu('スクリプト');
  menu.addItem('スクレイピング実行', 'scraping');
  menu.addToUi();
}

function scraping() {
  for(let i=0;i<urls.length;i++) {

    // URLが空欄であれば終了
    if(urls[i] == '') {　return;　}

    // 進行状況の表示
    SpreadsheetApp.getActiveSpreadsheet().toast((i+1) + '行目スクレイピング実施中');

    let response = UrlFetchApp.fetch(urls[i]);

    for(let m=0;m<targetInfo.length;m++) {
      try {
        let regexp = new RegExp(targetInfo[m]['regexp']);
        let result = response.getContentText().match(regexp);

        sheet.getRange(i + urlFirstRow, targetInfo[m]['col']).setValue(result[1]);
      } catch (e) {
        sheet.getRange(i + urlFirstRow, targetInfo[m]['col']).setValue('なし');
      }
    }
  }
}

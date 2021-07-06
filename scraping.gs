// 対象シート名
var targetSheetName = 'スクレイピング'; 

// URL入力欄の最初のセルを定義（B2）
var urlFirstRowNum = 2;
var urlFirstcolumnNum = 2;

// スクレイピングで取得するデータと出力先を定義
var targetInfo = [
  {
    'name'      : 'title',
    'regexp'    : '<title>(.*?)<\/title>',
    'columnNum' : 3, // 出力先：C列
  },
  {
    'name'      : 'description',
    'regexp'    : '<meta name="description" content=(.*?)>',
    'columnNum' : 4, // 出力先：D列
  },
  {
    'name'      : 'keywords',
    'regexp'    : '<meta name="keywords" content=(.*?)>',
    'columnNum' : 5, // 出力先：E列
  },
]

var sheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName);
var inputUrlRange = sheet.getRange(urlFirstRowNum, urlFirstcolumnNum, sheet.getLastRow()-1);
var urls = inputUrlRange.getValues(); // 表題の1行目を除外

// ファイルを開いたとき、スクリプト実行ボタンをメニューに追加
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu('スクリプト');
  menu.addItem('スクレイピング実行', 'scraping');
  menu.addToUi();
}

function scraping() {
  for(let i=0;i<urls.length;i++) {
    let response = UrlFetchApp.fetch(urls[i]);

    for(let m=0;m<targetInfo.length;m++) {
      try {
        let regexp = new RegExp(targetInfo[m]['regexp']);
        let result = response.getContentText().match(regexp);

        sheet.getRange(i + urlFirstRowNum, targetInfo[m]['columnNum']).setValue(result[1]);
      } catch (e) {
        sheet.getRange(i + urlFirstRowNum, targetInfo[m]['columnNum']).setValue('なし');
      }
    }
  }
}

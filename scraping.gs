// 対象シート名
var targetSheetName = 'スクレイピング'; 

// URL入力欄の最初のセルを定義
var urlFirstRow = 2;
var urlCol = 'B';

// スクレイピングで取得するデータと出力先を定義
var targetInfo = [
  {
    'name'      : 'title',
    'regexp'    : '<title>(.*?)<\/title>',
    'col'       : 'C',
  },
  {
    'name'      : 'description',
    'regexp'    : '<meta name="description" content=(.*?)>',
    'col'       : 'D',
  },
  {
    'name'      : 'keywords',
    'regexp'    : '<meta name="keywords" content=(.*?)>',
    'col'       : 'E',
  },
]

var sheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName);
var urls = sheet.getRange(urlCol + urlFirstRow + ':' + urlCol).getValues();

// ファイルを開いたとき、スクリプト実行ボタンをメニューに追加
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu('スクリプト');
  menu.addItem('スクレイピング実行', 'scraping');
  menu.addToUi();
}

function scraping() {
  inputUrlContainBlankRowCheck();

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

        sheet.getRange(targetInfo[m]['col'] + (i + urlFirstRow)).setValue(result[1]);
      } catch (e) {
        sheet.getRange(targetInfo[m]['col'] + (i + urlFirstRow)).setValue('なし');
      }
    }
  }
}

function inputUrlContainBlankRowCheck() {
  if(urls.length !== urls.filter(String).length) {
    throw new Error('入力したURLに空白行が含まれています。空白行を削除してください');
  }
}

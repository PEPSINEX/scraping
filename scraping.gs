// -----定数指定-----
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

// エラー文言
var errorMessage = {
  'httpAccess': 'アクセスエラー',
  'scraping'  : '該当する値なし',
  'initialCellBlank'  : urlCol + '列' + urlFirstRow + '行目より、URLを入力してください',
  'containsBlankCell' : '入力したURLに空白行が含まれています。空白行を削除してください',
}
// --------------------

// -----共通変数指定-----
var sheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName);
var urlRowValues = sheet.getRange(urlCol + urlFirstRow + ':' + urlCol).getValues();
var urlCount = urlRowValues.filter(String).length;
var urlValues = sheet.getRange(urlCol + urlFirstRow + ':' + urlCol + (urlFirstRow + urlCount - 1)).getValues();
// --------------------

// ファイルを開いたとき、スクリプト実行ボタンをメニューに追加
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu('スクリプト');
  menu.addItem('スクレイピング実行', 'scraping');
  menu.addToUi();
}

function scraping() {
  inputCheck();

  for(let i=0;i<urlValues.length;i++) {
    // URLが空欄であれば終了
    if(urlValues[i] == '') {　return;　}

    // 進行状況の表示
    SpreadsheetApp.getActiveSpreadsheet().toast((i+1) + '行目スクレイピング実施中');

    let response = UrlFetchApp.fetch(urlValues[i]);

    for(let m=0;m<targetInfo.length;m++) {
      try {
        let regexp = new RegExp(targetInfo[m]['regexp']);
        let result = response.getContentText().match(regexp);

        sheet.getRange(targetInfo[m]['col'] + (i + urlFirstRow)).setValue(result[1]);
      } catch (e) {
        sheet.getRange(targetInfo[m]['col'] + (i + urlFirstRow)).setValue('該当する値なし');
      }
    }
  }
}

// -----入力値チェック-----
function inputCheck() {
  initialCellExistCheck();
  midwayBlankCellCheck();
}

function initialCellExistCheck() {
  const initialCellRange = sheet.getRange(urlCol + urlFirstRow);

  if(initialCellRange.isBlank()) {
    throw new Error(errorMessage['initialCellBlank']);
  }
}

function midwayBlankCellCheck() {
  const continuousLastCell = sheet.getRange(urlCol + urlFirstRow).getNextDataCell(SpreadsheetApp.Direction.DOWN).getA1Notation();
  const continuousValues = sheet.getRange(urlCol + urlFirstRow + ':' + continuousLastCell).getValues();
  const continuousCount = continuousValues.filter(String).length;

  if(urlCount !== continuousCount) {
    throw new Error(errorMessage['containsBlankCell']);
  }
}
// --------------------

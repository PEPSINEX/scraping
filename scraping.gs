// -----定数指定-----
var targetSheetName = 'スクレイピング';

var dataInitialRow = 2;
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
  'httpRespons': 'アクセスエラー',
  'scraping'  : '該当する値なし',
  'initialCellBlank'  : urlCol + '列' + dataInitialRow + '行目より、URLを入力してください',
  'containsBlankCell' : '入力したURLに空白行が含まれています。空白行を削除してください',
}
// --------------------

// -----共通変数指定-----
var sheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName);
var urlRowValues = sheet.getRange(urlCol + dataInitialRow + ':' + urlCol).getValues();
var urlCount = urlRowValues.filter(String).length;
var urlValues = sheet.getRange(urlCol + dataInitialRow + ':' + urlCol + (dataInitialRow + urlCount - 1)).getValues();
// --------------------

// -----スクリプト実行ボタン追加処理-----
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu('スクリプト');
  menu.addItem('スクレイピング実行', 'main');
  menu.addToUi();
}
// -----------------------------------

function main() {
  inputCheck();
  resetResult();
  scraping();
}

// -----スクレイピング処理-----
function scraping() {
  // URLごとの処理開始
  for(let i=0;i<urlValues.length;i++) {
    // URLが空欄であれば終了
    if(urlValues[i] == '') {　return;　}

    // 進行状況の表示
    SpreadsheetApp.getActiveSpreadsheet().toast((i+1) + '行目スクレイピング実施中');

    // HTTPレスポンスの取得。取得エラー時は該当セルにメッセージを出力
    let response = getHttpResponsBody(urlValues[i]);
    if(response == 'httpResponsError') {
      for(let m=0;m<targetInfo.length;m++) {
        sheet.getRange(targetInfo[m]['col'] + (i + dataInitialRow)).setValue(errorMessage['httpRespons']);
      }
      return;
    }

    // 必要な項目だけ抽出
    for(let m=0;m<targetInfo.length;m++) {
      try {
        let regexp = new RegExp(targetInfo[m]['regexp']);
        let result = response.getContentText().match(regexp);

        sheet.getRange(targetInfo[m]['col'] + (i + dataInitialRow)).setValue(result[1]);
      } catch (e) {
        sheet.getRange(targetInfo[m]['col'] + (i + dataInitialRow)).setValue(errorMessage['scraping']);
      }
    }
  }
}

function getHttpResponsBody(url) {
  try {
    let httpResponsBody = UrlFetchApp.fetch(url);
    return httpResponsBody;
  } catch (e) {
    return 'httpResponsError';
  }
}
// --------------------

// -----入力値チェック処理-----
function inputCheck() {
  initialCellExistCheck();
  midwayBlankCellCheck();
}

function initialCellExistCheck() {
  const initialCellRange = sheet.getRange(urlCol + dataInitialRow);

  if(initialCellRange.isBlank()) {
    throw new Error(errorMessage['initialCellBlank']);
  }
}

function midwayBlankCellCheck() {
  const continuousLastCell = sheet.getRange(urlCol + dataInitialRow).getNextDataCell(SpreadsheetApp.Direction.DOWN).getA1Notation();
  const continuousValues = sheet.getRange(urlCol + dataInitialRow + ':' + continuousLastCell).getValues();
  const continuousCount = continuousValues.filter(String).length;

  if(urlCount !== continuousCount) {
    throw new Error(errorMessage['containsBlankCell']);
  }
}
// --------------------

// -----既存のスクレイピング結果の削除-----
function resetResult() {
  for(let i=0;i<targetInfo.length;i++) {
    const continuousLastCell = sheet.getRange(targetInfo[i]['col'] + dataInitialRow).getNextDataCell(SpreadsheetApp.Direction.DOWN).getA1Notation();
    sheet.getRange(targetInfo[i]['col'] + dataInitialRow + ':' + continuousLastCell).clearContent();
  }
  SpreadsheetApp.flush();
}
// --------------------------------------

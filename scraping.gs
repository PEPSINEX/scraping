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
  'scraping'  : '該当値なし',
  'initialCellBlank'  : urlCol + '列' + dataInitialRow + '行目より、URLを入力してください',
  'containsBlankCell' : '入力したURLに空白行が含まれています。空白行を削除してください',
}

// 説明文言
var description = [
  '■使用方法・ツール概要',
  '- URL入力後、上部メニュー「スクリプト⇒スクレイピング実行」を押下してください',
  '- スクリプトが実行され、title, description, keywordsが自動出力されます',
  '',
  '■注意点',
  '- 通信状況によりますが、1行につき2秒程度の時間がかかる場合があります',
  '- URLはB列2行目から入力してください。その際、途中に空白行を入れないでください',
  '- 1回のスクレイピングで入力するURLは、最大100件程度までとしてください。タイムアウトになる可能性があります',
  '- HTTPレスポンスエラーの際は「アクセスエラー」と出力されます',
  '- ページの中に該当項目がない場合は「該当値なし」と出力されます',
]
// --------------------

// -----共通変数指定-----
var sheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName);
var urlRowValues = sheet.getRange(urlCol + dataInitialRow + ':' + urlCol).getValues();
var urlCount = urlRowValues.filter(String).length;
var urlValues = sheet.getRange(urlCol + dataInitialRow + ':' + urlCol + (dataInitialRow + urlCount - 1)).getValues();
// --------------------

// -----スクリプト実行ボタン追加処理-----
function onOpen() {
  Browser.msgBox(description.join('\\n'));
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
    if(urlValues[i] == '') { return; }

    // 進行状況の表示
    SpreadsheetApp.getActiveSpreadsheet().toast((i+1) + '行目スクレイピング実施中');

    // HTTPレスポンスの取得
    let response = getHttpResponsBody(urlValues[i]);

    // レスポンス取得状況に応じて処理を分岐
    if(response == 'httpResponsError') {  
      for(let m=0;m<targetInfo.length;m++) {
        sheet.getRange(targetInfo[m]['col'] + (i + dataInitialRow)).setValue(errorMessage['httpRespons']);
      }
    }else {
      for(let m=0;m<targetInfo.length;m++) {
        let result = getMatchValue(response, targetInfo[m]['regexp']);
        sheet.getRange(targetInfo[m]['col'] + (i + dataInitialRow)).setValue(result);
      }
    }

    Utilities.sleep(1000);
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

function getMatchValue(htmlText, targetRegexp) {
  try {
    let regexp = new RegExp(targetRegexp);
    let matchValue = htmlText.getContentText().match(regexp);

    return matchValue[1];
  } catch (e) {
    return errorMessage['scraping'];
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

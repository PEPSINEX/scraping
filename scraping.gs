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
  'match'  : '該当値なし',
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
  '- B列以外は基本的に変更しないでください。スクリプトに影響を与える場合があります',
  '- 1回のスクレイピングで入力するURLは最大100件まで',
  '- HTTPレスポンスエラーの際は「アクセスエラー」と出力されます',
  '- ページの中に該当項目がない場合は「該当値なし」と出力されます',
  '',
  '■その他',
  '- 本文章を再表示する場合は、上部メニュー「スクリプト⇒説明文表示」を押下してください',
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
  displayDescription();
  let ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu('スクリプト');
  menu.addItem('スクレイピング実行', 'main');
  menu.addItem('説明文表示', 'displayDescription');
  menu.addToUi();
}
// -----------------------------------

// -----説明文表示----
function displayDescription() {
  Browser.msgBox(description.join('\\n'));
}
// -----------------

// -----スクリプトメイン処理-----
function main() {
  inputCheck();
  resetResult();
  scraping();
}
// ---------------------------

// -----スクレイピング処理-----
function scraping() {
  // URLごとの処理開始
  for(let i=0;i<urlValues.length;i++) {
    // URLが空欄であれば終了
    if(urlValues[i] == '') { return; }

    // 進行状況の表示
    SpreadsheetApp.getActiveSpreadsheet().toast((i+1) + '行目スクレイピング実施中');

    // HTTPレスポンスの取得
    let response = getHttpRespons(urlValues[i]);

    // レスポンス取得状況に応じて処理を分岐
    if(response == 'httpResponsError') {
      for(let m=0;m<targetInfo.length;m++) {
        sheet.getRange(targetInfo[m]['col'] + (i + dataInitialRow)).setValue(errorMessage['httpRespons']);
      }
    }else {
      let charset = getResponsCharset(response);

      for(let m=0;m<targetInfo.length;m++) {
        let result = getMatchValue(response, targetInfo[m]['regexp'], charset);
        sheet.getRange(targetInfo[m]['col'] + (i + dataInitialRow)).setValue(result);
      }
    }

    // DOS攻撃対策
    Utilities.sleep(1000);
  }
}

function getHttpRespons(url) {
  try {
    let httpRespons = UrlFetchApp.fetch(url);
    return httpRespons;
  } catch (e) {
    return 'httpResponsError';
  }
}

function getResponsCharset(response) {
  let charset = '';
  try {
    let responseCharset = response.getContentText().match(/charset=(.*?)>/gi)[0];

    if(responseCharset.match(/UTF-8/gi)) {
      charset = 'UTF-8';
    }else if(responseCharset.match(/Shift_JIS/gi)) {
      charset = 'Shift_JIS';
    }else if(responseCharset.match(/euc-jp/gi)) {
      charset = 'euc-jp';
    }else {
      charset = '';
    }
  } catch (e) {
    charset = '';
  }

  return charset;
}

function getMatchValue(response, targetRegexp, charset) {
  try {
    let regexp = new RegExp(targetRegexp);
    return response.getContentText(charset).match(regexp)[1];
  } catch (e) {
    return errorMessage['match'];
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

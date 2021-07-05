function onOpen() {
  let ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu('スクリプト');
  menu.addItem('スクレイピング実行', 'myFunction');
  menu.addToUi();
}

function myFunction() {
  // ----------定数定義start---------------
  // 対象シートを取得
  const sheet = SpreadsheetApp.getActive().getSheetByName('スクレイピング');

  // URL入力欄の最初のセルを定義（B2）
  const urlFirstRowNum = 2;
  const urlFirstcolumnNum = 2;

  // URL一覧を取得
  const urls = sheet.getRange(urlFirstRowNum, urlFirstcolumnNum, sheet.getLastRow()-1).getValues();

  // スクレイピングで取得するデータと出力先を定義
  const targetInfo = [
    {
      'name'      : 'title',
      'regexp'    : '<title>(.*?)<\/title>',
      'columnNum' : 3,  // C列
    },
    {
      'name'      : 'description',
      'regexp'    : '<meta name="description" content=(.*?)>',
      'columnNum' : 4,  // D列
    },
    {
      'name'      : 'keywords',
      'regexp'    : '<meta name="keywords" content=(.*?)>',
      'columnNum' : 5,  // E列   
    },
  ]
  // ----------定数定義end-----------------

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

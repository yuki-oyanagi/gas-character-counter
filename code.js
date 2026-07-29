function myFunction() {
  
}

/**
 * 【概要】
 * 選択範囲のセルの文字列から改行文字を除外した文字数をカウントし、
 * セルの内容の下に改行を挟んで「（〇〇文字）」という形式で追記するプログラムです。
 */
function appendCharacterCountToSelectedRange() {
  // アクティブなスプレッドシート（現在操作している表計算ソフト）の、現在表示しているシートを取得する
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // ユーザーが現在スプレッドシート上で選択しているセルの範囲（Active Range: アクティブ範囲）を取得する
  var range = sheet.getActiveRange();
  
  // 指定した範囲内のセルの値（Values）を二次元配列として一括取得する
  var values = range.getValues();
  
  // 書き戻すための新しいデータを格納する配列を準備する
  var newValues = [];
  
  // 取得したデータの行数分だけループ処理を行う
  for (var i = 0; i < values.length; i++) {
    var row = [];
    // 各行の列数分だけループ処理を行う
    for (var j = 0; j < values[i].length; j++) {
      var cellValue = values[i][j];
      
      // セルが空欄ではない場合のみ処理を実行する
      if (cellValue !== "" && cellValue !== null) {
        // セルの内容を文字列（String）として扱う
        var text = String(cellValue);
        
        // すでに「（〇〇文字）」という追記が含まれている場合は、重複を防ぐために一度古い追記部分を削除する
        text = text.replace(/\n（\d+文字）$/, "");
        
        // 改行文字（\n や \r）を除外した文字列を作成する
        // 正規表現（Regular Expression: 特定のパターンに一致する文字列を検索・置換するための表現方法）を用いて、
        // すべての改行を空文字に置き換えてから文字数をカウントする
        var textWithoutNewlines = text.replace(/[\r\n]/g, "");
        var charCount = textWithoutNewlines.length;
        
        // もとの内容の下に改行（\n）を入れて文字数を結合する
        var updatedText = text + "\n（" + charCount + "文字）";
        
        row.push(updatedText);
      } else {
        // セルが空欄の場合はそのまま空欄にする
        row.push("");
      }
    }
    newValues.push(row);
  }
  
  // 編集したデータを一括して選択中の範囲に書き戻す
  range.setValues(newValues);
}

/**
 * 【概要】
 * スプレッドシートを開いたときに自動実行され、上部メニューバーに独自のカスタムメニューを追加するプログラムです。
 */
function onOpen() {
  // スプレッドシートのUI（User Interface: ユーザーインターフェース＝画面操作機能）を取得する
  var ui = SpreadsheetApp.getUi();
  
  // 上部メニューバーに新しいメニューを追加する
  ui.createMenu('文字数ツール') // メニューバーに表示される名前
      .addItem('選択範囲の文字数をカウント', 'appendCharacterCountToSelectedRange') // メニューの中身と呼び出す関数名
      .addToUi(); // スプレッドシートの画面に反映させる
}
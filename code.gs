 let spreadsheet;/* スプレッドシートのデータ */
 let sheet;/* シートのデータ */
 let range_number_rondom;/* 乱数を保存しているセルの情報*/
 let value_number_random;/* 乱数のデータ */

 const console = 'console';/* 入出力画面の役割をするシートのシート名 */
 const buffer = 'buffer';/* 日本語を一時保存するシートのシート名 */
 const data = 'data_toeic';/* 英単語データを保存するシートのシート名 */
 const number_wordsize = 1000;/* dataに保存した単語の数 */

/* 問題を表示 */
function myQuestion() {
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  /* dataに保存されている単語数の範囲で乱数を生成 */
  value_number_random = Math.random();
  value_number_random = Math.floor(value_number_random*number_wordsize);
  value_number_random += 2;
  
  /* bufferに乱数を保存する */
  sheet = spreadsheet.getSheetByName(buffer);
  sheet.getRange(1,1).setValue(value_number_random);

  /* dataから英単語の情報を取得する */
  sheet = spreadsheet.getSheetByName(data);
  let range_word_english = sheet.getRange(value_number_random,1);
  let value_word_english = range_word_english.getValue();
  
  /* consoleに英単語を表示する */
  sheet = spreadsheet.getSheetByName(console);
  sheet.getRange(1,1).setValue(value_word_english);
  for(i = 0; i < 6; i++){
    sheet.getRange(i+2,1).setValue('-');
  }
}

/* 答えを表示 */
function myAnswer(){
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  /* bufferから乱数を取得する */
  sheet = spreadsheet.getSheetByName(buffer);
  range_number_random = sheet.getRange(1,1);
  value_number_random = range_number_random.getValue();

  /* dataから日本語の情報を取得する */
  sheet = spreadsheet.getSheetByName(data);
  let range_word_japanese = [];/* 日本語を保存しているセルの情報[日本語(名詞),日本語(動詞),日本語(形容詞),日本語(副詞),日本語(前置詞),日本語(接続詞)] */
  let value_word_japanese = [];/* 日本語のデータ[日本語(名詞),日本語(動詞),日本語(形容詞),日本語(副詞),日本語(前置詞),日本語(接続詞)] */
  for(i = 0; i < 6; i++){
    range_word_japanese[i] = sheet.getRange(value_number_random,i+2);
    value_word_japanese[i] = range_word_japanese[i].getValue();
  }

  /* consoleに日本語を表示する */
  sheet = spreadsheet.getSheetByName(console);
  if(value_word_japanese[0] != '-'){
    /* 日本語(名詞)の文頭に'(名)'をつける */
    sheet.getRange(2,1).setValue('(名)'+value_word_japanese[0]);
  }
  if(value_word_japanese[1] != '-'){
    /* 日本語(動詞)の文頭に'(動)'をつける */
    sheet.getRange(3,1).setValue('(動)'+value_word_japanese[1]);
  }
  if(value_word_japanese[2] != '-'){
    /* 日本語(形容詞)の文頭に'(形)'をつける */
    sheet.getRange(4,1).setValue('(形)'+value_word_japanese[2]);
  }
  if(value_word_japanese[3] != '-'){
    /* 日本語(副詞)の文頭に'(副)'をつける */
    sheet.getRange(5,1).setValue('(副)'+value_word_japanese[3]);
  }
  if(value_word_japanese[4] != '-'){
    /* 日本語(前置詞)の文頭に'(前)'をつける */
    sheet.getRange(6,1).setValue('(前)'+value_word_japanese[4]);
  }
  if(value_word_japanese[5] != '-'){
    /* 日本語(接続詞)の文頭に'(接)'をつける */
    sheet.getRange(7,1).setValue('(接)'+value_word_japanese[5]);
  }
}

/* イベントハンドラー(スプレッドシート編集時にトリガーする) */
function eventHandler(){
  sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  /* チェックボックスにチェックが入ったら関数myQuestionを実行する */
  if(sheet.getRange(8,1).getValue()){
    myQuestion();
  }
  
  /* チェックボックスのチェックを外したら関数myAnswerを実行する */
  else if(!(sheet.getRange(8,1).getValue())){
    myAnswer();
  }
}
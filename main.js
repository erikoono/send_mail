function SendMail(){
  // 対象のシートを指定
  var ss = SpreadsheetApp.getActive().getSheetByName('アドレス帳');

  // 送信時間をメールテキストで表現できるよう、フォーマット
  var today = Utilities.formatDate(new Date(), 'JST', 'yyyy年M月d日 H時m分s秒');

  // メールタイトルを変数で定義
  var MailTitle = "家計簿からのお知らせ";

  // スプレッドシートの最終行まで取得。よく使う
  var lastRow = ss.getLastRow();

  // 今月のグラフを取得
  var charts = getCharts();
  charts.forEach(function(chart){
    chart.getBlob().getAs('image/png').setName('summary.png');
  })

  // 繰り返し処理を実施
  for(var i=2; i<=lastRow; i++){
    if (ss.getRange(i, 3).getValue() == "送信対象"){
      var rangeA = ss.getRange('A' + i).getValue(); //送信対象のA列を最終行まで取得していく。それを ' rangeA ' という変数に格納
      var rangeB = ss.getRange('B' + i).getValue(); //送信対象のB列を最終行まで取得していく。それを ' rangeB ' という変数に格納

      // メール本文を変数で定義。rangeBの変数は送信時の宛名を入れておく
      // todayは送信時の日付と時間を表現させる
      var MailText = rangeB+"さん"+"\n\nお疲れ様です。\n本日"+today+"時点でのメールを送ります。\n当月の変動費の途中経過は以下のようになっています。";
      
    }


    GmailApp.sendEmail(rangeA,
                       MailTitle,
                       MailText,
                       {attachments: charts[0]}
                      )
  }
}

function getCharts() {
  var currentMonth = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/M');

  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var sheet      = ss.getSheetByName(currentMonth);
  var charts     = sheet.getCharts();

  return　charts;
}
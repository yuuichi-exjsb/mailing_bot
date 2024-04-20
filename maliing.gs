
function main(){
  const ss = SpreadsheetApp.openById("スプレッドシートのID");
  const sheet = ss.getSheetByName('mailing_list');
  var range = sheet.getRange("B2:B26").getValues();

  for(let i = 0;i<range.length;i++){
    getGmail(range[i])
  }
  
}



function chatGPT(prompt){
  //スクリプトプロパティに設定したOpenAIのAPIキーを取得
  const apiKey = ScriptProperties.getProperty('OpenAI_key');
  //ChatGPTのAPIのエンドポイントを設定
  const apiUrl = 'https://api.openai.com/v1/chat/completions';

   //ChatGPTに投げるメッセージを設定
  const messages = [
    {'role': 'system', 'content': '次の文章を要約してください'},
    {'role': 'user', 'content': prompt}
  ];

  //OpenAIのAPIリクエストに必要なヘッダー情報を設定
  const headers = {
    'Authorization':'Bearer '+ apiKey,
    'Content-type': 'application/json',
    'X-Slack-No-Retry': 1
  };

  //オプションの設定(モデルやトークン上限、プロンプト)
  const options = {
    'muteHttpExceptions' : true,
    'headers': headers,
    'method': 'POST',
    'payload': JSON.stringify({
      'model': 'gpt-3.5-turbo',
      'max_tokens' : 2048,
      'temperature' : 0.9,
      'messages': messages})
  };

  //OpenAIのChatGPTにAPIリクエストを送り、結果を変数に格納
  const response_ = JSON.parse(UrlFetchApp.fetch(apiUrl, options).getContentText());
  //ChatGPTのAPIレスポンスをログ出力
  console.log(response_);
  console.log(response_.choices[0].message.content);

  return response_.choices[0].message.content;
}



function getGmail(mailAddress) {
  const gmailSearchString = mailAddress;
  //メール内容を取得
  const threads = GmailApp.search(gmailSearchString, 0, 1); //最新の一件
  const latestMail = GmailApp.getMessagesForThreads(threads)[0][0];

  const mailtitle = latestMail.getSubject(); //メールタイトル
  const mailBody = latestMail.getPlainBody(); //メール本文
  const mailId = latestMail.getId(); //メールID

  console.log(mailBody)

  // シートに同じメールがある場合は処理終わり
  if(isExistMailInSheet(mailId)){
    return;
  }

  const response_gpt = chatGPT(mailBody)

  const response = LINEMessagingApiPush(mailtitle,response_gpt);
  //response.getContentText("UTF-8"); //デバッグ用、正常なら空の配列が返る

  writeMailInSheet(latestMail,response_gpt);
}

// シートにメールが存在する
function isExistMailInSheet(mailId){
  const sqlTarget = SpreadSheetsSQL.open("スプレッドシートID", "シート1");
  const data = sqlTarget.select(["メールID"]).filter('メールID = ' + mailId).result();

  console.log(data)

  if(data.length >= 1){
    return true;
  }

  return false;
}

// シートにデータの書き込み
function writeMailInSheet(mail,response_gpt){
  const sheet = SpreadsheetApp.getActive().getSheetByName("シート1");
  sheet.insertRowAfter(1); //空行の差し込み

  sheet.getRange("A2").setValue(mail.getId()); //メールID
  sheet.getRange("B2").setValue(mail.getDate()); //送信日時
  sheet.getRange("C2").setValue(mail.getSubject()); //メールタイトル
  sheet.getRange("D2").setValue(mail.getPlainBody()); //メール本文
  sheet.getRange("E2").setValue(response_gpt);
}

//LINE通知
function LINEMessagingApiPush(title,body) {
  const accessToken = "Messaging APIトークン";
  const to = "User ID";

  const text = "タイトル:"+title +
               "\n=========="+
               "\nChatGPTによる要約:"+body; 

  const url = "https://api.line.me/v2/bot/message/push";
  const headers = {
    "Content-Type" : "application/json; charset=UTF-8",
    'Authorization': 'Bearer ' + accessToken,
  };

  const postData = {
    "to" : to,
    "messages" : [
      {
        'type':'text',
        'text': text,
      }
    ]
  };

  const options = {
    "method" : "post",
    "headers" : headers,
    "payload" : JSON.stringify(postData)
  };

  return UrlFetchApp.fetch(url, options);
}

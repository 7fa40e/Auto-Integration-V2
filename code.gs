const LINE_ACCESS_TOKEN = '任意のアクセストークンを入力';
const GEMINI_API_KEY    = '任意のAPIキーを入力';
/**
 * LINEからのWebhook（画像送信など）を受け取るメイン関数
 */
function doPost(e) {
  let replyToken = null; 
  let progress = "システム起動"; //エラーメッセージを作るためのやつ
  
  try {
    if (!e || !e.postData || !e.postData.contents) return;
    const json = JSON.parse(e.postData.contents);
    if (!json.events || json.events.length === 0) return;

    const event = json.events[0];
    replyToken = event.replyToken;
    const userId = event.source.userId; 
    
    // 送られてきたのが「画像」ではない場合
    if (event.type !== 'message' || event.message.type !== 'image') {
      if (replyToken) replyToLine(replyToken, "レシートの「画像」を送信してください。");
      return;
    }
    
    // 1. LINEのプロフィールから送信者の名前を取得
    progress = "Retrieving username from LINE";
    const userName = getLineUserName(userId);

    // 2. LINEサーバーから画像データを取得
    progress = "Retrieving image binary data from LINE server.";
    const messageId = event.message.id;
    const imageBlob = getLineImage(messageId);
    
    // 3. Gemini APIで画像を解析
    progress = "Image analysis and JSON generation using the Gemini API.";
    const geminiResult = analyzeReceiptWithGemini(imageBlob);
    
    if (!geminiResult) {
      throw new Error("The response data (JSON) from Gemini could not be properly extracted.");
    }
    
    // 4. スプレッドシートにデータを書き込み
    progress = "Writing data to Google Sheets";
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["日時", "購入者", "店舗名", "合計金額(円)", "購入品目・メモ"]);
    }
    
    sheet.appendRow([
      geminiResult.date || new Date().toLocaleDateString('ja-JP'),
      userName,
      geminiResult.store || "不明な店舗",
      geminiResult.totalAmount || 0,
      geminiResult.items || ""
    ]);
    
    // 5. ユーザーに入力完了の通知をLINEで返信
    progress = "LINEへの完了通知メッセージ送信中";
    const successMessage = `以下の内容でスプレッドシートに記録しました\n\n` +
                           `購入者: ${userName}\n` +
                           `日付: ${geminiResult.date}\n` +
                           `店舗: ${geminiResult.store}\n` +
                           `金額: ¥${Number(geminiResult.totalAmount).toLocaleString()}\n` +
                           `品目: ${geminiResult.items}`;
                           
    replyToLine(replyToken, successMessage);
    
  } catch (error) {
    //GASのログにも残す(効果は限定的)
    console.error(`【失敗したステップ】: ${progress}\n【エラー内容】: ${error.toString()}`);
    
    //LINEに通知
    if (replyToken) {
      const errorMessage = `ERROR \n\n` +
                           `TAG = \n ${progress}\n\n` +
                           `ERROR_MSG\n${error.toString()}\n\n` +
                           `設定、APIキー、または画像の文字の読みやすさを確認してください。`;
      replyToLine(replyToken, errorMessage);
    }
  }
}

/**
 * LINEサーバーから画像バイナリを取得する関数
 */
function getLineImage(messageId) {
  const url = `https://api-data.line.me/v2/bot/message/${messageId}/content`;
  const options = {
    "method": "get",
    "headers": { "Authorization": "Bearer " + LINE_ACCESS_TOKEN }
  };
  const response = UrlFetchApp.fetch(url, options);
  return response.getBlob();
}

/**
 * LINEのユーザーIDからプロフィール名を取得する関数
 */
function getLineUserName(userId) {
  try {
    const url = `https://api.line.me/v2/bot/profile/${userId}`;
    const options = {
      "method": "get",
      "headers": { "Authorization": "Bearer " + LINE_ACCESS_TOKEN }
    };
    const response = UrlFetchApp.fetch(url, options);
    const profile = JSON.parse(response.getContentText());
    return profile.displayName; 
  } catch (e) {
    console.warn("ユーザー名の取得に失敗しました: " + e.toString());
    return "不明なユーザー";
  }
}

/**
 * Gemini API (gemini-2.5-flash) を呼び出して画像を解析する関数
 */
function analyzeReceiptWithGemini(imageBlob) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`;
  
  const base64Image = Utilities.base64Encode(imageBlob.getBytes());
  const mimeType = "image/jpeg"; // MimeTypeエラー対策として固定値に設定
  
  const prompt = "添付されたレシートの画像から「購入日付(date)」「店舗名(store)」「合計金額(totalAmount)」「購入品目の箇条書き(items)」を抽出してください。価格は数字のみにしてください。";

  const payload = {
    "contents": [{
      "parts": [
        { "text": prompt },
        { "inlineData": { "mimeType": mimeType, "data": base64Image } }
      ]
    }],
    "generationConfig": {
      "responseMimeType": "application/json",
      "responseSchema": {
        "type": "OBJECT",
        "properties": {
          "date": { "type": "STRING", "description": "YYYY/MM/DD形式の購入日付" },
          "store": { "type": "STRING", "description": "店舗名" },
          "totalAmount": { "type": "INTEGER", "description": "支払った合計金額の数値" },
          "items": { "type": "STRING", "description": "購入した主な品目をカンマ区切り" }
        },
        "required": ["date", "store", "totalAmount", "items"]
      }
    }
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseText = response.getContentText();
  
  if (response.getResponseCode() !== 200) {
    throw new Error(`Gemini APIエラー (コード: ${response.getResponseCode()})\n詳細: ${responseText}`);
  }
  
  try {
    const jsonResult = JSON.parse(responseText);
    const textAnswer = jsonResult.candidates[0].content.parts[0].text;
    return JSON.parse(textAnswer);
  } catch (e) {
    console.error("応答のパースに失敗。生データ: " + responseText);
    return null;
  }
}

/**
 * LINEにメッセージを返信する関数
 */
function replyToLine(replyToken, text) {
  const url = "https://api.line.me/v2/bot/message/reply";
  const payload = {
    "replyToken": replyToken,
    "messages": [{ "type": "text", "text": text }]
  };
  const options = {
    "method": "post",
    "contentType": "application/json",
    "headers": { "Authorization": "Bearer " + LINE_ACCESS_TOKEN },
    "payload": JSON.stringify(payload)
  };
  UrlFetchApp.fetch(url, options);
}

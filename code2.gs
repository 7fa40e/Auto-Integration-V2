const LINE_ACCESS_TOKEN = '';
const GEMINI_API_KEY = '';
const MASTER_SPREADSHEET_ID = ''

// ====================================================================

function doPost(e) {
  let replyToken = null;
  let progress = "システム起動";
  
  try {
    if (!e || !e.postData || !e.postData.contents) return;
    const json = JSON.parse(e.postData.contents);
    if (!json.events || json.events.length === 0) return;

    const event = json.events[0];
    replyToken = event.replyToken;
    const userId = event.source.userId;
    
    // ----------------------------------------------------------------
    // 1. テキストメッセージの処理（新規作成・共有・ヘルプなど）
    // ----------------------------------------------------------------
    if (event.type === 'message' && event.message.type === 'text') {
      const text = event.message.text.trim();
      
      // 【機能：ヘルプ表示】
      if (text === "help" || text === "ヘルプ" || text === "使い方") {
        progress = "ヘルプメッセージの送信中";
        const helpMessage = `[レシート管理Bot 使い方マニュアル]\n\n` +
                            `1. シートを新しく作る\n` +
                            `「新規作成 [シート名]」と送信します。\n` +
                            `例: 新規作成 家計簿\n` +
                            `※自動的にリンクを知っている全員が閲覧できる状態で作成されます。\n\n` +
                            `2. 友達のシートを自分のLINEに登録する\n` +
                            `友達から共有されたキーを使って「共有 [キー]」と送信します。\n` +
                            `例: 共有 a1b2c3d4\n\n` +
                            `3. レシートを記録する\n` +
                            `「レシート記録」または「シート選択」と送信すると、登録されているシートがボタンで表示されます。\n` +
                            `保存したいシートを選択後、カメラまたはアルバムから画像を送信してください。`;
        
        replyToLine(replyToken, helpMessage);
        return;
      }
      
      // 【機能：新規作成】
      if (text.startsWith("新規作成 ")) {
        progress = "スプレッドシートの新規作成中";
        const parts = text.split(" ");
        const sheetName = parts[1];
        
        if (!sheetName) {
          replyToLine(replyToken, "[作成エラー]\n「新規作成 [シート名]」の形で入力してください。\n例: 新規作成 サークル部費");
          return;
        }
        
        // 管理帳を開いて重複チェック
        const masterSheet = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID).getActiveSheet();
        const data = masterSheet.getDataRange().getValues();
        
        for (let i = 1; i < data.length; i++) {
          // 同じユーザーが同じ名前のシートを作ろうとしていないか確認
          if (data[i][0] === userId && data[i][2] === sheetName) {
            replyToLine(replyToken, `[作成エラー]\n「${sheetName}」という名前のシートは既に作成されています。別の名前を指定してください。`);
            return;
          }
        }
        
        const userName = getLineUserName(userId);
        const sharedKey = Utilities.getUuid().split("-")[0]; // 8文字のランダムキー生成
        
        // 新しいスプレッドシートを作成
        const newSs = SpreadsheetApp.create(`${sheetName}_(${userName})`);
        const newSsId = newSs.getId();
        const sheet = newSs.getActiveSheet();
        sheet.appendRow(["日時", "購入者", "店舗名", "合計金額(円)", "購入品目・メモ"]);
        
        // リンクを知っている全員が閲覧できるように権限を変更
        progress = "ファイルの共有権限変更中";
        const file = DriveApp.getFileById(newSsId);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        // 管理帳に記録
        masterSheet.appendRow([userId, newSsId, sheetName, sharedKey]);
        
        const confirmationMessage = `[作成完了]\n` +
                                    `スプレッドシート「${sheetName}」を新規作成しました。\n\n` +
                                    `閲覧用URL（リンクを知っている全員に公開済み）:\n${newSs.getUrl()}\n\n` +
                                    `このシートの共有用キー:\n${sharedKey}\n\n` +
                                    `[他の人と共有する方法]\n` +
                                    `共同編集したい相手にこのキーを伝えてください。\n` +
                                    `相手がこのトーク画面で「共有 ${sharedKey}」と送信すると、同じシートへの記録が可能になります。`;
        
        replyToLine(replyToken, confirmationMessage);
        return;
      }
      
      // 【機能：共有キーでシートを追加】
      if (text.startsWith("共有 ")) {
        progress = "共有キーによるシートの紐付け中";
        const inputKey = text.split(" ")[1];
        if (!inputKey) {
          replyToLine(replyToken, "[共有エラー]\n「共有 [共有用キー]」の形で入力してください。\n例: 共有 a1b2c3d4");
          return;
        }
        
        const masterSheet = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID).getActiveSheet();
        const data = masterSheet.getDataRange().getValues();
        let found = false;
        
        for (let i = 1; i < data.length; i++) {
          if (data[i][3] && data[i][3].toString() === inputKey) {
            // すでにそのシートを登録していないかも念のためチェック
            let alreadyLinked = false;
            for (let j = 1; j < data.length; j++) {
              if (data[j][0] === userId && data[j][1] === data[i][1]) {
                alreadyLinked = true;
                break;
              }
            }
            
            if (alreadyLinked) {
              replyToLine(replyToken, `既に「${data[i][2]}」はあなたのリストに登録されています。`);
              return;
            }
            
            masterSheet.appendRow([userId, data[i][1], data[i][2], data[i][3]]);
            replyToLine(replyToken, `連動成功しました。\n「${data[i][2]}」があなたの選択肢に追加されました。`);
            found = true;
            break;
          }
        }
        if (!found) replyToLine(replyToken, "該当する共有用キーのシートが見つかりませんでした。");
        return;
      }
      
      // 【機能：メニュー（利用可能シートの一覧をボタンで出す）】
      if (text === "レシート記録" || text === "シート選択") {
        progress = "操作可能シートの検索中";
        sendSheetChoices(userId, replyToken, "どのスプレッドシートに記録しますか？ボタンを選んでから画像を送信してください。");
        return;
      }
    }
    
    // ----------------------------------------------------------------
    // 2. 画像メッセージの処理（レシート解析）
    // ----------------------------------------------------------------
    if (event.type === 'message' && event.message.type === 'image') {
      progress = "画像メッセージの解析処理開始";
      
      const masterSheet = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID).getActiveSheet();
      const data = masterSheet.getDataRange().getValues();
      let targetSheetId = "";
      
      // ユーザーが最後に登録または選択したシートIDを取得
      for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][0] === userId) {
          targetSheetId = data[i][1]; 
          break;
        }
      }
      
      if (!targetSheetId) {
        replyToLine(replyToken, "エラー: まだ有効なシートがありません。「新規作成 [シート名]」で作成してください。");
        return;
      }
      
      const userName = getLineUserName(userId);
      const imageBlob = getLineImage(event.message.id);
      
      progress = "Gemini APIによるレシート解析中";
      const geminiResult = analyzeReceiptWithGemini(imageBlob);
      
      progress = "対象スプレッドシートへの書き込み中";
      const targetSs = SpreadsheetApp.openById(targetSheetId);
      const sheet = targetSs.getActiveSheet();
      
      const outputDate = geminiResult.date || new Date().toLocaleDateString('ja-JP');
      const outputStore = geminiResult.store || "不明な店舗";
      const outputAmount = geminiResult.totalAmount || 0;
      const outputItems = geminiResult.items || "";
      
      sheet.appendRow([
        outputDate,
        userName,
        outputStore,
        outputAmount,
        outputItems
      ]);
      
      // LINE返信用の結果テキストを構築（他の要素もすべて表示するよう拡充）
      const successMessage = `以下の内容で「${targetSs.getName().split("_(")[0]}」に記録しました。\n\n` +
                             `購入者: ${userName}\n` +
                             `日付: ${outputDate}\n` +
                             `店舗: ${outputStore}\n` +
                             `金額: ${Number(outputAmount).toLocaleString()}円\n` +
                             `品目: ${outputItems}`;
      
      replyToLine(replyToken, successMessage);
      return;
    }
    
    // ----------------------------------------------------------------
    // 3. ボタン（クイックリプライ）がタップされた時の処理（Postbackイベント）
    // ----------------------------------------------------------------
    if (event.type === 'postback') {
      progress = "選択されたシートへの画像送信要求を受信中";
      const postbackData = event.postback.data; 
      
      const params = {};
      postbackData.split("&").forEach(p => {
        const kv = p.split("=");
        params[kv[0]] = kv[1];
      });
      
      replyToCameraAction(replyToken, params.sheetId, params.name);
      return;
    }
    
  } catch (error) {
    console.error(`[失敗]: ${progress}\n${error.toString()}`);
    if (replyToken) {
      // 読み取り失敗等のエラーメッセージの最後に特定の案内文を追加
      replyToLine(replyToken, `エラーが発生しました。\n[タイミング]: ${progress}\n[詳細]: ${error.toString()}\n\nお手数ですが、もう一度画像を送信してください。`);
    }
  }
}

/**
 * ユーザーが操作できるシート一覧を「クイックリプライボタン」として送る
 */
function sendSheetChoices(userId, replyToken, titleText) {
  const masterSheet = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID).getActiveSheet();
  const data = masterSheet.getDataRange().getValues();
  const items = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      items.push({
        "type": "action",
        "action": {
          "type": "postback",
          "label": data[i][2], 
          "data": `sheetId=${data[i][1]}&name=${encodeURIComponent(data[i][2])}`, 
          "displayText": `「${data[i][2]}」を選択しました`
        }
      });
    }
  }
  
  if (items.length === 0) {
    replyToLine(replyToken, "まだ操作できるシートがありません。\n\n「新規作成 [シート名]」と送信して新しく作るか、\n「共有 [共有用キー]」と送信して友達のシートを登録してください。");
    return;
  }
  
  const url = "https://api.line.me/v2/bot/message/reply";
  const payload = {
    "replyToken": replyToken,
    "messages": [{
      "type": "text",
      "text": titleText,
      "quickReply": { "items": items }
    }]
  };
  UrlFetchApp.fetch(url, {
    "method": "post", "contentType": "application/json",
    "headers": { "Authorization": "Bearer " + LINE_ACCESS_TOKEN },
    "payload": JSON.stringify(payload)
  });
}

/**
 * シート選択後に、LINEのカメラ・アルバムを開かせるクイックリプライを返す関数
 */
function replyToCameraAction(replyToken, sheetId, sheetName) {
  const url = "https://api.line.me/v2/bot/message/reply";
  const decodedName = decodeURIComponent(sheetName);
  
  const payload = {
    "replyToken": replyToken,
    "messages": [{
      "type": "text",
      "text": `「${decodedName}」への記録準備ができました。\n下のボタンからカメラまたはアルバムを開いて、レシートの画像を送ってください。`,
      "quickReply": {
        "items": [
          { "type": "action", "action": { "type": "cameraRoll", "label": "アルバムから選択" } },
          { "type": "action", "action": { "type": "camera", "label": "カメラを起動" } }
        ]
      }
    }]
  };
  
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(replyToken + "_target", sheetId);
  
  UrlFetchApp.fetch(url, {
    "method": "post", "contentType": "application/json",
    "headers": { "Authorization": "Bearer " + LINE_ACCESS_TOKEN },
    "payload": JSON.stringify(payload)
  });
}

// ーーー 共通関数 ーーー
function getLineImage(messageId) {
  const url = `https://api-data.line.me/v2/bot/message/${messageId}/content`;
  return UrlFetchApp.fetch(url, { "method": "get", "headers": { "Authorization": "Bearer " + LINE_ACCESS_TOKEN } }).getBlob();
}

function getLineUserName(userId) {
  try {
    const response = UrlFetchApp.fetch(`https://api.line.me/v2/bot/profile/${userId}`, { "method": "get", "headers": { "Authorization": "Bearer " + LINE_ACCESS_TOKEN } });
    return JSON.parse(response.getContentText()).displayName; 
  } catch (e) { return "不明なユーザー"; }
}

function analyzeReceiptWithGemini(imageBlob) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`;
  const base64Image = Utilities.base64Encode(imageBlob.getBytes());
  const payload = {
    "contents": [{ "parts": [{ "text": "レシートからdate, store, totalAmount, itemsを抽出してください。" }, { "inlineData": { "mimeType": "image/jpeg", "data": base64Image } }] }],
    "generationConfig": {
      "responseMimeType": "application/json",
      "responseSchema": {
        "type": "OBJECT",
        "properties": {
          "date": { "type": "STRING" }, "store": { "type": "STRING" },
          "totalAmount": { "type": "INTEGER" }, "items": { "type": "STRING" }
        },
        "required": ["date", "store", "totalAmount", "items"]
      }
    }
  };
  const response = UrlFetchApp.fetch(url, { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true });
  return JSON.parse(JSON.parse(response.getContentText()).candidates[0].content.parts[0].text);
}

function replyToLine(replyToken, text) {
  const payload = { "replyToken": replyToken, "messages": [{ "type": "text", "text": text }] };
  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", { "method": "post", "contentType": "application/json", "headers": { "Authorization": "Bearer " + LINE_ACCESS_TOKEN }, "payload": JSON.stringify(payload) });
}

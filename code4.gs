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
    const cache = CacheService.getUserCache();
    
    // ----------------------------------------------------------------
    // 1. テキストメッセージの処理（新規作成・共有・ヘルプ、および【手動修正の入力】）
    // ----------------------------------------------------------------
    if (event.type === 'message' && event.message.type === 'text') {
      const text = event.message.text ? event.message.text.trim() : "";
      
      // 現在このユーザーが「修正入力待ち」状態かチェック
      const userState = cache.get(`state_${userId}`);
      
      if (userState) {
        progress = "手動修正入力の受付処理中";
        const [mode, cacheKey] = userState.split(":");
        const cachedString = cache.get(cacheKey);
        
        if (!cachedString) {
          replyToLine(replyToken, "エラー: 時間が経ちすぎたため修正を破棄しました。もう一度画像を送信してください。");
          cache.remove(`state_${userId}`);
          return;
        }
        
        const cacheData = JSON.parse(cachedString);
        
        if (mode === "edit_date")   cacheData.date = text;
        if (mode === "edit_store")  cacheData.store = text;
        if (mode === "edit_amount") cacheData.amount = text.replace(/[^0-9]/g, "");
        if (mode === "edit_quantity") cacheData.quantity = text.replace(/[^0-9]/g, ""); // ★追加：個数修正
        if (mode === "edit_items")  cacheData.items = text;
        
        cache.put(cacheKey, JSON.stringify(cacheData), 600);
        cache.remove(`state_${userId}`);
        
        //単価×個数の合計をその場で計算して表示
        const currentTotal = Number(cacheData.amount) * Number(cacheData.quantity);
        
        const confirmText = `[修正完了]\n内容を更新しました。再度確認してください。\n\n` +
                            `購入者: ${cacheData.user}\n` +
                            `日付: ${cacheData.date}\n` +
                            `品目: [${cacheData.store}] ${cacheData.items}\n` +
                            `単価: ${Number(cacheData.amount).toLocaleString()}円\n` +
                            `個数: ${cacheData.quantity} 個\n` +
                            `合計: ${currentTotal.toLocaleString()}円`;
                            
        sendConfirmReply(replyToken, confirmText, cacheKey);
        return;
      }
      
      // 【通常機能：ヘルプ表示】
      if (text === "help" || text === "ヘルプ" || text === "使い方") {
        progress = "ヘルプメッセージの送信中";
        const helpMessage = `[help]\n\n` +
                            `1. シートを新しく作る\n` +
                            `「新規作成 [シート名]」と送信します。\n\n` +
                            `2. 友達のシートを登録する\n` +
                            `「共有 [キー]」と送信します。\n\n` +
                            `3. レシートを記録する\n` +
                            `「レシート記録」からシートを選択後、画像を送信してください。\n` +
                            `確認画面の下部ボタンから、内容を自由に手動修正できます。`;
        
        replyToLine(replyToken, helpMessage);
        return;
      }
      
      // 【通常機能：新規作成】
      if (text && text.startsWith("新規作成 ")) {
        progress = "新規フォルダおよびスプレッドシートの作成中";
        const parts = text.split(" ");
        const sheetName = parts[1];
        
        if (!sheetName) {
          replyToLine(replyToken, "[作成エラー]\n「新規作成 [シート名]」の形で入力してください。");
          return;
        }
        
        const masterSheet = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID).getActiveSheet();
        const data = masterSheet.getDataRange().getValues();
        
        for (let i = 1; i < data.length; i++) {
          if (data[i][2] === sheetName) {
            replyToLine(replyToken, `[作成エラー]\n「${sheetName}」という名前のシートは既に存在します。別の名前を指定してください。`);
            return;
          }
        }
        
        const userName = getLineUserName(userId);
        const sharedKey = Utilities.getUuid().split("-")[0];
        
        const folder = DriveApp.createFolder(`${sheetName}_(${userName})`);
        const folderId = folder.getId();
        
        const newSs = SpreadsheetApp.create(`${sheetName}_(${userName})`);
        const newSsId = newSs.getId();
        const ssFile = DriveApp.getFileById(newSsId);
        folder.addFile(ssFile);
        DriveApp.getRootFolder().removeFile(ssFile);
        
        const sheet = newSs.getActiveSheet();
        //Excelの提出フォーマット（C〜H列）＋内部データ（I列〜）に見出しを統一
        sheet.appendRow(["レシート番号", "品目", "単価", "個数", "合計", "残金", "日時", "購入者", "画像URL"]);
        
        folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        masterSheet.appendRow([userId, newSsId, sheetName, sharedKey, folderId]);
        
        // ★修正：スプレッドシートURLに加え、画像フォルダのURL（folder.getUrl()）も一緒に送信
        const confirmationMessage = `[作成完了]\nスプレッドシート「${sheetName}」を新規作成しました。\n\n` +
                            `スプレッドシートURL:\n${newSs.getUrl()}\n\n` +
                            `画像フォルダURL:\n${folder.getUrl()}\n\n` +
                            `共有用キー:\n${sharedKey}`;
        replyToLine(replyToken, confirmationMessage);
        return;
      }
      
      // 【通常機能：共有キーでシートを追加】
      if (text && text.startsWith("共有 ")) {
        progress = "共有キーによるシートの紐付け中";
        const inputKey = text.split(" ")[1];
        if (!inputKey) {
          replyToLine(replyToken, "[共有エラー]\n「共有 [共有用キー]」の形で入力してください。");
          return;
        }
        
        const masterSheet = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID).getActiveSheet();
        const data = masterSheet.getDataRange().getValues();
        let found = false;
        let targetRowIndex = -1;
        
        for (let i = 1; i < data.length; i++) {
          if (data[i][3] && data[i][3].toString() === inputKey) {
            targetRowIndex = i;
            found = true;
            break;
          }
        }
        
        if (!found) {
          replyToLine(replyToken, "該当する共有用キーのシートが見つかりませんでした。");
          return;
        }
        
        const targetSheetId = data[targetRowIndex][1];
        const targetSheetName = data[targetRowIndex][2];
        const targetFolderId = data[targetRowIndex][4] || "";
        
        for (let j = 1; j < data.length; j++) {
          if (data[j][0] === userId && data[j][1] === targetSheetId) {
            replyToLine(replyToken, `既に「${targetSheetName}」はあなたのリストに登録されています。`);
            return;
          }
        }
        
        masterSheet.appendRow([userId, targetSheetId, targetSheetName, inputKey, targetFolderId]);
        
        // ★修正：共有に成功した際にも、スプレッドシートと画像フォルダのリンクを両方表示してあげる
        try {
          const targetSs = SpreadsheetApp.openById(targetSheetId);
          let folderUrlText = "取得できませんでした";
          if (targetFolderId) {
            try { folderUrlText = DriveApp.getFolderById(targetFolderId).getUrl(); } catch(e) {}
          }
          
          const successMessage = `\n「${targetSheetName}」が追加されました。\n\n` +
                                 `スプレッドシートURL:\n${targetSs.getUrl()}\n\n` +
                                 `画像フォルダURL:\n${folderUrlText}`;
          replyToLine(replyToken, successMessage);
        } catch(e) {
          replyToLine(replyToken, `共有に成功しました。\n「${targetSheetName}」が追加されました。`);
        }
        return;
      }
      
      // 【通常機能：メニュー】
      if (text === "シート選択") {
        progress = "操作可能シートの検索およびゴミ箱チェック中";
        sendSheetChoices(userId, replyToken, "どのスプレッドシートに記録しますか？ボタンを選んでから画像を送信してください。");
        return;
      }
    }
    
    // ----------------------------------------------------------------
    // 2. 画像メッセージの処理（確認画面の生成、データはキャッシュへ一時退避）
    // ----------------------------------------------------------------
    if (event.type === 'message' && event.message.type === 'image') {
      progress = "画像メッセージの解析と確認データの生成";
      
      const masterSheet = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID).getActiveSheet();
      const data = masterSheet.getDataRange().getValues();
      let targetSheetId = "";
      let targetFolderId = "";
      let sheetName = "";
      
      for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][0] === userId) {
          targetSheetId = data[i][1]; 
          sheetName = data[i][2];
          targetFolderId = data[i][4] || "";
          break;
        }
      }
      
      if (!targetSheetId) {
        replyToLine(replyToken, "エラー: まだ有効なシートがありません。「新規作成 [シート名]」で作成してください。");
        return;
      }
      
      const userName = getLineUserName(userId);
      const imageBlob = getLineImage(event.message.id);
      
      progress = "画像をGoogleドライブのフォルダに保存中";
      let imageUrl = "";
      if (targetFolderId) {
        try {
          const folder = DriveApp.getFolderById(targetFolderId);
          const timestamp = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd_HHmmss");
          const file = folder.createFile(imageBlob.setName(`receipt_${timestamp}.jpg`));
          imageUrl = file.getUrl();
        } catch(e) {
          const file = DriveApp.createFile(imageBlob);
          imageUrl = file.getUrl();
        }
      } else {
        const file = DriveApp.createFile(imageBlob);
        imageUrl = file.getUrl();
      }
      
      progress = "Gemini APIによる画像解析中";
      const geminiResult = analyzeReceiptWithGemini(imageBlob);
      
      const outputDate = geminiResult.date || new Date().toLocaleDateString('ja-JP');
      const outputStore = geminiResult.store || "不明な店舗";
      const outputAmount = geminiResult.totalAmount || 0;
      const outputItems = geminiResult.items || "";
      
      progress = "確認用メッセージの送信準備";
      const confirmText = `[内容確認]\n以下の内容で「${sheetName}」に記録します。よろしいですか？\n\n` +
                          `購入者: ${userName}\n` +
                          `日付: ${outputDate}\n` +
                          `品目: [${outputStore}] ${outputItems}\n` +
                          `単価: ${Number(outputAmount).toLocaleString()}円\n` +
                          `個数: ${defaultQuantity} 個\n` +
                          `合計: ${Number(outputAmount * defaultQuantity).toLocaleString()}円`;
      
      const cacheData = {
        sheetId: targetSheetId,
        sheetName: sheetName,
        date: outputDate,
        user: userName,
        store: outputStore,
        amount: outputAmount,   // 単価として保持
        quantity: defaultQuantity, // 追加：個数
        items: outputItems,
        img: imageUrl
      };
      
      const cacheKey = "rcpt_" + Utilities.getUuid().split("-")[0];
      cache.put(cacheKey, JSON.stringify(cacheData), 600);
      
      sendConfirmReply(replyToken, confirmText, cacheKey);
      return;
    }
    
    // ----------------------------------------------------------------
    // 3. ポストバックイベント（登録確定・各項目の修正要求モードの起動）
    // ----------------------------------------------------------------
    if (event.type === 'postback') {
      const postbackData = event.postback.data; 
      const params = {};
      postbackData.split("&").forEach(p => {
        const kv = p.split("=");
        params[kv[0]] = kv[1];
      });
      
      // 【処理A：確認画面で「この内容で登録する」が押された場合】
      // 【A：登録ボタンが押されたとき】
      if (params.action === "insert") {
        progress = "確定されたレシートデータの書き込み処理中";
        const cachedString = cache.get(params.cacheKey);
        
        if (!cachedString) {
          replyToLine(replyToken, "エラー: データが見つかりませんでした。お手数ですがもう一度画像を送信してください。");
          return;
        }
        
        const cacheData = JSON.parse(cachedString);
        const targetSs = SpreadsheetApp.openById(cacheData.sheetId);
        const sheet = targetSs.getActiveSheet();
        
        // 修正・追加：Excelの列(C〜H)に完全一致する並び順で書き込み、右側に管理データを逃がす
        const nextReceiptNo = sheet.getLastRow(); // 行数をそのままレシート番号に利用
        const finalTotal = Number(cacheData.amount) * Number(cacheData.quantity); // 単価×個数
        
        sheet.appendRow([
          nextReceiptNo,                                 // A列(ExcelのC)：レシート番号
          `[${cacheData.store}] ${cacheData.items}`,     // B列(ExcelのD)：品目
          Number(cacheData.amount),                      // C列(ExcelのE)：単価
          Number(cacheData.quantity),                    // D列(ExcelのF)：個数
          finalTotal,                                    // E列(ExcelのG)：合計
          "",                                            // F列(ExcelのH)：残金（空欄）
          cacheData.date,                                // G列：日時（内部データ）
          cacheData.user,                                // H列：購入者（内部データ）
          cacheData.img                                  // I列：画像URL（内部データ）
        ]);
        
        cache.remove(params.cacheKey);
        replyToLine(replyToken, `「${cacheData.sheetName}」へ正常に書き込みが完了しました。`);
        return;
      }
      
      // 【処理B：各種修正ボタンが押された場合】
      if (params.action && params.action.startsWith("edit_")) {
        progress = "修正受付状態への移行処理中";
        
        cache.put(`state_${userId}`, `${params.action}:${params.cacheKey}`, 300);
        
        let guideText = "";
        if (params.action === "edit_date")     guideText = "新しい「日付」を送信してください。\n例: 2026/06/08";
        if (params.action === "edit_store")    guideText = "新しい「店舗名」を送信してください。";
        if (params.action === "edit_amount")   guideText = "新しい「単価（1個あたりの金額）」を数字だけで送信してください。\n例: 150";
        if (params.action === "edit_quantity") guideText = "「個数」を数字だけで送信してください。\n例: 3"; // ★追加：個数ガイド
        if (params.action === "edit_items")    guideText = "新しい「品目・メモ」を送信してください。";
        
        replyToLine(replyToken, `[修正モード]\n${guideText}`);
        return;
      }
      
      // 【処理C：シート選択ボタンが押された場合】
      replyToCameraAction(replyToken, params.sheetId, params.name);
      return;
    }
    
  } catch (error) {
    console.error(`[失敗]: ${progress}\n${error.toString()}`);
    if (replyToken) {
      replyToLine(replyToken, `エラーが発生しました。\n[タイミング]: ${progress}\n[詳細]: ${error.toString()}\n\nもう一度画像を送信してください。`);
    }
  }
}

/**
 * 操作可能シート一覧をボタンで送る
 */
function sendSheetChoices(userId, replyToken, titleText) {
  const masterSheet = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID).getActiveSheet();
  const data = masterSheet.getDataRange().getValues();
  const items = [];
  const rowsToDelete = [];
  
  for (let i = 1; i < data.length; i++) {
    const sheetId = data[i][1];
    try {
      const file = DriveApp.getFileById(sheetId);
      if (file.isTrashed()) { rowsToDelete.push(i + 1); continue; }
    } catch (e) { rowsToDelete.push(i + 1); continue; }
    
    if (data[i][0] === userId) {
      items.push({
        "type": "action",
        "action": {
          "type": "postback",
          "label": data[i][2], 
          "data": `sheetId=${sheetId}&name=${encodeURIComponent(data[i][2])}`, 
          "displayText": `「${data[i][2]}」を選択しました`
        }
      });
    }
  }
  
  if (rowsToDelete.length > 0) {
    for (let r = rowsToDelete.length - 1; r >= 0; r--) { masterSheet.deleteRow(rowsToDelete[r]); }
  }
  if (items.length === 0) {
    replyToLine(replyToken, "まだ操作できるシートがありません。\n「新規作成 [シート名]」と送信して新しく作ってください。");
    return;
  }
  
  const payload = {
    "replyToken": replyToken,
    "messages": [{ "type": "text", "text": titleText, "quickReply": { "items": items } }]
  };
  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", {
    "method": "post", "contentType": "application/json", "headers": { "Authorization": "Bearer " + LINE_ACCESS_TOKEN }, "payload": JSON.stringify(payload)
  });
}

/**
 * 登録前の最終確認および【修正選択ボタン】を送信する
 */
function sendConfirmReply(replyToken, messageText, cacheKey) {
  const payload = {
    "replyToken": replyToken,
    "messages": [{
      "type": "text",
      "text": messageText,
      "quickReply": {
        "items": [
          { "type": "action", "action": { "type": "postback", "label": "この内容で登録", "data": `action=insert&cacheKey=${cacheKey}`, "displayText": "登録します" } },
          { "type": "action", "action": { "type": "postback", "label": "個数を修正", "data": `action=edit_quantity&cacheKey=${cacheKey}` } }, //追加：個数修正ボタン
          { "type": "action", "action": { "type": "postback", "label": "単価を修正", "data": `action=edit_amount&cacheKey=${cacheKey}` } },
          { "type": "action", "action": { "type": "postback", "label": "品名を修正", "data": `action=edit_items&cacheKey=${cacheKey}` } },
          { "type": "action", "action": { "type": "postback", "label": "店舗を修正", "data": `action=edit_store&cacheKey=${cacheKey}` } }
        ]
      }
    }]
  };
  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", {
    "method": "post", "contentType": "application/json", "headers": { "Authorization": "Bearer " + LINE_ACCESS_TOKEN }, "payload": JSON.stringify(payload)
  });
}

function replyToCameraAction(replyToken, sheetId, sheetName) {
  const decodedName = decodeURIComponent(sheetName);
  const payload = {
    "replyToken": replyToken,
    "messages": [{
      "type": "text",
      "text": `「${decodedName}」への記録準備ができました。\n下のボタンから画像を送ってください。`,
      "quickReply": {
        "items": [
          { "type": "action", "action": { "type": "cameraRoll", "label": "アルバムから選択" } },
          { "type": "action", "action": { "type": "camera", "label": "カメラを起動" } }
        ]
      }
    }]
  };
  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", {
    "method": "post", "contentType": "application/json", "headers": { "Authorization": "Bearer " + LINE_ACCESS_TOKEN }, "payload": JSON.stringify(payload)
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

/**
 * ========================================================================
 * 【Life Navigator Prime：doPost.gs (Ver 12.3.0)】
 * ========================================================================
 * [機能・修正履歴]
 * 1. JSONパースエラーの完全修正: Base64生データ受信時の try...catch 救済措置を復旧。
 * 2. 収入データ取り込み機能の完全復活: 給与明細とレシートのハイブリッド自動判別を再統合。
 * 3. 【税抜統一ロジック】: 税込/税抜をAIが自動判別し、全商品を税抜価格に統一変換。
 * 4. 【消費税独立行】: レシート配列の最後に消費税の合計額を自律的に追加する機能を実装。
 * 5. 指定モデル: gemini-3.5-flash を適用。
 * ========================================================================
 */

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("通信ログ検証");
  if (!logSheet) {
    logSheet = ss.insertSheet("通信ログ検証");
    logSheet.appendRow(["受信日時", "ステータス", "データサイズ", "受信データ"]);
  }
  
  try {
    // 1. データ受信
    var receivedData = (e.postData && e.postData.contents) ? e.postData.contents : JSON.stringify(e.parameter);
    
    // 2. Base64生データの救済パース（エラー回避シールド復旧）
    var parsedData;
    try {
      parsedData = JSON.parse(receivedData);
    } catch(err) {
      // JSONではない生テキスト（Base64画像文字列）が送られた場合の救済措置
      parsedData = { base64Data: receivedData };
    }
    var base64Image = (parsedData.base64Data || parsedData.image || parsedData.Base64Data || receivedData).replace(/^data:image\/[a-z]+;base64,/, "");
    
    // 通信ログへ受信記録を残す
    logSheet.appendRow([new Date(), "受信", receivedData.length, receivedData.substring(0, 200)]);
    
    // 3. APIキー取得
    var configSheet = ss.getSheetByName("設定管理") || ss.getSheetByName("マスタ設定");
    var apiKey = "";
    if (configSheet) {
      var data = configSheet.getDataRange().getValues();
      for (var i = 0; i < data.length; i++) {
        if (data[i][0] === "GEMINI_API_KEY") {
          apiKey = data[i][1];
          break;
        }
      }
    }
    if (!apiKey) throw new Error("APIキー未設定");
    
    // 4. シートとヘッダーの動的取得（収入・支出の両方）
    var salarySheet = ss.getSheetByName("収入データ");
    var receiptSheet = ss.getSheetByName("支出データ");
    
    if (!salarySheet || !receiptSheet) {
      throw new Error("「収入データ」または「支出データ」シートが見つかりません。");
    }
    
    var salaryHeaders = salarySheet.getRange(1, 1, 1, salarySheet.getLastColumn()).getValues()[0];
    var receiptHeaders = receiptSheet.getRange(1, 1, 1, receiptSheet.getLastColumn()).getValues()[0];
    
    // 5. AIプロンプトの動的構築（ハイブリッド判定 ＆ 税抜統一ロジック）
    var prompt = "あなたはプロの経理アシスタントです。送信画像を解析し、以下のルールで純粋なJSON(配列形式)のみを出力してください。\n" +
      "1. 給与明細・賞与明細の場合: \"dataType\": \"salary\" を含めた単一のオブジェクトを配列に入れて出力すること。\n" +
      "2. レシート・領収書の場合: 商品ごとにオブジェクトを分割した配列とし、各アイテムに \"dataType\": \"receipt\" を含めること。\n" +
      "※【重要・税計算ルール】レシートの価格が「税込」か「税抜」かを判別し、全商品の金額を【税抜価格】に統一して出力すること（税込表記の場合は、軽減税率なども考慮し消費税分を逆算して税抜金額とする）。\n" +
      "※さらに、配列の最後には必ず消費税の独立行を追加し、品名キーの値を「消費税（合計）」とし、金額キーにレシート全体の消費税額を出力すること。\n" +
      "3. 抽出キー名は以下のシートヘッダーと完全に一致させること:\n" +
      "   - salaryの場合: " + salaryHeaders.join(", ") + "\n" +
      "   - receiptの場合: " + receiptHeaders.join(", ") + "\n" +
      "4. 金額はカンマや円マークを除去した数値のみ出力すること。日付はレシートや明細に印字されている日時を優先すること。";

    // 6. 指定モデル gemini-3.5-flash の適用（指数バックオフ付）
    var apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key=" + apiKey;
    var payload = { "contents": [{ "parts": [{"text": prompt}, {"inlineData": {"mimeType": "image/jpeg", "data": base64Image}}] }] };
    
    var response;
    var success = false;
    for (var retry = 0; retry < 3; retry++) {
      response = UrlFetchApp.fetch(apiUrl, {
        "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true
      });
      if (response.getResponseCode() === 200) {
        success = true;
        break;
      }
      Utilities.sleep(Math.pow(2, retry) * 1000);
    }
    
    if (!success) throw new Error("API Error: " + response.getContentText());
    
    // 7. パースとマークダウン除去
    var responseText = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text.replace(/^```(json)?|```$/gm, "").trim();
    var resultsArray = JSON.parse(responseText);
    
    // オブジェクト単体で返ってきた場合の安全装置
    if (!Array.isArray(resultsArray)) {
      resultsArray = [resultsArray];
    }
    
    // 8. 動的マッピングとシートへの書き込み
    resultsArray.forEach(function(itemData) {
      // dataTypeによって対象シートと対象ヘッダーを切り替え
      var targetSheet = itemData.dataType === "salary" ? salarySheet : receiptSheet;
      var targetHeaders = itemData.dataType === "salary" ? salaryHeaders : receiptHeaders;
      
      var rowData = targetHeaders.map(function(header) {
        var val = itemData.hasOwnProperty(header) ? itemData[header] : "";
        // 「円」やカンマの除去
        if (typeof val === 'string') val = val.replace(/[円,]/g, "");
        // 日付フォールバック（空白の場合のみ現在時刻を補完）
        if ((val === "" || val === null) && (header === "登録日時" || header === "日時" || header === "日付" || header === "購入日" || header === "支給日")) return new Date();
        return val;
      });
      targetSheet.appendRow(rowData);
    });
    
    logSheet.appendRow([new Date(), "成功", responseText.length, "シートへの動的書き込み完了"]);
    return ContentService.createTextOutput(JSON.stringify({"status": "success"})).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    logSheet.appendRow([new Date(), "エラー", "", error.message]);
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": error.message})).setMimeType(ContentService.MimeType.JSON);
  }
}
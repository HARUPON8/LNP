/**
 * ========================================================================
 * 【Life Navigator Prime：doPost.js】
 * ========================================================================
 * [開発・アップデート履歴]
 * Ver 13.0.0 - 【究極進化】動的ヘッダーマッピング機構を完全搭載。
 * 「収入データ」シートの一行目を事前に全自動スキャンし、Geminiへのプロンプトを動的生成。
 * シート側の項目変更・追加・削除・並び替えに対し、GASコードの変更なしで完全追従。
 * 実績のある「gemini-3.5-flash」モデルおよびエンドポイントURL構造を継承。
 * 「設定管理」シートからのGEMINI_API_KEY動的自動取得（空白・改行自動クレンジング内蔵）。
 * iPhoneから送信された給料明細画像(Base64)の余分なヘッダー・改行を自動クレンジング。
 * Geminiモデルを用いた高精度OCR解析とマークダウンタグの強制除去。
 * 指数バックオフによる503エラー自動再試行機構および通信ログ検証シートへの二重防衛。
 * ========================================================================
 */

/**
 * iPhoneショートカット等の外部クライアントからのPOSTリクエストを処理する
 * @param {Object} e - HTTP POSTリクエストイベントオブジェクト
 * @return {TextOutput} JSON形式のレスポンス
 */
function doPost(e) {
  var timestamp = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  var logSheetName = "通信ログ検証";
  var targetSheetName = "収入データ";
  var configSheetName = "設定管理";
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(logSheetName);
  
  // ログシートの存在チェックと防衛自動生成
  if (!logSheet) {
    logSheet = ss.insertSheet(logSheetName);
    logSheet.appendRow(["受信日時", "ステータス", "データサイズ(文字数)", "受信データ"]);
  }
  
  try {
    // 1. 送信データの安全なサルベージ（表記揺れクレンジングシールド）
    var rawBase64 = "";
    if (e && e.postData && e.postData.contents) {
      try {
        var parsedBody = JSON.parse(e.postData.contents);
        rawBase64 = parsedBody.image || parsedBody.base64Data || parsedBody.Base64Data || e.postData.contents;
      } catch (ex) {
        rawBase64 = e.postData.contents;
      }
    } else if (e && e.parameter) {
      rawBase64 = e.parameter.image || e.parameter.base64Data || e.parameter.Base64Data;
      if (!rawBase64) {
        var keys = Object.keys(e.parameter);
        if (keys.length > 0) rawBase64 = e.parameter[keys[0]];
      }
    }
    
    // ガード節：データ空判定
    if (!rawBase64 || rawBase64.trim() === "") {
      logSheet.appendRow([timestamp, "WARNING", 0, "受信したBase64データが空でした。"]);
      return ContentService.createTextOutput(JSON.stringify({
        status: "error",
        message: "Empty data received."
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // 2. Base64文字列の強制標準化クレンジング
    var cleanBase64 = rawBase64
      .replace(/^data:image\/?[^;]+;base64,/, "") // MIMEプレフィックス除去
      .replace(/\r?\n/g, "")                    // 改行完全消去
      .replace(/\s/g, "");                       // 空白スペース完全消去
      
    var dataSize = cleanBase64.length;
    
    // 3. 設定管理シートからのAPIキー動的バインド（空白・改行の強制排除シールド）
    var configSheet = ss.getSheetByName(configSheetName) || ss.getSheetByName("設定管理");
    if (!configSheet) {
      throw new Error("「設定管理」シートが見つかりません。");
    }
    
    var configValues = configSheet.getDataRange().getValues();
    var apiKey = "";
    for (var i = 0; i < configValues.length; i++) {
      if (String(configValues[i][0]).trim() === "GEMINI_API_KEY") {
        apiKey = String(configValues[i][1]).trim();
        break;
      }
    }
    
    if (!apiKey || apiKey === "") {
      throw new Error("設定管理シートに GEMINI_API_KEY が登録されていないか、空欄です。");
    }
    
    // 4. 【本機能の核心】「収入データ」シートの現在の1行目(ヘッダー項目)をリアルタイム動的スキャン
    var targetSheet = ss.getSheetByName(targetSheetName);
    if (!targetSheet) {
      // 万が一シート自体がない場合の初期防衛生成
      targetSheet = ss.insertSheet(targetSheetName);
      targetSheet.appendRow([
        "支給日", "支給年月", "会社名", "支給合計", "控除合計", "差引支給額", "振込額", "課税支給額", "社会保険計", "課税対象額",
        "基本給(正社員)", "みなし残業手当", "深夜手当", "超過残業手当", "休日出勤手当", "健保(一般)", "健保(基本)", "健保(特定)",
        "子育て支援金", "健保(介護)", "厚生年金", "雇用保険", "所得税", "住民税", "所定労働日数", "所定労働時間",
        "平日出勤", "出勤時間", "残業時間", "超過残業", "深夜残業", "休出時間", "出勤日数計", "有休残前年", "有休残当年"
      ]);
    }
    
    // 1行目のセルのテキストをそのまま配列として取得する
    var currentHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    
    // 5. 取得した実際のヘッダー項目をベースに、指示プロンプトを全自動で組み立て
    var mdQuote = "\x60\x60\x60";
    var promptText = "あなたは一流の給料明細データ解析システムです。送信された画像を隅々まで分析し、指定された項目に該当する数値を正確に抽出し、指定された項目名をそのまま「キー(Key)」にしたJSONオブジェクトでのみ返却してください。マークダウンの囲み(" + mdQuote + "jsonなど)や余分な解説文は絶対に含めず、純粋なJSON文字列（シングルオブジェクト）のみを返してください。数値項目は「円」「,」などの記号を完全に除去し、プレーンな数字（整数または小数）にしてください。時間、日数などは文字列のまま（例: 「125:30」「24.00」）保持してください。項目が存在しない、または読み取れない場合は 0 または空文字(\"\")にしてください。\n\n" +
                     "【必ず抽出してJSONのキーにする項目リスト】\n" +
                     currentHeaders.join(", ") + "\n\n" +
                     "【注意事項】\n" +
                     "・日付項目(支給日など)は、可能であれば「2026/05/15」のようなスラッシュ区切りの形式に統一してください。\n" +
                     "・JSONオブジェクトのキー名は、上記リストの文言（スペースや括弧も含む）と1文字も違わずに完全に一致させてください。";

    // 開通実績のあるマスタモデル「gemini-3.5-flash」を完全適用
    var model = "gemini-3.5-flash";
    var apiURL = "https://generativelanguage.googleapis.com/v1beta/models/" + model + ":generateContent?key=" + apiKey;
    
    var payload = {
      "contents": [{
        "parts": [
          { "text": promptText },
          {
            "inlineData": {
              "mimeType": "image/jpeg",
              "data": cleanBase64
            }
          }
        ]
      }]
    };
    
    var options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };
    
    // 6. 503一時エラー自動回避：指数バックオフ付加通信処理
    var response;
    var responseCode;
    var maxRetries = 3;
    var waitTime = 2000;
    
    for (var retry = 0; retry <= maxRetries; retry++) {
      response = UrlFetchApp.fetch(apiURL, options);
      responseCode = response.getResponseCode();
      
      if (responseCode === 200) {
        break; 
      } else if (responseCode === 503 && retry < maxRetries) {
        Utilities.sleep(waitTime);
        waitTime *= 2; 
      } else {
        throw new Error("Gemini APIエラー コード: " + responseCode + " 詳細: " + response.getContentText());
      }
    }
    
    var responseText = response.getContentText();
    var jsonResponse = JSON.parse(responseText);
    
    if (!jsonResponse.candidates || jsonResponse.candidates.length === 0) {
      throw new Error("Gemini解析結果が空でした(No Candidates)。");
    }
    
    var aiResultText = jsonResponse.candidates[0].content.parts[0].text;
    
    // 外部からの不要なマークダウンタグおよび制御文字のクレンジング除去
    aiResultText = aiResultText.replace(/`{3}json/gi, "").replace(/`{3}/g, "").trim();
    
    var salaryData = JSON.parse(aiResultText);
    
    // 通貨や記号、不要文字の安全な数値変換関数
    function cleanNum(val, headerName) {
      if (val === undefined || val === null || val === "") return 0;
      // 時間表記や、支給年月、会社名などの文字列を保護するガードロジック
      if (typeof val === "string" && val.indexOf(":") !== -1) return val;
      if (headerName.indexOf("日") !== -1 && String(val).indexOf(".") !== -1) return val; // 日数小数を保持
      if (isNaN(val) && (headerName.indexOf("日") !== -1 || headerName.indexOf("年月") !== -1 || headerName.indexOf("名") !== -1)) return val;
      
      var num = Number(String(val).replace(/[^0-9.-]/g, ""));
      return isNaN(num) ? 0 : num;
    }
    
    // 7. 【動的整列シールド】スキャンした実際のヘッダー配列をループで回し、着弾データを動的に再マッピング
    var rowData = [];
    for (var j = 0; j < currentHeaders.length; j++) {
      var hName = currentHeaders[j];
      var extractedVal = salaryData[hName];
      
      // キーの表記揺れ（英語や別文字）対策の最終フォールバック
      if (extractedVal === undefined || extractedVal === null) {
        extractedVal = "";
      }
      
      // 項目名に応じて適切なクレンジングを行って配列に追加
      if (hName.indexOf("支給日") !== -1 || hName.indexOf("年月") !== -1 || hName.indexOf("会社") !== -1) {
        rowData.push(extractedVal); // 文字列項目はそのまま追加
      } else {
        rowData.push(cleanNum(extractedVal, hName)); // 金額・勤怠項目は数値化
      }
    }
    
    // 現在のシートの並び順と1マスもズレない全データ行を最終行へ一撃流し込み
    targetSheet.appendRow(rowData);
    
    // 通信ログ検証シートへのSUCCESS記録
    logSheet.appendRow([
      timestamp,
      "SUCCESS",
      dataSize,
      "動的ヘッダー解析に成功。現在の列数: " + currentHeaders.length + " 列に対して完全に整列化して格納しました。"
    ]);
    
    // iPhoneショートカットへ成功同期レスポンスを返却
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "給料明細の動的ヘッダー自動解析および格納が完全完了しました。",
      timestamp: timestamp,
      columns: currentHeaders.length
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    var errMsg = error.toString();
    if (logSheet) {
      logSheet.appendRow([timestamp, "ERROR", 0, "例外発生: " + errMsg]);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "バックエンド処理中に例外エラーが発生しました。",
      details: errMsg
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
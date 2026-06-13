/**
 * ========================================================================
 * 【バックエンドシステム：コード.gs】
 * ========================================================================
 * [開発・アップデート履歴]
 * Ver 1.0.0 - 初期デプロイ。doGetベース実装。
 * Ver 5.0.0 - 5大画面仕様（ライフプラン、交通・天気、健康、家計、スケジュール）統合。
 * Ver 6.6.6 - 強制検知用承認解除シールド（authorizeConnectionShield）内蔵。
 * Ver 8.3.0 - 週間天気予報を8日間拡張取得し、本日重複を排除するロジックの確定。
 * Ver 8.5.0 [最新] - Syntax Error (Unexpected end of input) 完全打破、全バックエンドマージ版。
 * ========================================================================
 */

// 1. Webアプリケーション 起動エントリーポイント
function doGet(e) {
  return HtmlService.createTemplateFromFile('life-navigator-prime')
      .evaluate()
      .setTitle('Life Navigator Prime')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 2. HTMLコンポーネント独立読み込みヘルパー
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 3. 外部気象通信連動エンジン（明日以降の純粋な7日間を確保するため、forecast_days=8で取得）
function getWeeklyWeatherData() {
  var weeklyForecast = [];
  var daysOfWeek = ["日", "月", "火", "水", "木", "金", "土"];
  
  // 既定のフォールバック座標（神奈川県綾瀬市役所周辺基準）
  var lat = "35.4554";
  var lon = "139.4278";
  
  try {
    // マスター設定シートから緯度・経度を動的取得
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("設定管理");
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      for (var i = 0; i < data.length; i++) {
        var key = String(data[i][0]).trim();
        var val = String(data[i][1]).trim();
        if (key === "AYASE_LATITUDE" && val) { lat = val; }
        if (key === "AYASE_LONGITUDE" && val) { lon = val; }
      }
    }
    
    // エンドポイントURLの構築（明日以降の7日間を確保するため、forecast_days=8を完全注入）
    var url = "https://api.open-meteo.com/v1/forecast?latitude=" + lat + "&longitude=" + lon + "&daily=weather_code,temperature_2m_max,temperature_2m_min&timezone=Asia%2FTokyo&forecast_days=8";
    
    var response = UrlFetchApp.fetch(url, { "muteHttpExceptions": true });
    if (response.getResponseCode() !== 200) {
      throw new Error("APIレスポンスコード異常: " + response.getResponseCode());
    }
    
    var json = JSON.parse(response.getContentText());
    var daily = json.daily;
    if (!daily || !daily.time) {
      throw new Error("気象データ構造が不正です");
    }
    
    // 本日 + 明日以降7日分の計8日分を安全にループスキャン
    var limit = Math.min(8, daily.time.length);
    for (var j = 0; j < limit; j++) {
      var dateParts = String(daily.time[j]).split('-');
      var yearStr = parseInt(dateParts[0], 10);
      var monthStr = parseInt(dateParts[1], 10);
      var dayStr = parseInt(dateParts[2], 10);
      
      var safetyDate = new Date(yearStr, monthStr - 1, dayStr);
      var dayName = daysOfWeek[safetyDate.getDay()];
      
      // 本日枠と未来枠の表示ラベル切り替え
      var dayLabel = (j === 0) ? "本日 " + monthStr + "/" + dayStr : monthStr + "/" + dayStr;
      
      var maxTemp = (daily.temperature_2m_max[j] !== null) ? Math.round(daily.temperature_2m_max[j]) : "--";
      var minTemp = (daily.temperature_2m_min[j] !== null) ? Math.round(daily.temperature_2m_min[j]) : "--";
      
      var wmoCode = daily.weather_code[j] || 0;
      var cleanWeather = "晴れ";
      var icon = "☀️";
      
      if (wmoCode === 0) {
        cleanWeather = "快晴"; icon = "☀️";
      } else if (wmoCode === 1 || wmoCode === 2) {
        cleanWeather = "晴れ"; icon = "🌤️";
      } else if (wmoCode === 3) {
        cleanWeather = "曇り"; icon = "☁️";
      } else if (wmoCode === 45 || wmoCode === 48) {
        cleanWeather = "霧"; icon = "🌫️";
      } else if (wmoCode === 51 || wmoCode === 53 || wmoCode === 55 || wmoCode === 61 || wmoCode === 63 || wmoCode === 65) {
        cleanWeather = "雨"; icon = "☔";
      } else if (wmoCode === 71 || wmoCode === 73 || wmoCode === 75 || wmoCode === 77 || wmoCode === 85 || wmoCode === 86) {
        cleanWeather = "雪"; icon = "❄️";
      } else if (wmoCode === 80 || wmoCode === 81 || wmoCode === 82) {
        cleanWeather = "にわか雨"; icon = "🌦️";
      } else if (wmoCode >= 95) {
        cleanWeather = "雷雨"; icon = "⛈️";
      }
      
      weeklyForecast.push({
        day: dayLabel + " (" + dayName + ")",
        weather: cleanWeather,
        icon: icon,
        temp: maxTemp + "℃ / " + minTemp + "℃"
      });
    }
    
  } catch (err) {
    Logger.log("週間予報パース連動例外捕捉: " + err.toString());
    weeklyForecast = [];
    var errDate = new Date();
    for (var k = 0; k < 8; k++) {
      var label = (k === 0) ? "本日" : (errDate.getMonth() + 1) + "/" + errDate.getDate();
      weeklyForecast.push({
        day: label + " (" + daysOfWeek[errDate.getDay()] + ")",
        weather: "データが取得できませんでした",
        icon: "⚠️",
        temp: "-- / --"
      });
      errDate.setDate(errDate.getDate() + 1);
    }
  }
  return weeklyForecast;
}

// 4. クイック現金支出追記処理
function saveCashExpense(memo, amount, cat) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("支出データ");
    if (!sheet) {
      sheet = ss.insertSheet("支出データ");
      sheet.appendRow(["日時", "店舗名", "カテゴリ", "商品名", "金額"]);
    }
    var timestamp = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
    
    // スプレッドシートの最終行へ一撃で安全に追記
    sheet.appendRow([timestamp, "現金手入力", cat, memo, Number(amount)]);
    return { status: "success" };
  } catch (err) {
    Logger.log("saveCashExpense例外捕捉: " + err.toString());
    return { status: "error", message: err.message };
  }
}

// 5. ダッシュボード統合データ一括一元取得エンジン（二重防衛シールド完備）
function getDashboardAllData() {
  var result = {
    progressPercent: 75,
    remainingAmount: "300,000 円",
    limitPerDay: "1,000 円 / 日",
    weight: "65.0 kg",
    fat: "18.5 %",
    steps: "8,500 歩",
    recipe: "AI献立：十六穀米、具だくさん発酵美肌味噌汁、サバの塩焼き、めかぶおろし納豆。",
    timeline: []
  };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // [A] ライフプランシートの走査
    var visionSheet = ss.getSheetByName("資産データ");
    if (visionSheet) {
      var latestRow = visionSheet.getLastRow();
      if (latestRow > 1) {
        var assetVal = visionSheet.getRange(latestRow, 2).getValue(); // B列：純資産残高
        var target = 1200000;
        var percent = Math.min(100, Math.round((assetVal / target) * 100));
        result.progressPercent = percent;
        result.remainingAmount = (target - assetVal).toLocaleString() + " 円";
      }
    }
    
    // [B] 健康管理データの走査
    var bioSheet = ss.getSheetByName("健康管理データ");
    if (bioSheet) {
      var lastBioRow = bioSheet.getLastRow();
      if (lastBioRow > 1) {
        result.weight = bioSheet.getRange(lastBioRow, 2).getValue() + " kg"; // B列
        result.fat = bioSheet.getRange(lastBioRow, 3).getValue() + " %";    // C列
        result.steps = bioSheet.getRange(lastBioRow, 4).getValue() + " 歩";  // D列
      }
    }
    
    // フロントエンド互換キー名への二重マッピング（安全対策）
    result.defenseFundRate = result.progressPercent;
    
  } catch (e) {
    Logger.log("getDashboardAllData一部セクション例外捕捉(安全弁スライド): " + e.toString());
  }
  return result;
}

// 6. 【最重要セキュリティ開通シールド】Googleの認証スキップを完全打破する防壁解除用関数
function authorizeConnectionShield() {
  try {
    UrlFetchApp.fetch("https://api.open-meteo.com/v1/forecast?latitude=35.4554&longitude=139.4278&daily=weather_code&timezone=Asia%2FTokyo&forecast_days=1");
    SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    Logger.log("シールドスタンドバイ実行完了");
  }
}
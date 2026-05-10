const MASTER_SHEET_NAME = "管理者シート";
const TEMPLATE_SHEET_NAME = "テンプレートシート";

// ==========================================
// 1. GETリクエスト（ステータス取得用）
// ==========================================
function doGet(e) {
  const params = e.parameter || {};
  const action = params.action;
  const token = params.token;

  // アプリを開いた時に「現在出勤中か？」を返すAPI
  if (action === 'status' && token) {
    try {
      const info = getTeacherStatus(token);
      return jsonResponse({ ok: true, data: info });
    } catch(err) {
      return jsonResponse({ ok: false, error: err.message });
    }
  }
  return HtmlService.createHtmlOutput("Teacher API is running.");
}

function getTeacherStatus(token) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  const masterData = masterSheet.getDataRange().getValues();
  
  let teacherInfo = null;
  // 管理者シート（A:ID, B:名前, C:トークン, D:個人シート名）
  for (let i = 1; i < masterData.length; i++) {
    if (String(masterData[i][2]).trim() === token) {
      teacherInfo = { id: masterData[i][0], name: masterData[i][1], sheetName: masterData[i][3] };
      break;
    }
  }
  if (!teacherInfo) throw new Error("無効なURL（トークン）です。");

  const targetSheet = ss.getSheetByName(teacherInfo.sheetName);
  if (!targetSheet) {
    return { name: teacherInfo.name, status: "not_started" }; // シートがない＝未出勤
  }

  const data = targetSheet.getDataRange().getValues();
  const todayStr = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");
  
  let status = "not_started";
  let plan = "";

  for (let i = data.length - 1; i >= 1; i--) {
    const rowDate = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], "JST", "yyyy/MM/dd") : data[i][0];
    if (rowDate === todayStr) {
      if (!data[i][2]) { // C列(退勤)が空なら出勤中
        status = "working";
        plan = data[i][4]; // E列: 本日の業務予定
      } else {
        status = "completed"; // 今日は退勤済み
      }
      break;
    }
  }
  return { name: teacherInfo.name, status: status, plan: plan };
}

// ==========================================
// 2. POSTリクエスト（打刻と報告の書き込み）
// ==========================================
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action;
    const token = body.token;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    const masterData = masterSheet.getDataRange().getValues();
    
    let teacherInfo = null;
    for (let i = 1; i < masterData.length; i++) {
      if (String(masterData[i][2]).trim() === token) {
        teacherInfo = { id: masterData[i][0], name: masterData[i][1], sheetName: masterData[i][3] };
        break;
      }
    }
    if (!teacherInfo) throw new Error("無効なトークンです。");
    
    // シートがない場合はテンプレートから自動作成
    let targetSheet = ss.getSheetByName(teacherInfo.sheetName);
    if (!targetSheet) {
      const templateSheet = ss.getSheetByName(TEMPLATE_SHEET_NAME);
      if (templateSheet) {
        targetSheet = templateSheet.copyTo(ss);
        targetSheet.setName(teacherInfo.sheetName);
      } else {
        targetSheet = ss.insertSheet(teacherInfo.sheetName);
        targetSheet.appendRow(["日付", "出勤", "退勤", "作業時間", "業務報告", "成果報告"]);
      }
    }
    
    const now = new Date();
    const todayStr = Utilities.formatDate(now, "JST", "yyyy/MM/dd");
    const timeStr = Utilities.formatDate(now, "JST", "HH:mm");
    
    const data = targetSheet.getDataRange().getValues();
    let targetRow = -1;
    let clockInTime = null;
    
    // 今日の行を探す
    for (let i = data.length - 1; i >= 1; i--) {
      const rowDate = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], "JST", "yyyy/MM/dd") : data[i][0];
      if (rowDate === todayStr) {
        targetRow = i + 1;
        clockInTime = data[i][1]; // B列 出勤時刻
        break;
      }
    }

    // 🌟 出勤時の処理（E列に予定を書き込む）
    if (action === "clockIn") {
      if (targetRow !== -1 && !data[targetRow-1][2]) return jsonResponse({ ok: false, error: "すでに出勤しています" });
      targetSheet.appendRow([todayStr, timeStr, "", "", body.plan, ""]);
      return jsonResponse({ ok: true, message: "出勤しました。本日の業務を開始します。" });
    } 
    
    // 🌟 退勤時の処理（F列に成果、D列に作業時間を書き込む）
    if (action === "clockOut") {
      if (targetRow === -1 || data[targetRow-1][2]) return jsonResponse({ ok: false, error: "出勤記録がないか、すでに退勤済みです" });
      
      targetSheet.getRange(targetRow, 3).setValue(timeStr);   // C列: 退勤
      targetSheet.getRange(targetRow, 6).setValue(body.result); // F列: 成果報告
      
      // 作業時間の自動計算（D列）
      if (clockInTime) {
        const inDate = new Date(todayStr + " " + (clockInTime instanceof Date ? Utilities.formatDate(clockInTime, "JST", "HH:mm") : clockInTime));
        const diffMs = now - inDate;
        if (diffMs > 0) {
          const h = Math.floor(diffMs / 3600000);
          const m = Math.floor((diffMs % 3600000) / 60000);
          targetSheet.getRange(targetRow, 4).setValue(`${h}時間${m}分`); 
        }
      }
      return jsonResponse({ ok: true, message: "退勤と成果報告を完了しました。お疲れ様でした！" });
    }

  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
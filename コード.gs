const MASTER_SHEET_NAME = "管理者シート";
const TEMPLATE_SHEET_NAME = "テンプレートシート";

// ==========================================
// ★ 管理者用：ドキュメントアクセス確認関数
// ==========================================
function checkAllDocsAccess() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    console.error("管理者シートが見つかりません。");
    return;
  }
  
  const data = masterSheet.getDataRange().getValues();
  let successCount = 0;
  let errorCount = 0;
  
  console.log("=== ドキュメントアクセス確認開始 ===");
  
  for (let i = 1; i < data.length; i++) {
    const teacherName = data[i][1];
    const docUrl = data[i][4]; // E列
    
    if (!docUrl) {
      console.log(`[スキップ] ${teacherName} 先生: URLが設定されていません。`);
      continue;
    }
    
    try {
      const docStr = String(docUrl).trim();
      const doc = docStr.startsWith('http') ? DocumentApp.openByUrl(docStr) : DocumentApp.openById(docStr);
      const name = doc.getName();
      console.log(`[OK] ${teacherName} 先生: アクセス成功 (ドキュメント名: ${name})`);
      successCount++;
    } catch (err) {
      console.error(`[エラー] ${teacherName} 先生: アクセス失敗 (${err.message}) URL: ${docUrl}`);
      errorCount++;
    }
  }
  
  console.log(`=== 確認終了 (成功: ${successCount}, 失敗: ${errorCount}) ===`);
}

// ==========================================
// 1. GETリクエスト（ステータス＆統計データ取得）
// ==========================================
function doGet(e) {
  const params = e.parameter || {};
  const action = params.action;
  const token = params.token;

  if (action === 'status' && token) {
    try {
      const info = getTeacherData(token);
      return jsonResponse({ ok: true, data: info });
    } catch(err) {
      return jsonResponse({ ok: false, error: err.message });
    }
  }
  return HtmlService.createHtmlOutput("Teacher API is running.");
}

function getTeacherData(token) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  const masterData = masterSheet.getDataRange().getValues();
  
  let teacherInfo = null;
  for (let i = 1; i < masterData.length; i++) {
    if (String(masterData[i][2]).trim() === token) {
      teacherInfo = { 
        id: masterData[i][0], 
        name: masterData[i][1], 
        sheetName: masterData[i][3],
        docUrl: masterData[i][4] 
      };
      break;
    }
  }
  if (!teacherInfo) throw new Error("無効なURL（トークン）です。");

  const targetSheet = ss.getSheetByName(teacherInfo.sheetName);
  const emptyStats = {
      daily: { labels: [], values: [] },
      weekly: { labels: [], values: [] },
      monthly: { labels: [], values: [] },
      history: [],
      today: "0時間0分"
  };

  if (!targetSheet) {
    return { name: teacherInfo.name, status: "not_started", plan: "", stats: emptyStats };
  }

  const data = targetSheet.getDataRange().getValues();
  const todayStr = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");
  
  let status = "not_started";
  let currentPlan = "";

  const dailyMap = {};
  const weeklyMap = {};
  const monthlyMap = {};
  const historyList = [];
  let todayMs = 0;

  const getMonday = (d) => {
    let day = d.getDay(), diff = d.getDate() - day + (day == 0 ? -6 : 1);
    return new Date(d.getFullYear(), d.getMonth(), diff);
  };

  for (let i = 1; i < data.length; i++) {
    const rowDateStr = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], "JST", "yyyy/MM/dd") : data[i][0];
    if (!rowDateStr) continue;
    
    const inTimeStr = data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], "JST", "HH:mm") : data[i][1];
    const outTimeStr = data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], "JST", "HH:mm") : data[i][2];
    const durationStr = data[i][3]; 
    const plan = data[i][4];
    const result = data[i][5];

    if (rowDateStr === todayStr) {
      if (!outTimeStr) {
        status = "working";
        currentPlan = plan;
      } else {
        status = "not_started";
      }
    }

    let diffMs = 0;
    if (inTimeStr && outTimeStr) {
       const inDate = new Date(rowDateStr + " " + inTimeStr);
       const outDate = new Date(rowDateStr + " " + outTimeStr);
       diffMs = outDate - inDate;
    }
    
    if (diffMs > 0) {
       if (rowDateStr === todayStr) {
         todayMs += diffMs;
       }
       const dObj = new Date(rowDateStr);
       const dayKey = Utilities.formatDate(dObj, "JST", "MM/dd");
       dailyMap[dayKey] = (dailyMap[dayKey] || 0) + diffMs;

       const monday = getMonday(dObj);
       const weekKey = Utilities.formatDate(monday, "JST", "MM/dd") + "週";
       weeklyMap[weekKey] = (weeklyMap[weekKey] || 0) + diffMs;

       const monthKey = Utilities.formatDate(dObj, "JST", "yyyy/MM");
       monthlyMap[monthKey] = (monthlyMap[monthKey] || 0) + diffMs;
    }

    historyList.unshift({
      date: rowDateStr,
      in: inTimeStr || "---",
      out: outTimeStr || "---",
      time: durationStr || "---",
      plan: plan || "",
      result: result || ""
    });
  }

  const formatChartData = (map) => {
    let labels = Object.keys(map).sort();
    let values = labels.map(k => (map[k]/3600000).toFixed(1));
    return { labels, values };
  };

  const formatTime = (ms) => { return `${Math.floor(ms/3600000)}時間${Math.floor((ms%3600000)/60000)}分`; };

  return { 
    name: teacherInfo.name, 
    status: status, 
    plan: currentPlan,
    stats: {
      daily: formatChartData(dailyMap),
      weekly: formatChartData(weeklyMap),
      monthly: formatChartData(monthlyMap),
      history: historyList.slice(0, 30),
      today: formatTime(todayMs)
    }
  };
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
        teacherInfo = { 
          id: masterData[i][0], 
          name: masterData[i][1], 
          sheetName: masterData[i][3],
          docUrl: masterData[i][4] 
        };
        break;
      }
    }
    if (!teacherInfo) throw new Error("無効なトークンです。");
    
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
    
    for (let i = data.length - 1; i >= 1; i--) {
      const rowDate = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], "JST", "yyyy/MM/dd") : data[i][0];
      if (rowDate === todayStr) {
        targetRow = i + 1;
        clockInTime = data[i][1];
        break; 
      }
    }

    if (action === "clockIn") {
      if (targetRow !== -1 && !data[targetRow-1][2]) return jsonResponse({ ok: false, error: "すでに出勤しています" });
      targetSheet.appendRow([todayStr, timeStr, "", "", body.plan, ""]);
      return jsonResponse({ ok: true, message: "出勤しました。本日の業務を開始します。" });
    } 
    
    if (action === "clockOut") {
      if (targetRow === -1 || data[targetRow-1][2]) return jsonResponse({ ok: false, error: "出勤記録がないか、すでに退勤済みです" });
      
      targetSheet.getRange(targetRow, 3).setValue(timeStr);   
      targetSheet.getRange(targetRow, 6).setValue(body.result); 
      
      let workTimeStr = "---";
      if (clockInTime) {
        const inDate = new Date(todayStr + " " + (clockInTime instanceof Date ? Utilities.formatDate(clockInTime, "JST", "HH:mm") : clockInTime));
        const diffMs = now - inDate;
        if (diffMs > 0) {
          const h = Math.floor(diffMs / 3600000);
          const m = Math.floor((diffMs % 3600000) / 60000);
          workTimeStr = `${h}時間${m}分`;
          targetSheet.getRange(targetRow, 4).setValue(workTimeStr); 
        }
      }

      // ★ Googleドキュメントへの【表形式】での追記処理
      if (teacherInfo.docUrl) {
        try {
          const docUrlStr = String(teacherInfo.docUrl).trim();
          const doc = docUrlStr.startsWith('http') ? DocumentApp.openByUrl(docUrlStr) : DocumentApp.openById(docUrlStr);
          const docBody = doc.getBody();
          
          // スプレッドシートに記録されている「今日の業務内容（Plan）」を取得
          const planFromSheet = data[targetRow-1][4];
          
          // 表（テーブル）のデータを作成（2列×4行）
          const tableData = [
            ["日時", `${todayStr} ${timeStr}`],
            ["作業時間", workTimeStr],
            ["今日の業務内容", planFromSheet || "（記載なし）"],
            ["成果報告", body.result || "（記載なし）"]
          ];
          
          // ドキュメントに表を追加
          const table = docBody.appendTable(tableData);
          
          // 見やすくするために1列目（項目名）の背景色をグレーにし、幅を固定する
          for (let i = 0; i < 4; i++) {
            const headerCell = table.getCell(i, 0);
            headerCell.setBackgroundColor("#f3f3f3");
            headerCell.setWidth(120);
          }
          
          docBody.appendParagraph(""); // 次の記録との隙間を空ける

        } catch (docErr) {
          console.error("ドキュメント書き込みエラー: ", docErr.message);
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

function setupPermissions() {
  DocumentApp.create("権限確認用（すぐに削除してOKです）");
  SpreadsheetApp.getActiveSpreadsheet();
}
// ====================================================================
// doPost関数（生徒用と統合した完成形イメージ）
// ====================================================================
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action;

    // ----- 【生徒用】設定保存 -----
    if (action === 'saveSettings') {
      const msg = saveStudentSettings(body.token, body.examName, body.examDate);
      return jsonResponse({ ok: true, message: msg });
    }

    // ----- 【講師用】出勤・退勤報告 -----
    if (action === 'clockIn' || action === 'clockOut') {
      return jsonResponse(processTeacherAttendance(body));
    }

    return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err && err.message || err) });
  }
}

// ====================================================================
// 講師用の勤怠処理ロジック
// ====================================================================
function processTeacherAttendance(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adminSheet = ss.getSheetByName('管理者シート');
  const adminData = adminSheet.getDataRange().getValues();

  // 1. 講師の認証
  let teacher = null;
  for (let i = 1; i < adminData.length; i++) {
    if (adminData[i][2] === body.token) { // C列がトークン
      teacher = {
        id: adminData[i][0],
        name: adminData[i][1],
        sheetName: adminData[i][3] // D列が個別シート名
      };
      break;
    }
  }
  if (!teacher) return { ok: false, error: 'トークンが無効です。URLを確認してください。' };

  // 2. 個別シートの取得（なければテンプレートから作成）
  let targetSheet = ss.getSheetByName(teacher.sheetName);
  if (!targetSheet) {
    const template = ss.getSheetByName('テンプレートシート');
    if (!template) return { ok: false, error: 'テンプレートシートが見つかりません。' };
    targetSheet = template.copyTo(ss).setName(teacher.sheetName);
  }

  const now = new Date();
  const todayStr = Utilities.formatDate(now, "JST", "yyyy/MM/dd");
  const timeStr = Utilities.formatDate(now, "JST", "HH:mm");

  const data = targetSheet.getDataRange().getValues();
  let targetRow = -1;

  // 今日の記録を探す
  for (let i = data.length - 1; i >= 1; i--) {
    const rowDate = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], "JST", "yyyy/MM/dd") : data[i][0];
    if (rowDate === todayStr) {
      targetRow = i + 1;
      break;
    }
  }

  // --- 出勤処理 ---
  if (body.action === 'clockIn') {
    if (targetRow !== -1 && data[targetRow - 1][1] !== "") {
      return { ok: false, error: '本日はすでに出勤記録があります。' };
    }
    // A:日付, B:出勤(時間), C:退勤, D:作業時間, E:業務報告(予定), F:成果報告
    targetSheet.appendRow([todayStr, timeStr, "", "", body.plannedTask, ""]);
    return { ok: true, message: `${teacher.name} 先生、おはようございます！\n本日の業務予定を記録しました。` };
  }

  // --- 退勤処理 ---
  if (body.action === 'clockOut') {
    if (targetRow === -1) return { ok: false, error: '本日の出勤記録が見つかりません。' };
    if (data[targetRow - 1][2] !== "") return { ok: false, error: '本日はすでに退勤済みです。' };

    // 退勤時間と成果報告を記録
    targetSheet.getRange(targetRow, 3).setValue(timeStr);
    targetSheet.getRange(targetRow, 6).setValue(body.resultTask);

    // 作業時間の計算（C列 - B列）
    const inTimeStr = String(data[targetRow - 1][1]);
    if (inTimeStr) {
      const inParts = inTimeStr.split(':');
      const inDate = new Date(now.getFullYear(), now.getMonth(), now.getDate(), parseInt(inParts[0]), parseInt(inParts[1]));
      const diffMs = now.getTime() - inDate.getTime();
      const hours = Math.floor(diffMs / 3600000);
      const minutes = Math.floor((diffMs % 3600000) / 60000);
      targetSheet.getRange(targetRow, 4).setValue(`${hours}時間${minutes}分`);
    }
    
    return { ok: true, message: `退勤と成果報告を記録しました。\n${teacher.name} 先生、本日もお疲れ様でした！` };
  }
}
/**
 * 代課費計算與管理系統 - Google Apps Script (GAS) 雲端資料庫後台
 * 【按月分表儲存 (總表 + 明細表)】
 *
 * 說明：
 * 1. 依代課日期的月份，自動建立 `YYYY-MM_總表` 與 `YYYY-MM_明細` 兩個分頁。
 * 2. 歷史月份各自獨立，不互相干擾。
 * 3. 行政端帶 `?month=YYYY-MM` 即可精準取得該月資料。
 *
 * ⚠️ 欄位相容性（勿隨意刪欄）：
 *   admin.html 需要 subTeacher / periodsCount / className / classFee / mentorFee / totalFee
 *   generate_voucher.py 另需 rate / payMode / actingDays / isActingMentor / periodsDetail[].paid
 *   總表與明細的欄位順序由下方 SUMMARY_HEADERS / DETAIL_HEADERS 定義，讀寫共用同一份定義。
 */

/**
 * 讀取金鑰（共用密碼）
 *
 * 用途：teacher.html 是公開網頁，網址會寫在原始碼裡任何人都看得到。
 * 若後端不設限，任何人都能用 `?month=all` 抓走全部代課費資料（教師姓名、金額）。
 * 因此「讀取」一律要求帶上金鑰，只有 admin.html（僅在承辦人電腦上）知道。
 *
 * ⚠️ 注意事項：
 * - 「寫入」(doPost) 維持開放，老師才能免登入填報。
 * - 課表 (?action=timetable) 維持開放，teacher.html 需要它顯示課程。
 * - 金鑰外流時，改掉這一行後重新部署即可（網址不變）。
 * - admin.html 不會寫死金鑰，由承辦人在齒輪設定中填入、存於瀏覽器。
 */
var SECRET_KEY = '請自訂一串隨機的長密碼（不要用範例、不要上傳）';

// 課表分頁名稱（每學年更新課表時，直接覆蓋這個分頁的內容即可，不必改程式碼）
var TIMETABLE_SHEET = 'Class_Timetables';

var SUMMARY_HEADERS = [
  "申請編號(ID)", "填報時間(Timestamp)", "代課日期(Date)", "學校名稱(School)",
  "教育階段(Level)", "請假教師(AbsentTeacher)", "代課教師(SubTeacher)", "班級(ClassName)",
  "假別(LeaveType)", "計費節數(PeriodsCount)", "單節費率(Rate)", "課堂鐘點費(ClassFee)",
  "是否代理導師(IsActingMentor)", "代理導師天數(ActingDays)", "導師費加給(MentorFee)",
  "整單總金額(TotalFee)", "計費模式(PayMode)", "備註(PayNote)"
];

// 註：「學校(School)」刻意加在最後一欄 —— 這樣既有資料（12 欄）讀取時位置不變，
// 且 DETAIL_COL_STATUS / DETAIL_COL_CONFIRM_DATE 的欄位編號不受影響。
// 跨校共聘老師請假時，同一天不同節可能在不同學校，故學校需記在「每一節」而非整單。
var DETAIL_HEADERS = [
  "申請編號(ID)", "代課日期(Date)", "班級(ClassName)", "節次(Period)",
  "科目(Subject)", "代課教師(SubTeacher)", "是否計費(Paid)", "單節費率(Rate)",
  "預估金額(Fee)", "核銷狀態(Status)", "交代事項(Handover)", "確認日期(ConfirmDate)",
  "學校(School)"
];

// 明細表欄位位置（1-indexed，供狀態更新使用）
var DETAIL_COL_STATUS = 10;
var DETAIL_COL_CONFIRM_DATE = 12;

function getCorsResponse(content, type) {
  var output = ContentService.createTextOutput(content);
  if (type === 'json') {
    output.setMimeType(ContentService.MimeType.JSON);
  } else {
    output.setMimeType(ContentService.MimeType.TEXT);
  }
  return output;
}

// 取得（或建立）指定月份的分頁，並確保標頭列存在
function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 把儲存格的日期值正規化成 'YYYY-MM-DD' 字串。
 * 試算表可能把日期存成 Date 物件（而非字串），比對前必須統一格式。
 */
function normalizeDateValue(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v || '').trim().substring(0, 10);
}

/**
 * 找出「同日期＋同代課教師＋同班級」的既有紀錄（疑似重複送出）。
 *
 * 為什麼需要：教師端送出鈕在上傳期間就會 disable，能擋手滑連點；
 * 但擋不住老師「重新再填一次」（2026-07-31 那三筆間隔 20 秒~3 分鐘就是這樣來的）。
 * 若無人察覺，核銷時 3 節會被算成 9 節。
 *
 * 回傳既有紀錄的摘要陣列，供前端顯示給老師確認。
 */
function findPossibleDuplicates(summarySheet, dateStr, subTeacher, className) {
  var lastRow = summarySheet.getLastRow();
  if (lastRow < 2) return [];

  var values = summarySheet.getRange(2, 1, lastRow - 1, SUMMARY_HEADERS.length).getValues();
  var target = {
    date: normalizeDateValue(dateStr),
    sub: String(subTeacher || '').trim(),
    cls: String(className || '').trim()
  };

  var hits = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    // 欄位順序見 SUMMARY_HEADERS：2=代課日期 6=代課教師 7=班級
    if (normalizeDateValue(row[2]) !== target.date) continue;
    if (String(row[6] || '').trim() !== target.sub) continue;
    if (String(row[7] || '').trim() !== target.cls) continue;

    hits.push({
      appId: row[0],
      timestamp: String(row[1] || ''),
      absentTeacher: row[5],
      periodsCount: Number(row[9]) || 0,
      totalFee: Number(row[15]) || 0
    });
  }
  return hits;
}

/**
 * 接收教師端網頁送出的資料 (POST)
 *
 * ⚠️ 重複偵測在這裡「只標記、不拒絕」。
 * 原因：教師端的 fetch 用 mode:'no-cors'（GAS 轉址的 CORS 限制），
 * 前端**讀不到 POST 的回應**。若這裡拒絕寫入，網頁仍會顯示送出成功，
 * 資料就靜悄悄不見了 —— 那比重複更糟。
 *
 * 真正讓老師看到確認框的關卡在 doGet 的 action=checkDuplicate（GET 讀得到回應），
 * 由教師端在送出前先問一次。這裡則是後防線：偵測到重複仍照常寫入，
 * 但在「備註」欄標上 ⚠️疑似重複，後台核銷時一眼看得到。
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var dateStr = data.date; // 格式為 YYYY-MM-DD
    if (!dateStr) {
      return getCorsResponse(JSON.stringify({ status: "error", message: "缺少代課日期 (date)！" }), 'json');
    }

    var month = dateStr.substring(0, 7); // YYYY-MM
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var summarySheet = getOrCreateSheet(ss, month + "_總表", SUMMARY_HEADERS);
    var detailSheet = getOrCreateSheet(ss, month + "_明細", DETAIL_HEADERS);

    // 重複偵測：只標記、不拒絕（原因見上方註解）
    // data.force === true 代表老師已在教師端的確認框按過「確定」，屬有意識的重送
    var dupNote = "";
    var dups = findPossibleDuplicates(summarySheet, data.date, data.subTeacher, data.className);
    if (dups.length > 0) {
      dupNote = data.force === true
        ? "⚠️疑似重複(老師已確認)：另有 " + dups.length + " 筆同日同班同代課教師"
        : "⚠️疑似重複：另有 " + dups.length + " 筆同日同班同代課教師，請核銷前確認";
    }

    var rate = Number(data.rate) || 0;
    var payMode = data.payMode || 'perPeriod';
    // 整天代理導師模式：金額由行政人員自行計算，此處不臆測費用
    var isActingMentor = data.isActingMentor === true || payMode === 'mentorDaily';

    var appId = "SUB-" + month.replace("-", "") + "-" + Math.floor(100 + Math.random() * 900);
    summarySheet.appendRow([
      appId,
      new Date().toISOString(),
      data.date,
      data.school || "",
      data.level || "",
      data.absentTeacher || "",
      data.subTeacher || "",
      data.className || "",
      data.leaveType || "",
      Number(data.periodsCount) || 0,
      rate,
      Number(data.classFee) || 0,
      isActingMentor ? "是" : "否",
      Number(data.actingDays) || 0,
      Number(data.mentorFee) || 0,
      Number(data.totalFee) || 0,
      payMode,
      // 備註：原本的備註在前，重複警告接在後面（兩者都可能有）
      [data.payNote || "", dupNote].filter(function (s) { return s; }).join("　")
    ]);

    // 明細：一節課一列
    if (data.periodsDetail && Array.isArray(data.periodsDetail)) {
      for (var i = 0; i < data.periodsDetail.length; i++) {
        var p = data.periodsDetail[i];
        var isPaid = p.paid === true;
        // 未計費的節次（午餐、午休、放學等指導）費率與金額皆為 0
        var periodRate = isPaid ? rate : 0;

        detailSheet.appendRow([
          appId,
          data.date,
          // 跨校時每節的班級可能不同（例：甲校五甲→乙校六班），優先用該節自帶的值
          p.className || data.className || "",
          p.period,
          p.subject || "",
          p.subTeacher || data.subTeacher || "",
          isPaid ? "是" : "否",
          periodRate,
          periodRate,
          "待核對", // 預設狀態
          p.handover || "",
          "", // 確認日期初始為空
          p.school || data.school || HOME_SCHOOL // 該節實際上課的學校
        ]);
      }
    }

    return getCorsResponse(JSON.stringify({ status: "success", appId: appId, message: "資料已成功按月寫入 Google 試算表！" }), 'json');
  } catch (error) {
    return getCorsResponse(JSON.stringify({ status: "error", message: error.toString() }), 'json');
  }
}

// 本校名稱：用來判斷「跨校共聘老師」（在本校以外的學校也有課的人）
var HOME_SCHOOL = '○○國小';

/**
 * 讀取 Class_Timetables 分頁。
 *
 * 支援兩種欄位格式，依標頭第一欄自動偵測（更新試算表與更新程式碼的順序不拘，不會中斷）：
 *
 *  【舊 5 欄】A=班級 B=星期 C=節次 D=科目 E=授課教師
 *     → 回傳 { "四甲": { 1: { 1: { subject, teacher } } } }（v1，向後相容）
 *
 *  【新 6 欄】A=學校 B=班級 C=星期 D=節次 E=科目 F=授課教師
 *     → 回傳 {
 *          schema: 'v2',
 *          homeSchool: '○○國小',
 *          byClass:   { 學校: { 班級: { 星期: { 節次: { subject, teacher } } } } },
 *          byTeacher: { 教師: { 星期: { 節次: { school, className, subject } } } },  // 所有老師（2026-09 起不再只限跨校）
 *          crossTeachers: [ 跨校老師姓名 ]  // 僅供相容判讀，前端已不依此決定流程
 *        }
 *
 * 2026-09-10 變更：byTeacher 改為涵蓋所有老師（原本只收跨校共聘老師）。
 * 前端因此可一律「選老師→查課表」，不必先問使用者「本校/跨校」。
 * 他校班級課表仍不完整（只有共聘老師零星節次），故 byClass 仍只暴露本校完整班級課表。
 */
function buildTimetable(ss) {
  var sheet = ss.getSheetByName(TIMETABLE_SHEET);
  if (!sheet) return {};

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return {};

  var header = sheet.getRange(1, 1, 1, 6).getValues()[0];
  var isV2 = String(header[0] || '').indexOf('學校') !== -1;

  if (!isV2) {
    // ---- 舊 5 欄格式 ----
    var values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    var db = {};
    for (var i = 0; i < values.length; i++) {
      var cls = String(values[i][0] || '').trim();
      var day = parseInt(values[i][1], 10);
      var period = parseInt(values[i][2], 10);
      if (!cls || !day || !period) continue;
      if (!db[cls]) db[cls] = {};
      if (!db[cls][day]) db[cls][day] = {};
      db[cls][day][period] = {
        subject: String(values[i][3] || '').trim(),
        teacher: String(values[i][4] || '').trim()
      };
    }
    return db;
  }

  // ---- 新 6 欄格式 ----
  var rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var byClass = {};
  var teacherRows = {};   // 教師 -> [列]
  var teacherSchools = {}; // 教師 -> { 學校: true }

  for (var j = 0; j < rows.length; j++) {
    var school = String(rows[j][0] || '').trim();
    var cls2 = String(rows[j][1] || '').trim();
    var day2 = parseInt(rows[j][2], 10);
    var period2 = parseInt(rows[j][3], 10);
    var subject2 = String(rows[j][4] || '').trim();
    var teacher2 = String(rows[j][5] || '').trim();
    if (!school || !cls2 || !day2 || !period2) continue;

    if (!byClass[school]) byClass[school] = {};
    if (!byClass[school][cls2]) byClass[school][cls2] = {};
    if (!byClass[school][cls2][day2]) byClass[school][cls2][day2] = {};
    byClass[school][cls2][day2][period2] = { subject: subject2, teacher: teacher2 };

    if (teacher2) {
      if (!teacherRows[teacher2]) teacherRows[teacher2] = [];
      teacherRows[teacher2].push({
        day: day2, period: period2, school: school, className: cls2, subject: subject2
      });
      if (!teacherSchools[teacher2]) teacherSchools[teacher2] = {};
      teacherSchools[teacher2][school] = true;
    }
  }

  // byTeacher 現在收「所有」老師（含本校），前端才能一律依老師查課表，
  // 不必再讓使用者自己判斷「我算本校還是跨校」。
  // crossTeachers 仍只列在本校以外也有課的人 —— 保留供舊版前端／其他消費者相容判讀用途。
  var byTeacher = {};
  var crossTeachers = [];
  for (var t in teacherRows) {
    byTeacher[t] = {};
    var list = teacherRows[t];
    for (var k = 0; k < list.length; k++) {
      var it = list[k];
      if (!byTeacher[t][it.day]) byTeacher[t][it.day] = {};
      byTeacher[t][it.day][it.period] = {
        school: it.school, className: it.className, subject: it.subject
      };
    }
    var hasOther = false;
    for (var s in teacherSchools[t]) {
      if (s !== HOME_SCHOOL) { hasOther = true; break; }
    }
    if (hasOther) crossTeachers.push(t);
  }
  crossTeachers.sort();

  return {
    schema: 'v2',
    homeSchool: HOME_SCHOOL,
    byClass: byClass,
    byTeacher: byTeacher,
    crossTeachers: crossTeachers
  };
}

// 列出試算表中所有有資料的月份（依 `YYYY-MM_總表` 分頁名稱推斷），由舊到新排序
function listDataMonths(ss) {
  var months = [];
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var m = sheets[i].getName().match(/^(\d{4}-\d{2})_總表$/);
    if (m) months.push(m[1]);
  }
  months.sort();
  return months;
}

// 讀取單一月份，組裝成主從結構 (Master-Detail)。找不到分頁或無資料時回傳空陣列。
function buildRecordsForMonth(ss, targetMonth) {
  var summarySheet = ss.getSheetByName(targetMonth + "_總表");
  var detailSheet = ss.getSheetByName(targetMonth + "_明細");
  if (!summarySheet || !detailSheet) return [];

  var sumLastRow = summarySheet.getLastRow();
  var detLastRow = detailSheet.getLastRow();
  if (sumLastRow <= 1) return [];

  var sumValues = summarySheet.getRange(2, 1, sumLastRow - 1, SUMMARY_HEADERS.length).getValues();
  var detValues = detLastRow > 1
    ? detailSheet.getRange(2, 1, detLastRow - 1, DETAIL_HEADERS.length).getValues()
    : [];

  // 明細依 appId 分組
  var detailsMap = {};
  for (var j = 0; j < detValues.length; j++) {
    var d = detValues[j];
    var key = d[0];
    if (!detailsMap[key]) detailsMap[key] = [];
    detailsMap[key].push({
      period: d[3].toString(),
      className: d[2],
      subject: d[4],
      subTeacher: d[5],
      paid: d[6] === "是",
      rate: Number(d[7]) || 0,
      fee: Number(d[8]) || 0,
      status: d[9],
      handover: d[10],
      confirmDate: d[11],
      // 舊資料（12 欄）此欄為空，視為本校，避免顯示空白
      school: d[12] || HOME_SCHOOL
    });
  }

  var records = [];
  for (var k = 0; k < sumValues.length; k++) {
    var s = sumValues[k];
    var currentAppId = s[0];
    records.push({
      id: currentAppId,
      timestamp: s[1],
      date: formatDate(s[2]),
      school: s[3],
      level: s[4],
      absentTeacher: s[5],
      subTeacher: s[6],
      className: s[7],
      leaveType: s[8],
      periodsCount: Number(s[9]) || 0,
      rate: Number(s[10]) || 0,
      classFee: Number(s[11]) || 0,
      isActingMentor: s[12] === "是",
      actingDays: Number(s[13]) || 0,
      mentorFee: Number(s[14]) || 0,
      totalFee: Number(s[15]) || 0,
      payMode: s[16] || 'perPeriod',
      payNote: s[17] || "",
      periodsDetail: detailsMap[currentAppId] || []
    });
  }
  return records;
}

/**
 * 提供行政端網頁讀取與修改資料 (GET)
 *
 * 【公開】不需金鑰：
 * - 取得課表：GET ?action=timetable        （teacher.html 使用）
 * - 重複檢查：GET ?action=checkDuplicate&date=&subTeacher=&className=
 *                                          （teacher.html 送出前使用）
 * - 常用代課教師名單：GET ?action=subTeacherNames
 *                                          （teacher.html 代課教師下拉選單用；
 *                                           只回傳姓名，不含日期/班級/金額等其他欄位）
 *
 * 【受保護】需帶 &key=<金鑰>：
 * - 讀取某月：GET ?month=YYYY-MM&key=...
 * - 讀取全部：GET ?month=all&key=...        （admin.html「備份所有 JSON」）
 * - 修改狀態：GET ?action=confirm&month=YYYY-MM&appId=SUB-XXX&period=節次&key=...
 */
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var params = e.parameter;

    // --- 課表（公開，不需金鑰）---
    // teacher.html 為公開網頁，無法安全保存金鑰；課表本身不含金額資訊。
    if (params.action === "timetable") {
      return getCorsResponse(JSON.stringify(buildTimetable(ss)), 'json');
    }

    // --- 重複檢查（公開，不需金鑰）---
    // teacher.html 送出前呼叫；查詢者本來就得自己提供日期／代課教師／班級，
    // 故意只回傳「筆數與節次」，不回金額與請假教師，避免變成資料外洩管道。
    if (params.action === "checkDuplicate") {
      var qDate = params.date || "";
      var qSub = params.subTeacher || "";
      var qCls = params.className || "";
      if (!qDate || !qSub) {
        return getCorsResponse(JSON.stringify({ status: "error", message: "缺少 date 或 subTeacher" }), 'json');
      }
      var qSheet = ss.getSheetByName(qDate.substring(0, 7) + "_總表");
      if (!qSheet) {
        return getCorsResponse(JSON.stringify({ status: "ok", count: 0, existing: [] }), 'json');
      }
      var found = findPossibleDuplicates(qSheet, qDate, qSub, qCls);
      var brief = found.map(function (r) {
        return {
          periodsCount: r.periodsCount,
          // 只給到分鐘，讓老師辨認「是不是我剛剛送的那筆」
          timestamp: String(r.timestamp).replace('T', ' ').substring(0, 16)
        };
      });
      return getCorsResponse(JSON.stringify({ status: "ok", count: brief.length, existing: brief }), 'json');
    }

    // --- 常用代課教師名單（公開，不需金鑰）---
    // 給教師端「代課教師」下拉選單用：把歷來所有月份總表出現過的代課教師姓名去重回傳。
    // 只回傳姓名，不含日期/班級/金額，比對 checkDuplicate 的公開尺度一致，不構成資料外洩。
    if (params.action === "subTeacherNames") {
      var monthsForNames = listDataMonths(ss);
      var nameSet = {};
      for (var nm = 0; nm < monthsForNames.length; nm++) {
        var nameSheet = ss.getSheetByName(monthsForNames[nm] + "_總表");
        if (!nameSheet) continue;
        var nameLastRow = nameSheet.getLastRow();
        if (nameLastRow < 2) continue;
        // SUMMARY_HEADERS 第 7 欄 = 代課教師(SubTeacher)
        var nameVals = nameSheet.getRange(2, 7, nameLastRow - 1, 1).getValues();
        for (var nv = 0; nv < nameVals.length; nv++) {
          var nameStr = String(nameVals[nv][0] || '').trim();
          if (nameStr) nameSet[nameStr] = true;
        }
      }
      var names = Object.keys(nameSet).sort();
      return getCorsResponse(JSON.stringify({ status: "ok", names: names }), 'json');
    }

    // --- 以下皆需金鑰 ---
    // 沒有這道檢查，任何人拿到網址就能用 ?month=all 抓走全部代課費資料。
    if (params.key !== SECRET_KEY) {
      return getCorsResponse(JSON.stringify({
        status: "error",
        message: "未授權：缺少或錯誤的金鑰。請在 admin.html 右上角齒輪設定中填入正確金鑰。"
      }), 'json');
    }

    // --- 狀態修改 (action = confirm) ---
    if (params.action === "confirm" && params.month && params.appId && params.period) {
      var detailSheet = ss.getSheetByName(params.month + "_明細");
      if (!detailSheet) {
        return getCorsResponse(JSON.stringify({ status: "error", message: "找不到該月份的明細分頁！" }), 'json');
      }

      var lastRow = detailSheet.getLastRow();
      if (lastRow <= 1) {
        return getCorsResponse(JSON.stringify({ status: "error", message: "該月份明細分頁尚無資料！" }), 'json');
      }

      var values = detailSheet.getRange(2, 1, lastRow - 1, DETAIL_HEADERS.length).getValues();
      var success = false;

      for (var i = 0; i < values.length; i++) {
        // 比對 申請編號(ID) 與 節次(Period)
        if (values[i][0] === params.appId && values[i][3].toString() === params.period.toString()) {
          var rowNum = i + 2; // 還原為試算表列號（1-indexed 且跳過標頭）
          detailSheet.getRange(rowNum, DETAIL_COL_STATUS).setValue("已確認");
          detailSheet.getRange(rowNum, DETAIL_COL_CONFIRM_DATE).setValue(new Date().toISOString());
          success = true;
          break;
        }
      }

      return getCorsResponse(JSON.stringify(success
        ? { status: "success", message: "狀態已更新為已確認！" }
        : { status: "error", message: "找不到對應的代課節次明細！" }), 'json');
    }

    // --- 讀取全部月份 (month=all) ---
    // 供 admin.html「備份所有 JSON」使用。按月分表後，單月請求只會拿到一個月，
    // 若沒有這個分支，「全部備份」會變成只備份單月的假備份。
    if (params.month === "all") {
      var months = listDataMonths(ss);
      var allRecords = [];
      for (var n = 0; n < months.length; n++) {
        allRecords = allRecords.concat(buildRecordsForMonth(ss, months[n]));
      }
      allRecords.sort(function (a, b) {
        return a.date < b.date ? -1 : (a.date > b.date ? 1 : 0);
      });
      return getCorsResponse(JSON.stringify(allRecords), 'json');
    }

    // --- 讀取單一月份 ---
    var targetMonth = params.month;
    if (!targetMonth) {
      var today = new Date();
      targetMonth = today.getFullYear() + "-" + ("0" + (today.getMonth() + 1)).slice(-2);
    }

    return getCorsResponse(JSON.stringify(buildRecordsForMonth(ss, targetMonth)), 'json');
  } catch (error) {
    return getCorsResponse(JSON.stringify({ status: "error", message: error.toString() }), 'json');
  }
}

// 輔助函式：確保日期格式為 YYYY-MM-DD
function formatDate(dateVal) {
  if (dateVal instanceof Date) {
    return dateVal.getFullYear() + "-" +
      ("0" + (dateVal.getMonth() + 1)).slice(-2) + "-" +
      ("0" + dateVal.getDate()).slice(-2);
  }
  var str = dateVal.toString();
  if (str.indexOf("GMT") !== -1 || str.indexOf("T") !== -1) {
    var dObj = new Date(str);
    return dObj.getFullYear() + "-" +
      ("0" + (dObj.getMonth() + 1)).slice(-2) + "-" +
      ("0" + dObj.getDate()).slice(-2);
  }
  return str;
}

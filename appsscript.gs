// ── nu_engagementstream 整合版（問卷 + 圖表標註）──
const SHEET_ID  = '1v-oOkUm3ESFVKYBQLZqOdUjBZgSvFrU36bEKdtI4hHA';
const FOLDER_ID = '1Px-JhX2UH_dYWJnbgczRuWo7pA7ozkYG';
// ─────────────────────────────────────────────────

// 表頭（欄位順序即寫入順序）
const HEADERS = [
  '時間戳記', '日期', '桌號', '身分',
  // A. 基本資訊
  'A1_研究代碼', 'A3_課程', 'A3_其他', 'A4_第幾次', 'A5_閱讀狀況',
  'A6_做了哪些事', 'A7_花費時間', 'A8_回看段落',
  // B. 理解/信任/感受（1-5）
  'B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8', 'B9', 'B10', 'B11',
  // C. 看見了什麼
  'C1_情況', 'C1_其他',
  'T1_起(min)', 'T1_訖(min)', 'T1_類型', 'T1_類型_其他', 'T1_描述',
  'T2_起(min)', 'T2_訖(min)', 'T2_類型', 'T2_類型_其他', 'T2_描述',
  'T3_起(min)', 'T3_訖(min)', 'T3_類型', 'T3_類型_其他', 'T3_描述',
  'C3_原因', 'C3_其他', 'C4_想釐清',
  // D. 行動
  'D0_是否調整', 'D0-1_原因', 'D0-1_其他',
  // D-T 教師
  'DT1_是否回看', 'DT2_1_觀察', 'DT2_2_推測原因', 'DT2_3_調整策略', 'DT2_4_驗證方式',
  'DT3_調整類型', 'DT3_其他', 'DT4_協助', 'DT4_其他',
  // D-S 學生
  'DS1_校準意願', 'DS2_行為目標', 'DS2_其他', 'DS3_具體行為',
  'DS4_角色分工', 'DS5_角色', 'DS5_評分',
  // E. 效益（1-5）
  'E1', 'E2', 'E3', 'E4', 'E5', 'E6', 'E7',
  // F. 介面
  'F1_圖例', 'F1_段落標註', 'F1_關鍵時間點', 'F1_讀圖提示', 'F1_分組排序', 'F1_摘要指標',
  'F2_改進項目', 'F2_其他', 'F3_補充資訊', 'F4_新增功能',
  // G. 其他
  'G1_受訪意願', 'G2_補充'
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const { date, group, role, segments, answers, screenshot, screenshot_left, screenshot_right, timestamp } = data;
    const a = answers || {};
    const s = segments || [{}, {}, {}];

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];

    // 第一次寫入：建立表頭
    if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);

    // 將陣列（複選）用 ` | ` 連接為字串
    const join = (v) => Array.isArray(v) ? v.join(' | ') : (v || '');

    const row = [
      timestamp, date, group, role,
      // A
      a.A1 || '', a.A3 || '', a.A3_other || '', a.A4 || '', a.A5 || '',
      join(a.A6), a.A7 || '', a.A8 || '',
      // B
      a.B1||'', a.B2||'', a.B3||'', a.B4||'', a.B5||'', a.B6||'',
      a.B7||'', a.B8||'', a.B9||'', a.B10||'', a.B11||'',
      // C
      join(a.C1), a.C1_other || '',
      s[0].start || '', s[0].end || '', a.C2_1_type || '', a.C2_1_type_other || '', a.C2_1_desc || '',
      s[1].start || '', s[1].end || '', a.C2_2_type || '', a.C2_2_type_other || '', a.C2_2_desc || '',
      s[2].start || '', s[2].end || '', a.C2_3_type || '', a.C2_3_type_other || '', a.C2_3_desc || '',
      join(a.C3), a.C3_other || '', a.C4 || '',
      // D
      a.D0 || '', join(a.D0_1), a.D0_1_other || '',
      // D-T
      a.DT1 || '', a.DT2_1 || '', a.DT2_2 || '', a.DT2_3 || '', a.DT2_4 || '',
      join(a.DT3), a.DT3_other || '', join(a.DT4), a.DT4_other || '',
      // D-S
      a.DS1 || '', a.DS2 || '', a.DS2_other || '', a.DS3 || '',
      a.DS4 || '', a.DS5_role || '', a.DS5_rating || '',
      // E
      a.E1||'', a.E2||'', a.E3||'', a.E4||'', a.E5||'', a.E6||'', a.E7||'',
      // F
      a.F1_legend||'', a.F1_segment||'', a.F1_keytime||'', a.F1_readhint||'', a.F1_grouporder||'', a.F1_summary||'',
      join(a.F2), a.F2_other || '', a.F3 || '', a.F4 || '',
      // G
      a.G1 || '', a.G2 || ''
    ];
    sheet.appendRow(row);

    // 儲存截圖到 Drive（左右兩張分開存；保留舊 screenshot 欄位向後相容）
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const safeRole = (role || 'x').replace(/[^\w]/g, '');
    const stamp = timestamp.replace(/[:.]/g, '-');
    const baseName = `${date}_${group}_${safeRole}_${stamp}`;

    if (screenshot_left) {
      const blob = Utilities.newBlob(Utilities.base64Decode(screenshot_left), 'image/png', `${baseName}_left.png`);
      folder.createFile(blob);
    }
    if (screenshot_right) {
      const blob = Utilities.newBlob(Utilities.base64Decode(screenshot_right), 'image/png', `${baseName}_right.png`);
      folder.createFile(blob);
    }
    if (screenshot) {
      const blob = Utilities.newBlob(Utilities.base64Decode(screenshot), 'image/png', `${baseName}.png`);
      folder.createFile(blob);
    }

    return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput('ERROR: ' + err.message).setMimeType(ContentService.MimeType.TEXT);
  }
}

function doGet(e) {
  return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
}

// ── 工具：重設表頭（在 Apps Script 編輯器手動執行一次即可）──
// 用法：選擇函式「resetHeaders」→ 按「執行」。
// 會把第 1 列換成最新的 HEADERS，不會動到其他列的資料。
function resetHeaders() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  // 清除第 1 列舊內容
  sheet.getRange(1, 1, 1, sheet.getMaxColumns()).clearContent();
  // 寫入新表頭
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  // 凍結第 1 列、粗體
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  SpreadsheetApp.flush();
  Logger.log('已寫入 ' + HEADERS.length + ' 個欄位');
}

// ── 工具：清除全部資料（保留表頭）──
// 小心使用！會刪掉所有回應
function clearAllResponses() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getMaxColumns()).clearContent();
  Logger.log('已清除 ' + (lastRow - 1) + ' 列');
}

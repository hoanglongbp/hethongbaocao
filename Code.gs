// ==========================================================
// MÃ NGUỒN GOOGLE APPS SCRIPT HOÀN CHỈNH (Code.gs)
// CƠ QUAN: ỦY BAN MTTQ VIỆT NAM THÀNH PHỐ ĐỒNG NAI
// BẢN QUYỀN: LÊ VĂN NGUYÊN
// ==========================================================

const SHEET_NAME = "Báo Cáo Đơn Vị";
const DRIVE_FOLDER_NAME = "Thư Mục Báo Cáo MTTQ thành phố Đồng Nai";
const SPREADSHEET_ID = "1t0PhCETWuFY-b_nJRBb7BtIShMky6sUvPJtUTYkSJUM";

const ADMIN_SHEET_NAME = "Tài Khoản Quản Trị";
const UNIT_SHEET_NAME = "Danh Mục Đơn Vị";
const PERIOD_SHEET_NAME = "Danh Mục Kỳ Đột Xuất";
const AGENCY_NAME = "Ủy ban MTTQ Việt Nam thành phố Đồng Nai";

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Hệ Thống Báo Cáo - ' + AGENCY_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function recordVisit() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(3000); 
    const props = PropertiesService.getScriptProperties();
    let totalCount = parseInt(props.getProperty("SYSTEM_VISIT_COUNT") || "0", 10);
    totalCount++;
    props.setProperty("SYSTEM_VISIT_COUNT", totalCount.toString());
    
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    let lastDate = props.getProperty("LAST_VISIT_DATE");
    let dailyCount = parseInt(props.getProperty("DAILY_VISIT_COUNT") || "0", 10);
    
    if (todayStr !== lastDate) {
      dailyCount = 1;
      props.setProperty("LAST_VISIT_DATE", todayStr);
    } else {
      dailyCount++;
    }
    props.setProperty("DAILY_VISIT_COUNT", dailyCount.toString());
    
    return { total: totalCount, daily: dailyCount };
  } catch (e) {
    return { total: "...", daily: "..." };
  } finally {
    lock.releaseLock();
  }
}

function getSpreadsheet_() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch(e) {
    throw new Error("Lỗi quyền truy cập Database. Vui lòng kiểm tra lại thiết lập Deploy (Execute as: Me).");
  }
}

function getOutputParentFolder_() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function getPeriodFolder_(periodValue) {
  const parentFolder = getOutputParentFolder_();
  const safeName = String(periodValue || "Khác").replace(/[\\/:*?"<>|]/g, "-").trim();
  const folders = parentFolder.getFoldersByName(safeName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(safeName);
}

// === QUẢN LÝ KỲ ĐỘT XUẤT ===
function ensurePeriodSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(PERIOD_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(PERIOD_SHEET_NAME);
    sh.getRange(1, 1, 1, 5).setValues([["Tên Kỳ Đột Xuất", "Người Tạo", "Ngày Tạo", "Trạng Thái", "Yêu Cầu (Form)"]]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function getAdhocPeriods() {
  try {
    const sh = ensurePeriodSheet_();
    if (sh.getLastRow() <= 1) return [];
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getDisplayValues();
    return data.map(r => {
      return {
        name: String(r[0]).trim(),
        creator: String(r[1]).trim(),
        date: String(r[2]).trim(),
        status: String(r[3] ? r[3] : "MỞ").trim().toUpperCase(),
        desc: String(r[4] || "").trim()
      };
    }).filter(x => x.name);
  } catch(e) { return []; }
}

function addAdhocPeriod(role, username, periodName, desc) {
  if (role !== "SUPER_ADMIN" && role !== "ADMIN") return { success: false, message: "Không có quyền quản trị!" };
  if (!periodName || !periodName.trim()) return { success: false, message: "Tên kỳ không hợp lệ!" };
  let name = periodName.trim();
  if (!name.toLowerCase().startsWith("đột xuất:")) name = "Đột xuất: " + name;

  const sh = ensurePeriodSheet_();
  const existing = getAdhocPeriods().map(p => p.name);
  if (existing.includes(name)) return { success: false, message: "Kỳ báo cáo này đã tồn tại!" };
  
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  sh.appendRow([name, username, now, "MỞ", String(desc || "").trim()]);
  return { success: true, message: "Thêm đợt báo cáo đột xuất thành công!" };
}

function editAdhocPeriod(role, oldName, newNameRaw, desc) {
  if (role !== "SUPER_ADMIN" && role !== "ADMIN") return { success: false, message: "Không có quyền quản trị!" };
  if (!newNameRaw || !newNameRaw.trim()) return { success: false, message: "Tên mới không hợp lệ!" };
  let newName = newNameRaw.trim();
  if (!newName.toLowerCase().startsWith("đột xuất:")) newName = "Đột xuất: " + newName;

  const sh = ensurePeriodSheet_();
  const data = sh.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === oldName.trim()) {
      sh.getRange(i + 1, 1).setValue(newName);
      sh.getRange(i + 1, 5).setValue(String(desc || "").trim());
      found = true; break;
    }
  }
  if (!found) return { success: false, message: "Không tìm thấy đợt này!" };

  if (oldName !== newName) {
    const rSheet = getSpreadsheet_().getSheetByName(SHEET_NAME);
    if(rSheet && rSheet.getLastRow() > 1) {
      const rData = rSheet.getDataRange().getValues();
      for(let j = 1; j < rData.length; j++){
        if(String(rData[j][4]).trim() === oldName.trim()) rSheet.getRange(j + 1, 5).setValue(newName);
      }
    }
  }
  return { success: true, message: "Đã cập nhật thông tin đợt thành công!" };
}

function toggleAdhocPeriod(role, periodName, currentStatus) {
  if (role !== "SUPER_ADMIN" && role !== "ADMIN") return { success: false, message: "Không có quyền!" };
  const sh = ensurePeriodSheet_();
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === periodName.trim()) {
      const newStatus = (currentStatus === "KHÓA") ? "MỞ" : "KHÓA";
      sh.getRange(i + 1, 4).setValue(newStatus);
      return { success: true, message: "Trạng thái đợt báo cáo đã được " + (newStatus==="KHÓA" ? "Khóa" : "Mở") };
    }
  }
  return { success: false, message: "Không tìm thấy đợt này!" };
}

function deleteAdhocPeriod(role, periodName) {
  if (role !== "SUPER_ADMIN" && role !== "ADMIN") return { success: false, message: "Không có quyền quản trị!" };
  const sh = ensurePeriodSheet_();
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === periodName.trim()) {
      sh.deleteRow(i + 1);
      return { success: true, message: "Đã xóa đợt báo cáo đột xuất!" };
    }
  }
  return { success: false, message: "Không tìm thấy đợt này!" };
}

// === QUẢN LÝ BÁO CÁO CỦA ĐƠN VỊ ===
function saveReport(data) {
  try {
    const ss = getSpreadsheet_();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(["ID", "Thời Gian Gửi", "Đơn Vị", "Loại Kỳ", "Kỳ Báo Cáo", "Tên File", "Link Google Drive", "Trạng Thái", "Nội Dung Báo Cáo"]);
      sheet.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#fee2e2");
      sheet.setFrozenRows(1);
    }

    const existingReports = getSubmittedReportsInternal_();
    const unitsToSubmit = String(data.unit).split(',').map(u => u.trim()).filter(Boolean);
    if (unitsToSubmit.length === 0) return { success: false, message: "Không có đơn vị nào được chọn!" };

    let fileUrl = ""; let fileName = "";
    if (data.fileData && data.fileName) {
      const folder = getPeriodFolder_(data.periodValue);
      const parts = String(data.fileData).split(',');
      if (parts.length < 2) throw new Error("Dữ liệu tệp không hợp lệ.");
      const decodedData = Utilities.base64Decode(parts[1]);
      const blob = Utilities.newBlob(decodedData, data.mimeType || MimeType.PLAIN_TEXT, data.fileName);
      const file = folder.createFile(blob);
      fileUrl = file.getUrl(); fileName = file.getName();
    }

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    const textContent = String(data.textContent || "");
    let successCount = 0; let dupCount = 0;

    unitsToSubmit.forEach((u, index) => {
      const isDuplicate = existingReports.some(item =>
        String(item.unit).trim() === u && String(item.periodValue).trim().toLowerCase() === String(data.periodValue).trim().toLowerCase()
      );
      if (isDuplicate) { dupCount++; } 
      else {
        const id = Date.now().toString() + index.toString();
        sheet.appendRow([id, timestamp, u, String(data.periodType).trim(), String(data.periodValue).trim(), fileName, fileUrl, "Đã nộp", textContent]);
        successCount++;
      }
    });

    if (successCount === 0) return { success: false, message: "Tất cả các đơn vị bạn chọn đều đã nộp báo cáo cho kỳ này rồi!" };
    if (dupCount > 0) return { success: true, message: `Gửi thành công cho ${successCount} đơn vị. (Bỏ qua ${dupCount} đơn vị do đã nộp trước đó)` };
    return { success: true, message: `Gửi báo cáo thành công cho ${successCount} đơn vị!` };
  } catch (error) { return { success: false, message: "Lỗi Server: " + error.toString() }; }
}

function getSubmittedReportsInternal_() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const rawData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  const displayData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getDisplayValues();
  
  return rawData.map((r, i) => ({
    id: String(displayData[i][0]).replace(/,/g, "").trim(), timestamp: displayData[i][1],
    unit: String(r[2]).trim(), periodType: String(r[3]).trim(), periodValue: String(r[4]).trim(),
    fileName: String(r[5]), fileUrl: String(r[6]), status: String(r[7] || "Đã nộp"), textContent: String(r[8] || "")
  })).filter(x => x.unit);
}

function getSubmittedReports(isAdmin) {
  try {
    const reports = getSubmittedReportsInternal_();
    return reports.map(r => ({
      id: r.id, timestamp: r.timestamp, unit: r.unit, periodType: r.periodType, periodValue: r.periodValue,
      fileName: isAdmin ? r.fileName : "", fileUrl: isAdmin ? r.fileUrl : "", status: r.status, textContent: r.textContent
    })).reverse();
  } catch (e) { return []; }
}

function deleteReportAdvanced(id, unitName, periodValue) {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { success: false, message: "Sheet không tồn tại" };
    const dataDisp = sheet.getDataRange().getDisplayValues();
    const dataRaw = sheet.getDataRange().getValues();
    const targetId = String(id).replace(/,/g, "").trim();
    const targetUnit = String(unitName).trim();
    const targetPeriod = String(periodValue).trim();
    
    for (let i = 1; i < dataDisp.length; i++) {
      const rIdDisp = String(dataDisp[i][0]).replace(/,/g, "").trim();
      const rIdRaw = String(dataRaw[i][0]).replace(/,/g, "").trim();
      const rUnit = String(dataDisp[i][2]).trim();
      const rPeriod = String(dataDisp[i][4]).trim();
      
      if ( (targetId && (rIdDisp === targetId || rIdRaw === targetId)) || (targetUnit === rUnit && targetPeriod === rPeriod) ) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "Đã xóa báo cáo và mở khóa thành công!" };
      }
    }
    return { success: false, message: "Không tìm thấy báo cáo trong CSDL để xóa!" };
  } catch(err) { return { success: false, message: err.toString() }; }
}

function extractFileIdFromUrl_(url) {
  const m = url ? url.match(/[-\w]{25,}/) : null; return m ? m[0] : "";
}

// BỘ CHUYỂN ĐỔI FILE DỮ LIỆU
function readContentFromDriveFile_(fileUrl) {
  try {
    const fileId = extractFileIdFromUrl_(fileUrl); if (!fileId) return "";
    const file = DriveApp.getFileById(fileId);
    const mimeType = file.getMimeType(); const blob = file.getBlob();

    if (mimeType.includes("spreadsheet") || mimeType.includes("excel") || mimeType.includes("csv")) {
      try {
        const tempFile = Drive.Files.insert({ title: "temp_excel_" + fileId, mimeType: MimeType.GOOGLE_SHEETS }, blob);
        const ss = SpreadsheetApp.openById(tempFile.id);
        const sheets = ss.getSheets();
        let excelText = "";
        
        for (let i = 0; i < sheets.length; i++) {
          const data = sheets[i].getDataRange().getDisplayValues();
          if (data.length > 0 && data[0].join("").trim() !== "") {
            excelText += `--- BẢNG DỮ LIỆU: ${sheets[i].getName()} ---\n`;
            data.forEach(row => {
              if(row.join("").trim() !== "") {
                 excelText += row.join(" | ") + "\n";
              }
            });
            excelText += "\n";
          }
        }
        DriveApp.getFileById(tempFile.id).setTrashed(true);
        return excelText ? excelText : "Tệp Excel trống hoặc không có số liệu.";
      } catch (e) { return "Tệp Excel: " + file.getName() + " (Đã tiếp nhận tệp)"; }
    }

    if (mimeType.includes("word") || mimeType.includes("officedocument.wordprocessingml")) {
      try {
        const tempDoc = Drive.Files.insert({ title: "temp_read_" + fileId, mimeType: MimeType.GOOGLE_DOCS }, blob);
        const text = DocumentApp.openById(tempDoc.id).getBody().getText();
        DriveApp.getFileById(tempDoc.id).setTrashed(true); return text;
      } catch (e) { return "Tệp: " + file.getName(); }
    }
    
    if (mimeType.includes("pdf") || mimeType.includes("image")) {
      try {
        const tempDoc = Drive.Files.insert({ title: "temp_ocr_" + fileId }, blob, { ocr: true, ocrLanguage: "vi" });
        const text = DocumentApp.openById(tempDoc.id).getBody().getText();
        DriveApp.getFileById(tempDoc.id).setTrashed(true); return text;
      } catch (e) { return "Tệp: " + file.getName(); }
    }
    
    try { return blob.getDataAsString("UTF-8"); } catch(e) { return "Tệp: " + file.getName(); }
  } catch (err) { return ""; }
}

function getAggregateReport(filterType, filterValue) {
  try {
    let baseUnits = getUnitCatalog(); 
    if (filterType === "Đột xuất" && filterValue !== "ALL") {
      const adhocList = getAdhocPeriods();
      const currentPeriod = adhocList.find(p => p.name === filterValue);
      if (currentPeriod && currentPeriod.desc) {
        try {
          const config = JSON.parse(currentPeriod.desc);
          if (config.targets && config.targets.length > 0) {
            baseUnits = baseUnits.filter(u => config.targets.includes(u.unitName));
          }
        } catch(e) {} 
      }
    }

    const allReports = getSubmittedReportsInternal_();
    let matchedReports = allReports;
    if (filterType && filterValue && filterValue !== "ALL") {
      matchedReports = allReports.filter(r => {
        if (filterType === "Đột xuất") return r.periodType === "Đột xuất" && r.periodValue.toLowerCase().includes(filterValue.toLowerCase().trim());
        
        // Sửa lỗi "Tháng Tháng" để ghi nhận báo cáo cũ:
        let fv = filterValue.trim();
        let fvBug = fv.startsWith("Tháng ") ? "Tháng " + fv : fv;
        return r.periodType === filterType && (r.periodValue.trim() === fv || r.periodValue.trim() === fvBug);
      });
    }

    const submittedMap = new Map();
    matchedReports.forEach(r => submittedMap.set(r.unit.trim(), r));

    const unitDetails = baseUnits.map(u => {
      const rep = submittedMap.get(u.unitName.trim());
      return {
        unitName: u.unitName, 
        unitType: u.unitType, 
        isSubmitted: !!rep, 
        timestamp: rep ? rep.timestamp : "-",
        periodType: rep ? rep.periodType : (filterType !== "ALL" ? filterType : "-"), // Fixed undefined Loại Kỳ
        periodValue: rep ? rep.periodValue : (filterValue !== "ALL" ? filterValue : "-"),
        fileName: rep ? rep.fileName : "", 
        fileUrl: rep ? rep.fileUrl : "", 
        textContent: rep ? rep.textContent : "", 
        id: rep ? rep.id : ""
      };
    });

    const total = baseUnits.length; 
    const submittedCount = unitDetails.filter(x => x.isSubmitted).length;
    const pendingCount = unitDetails.filter(x => !x.isSubmitted).length;
    const rate = total ? Math.round((submittedCount / total) * 100) : 0;
    
    let folderUrl = "";
    if (filterValue && filterValue !== "ALL") { try { folderUrl = getPeriodFolder_(filterValue).getUrl(); } catch (e) {} }

    return { success: true, summary: { total, submittedCount, pendingCount, rate, filterType, filterValue, folderUrl }, details: unitDetails };
  } catch (err) { return { success: false, message: err.toString() }; }
}

function prepareDataForClientGemini(filterType, filterValue) {
  try {
    let baseUnits = getUnitCatalog();
    let isForm = false;
    let formQuestions = [];

    if (filterType === "Đột xuất" && filterValue !== "ALL") {
      const adhocList = getAdhocPeriods();
      const currentPeriod = adhocList.find(p => p.name === filterValue);
      if (currentPeriod && currentPeriod.desc) {
        try {
          const config = JSON.parse(currentPeriod.desc);
          if (config.targets && config.targets.length > 0) baseUnits = baseUnits.filter(u => config.targets.includes(u.unitName));
          if (config.questions && config.questions.length > 0 && config.questions[0].q) {
            isForm = true;
            formQuestions = config.questions;
          }
        } catch(e) {}
      }
    }

    const allReports = getSubmittedReportsInternal_();
    let matchedReports = allReports;
    if (filterType && filterValue && filterValue !== "ALL") {
      matchedReports = allReports.filter(r => {
        if (filterType === "Đột xuất") return r.periodType === "Đột xuất" && r.periodValue.toLowerCase().includes(filterValue.toLowerCase().trim());
        
        let fv = filterValue.trim();
        let fvBug = fv.startsWith("Tháng ") ? "Tháng " + fv : fv;
        return r.periodType === filterType && (r.periodValue.trim() === fv || r.periodValue.trim() === fvBug);
      });
    }

    const submittedMap = new Map();
    matchedReports.forEach(r => submittedMap.set(r.unit.trim(), r));

    if (isForm) {
      const safeName = filterValue.replace(/[\\/:*?"<>|]/g, "_");
      let ssNew = SpreadsheetApp.create(`Tong_Hop_So_Lieu_${safeName}`);
      let shNew = ssNew.getActiveSheet();
      
      let headers = ["STT", "Đơn vị"];
      formQuestions.forEach(q => headers.push(q.q));
      shNew.appendRow(headers);
      shNew.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f1f5f9");

      let totals = new Array(formQuestions.length).fill(0);
      let count = 1;
      
      baseUnits.forEach((u) => {
        const rep = submittedMap.get(u.unitName.trim());
        if (rep) {
          let rowData = [count++, u.unitName];
          let answers = {};
          if (rep.textContent) {
            rep.textContent.split('\n').forEach(line => {
              let splitIdx = line.indexOf(':');
              if (splitIdx > -1) {
                let k = line.substring(0, splitIdx).trim();
                let v = line.substring(splitIdx + 1).trim();
                answers[k] = v;
              }
            });
          }

          formQuestions.forEach((q, idx) => {
            let val = answers[q.q] || "0";
            rowData.push(val);
            if (q.t === 'number') {
               let num = parseFloat(val.replace(/,/g, ''));
               if (!isNaN(num)) totals[idx] += num;
            }
          });
          shNew.appendRow(rowData);
        }
      });

      let totalRow = ["", "TỔNG CỘNG ĐÃ NỘP"];
      formQuestions.forEach((q, idx) => {
         if(q.t === 'number') {
             totalRow.push(totals[idx]);
         } else {
             totalRow.push("-");
         }
      });
      shNew.appendRow(totalRow);
      shNew.getRange(shNew.getLastRow(), 1, 1, totalRow.length).setFontWeight("bold").setBackground("#fffbeb").setFontColor("#9a3412");

      let fileUrl = ssNew.getUrl();
      
      try {
        let file = DriveApp.getFileById(ssNew.getId());
        let parentFolder = getPeriodFolder_(filterValue);
        file.moveTo(parentFolder);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch(e) {}

      return {
         success: true, dataType: 'sheet', sheetUrl: fileUrl,
         periodValue: filterValue, total: baseUnits.length,
         submittedCount: submittedMap.size, pendingCount: baseUnits.length - submittedMap.size,
         rate: baseUnits.length ? Math.round((submittedMap.size / baseUnits.length) * 100) : 0
      };
    } 
    else {
      const submittedUnits = []; const pendingUnits = []; let extractedFileCorpus = "";

      baseUnits.forEach((u, i) => {
        const rep = submittedMap.get(u.unitName.trim());
        if (rep) {
          submittedUnits.push({ unitName: u.unitName });
          let unitText = "";
          if (rep.textContent) unitText += `\nNội dung văn bản: ${rep.textContent}\n`;
          if (rep.fileUrl) {
            const fileContent = readContentFromDriveFile_(rep.fileUrl);
            if (fileContent) unitText += `\nNội dung file: ${fileContent.substring(0, 3500)}`;
          }
          extractedFileCorpus += `\n--- [ĐƠN VỊ ${i+1}: ${u.unitName}] ---\n${unitText || "Không có dữ liệu văn bản."}\n`;
        } else { pendingUnits.push({ unitName: u.unitName }); }
      });

      const systemPrompt = `Bạn là Chuyên viên Tổng hợp của Ủy ban MTTQ Việt Nam thành phố Đồng Nai.
Vui lòng đọc dữ liệu báo cáo bên dưới và thực hiện các việc sau:
1. TÌM VÀ CỘNG TỔNG SỐ LIỆU: Tìm tất cả các con số (tiền, gạo, thẻ BHYT, sửa nhà...) trong các báo cáo và cộng dồn lại theo từng hạng mục.
2. TÓM TẮT NỘI DUNG SƠ BỘ: Nêu tình hình chung, các điểm nổi bật và khó khăn kiến nghị (nếu có).
3. ĐÁNH GIÁ CHUNG: Viết 1 đoạn văn nhận xét ngắn gọn.

DỮ LIỆU CÁC ĐƠN VỊ ĐÃ NỘP BÁO CÁO KỲ NÀY:
${extractedFileCorpus}`;

      return {
        success: true, dataType: 'prompt', promptText: systemPrompt,
        periodValue: filterValue, total: baseUnits.length,
        submittedCount: submittedUnits.length, pendingCount: pendingUnits.length,
        rate: baseUnits.length ? Math.round((submittedUnits.length / baseUnits.length) * 100) : 0
      };
    }
  } catch (err) { return { success: false, message: err.toString() }; }
}

function zipAndDownloadReports(filterType, filterValue) {
  try {
    const allReports = getSubmittedReportsInternal_();
    let matchedReports = allReports;
    if (filterType && filterType !== "ALL") matchedReports = matchedReports.filter(r => r.periodType === filterType);
    if (filterValue && filterValue !== "ALL") {
      matchedReports = matchedReports.filter(r => {
        if (filterType === "Đột xuất") return r.periodValue.toLowerCase().includes(filterValue.toLowerCase().trim());
        let fv = filterValue.trim();
        let fvBug = fv.startsWith("Tháng ") ? "Tháng " + fv : fv;
        return r.periodValue.trim() === fv || r.periodValue.trim() === fvBug;
      });
    }

    if (matchedReports.length === 0) return { success: false, message: "Không có báo cáo nào để tải!" };

    const blobs = [];
    matchedReports.forEach(r => {
      if (r.fileUrl) {
        const fileId = extractFileIdFromUrl_(r.fileUrl);
        if (fileId) {
          try {
            const file = DriveApp.getFileById(fileId);
            const safeUnitName = r.unit.replace(/[\\/:*?"<>|]/g, "_");
            const newName = `[${safeUnitName}] ${file.getName()}`;
            const blob = file.getBlob(); blob.setName(newName); blobs.push(blob);
          } catch (e) {}
        }
      }
    });

    if (blobs.length === 0) return { success: false, message: "Không lấy được tệp đính kèm nào từ Drive!" };

    const safeName = (filterValue && filterValue !== "ALL") ? filterValue : (filterType && filterType !== "ALL" ? filterType : "Tat_Ca");
    const zipName = `Toan_Bo_Bao_Cao_${safeName.replace(/[\\/:*?"<>| ]/g, "_")}.zip`;
    const zipBlob = Utilities.zip(blobs, zipName);
    const zipFile = getOutputParentFolder_().createFile(zipBlob);
    try { zipFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    
    return { success: true, downloadUrl: `https://drive.google.com/uc?export=download&id=${zipFile.getId()}`, count: blobs.length };
  } catch (err) { return { success: false, message: err.toString() }; }
}

function hashPassword_(p) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(p || ""), Utilities.Charset.UTF_8).map(b => (b + 256).toString(16).slice(-2)).join("");
}

function ensureAdminSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(ADMIN_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(ADMIN_SHEET_NAME);
    sh.getRange(1, 1, 1, 6).setValues([["Tên Đăng Nhập", "Mật Khẩu SHA-256", "Họ Tên", "Vai Trò", "Trạng Thái", "Ngày Tạo"]]).setFontWeight("bold");
    sh.appendRow(["admin", hashPassword_("Dongnai@2026"), "Quản Trị Viên Cao Cấp", "SUPER_ADMIN", "ACTIVE", ""]);
  }
  return sh;
}

function authenticateAdmin(user, password) {
  const u = String(user || "").trim().toLowerCase();
  const p = hashPassword_(password);
  const sh = ensureAdminSheet_();
  const rows = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 6).getDisplayValues();
  const found = rows.find(r => String(r[0]).trim().toLowerCase() === u && String(r[1]).trim() === p && String(r[4] || "ACTIVE").toUpperCase() === "ACTIVE");
  if (!found) return { success: false, message: "Tài khoản hoặc mật khẩu không chính xác!" };
  return { success: true, user: found[0], fullName: found[2] || found[0], role: found[3] || "VIEWER" };
}

function changeAdminPassword(u, oldP, newP) {
  try {
    const sh = ensureAdminSheet_();
    const rows = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 6).getDisplayValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toLowerCase() === String(u).trim().toLowerCase()) {
        if (rows[i][1] !== hashPassword_(oldP)) return { success: false, message: "Mật khẩu cũ không chính xác!" };
        sh.getRange(i + 2, 2).setValue(hashPassword_(newP));
        return { success: true, message: "Đổi mật khẩu thành công!" };
      }
    }
    return { success: false, message: "Lỗi!" };
  } catch (err) { return { success: false, message: err.toString() }; }
}

function listAccounts(role) {
  if (role !== "SUPER_ADMIN") return { success: false };
  return { success: true, accounts: ensureAdminSheet_().getRange(2, 1, Math.max(ensureAdminSheet_().getLastRow() - 1, 1), 6).getDisplayValues().filter(r => r[0]).map((r, i) => ({ row: i + 2, username: r[0], fullName: r[2], role: r[3], status: r[4], createdAt: r[5] })) };
}

function createAccount(r, d) {
  if (r !== "SUPER_ADMIN") return { success: false, message: "Chỉ Super Admin!" };
  ensureAdminSheet_().appendRow([d.username.toLowerCase(), hashPassword_(d.password), d.fullName, d.role, "ACTIVE", ""]);
  return { success: true, message: "Tạo thành công!" };
}

function resetUserPasswordBySuperAdmin(r, u, p) {
  if (r !== "SUPER_ADMIN") return { success: false };
  const sh = ensureAdminSheet_();
  const rows = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 6).getDisplayValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === String(u).toLowerCase()) {
      sh.getRange(i + 2, 2).setValue(hashPassword_(p));
      return { success: true, message: "Đã đổi mật khẩu!" };
    }
  }
  return { success: false, message: "Không tìm thấy!" };
}

function deleteAccountBySuperAdmin(r, row, u) {
  if (r !== "SUPER_ADMIN" || u === "admin") return { success: false, message: "Không thể xóa!" };
  ensureAdminSheet_().deleteRow(Number(row));
  return { success: true, message: "Đã xóa!" };
}

function updateAccountStatus(r, row, s) {
  if (r !== "SUPER_ADMIN") return { success: false };
  ensureAdminSheet_().getRange(Number(row), 5).setValue(s);
  return { success: true, message: "Cập nhật thành công!" };
}

// CẬP NHẬT ĐẦY ĐỦ DANH MỤC ĐƠN VỊ VÀ TỔ CHỨC
const DEFAULT_UNITS_DONG_NAI = [
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường An Lộc", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã An Phước", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã An Viễn", "Xã", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Bảo Vinh", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Bàu Hàm", "Xã", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Biên Hoà", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Bình An", "Xã", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Bình Lộc", "Phường", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Bình Long", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Bình Minh", "Xã", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Bình Phước", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Bình Tân", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Bom Bo", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Bù Đăng", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Cẩm Mỹ", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Nam Cát Tiên", "Xã", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Chơn Thành", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Đa Kia", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Đại Phước", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Đak Lua", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Đak Nhau", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Đăk Ơ", "Xã", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Dầu Giây", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Định Quán", "Xã", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Đồng Phú", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Đồng Tâm", "Xã", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Đồng Xoài", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Gia Kiệm", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Bù Đăng", "Xã", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Hàng Gòn", "Phường", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Hố Nai", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Hưng Phước", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Hưng Thịnh", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã La Ngà", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Lộc Hưng", "Xã", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Lộc Ninh", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Lộc Quang", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Lộc Tấn", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Lộc Thành", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Lộc Thạnh", "Xã", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Long Bình", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Long Hà", "Xã", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Long Hưng", "Phường", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Long Khánh", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Long Phước", "Xã", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Long Thành", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Minh Đức", "Xã", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Minh Hưng", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Nghĩa Trung", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Nha Bích", "Xã", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Nhơn Trạch", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Phú Hoà", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Phú Lâm", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Phú Lý", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Phú Nghĩa", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Phú Riềng", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Phú Trung", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Phú Vinh", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Phước An", "Xã", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Phước Bình", "Phường", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Phước Long", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Phước Sơn", "Xã", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Phước Tân", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Phước Thái", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Sông Ray", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Tà Lài", "Xã", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Tam Hiệp", "Phường", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Tam Phước", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Tân An", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Tân Hưng", "Xã", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Tân Khai", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Tân Lợi", "Xã", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Tân Phú", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Tân Quan", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Tân Tiến", "Xã", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Tân Triều", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Thanh Sơn", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Thiện Hưng", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Thọ Sơn", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Thống Nhất", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Thuận Lợi", "Xã", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Trấn Biên", "Phường", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Trảng Bom", "Phường", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Trảng Dài", "Phường", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Trị An", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Xuân Bắc", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Xuân Định", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Xuân Đông", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Xuân Đường", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Xuân Hoà", "Xã", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Xuân Lập", "Phường", "Có"],
  ["KHOI_C_PHUONG", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Phường Xuân Lộc", "Phường", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Xuân Phú", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Xuân Quế", "Xã", "Có"],
  ["KHOI_C_XA", "MTTQ cấp xã", "Ủy ban MTTQ Việt Nam Xã Xuân Thành", "Xã", "Có"],
  ["KHOI_DOAN_THE", "Các tổ chức chính trị xã hội", "Hội Liên hiệp Phụ nữ thành phố Đồng Nai", "Hội", "Có"],
  ["KHOI_DOAN_THE", "Các tổ chức chính trị xã hội", "Đoàn TNCS Hồ Chí Minh thành phố Đồng Nai", "Đoàn", "Có"],
  ["KHOI_DOAN_THE", "Các tổ chức chính trị xã hội", "Hội Nông dân thành phố Đồng Nai", "Hội", "Có"],
  ["KHOI_DOAN_THE", "Các tổ chức chính trị xã hội", "Hội Cựu chiến binh thành phố Đồng Nai", "Hội", "Có"],
  ["KHOI_DOAN_THE", "Các tổ chức chính trị xã hội", "Liên đoàn Lao động thành phố Đồng Nai", "Hội", "Có"],
  ["KHOI_DOAN_THE", "Các tổ chức chính trị xã hội", "Nhà Văn hóa Thanh thiếu nhi thành phố Đồng Nai", "Đơn vị", "Có"],
  ["KHOI_DOAN_THE", "Các tổ chức chính trị xã hội", "Hội Chữ thập đỏ thành phố Đồng Nai", "Hội", "Có"],
  ["KHOI_DOAN_THE", "Các tổ chức chính trị xã hội", "Hội Khuyến học thành phố Đồng Nai", "Hội", "Có"],
  ["KHOI_DOAN_THE", "Các tổ chức chính trị xã hội", "Hội Người cao tuổi thành phố Đồng Nai", "Hội", "Có"],
  ["KHOI_DOAN_THE", "Các tổ chức chính trị xã hội", "Hội Cựu thanh niên xung phong thành phố Đồng Nai", "Hội", "Có"],
  ["KHOI_DOAN_THE", "Các tổ chức chính trị xã hội", "Hội Nạn nhân chất độc da cam/dioxin thành phố Đồng Nai", "Hội", "Có"]
];

function ensureUnitSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(UNIT_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(UNIT_SHEET_NAME);
  if (sh.getLastRow() <= 100) {
    sh.clear();
    sh.getRange(1, 1, 1, 5).setValues([["Mã nhóm", "Tên nhóm", "Tên đơn vị", "Loại đơn vị", "Hoạt động"]]).setFontWeight("bold");
    sh.getRange(2, 1, DEFAULT_UNITS_DONG_NAI.length, 5).setValues(DEFAULT_UNITS_DONG_NAI);
  }
  return sh;
}

function getUnitCatalog() {
  try {
    const sh = ensureUnitSheet_();
    if (sh.getLastRow() > 1) {
      const list = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getDisplayValues().map((r, i) => {
        let unitName = String(r[2] ? r[2] : "").trim();
        let unitType = String(r[3] ? r[3] : "").trim();
        let groupCode = String(r[0] ? r[0] : "").trim().toUpperCase();
        
        if (unitName.includes("Trảng Bom")) {
          unitName = "Ủy ban MTTQ Việt Nam Phường Trảng Bom";
          unitType = "Phường";
          groupCode = "KHOI_C_PHUONG";
        }
        return {
          row: i + 2,
          groupCode: groupCode,
          groupName: String(r[1] ? r[1] : "MTTQ cấp xã").trim(),
          unitName: unitName,
          unitType: unitType,
          active: !["không", "khong", "0", "false", "ngưng"].includes(String(r[4] ? r[4] : "Có").trim().toLowerCase())
        };
      }).filter(x => x.unitName && x.active); 
      
      if(list.length > 0) return list;
    }
  } catch (err) {}
  
  return DEFAULT_UNITS_DONG_NAI.map((r, i) => {
    let unitName = String(r[2]).trim();
    let unitType = String(r[3]).trim();
    let groupCode = String(r[0]).trim();
    if (unitName.includes("Trảng Bom")) {
        unitName = "Ủy ban MTTQ Việt Nam Phường Trảng Bom";
        unitType = "Phường";
        groupCode = "KHOI_C_PHUONG";
    }
    return {
      row: i + 2, groupCode: groupCode, groupName: r[1],
      unitName: unitName, unitType: unitType, active: true
    };
  });
}

function getAllUnitsAdmin(role) {
  if (role !== "SUPER_ADMIN" && role !== "ADMIN") return { success: false, message: "Không có quyền!" };
  const sh = ensureUnitSheet_();
  if (sh.getLastRow() <= 1) return { success: true, units: [] };
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getDisplayValues();
  const list = data.map((r, i) => ({
    row: i + 2,
    groupCode: String(r[0] || "").trim(),
    groupName: String(r[1] || "").trim(),
    unitName: String(r[2] || "").trim(),
    unitType: String(r[3] || "").trim(),
    active: String(r[4] || "Có").trim()
  }));
  return { success: true, units: list };
}

function addUnitAdmin(role, data) {
  if (role !== "SUPER_ADMIN" && role !== "ADMIN") return { success: false, message: "Không có quyền!" };
  if (!data.unitName) return { success: false, message: "Tên đơn vị không được bỏ trống!" };
  const sh = ensureUnitSheet_();
  sh.appendRow([data.groupCode, data.groupName, data.unitName, data.unitType, data.active]);
  return { success: true, message: "Thêm đơn vị thành công!" };
}

function editUnitAdmin(role, row, data) {
  if (role !== "SUPER_ADMIN" && role !== "ADMIN") return { success: false, message: "Không có quyền!" };
  if (!data.unitName) return { success: false, message: "Tên đơn vị không được bỏ trống!" };
  const sh = ensureUnitSheet_();
  sh.getRange(Number(row), 1, 1, 5).setValues([[data.groupCode, data.groupName, data.unitName, data.unitType, data.active]]);
  return { success: true, message: "Cập nhật đơn vị thành công!" };
}

function deleteUnitAdmin(role, row) {
  if (role !== "SUPER_ADMIN" && role !== "ADMIN") return { success: false, message: "Không có quyền!" };
  const sh = ensureUnitSheet_();
  sh.deleteRow(Number(row));
  return { success: true, message: "Đã xóa đơn vị khỏi hệ thống!" };
}
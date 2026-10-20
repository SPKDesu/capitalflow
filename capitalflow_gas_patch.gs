/**
 * =========================================================================
 * CapitalFlow Google Apps Script - Code.gs (Full Master Version)
 * รองรับครบทั้ง:
 *  1. รายจ่ายด่วน (Expenses)
 *  2. ผ่อนชำระ (Installments)
 *  3. ยืม-คืนด่วน (Loan Tracker)
 * =========================================================================
 */

const LOAN_SHEET_NAME = "ยืม-คืน";

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var data = JSON.parse(e.postData.contents);

    // 1. จัดการระบบยืม-คืนด่วน (Loan Tracker)
    if (data.type === 'loan_sync') {
      return handleSyncLoans(data.loans || []);
    }

    // 2. จัดการรายการผ่อนชำระ (Installments)
    if (data.type === 'installment') {
      var sheet = ss.getSheetByName('Installments');
      if (!sheet) {
        sheet = ss.insertSheet('Installments');
        sheet.appendRow(['Date', 'Item Name', 'Source', 'Total Amount', 'Months', 'Interest Rate', 'Start Month', 'End Month', 'Monthly Amount', 'Status', 'Synced']);
      }
      
      sheet.appendRow([
        data.date,
        data.item,
        data.source,
        data.total,
        data.months,
        data.interest,
        data.start_month,
        data.end_month,
        data.monthly_amount,
        'Active',
        'Pending' // Synced column
      ]);
      
      return ContentService.createTextOutput("SUCCESS");
    }

    // 3. จัดการรายจ่ายด่วน (Expenses)
    var expSheet = ss.getSheetByName('Expenses') || ss.getSheets()[0];
    expSheet.appendRow([
      data.date,
      data.type,
      data.category,
      data.frequency,
      data.item,
      data.amount,
      'Pending'
    ]);
    return ContentService.createTextOutput("SUCCESS");

  } catch (error) {
    return ContentService.createTextOutput("ERROR: " + error.toString());
  }
}

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var action = e ? e.parameter.action : '';

  // ตรวจสอบสถานะการเชื่อมต่อ (Ping/Pre-warm)
  if (action === 'test') {
    return ContentService.createTextOutput("OK");
  }

  // ดึงรายการยืม-คืนด่วน
  if (action === 'get_loans') {
    return handleGetLoans();
  }

  // ซิงค์รายการผ่อนชำระเข้าคอม
  if (action === 'sync_installments') {
    var sheet = ss.getSheetByName('Installments');
    if (!sheet) return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
    
    var data = sheet.getDataRange().getValues();
    var pending = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[10] !== 'Synced') {
        pending.push({
          row_index: i + 1,
          date: row[0],
          item: row[1],
          source: row[2],
          total: row[3],
          months: row[4],
          interest: row[5],
          start_month: row[6],
          end_month: row[7],
          monthly_amount: row[8],
          status: row[9] || 'Active'
        });
      }
    }
    return ContentService.createTextOutput(JSON.stringify(pending)).setMimeType(ContentService.MimeType.JSON);
  }

  // ยืนยันการซิงค์รายการผ่อนชำระ
  if (action === 'confirm_sync_installments') {
    var sheet = ss.getSheetByName('Installments');
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (data[i][10] !== 'Synced') {
          sheet.getRange(i + 1, 11).setValue('Synced');
        }
      }
    }
    return ContentService.createTextOutput("CONFIRMED");
  }

  // ซิงค์รายจ่ายด่วนเข้าคอม (Expenses)
  if (action === 'sync') {
    var sheet = ss.getSheetByName('Expenses') || ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    var pending = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][6] !== 'Synced') {
        pending.push({
          date: data[i][0],
          type: data[i][1],
          category: data[i][2],
          frequency: data[i][3],
          item: data[i][4],
          amount: data[i][5]
        });
        sheet.getRange(i + 1, 7).setValue('Synced');
      }
    }
    return ContentService.createTextOutput(JSON.stringify(pending)).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput("CapitalFlow Web App is running.");
}

/**
 * =========================================================================
 * Helper Functions สำหรับระบบยืม-คืนด่วน (Loan Tracker)
 * =========================================================================
 */

function getOrCreateLoanSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOAN_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOAN_SHEET_NAME);
    const headers = [["วันที่ยืม", "ชื่อคนยืม", "ยอดคงค้าง (บาท)", "หมายเหตุ", "อัปเดตล่าสุด", "ประวัติการคืน/ยืม"]];
    const headerRange = sheet.getRange(1, 1, 1, 6);
    headerRange.setValues(headers);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#1e293b");
    headerRange.setFontColor("#f8fafc");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 110);
    sheet.setColumnWidth(2, 130);
    sheet.setColumnWidth(3, 130);
    sheet.setColumnWidth(4, 180);
    sheet.setColumnWidth(5, 150);
    sheet.setColumnWidth(6, 250);
  }
  return sheet;
}

function handleGetLoans() {
  try {
    const sheet = getOrCreateLoanSheet();
    const lastRow = sheet.getLastRow();
    const loans = [];

    if (lastRow > 1) {
      const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const person = String(row[1] || "").trim();
        const amount = Number(row[2]) || 0;
        if (person && amount > 0) {
          loans.push({
            id: "loan_" + (i + 1) + "_" + Date.now(),
            date: formatDateOnly(row[0]),
            person: person,
            amount: amount,
            note: String(row[3] || ""),
            updatedAt: formatDateTime(row[4]),
            history: String(row[5] || "")
          });
        }
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", loans: loans }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleSyncLoans(loansList) {
  try {
    const sheet = getOrCreateLoanSheet();
    const lastRow = sheet.getLastRow();
    
    // เคลียร์ข้อมูลเก่าทั้งหมด (คงแถวหัวตารางแถว 1 ไว้)
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 6).clearContent();
    }

    if (Array.isArray(loansList) && loansList.length > 0) {
      const rows = loansList.map(item => [
        item.date || "",
        item.person || "",
        Number(item.amount) || 0,
        item.note || "",
        item.updatedAt || new Date().toLocaleString("th-TH"),
        item.history || ""
      ]);

      const targetRange = sheet.getRange(2, 1, rows.length, 6);
      targetRange.setValues(rows);
      sheet.getRange(2, 3, rows.length, 1).setNumberFormat("#,##0.00");
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        status: "success",
        message: "บันทึกรายการยืม-คืนลง Google Sheet สำเร็จ",
        activeCount: Array.isArray(loansList) ? loansList.length : 0
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function formatDateOnly(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Bangkok", "yyyy-MM-dd");
  }
  return String(val).substring(0, 10);
}

function formatDateTime(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Bangkok", "yyyy-MM-dd HH:mm");
  }
  return String(val);
}

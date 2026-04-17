/**
 * =========================================================================
 * ADVANCE & RECEIPT TRACKER - COMPLETE BACKEND WITH API
 * =========================================================================
 * Version: 2.1 - Using google.script.run (No CORS Issues)
 * Purpose: Google Apps Script backend with HtmlService
 * Features: Authentication, CRUD, Reporting, Audit Logging
 * =========================================================================
 */

// ==================== GLOBAL CONSTANTS ====================

const SPREADSHEET_ID = '1mdWN2lHUTGphpalLnobDHpk_BkpATk2BTcg5mHqUQh0';
const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

const SHEET_NAMES = {
  USERS: 'USERS',
  EMPLOYEES: 'EMPLOYEES',
  ADVANCES: 'ADVANCES',
  RECEIPTS: 'RECEIPTS',
  AUDIT_LOG: 'AUDIT_LOG',
  DEPARTMENTS: 'DEPARTMENTS',
  SETTINGS: 'SETTINGS',
  REPORTS_CACHE: 'REPORTS_CACHE',
  APPROVAL_QUEUE: 'APPROVAL_QUEUE'
};

const ROLES = {
  ADMIN: 'ADMIN',
  MANAGER: 'MANAGER',
  EMPLOYEE: 'EMPLOYEE',
  VIEWER: 'VIEWER'
};

// ==================== WEB APP FUNCTIONS ====================

/**
 * Serve the HTML file
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Server function to execute actions from frontend via google.script.run
 * This replaces fetch and eliminates CORS issues completely
 */
function executeServerFunction(action, params) {
  try {
    const handlers = {
      'login': () => login(params.email, params.password),
      'logout': () => logout(),
      'getCurrentUserInfo': () => getCurrentUserInfo(),
      'getAllEmployees': () => getAllEmployees(),
      'recordEmployee': () => recordEmployee(params),
      'deleteEmployee': () => deleteEmployee(params.employeeId),
      'getAllAdvances': () => getAllAdvances(),
      'recordAdvance': () => recordAdvance(params),
      'deleteAdvance': () => deleteAdvance(params.advanceId),
      'getAllReceipts': () => getAllReceipts(),
      'recordReceipt': () => recordReceipt(params),
      'deleteReceipt': () => deleteReceipt(params.receiptId),
      'verifyReceipt': () => verifyReceipt(params.receiptId),
      'getAdjustmentReportData': () => getAdjustmentReportData(params.employeeId),
      'getTotalOutstandingReportData': () => getTotalOutstandingReportData(),
      'getDashboardData': () => getDashboardData()
    };

    if (handlers[action]) {
      return handlers[action]();
    }

    return { success: false, message: 'Unknown action' };
  } catch (e) {
    return {
      success: false,
      message: 'Error: ' + e.message
    };
  }
}

// ==================== UTILITY FUNCTIONS ====================

function getSheet(sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
  return sheet;
}

function getSheetData(sheetName) {
  const sheet = getSheet(sheetName);
  const values = sheet.getDataRange().getDisplayValues();
  
  if (values.length === 0) return [];
  
  const headers = values[0];
  const data = [];
  
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowObject = {};
    for (let j = 0; j < headers.length; j++) {
      rowObject[headers[j]] = row[j];
    }
    data.push(rowObject);
  }
  
  return data;
}

function getCurrentUser() {
  return Session.getEffectiveUser().getEmail();
}

function getUserByEmail(email) {
  const users = getSheetData(SHEET_NAMES.USERS);
  return users.find(u => u.EMAIL === email) || null;
}

function generateUniqueId(prefix) {
  return `${prefix}-${Date.now()}`;
}

function formatDateDMY(date) {
  if (!date) return '';
  if (typeof date === 'string') {
    if (date.match(/^\d{2}\/\d{2}\/\d{4}$/)) return date;
    date = new Date(date);
  }
  
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}/${month}/${year}`;
}

function logAudit(action, recordType, recordId, fieldChanged = '', oldValue = '', newValue = '', reason = '') {
  try {
    const user = getUserByEmail(getCurrentUser());
    if (!user) return;
    
    const sheet = getSheet(SHEET_NAMES.AUDIT_LOG);
    const logId = generateUniqueId('LOG');
    const timestamp = new Date();
    
    const logRow = [
      logId,
      timestamp,
      user.USER_ID,
      user.EMAIL,
      action,
      recordType,
      recordId,
      fieldChanged,
      oldValue,
      newValue,
      reason,
      '',
      'Success'
    ];
    
    sheet.appendRow(logRow);
  } catch (e) {
    console.error('Audit logging error: ' + e.message);
  }
}

// ==================== AUTHENTICATION ====================

function login(email, password) {
  try {
    const users = getSheetData(SHEET_NAMES.USERS);
    const user = users.find(u => u.EMAIL === email);
    
    if (!user) {
      return { success: false, message: 'User not found' };
    }
    
    if (user.STATUS !== 'Active') {
      return { success: false, message: 'Account is not active' };
    }
    
    if (user.PASSWORD_HASH !== password) {
      return { success: false, message: 'Invalid password' };
    }
    
    updateLastLogin(user.USER_ID);
    logAudit('LOGIN', 'USER', user.USER_ID, '', '', '', 'User logged in');
    
    return {
      success: true,
      message: 'Login successful',
      user: {
        userId: user.USER_ID,
        email: user.EMAIL,
        name: user.FULL_NAME,
        role: user.ROLE,
        department: user.DEPARTMENT
      }
    };
  } catch (e) {
    return { success: false, message: 'Login error: ' + e.message };
  }
}

function updateLastLogin(userId) {
  try {
    const sheet = getSheet(SHEET_NAMES.USERS);
    const data = getSheetData(SHEET_NAMES.USERS);
    const rowIndex = data.findIndex(u => u.USER_ID === userId);
    
    if (rowIndex !== -1) {
      const user = data[rowIndex];
      const updateRow = sheet.getRange(rowIndex + 2, 1, 1, 12);
      const values = [
        user.USER_ID,
        user.EMAIL,
        user.PASSWORD_HASH,
        user.FULL_NAME,
        user.ROLE,
        user.DEPARTMENT,
        user.STATUS,
        user.CREATED_DATE,
        new Date(),
        user.PHONE,
        user.CREATED_BY,
        user.NOTES
      ];
      updateRow.setValues([values]);
    }
  } catch (e) {
    console.error('Error updating last login: ' + e.message);
  }
}

function logout() {
  try {
    logAudit('LOGOUT', 'USER', getCurrentUser(), '', '', '', 'User logged out');
    return { success: true, message: 'Logged out successfully' };
  } catch (e) {
    return { success: false, message: 'Logout error: ' + e.message };
  }
}

function getCurrentUserInfo() {
  try {
    const user = getUserByEmail(getCurrentUser());
    if (!user) {
      return { success: false, message: 'User not found' };
    }
    
    return {
      success: true,
      user: {
        userId: user.USER_ID,
        email: user.EMAIL,
        name: user.FULL_NAME,
        role: user.ROLE,
        department: user.DEPARTMENT
      }
    };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ==================== EMPLOYEE MANAGEMENT ====================

function getAllEmployees() {
  try {
    return getSheetData(SHEET_NAMES.EMPLOYEES);
  } catch (e) {
    return [];
  }
}

function getEmployee(employeeId) {
  try {
    const employees = getSheetData(SHEET_NAMES.EMPLOYEES);
    return employees.find(e => e.UNIQUE_EMPLOYEE_ID === employeeId) || null;
  } catch (e) {
    return null;
  }
}

function recordEmployee(employeeData) {
  try {
    const sheet = getSheet(SHEET_NAMES.EMPLOYEES);
    const user = getUserByEmail(getCurrentUser());
    const employees = getSheetData(SHEET_NAMES.EMPLOYEES);
    
    if (employees.find(e => e.EMPLOYEE_NO === employeeData.employeeNo)) {
      return { success: false, message: 'Employee number already exists' };
    }
    
    const employeeId = generateUniqueId('EMP');
    const employeeRow = [
      employeeId,
      employeeData.employeeNo,
      employeeData.firstName,
      employeeData.lastName,
      employeeData.email || '',
      employeeData.phone || '',
      employeeData.department || '',
      employeeData.designation,
      employeeData.office,
      employeeData.managerId || '',
      'Active',
      formatDateDMY(new Date(employeeData.dateOfJoining || new Date())),
      employeeData.salaryBand || '',
      new Date(),
      new Date(),
      user.USER_ID
    ];
    
    sheet.appendRow(employeeRow);
    logAudit('CREATE', 'EMPLOYEE', employeeId, 'EMPLOYEE_NO', '', employeeData.employeeNo);
    
    return { success: true, message: 'Employee recorded successfully', employeeId };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

function deleteEmployee(employeeId) {
  try {
    const sheet = getSheet(SHEET_NAMES.EMPLOYEES);
    const data = getSheetData(SHEET_NAMES.EMPLOYEES);
    const rowIndex = data.findIndex(e => e.UNIQUE_EMPLOYEE_ID === employeeId);
    
    if (rowIndex === -1) {
      return { success: false, message: 'Employee not found' };
    }
    
    sheet.deleteRow(rowIndex + 2);
    logAudit('DELETE', 'EMPLOYEE', employeeId);
    
    return { success: true, message: 'Employee deleted successfully' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ==================== ADVANCE MANAGEMENT ====================

function getAllAdvances() {
  try {
    return getSheetData(SHEET_NAMES.ADVANCES);
  } catch (e) {
    return [];
  }
}

function getEmployeeAdvances(employeeId) {
  try {
    const advances = getSheetData(SHEET_NAMES.ADVANCES);
    return advances.filter(a => a.UNIQUE_EMPLOYEE_ID === employeeId);
  } catch (e) {
    return [];
  }
}

function recordAdvance(advanceData) {
  try {
    const sheet = getSheet(SHEET_NAMES.ADVANCES);
    const user = getUserByEmail(getCurrentUser());
    const employees = getSheetData(SHEET_NAMES.EMPLOYEES);
    const employee = employees.find(e => e.UNIQUE_EMPLOYEE_ID === advanceData.employeeId);
    
    if (!employee) {
      return { success: false, message: 'Employee not found' };
    }
    
    const amount = parseFloat(advanceData.amount);
    if (isNaN(amount) || amount <= 0) {
      return { success: false, message: 'Invalid amount' };
    }
    
    const advanceId = generateUniqueId('ADV');
    const advanceDate = formatDateDMY(new Date(advanceData.advanceDate));
    
    const advanceRow = [
      advanceId,
      advanceData.employeeId,
      employee.EMPLOYEE_NO,
      `${employee.FIRST_NAME} ${employee.LAST_NAME}`,
      employee.DEPARTMENT,
      advanceDate,
      amount,
      advanceData.purpose,
      0,
      amount,
      'PENDING',
      'PENDING',
      '',
      '',
      advanceData.notes || '',
      advanceData.dueDate ? formatDateDMY(new Date(advanceData.dueDate)) : '',
      new Date(),
      new Date(),
      user.USER_ID
    ];
    
    sheet.appendRow(advanceRow);
    logAudit('CREATE', 'ADVANCE', advanceId, 'AMOUNT', '', amount.toString());
    
    return { success: true, message: 'Advance recorded successfully', advanceId };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

function deleteAdvance(advanceId) {
  try {
    const sheet = getSheet(SHEET_NAMES.ADVANCES);
    const data = getSheetData(SHEET_NAMES.ADVANCES);
    const rowIndex = data.findIndex(a => a.ADVANCE_ID === advanceId);
    
    if (rowIndex === -1) {
      return { success: false, message: 'Advance not found' };
    }
    
    sheet.deleteRow(rowIndex + 2);
    logAudit('DELETE', 'ADVANCE', advanceId);
    
    return { success: true, message: 'Advance deleted successfully' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ==================== RECEIPT MANAGEMENT ====================

function getAllReceipts() {
  try {
    return getSheetData(SHEET_NAMES.RECEIPTS);
  } catch (e) {
    return [];
  }
}

function getEmployeeReceipts(employeeId) {
  try {
    const receipts = getSheetData(SHEET_NAMES.RECEIPTS);
    return receipts.filter(r => r.UNIQUE_EMPLOYEE_ID === employeeId);
  } catch (e) {
    return [];
  }
}

function recordReceipt(receiptData) {
  try {
    const sheet = getSheet(SHEET_NAMES.RECEIPTS);
    const user = getUserByEmail(getCurrentUser());
    const employees = getSheetData(SHEET_NAMES.EMPLOYEES);
    const employee = employees.find(e => e.UNIQUE_EMPLOYEE_ID === receiptData.employeeId);
    
    if (!employee) {
      return { success: false, message: 'Employee not found' };
    }
    
    const amount = parseFloat(receiptData.amount);
    if (isNaN(amount) || amount <= 0) {
      return { success: false, message: 'Invalid amount' };
    }
    
    const receiptId = generateUniqueId('REC');
    const receiptDate = formatDateDMY(new Date(receiptData.receiptDate));
    
    const receiptRow = [
      receiptId,
      receiptData.advanceId || 'N/A',
      receiptData.employeeId,
      employee.EMPLOYEE_NO,
      `${employee.FIRST_NAME} ${employee.LAST_NAME}`,
      employee.DEPARTMENT,
      receiptDate,
      amount,
      receiptData.paymentMethod || 'Other',
      receiptData.referenceNo || '',
      receiptData.remarks || '',
      '',
      '',
      'PENDING',
      new Date(),
      new Date(),
      user.USER_ID
    ];
    
    sheet.appendRow(receiptRow);
    
    if (receiptData.advanceId && receiptData.advanceId !== 'N/A') {
      updateAdvanceSettlement(receiptData.advanceId, amount);
    }
    
    logAudit('CREATE', 'RECEIPT', receiptId, 'AMOUNT', '', amount.toString());
    
    return { success: true, message: 'Receipt recorded successfully', receiptId };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

function updateAdvanceSettlement(advanceId, receiptAmount) {
  try {
    const sheet = getSheet(SHEET_NAMES.ADVANCES);
    const data = getSheetData(SHEET_NAMES.ADVANCES);
    const rowIndex = data.findIndex(a => a.ADVANCE_ID === advanceId);
    
    if (rowIndex === -1) return;
    
    const advance = data[rowIndex];
    const currentSettled = parseFloat(advance.SETTLED_AMOUNT) || 0;
    const newSettled = currentSettled + receiptAmount;
    const outstandingAmount = parseFloat(advance.AMOUNT) - newSettled;
    
    let newStatus = 'PARTIAL';
    if (outstandingAmount <= 0) {
      newStatus = 'SETTLED';
    }
    
    const updateRow = sheet.getRange(rowIndex + 2, 1, 1, 19);
    const values = [
      advance.ADVANCE_ID,
      advance.UNIQUE_EMPLOYEE_ID,
      advance.EMPLOYEE_NO,
      advance.EMPLOYEE_NAME,
      advance.DEPARTMENT,
      advance.ADVANCE_DATE,
      advance.AMOUNT,
      advance.PURPOSE,
      newSettled,
      Math.max(0, outstandingAmount),
      newStatus,
      advance.APPROVAL_STATUS,
      advance.APPROVED_BY,
      advance.APPROVAL_DATE,
      advance.NOTES,
      advance.DUE_DATE,
      advance.CREATED_DATE,
      new Date(),
      advance.CREATED_BY
    ];
    
    updateRow.setValues([values]);
  } catch (e) {
    console.error('Error in updateAdvanceSettlement: ' + e.message);
  }
}

function verifyReceipt(receiptId) {
  try {
    const user = getUserByEmail(getCurrentUser());
    
    const sheet = getSheet(SHEET_NAMES.RECEIPTS);
    const data = getSheetData(SHEET_NAMES.RECEIPTS);
    const rowIndex = data.findIndex(r => r.RECEIPT_ID === receiptId);
    
    if (rowIndex === -1) {
      return { success: false, message: 'Receipt not found' };
    }
    
    const oldReceipt = data[rowIndex];
    const updateRow = sheet.getRange(rowIndex + 2, 1, 1, 17);
    const values = [
      receiptId,
      oldReceipt.ADVANCE_ID,
      oldReceipt.UNIQUE_EMPLOYEE_ID,
      oldReceipt.EMPLOYEE_NO,
      oldReceipt.EMPLOYEE_NAME,
      oldReceipt.DEPARTMENT,
      oldReceipt.RECEIPT_DATE,
      oldReceipt.AMOUNT,
      oldReceipt.PAYMENT_METHOD,
      oldReceipt.REFERENCE_NO,
      oldReceipt.REMARKS,
      user.USER_ID,
      formatDateDMY(new Date()),
      'VERIFIED',
      oldReceipt.CREATED_DATE,
      new Date(),
      oldReceipt.CREATED_BY
    ];
    
    updateRow.setValues([values]);
    logAudit('VERIFY', 'RECEIPT', receiptId, 'STATUS', 'PENDING', 'VERIFIED');
    
    return { success: true, message: 'Receipt verified successfully' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

function deleteReceipt(receiptId) {
  try {
    const sheet = getSheet(SHEET_NAMES.RECEIPTS);
    const data = getSheetData(SHEET_NAMES.RECEIPTS);
    const rowIndex = data.findIndex(r => r.RECEIPT_ID === receiptId);
    
    if (rowIndex === -1) {
      return { success: false, message: 'Receipt not found' };
    }
    
    sheet.deleteRow(rowIndex + 2);
    logAudit('DELETE', 'RECEIPT', receiptId);
    
    return { success: true, message: 'Receipt deleted successfully' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ==================== REPORTING ====================

function getAdjustmentReportData(employeeId) {
  try {
    const employees = getSheetData(SHEET_NAMES.EMPLOYEES);
    const advances = getSheetData(SHEET_NAMES.ADVANCES);
    const receipts = getSheetData(SHEET_NAMES.RECEIPTS);
    
    const employee = employees.find(e => e.UNIQUE_EMPLOYEE_ID === employeeId);
    if (!employee) {
      throw new Error('Employee not found');
    }
    
    const employeeAdvances = advances.filter(a => a.UNIQUE_EMPLOYEE_ID === employeeId);
    const employeeReceipts = receipts.filter(r => r.UNIQUE_EMPLOYEE_ID === employeeId);
    
    let totalAdvances = 0;
    employeeAdvances.forEach(a => {
      totalAdvances += parseFloat(a.AMOUNT) || 0;
    });
    
    let totalReceipts = 0;
    employeeReceipts.forEach(r => {
      totalReceipts += parseFloat(r.AMOUNT) || 0;
    });
    
    const netBalance = totalAdvances - totalReceipts;
    
    return {
      success: true,
      employeeName: `${employee.FIRST_NAME} ${employee.LAST_NAME}`,
      employeeNo: employee.EMPLOYEE_NO,
      department: employee.DEPARTMENT,
      totalAdvances,
      totalReceipts,
      netBalance,
      advances: employeeAdvances,
      receipts: employeeReceipts,
      generatedDate: new Date()
    };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

function getTotalOutstandingReportData() {
  try {
    const advances = getSheetData(SHEET_NAMES.ADVANCES);
    const receipts = getSheetData(SHEET_NAMES.RECEIPTS);
    const employees = getSheetData(SHEET_NAMES.EMPLOYEES);
    
    let totalAllAdvances = 0;
    let pendingAdvances = 0;
    advances.forEach(a => {
      if (a.STATUS !== 'REJECTED') {
        totalAllAdvances += parseFloat(a.AMOUNT) || 0;
        if (a.STATUS === 'PENDING') pendingAdvances++;
      }
    });
    
    let totalAllReceipts = 0;
    receipts.forEach(r => {
      if (r.VERIFICATION_STATUS === 'VERIFIED') {
        totalAllReceipts += parseFloat(r.AMOUNT) || 0;
      }
    });
    
    const netTotalOutstanding = totalAllAdvances - totalAllReceipts;
    
    const departmentSummary = {};
    employees.forEach(emp => {
      const dept = emp.DEPARTMENT || 'Unknown';
      if (!departmentSummary[dept]) {
        departmentSummary[dept] = {
          department: dept,
          totalAdvances: 0,
          totalReceipts: 0
        };
      }
      
      const empAdvances = advances.filter(a => a.UNIQUE_EMPLOYEE_ID === emp.UNIQUE_EMPLOYEE_ID);
      const empReceipts = receipts.filter(r => r.UNIQUE_EMPLOYEE_ID === emp.UNIQUE_EMPLOYEE_ID);
      
      empAdvances.forEach(a => {
        departmentSummary[dept].totalAdvances += parseFloat(a.AMOUNT) || 0;
      });
      empReceipts.forEach(r => {
        departmentSummary[dept].totalReceipts += parseFloat(r.AMOUNT) || 0;
      });
    });
    
    return {
      success: true,
      totalAllAdvances,
      totalAllReceipts,
      netTotalOutstanding,
      pendingAdvances,
      advances,
      receipts,
      departmentSummary: Object.values(departmentSummary),
      generatedDate: new Date()
    };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

function getDashboardData() {
  try {
    const user = getUserByEmail(getCurrentUser());
    if (!user) {
      return { success: false, message: 'User not found' };
    }
    
    const advances = getSheetData(SHEET_NAMES.ADVANCES);
    const receipts = getSheetData(SHEET_NAMES.RECEIPTS);
    const employees = getSheetData(SHEET_NAMES.EMPLOYEES);
    
    let totalAdvances = 0;
    let totalReceipts = 0;
    let pendingCount = 0;
    
    advances.forEach(a => {
      totalAdvances += parseFloat(a.AMOUNT) || 0;
      if (a.STATUS === 'PENDING') pendingCount++;
    });
    
    receipts.forEach(r => {
      totalReceipts += parseFloat(r.AMOUNT) || 0;
    });
    
    return {
      success: true,
      userRole: user.ROLE,
      userName: user.FULL_NAME,
      userEmail: user.EMAIL,
      totalAdvances,
      totalReceipts,
      netOutstanding: totalAdvances - totalReceipts,
      pendingCount,
      totalEmployees: employees.length
    };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

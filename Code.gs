const SHEET_NAME = 'Sheet1';
const FIXED_EXPENSES_SHEET_NAME = 'FixedExpenses';
const FIXED_EXPENSE_HEADERS = [
  'ID',
  'Name',
  'Description',
  'Amount',
  'Category',
  'Start Date',
  'End Date',
  'Details JSON',
  'Created At',
  'Updated At'
];
const CATEGORIES = [
  'Food', 'Transport', 'Shopping', 'Bills', 'Market', 'Grocery',
  'Entertainment', 'Health', 'Fuel', 'Other'
];

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getExpenseSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" was not found`);
  return sheet;
}

function getFixedExpenseSheet_(createIfMissing) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(FIXED_EXPENSES_SHEET_NAME);
  if (!sheet && createIfMissing) {
    sheet = spreadsheet.insertSheet(FIXED_EXPENSES_SHEET_NAME);
    sheet.getRange(1, 1, 1, FIXED_EXPENSE_HEADERS.length)
      .setValues([FIXED_EXPENSE_HEADERS])
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function normalizeIsoDate_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(
      value,
      SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),
      'yyyy-MM-dd'
    );
  }
  const text = String(value).trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(text) ? text : '';
}

function validateFixedExpense_(data) {
  const amount = Number(data.amount);
  const item = String(data.item || '').trim();
  const description = String(data.description || '').trim();
  const category = String(data.category || '');
  const startDate = normalizeIsoDate_(data.startDate);
  const endDate = normalizeIsoDate_(data.endDate);

  if (!item || !Number.isFinite(amount) || amount <= 0 || CATEGORIES.indexOf(category) === -1) {
    throw new Error('Invalid fixed expense details');
  }
  if ((data.startDate && !startDate) || (data.endDate && !endDate)) {
    throw new Error('Invalid fixed expense date');
  }
  if (startDate && endDate && startDate > endDate) {
    throw new Error('End date must be on or after start date');
  }

  let details = null;
  if (data.details && typeof data.details === 'object' && !Array.isArray(data.details)) {
    details = data.details;
  }

  return { item, description, amount, category, startDate, endDate, details };
}

function findFixedExpenseRow_(sheet, detailId) {
  if (!sheet || sheet.getLastRow() < 2) return -1;
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  const index = ids.findIndex(row => row[0] === detailId);
  return index === -1 ? -1 : index + 2;
}

function fixedExpenseFromRow_(row) {
  let details = null;
  if (row[7]) {
    try {
      const parsed = JSON.parse(String(row[7]));
      if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) details = parsed;
    } catch (error) {
      console.warn(`Ignoring invalid details JSON for fixed expense ${row[0]}`);
    }
  }
  return {
    detailId: String(row[0]),
    item: String(row[1]),
    description: String(row[2] || ''),
    amount: Number(row[3]),
    category: String(row[4]),
    startDate: normalizeIsoDate_(row[5]),
    endDate: normalizeIsoDate_(row[6]),
    details: details
  };
}

function listFixedExpenses_() {
  const sheet = getFixedExpenseSheet_(false);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, FIXED_EXPENSE_HEADERS.length)
    .getValues()
    .map(fixedExpenseFromRow_)
    .filter(expense => expense.detailId && expense.item && expense.amount > 0);
}

function addFixedExpense_(data) {
  const sheet = getFixedExpenseSheet_(true);
  const fixed = validateFixedExpense_(data);
  const requestedId = String(data.detailId || '');
  const detailId = /^customFixed-[\w-]+$/.test(requestedId)
    ? requestedId
    : `customFixed-${Utilities.getUuid()}`;
  if (findFixedExpenseRow_(sheet, detailId) !== -1) {
    throw new Error('A fixed expense with this ID already exists');
  }
  const now = new Date();
  sheet.appendRow([
    detailId,
    fixed.item,
    fixed.description,
    fixed.amount,
    fixed.category,
    fixed.startDate,
    fixed.endDate,
    fixed.details ? JSON.stringify(fixed.details) : '',
    now,
    now
  ]);
  return { detailId, ...fixed };
}

function editFixedExpense_(data) {
  const sheet = getFixedExpenseSheet_(false);
  const detailId = String(data.detailId || '');
  const rowNumber = findFixedExpenseRow_(sheet, detailId);
  if (rowNumber === -1) throw new Error('Fixed expense was not found');

  const fixed = validateFixedExpense_(data);
  const createdAt = sheet.getRange(rowNumber, 9).getValue() || new Date();
  sheet.getRange(rowNumber, 1, 1, FIXED_EXPENSE_HEADERS.length).setValues([[
    detailId,
    fixed.item,
    fixed.description,
    fixed.amount,
    fixed.category,
    fixed.startDate,
    fixed.endDate,
    fixed.details ? JSON.stringify(fixed.details) : '',
    createdAt,
    new Date()
  ]]);
  return { detailId, ...fixed };
}

function deleteFixedExpense_(data) {
  const sheet = getFixedExpenseSheet_(false);
  const detailId = String(data.detailId || '');
  const rowNumber = findFixedExpenseRow_(sheet, detailId);
  if (rowNumber === -1) throw new Error('Fixed expense was not found');
  sheet.deleteRow(rowNumber);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'addFixedExpense') {
      return jsonResponse({ status: 'success', fixedExpense: addFixedExpense_(data) });
    }
    if (data.action === 'editFixedExpense') {
      return jsonResponse({ status: 'success', fixedExpense: editFixedExpense_(data) });
    }
    if (data.action === 'deleteFixedExpense') {
      deleteFixedExpense_(data);
      return jsonResponse({ status: 'success', message: 'Fixed expense deleted' });
    }

    const sheet = getExpenseSheet_();
    if (data.action === 'delete') {
      const row = Number(data.rowNumber ?? data.row);
      if (!Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) {
        return jsonResponse({ status: 'error', message: 'Invalid row' });
      }
      sheet.deleteRow(row);
      return jsonResponse({ status: 'success', message: 'Row deleted' });
    }

    if (data.action === 'edit') {
      const row = Number(data.rowNumber ?? data.row);
      const amount = Number(data.amount);
      if (!Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) {
        return jsonResponse({ status: 'error', message: 'Invalid row' });
      }
      if (!data.date || !data.item || !data.category || !Number.isFinite(amount) || amount <= 0) {
        return jsonResponse({ status: 'error', message: 'Invalid expense details' });
      }

      sheet.getRange(row, 1, 1, 4).setValues([[
        data.date,
        String(data.item).trim(),
        String(data.category),
        amount
      ]]);
      return jsonResponse({ status: 'success', message: 'Expense updated' });
    }

    const amount = Number(data.amount);
    if (!data.date || !data.item || !data.category || !Number.isFinite(amount) || amount <= 0) {
      return jsonResponse({ status: 'error', message: 'Invalid expense details' });
    }
    sheet.appendRow([data.date, String(data.item).trim(), String(data.category), amount]);
    return jsonResponse({ status: 'success' });
  } catch (error) {
    return jsonResponse({ status: 'error', message: error.message });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function doGet() {
  try {
    const sheet = getExpenseSheet_();
    const rows = sheet.getDataRange().getValues().slice(1);
    const data = rows.map((row, index) => ({
      rowNumber: index + 2,
      date: row[0],
      item: row[1],
      category: row[2],
      amount: row[3]
    }));
    return jsonResponse({
      status: 'success',
      data: data,
      fixedExpenses: listFixedExpenses_(),
      features: { fixedExpenses: true }
    });
  } catch (error) {
    return jsonResponse({ status: 'error', message: error.message });
  }
}

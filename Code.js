// ═══════════════════════════════════════════════════════════════════
//  SignQuote — Code.gs  (FIXED)
// ═══════════════════════════════════════════════════════════════════

const ACCOUNTS_SPREADSHEET_ID = '1ClnO3Z6xGXa2V6AVijXIfqIJDlI2zLVVKRJdEeE4ebM';
const ACCOUNTS_SHEET_NAME     = 'Accounts';
const SESSION_SHEET           = 'Sessions';
const SHEET_QUOTATIONS        = 'Quotations';
const SHEET_DATABASE          = 'Database';
const TARP_SHEET              = 'Tarp Quotations';   // ✅ FIX #1 — constant was missing
const RECEIPT_SHEET           = 'Receipt Quotations';
const BOOKBIND_SHEET          = 'Bookbind Quotations';
const FRAME_SHEET             = 'Frame Quotations';
const TSHIRT_SHEET            = 'Tshirt Quotations';
const MUG_SHEET               = 'Mug Quotations';
const CUSTOMER_SHEET          = 'Customer Quotations';
const CUSTOMER_SS_ID          = '1SKuJe0ocRgiTLMOtqp9gerdOkGDiP-86Z6QiWXu4R5Y';

function setupMainSSId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('MAIN_SS_ID', ss.getId());
  Logger.log('Main SS ID saved: ' + ss.getId());
}

function getMainSS_() {
  const id = PropertiesService.getScriptProperties().getProperty('MAIN_SS_ID');
  if (id) return SpreadsheetApp.openById(id);
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getCustomerSS_() {
  return SpreadsheetApp.openById(CUSTOMER_SS_ID);
}

// ══════════════════════════════════════════════════════════════════
//  doGet
// ══════════════════════════════════════════════════════════════════
function doGet(e) {
  const token  = String(e?.parameter?.token || '').trim();
  const page   = String(e?.parameter?.page  || '').trim().toLowerCase();
  const appUrl = ScriptApp.getService().getUrl();

  // ── PUBLIC CUSTOMER PAGES (no auth required) ──
  if (page === 'customer') {
    const tpl  = HtmlService.createTemplateFromFile('Customer');
    tpl.appUrl = appUrl;
    tpl.injectedProduct = String(e?.parameter?.product || '').trim().toLowerCase();
    return tpl.evaluate()
      .setTitle('Get a Quote — Ormoc Printshoppe')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'customer-tarp') {
    const tpl  = HtmlService.createTemplateFromFile('CustomerTarp');
    tpl.appUrl = appUrl;
    return tpl.evaluate()
      .setTitle('Tarpaulin Quote — Ormoc Printshoppe')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  // Product-specific customer portal links: customer-bookbind, customer-receipt, etc.
  if (page.indexOf('customer-') === 0) {
    const tpl  = HtmlService.createTemplateFromFile('Customer');
    tpl.appUrl = appUrl;
    tpl.injectedProduct = page.slice(9);   // 'customer-bookbind' -> 'bookbind'
    return tpl.evaluate()
      .setTitle('Get a Quote — Ormoc Printshoppe')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (!token) return serveLogin_(appUrl);

  const session = getSessionData_(token);
  if (!session) return serveLogin_(appUrl);

  const role = (session.role || 'sales').toLowerCase();

  if (page === 'index') {
    return serveWithToken_('Index', 'Quotation System — Quotation', token, appUrl);
  }
  if (page === 'dashboard') {
    return serveWithToken_('Dashboard', 'Quotation System — Dashboard', token, appUrl);
  }
  if (page === 'quotation') {
    const qn  = String(e?.parameter?.qn || '').trim();
    const tpl = HtmlService.createTemplateFromFile('Quotation');
    tpl.injectedToken    = token;
    tpl.injectedQuoteNum = qn;
    return tpl.evaluate()
      .setTitle('Quotation — Ormoc Printshoppe')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'tarpaulin') {
    return serveWithToken_('Tarpauline', 'Quotation System — Tarpaulin', token, appUrl);
  }
  if (page === 'receipt') {
    return serveWithToken_('Receipt', 'Quotation System — Receipt', token, appUrl);
  }
  if (page === 'bookbind') {
    return serveWithToken_('Bookbind', 'Quotation System — Bookbinding', token, appUrl);
  }
  if (page === 'frame') {
    return serveWithToken_('Frame', 'Quotation System — Frame', token, appUrl);
  }
  if (page === 'tshirt') {
    return serveWithToken_('Tshirt', 'Quotation System — T-Shirt', token, appUrl);
  }
  if (page === 'mug') {
    return serveWithToken_('Mug', 'Quotation System — Mug', token, appUrl);
  }
  if (role === 'sales' || role === 'staff') {
    return serveWithToken_('Index', 'Quotation System — Quotation', token, appUrl);
  }
  return serveWithToken_('Dashboard', 'Quotation System — Dashboard', token, appUrl);
}

function serveLogin_(appUrl) {
  const tpl = HtmlService.createTemplateFromFile('Login');
  tpl.appUrl = appUrl;
  return tpl.evaluate()
    .setTitle('Quotation System — Login')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function serveWithToken_(file, title, token, appUrl) {
  const session = getSessionData_(token);
  const tpl = HtmlService.createTemplateFromFile(file);
  tpl.injectedToken = token;
  tpl.appUrl        = appUrl || ScriptApp.getService().getUrl();
  tpl.userName      = session ? session.name : '';
  return tpl.evaluate()
    .setTitle(title)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ══════════════════════════════════════════════════════════════════
//  LOGIN
// ══════════════════════════════════════════════════════════════════
function loginUser(username, password) {
  try {
    const ss    = getMainSS_();
    const sheet = ss.getSheetByName(ACCOUNTS_SHEET_NAME);
    if (!sheet) return { success: false, message: 'Accounts sheet not found.' };

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row        = data[i];
      const storedUser = String(row[0] || '').trim();
      const storedPass = String(row[1] || '').trim();
      const role       = String(row[2] || 'sales').trim();
      const name       = String(row[3] || storedUser).trim();

      if (!storedUser) continue;

      const inputUser = username.trim().toLowerCase();
      const matched   = inputUser === storedUser.toLowerCase()
                     || inputUser === storedUser.split('@')[0].toLowerCase();

      if (!matched) continue;
      if (String(password) !== String(storedPass)) {
        return { success: false, message: 'Invalid username or password.' };
      }

      const token = createSession_(storedUser, role, name);
      logLogin_(storedUser, role);
      return { success: true, username: storedUser, role, name, token };
    }

    return { success: false, message: 'Invalid username or password.' };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

// ══════════════════════════════════════════════════════════════════
//  SESSION
// ══════════════════════════════════════════════════════════════════
function createSession_(username, role, name) {
  const ss  = getMainSS_();
  let sheet = ss.getSheetByName(SESSION_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SESSION_SHEET);
    sheet.appendRow(['Token','Username','Role','Expires','Active','Name','Created']);
    sheet.getRange(1,1,1,7).setBackground('#0A0A0A').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  try {
    const rows = sheet.getDataRange().getValues();
    const now  = new Date();
    const toDelete = [];
    for (let i = rows.length - 1; i >= 1; i--) {
      const active = String(rows[i][4] || '').trim().toUpperCase();
      const expiry = rows[i][3] instanceof Date ? rows[i][3] : new Date(rows[i][3]);
      if (active === 'FALSE' || (expiry < now && !isNaN(expiry))) {
        toDelete.push(i + 1);
      }
    }
    toDelete.forEach(r => sheet.deleteRow(r));
  } catch(e) {}

  const token   = Utilities.getUuid();
  const now     = new Date();
  const expires = new Date(now.getTime() + 8 * 60 * 60 * 1000);
  sheet.appendRow([token, username, role, expires, 'TRUE', name || username, now]);
  return token;
}

function getSessionData_(token) {
  try {
    if (!token) return null;
    const cleanToken = String(token).trim();
    const ss    = getMainSS_();
    const sheet = ss.getSheetByName(SESSION_SHEET);
    if (!sheet) return null;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const storedToken = String(data[i][0] || '').trim();
      if (storedToken !== cleanToken) continue;
      const isActive = String(data[i][4] || '').trim().toUpperCase();
      if (isActive !== 'TRUE') continue;
      return {
        token:    cleanToken,
        username: String(data[i][1] || '').trim(),
        role:     String(data[i][2] || '').replace(/[^a-zA-Z_]/g,'').trim() || 'sales',
        name:     String(data[i][5] || data[i][1] || '').trim(),
      };
    }
    return null;
  } catch (err) {
    Logger.log('getSessionData_ error: ' + err.message);
    return null;
  }
}

function getCurrentSession(token) {
  return getSessionData_(token);
}

function logoutUser(token) {
  try {
    const ss    = getMainSS_();
    const sheet = ss.getSheetByName(SESSION_SHEET);
    if (!sheet) return { success: false };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === token) {
        sheet.getRange(i + 1, 5).setValue('FALSE');
        return { success: true };
      }
    }
    return { success: false };
  } catch (err) { return { success: false }; }
}

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

function getActiveUserEmail() {
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch(e) { return ''; }
}

function googleLogin() {
  try {
    const email = Session.getActiveUser().getEmail().toLowerCase().trim();
    if (!email) return { success: false, message: 'Could not detect your Google account.' };

    const ss    = getMainSS_();
    const sheet = ss.getSheetByName(ACCOUNTS_SHEET_NAME);
    if (!sheet) return { success: false, message: 'Accounts sheet not found.' };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row     = data[i];
      const rawRole = String(row[0] || '').trim();
      const emails  = String(row[1] || '').split(',').map(e => e.trim().toLowerCase()).filter(Boolean);

      if (!rawRole || emails.length === 0) continue;
      if (!emails.includes(email)) continue;

      let role = 'sales';
      if (/tech/i.test(rawRole))            role = 'tech_support';
      else if (/admin/i.test(rawRole))       role = 'admin';
      else if (/sales|staff/i.test(rawRole)) role = 'sales';

      const name  = email.split('@')[0];
      const token = createSession_(email, role, name);
      logLogin_(email, role);
      return { success: true, username: email, role, name, token };
    }

    return { success: false, message: 'Account "' + email + '" not authorized. Contact your administrator.' };
  } catch(err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

function loginByEmail(email) {
  try {
    const inputEmail = String(email || '').toLowerCase().trim();
    if (!inputEmail) return { success: false, message: 'Please enter your email.' };

    const ss    = getMainSS_();
    const sheet = ss.getSheetByName(ACCOUNTS_SHEET_NAME);
    if (!sheet) return { success: false, message: 'Accounts sheet not found.' };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row     = data[i];
      const rawRole = String(row[0] || '').trim();
      const emails  = String(row[1] || '').split(',').map(e => e.trim().toLowerCase()).filter(Boolean);

      if (!rawRole || emails.length === 0) continue;
      if (!emails.includes(inputEmail)) continue;

      let role = 'sales';
      if (/tech/i.test(rawRole))            role = 'tech_support';
      else if (/admin/i.test(rawRole))       role = 'admin';
      else if (/sales|staff/i.test(rawRole)) role = 'sales';

      const name  = inputEmail.split('@')[0];
      const token = createSession_(inputEmail, role, name);
      logLogin_(inputEmail, role);
      return { success: true, username: inputEmail, role, name, token };
    }

    return { success: false, message: 'Email "' + inputEmail + '" is not authorized. Contact your administrator.' };
  } catch(err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

function logLogin_(username, role) {
  const ss  = getMainSS_();
  let sheet = ss.getSheetByName('LoginLog');
  if (!sheet) {
    sheet = ss.insertSheet('LoginLog');
    sheet.appendRow(['Timestamp','Username','Role']);
    sheet.getRange(1,1,1,3).setBackground('#0A0A0A').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([new Date(), username, role]);
}

function cleanExpiredSessions() {
  const ss    = getMainSS_();
  const sheet = ss.getSheetByName(SESSION_SHEET);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const now  = new Date();
  for (let i = data.length - 1; i >= 1; i--) {
    const expiry   = new Date(data[i][3]);
    const isActive = String(data[i][4] || '').trim().toUpperCase();
    if (isActive === 'FALSE' || expiry < now) sheet.deleteRow(i + 1);
  }
}

function clearAllSessions() {
  const ss = getMainSS_();
  const sheet = ss.getSheetByName('Sessions');
  if (!sheet) { Logger.log('NO SESSIONS SHEET'); return; }
  const lastRow = sheet.getLastRow();
  Logger.log('Sessions rows: ' + lastRow);
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
    Logger.log('All sessions cleared!');
  }
}

// ══════════════════════════════════════════════════════════════════
//  TEST FUNCTIONS
// ══════════════════════════════════════════════════════════════════
function testGetDashboard() {
  const ss    = getMainSS_();
  const sheet = ss.getSheetByName(SESSION_SHEET);
  if (!sheet) { Logger.log('NO SESSIONS SHEET'); return; }
  const data = sheet.getDataRange().getValues();
  Logger.log('Sessions count: ' + (data.length - 1));
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][4]).trim().toUpperCase() === 'TRUE') {
      const token = String(data[i][0]).trim();
      Logger.log('Testing token: ' + token);
      const result = getDashboardData(token);
      Logger.log('Result: ' + JSON.stringify(result ? { name: result.name, role: result.role, quotes: result.quotes?.length } : null));
      return;
    }
  }
  Logger.log('No active sessions found');
}

function testGoogleLogin() {
  const result = googleLogin();
  Logger.log(JSON.stringify(result));
}

function testEmail() {
  Logger.log(Session.getActiveUser().getEmail());
  Logger.log(Session.getEffectiveUser().getEmail());
}

// ══════════════════════════════════════════════════════════════════
//  DASHBOARD DATA
// ══════════════════════════════════════════════════════════════════
function getDashboardData(token) {
  try {
    if (!token) return null;

    const session = getSessionData_(token);
    if (!session) {
      Logger.log('getDashboardData: session not found for token: ' + token);
      return null;
    }

    const ss   = getMainSS_();
    const role = session.role.toLowerCase();

    // ── SIGNAGE QUOTES ──────────────────────────────────────────
    let quotes = [];
    const sigSheet = ss.getSheetByName(SHEET_QUOTATIONS);
    if (sigSheet) {
      const data = sigSheet.getDataRange().getValues();
      quotes = data.slice(1).filter(r => r[0]).map(row => {
        let dateStr = '';
        try { dateStr = row[1] ? new Date(row[1]).toISOString() : ''; } catch(e) {}
        return {
          quoteNum:            String(row[0]  || ''),
          date:                dateStr,
          clientName:          String(row[2]  || ''),
          contact:             String(row[3]  || ''),
          email:               String(row[4]  || ''),
          address:             String(row[5]  || ''),
          delivery:            String(row[6]  || ''),
          signageType:         String(row[7]  || ''),
          lighting:            String(row[8]  || ''),
          material:            String(row[9]  || ''),
          width:               row[10] || 0,
          height:              row[11] || 0,
          sqft:                row[12] || 0,
          rate:                row[13] || 0,
          mounting:            String(row[16] || ''),
          baseAmount:          row[21] || 0,
          mountSurcharge:      row[22] || 0,
          complexitySurcharge: row[23] || 0,
          designFee:           row[24] || 0,
          totalAmount:         row[25] || 0,
          downpayment:         row[26] || 0,
          balance:             row[27] || 0,
          status:              String(row[28] || 'Pending'),
          approvedBy:          String(row[29] || ''),
          salesStaff:          String(row[30] || ''),
          dateNeeded:          String(row[31] || ''),
          addonDesign:         String(row[32] || ''),
          addonDesignFee:      row[33] || 0,
          addonRush:           String(row[34] || ''),
          addonRushFee:        row[35] || 0,
          addonElec:           String(row[36] || ''),
          addonElecFee:        row[37] || 0,
          addonTransport:      String(row[38] || ''),
          addonTransportFee:   row[39] || 0,
          addonTransportLocation: String(row[40] || ''),
          paymentTermLabel: String(row[41] || ''),
          paymentTermValue: String(row[41] || '').includes('No Down') ? 0 : String(row[41] || '').includes('25%') ? 0.25 : String(row[41] || '').includes('Full') ? 1 : String(row[41] || '') === '' ? 0.5 : 0.5,
          taxType:      String(row[42] || 'non-vat'),
          taxAmount:    parseFloat(row[43]) || 0,
          quoteType:           'signage',
        };
      });
    }

    // ── TARP QUOTES ─────────────────────────────────────────────
    let tarpQuotes = [];
    const tarpSheet = ss.getSheetByName(TARP_SHEET);
    if (tarpSheet) {
      const tdata = tarpSheet.getDataRange().getValues();
      tarpQuotes = tdata.slice(1).filter(r => r[0]).map(row => {
        let dateStr = '';
        try { dateStr = row[1] ? new Date(row[1]).toISOString() : ''; } catch(e) {}
        return {
          quoteNum:     String(row[0]  || ''),
          date:         dateStr,
          clientName:   String(row[2]  || ''),
          contact:      String(row[3]  || ''),
          email:        String(row[4]  || ''),
          width:        row[5]  || 0,
          height:       row[6]  || 0,
          sqft:         row[7]  || 0,
          quantity:     row[8]  || 1,
          totalSqft:    row[9]  || 0,
          eyelet:       String(row[10] || ''),
          printLayout:  String(row[11] || ''),
          rushOrder:    String(row[12] || ''),
          designCharge: String(row[13] || ''),
          rate:         row[14] || 0,
          rushFee:      row[15] || 0,
          designFee:    row[16] || 0,
          baseAmount:   row[17] || 0,
          rushFeeAmt:   row[18] || 0,
          designFeeAmt: row[19] || 0,
          totalAmount:  row[20] || 0,
          balance:      row[21] || 0,
          dateNeeded:   String(row[22] || ''),
          status:       String(row[23] || 'Pending'),
          approvedBy:   String(row[24] || ''),
          salesStaff:   String(row[25] || ''),
          paymentTermLabel: String(row[26] || ''),
          paymentTermValue: String(row[26] || '').includes('No Down') ? 0 : String(row[26] || '').includes('25%') ? 0.25 : String(row[26] || '').includes('Full') ? 1 : 0.5,
          items: (function() { try { const j = String(row[27]||''); if (!j || j==='[]') return []; return JSON.parse(j); } catch(e) { return []; } })(),
          taxType:      String(row[28] || 'non-vat'),
          taxAmount:    parseFloat(row[29]) || 0,
          quoteType:    'tarpaulin',
          signageType:  'Tarpaulin',
          address: '', delivery: '', lighting: '', material: '',
          mounting: '', mountSurcharge: 0, complexitySurcharge: 0,
          addonDesign:'', addonDesignFee:0,
          addonRush: String(row[12] || ''),    addonRushFee: row[18] || 0,
          addonElec:'',   addonElecFee:0,
          addonTransport:'', addonTransportFee:0,
        };
      });
    }

    // ── RECEIPT QUOTES ──────────────────────────────────────────
    // ✅ FIX #2 — receiptData was referenced but never loaded; now properly loaded here
    let receiptQuotes = [];
    const receiptSheet = ss.getSheetByName(RECEIPT_SHEET);
    if (receiptSheet) {
      const receiptData = receiptSheet.getDataRange().getValues();
      receiptQuotes = receiptData.slice(1).filter(r => r[0]).map(row => ({
  quoteNum:    String(row[0]  || ''),
  date:        row[1] ? new Date(row[1]).toISOString() : '',
  clientName:  String(row[2]  || ''),
  contact:     String(row[6]  || ''),
  email:       String(row[5]  || ''),
  signageType: 'Receipt — ' + String(row[8] || ''),
  size: row[9] instanceof Date ? '' : String(row[9] || ''),
  paperType:   String(row[11] || ''),
  paperColors: String(row[12] || ''),
  perforation: String(row[13] || ''),
  numbering:   String(row[14] || ''),
  quantity:    row[16] || 0,
  dateNeeded:  String(row[17] || ''),
  totalAmount: row[18] || 0,
  paymentTermLabel: String(row[22] || ''),
  paymentTermValue: String(row[22] || '').includes('No Down') ? 0 : String(row[22] || '').includes('25%') ? 0.25 : String(row[22] || '').includes('Full') ? 1 : 0.5,
  taxType:      String(row[23] || 'non-vat'),
  taxAmount:    parseFloat(row[24]) || 0,
  rushOrder:    String(row[25] || ''),
  rushFee:      parseFloat(row[26]) || 0,
  addonRush:    String(row[25] || ''),
  addonRushFee: parseFloat(row[26]) || 0,
  status:      String(row[19] || 'Pending'),
  approvedBy:  String(row[20] || ''),
  salesStaff:  String(row[20] || ''),
}));
    }

    // ── BOOKBIND QUOTES ─────────────────────────────────────────
    let bookbindQuotes = [];
    const bookbindSheet = ss.getSheetByName(BOOKBIND_SHEET);
    if (bookbindSheet) {
      const bdata = bookbindSheet.getDataRange().getValues();
      // If first row is data (no header), start from 0; otherwise skip header row 0
      const bStart = bdata.length > 0 && String(bdata[0][0]).startsWith('BQ-') ? 0 : 1;
      bookbindQuotes = bdata.slice(bStart).filter(r => r[0] && String(r[0]).startsWith('BQ-')).map(row => {
        let dateStr = '';
        try { dateStr = row[1] ? new Date(row[1]).toISOString() : ''; } catch(e) {}
        const ptLabel = String(row[27] || '');
        return {
          quoteNum:         String(row[0]  || ''),
          date:             dateStr,
          clientName:       String(row[2]  || ''),
          contact:          String(row[3]  || ''),
          email:            String(row[4]  || ''),
          dateNeeded:       String(row[5]  || ''),
          signageType:      'Bookbinding — ' + String(row[6] || ''),
          totalAmount:      row[20] || 0,
          downpayment:      row[21] || 0,
          balance:          row[22] || 0,
          salesStaff:       String(row[24] || ''),
          status:           String(row[25] || 'Pending'),
          approvedBy:       String(row[26] || ''),
          paymentTermLabel: ptLabel,
          paymentTermValue: ptLabel.includes('No Down') ? 0 : ptLabel.includes('25%') ? 0.25 : ptLabel.includes('Full') ? 1 : 0.5,
          taxType:          String(row[28] || 'non-vat'),
          taxAmount:        parseFloat(row[29]) || 0,
          quoteType:        'bookbind',
          address: '', delivery: '', lighting: '', material: '',
          mounting: '', mountSurcharge: 0, complexitySurcharge: 0,
          addonDesign: '', addonDesignFee: 0,
          addonRush: '', addonRushFee: 0,
          addonElec: '', addonElecFee: 0,
          addonTransport: '', addonTransportFee: 0,
        };
      });
    }

    // ── FRAME QUOTES ────────────────────────────────────────────
    let frameQuotes = [];
    const frameSheet = ss.getSheetByName(FRAME_SHEET);
    if (frameSheet) {
      const fdata = frameSheet.getDataRange().getValues();
      // If first row is data (no header), start from 0; otherwise skip header row
      const fStart = fdata.length > 0 && String(fdata[0][0]).startsWith('FQ-') ? 0 : 1;
      frameQuotes = fdata.slice(fStart).filter(r => r[0] && String(r[0]).startsWith('FQ-')).map(row => {
        let dateStr = '';
        try { dateStr = row[1] ? new Date(row[1]).toISOString() : ''; } catch(e) {}
        const ptLabel = String(row[21] || '');
        return {
          quoteNum:         String(row[0]  || ''),
          date:             dateStr,
          clientName:       String(row[2]  || ''),
          contact:          String(row[3]  || ''),
          email:            String(row[4]  || ''),
          dateNeeded:       String(row[5]  || ''),
          width:            row[6]  || 0,
          height:           row[7]  || 0,
          quantity:         row[8]  || 1,
          sqft:             row[9]  || 0,
          totalSqft:        row[10] || 0,
          matting:          String(row[11] || ''),
          rate:             row[12] || 0,
          baseAmount:       row[13] || 0,
          totalAmount:      row[14] || 0,
          downpayment:      row[15] || 0,
          balance:          row[16] || 0,
          notes:            String(row[17] || ''),
          salesStaff:       String(row[18] || ''),
          status:           String(row[19] || 'Pending'),
          approvedBy:       String(row[20] || ''),
          paymentTermLabel: ptLabel,
          paymentTermValue: ptLabel.includes('No Down') ? 0 : ptLabel.includes('25%') ? 0.25 : ptLabel.includes('Full') ? 1 : 0.5,
          taxType:          String(row[22] || 'non-vat'),
          taxAmount:        parseFloat(row[23]) || 0,
          rushOrder:        String(row[24] || ''),
          rushFee:          parseFloat(row[25]) || 0,
          quoteType:        'frame',
          signageType:      'Frame — ' + String(row[11] || ''),
          address: '', delivery: '', lighting: '', material: '',
          mounting: '', mountSurcharge: 0, complexitySurcharge: 0,
          addonDesign: '', addonDesignFee: 0,
          addonRush: String(row[24] || ''),  addonRushFee: parseFloat(row[25]) || 0,
          addonElec: '', addonElecFee: 0,
          addonTransport: '', addonTransportFee: 0,
        };
      });
    }

    // ── T-SHIRT QUOTES ──────────────────────────────────────────
    let tshirtQuotes = [];
    const tshirtSheet = ss.getSheetByName(TSHIRT_SHEET);
    if (tshirtSheet) {
      const tsdata = tshirtSheet.getDataRange().getValues();
      const tStart = tsdata.length > 0 && String(tsdata[0][0]).startsWith('SH-') ? 0 : 1;
      tshirtQuotes = tsdata.slice(tStart).filter(r => r[0] && String(r[0]).startsWith('SH-')).map(row => {
        let dateStr = '';
        try { dateStr = row[1] ? new Date(row[1]).toISOString() : ''; } catch(e) {}
        const ptLabel = String(row[29] || '');
        const printType   = String(row[6]  || '');
        const garmentType = String(row[10] || '');
        const shirtChoice = String(row[9]  || '');
        const sigType = printType === 'Full Sublimation'
          ? 'T-Shirt — Full Sub ' + garmentType
          : 'T-Shirt — ' + printType + (row[7] ? ' ' + row[7] : '');
        return {
          quoteNum:         String(row[0]  || ''),
          date:             dateStr,
          clientName:       String(row[2]  || ''),
          contact:          String(row[3]  || ''),
          email:            String(row[4]  || ''),
          dateNeeded:       String(row[5]  || ''),
          printType:        printType,
          logoSize:         String(row[7]  || ''),
          hasOwnShirt:      String(row[8]  || ''),
          shirtChoice:      shirtChoice,
          garmentType:      garmentType,
          shirtColor:       String(row[11] || ''),
          sizeBreakdown:    String(row[12] || ''),
          quantity:         row[13] || 1,
          printPrice:       parseFloat(row[14]) || 0,
          shirtPrice:       parseFloat(row[15]) || 0,
          unitPrice:        parseFloat(row[16]) || 0,
          baseAmount:       parseFloat(row[17]) || 0,
          rushOrder:        String(row[18] || ''),
          rushFee:          parseFloat(row[19]) || 0,
          designService:    String(row[20] || ''),
          designFee:        parseFloat(row[21]) || 0,
          totalAmount:      parseFloat(row[22]) || 0,
          downpayment:      parseFloat(row[23]) || 0,
          balance:          parseFloat(row[24]) || 0,
          notes:            String(row[25] || ''),
          salesStaff:       String(row[26] || ''),
          status:           String(row[27] || 'Pending'),
          approvedBy:       String(row[28] || ''),
          paymentTermLabel: ptLabel,
          paymentTermValue: ptLabel.includes('No Down') ? 0 : ptLabel.includes('25%') ? 0.25 : ptLabel.includes('Full') ? 1 : 0.5,
          taxType:          String(row[30] || 'non-vat'),
          taxAmount:        parseFloat(row[31]) || 0,
          quoteType:        'tshirt',
          signageType:      sigType,
          address: '', delivery: '', lighting: '', material: '',
          mounting: '', mountSurcharge: 0, complexitySurcharge: 0,
          addonDesign: String(row[20] || ''), addonDesignFee: parseFloat(row[21]) || 0,
          addonRush:   String(row[18] || ''), addonRushFee:   parseFloat(row[19]) || 0,
          addonElec: '', addonElecFee: 0,
          addonTransport: '', addonTransportFee: 0,
        };
      });
    }

    // ── MUG QUOTES ──────────────────────────────────────────────
    let mugQuotes = [];
    const mugSheet = ss.getSheetByName(MUG_SHEET);
    if (mugSheet) {
      const mdata  = mugSheet.getDataRange().getValues();
      const mStart = mdata.length > 0 && String(mdata[0][0]).startsWith('MUG-') ? 0 : 1;
      mugQuotes = mdata.slice(mStart).filter(r => r[0] && String(r[0]).startsWith('MUG-')).map(row => {
        let dateStr = '';
        try { dateStr = row[1] ? new Date(row[1]).toISOString() : ''; } catch(e) {}
        const ptLabel = String(row[23] || '');
        return {
          quoteNum:         String(row[0]  || ''),
          date:             dateStr,
          clientName:       String(row[2]  || ''),
          contact:          String(row[3]  || ''),
          email:            String(row[4]  || ''),
          dateNeeded:       String(row[5]  || ''),
          mugType:          String(row[6]  || ''),
          quantity:         row[7]  || 0,
          unitPrice:        parseFloat(row[10]) || 0,
          baseAmount:       parseFloat(row[11]) || 0,
          rushOrder:        String(row[12] || ''),
          rushFee:          parseFloat(row[13]) || 0,
          designService:    String(row[14] || ''),
          designFee:        parseFloat(row[15]) || 0,
          totalAmount:      parseFloat(row[16]) || 0,
          downpayment:      parseFloat(row[17]) || 0,
          balance:          parseFloat(row[18]) || 0,
          notes:            String(row[19] || ''),
          salesStaff:       String(row[20] || ''),
          status:           String(row[21] || 'Pending'),
          approvedBy:       String(row[22] || ''),
          paymentTermLabel: ptLabel,
          paymentTermValue: ptLabel.includes('No Down') ? 0 : ptLabel.includes('25%') ? 0.25 : ptLabel.includes('Full') ? 1 : 0.5,
          taxType:          String(row[24] || 'non-vat'),
          taxAmount:        parseFloat(row[25]) || 0,
          quoteType:        'mug',
          signageType:      'Mug — ' + String(row[6] || ''),
          address: '', delivery: '', lighting: '', material: '',
          mounting: '', mountSurcharge: 0, complexitySurcharge: 0,
          addonDesign: String(row[14] || ''), addonDesignFee: parseFloat(row[15]) || 0,
          addonRush:   String(row[12] || ''), addonRushFee:   parseFloat(row[13]) || 0,
          addonElec: '', addonElecFee: 0,
          addonTransport: '', addonTransportFee: 0,
        };
      });
    }

    // ── COMBINE & FILTER ────────────────────────────────────────
    const allQuotes = [...quotes, ...tarpQuotes, ...receiptQuotes, ...bookbindQuotes, ...frameQuotes, ...tshirtQuotes, ...mugQuotes];

    const filtered = (role === 'sales' || role === 'staff')
      ? allQuotes.filter(q => q.salesStaff === session.username || q.salesStaff === session.name)
      : allQuotes;

    return { name: session.name || session.username || 'User', username: session.username, role: session.role || 'sales', quotes: filtered };

  } catch(err) {
    Logger.log('getDashboardData error: ' + err.message);
    return null;
  }
}

// ══════════════════════════════════════════════════════════════════
//  GET QUOTE FOR PDF
// ══════════════════════════════════════════════════════════════════
function getQuoteForPDF(token, quoteNum) {
  const session = getSessionData_(token);
  if (!session) return null;

  const ss = getMainSS_();
  const qn = String(quoteNum || '').trim();
  if (!qn) return null;

  // ── Receipt ─────────────────────────────────────────────────────
  if (qn.startsWith('RQ-')) {
    const sheet = ss.getSheetByName(RECEIPT_SHEET);
    if (!sheet) return null;
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() !== qn) continue;
      const r = rows[i];
      return {
        quoteNum:         qn,
        date:             r[1] ? new Date(r[1]).toISOString() : '',
        clientName:       String(r[2]  || ''),
        contact:          String(r[6]  || ''),
        email:            String(r[5]  || ''),
        signageType:      'Receipt — ' + String(r[8] || ''),
        size:             r[9] instanceof Date ? '' : String(r[9] || ''),
        paperType:        String(r[11] || ''),
        paperColors:      String(r[12] || ''),
        perforation:      String(r[13] || ''),
        numbering:        String(r[14] || ''),
        quantity:         r[16] || 0,
        dateNeeded:       String(r[17] || ''),
        totalAmount:      r[18] || 0,
        status:           String(r[19] || 'Pending'),
        salesStaff:       String(r[20] || ''),
        paymentTermLabel: String(r[22] || ''),
        paymentTermValue: String(r[22]||'').includes('No Down') ? 0
                        : String(r[22]||'').includes('25%')   ? 0.25
                        : String(r[22]||'').includes('Full')  ? 1 : 0.5,
        taxType:   String(r[23] || 'non-vat'),
        taxAmount: parseFloat(r[24]) || 0,
        rushOrder: String(r[25] || ''),
        rushFee:   parseFloat(r[26]) || 0,
        addonRush: String(r[25] || ''),
        addonRushFee: parseFloat(r[26]) || 0,
      };
    }
    return null;
  }

  // ── Tarpaulin ────────────────────────────────────────────────────
  if (qn.startsWith('TQ-')) {
    const sheet = ss.getSheetByName(TARP_SHEET);
    if (!sheet) return null;
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() !== qn) continue;
      const r = rows[i];
      let dateStr = '';
      try { dateStr = r[1] ? new Date(r[1]).toISOString() : ''; } catch(e) {}
      return {
        quoteNum:         qn,
        date:             dateStr,
        clientName:       String(r[2]  || ''),
        contact:          String(r[3]  || ''),
        email:            String(r[4]  || ''),
        signageType:      'Tarpaulin',
        width:            r[5]  || 0,
        height:           r[6]  || 0,
        sqft:             r[7]  || 0,
        quantity:         r[8]  || 1,
        eyelet:           String(r[10] || ''),
        baseAmount:       r[17] || 0,
        addonRushFee:     r[18] || 0,
        designFee:        r[19] || 0,
        totalAmount:      r[20] || 0,
        balance:          r[21] || 0,
        dateNeeded:       String(r[22] || ''),
        status:           String(r[23] || 'Pending'),
        salesStaff:       String(r[25] || ''),
        paymentTermLabel: String(r[26] || ''),
        paymentTermValue: String(r[26]||'').includes('No Down') ? 0
                        : String(r[26]||'').includes('25%')   ? 0.25
                        : String(r[26]||'').includes('Full')  ? 1 : 0.5,
        items: (function() { try { const j = String(r[27]||''); if (!j || j==='[]') return []; return JSON.parse(j); } catch(e) { return []; } })(),
        taxType:   String(r[28] || 'non-vat'),
        taxAmount: parseFloat(r[29]) || 0,
        address: '', delivery: '', lighting: '', material: '',
        mounting: '', mountFee: 0, complexitySurcharge: 0,
        addonDesign: '', addonDesignFee: 0,
        addonRush: String(r[12] || ''),
        addonElec: '', addonElecFee: 0,
        addonTransport: '', addonTransportFee: 0,
      };
    }
    return null;
  }

  // ── Bookbind ─────────────────────────────────────────────────────
  if (qn.startsWith('BQ-')) {
    const sheet = ss.getSheetByName(BOOKBIND_SHEET);
    if (!sheet) return null;
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() !== qn) continue;
      const r = rows[i];
      let dateStr = '';
      try { dateStr = r[1] ? new Date(r[1]).toISOString() : ''; } catch(e) {}
      const ptLabel = String(r[27] || '');
      return {
        quoteNum:              qn,
        date:                  dateStr,
        clientName:            String(r[2]  || ''),
        contact:               String(r[3]  || ''),
        email:                 String(r[4]  || ''),
        dateNeeded:            String(r[5]  || ''),
        address:               '',
        delivery:              '',
        signageType:           'Bookbinding — ' + String(r[6] || ''),
        bindType:              String(r[6]  || ''),
        withLettering:         String(r[7]  || ''),
        bindQty:               r[8]  || 1,
        bindPages:             r[9]  || 0,
        paperSize:             String(r[10] || ''),
        orientation:           String(r[11] || ''),
        bindingSide:           String(r[12] || ''),
        coverColor:            String(r[13] || ''),
        printedMaterialsReady: String(r[14] || ''),
        printingType:          String(r[15] || ''),
        printingFee:           parseFloat(r[16]) || 0,
        rushOrder:             String(r[17] || ''),
        bindingPrice:          parseFloat(r[18]) || 0,
        rushFee:               parseFloat(r[19]) || 0,
        totalAmount:           parseFloat(r[20]) || 0,
        salesStaff:            String(r[24] || ''),
        status:                String(r[25] || 'Pending'),
        approvedBy:            String(r[26] || ''),
        paymentTermLabel:      ptLabel,
        paymentTermValue:      ptLabel.includes('No Down') ? 0 : ptLabel.includes('25%') ? 0.25 : ptLabel.includes('Full') ? 1 : 0.5,
        taxType:               String(r[28] || 'non-vat'),
        taxAmount:             parseFloat(r[29]) || 0,
        quoteType:             'bookbind',
        baseAmount: 0, mountFee: 0, complexitySurcharge: 0, designFee: 0,
        addonDesign: '', addonDesignFee: 0,
        addonRush: '', addonRushFee: 0,
        addonElec: '', addonElecFee: 0,
        addonTransport: '', addonTransportFee: 0,
      };
    }
    return null;
  }

  // ── Frame ────────────────────────────────────────────────────────
  if (qn.startsWith('FQ-')) {
    const sheet = ss.getSheetByName(FRAME_SHEET);
    if (!sheet) return null;
    const rows = sheet.getDataRange().getValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() !== qn) continue;
      const r = rows[i];
      let dateStr = '';
      try { dateStr = r[1] ? new Date(r[1]).toISOString() : ''; } catch(e) {}
      const ptLabel = String(r[21] || '');
      return {
        quoteNum:         qn,
        date:             dateStr,
        clientName:       String(r[2]  || ''),
        contact:          String(r[3]  || ''),
        email:            String(r[4]  || ''),
        dateNeeded:       String(r[5]  || ''),
        width:            r[6]  || 0,
        height:           r[7]  || 0,
        quantity:         r[8]  || 1,
        sqft:             r[9]  || 0,
        totalSqft:        r[10] || 0,
        matting:          String(r[11] || ''),
        rate:             r[12] || 0,
        baseAmount:       r[13] || 0,
        totalAmount:      r[14] || 0,
        downpayment:      r[15] || 0,
        balance:          r[16] || 0,
        notes:            String(r[17] || ''),
        salesStaff:       String(r[18] || ''),
        status:           String(r[19] || 'Pending'),
        approvedBy:       String(r[20] || ''),
        paymentTermLabel: ptLabel,
        paymentTermValue: ptLabel.includes('No Down') ? 0
                        : ptLabel.includes('25%')   ? 0.25
                        : ptLabel.includes('Full')  ? 1 : 0.5,
        taxType:          String(r[22] || 'non-vat'),
        taxAmount:        parseFloat(r[23]) || 0,
        rushOrder:        String(r[24] || ''),
        rushFee:          parseFloat(r[25]) || 0,
        quoteType:        'frame',
        signageType:      'Frame — ' + String(r[11] || ''),
        address: '', delivery: '', lighting: '', material: '',
        mounting: '', mountFee: 0, complexitySurcharge: 0, designFee: 0,
        addonDesign: '', addonDesignFee: 0,
        addonRush: String(r[24] || ''), addonRushFee: parseFloat(r[25]) || 0,
        addonElec: '', addonElecFee: 0,
        addonTransport: '', addonTransportFee: 0,
      };
    }
    return null;
  }

  // ── T-Shirt ──────────────────────────────────────────────────────
  if (qn.startsWith('SH-')) {
    const sheet = ss.getSheetByName(TSHIRT_SHEET);
    if (!sheet) return null;
    const rows = sheet.getDataRange().getValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() !== qn) continue;
      const r = rows[i];
      let dateStr = '';
      try { dateStr = r[1] ? new Date(r[1]).toISOString() : ''; } catch(e) {}
      const ptLabel = String(r[29] || '');
      const printType   = String(r[6]  || '');
      const garmentType = String(r[10] || '');
      const sigType = printType === 'Full Sublimation'
        ? 'T-Shirt — Full Sub ' + garmentType
        : 'T-Shirt — ' + printType + (r[7] ? ' ' + r[7] : '');
      return {
        quoteNum:         qn,
        date:             dateStr,
        clientName:       String(r[2]  || ''),
        contact:          String(r[3]  || ''),
        email:            String(r[4]  || ''),
        dateNeeded:       String(r[5]  || ''),
        printType:        printType,
        logoSize:         String(r[7]  || ''),
        hasOwnShirt:      String(r[8]  || ''),
        shirtChoice:      String(r[9]  || ''),
        garmentType:      garmentType,
        shirtColor:       String(r[11] || ''),
        sizeBreakdown:    String(r[12] || ''),
        quantity:         r[13] || 1,
        printPrice:       parseFloat(r[14]) || 0,
        shirtPrice:       parseFloat(r[15]) || 0,
        unitPrice:        parseFloat(r[16]) || 0,
        baseAmount:       parseFloat(r[17]) || 0,
        rushOrder:        String(r[18] || ''),
        rushFee:          parseFloat(r[19]) || 0,
        designService:    String(r[20] || ''),
        designFee:        parseFloat(r[21]) || 0,
        totalAmount:      parseFloat(r[22]) || 0,
        downpayment:      parseFloat(r[23]) || 0,
        balance:          parseFloat(r[24]) || 0,
        notes:            String(r[25] || ''),
        salesStaff:       String(r[26] || ''),
        status:           String(r[27] || 'Pending'),
        approvedBy:       String(r[28] || ''),
        paymentTermLabel: ptLabel,
        paymentTermValue: ptLabel.includes('No Down') ? 0
                        : ptLabel.includes('25%')   ? 0.25
                        : ptLabel.includes('Full')  ? 1 : 0.5,
        taxType:          String(r[30] || 'non-vat'),
        taxAmount:        parseFloat(r[31]) || 0,
        quoteType:        'tshirt',
        signageType:      sigType,
        address: '', delivery: '', lighting: '', material: '',
        mounting: '', mountFee: 0, complexitySurcharge: 0,
        addonDesign: String(r[20] || ''), addonDesignFee: parseFloat(r[21]) || 0,
        addonRush:   String(r[18] || ''), addonRushFee:   parseFloat(r[19]) || 0,
        addonElec: '', addonElecFee: 0,
        addonTransport: '', addonTransportFee: 0,
      };
    }
    return null;
  }

  // ── Signage ──────────────────────────────────────────────────────
  const sheet = ss.getSheetByName(SHEET_QUOTATIONS);
  if (!sheet) return null;
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() !== qn) continue;
    const r = rows[i];
    let dateStr = '';
    try { dateStr = r[1] ? new Date(r[1]).toISOString() : ''; } catch(e) {}
    return {
      quoteNum:               qn,
      date:                   dateStr,
      clientName:             String(r[2]  || ''),
      contact:                String(r[3]  || ''),
      email:                  String(r[4]  || ''),
      address:                String(r[5]  || ''),
      delivery:               String(r[6]  || ''),
      signageType:            String(r[7]  || ''),
      lighting:               String(r[8]  || ''),
      material:               String(r[9]  || ''),
      width:                  r[10] || 0,
      height:                 r[11] || 0,
      sqft:                   r[12] || 0,
      rate:                   r[13] || 0,
      mounting:               String(r[16] || ''),
      baseAmount:             r[21] || 0,
      mountFee:               r[22] || 0,
      complexitySurcharge:    r[23] || 0,
      designFee:              r[24] || 0,
      totalAmount:            r[25] || 0,
      downpayment:            r[26] || 0,
      balance:                r[27] || 0,
      status:                 String(r[28] || 'Pending'),
      salesStaff:             String(r[30] || ''),
      dateNeeded:             String(r[31] || ''),
      addonDesign:            String(r[32] || ''),
      addonDesignFee:         r[33] || 0,
      addonRush:              String(r[34] || ''),
      addonRushFee:           r[35] || 0,
      addonElec:              String(r[36] || ''),
      addonElecFee:           r[37] || 0,
      addonTransport:         String(r[38] || ''),
      addonTransportFee:      r[39] || 0,
      addonTransportLocation: String(r[40] || ''),
      paymentTermLabel:       String(r[41] || ''),
      paymentTermValue: String(r[41]||'').includes('No Down') ? 0
                      : String(r[41]||'').includes('25%')   ? 0.25
                      : String(r[41]||'').includes('Full')  ? 1 : 0.5,
      taxType:   String(r[42] || 'non-vat'),
      taxAmount: parseFloat(r[43]) || 0,
    };
  }
  return null;
}

// ══════════════════════════════════════════════════════════════════
//  UPDATE QUOTE STATUS
// ══════════════════════════════════════════════════════════════════
// ✅ FIX #3 — removed nested duplicate function; receipt branch now properly placed
function updateQuoteStatus(token, quoteNum, status) {
  const session = getSessionData_(token);
  if (!session) throw new Error('Session expired.');
  const role = session.role.toLowerCase();
  if (role === 'sales' || role === 'staff') throw new Error('Access denied.');

  const ss          = getMainSS_();
  const isTarp      = String(quoteNum).startsWith('TQ-');
  const isReceipt   = String(quoteNum).startsWith('RQ-');
  const isBookbind  = String(quoteNum).startsWith('BQ-');
  const isFrame     = String(quoteNum).startsWith('FQ-');
  const isTshirt    = String(quoteNum).startsWith('SH-');

  if (isReceipt) {
    const sheet = ss.getSheetByName(RECEIPT_SHEET);
    if (!sheet) throw new Error('Receipt Quotations sheet not found.');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === quoteNum) {
        sheet.getRange(i+1, 20).setValue(status);        // col T (index 19) = Status
        sheet.getRange(i+1, 21).setValue(session.name + ' — ' + new Date().toLocaleString('en-PH')); // col U = Approved By
        const color = status === 'Approved' ? '#E6FFF3' : status === 'Rejected' ? '#FFF0F0' : '#FFFFFF';
        sheet.getRange(i+1, 1, 1, 27).setBackground(color);
        return { success: true };
      }
    }
    throw new Error('Receipt order not found: ' + quoteNum);
  }

  if (isBookbind) {
    const sheet = ss.getSheetByName(BOOKBIND_SHEET);
    if (!sheet) throw new Error('Bookbind Quotations sheet not found.');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === quoteNum) {
        sheet.getRange(i+1, 26).setValue(status);
        sheet.getRange(i+1, 27).setValue(session.name + ' — ' + new Date().toLocaleString('en-PH'));
        const color = status === 'Approved' ? '#E6FFF3' : status === 'Rejected' ? '#FFF0F0' : '#FFFFFF';
        sheet.getRange(i+1, 1, 1, 27).setBackground(color);
        return { success: true };
      }
    }
    throw new Error('Bookbind order not found: ' + quoteNum);
  }

  if (isFrame) {
    const sheet = ss.getSheetByName(FRAME_SHEET);
    if (!sheet) throw new Error('Frame Quotations sheet not found.');
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === quoteNum) {
        sheet.getRange(i+1, 20).setValue(status);  // col T = Status
        sheet.getRange(i+1, 21).setValue(session.name + ' — ' + new Date().toLocaleString('en-PH'));  // col U = Approved By
        const color = status === 'Approved' ? '#E6FFF3' : status === 'Rejected' ? '#FFF0F0' : '#FFFFFF';
        sheet.getRange(i+1, 1, 1, 26).setBackground(color);
        return { success: true };
      }
    }
    throw new Error('Frame order not found: ' + quoteNum);
  }

  if (isTshirt) {
    const sheet = ss.getSheetByName(TSHIRT_SHEET);
    if (!sheet) throw new Error('Tshirt Quotations sheet not found.');
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === quoteNum) {
        sheet.getRange(i+1, 28).setValue(status);  // col AB = Status
        sheet.getRange(i+1, 29).setValue(session.name + ' — ' + new Date().toLocaleString('en-PH'));  // col AC = Approved By
        const color = status === 'Approved' ? '#E6FFF3' : status === 'Rejected' ? '#FFF0F0' : '#FFFFFF';
        sheet.getRange(i+1, 1, 1, 32).setBackground(color);
        return { success: true };
      }
    }
    throw new Error('T-Shirt order not found: ' + quoteNum);
  }

  if (isTarp) {
    const sheet = ss.getSheetByName(TARP_SHEET);
    if (!sheet) throw new Error('Tarp Quotations sheet not found.');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === quoteNum) {
        sheet.getRange(i+1, 24).setValue(status);        // ✅ FIX #4 — col X (index 23) = Status (was col 21)
        sheet.getRange(i+1, 25).setValue(session.name + ' — ' + new Date().toLocaleString('en-PH')); // col Y = Approved By
        const color = status === 'Approved' ? '#E6FFF3' : status === 'Rejected' ? '#FFF0F0' : '#FFFFFF';
        sheet.getRange(i+1, 1, 1, 26).setBackground(color);
        return { success: true };
      }
    }
    throw new Error('Quote not found: ' + quoteNum);
  }

  // Signage
  const sheet = ss.getSheetByName(SHEET_QUOTATIONS);
  if (!sheet) throw new Error('Quotations sheet not found.');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === quoteNum) {
      sheet.getRange(i+1, 29).setValue(status);          // col AC = Status
      sheet.getRange(i+1, 30).setValue(session.name + ' — ' + new Date().toLocaleString('en-PH'));
      const color = status === 'Approved' ? '#E6FFF3' : status === 'Rejected' ? '#FFF0F0' : '#FFFFFF';
      sheet.getRange(i+1, 1, 1, 31).setBackground(color);
      return { success: true };
    }
  }
  throw new Error('Quote not found: ' + quoteNum);
}

// ══════════════════════════════════════════════════════════════════
//  USER MANAGEMENT
// ══════════════════════════════════════════════════════════════════
function getUsers(token) {
  const session = getSessionData_(token);
  if (!session || session.role.toLowerCase() !== 'tech_support') throw new Error('Access denied.');
  const ss    = getMainSS_();
  const sheet = ss.getSheetByName(ACCOUNTS_SHEET_NAME);
  if (!sheet) throw new Error('Accounts sheet not found.');
  const data = sheet.getDataRange().getValues();
  return data.slice(1).filter(r => r[0]).map(r => ({
    username: String(r[0]),
    name:     String(r[3] || r[0]),
    role:     String(r[2] || 'sales'),
  }));
}

function addUser(token, userData) {
  const session = getSessionData_(token);
  if (!session || session.role.toLowerCase() !== 'tech_support') throw new Error('Access denied.');
  const ss    = getMainSS_();
  const sheet = ss.getSheetByName(ACCOUNTS_SHEET_NAME);
  if (!sheet) throw new Error('Accounts sheet not found.');
  sheet.appendRow([userData.username, userData.password, userData.role, userData.name]);
  return { success: true };
}

function deleteUser(token, username) {
  const session = getSessionData_(token);
  if (!session || session.role.toLowerCase() !== 'tech_support') throw new Error('Access denied.');
  const ss    = getMainSS_();
  const sheet = ss.getSheetByName(ACCOUNTS_SHEET_NAME);
  if (!sheet) throw new Error('Accounts sheet not found.');
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === username) { sheet.deleteRow(i+1); break; }
  }
  return { success: true };
}

// ══════════════════════════════════════════════════════════════════
//  LIVE PRICING
// ══════════════════════════════════════════════════════════════════
function getPricing() {
  const ss      = getMainSS_();
  const dbSheet = ss.getSheetByName(SHEET_DATABASE);
  if (!dbSheet) throw new Error('Database sheet not found!');

  const dbData = dbSheet.getDataRange().getValues();
  let externalId = null;
  for (let i = 0; i < dbData.length; i++) {
    if (String(dbData[i][0]).trim() === 'PriceDatabase') {
      externalId = String(dbData[i][1]).trim();
      break;
    }
  }
  if (!externalId) throw new Error('PriceDatabase not found in Database sheet!');

  const match = externalId.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  const ssId  = match ? match[1] : externalId;

  const extSS    = SpreadsheetApp.openById(ssId);
  const extSheet = extSS.getSheetByName('Signage New');
  if (!extSheet) throw new Error('Signage New sheet not found!');

  const data                 = extSheet.getDataRange().getValues();
  const products             = [];
  const rates                = {};
  const minCharge            = {};
  const minArea              = {};
  const mountingOptions      = [];
  const complexitySurcharges = [];
  const designOptions        = [];

  function parseRate(val) {
    if (typeof val === 'number' && val > 0) return val;
    const cleaned = String(val).replace(/[₱,\s]/g, '');
    const m = cleaned.match(/(\d+(?:\.\d+)?)/);
    return m ? parseFloat(m[1]) : NaN;
  }

  function parseSurcharge(raw) {
    const s = String(raw || '').trim();
    if (!s) return null;
    const pctM = s.match(/\+?(\d+(?:\.\d+)?)\s*%/);
    if (pctM) return { pct: parseFloat(pctM[1]), flat: 0, display: s, type: 'percent' };
    const flatNum = parseRate(s);
    if (!isNaN(flatNum) && flatNum > 0) return { pct: 0, flat: flatNum, display: s, type: 'flat' };
    if (/by quotation/i.test(s)) return { pct: 0, flat: 0, display: s, type: 'quotation' };
    return null;
  }

  let inMounting = false, inComplexity = false, inProducts = true;
  let currentSection = 'Other';

  for (let i = 0; i < data.length; i++) {
    const row     = data[i];
    const rawName = String(row[0] || '').trim();
    if (!rawName) continue;
    const firstLine = rawName.split('\n')[0].trim();

    if (/^SECTION\s+1/i.test(firstLine) || /SINGLE FACE SIGNAGE/i.test(firstLine)) {
      inProducts=true; inMounting=false; inComplexity=false;
      currentSection = 'Single Face Signage (Metal Frame)'; continue;
    }
    if (/^SECTION\s+2/i.test(firstLine) || /DOUBLE FACE SIGNAGE/i.test(firstLine)) {
      inProducts=true; inMounting=false; inComplexity=false;
      currentSection = 'Double Face Signage (Metal Frame)'; continue;
    }
    if (/^SECTION\s+3/i.test(firstLine) || /3D BUILD UP SIGNAGE/i.test(firstLine)) {
      inProducts=true; inMounting=false; inComplexity=false;
      currentSection = '3D Build Up Signage'; continue;
    }
    if (/SECTION\s+4/i.test(firstLine)) { inProducts=false; inMounting=false; inComplexity=false; continue; }
    if (/^INSTALLATION\s*&\s*MOUNTING/i.test(firstLine)) { inMounting=true; inComplexity=false; inProducts=false; continue; }
    if (/^COMPLEXITY\s+SURCHARGES/i.test(firstLine))     { inComplexity=true; inMounting=false; inProducts=false; continue; }
    // End of a collecting section — stop adding to products / mounting / complexity
    if (/^(ELECTRICAL\s*&|TRANSPORT\s*&|GENERAL\s+CONDITION|DESIGN\s*&|RUSH\s*&)/i.test(firstLine)) { inProducts=false; inMounting=false; inComplexity=false; continue; }
    if (/^(RUSH|ELECTRICAL|TRANSPORT|GENERAL|DESIGN\s*&|CLUSTER|TIMER|Product\s*\/|Type|Service|Item|ORMOC|Effective:)/i.test(firstLine)) continue;
    if (/^(Product|Description|Remarks)/i.test(firstLine)) continue;

    if (inMounting) {
      const parsed = parseSurcharge(String(row[1] || '').trim());
      if (!firstLine || /^(Mounting|Type)/i.test(firstLine)) continue;
      mountingOptions.push({ label: firstLine, fee: parsed?parsed.flat:0, pct: parsed?parsed.pct:0, notes: String(row[3]||'').trim(), isQuotation: parsed?parsed.type==='quotation':false, rateDisplay: String(row[1]||'').trim() });
      continue;
    }
    if (inComplexity) {
      const parsed = parseSurcharge(String(row[1] || '').trim());
      if (!parsed) continue;
      complexitySurcharges.push({ label: firstLine, surchargeDisplay: String(row[1]||'').trim(), pct: parsed.pct, flat: parsed.flat, type: parsed.type });
      continue;
    }
    if (!inProducts) continue;

    const rate = parseRate(row[1]);
    if (isNaN(rate) || rate <= 0) continue;
    const charge = parseRate(row[2]) || 0;
    const areaM  = String(row[3]||'').replace(/[₱,]/g,'').match(/(\d+(?:\.\d+)?)/);
    const area   = areaM ? parseFloat(areaM[1]) : 0;
    rates[firstLine] = rate; minCharge[firstLine] = charge; minArea[firstLine] = area;
    products.push({ name: firstLine, rate, minCharge: charge, minArea: area, section: currentSection });
  }

  const section4 = { design: [], rush: [], electrical: [], transport: [] };
  try {
    let s4Section = '';
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const col0 = String(row[0]||'').trim(), col1 = String(row[1]||'').trim();
      const col2 = String(row[2]||'').trim(), col3 = String(row[3]||'').trim();
      if (!col0) continue;
      const firstLine = col0.split('\n')[0].trim();
      if (/SECTION\s+4/i.test(col0)) { s4Section = ''; continue; }
      if (i < 29) continue;
      if (/^DESIGN\s*&\s*ARTWORK/i.test(firstLine))     { s4Section = 'design'; continue; }
      if (/^RUSH\s*&\s*PRODUCTION/i.test(firstLine))     { s4Section = 'rush'; continue; }
      if (/^ELECTRICAL\s*&\s*LIGHTING/i.test(firstLine)) { s4Section = 'electrical'; continue; }
      if (/^TRANSPORT\s*&\s*DELIVERY/i.test(firstLine))  { s4Section = 'transport'; continue; }
      if (/^(COMPLEXITY|INSTALLATION|GENERAL|MOUNTING)/i.test(firstLine)) { s4Section = ''; continue; }
      if (/^(Service|Type|Item|Cluster)/i.test(firstLine) && !col2) continue;
      if (!s4Section || !col1) continue;
      const nameLine = firstLine;
      const noteLine = col0.includes('\n') ? col0.split('\n').slice(1).join(' ').trim() : (col3||'');
      if (s4Section === 'design') {
        const rate = parseRate(col1), rateMax = col1.includes('–') ? parseRate(col1.split('–')[1]) : 0;
        if (!isNaN(rate) && rate > 0) section4.design.push({ label: nameLine, rate, rateMax: rateMax||0, notes: noteLine });
      } else if (s4Section === 'rush') {
        const pctM = col1.match(/\+?(\d+)%/), flatM = col1.match(/\+?₱?(\d+(?:,\d+)?)\s*\/\s*day/i);
        if (pctM) section4.rush.push({ label: nameLine, pct: parseFloat(pctM[1]), flat: 0, notes: noteLine });
        else if (flatM) section4.rush.push({ label: nameLine, pct: 0, flat: parseFloat(flatM[1].replace(',','')), notes: noteLine });
      } else if (s4Section === 'electrical') {
        const perMeter = /per.*meter/i.test(col1)||/\/\s*meter/i.test(col1);
        const rateNum  = parseRate(col1.replace(/\/\s*(meter|m)/i,'')), rateMax = col1.includes('–') ? parseRate(col1.split('–')[1]) : 0;
        if (!isNaN(rateNum)||perMeter) section4.electrical.push({ label: nameLine, flat: perMeter?0:(rateNum||0), flatMax: rateMax||0, perUnit: perMeter, unitRate: perMeter?(parseRate(col1)||100):0, notes: noteLine });
      } else if (s4Section === 'transport') {
        if (/surcharge/i.test(col1)) continue;
        const isFree = /FREE/i.test(col1), flatNum = isFree?0:parseRate(col1.split('–')[0]), flatMax = col1.includes('–')?parseRate(col1.split('–')[1]):0;
        section4.transport.push({ label: nameLine, flat: isFree?0:(flatNum||0), flatMax: flatMax||0, notes: col2||noteLine });
      }
    }
    section4.rush.unshift({ label: 'No rush — standard lead time (5–7 days)', pct: 0, flat: 0, notes: '' });
  } catch(s4err) { Logger.log('Section4 parse error: ' + s4err.message); }

  return { products, rates, minCharge, minArea, mountingOptions, complexitySurcharges, designOptions, section4, designFee: 800 };
}

// ══════════════════════════════════════════════════════════════════
//  SAVE TARP QUOTATION
// ══════════════════════════════════════════════════════════════════
function saveTarpQuotation(data) {
  const ss  = getMainSS_();
  let sheet = ss.getSheetByName(TARP_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(TARP_SHEET);
    const headers = [
      'Quote #', 'Date', 'Client Name', 'Contact', 'Email',
      'Width (ft)', 'Height (ft)', 'Area/pc (sqft)', 'Quantity', 'Total Sqft',
      'Eyelet', 'Print Layout', 'Rush Order', 'Design Charge',
      'Rate/sqft', 'Rush Fee', 'Design Fee',
      'Base Amount', 'Rush Fee Amt', 'Design Fee Amt',
      'TOTAL AMOUNT', 'Balance', 'Date Needed', 'Status',
      'Approved By', 'Sales Staff', 'Payment Term', 'Items JSON',
      'Tax Type', 'Tax Amount',
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#E8151B').setFontColor('#fff')
      .setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
  }

  const lastRow  = sheet.getLastRow();
  const quoteNum = `TQ-${String(lastRow).padStart(4,'0')}`;

  const session2  = data.token ? getSessionData_(data.token) : null;
  const staffName = session2 ? (session2.username || session2.name) : (data.salesStaff || '');

  const w         = parseFloat(data.width)  || 0;
  const h         = parseFloat(data.height) || 0;
  const qty       = parseInt(data.quantity) || 1;
  const area      = w * h;
  const totalSqft = area * qty;
  const baseAmt   = totalSqft * (parseFloat(data.ratePerSqft) || 0);
  const rushAmt   = data.rushOrder    ? (parseFloat(data.rushFee)   || 0) : 0;
  const desAmt    = data.designCharge ? (parseFloat(data.designFee) || 0) : 0;
  // Use grand total from payload when multiple items exist, else calculate from single item
  const calcTotal = baseAmt + rushAmt + desAmt;
  const totalAmt  = parseFloat(data.totalAmount) > 0 ? parseFloat(data.totalAmount) : calcTotal;
  const bal       = totalAmt * 0.5;
  const itemsJson = (data.items && data.items.length > 0) ? JSON.stringify(data.items) : '[]';

  sheet.appendRow([
    quoteNum,                           // A  col 1  - Quote #
    new Date(),                         // B  col 2  - Date
    data.clientName  || '',             // C  col 3  - Client Name
    data.contact     || '',             // D  col 4  - Contact
    data.email       || '',             // E  col 5  - Email
    w,                                  // F  col 6  - Width (last item)
    h,                                  // G  col 7  - Height (last item)
    parseFloat(area.toFixed(2)),        // H  col 8  - Area/pc
    qty,                                // I  col 9  - Quantity
    parseFloat(totalSqft.toFixed(2)),   // J  col 10 - Total Sqft
    data.eyelet      || '',             // K  col 11 - Eyelet (last item)
    data.printLayout || '',             // L  col 12 - Print Layout
    data.rushOrder    ? 'Yes' : 'No',  // M  col 13 - Rush Order
    data.designCharge ? 'Yes' : 'No',  // N  col 14 - Design Charge
    parseFloat(data.ratePerSqft) || 0, // O  col 15 - Rate/sqft
    parseFloat(data.rushFee)     || 0, // P  col 16 - Rush Fee
    parseFloat(data.designFee)   || 0, // Q  col 17 - Design Fee
    parseFloat(baseAmt.toFixed(2)),    // R  col 18 - Base Amount (last item)
    parseFloat(rushAmt.toFixed(2)),    // S  col 19 - Rush Fee Amt (last item)
    parseFloat(desAmt.toFixed(2)),     // T  col 20 - Design Fee Amt (last item)
    parseFloat(totalAmt.toFixed(2)),   // U  col 21 - TOTAL AMOUNT (grand total)
    parseFloat(bal.toFixed(2)),        // V  col 22 - Balance
    data.dateNeeded  || '',            // W  col 23 - Date Needed
    data.status || 'Pending',          // X  col 24 - Status
    '',                                // Y  col 25 - Approved By
    staffName,                         // Z  col 26 - Sales Staff
    '',                                // AA col 27 - Payment Term (set by savePaymentTerm)
    itemsJson,                         // AB col 28 - Items JSON
    data.taxType  || 'non-vat',       // AC col 29 - Tax Type
    parseFloat(data.taxAmount) || 0,  // AD col 30 - Tax Amount
  ]);

  sheet.getRange(sheet.getLastRow(), 18, 1, 4).setNumberFormat('₱#,##0.00');
  return quoteNum;
}

// ══════════════════════════════════════════════════════════════════
//  TARPAULIN PRICING
// ══════════════════════════════════════════════════════════════════
function getTarpPricing() {
  const defaults = { ratePerSqft: 50, rushFee: 150, designFee: 250 };
  try {
    const ss    = SpreadsheetApp.openById('1uZQlQWBSAvee0g8gBiZytATD8T8VxN9V1DJxwGz5N7o');
    const sheet = ss.getSheetByName('Banner');
    if (!sheet) return defaults;

    // Sheet layout: col A = category ("Banner"), col B = label, col C = price
    const rows = sheet.getDataRange().getValues();
    const result = Object.assign({}, defaults);

    for (let i = 0; i < rows.length; i++) {
      const key = String(rows[i][1] || '').trim().toLowerCase();
      const raw = rows[i][2];
      const val = typeof raw === 'number'
        ? raw
        : parseFloat(String(raw || '').replace(/[^\d.]/g, '')) || 0;
      if (!key || val <= 0) continue;
      if (key.includes('rush'))   result.rushFee     = val;
      if (key.includes('design')) result.designFee   = val;
      if (key.includes('rate') || key.includes('sqft') || key.includes('sq ft') || key.includes('per sq') || key.includes('price per')) result.ratePerSqft = val;
    }

    return result;
  } catch(e) {
    return defaults;
  }
}

// ══════════════════════════════════════════════════════════════════
//  FRAME PRICING  (live from external spreadsheet, Frame sheet)
// ══════════════════════════════════════════════════════════════════
function getFramePricing() {
  const defaults = { withMatting: 600, withoutMatting: 550, rushFee: 150 };
  try {
    const ss    = SpreadsheetApp.openById('1uZQlQWBSAvee0g8gBiZytATD8T8VxN9V1DJxwGz5N7o');
    const sheet = ss.getSheetByName('Frame');
    if (!sheet) return defaults;

    // Sheet layout: col A = category ("Frame"), col B = label, col C = price
    const rows = sheet.getDataRange().getValues();
    const result = Object.assign({}, defaults);

    for (let i = 0; i < rows.length; i++) {
      const key = String(rows[i][1] || '').trim().toLowerCase();
      const raw = rows[i][2];
      const val = typeof raw === 'number'
        ? raw
        : parseFloat(String(raw || '').replace(/[^\d.]/g, '')) || 0;
      if (!key || val <= 0) continue;
      // Order matters: check "without" and "rush" before generic "with" / "matting"
      if (key.includes('rush')) {
        result.rushFee = val;
      } else if (key.includes('without') || key.includes('no matting') || key.includes('w/o')) {
        result.withoutMatting = val;
      } else if (key.includes('with') || key.includes('matting')) {
        result.withMatting = val;
      }
    }

    return result;
  } catch(e) {
    return defaults;
  }
}

// ══════════════════════════════════════════════════════════════════
//  SAVE FRAME ORDER
// ══════════════════════════════════════════════════════════════════
function saveFrameOrder(data) {
  const ss  = getMainSS_();
  let sheet = ss.getSheetByName(FRAME_SHEET);

  if (!sheet) sheet = ss.insertSheet(FRAME_SHEET);

  const headers = [
    'Quote #', 'Date', 'Client Name', 'Contact', 'Email',
    'Date Needed',
    'Width (ft)', 'Height (ft)', 'Quantity', 'Area/pc (sqft)', 'Total Sqft',
    'Matting', 'Rate/sqft',
    'Base Amount', 'Total Amount', 'Downpayment', 'Balance',
    'Special Instructions', 'Sales Staff',
    'Status', 'Approved By', 'Payment Term', 'Tax Type', 'Tax Amount',
    'Rush Order', 'Rush Fee',
  ];

  // Auto-insert header row if missing
  const firstCell = sheet.getLastRow() > 0 ? String(sheet.getRange(1,1).getValue()) : '';
  if (sheet.getLastRow() === 0 || firstCell.startsWith('FQ-')) {
    if (firstCell.startsWith('FQ-')) sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#E8151B').setFontColor('#fff')
      .setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
  }

  const lastRow  = sheet.getLastRow();
  const quoteNum = 'FQ-' + String(lastRow).padStart(4, '0');

  const session2  = data.token ? getSessionData_(data.token) : null;
  const staffName = session2 ? (session2.username || session2.name) : (data.salesStaff || '');

  const w        = parseFloat(data.width)  || 0;
  const h        = parseFloat(data.height) || 0;
  const qty      = parseInt(data.quantity) || 1;
  const area     = w * h;
  const totalSq  = area * qty;
  const rate     = parseFloat(data.ratePerSqft) || 0;
  const baseAmt  = totalSq * rate;
  const rushFee  = parseFloat(data.rushFee) || 0;
  const totalAmt = parseFloat(data.totalAmount) > 0 ? parseFloat(data.totalAmount) : (baseAmt + rushFee);
  const dp       = totalAmt * 0.5;
  const bal      = totalAmt - dp;

  sheet.appendRow([
    quoteNum,                                  // A  col 1  - Quote #
    new Date(),                                // B  col 2  - Date
    data.clientName       || '',               // C  col 3  - Client Name
    data.contact          || '',               // D  col 4  - Contact
    data.email            || '',               // E  col 5  - Email
    data.dateNeeded       || '',               // F  col 6  - Date Needed
    w,                                         // G  col 7  - Width
    h,                                         // H  col 8  - Height
    qty,                                       // I  col 9  - Quantity
    parseFloat(area.toFixed(2)),               // J  col 10 - Area/pc
    parseFloat(totalSq.toFixed(2)),            // K  col 11 - Total Sqft
    data.matting          || '',               // L  col 12 - Matting (With/Without)
    rate,                                      // M  col 13 - Rate/sqft
    parseFloat(baseAmt.toFixed(2)),            // N  col 14 - Base Amount
    parseFloat(totalAmt.toFixed(2)),           // O  col 15 - Total Amount
    parseFloat(dp.toFixed(2)),                 // P  col 16 - Downpayment
    parseFloat(bal.toFixed(2)),                // Q  col 17 - Balance
    data.notes            || '',               // R  col 18 - Special Instructions
    staffName,                                 // S  col 19 - Sales Staff
    'Pending',                                 // T  col 20 - Status
    '',                                        // U  col 21 - Approved By
    '',                                        // V  col 22 - Payment Term
    data.taxType          || 'non-vat',        // W  col 23 - Tax Type
    parseFloat(data.taxAmount) || 0,           // X  col 24 - Tax Amount
    data.rushOrder        || '',               // Y  col 25 - Rush Order
    parseFloat(rushFee.toFixed(2)),            // Z  col 26 - Rush Fee
  ]);

  sheet.getRange(sheet.getLastRow(), 14, 1, 4).setNumberFormat('₱#,##0.00');
  return quoteNum;
}

// ══════════════════════════════════════════════════════════════════
//  T-SHIRT PRICING  (live from external spreadsheet, T-Shirt sheet)
// ══════════════════════════════════════════════════════════════════
function getTshirtPricing() {
  const defaults = {
    sublimation: {}, dtf: {}, fullSub: {}, shirts: {},
    rushFee: 150, designFee: 250, fullSubMinQty: 15,
  };
  try {
    const ss = SpreadsheetApp.openById('1uZQlQWBSAvee0g8gBiZytATD8T8VxN9V1DJxwGz5N7o');
    // Try multiple capitalizations since tab name may vary
    let sheet = ss.getSheetByName('T-shirt') || ss.getSheetByName('T-Shirt')
             || ss.getSheetByName('Tshirt')  || ss.getSheetByName('TShirt');
    if (!sheet) return defaults;

    const rows = sheet.getDataRange().getValues();
    const result = {
      sublimation: {}, dtf: {}, fullSub: {}, shirts: {},
      rushFee: defaults.rushFee, designFee: defaults.designFee, fullSubMinQty: defaults.fullSubMinQty,
    };

    function parsePrice(raw) {
      if (typeof raw === 'number') return raw;
      return parseFloat(String(raw || '').replace(/[^\d.]/g, '')) || 0;
    }
    // Normalize label to Title Case so "Logo only" and "Logo Only" become one canonical key
    function canon(s) {
      return String(s || '').trim().replace(/\s+/g, ' ')
        .toLowerCase()
        .replace(/\b\w/g, c => c.toUpperCase());
    }

    // Step 1: find category header columns by scanning every row for the labels
    let subCol = -1, dtfCol = -1, fullSubCol = -1, headerRowIdx = -1;
    for (let r = 0; r < rows.length && headerRowIdx === -1; r++) {
      for (let c = 0; c < rows[r].length; c++) {
        const v = String(rows[r][c] || '').trim().toLowerCase();
        if (v === 'sublimation' && subCol === -1) { subCol = c; headerRowIdx = r; }
        if (v === 'dtf' && dtfCol === -1) { dtfCol = c; headerRowIdx = r; }
        if ((v === 'full sublimation' || v === 'full sub') && fullSubCol === -1) { fullSubCol = c; headerRowIdx = r; }
      }
    }

    // Step 2: find "If with shirt, add" row FIRST so we can stop category reads before it
    let shirtSectionStart = -1;
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < rows[r].length; c++) {
        const v = String(rows[r][c] || '').trim().toLowerCase();
        if (v.indexOf('if with shirt') === 0 || v === 'if with shirt, add') {
          shirtSectionStart = r;  // the heading row itself; data starts at +1
          break;
        }
      }
      if (shirtSectionStart >= 0) break;
    }
    // Category sections are read from header+1 up to (but not including) the shirt section
    const categoryEnd = shirtSectionStart >= 0 ? shirtSectionStart : rows.length;

    // Step 3: read items below the header row, BUT only up to categoryEnd
    if (headerRowIdx >= 0) {
      for (let r = headerRowIdx + 1; r < categoryEnd; r++) {
        const row = rows[r];

        if (subCol >= 0) {
          const label = canon(row[subCol]);
          const price = parsePrice(row[subCol + 1]);
          if (label && price > 0) result.sublimation[label] = price;
        }
        if (dtfCol >= 0) {
          const label = canon(row[dtfCol]);
          const price = parsePrice(row[dtfCol + 1]);
          if (label && price > 0) result.dtf[label] = price;
        }
        if (fullSubCol >= 0) {
          const label = canon(row[fullSubCol]);
          const price = parsePrice(row[fullSubCol + 1]);
          if (label && price > 0) result.fullSub[label] = price;
        }
      }
    }

    // Step 4: read the shirt-add section (rows after "If with shirt, add")
    if (shirtSectionStart >= 0) {
      for (let r = shirtSectionStart + 1; r < rows.length; r++) {
        const label = canon(rows[r][0]);
        const price = parsePrice(rows[r][1]);
        if (label && price > 0) result.shirts[label] = price;
      }
    }

    // Step 5: scan everywhere for optional Rush / Design / Min Qty rows
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < rows[r].length - 1; c++) {
        const v = String(rows[r][c] || '').trim().toLowerCase();
        const next = parsePrice(rows[r][c + 1]);
        if (!v || next <= 0) continue;
        if (v === 'rush' || v.indexOf('rush ') === 0) result.rushFee = next;
        else if (v.indexOf('design') === 0 || v.indexOf('artwork') === 0) result.designFee = next;
        else if (v.indexOf('min') === 0 && (v.indexOf('qty') >= 0 || v.indexOf('order') >= 0)) result.fullSubMinQty = next;
      }
    }

    return result;
  } catch(e) {
    return defaults;
  }
}

// ══════════════════════════════════════════════════════════════════
//  SAVE T-SHIRT ORDER
// ══════════════════════════════════════════════════════════════════
function saveTshirtOrder(data) {
  const ss  = getMainSS_();
  let sheet = ss.getSheetByName(TSHIRT_SHEET);

  if (!sheet) sheet = ss.insertSheet(TSHIRT_SHEET);

  const headers = [
    'Quote #', 'Date', 'Client Name', 'Contact', 'Email',
    'Date Needed',
    'Print Type', 'Logo Size', 'Has Own Shirt', 'Shirt Choice', 'Garment Type',
    'Shirt Color', 'Size Breakdown', 'Quantity',
    'Print Price/Unit', 'Shirt Price/Unit', 'Unit Price', 'Base Amount',
    'Rush Order', 'Rush Fee', 'Design Service', 'Design Fee',
    'Total Amount', 'Downpayment', 'Balance',
    'Special Instructions', 'Sales Staff',
    'Status', 'Approved By', 'Payment Term', 'Tax Type', 'Tax Amount',
  ];

  // Auto-insert header row if missing
  const firstCell = sheet.getLastRow() > 0 ? String(sheet.getRange(1,1).getValue()) : '';
  if (sheet.getLastRow() === 0 || firstCell.startsWith('SH-')) {
    if (firstCell.startsWith('SH-')) sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#E8151B').setFontColor('#fff')
      .setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
  }

  const lastRow  = sheet.getLastRow();
  const quoteNum = 'SH-' + String(lastRow).padStart(4, '0');

  const session2  = data.token ? getSessionData_(data.token) : null;
  const staffName = session2 ? (session2.username || session2.name) : (data.salesStaff || '');

  const qty        = parseInt(data.quantity) || 1;
  const printPrice = parseFloat(data.printPrice) || 0;
  const shirtPrice = parseFloat(data.shirtPrice) || 0;
  const unitPrice  = parseFloat(data.unitPrice)  || (printPrice + shirtPrice);
  const baseAmt    = unitPrice * qty;
  const rushFee    = parseFloat(data.rushFee)   || 0;
  const designFee  = parseFloat(data.designFee) || 0;
  const totalAmt   = parseFloat(data.totalAmount) > 0 ? parseFloat(data.totalAmount) : (baseAmt + rushFee + designFee);
  const dp         = totalAmt * 0.5;
  const bal        = totalAmt - dp;

  sheet.appendRow([
    quoteNum,                                  // A  col 1  - Quote #
    new Date(),                                // B  col 2  - Date
    data.clientName       || '',               // C  col 3  - Client Name
    data.contact          || '',               // D  col 4  - Contact
    data.email            || '',               // E  col 5  - Email
    data.dateNeeded       || '',               // F  col 6  - Date Needed
    data.printType        || '',               // G  col 7  - Print Type
    data.logoSize         || '',               // H  col 8  - Logo Size
    data.hasOwnShirt      || '',               // I  col 9  - Has Own Shirt
    data.shirtChoice      || '',               // J  col 10 - Shirt Choice
    data.garmentType      || '',               // K  col 11 - Garment Type (Full Sub)
    data.shirtColor       || '',               // L  col 12 - Shirt Color
    data.sizeBreakdown    || '',               // M  col 13 - Size Breakdown
    qty,                                       // N  col 14 - Quantity
    parseFloat(printPrice.toFixed(2)),         // O  col 15 - Print Price/Unit
    parseFloat(shirtPrice.toFixed(2)),         // P  col 16 - Shirt Price/Unit
    parseFloat(unitPrice.toFixed(2)),          // Q  col 17 - Unit Price
    parseFloat(baseAmt.toFixed(2)),            // R  col 18 - Base Amount
    data.rushOrder        || '',               // S  col 19 - Rush Order
    parseFloat(rushFee.toFixed(2)),            // T  col 20 - Rush Fee
    data.designService    || '',               // U  col 21 - Design Service
    parseFloat(designFee.toFixed(2)),          // V  col 22 - Design Fee
    parseFloat(totalAmt.toFixed(2)),           // W  col 23 - Total Amount
    parseFloat(dp.toFixed(2)),                 // X  col 24 - Downpayment
    parseFloat(bal.toFixed(2)),                // Y  col 25 - Balance
    data.notes            || '',               // Z  col 26 - Special Instructions
    staffName,                                 // AA col 27 - Sales Staff
    data.status || 'Pending',                 // AB col 28 - Status
    '',                                        // AC col 29 - Approved By
    '',                                        // AD col 30 - Payment Term
    data.taxType          || 'non-vat',        // AE col 31 - Tax Type
    parseFloat(data.taxAmount) || 0,           // AF col 32 - Tax Amount
  ]);

  sheet.getRange(sheet.getLastRow(), 15, 1, 11).setNumberFormat('₱#,##0.00');
  return quoteNum;
}

// ══════════════════════════════════════════════════════════════════
//  PUBLIC CUSTOMER FUNCTIONS
// ══════════════════════════════════════════════════════════════════
function getPublicPricing() {
  return {
    mug:      getMugPricing(),
    tshirt:   getTshirtPricing(),
    tarp:     getTarpPricing(),
    frame:    getFramePricing(),
    bookbind: getBookbindPricing(),
    receipt:  (function(){ try{ return getReceiptPricing(); }catch(e){ return null; } })(),
    signage:  (function(){ try{ return getPricing(); }      catch(e){ return null; } })(),
  };
}

function submitCustomerRequest(data) {
  try {
    const ss    = getCustomerSS_();
    let   sheet = ss.getSheetByName(CUSTOMER_SHEET);
    if (!sheet) sheet = ss.insertSheet(CUSTOMER_SHEET);

    const headers = [
      'Request #', 'Date Submitted', 'Client Name', 'Contact', 'Email',
      'Product Type', 'Specs Summary', 'Quantity',
      'Rush Order', 'Design Service',
      'Total Amount', 'Downpayment (50%)', 'Balance',
      'Date Needed', 'Notes', 'Status', 'Assigned To',
    ];

    const firstCell = sheet.getLastRow() > 0 ? String(sheet.getRange(1,1).getValue()) : '';
    if (sheet.getLastRow() === 0 || firstCell.startsWith('CUST-')) {
      if (firstCell.startsWith('CUST-')) sheet.insertRowBefore(1);
      sheet.getRange(1,1,1,headers.length).setValues([headers])
        .setBackground('#E8151B').setFontColor('#fff')
        .setFontWeight('bold').setFontSize(11);
      sheet.setFrozenRows(1);
    }

    const lastRow  = sheet.getLastRow();
    const reqNum   = 'CUST-' + String(lastRow).padStart(4, '0');
    const type     = String(data.productType || '').toUpperCase();
    const total    = parseFloat(data.totalAmount) || 0;
    const dp       = total * 0.5;
    const bal      = total - dp;

    // Page-size labels for receipt size divisions
    const RECEIPT_SIZE_LABELS = {
      '1': 'Full Page', '2': 'Half Page', '3': 'Third Page',
      '4': 'Quarter Page', '6': 'Sixth Page', '8': 'Eighth Page',
    };

    // Build a human-readable specs summary
    let specs = '';
    if (data.productType === 'mug') {
      specs = (data.mugType || '') + ' × ' + (data.quantity || 0) + ' pcs';
    } else if (data.productType === 'tshirt') {
      specs = (data.printType || '') + (data.logoSize ? ' ' + data.logoSize : '') + ' × ' + (data.quantity || 0) + ' pcs';
      if (data.garmentType) specs += ' | ' + data.garmentType;
      if (data.hasOwnShirt === 'Yes') specs += ' | Own shirt';
      else if (data.shirtChoice)      specs += ' | Shirt: ' + data.shirtChoice;
      if (data.shirtColor) specs += ' | Color: ' + data.shirtColor;
    } else if (data.productType === 'tarp') {
      specs = (data.width || '?') + ' × ' + (data.height || '?') + ' ft × ' + (data.quantity || 1) + ' pc(s)';
    } else if (data.productType === 'frame') {
      specs = (data.frameWidth || data.width || '?') + ' × ' + (data.frameHeight || data.height || '?') + ' ' + (data.frameUnit || data.unit || 'in') + ' | '
            + ((data.hasMatting === 'With Matting' || data.hasMatting === 'Yes') ? 'With matting' : 'No matting')
            + ' × ' + (data.quantity || 1) + ' pc(s)';
    } else if (data.productType === 'bookbind') {
      specs = (data.bindingType || data.bindType || '—') + ' × ' + (data.quantity || 0) + ' volume(s)';
      if (data.pages)       specs += ' | ' + data.pages + ' pages';
      if (data.paperSize)   specs += ' | ' + data.paperSize + (data.orientation ? ' ' + data.orientation : '');
      if (data.bindingSide) specs += ' | ' + data.bindingSide + ' bind';
      if (data.coverColor)  specs += ' | Cover: ' + data.coverColor;
      if (data.textColor)   specs += ' | Text: ' + data.textColor;
      if (data.fontStyle)   specs += ' | Font: ' + data.fontStyle;
      if (data.printingType && data.printedMaterialsReady === 'No') specs += ' | Print: ' + data.printingType;
    } else if (data.productType === 'receipt') {
      const sizeLabel = RECEIPT_SIZE_LABELS[String(data.sizeDiv)] || data.sizeDiv || '—';
      specs = (data.paperType || '—') + ' · ' + (data.copies || '—') + ' · '
            + sizeLabel + ' × ' + (data.quantity || 0) + ' booklet(s)';
      if (data.colors)      specs += ' | Colors: ' + data.colors;
      if (data.perforation) specs += ' | Perf: ' + data.perforation;
      if (data.numbering)   specs += ' | Numbering: ' + data.numbering;
      if (data.numbering === 'Yes' && data.startingNo) specs += ' (from ' + data.startingNo + ')';
    } else if (data.productType === 'signage') {
      specs = (data.signageType || '—') + ' · ' + (data.width || '?') + ' × ' + (data.height || '?') + ' ' + (data.unit || 'ft') + ' × ' + (data.quantity || 1) + ' pc(s)';
      if (data.lighting)  specs += ' | ' + data.lighting;
      if (data.material)  specs += ' | ' + data.material;
      if (data.mounting)  specs += ' | Mount: ' + data.mounting;
      if (data.transport) specs += ' | Transport: ' + data.transport + (data.transportLocation ? ' (' + data.transportLocation + ')' : '');
      if (data.electrical) specs += ' | Electrical: ' + data.electrical;
      if (data.designService === 'Yes') specs += ' | Design';
    }

    sheet.appendRow([
      reqNum,                          // A - Request #
      new Date(),                      // B - Date Submitted
      data.clientName    || '',        // C - Client Name
      data.contact       || '',        // D - Contact
      data.email         || '',        // E - Email
      type,                            // F - Product Type
      specs,                           // G - Specs Summary
      data.quantity      || '',        // H - Quantity
      data.rushOrder     || 'No',      // I - Rush Order
      data.designService || data.designCharge || 'No', // J - Design
      parseFloat(total.toFixed(2)),    // K - Total Amount
      parseFloat(dp.toFixed(2)),       // L - Downpayment
      parseFloat(bal.toFixed(2)),      // M - Balance
      data.dateNeeded    || '',        // N - Date Needed
      data.notes         || '',        // O - Notes
      'Quote Request',                 // P - Status
      '',                              // Q - Assigned To
    ]);

    sheet.getRange(sheet.getLastRow(), 11, 1, 3).setNumberFormat('₱#,##0.00');
    return reqNum;
  } catch(e) {
    throw new Error('Submit failed: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════════
//  CUSTOMER DASHBOARD DATA
// ══════════════════════════════════════════════════════════════════
function getCustomerDashboardData(token) {
  try {
    if (!getSessionData_(token)) return { success: false, message: 'Not authorized.' };
    const ss    = getCustomerSS_();
    const sheet = ss.getSheetByName(CUSTOMER_SHEET);
    if (!sheet) return { success: true, rows: [] };
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, rows: [] };
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (!r[0]) continue;
      rows.push({
        rowNum:       i + 1,
        reqNum:       String(r[0]  || ''),
        dateSubmitted:r[1] ? new Date(r[1]).toLocaleDateString('en-PH',{year:'numeric',month:'short',day:'numeric'}) : '',
        clientName:   String(r[2]  || ''),
        contact:      String(r[3]  || ''),
        email:        String(r[4]  || ''),
        productType:  String(r[5]  || ''),
        specs:        String(r[6]  || ''),
        quantity:     r[7] || '',
        rushOrder:    String(r[8]  || ''),
        designService:String(r[9]  || ''),
        totalAmount:  parseFloat(r[10]) || 0,
        downpayment:  parseFloat(r[11]) || 0,
        balance:      parseFloat(r[12]) || 0,
        dateNeeded:   r[13] ? new Date(r[13]).toLocaleDateString('en-PH',{year:'numeric',month:'short',day:'numeric'}) : '',
        notes:        String(r[14] || ''),
        status:       String(r[15] || 'Quote Request'),
        assignedTo:   String(r[16] || ''),
      });
    }
    rows.reverse();
    return { success: true, rows: rows };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updateCustomerQuoteStatus(token, rowNum, status) {
  try {
    if (!getSessionData_(token)) return { success: false, message: 'Not authorized.' };
    const ss    = getCustomerSS_();
    const sheet = ss.getSheetByName(CUSTOMER_SHEET);
    if (!sheet) return { success: false, message: 'Sheet not found.' };
    sheet.getRange(rowNum, 16).setValue(status);
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ══════════════════════════════════════════════════════════════════
//  GET MUG PRICING
// ══════════════════════════════════════════════════════════════════
function getMugPricing() {
  const defaults = {
    mugs: {
      'White Mug':                      [{minQty:5,price:180},{minQty:10,price:150},{minQty:25,price:125},{minQty:50,price:100}],
      'Inner Color Mug':                [{minQty:5,price:190},{minQty:10,price:160},{minQty:25,price:140},{minQty:50,price:130}],
      'Neon Color Mug':                 [{minQty:5,price:210},{minQty:10,price:195},{minQty:25,price:185},{minQty:50,price:175}],
      'Enamel Mug':                     [{minQty:5,price:280},{minQty:10,price:260},{minQty:25,price:250},{minQty:50,price:240}],
      'Gold/Silver Glitter Mug':        [{minQty:5,price:240},{minQty:10,price:230},{minQty:25,price:225},{minQty:50,price:220}],
      'Clear/Frosted Mug':              [{minQty:5,price:220},{minQty:10,price:210},{minQty:25,price:200},{minQty:50,price:195}],
      'Heart Handle Glitter Magic Mug': [{minQty:5,price:240},{minQty:10,price:230},{minQty:25,price:225},{minQty:50,price:220}],
      'Clear/Frosted Beer Mug':         [{minQty:5,price:390},{minQty:10,price:375},{minQty:25,price:360},{minQty:50,price:350}],
    },
    rushFee: 150, designFee: 250,
  };
  try {
    const ss    = SpreadsheetApp.openById('1uZQlQWBSAvee0g8gBiZytATD8T8VxN9V1DJxwGz5N7o');
    const sheet = ss.getSheetByName('Mugs') || ss.getSheetByName('Mug') || ss.getSheetByName('Table2');
    if (!sheet) return defaults;

    const rows = sheet.getDataRange().getValues();

    // Find header row: first row where multiple cells contain "mug"
    let headerIdx = -1;
    for (let i = 0; i < rows.length; i++) {
      const mugCols = rows[i].filter(c => /mug/i.test(String(c))).length;
      if (mugCols >= 2) { headerIdx = i; break; }
    }
    if (headerIdx < 0) return defaults;

    const headers  = rows[headerIdx];
    const tierRows = rows.slice(headerIdx + 1).filter(r => r.some(c => /\d+\s*pcs?/i.test(String(c))));
    if (!tierRows.length) return defaults;

    function parseTierCell(cell) {
      const s = String(cell || '').trim();
      const qtyM   = s.match(/(\d+)\s*pcs?/i);
      const priceM = s.match(/[P₱]\s*(\d+)/) || s.match(/\.?\s*(\d{2,})\s*$/);
      if (!qtyM || !priceM) return null;
      return { minQty: parseInt(qtyM[1]), price: parseInt(priceM[1]) };
    }

    const result = { mugs: {}, rushFee: defaults.rushFee, designFee: defaults.designFee };
    for (let col = 0; col < headers.length; col++) {
      const mugName = String(headers[col]).trim();
      if (!mugName || !/mug/i.test(mugName)) continue;
      const tiers = tierRows.map(r => parseTierCell(r[col])).filter(Boolean);
      tiers.sort((a, b) => a.minQty - b.minQty);
      if (tiers.length) result.mugs[mugName] = tiers;
    }

    return Object.keys(result.mugs).length ? result : defaults;
  } catch(e) {
    Logger.log('getMugPricing error: ' + e);
    return defaults;
  }
}

// ══════════════════════════════════════════════════════════════════
//  SAVE MUG ORDER
// ══════════════════════════════════════════════════════════════════
function saveMugOrder(data) {
  const ss    = getMainSS_();
  let   sheet = ss.getSheetByName(MUG_SHEET);
  if (!sheet) sheet = ss.insertSheet(MUG_SHEET);

  const headers = [
    'Quote #', 'Date', 'Client Name', 'Contact', 'Email', 'Date Needed',
    'Mug Type', 'Quantity',
    'Base Price/Unit', 'Discount %', 'Unit Price', 'Base Amount',
    'Rush Order', 'Rush Fee', 'Design Service', 'Design Fee',
    'Total Amount', 'Downpayment', 'Balance',
    'Special Instructions', 'Sales Staff',
    'Status', 'Approved By', 'Payment Term', 'Tax Type', 'Tax Amount',
  ];

  const firstCell = sheet.getLastRow() > 0 ? String(sheet.getRange(1, 1).getValue()) : '';
  if (sheet.getLastRow() === 0 || firstCell.startsWith('MUG-')) {
    if (firstCell.startsWith('MUG-')) sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#E8151B').setFontColor('#fff')
      .setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
  }

  const lastRow  = sheet.getLastRow();
  const quoteNum = 'MUG-' + String(lastRow).padStart(4, '0');

  const session   = data.token ? getSessionData_(data.token) : null;
  const staffName = session ? (session.username || session.name) : (data.salesStaff || '');

  const qty       = parseInt(data.quantity)   || 1;
  const baseP     = parseFloat(data.basePrice) || 0;
  const discPct   = parseFloat(data.discountPct) || 0;
  const unitP     = parseFloat(data.unitPrice) || baseP * (1 - discPct / 100);
  const baseAmt   = unitP * qty;
  const rushFee   = parseFloat(data.rushFee)   || 0;
  const designFee = parseFloat(data.designFee) || 0;
  const totalAmt  = parseFloat(data.totalAmount) > 0 ? parseFloat(data.totalAmount) : (baseAmt + rushFee + designFee);
  const dp        = totalAmt * 0.5;
  const bal       = totalAmt - dp;

  sheet.appendRow([
    quoteNum,                               // A  - Quote #
    new Date(),                             // B  - Date
    data.clientName    || '',               // C  - Client Name
    data.contact       || '',               // D  - Contact
    data.email         || '',               // E  - Email
    data.dateNeeded    || '',               // F  - Date Needed
    data.mugType       || '',               // G  - Mug Type
    qty,                                    // H  - Quantity
    parseFloat(baseP.toFixed(2)),           // I  - Base Price/Unit
    discPct,                                // J  - Discount %
    parseFloat(unitP.toFixed(2)),           // K  - Unit Price
    parseFloat(baseAmt.toFixed(2)),         // L  - Base Amount
    data.rushOrder     || '',               // M  - Rush Order
    parseFloat(rushFee.toFixed(2)),         // N  - Rush Fee
    data.designService || '',               // O  - Design Service
    parseFloat(designFee.toFixed(2)),       // P  - Design Fee
    parseFloat(totalAmt.toFixed(2)),        // Q  - Total Amount
    parseFloat(dp.toFixed(2)),              // R  - Downpayment
    parseFloat(bal.toFixed(2)),             // S  - Balance
    data.notes         || '',               // T  - Special Instructions
    staffName,                              // U  - Sales Staff
    data.status || 'Pending',              // V  - Status
    '',                                     // W  - Approved By
    '',                                     // X  - Payment Term
    data.taxType       || 'non-vat',        // Y  - Tax Type
    parseFloat(data.taxAmount) || 0,        // Z  - Tax Amount
  ]);

  sheet.getRange(sheet.getLastRow(), 9, 1, 11).setNumberFormat('₱#,##0.00');
  return quoteNum;
}

// ══════════════════════════════════════════════════════════════════
//  FIX TARP HEADERS (utility/one-time runner)
// ══════════════════════════════════════════════════════════════════
function fixTarpHeaders() {
  const ss    = getMainSS_();
  const sheet = ss.getSheetByName(TARP_SHEET);
  if (!sheet) { Logger.log('Sheet not found'); return; }

  const firstCell = String(sheet.getRange(1,1).getValue()).trim();
  if (firstCell === 'Quote #') { Logger.log('Headers already exist!'); return; }

  sheet.insertRowBefore(1);
  const headers = [
    'Quote #',        // A  col 1
    'Date',           // B  col 2
    'Client Name',    // C  col 3
    'Contact',        // D  col 4
    'Email',          // E  col 5
    'Width (ft)',     // F  col 6
    'Height (ft)',    // G  col 7
    'Area/pc (sqft)', // H  col 8
    'Quantity',       // I  col 9
    'Total Sqft',     // J  col 10
    'Eyelet',         // K  col 11
    'Print Layout',   // L  col 12
    'Rush Order',     // M  col 13
    'Design Charge',  // N  col 14
    'Rate/sqft',      // O  col 15
    'Rush Fee',       // P  col 16
    'Design Fee',     // Q  col 17
    'Base Amount',    // R  col 18
    'Rush Fee Amt',   // S  col 19
    'Design Fee Amt', // T  col 20
    'TOTAL AMOUNT',   // U  col 21
    'Balance',        // V  col 22
    'Date Needed',    // W  col 23
    'Status',         // X  col 24
    'Approved By',    // Y  col 25
    'Sales Staff',    // Z  col 26
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#E8151B')
    .setFontColor('#fff')
    .setFontWeight('bold')
    .setFontSize(11);
  sheet.setFrozenRows(1);
  Logger.log('Headers fixed!');
}

// ══════════════════════════════════════════════════════════════════
//  HELPERS
// ══════════════════════════════════════════════════════════════════
function getSheetUrl() {
  return getMainSS_().getUrl();
}

// ══════════════════════════════════════════════════════════════════
//  GET TARP DASHBOARD DATA
// ══════════════════════════════════════════════════════════════════
function getTarpDashboardData(token) {
  try {
    if (!token) return null;
    const session = getSessionData_(token);
    if (!session) return null;

    const ss    = getMainSS_();
    const sheet = ss.getSheetByName(TARP_SHEET);
    if (!sheet) return { name: session.name, username: session.username, role: session.role, quotes: [] };

    const data  = sheet.getDataRange().getValues();
    const role  = session.role.toLowerCase();

    const quotes = data.slice(1).filter(r => r[0]).map(row => {
      let dateStr = '';
      try { dateStr = row[1] ? new Date(row[1]).toISOString() : ''; } catch(e) { dateStr = ''; }
      return {
        quoteNum:     String(row[0]  || ''),
        date:         dateStr,
        clientName:   String(row[2]  || ''),
        contact:      String(row[3]  || ''),
        email:        String(row[4]  || ''),
        width:        row[5]  || 0,
        height:       row[6]  || 0,
        sqft:         row[7]  || 0,
        quantity:     row[8]  || 1,
        totalSqft:    row[9]  || 0,
        eyelet:       String(row[10] || ''),
        printLayout:  String(row[11] || ''),
        rushOrder:    String(row[12] || ''),
        designCharge: String(row[13] || ''),
        ratePerSqft:  row[14] || 0,
        rushFee:      row[15] || 0,
        designFee:    row[16] || 0,
        baseAmount:   row[17] || 0,
        rushFeeAmt:   row[18] || 0,
        designFeeAmt: row[19] || 0,
        totalAmount:  row[20] || 0,
        balance:      row[21] || 0,
        dateNeeded:   String(row[22] || ''),
        status:       String(row[23] || 'Pending'),  // X = col 24
        approvedBy:   String(row[24] || ''),          // Y = col 25
        salesStaff:   String(row[25] || ''),          // Z = col 26
      };
    });

    const filtered = (role === 'sales' || role === 'staff')
      ? quotes.filter(q => q.salesStaff === session.username || q.salesStaff === session.name)
      : quotes;

    return { name: session.name, username: session.username, role: session.role, quotes: filtered };
  } catch(err) {
    Logger.log('getTarpDashboardData error: ' + err.message);
    return null;
  }
}

// ══════════════════════════════════════════════════════════════════
//  UPDATE TARP QUOTE STATUS  (kept for backward compat)
// ══════════════════════════════════════════════════════════════════
function updateTarpQuoteStatus(token, quoteNum, status) {
  return updateQuoteStatus(token, quoteNum, status);
}

// ══════════════════════════════════════════════════════════════════
//  SAVE SIGNAGE QUOTATION
// ══════════════════════════════════════════════════════════════════
function saveQuotation(data) {
  const ss  = getMainSS_();
  let sheet = ss.getSheetByName(SHEET_QUOTATIONS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_QUOTATIONS);
    const headers = [
      'Quote #','Date','Client Name','Contact','Email','Address','Delivery',
      'Signage Type','Lighting','Material','Width (ft)','Height (ft)',
      'Sq Ft','Rate/sqft','Min Charge','Min Area','Mounting',
      'Install Notes','Unit','Actual Sqft','Type Notes',
      'Base Amount','Mount Fee','Complexity Surcharge','Design Fee','Total Amount',
      'Downpayment','Balance','Status','Approved By','Sales Staff','Date Needed',
      'Addon Design','Addon Design Fee','Addon Rush','Addon Rush Fee',
      'Addon Elec','Addon Elec Fee','Addon Transport','Addon Transport Fee','Addon Transport Location',
      'Payment Term','Tax Type','Tax Amount',
    ];
    sheet.appendRow(headers);
    sheet.getRange(1,1,1,headers.length)
      .setBackground('#E8151B').setFontColor('#fff')
      .setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
  }

  const lastRow  = sheet.getLastRow();
  const quoteNum = `SQ-${String(lastRow).padStart(4,'0')}`;

  sheet.appendRow([
    quoteNum,                                    // A  col 1  - Quote #
    new Date(),                                  // B  col 2  - Date
    data.clientName       || '',                 // C  col 3  - Client Name
    data.contact          || '',                 // D  col 4  - Contact
    data.email            || '',                 // E  col 5  - Email
    data.address          || '',                 // F  col 6  - Address
    data.delivery         || '',                 // G  col 7  - Delivery
    data.signageType      || '',                 // H  col 8  - Signage Type
    data.lighting         || '',                 // I  col 9  - Lighting
    data.material         || '',                 // J  col 10 - Material
    parseFloat(data.width)         || 0,         // K  col 11 - Width (ft)
    parseFloat(data.height)        || 0,         // L  col 12 - Height (ft)
    parseFloat(data.sqft)          || 0,         // M  col 13 - Sq Ft
    parseFloat(data.rate)          || 0,         // N  col 14 - Rate/sqft
    parseFloat(data.minCharge)     || 0,         // O  col 15 - Min Charge
    parseFloat(data.minArea)       || 0,         // P  col 16 - Min Area
    data.mounting         || '',                 // Q  col 17 - Mounting
    data.installNotes     || '',                 // R  col 18 - Install Notes
    data.unit             || 'ft',               // S  col 19 - Unit
    parseFloat(data.actualSqft)    || 0,         // T  col 20 - Actual Sqft
    data.typeNotes        || '',                 // U  col 21 - Type Notes
    parseFloat(data.baseAmount)    || 0,         // V  col 22 - Base Amount
    parseFloat(data.mountFee)      || 0,         // W  col 23 - Mount Fee
    parseFloat(data.complexitySurcharge) || 0,   // X  col 24 - Complexity Surcharge
    parseFloat(data.designFee)     || 0,         // Y  col 25 - Design Fee
    parseFloat(data.totalAmount)   || 0,         // Z  col 26 - Total Amount
    parseFloat(data.downpayment)   || 0,         // AA col 27 - Downpayment
    parseFloat(data.balance)       || 0,         // AB col 28 - Balance
    'Pending',                                   // AC col 29 - Status
    '',                                          // AD col 30 - Approved By
    data.salesStaff       || '',                 // AE col 31 - Sales Staff
    data.dateNeeded       || '',                 // AF col 32 - Date Needed
    data.addonDesign      || '',                 // AG col 33 - Addon Design
    parseFloat(data.addonDesignFee)    || 0,     // AH col 34 - Addon Design Fee
    data.addonRush        || '',                 // AI col 35 - Addon Rush
    parseFloat(data.addonRushFee)      || 0,     // AJ col 36 - Addon Rush Fee
    data.addonElec        || '',                 // AK col 37 - Addon Elec
    parseFloat(data.addonElecFee)      || 0,     // AL col 38 - Addon Elec Fee
    data.addonTransport   || '',                 // AM col 39 - Addon Transport
    parseFloat(data.addonTransportFee) || 0,     // AN col 40 - Addon Transport Fee
    data.addonTransportLocation || '',           // AO col 41 - Addon Transport Location
    '',                                          // AP col 42 - Payment Term (set by savePaymentTerm)
    data.taxType  || 'non-vat',                 // AQ col 43 - Tax Type
    parseFloat(data.taxAmount) || 0,            // AR col 44 - Tax Amount
  ]);

  sheet.getRange(sheet.getLastRow(), 22, 1, 7).setNumberFormat('₱#,##0.00');
  return quoteNum;
}

// ================================================
// SAVE RECEIPT ORDER
// ══════════════════════════════════════════════════════════════════
//  SAVE RECEIPT ORDER
// ✅ FIX #5 — removed duplicate declaration; merged both versions
//             (auto-create sheet + proper orderNum generation)
// ══════════════════════════════════════════════════════════════════
function saveReceiptOrder(payload) {
  const ss  = getMainSS_();
  let sheet = ss.getSheetByName(RECEIPT_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(RECEIPT_SHEET);
    const headers = [
      'Order #','Date','Company','Last Name','First Name','Email','Mobile',
      'Customer Type','Copies','Size','Pages/Booklet','Paper Type','Colors',
      'Perforation','Numbering','Starting No','Quantity','Date Needed',
      'Total Price','Status','Sales Staff','Notes',
      'Payment Term','Tax Type','Tax Amount',
      'Rush Order','Rush Fee',
    ];
    sheet.appendRow(headers);
    sheet.getRange(1,1,1,headers.length)
      .setBackground('#E8151B').setFontColor('#fff')
      .setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
  }

   
    // ← line 1040 (closing brace of if(!sheet))
     // ← line 1041 (blank)
    const session   = payload.salesStaff ? getSessionData_(payload.salesStaff) : null;  // ← NEW
    const staffName = session ? session.name : '';  // ← NEW
    const lastRow   = sheet.getLastRow();   // ← line 1042 (unchanged)
    const orderNum  = 'RQ-' + String(lastRow).padStart(4, '0');  // ← line 1043 (unchanged)

  sheet.appendRow([
    orderNum,                // A  col 1  - Order #
    new Date(),              // B  col 2  - Date
    payload.company  || '',  // C  col 3  - Company
    payload.lastName || '',  // D  col 4  - Last Name
    payload.firstName|| '',  // E  col 5  - First Name
    payload.email    || '',  // F  col 6  - Email
    payload.mobile   || '',  // G  col 7  - Mobile
    payload.customerType||'',// H  col 8  - Customer Type
    payload.copies   || '',  // I  col 9  - Copies
    payload.size     || '',  // J  col 10 - Size
    50,                      // K  col 11 - Pages per booklet
    payload.paperType|| '',  // L  col 12 - Paper Type
    payload.paperColors||'', // M  col 13 - Colors
    payload.perforation||'', // N  col 14 - Perforation
    payload.numbering|| '',  // O  col 15 - Numbering
    payload.startingNo||'',  // P  col 16 - Starting No
    payload.quantity || '',  // Q  col 17 - Quantity
    payload.dateNeeded||'',  // R  col 18 - Date Needed
    payload.totalPrice|| 0,           // S  col 19 - Total Price
    'Pending',                        // T  col 20 - Status
    staffName,                        // U  col 21 - Sales Staff
    payload.notes    || '',           // V  col 22 - Notes
    '',                               // W  col 23 - Payment Term (set by savePaymentTerm)
    payload.taxType  || 'non-vat',   // X  col 24 - Tax Type
    parseFloat(payload.taxAmount)||0, // Y  col 25 - Tax Amount
    payload.rushOrder|| '',           // Z  col 26 - Rush Order
    parseFloat(payload.rushFee)||0,   // AA col 27 - Rush Fee
  ]);

  sheet.getRange(sheet.getLastRow(), 19, 1, 1).setNumberFormat('₱#,##0.00');
  return orderNum;
}
function getReceiptPricing() {
  try {
    const ss = getMainSS_();
    const db = ss.getSheetByName(SHEET_DATABASE);
    if (!db) return null;

    // Get external pricing spreadsheet ID (same as signage)
    const dbData = db.getDataRange().getValues();
    let externalId = null;
    for (let i = 0; i < dbData.length; i++) {
      if (String(dbData[i][0]).trim() === 'PriceDatabase') {
        externalId = String(dbData[i][1]).trim();
        break;
      }
    }
    if (!externalId) return null;

    const match = externalId.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    const ssId  = match ? match[1] : externalId;

    const extSS    = SpreadsheetApp.openById(ssId);
    const extSheet = extSS.getSheetByName('Receipt');
    if (!extSheet) return null;

    const data = extSheet.getDataRange().getValues();
    const rows = [];

    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      const printType = String(row[0]).trim();
      // Skip the Rush row — it's read by getReceiptRushFee() instead
      if (printType.toLowerCase().includes('rush')) continue;
      rows.push({
        printType:   printType,                // Carbonless / Newsprint
        copies:      row[1],                   // 2, 3, 4
        sizeDiv:     row[2],                   // 1,2,3,4,6,8
        booklets:    row[3],                   // quantity
        sellingPrice: row[14],                 // col L = Selling Price/Booklet
      });
    }
    return rows;
  } catch(e) {
    Logger.log('getReceiptPricing error: ' + e.message);
    return null;
  }
}

// ══════════════════════════════════════════════════════════════════
//  GET RECEIPT RUSH FEE  (looks for a row labeled "Rush" in Receipt sheet)
// ══════════════════════════════════════════════════════════════════
function getReceiptRushFee() {
  const defaultFee = 150;
  try {
    const extSS    = SpreadsheetApp.openById('1uZQlQWBSAvee0g8gBiZytATD8T8VxN9V1DJxwGz5N7o');
    const extSheet = extSS.getSheetByName('Receipt');
    if (!extSheet) return defaultFee;

    const data = extSheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      // Search any cell in the row for the "Rush" label
      for (let c = 0; c < row.length; c++) {
        const cell = String(row[c] || '').trim().toLowerCase();
        if (cell === 'rush' || cell.startsWith('rush ')) {
          // The price is expected in the next non-empty numeric cell
          for (let n = c + 1; n < row.length; n++) {
            const raw = row[n];
            const val = typeof raw === 'number'
              ? raw
              : parseFloat(String(raw || '').replace(/[^\d.]/g, '')) || 0;
            if (val > 0) return val;
          }
        }
      }
    }
    return defaultFee;
  } catch(e) {
    return defaultFee;
  }
}
// ══════════════════════════════════════════════════════════════════
//  FIX RECEIPT HEADERS (utility/one-time runner)
// ══════════════════════════════════════════════════════════════════
function fixReceiptHeaders() {
  const ss    = getMainSS_();
  const sheet = ss.getSheetByName(RECEIPT_SHEET);
  if (!sheet) { Logger.log('Receipt sheet not found'); return; }

  const firstCell = String(sheet.getRange(1,1).getValue()).trim();
  if (firstCell === 'Order #') { Logger.log('Headers already exist!'); return; }

  sheet.insertRowBefore(1);
  const headers = [
    'Order #',        // A  col 1
    'Date',           // B  col 2
    'Company',        // C  col 3
    'Last Name',      // D  col 4
    'First Name',     // E  col 5
    'Email',          // F  col 6
    'Mobile',         // G  col 7
    'Customer Type',  // H  col 8
    'Copies',         // I  col 9
    'Size',           // J  col 10
    'Pages/Booklet',  // K  col 11
    'Paper Type',     // L  col 12
    'Colors',         // M  col 13
    'Perforation',    // N  col 14
    'Numbering',      // O  col 15
    'Starting No',    // P  col 16
    'Quantity',       // Q  col 17
    'Date Needed',    // R  col 18
    'Total Price',    // S  col 19
    'Status',         // T  col 20
    'Sales Staff',    // U  col 21
    'Notes',          // V  col 22
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#E8151B')
    .setFontColor('#fff')
    .setFontWeight('bold')
    .setFontSize(11);
  sheet.setFrozenRows(1);
  Logger.log('Receipt headers fixed!');
}
// ── SAVE PAYMENT TERM ──
function savePaymentTerm(token, quoteNum, termLabel, termValue) {
  const user = getSessionData_(token);
  if (!user) throw new Error('Unauthorized');
  
  const ss = getMainSS_();
  
  let sheetName;
  if (quoteNum.startsWith('SQ-')) sheetName = 'Quotations';
  else if (quoteNum.startsWith('TQ-')) sheetName = 'Tarp Quotations';
  else if (quoteNum.startsWith('RQ-')) sheetName = 'Receipt Quotations';
  else if (quoteNum.startsWith('BQ-')) sheetName = BOOKBIND_SHEET;
  else if (quoteNum.startsWith('FQ-')) sheetName = FRAME_SHEET;
  else if (quoteNum.startsWith('SH-')) sheetName = TSHIRT_SHEET;
  else throw new Error('Unknown quote type');

  const sh = ss.getSheetByName(sheetName);
  const data = sh.getDataRange().getValues();
  const headers = data[0];

  // Find or create Payment Term column
  // Fixed column per sheet type (1-based)
let ptCol;
if (quoteNum.startsWith('SQ-')) ptCol = 42;      // col AP
else if (quoteNum.startsWith('TQ-')) ptCol = 27; // col AA
else if (quoteNum.startsWith('RQ-')) ptCol = 23; // col W
else if (quoteNum.startsWith('BQ-')) ptCol = 28; // col AB
else if (quoteNum.startsWith('FQ-')) ptCol = 22; // col V
else if (quoteNum.startsWith('SH-')) ptCol = 30; // col AD

// Find the row
for (let i = 1; i < data.length; i++) {
  if (String(data[i][0]) === String(quoteNum)) {
    sh.getRange(i + 1, ptCol).setValue(termLabel);
    return true;
  }
}
throw new Error('Quote not found');
}

// ══════════════════════════════════════════════════════════════════
//  GET BOOKBIND PRICING  (live from external spreadsheet)
// ══════════════════════════════════════════════════════════════════
function getBookbindPricing() {
  const defaults = {
    'Hardbound':                   430,
    'Softbound with lettering':    150,
    'Softbound without lettering': 100,
    'Ring bind':                   100,
    'Rush fee':                    150,
  };
  try {
    const ss    = SpreadsheetApp.openById('1uZQlQWBSAvee0g8gBiZytATD8T8VxN9V1DJxwGz5N7o');
    const sheet = ss.getSheetByName('BookBind');
    if (!sheet) return defaults;
    const rows = sheet.getDataRange().getValues();
    const result = JSON.parse(JSON.stringify(defaults));
    for (let i = 1; i < rows.length; i++) {
      const type = String(rows[i][0] || '').trim();
      const raw  = rows[i][1];
      const price = typeof raw === 'number'
        ? raw
        : parseFloat(String(raw || '').replace(/[^\d.]/g, '')) || 0;
      if (type && price > 0) result[type] = price;
    }
    return result;
  } catch(e) {
    return defaults;
  }
}

// ══════════════════════════════════════════════════════════════════
//  SAVE BOOKBIND ORDER
// ══════════════════════════════════════════════════════════════════
function saveBookbindOrder(data) {
  const ss  = getMainSS_();
  let sheet = ss.getSheetByName(BOOKBIND_SHEET);

  if (!sheet) sheet = ss.insertSheet(BOOKBIND_SHEET);

  const headers = [
    'Quote #', 'Date', 'Client Name', 'Contact', 'Email',
    'Target Pickup Date', 'Binding Type', 'With Lettering', 'Quantity', 'Pages',
    'Paper Size', 'Orientation', 'Binding Side', 'Cover Color',
    'Printed Materials Ready', 'Printing Needed', 'File Final for Printing', 'Rush Order',
    'Binding Price/Unit', 'Rush Fee',
    'Total Amount', 'Downpayment', 'Balance',
    'Special Instructions', 'Sales Staff',
    'Status', 'Approved By', 'Payment Term', 'Tax Type', 'Tax Amount',
    'Text Color', 'Font Style',
  ];

  // Add header row if sheet is empty or first row is data (no header)
  const firstCell = sheet.getLastRow() > 0 ? String(sheet.getRange(1,1).getValue()) : '';
  if (sheet.getLastRow() === 0 || firstCell.startsWith('BQ-')) {
    if (firstCell.startsWith('BQ-')) {
      sheet.insertRowBefore(1);  // push existing data down
    }
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#E8151B').setFontColor('#fff')
      .setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
  }

  const lastRow  = sheet.getLastRow();
  const quoteNum = 'BQ-' + String(lastRow).padStart(4, '0');

  sheet.appendRow([
    quoteNum,                                  // A  col 1  - Quote #
    new Date(),                                // B  col 2  - Date
    data.clientName           || '',           // C  col 3  - Client Name
    data.contact              || '',           // D  col 4  - Contact
    data.email                || '',           // E  col 5  - Email
    data.dateNeeded           || '',           // F  col 6  - Target Pickup Date
    data.bindType             || '',           // G  col 7  - Binding Type
    data.withLettering        || '',           // H  col 8  - With Lettering
    parseInt(data.quantity)   || 1,            // I  col 9  - Quantity
    parseInt(data.pages)      || 0,            // J  col 10 - Pages
    data.paperSize            || '',           // K  col 11 - Paper Size
    data.orientation          || '',           // L  col 12 - Orientation
    data.bindingSide          || '',           // M  col 13 - Binding Side
    data.coverColor           || '',           // N  col 14 - Cover Color
    data.printedMaterialsReady|| '',           // O  col 15 - Printed Materials Ready
    data.printingType         || '',           // P  col 16 - Printing Type
    parseFloat(data.printingFee)||0,           // Q  col 17 - Printing Fee
    data.rushOrder            || '',           // R  col 18 - Rush Order
    parseFloat(data.bindingPrice) || 0,        // S  col 19 - Binding Price/Unit
    parseFloat(data.rushFee)      || 0,        // T  col 20 - Rush Fee
    parseFloat(data.totalAmount)  || 0,        // U  col 21 - Total Amount
    parseFloat(data.downpayment)  || 0,        // V  col 22 - Downpayment
    parseFloat(data.balance)      || 0,        // W  col 23 - Balance
    data.notes                || '',           // X  col 24 - Special Instructions
    data.salesStaff           || '',           // Y  col 25 - Sales Staff
    'Pending',                                 // Z  col 26 - Status
    '',                                        // AA col 27 - Approved By
    '',                                        // AB col 28 - Payment Term
    data.taxType              || 'non-vat',    // AC col 29 - Tax Type
    parseFloat(data.taxAmount)    || 0,        // AD col 30 - Tax Amount
    data.textColor            || '',           // AE col 31 - Text Color
    data.fontStyle            || '',           // AF col 32 - Font Style
  ]);

  sheet.getRange(sheet.getLastRow(), 21, 1, 3).setNumberFormat('₱#,##0.00');
  return quoteNum;
}
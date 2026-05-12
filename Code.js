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

// ══════════════════════════════════════════════════════════════════
//  doGet
// ══════════════════════════════════════════════════════════════════
function doGet(e) {
  const token  = String(e?.parameter?.token || '').trim();
  const page   = String(e?.parameter?.page  || '').trim().toLowerCase();
  const appUrl = ScriptApp.getService().getUrl();

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
          paymentTermValue: String(row[41] || '').includes('No Down') ? 0 : String(row[41] || '').includes('25%') ? 0.25 : String(row[41] || '').includes('Full') ? 1 : 0.5,
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
  status:      String(row[19] || 'Pending'),
  approvedBy:  String(row[20] || ''),
  salesStaff:  String(row[20] || ''),
}));
    }

    // ── COMBINE & FILTER ────────────────────────────────────────
    const allQuotes = [...quotes, ...tarpQuotes, ...receiptQuotes];

    const filtered = (role === 'sales' || role === 'staff')
      ? allQuotes.filter(q => q.salesStaff === session.username || q.salesStaff === session.name)
      : allQuotes;

    return { name: session.name, username: session.username, role: session.role, quotes: filtered };

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

  const ss        = getMainSS_();
  const isTarp    = String(quoteNum).startsWith('TQ-');
  const isReceipt = String(quoteNum).startsWith('RQ-');

  if (isReceipt) {
    const sheet = ss.getSheetByName(RECEIPT_SHEET);
    if (!sheet) throw new Error('Receipt Quotations sheet not found.');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === quoteNum) {
        sheet.getRange(i+1, 20).setValue(status);        // col T (index 19) = Status
        sheet.getRange(i+1, 21).setValue(session.name + ' — ' + new Date().toLocaleString('en-PH')); // col U = Approved By
        const color = status === 'Approved' ? '#E6FFF3' : status === 'Rejected' ? '#FFF0F0' : '#FFFFFF';
        sheet.getRange(i+1, 1, 1, 22).setBackground(color);
        return { success: true };
      }
    }
    throw new Error('Receipt order not found: ' + quoteNum);
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

    if (/^SECTION\s+1/i.test(firstLine) || /SINGLE FACE/i.test(firstLine)) {
      inProducts=true; inMounting=false; inComplexity=false;
      currentSection = 'Single Face Signage (Metal Frame)'; continue;
    }
    if (/^SECTION\s+2/i.test(firstLine) || /DOUBLE FACE/i.test(firstLine)) {
      inProducts=true; inMounting=false; inComplexity=false;
      currentSection = 'Double Face Signage (Metal Frame)'; continue;
    }
    if (/^SECTION\s+3/i.test(firstLine) || /3D BUILD UP/i.test(firstLine)) {
      inProducts=true; inMounting=false; inComplexity=false;
      currentSection = '3D Build Up Signage'; continue;
    }
    if (/^SECTION\s+4/i.test(firstLine)) { inProducts=false; inMounting=false; inComplexity=false; continue; }
    if (/^INSTALLATION\s*&\s*MOUNTING/i.test(firstLine)) { inMounting=true; inComplexity=false; inProducts=false; continue; }
    if (/^COMPLEXITY\s+SURCHARGES/i.test(firstLine))     { inComplexity=true; inMounting=false; inProducts=false; continue; }
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
    'Pending',                         // X  col 24 - Status
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
  try {
    const ss      = getMainSS_();
    const dbSheet = ss.getSheetByName(SHEET_DATABASE);
    if (!dbSheet) return { ratePerSqft: 20, rushFee: 100, designFee: 250 };

    const data = dbSheet.getDataRange().getValues();
    let ratePerSqft = 20, rushFee = 100, designFee = 250;

    for (let i = 0; i < data.length; i++) {
      const key = String(data[i][0]).trim();
      const val = parseFloat(data[i][1]) || 0;
      if (key === 'TarpRate')   ratePerSqft = val;
      if (key === 'TarpRush')   rushFee     = val;
      if (key === 'TarpDesign') designFee   = val;
    }

    return { ratePerSqft, rushFee, designFee };
  } catch(e) {
    return { ratePerSqft: 20, rushFee: 100, designFee: 250 };
  }
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

function getReceiptPricing() {
  const ss = SpreadsheetApp.openById('1ClnO3Z6xGXa2V6AVijXIfqIJDlI2zLVVKRJdEeE4ebM');
  const sheet = ss.getSheetByName('Pricing'); // i-change kung lain ang sheet name
  const data = sheet.getDataRange().getValues();
  return data.slice(1).filter(r => r[0]).map(row => ({
    printType:    String(row[0] || ''),
    copies:       Number(row[1] || 0),
    sizeDiv:      Number(row[2] || 0),
    booklets:     Number(row[3] || 0),
    sellingPrice: Number(row[4] || 0),
  }));
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
      rows.push({
        printType:   String(row[0]).trim(),   // Carbonless / Newsprint
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
  else throw new Error('Unknown quote type');
  
  const sh = ss.getSheetByName(sheetName);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  
  // Find or create Payment Term column
  let ptCol = headers.indexOf('Payment Term');
  if (ptCol === -1) {
    ptCol = headers.length;
    sh.getRange(1, ptCol + 1).setValue('Payment Term');
  }
  
  // Find the row
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(quoteNum)) {
      sh.getRange(i + 1, ptCol + 1).setValue(termLabel);
      return true;
    }
  }
  throw new Error('Quote not found');
}
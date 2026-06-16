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
const STICKER_SHEET           = 'Sticker Quotations';
const RISOGRAPH_SHEET         = 'Risograph Quotations';
const TOTEBAG_SHEET           = 'Totebag Quotations';
const TICKET_SHEET            = 'Ticket Quotations';
const NEWSPRINT_SHEET         = 'Newsprint Quotations';
const SOUVENIR_SHEET          = 'Souvenir Quotations';
const KEYCHAIN_SHEET          = 'Keychain Quotations';
const ACRYLICSIGN_SHEET       = 'Acrylic Signage Quotations';
const ACRYLICPLATE_SHEET      = 'Acrylic Plate Quotations';
const CUSTOMER_SHEET          = 'Customer Quotations';
const CUSTOMER_INFO_SHEET     = 'Customer Info';
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

// Resolve the external pricing spreadsheet from the Database sheet
// ('PriceDatabase' row, accepts a raw ID or a full URL). Falls back to
// the legacy hardcoded ID so pricing keeps working if the row is missing.
function getPriceDbSS_() {
  const FALLBACK_ID = '1uZQlQWBSAvee0g8gBiZytATD8T8VxN9V1DJxwGz5N7o';
  try {
    const db = getMainSS_().getSheetByName(SHEET_DATABASE);
    if (db) {
      const dbData = db.getDataRange().getValues();
      for (let i = 0; i < dbData.length; i++) {
        if (String(dbData[i][0]).trim() === 'PriceDatabase') {
          const raw = String(dbData[i][1]).trim();
          if (raw) {
            const m = raw.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
            return SpreadsheetApp.openById(m ? m[1] : raw);
          }
          break;
        }
      }
    }
  } catch (e) {}
  return SpreadsheetApp.openById(FALLBACK_ID);
}

// Serialize quote-number generation: two simultaneous saves would otherwise
// read the same getLastRow() and produce duplicate quote numbers. The lock
// is auto-released when the script execution ends.
function lockQuoteNumbering_() {
  try { LockService.getScriptLock().waitLock(10000); } catch (e) {}
}

// ══════════════════════════════════════════════════════════════════
//  CUSTOMER INFO  (autofill returning customers + save on every quote)
// ══════════════════════════════════════════════════════════════════
const CUSTOMER_INFO_HEADERS = [
  'Name', 'Contact', 'Company', 'Email', 'Address',
  'First Name', 'Last Name', 'First Saved', 'Last Updated',
];

function normName_(s) {
  return String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' ');
}

function getCustomerInfoSheet_() {
  const ss = getMainSS_();
  let sheet = ss.getSheetByName(CUSTOMER_INFO_SHEET);
  if (!sheet) sheet = ss.insertSheet(CUSTOMER_INFO_SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, CUSTOMER_INFO_HEADERS.length).setValues([CUSTOMER_INFO_HEADERS])
      .setBackground('#E8151B').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Look up a saved customer by name (case-insensitive). Returns the most
// recent matching record, or null. Called from the quotation pages to
// autofill the contact/email/etc. of a returning customer.
function lookupCustomerByName(name) {
  const key = normName_(name);
  if (key.length < 2) return null;
  try {
    const sheet = getMainSS_().getSheetByName(CUSTOMER_INFO_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return null;
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {   // newest row wins
      if (normName_(data[i][0]) === key) {
        return {
          name:      String(data[i][0] || ''),
          contact:   String(data[i][1] || ''),
          company:   String(data[i][2] || ''),
          email:     String(data[i][3] || ''),
          address:   String(data[i][4] || ''),
          firstName: String(data[i][5] || ''),
          lastName:  String(data[i][6] || ''),
        };
      }
    }
  } catch (e) {
    Logger.log('lookupCustomerByName error: ' + (e && e.message));
  }
  return null;
}

// Live autocomplete for the staff quotation pages. Returns up to 8 saved
// customers whose full name OR any name-word starts with the typed prefix
// (case-insensitive), newest first, de-duplicated by name.
function searchCustomersByName(prefix) {
  const key = normName_(prefix);
  if (key.length < 2) return [];
  try {
    const sheet = getMainSS_().getSheetByName(CUSTOMER_INFO_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getDataRange().getValues();
    const seen = {};
    const out  = [];
    for (let i = data.length - 1; i >= 1; i--) {   // newest first
      const full = normName_(data[i][0]);
      if (!full || seen[full]) continue;
      const startsWord = full.indexOf(key) === 0 || full.split(' ').some(function (w) { return w.indexOf(key) === 0; });
      if (!startsWord) continue;
      seen[full] = true;
      out.push({
        name:      String(data[i][0] || ''),
        contact:   String(data[i][1] || ''),
        company:   String(data[i][2] || ''),
        email:     String(data[i][3] || ''),
        firstName: String(data[i][5] || ''),
        lastName:  String(data[i][6] || ''),
      });
      if (out.length >= 8) break;
    }
    return out;
  } catch (e) {
    Logger.log('searchCustomersByName error: ' + (e && e.message));
    return [];
  }
}

// Insert or update a customer (matched by name). New non-empty fields
// overwrite blanks but never wipe existing data. Best-effort: never throws.
function upsertCustomerInfo_(rec) {
  try {
    const name = String(rec && rec.name || '').trim();
    if (!name) return;
    const sheet = getCustomerInfoSheet_();
    const key   = normName_(name);
    const data  = sheet.getDataRange().getValues();
    const now   = new Date();
    let rowIdx  = -1;
    for (let i = 1; i < data.length; i++) {
      if (normName_(data[i][0]) === key) { rowIdx = i; break; }
    }
    if (rowIdx >= 0) {
      const r = data[rowIdx];
      const merged = [
        name,
        rec.contact   || r[1] || '',
        rec.company   || r[2] || '',
        rec.email     || r[3] || '',
        rec.address   || r[4] || '',
        rec.firstName || r[5] || '',
        rec.lastName  || r[6] || '',
        r[7] || now,   // First Saved (preserved)
        now,           // Last Updated
      ];
      sheet.getRange(rowIdx + 1, 1, 1, merged.length).setValues([merged]);
    } else {
      sheet.appendRow([
        name, rec.contact || '', rec.company || '', rec.email || '', rec.address || '',
        rec.firstName || '', rec.lastName || '', now, now,
      ]);
    }
  } catch (e) {
    Logger.log('upsertCustomerInfo_ error: ' + (e && e.message));
  }
}

// Derive a customer record from any quotation payload and upsert it.
// Handles both the single-name pages (clientName) and Receipt (first/last/company).
function upsertCustomerFromPayload_(data) {
  if (!data) return;
  const first = data.firstName || '';
  const last  = data.lastName  || '';
  const name  = String(data.clientName || [first, last].filter(String).join(' ') || data.company || '').trim();
  if (!name) return;
  upsertCustomerInfo_({
    name:      name,
    contact:   data.contact || data.mobile || '',
    company:   data.company || '',
    email:     data.email   || '',
    address:   data.address || data.delivery || '',
    firstName: first,
    lastName:  last,
  });
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
  if (page === 'sticker') {
    return serveWithToken_('Sticker', 'Quotation System — Sticker', token, appUrl);
  }
  if (page === 'risograph') {
    return serveWithToken_('Risograph', 'Quotation System — Risograph', token, appUrl);
  }
  if (page === 'totebag') {
    return serveWithToken_('Totebag', 'Quotation System — Tote Bag', token, appUrl);
  }
  if (page === 'ticket') {
    return serveWithToken_('Ticket', 'Quotation System — Tickets', token, appUrl);
  }
  if (page === 'newsprint') {
    return serveWithToken_('Newsprint', 'Quotation System — Newsletter & Newspaper', token, appUrl);
  }
  if (page === 'souvenir') {
    return serveWithToken_('Souvenir', 'Quotation System — Souvenir Program', token, appUrl);
  }
  if (page === 'keychain') {
    return serveWithToken_('Keychain', 'Quotation System — Acrylic Keychain', token, appUrl);
  }
  if (page === 'acrylicsign') {
    return serveWithToken_('AcrylicSignage', 'Quotation System — Acrylic Signage', token, appUrl);
  }
  if (page === 'acrylicplate') {
    return serveWithToken_('AcrylicPlate', 'Quotation System — Acrylic Plate Number', token, appUrl);
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

    // ── STICKER QUOTES ──────────────────────────────────────────
    let stickerQuotes = [];
    const stickerSheet = ss.getSheetByName(STICKER_SHEET);
    if (stickerSheet) {
      const sdata  = stickerSheet.getDataRange().getValues();
      const sStart = sdata.length > 0 && String(sdata[0][0]).startsWith('STK-') ? 0 : 1;
      stickerQuotes = sdata.slice(sStart).filter(r => r[0] && String(r[0]).startsWith('STK-')).map(row => {
        let dateStr = '';
        try { dateStr = row[1] ? new Date(row[1]).toISOString() : ''; } catch(e) {}
        const ptLabel = String(row[22] || '');
        return {
          quoteNum:         String(row[0]  || ''),
          date:             dateStr,
          clientName:       String(row[2]  || ''),
          contact:          String(row[3]  || ''),
          email:            String(row[4]  || ''),
          dateNeeded:       String(row[5]  || ''),
          stickerType:      String(row[6]  || ''),
          layout:           String(row[7]  || ''),
          width:            parseFloat(row[8])  || 0,
          height:           parseFloat(row[9])  || 0,
          unit:             String(row[10] || 'in'),
          sqft:             parseFloat(row[11]) || 0,
          quantity:         row[12] || 0,
          ratePerSqft:      parseFloat(row[13]) || 0,
          baseAmount:       parseFloat(row[14]) || 0,
          rushOrder:        String(row[15] || ''),
          rushFee:          parseFloat(row[16]) || 0,
          designService:    String(row[17] || ''),
          designFee:        parseFloat(row[18]) || 0,
          totalAmount:      parseFloat(row[19]) || 0,
          notes:            String(row[20] || ''),
          salesStaff:       String(row[21] || ''),
          paymentTermLabel: ptLabel,
          paymentTermValue: ptLabel.includes('No Down') ? 0 : ptLabel.includes('25%') ? 0.25 : ptLabel.includes('Full') ? 1 : 0.5,
          status:           String(row[23] || 'Pending'),
          approvedBy:       String(row[24] || ''),
          taxType:          String(row[25] || 'non-vat'),
          taxAmount:        parseFloat(row[26]) || 0,
          quoteType:        'sticker',
          signageType:      'Sticker — ' + String(row[6] || ''),
          address: '', delivery: '', lighting: '', material: '',
          mounting: '', mountSurcharge: 0, complexitySurcharge: 0,
          addonDesign: String(row[17] || ''), addonDesignFee: parseFloat(row[18]) || 0,
          addonRush:   String(row[15] || ''), addonRushFee:   parseFloat(row[16]) || 0,
          addonElec: '', addonElecFee: 0,
          addonTransport: '', addonTransportFee: 0,
        };
      });
    }

    // ── RISOGRAPH QUOTES ────────────────────────────────────────
    let risoQuotes = [];
    const risoSheet = ss.getSheetByName(RISOGRAPH_SHEET);
    if (risoSheet) {
      const rdata  = risoSheet.getDataRange().getValues();
      const rStart = rdata.length > 0 && String(rdata[0][0]).startsWith('RG-') ? 0 : 1;
      risoQuotes = rdata.slice(rStart).filter(r => r[0] && String(r[0]).startsWith('RG-')).map(row => {
        let dateStr = '';
        try { dateStr = row[1] ? new Date(row[1]).toISOString() : ''; } catch(e) {}
        const ptLabel = String(row[19] || '');
        return {
          quoteNum:         String(row[0]  || ''),
          date:             dateStr,
          clientName:       String(row[2]  || ''),
          contact:          String(row[3]  || ''),
          email:            String(row[4]  || ''),
          dateNeeded:       String(row[5]  || ''),
          paperType:        String(row[6]  || ''),
          paperSize:        String(row[7]  || ''),
          service:          String(row[8]  || ''),
          sides:            String(row[9]  || ''),
          quantity:         row[10] || 0,
          rate:             parseFloat(row[11]) || 0,
          baseAmount:       parseFloat(row[12]) || 0,
          sortStaple:       String(row[13] || ''),
          sortStapleFee:    parseFloat(row[14]) || 0,
          rushOrder:        String(row[15] || ''),
          rushFee:          parseFloat(row[16]) || 0,
          designService:    String(row[17] || ''),
          designFee:        parseFloat(row[18]) || 0,
          paymentTermLabel: ptLabel,
          paymentTermValue: ptLabel.includes('No Down') ? 0 : ptLabel.includes('25%') ? 0.25 : ptLabel.includes('Full') ? 1 : 0.5,
          totalAmount:      parseFloat(row[20]) || 0,
          notes:            String(row[21] || ''),
          salesStaff:       String(row[22] || ''),
          status:           String(row[23] || 'Pending'),
          approvedBy:       String(row[24] || ''),
          taxType:          String(row[25] || 'non-vat'),
          taxAmount:        parseFloat(row[26]) || 0,
          quoteType:        'risograph',
          signageType:      'Risograph — ' + String(row[6] || '') + ' ' + String(row[7] || ''),
          address: '', delivery: '', lighting: '', material: '',
          mounting: '', mountSurcharge: 0, complexitySurcharge: 0,
          addonDesign: String(row[17] || ''), addonDesignFee: parseFloat(row[18]) || 0,
          addonRush:   String(row[15] || ''), addonRushFee:   parseFloat(row[16]) || 0,
          addonElec: '', addonElecFee: 0,
          addonTransport: '', addonTransportFee: 0,
        };
      });
    }

    // ── TOTE BAG QUOTES ─────────────────────────────────────────
    let totebagQuotes = [];
    const tbSheet = ss.getSheetByName(TOTEBAG_SHEET);
    if (tbSheet) {
      const tdata  = tbSheet.getDataRange().getValues();
      const tStart = tdata.length > 0 && String(tdata[0][0]).startsWith('TB-') ? 0 : 1;
      totebagQuotes = tdata.slice(tStart).filter(r => r[0] && String(r[0]).startsWith('TB-')).map(row => {
        let dateStr = '';
        try { dateStr = row[1] ? new Date(row[1]).toISOString() : ''; } catch(e) {}
        const ptLabel = String(row[16] || '');
        return {
          quoteNum:         String(row[0]  || ''),
          date:             dateStr,
          clientName:       String(row[2]  || ''),
          contact:          String(row[3]  || ''),
          email:            String(row[4]  || ''),
          dateNeeded:       String(row[5]  || ''),
          totebagSize:      String(row[6]  || ''),
          printMethod:      String(row[7]  || ''),
          material:         String(row[8]  || ''),
          quantity:         row[9]  || 0,
          unitPrice:        parseFloat(row[10]) || 0,
          baseAmount:       parseFloat(row[11]) || 0,
          rushOrder:        String(row[12] || ''),
          rushFee:          parseFloat(row[13]) || 0,
          designService:    String(row[14] || ''),
          designFee:        parseFloat(row[15]) || 0,
          paymentTermLabel: ptLabel,
          paymentTermValue: ptLabel.includes('No Down') ? 0 : ptLabel.includes('25%') ? 0.25 : ptLabel.includes('Full') ? 1 : 0.5,
          totalAmount:      parseFloat(row[17]) || 0,
          notes:            String(row[18] || ''),
          salesStaff:       String(row[19] || ''),
          status:           String(row[20] || 'Pending'),
          approvedBy:       String(row[21] || ''),
          taxType:          String(row[22] || 'non-vat'),
          taxAmount:        parseFloat(row[23]) || 0,
          quoteType:        'totebag',
          signageType:      'Tote Bag — ' + String(row[6] || ''),
          address: '', delivery: '', lighting: '', material2: '',
          mounting: '', mountSurcharge: 0, complexitySurcharge: 0,
          addonDesign: String(row[14] || ''), addonDesignFee: parseFloat(row[15]) || 0,
          addonRush:   String(row[12] || ''), addonRushFee:   parseFloat(row[13]) || 0,
          addonElec: '', addonElecFee: 0,
          addonTransport: '', addonTransportFee: 0,
        };
      });
    }

    // ── TICKET QUOTES ───────────────────────────────────────────
    let ticketQuotes = [];
    const tkSheet = ss.getSheetByName(TICKET_SHEET);
    if (tkSheet) {
      const kdata  = tkSheet.getDataRange().getValues();
      const kStart = kdata.length > 0 && String(kdata[0][0]).startsWith('TKT-') ? 0 : 1;
      ticketQuotes = kdata.slice(kStart).filter(r => r[0] && String(r[0]).startsWith('TKT-')).map(row => {
        let dateStr = '';
        try { dateStr = row[1] ? new Date(row[1]).toISOString() : ''; } catch(e) {}
        const ptLabel = String(row[14] || '');
        return {
          quoteNum:         String(row[0]  || ''),
          date:             dateStr,
          clientName:       String(row[2]  || ''),
          contact:          String(row[3]  || ''),
          email:            String(row[4]  || ''),
          dateNeeded:       String(row[5]  || ''),
          ticketType:       String(row[6]  || ''),
          quantity:         row[7]  || 0,
          unitPrice:        parseFloat(row[8])  || 0,
          baseAmount:       parseFloat(row[9])  || 0,
          rushOrder:        String(row[10] || ''),
          rushFee:          parseFloat(row[11]) || 0,
          designService:    String(row[12] || ''),
          designFee:        parseFloat(row[13]) || 0,
          paymentTermLabel: ptLabel,
          paymentTermValue: ptLabel.includes('No Down') ? 0 : ptLabel.includes('25%') ? 0.25 : ptLabel.includes('Full') ? 1 : 0.5,
          totalAmount:      parseFloat(row[15]) || 0,
          notes:            String(row[16] || ''),
          salesStaff:       String(row[17] || ''),
          status:           String(row[18] || 'Pending'),
          approvedBy:       String(row[19] || ''),
          taxType:          String(row[20] || 'non-vat'),
          taxAmount:        parseFloat(row[21]) || 0,
          quoteType:        'ticket',
          signageType:      'Tickets — ' + String(row[6] || ''),
          address: '', delivery: '', lighting: '',
          mounting: '', mountSurcharge: 0, complexitySurcharge: 0,
          addonDesign: String(row[12] || ''), addonDesignFee: parseFloat(row[13]) || 0,
          addonRush:   String(row[10] || ''), addonRushFee:   parseFloat(row[11]) || 0,
          addonElec: '', addonElecFee: 0,
          addonTransport: '', addonTransportFee: 0,
        };
      });
    }

    // ── NEWSLETTER / NEWSPAPER QUOTES ───────────────────────────
    let newsprintQuotes = [];
    const npSheet = ss.getSheetByName(NEWSPRINT_SHEET);
    if (npSheet) {
      const ndata  = npSheet.getDataRange().getValues();
      const nStart = ndata.length > 0 && String(ndata[0][0]).startsWith('NL-') ? 0 : 1;
      newsprintQuotes = ndata.slice(nStart).filter(r => r[0] && String(r[0]).startsWith('NL-')).map(row => {
        let dateStr = '';
        try { dateStr = row[1] ? new Date(row[1]).toISOString() : ''; } catch(e) {}
        const ptLabel = String(row[17] || '');
        return {
          quoteNum:         String(row[0]  || ''),
          date:             dateStr,
          clientName:       String(row[2]  || ''),
          contact:          String(row[3]  || ''),
          email:            String(row[4]  || ''),
          dateNeeded:       String(row[5]  || ''),
          newsCategory:     String(row[6]  || ''),
          newsOption:       String(row[7]  || ''),
          newsSize:         String(row[8]  || ''),
          newsMaterial:     String(row[9]  || ''),
          quantity:         row[10] || 0,
          unitPrice:        parseFloat(row[11]) || 0,
          baseAmount:       parseFloat(row[12]) || 0,
          rushOrder:        String(row[13] || ''),
          rushFee:          parseFloat(row[14]) || 0,
          designService:    String(row[15] || ''),
          designFee:        parseFloat(row[16]) || 0,
          paymentTermLabel: ptLabel,
          paymentTermValue: ptLabel.includes('No Down') ? 0 : ptLabel.includes('25%') ? 0.25 : ptLabel.includes('Full') ? 1 : 0.5,
          totalAmount:      parseFloat(row[18]) || 0,
          notes:            String(row[19] || ''),
          salesStaff:       String(row[20] || ''),
          status:           String(row[21] || 'Pending'),
          approvedBy:       String(row[22] || ''),
          taxType:          String(row[23] || 'non-vat'),
          taxAmount:        parseFloat(row[24]) || 0,
          quoteType:        'newsprint',
          signageType:      (String(row[6] || 'Newsprint')) + ' — ' + String(row[7] || ''),
          address: '', delivery: '', lighting: '',
          mounting: '', mountSurcharge: 0, complexitySurcharge: 0,
          addonDesign: String(row[15] || ''), addonDesignFee: parseFloat(row[16]) || 0,
          addonRush:   String(row[13] || ''), addonRushFee:   parseFloat(row[14]) || 0,
          addonElec: '', addonElecFee: 0,
          addonTransport: '', addonTransportFee: 0,
        };
      });
    }

    // ── SOUVENIR PROGRAM QUOTES ─────────────────────────────────
    let souvenirQuotes = [];
    const spSheet = ss.getSheetByName(SOUVENIR_SHEET);
    if (spSheet) {
      const sdata2  = spSheet.getDataRange().getValues();
      const sStart2 = sdata2.length > 0 && String(sdata2[0][0]).startsWith('SP-') ? 0 : 1;
      souvenirQuotes = sdata2.slice(sStart2).filter(r => r[0] && String(r[0]).startsWith('SP-')).map(row => {
        let dateStr = '';
        try { dateStr = row[1] ? new Date(row[1]).toISOString() : ''; } catch(e) {}
        const ptLabel = String(row[17] || '');
        return {
          quoteNum:         String(row[0]  || ''),
          date:             dateStr,
          clientName:       String(row[2]  || ''),
          contact:          String(row[3]  || ''),
          email:            String(row[4]  || ''),
          dateNeeded:       String(row[5]  || ''),
          souvenirMaterial: String(row[6]  || ''),
          souvenirPageSize: String(row[7]  || ''),
          souvenirMethod:   String(row[8]  || ''),
          souvenirPages:    row[9]  || 0,
          quantity:         row[10] || 0,
          unitPrice:        parseFloat(row[11]) || 0,
          baseAmount:       parseFloat(row[12]) || 0,
          rushOrder:        String(row[13] || ''),
          rushFee:          parseFloat(row[14]) || 0,
          designService:    String(row[15] || ''),
          designFee:        parseFloat(row[16]) || 0,
          paymentTermLabel: ptLabel,
          paymentTermValue: ptLabel.includes('No Down') ? 0 : ptLabel.includes('25%') ? 0.25 : ptLabel.includes('Full') ? 1 : 0.5,
          totalAmount:      parseFloat(row[18]) || 0,
          notes:            String(row[19] || ''),
          salesStaff:       String(row[20] || ''),
          status:           String(row[21] || 'Pending'),
          approvedBy:       String(row[22] || ''),
          taxType:          String(row[23] || 'non-vat'),
          taxAmount:        parseFloat(row[24]) || 0,
          quoteType:        'souvenir',
          signageType:      'Souvenir Program — ' + String(row[6] || ''),
          address: '', delivery: '', lighting: '',
          mounting: '', mountSurcharge: 0, complexitySurcharge: 0,
          addonDesign: String(row[15] || ''), addonDesignFee: parseFloat(row[16]) || 0,
          addonRush:   String(row[13] || ''), addonRushFee:   parseFloat(row[14]) || 0,
          addonElec: '', addonElecFee: 0,
          addonTransport: '', addonTransportFee: 0,
        };
      });
    }

    // ── ACRYLIC KEYCHAIN QUOTES ─────────────────────────────────
    let keychainQuotes = [];
    const kcSheet = ss.getSheetByName(KEYCHAIN_SHEET);
    if (kcSheet) {
      const cdata  = kcSheet.getDataRange().getValues();
      const cStart = cdata.length > 0 && String(cdata[0][0]).startsWith('KC-') ? 0 : 1;
      keychainQuotes = cdata.slice(cStart).filter(r => r[0] && String(r[0]).startsWith('KC-')).map(row => {
        let dateStr = '';
        try { dateStr = row[1] ? new Date(row[1]).toISOString() : ''; } catch(e) {}
        const ptLabel = String(row[18] || '');
        return {
          quoteNum:         String(row[0]  || ''),
          date:             dateStr,
          clientName:       String(row[2]  || ''),
          contact:          String(row[3]  || ''),
          email:            String(row[4]  || ''),
          dateNeeded:       String(row[5]  || ''),
          keychainSize:     String(row[6]  || ''),
          keychainSqin:     parseFloat(row[7]) || 0,
          keychainCutType:  String(row[8]  || ''),
          keychainMaterial: String(row[9]  || ''),
          quantity:         row[10] || 0,
          unitPrice:        parseFloat(row[11]) || 0,
          baseAmount:       parseFloat(row[12]) || 0,
          rushOrder:        String(row[13] || ''),
          rushFee:          parseFloat(row[14]) || 0,
          designService:    String(row[15] || ''),
          designFee:        parseFloat(row[16]) || 0,
          keychainDesignRef:String(row[17] || ''),
          paymentTermLabel: ptLabel,
          paymentTermValue: ptLabel.includes('No Down') ? 0 : ptLabel.includes('25%') ? 0.25 : ptLabel.includes('Full') ? 1 : 0.5,
          totalAmount:      parseFloat(row[19]) || 0,
          notes:            String(row[20] || ''),
          salesStaff:       String(row[21] || ''),
          status:           String(row[22] || 'Pending'),
          approvedBy:       String(row[23] || ''),
          taxType:          String(row[24] || 'non-vat'),
          taxAmount:        parseFloat(row[25]) || 0,
          quoteType:        'keychain',
          signageType:      'Acrylic Keychain — ' + String(row[6] || ''),
          address: '', delivery: '', lighting: '',
          mounting: '', mountSurcharge: 0, complexitySurcharge: 0,
          addonDesign: String(row[15] || ''), addonDesignFee: parseFloat(row[16]) || 0,
          addonRush:   String(row[13] || ''), addonRushFee:   parseFloat(row[14]) || 0,
          addonElec: '', addonElecFee: 0,
          addonTransport: '', addonTransportFee: 0,
        };
      });
    }

    // ── ACRYLIC SIGNAGE QUOTES ──────────────────────────────────
    let acrylicSignQuotes = [];
    const asSheet = ss.getSheetByName(ACRYLICSIGN_SHEET);
    if (asSheet) {
      const adata  = asSheet.getDataRange().getValues();
      const aStart = adata.length > 0 && String(adata[0][0]).startsWith('AS-') ? 0 : 1;
      acrylicSignQuotes = adata.slice(aStart).filter(r => r[0] && String(r[0]).startsWith('AS-')).map(row => {
        let dateStr = '';
        try { dateStr = row[1] ? new Date(row[1]).toISOString() : ''; } catch(e) {}
        const ptLabel = String(row[20] || '');
        const typeName = String(row[6] || '');
        return {
          quoteNum:         String(row[0]  || ''),
          date:             dateStr,
          clientName:       String(row[2]  || ''),
          contact:          String(row[3]  || ''),
          email:            String(row[4]  || ''),
          dateNeeded:       String(row[5]  || ''),
          acsignType:       typeName,
          width:            parseFloat(row[7]) || 0,
          height:           parseFloat(row[8]) || 0,
          unit:             String(row[9] || 'ft'),
          sqft:             parseFloat(row[10]) || 0,
          acsignBilledSqft: parseFloat(row[11]) || 0,
          quantity:         row[12] || 0,
          ratePerSqft:      parseFloat(row[13]) || 0,
          unitPrice:        parseFloat(row[14]) || 0,
          baseAmount:       parseFloat(row[15]) || 0,
          rushOrder:        String(row[16] || ''),
          rushFee:          parseFloat(row[17]) || 0,
          designService:    String(row[18] || ''),
          designFee:        parseFloat(row[19]) || 0,
          paymentTermLabel: ptLabel,
          paymentTermValue: ptLabel.includes('No Down') ? 0 : ptLabel.includes('25%') ? 0.25 : ptLabel.includes('Full') ? 1 : 0.5,
          totalAmount:      parseFloat(row[21]) || 0,
          notes:            String(row[22] || ''),
          salesStaff:       String(row[23] || ''),
          status:           String(row[24] || 'Pending'),
          approvedBy:       String(row[25] || ''),
          taxType:          String(row[26] || 'non-vat'),
          taxAmount:        parseFloat(row[27]) || 0,
          quoteType:        'acrylicsign',
          signageType:      /acrylic|signage/i.test(typeName) ? typeName : ('Acrylic Signage — ' + typeName),
          address: '', delivery: '', lighting: '', material: '',
          mounting: '', mountSurcharge: 0, complexitySurcharge: 0,
          addonDesign: String(row[18] || ''), addonDesignFee: parseFloat(row[19]) || 0,
          addonRush:   String(row[16] || ''), addonRushFee:   parseFloat(row[17]) || 0,
          addonElec: '', addonElecFee: 0,
          addonTransport: '', addonTransportFee: 0,
        };
      });
    }

    // ── ACRYLIC PLATE NUMBER QUOTES ─────────────────────────────
    let acrylicPlateQuotes = [];
    const apSheet = ss.getSheetByName(ACRYLICPLATE_SHEET);
    if (apSheet) {
      const pdata  = apSheet.getDataRange().getValues();
      const pStart = pdata.length > 0 && String(pdata[0][0]).startsWith('AP-') ? 0 : 1;
      acrylicPlateQuotes = pdata.slice(pStart).filter(r => r[0] && String(r[0]).startsWith('AP-')).map(row => {
        let dateStr = '';
        try { dateStr = row[1] ? new Date(row[1]).toISOString() : ''; } catch(e) {}
        const ptLabel = String(row[15] || '');
        const plateType = String(row[6] || '');
        return {
          quoteNum:         String(row[0]  || ''),
          date:             dateStr,
          clientName:       String(row[2]  || ''),
          contact:          String(row[3]  || ''),
          email:            String(row[4]  || ''),
          dateNeeded:       String(row[5]  || ''),
          plateType:        plateType,
          plateText:        String(row[7] || ''),
          quantity:         row[8]  || 0,
          unitPrice:        parseFloat(row[9])  || 0,
          baseAmount:       parseFloat(row[10]) || 0,
          rushOrder:        String(row[11] || ''),
          rushFee:          parseFloat(row[12]) || 0,
          designService:    String(row[13] || ''),
          designFee:        parseFloat(row[14]) || 0,
          paymentTermLabel: ptLabel,
          paymentTermValue: ptLabel.includes('No Down') ? 0 : ptLabel.includes('25%') ? 0.25 : ptLabel.includes('Full') ? 1 : 0.5,
          totalAmount:      parseFloat(row[16]) || 0,
          notes:            String(row[17] || ''),
          salesStaff:       String(row[18] || ''),
          status:           String(row[19] || 'Pending'),
          approvedBy:       String(row[20] || ''),
          taxType:          String(row[21] || 'non-vat'),
          taxAmount:        parseFloat(row[22]) || 0,
          items: (function(){ try { const j = String(row[23]||''); if (!j || j==='[]') return []; return JSON.parse(j); } catch(e){ return []; } })(),
          quoteType:        'acrylicplate',
          signageType:      'Acrylic Plate — ' + plateType,
          address: '', delivery: '', lighting: '', material: '',
          mounting: '', mountSurcharge: 0, complexitySurcharge: 0,
          addonDesign: String(row[13] || ''), addonDesignFee: parseFloat(row[14]) || 0,
          addonRush:   String(row[11] || ''), addonRushFee:   parseFloat(row[12]) || 0,
          addonElec: '', addonElecFee: 0,
          addonTransport: '', addonTransportFee: 0,
        };
      });
    }

    // ── COMBINE & FILTER ────────────────────────────────────────
    const allQuotes = [...quotes, ...tarpQuotes, ...receiptQuotes, ...bookbindQuotes, ...frameQuotes, ...tshirtQuotes, ...mugQuotes, ...stickerQuotes, ...risoQuotes, ...totebagQuotes, ...ticketQuotes, ...newsprintQuotes, ...souvenirQuotes, ...keychainQuotes, ...acrylicSignQuotes, ...acrylicPlateQuotes];

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

  // ── Newer products (Mug, Sticker, Risograph, Tote Bag, Tickets,
  //     Newsletter/Newspaper, Souvenir, Keychain) — reuse the dashboard
  //     data builder, which already returns a render-ready quote object. ──
  const PDF_VIA_DASHBOARD = ['MUG-', 'STK-', 'RG-', 'TB-', 'TKT-', 'NL-', 'SP-', 'KC-', 'AS-', 'AP-'];
  if (PDF_VIA_DASHBOARD.some(function(p){ return qn.indexOf(p) === 0; })) {
    try {
      const dash = getDashboardData(token);
      if (dash && dash.quotes && dash.quotes.length) {
        const found = dash.quotes.find(function(x){ return String(x.quoteNum).trim() === qn; });
        if (found) {
          // The PDF renders Design Fee from `designFee` and Rush from the add-on
          // section — clear the add-on Design mirror so it isn't billed twice.
          found.addonDesignFee = 0;
          found.addonDesign    = '';
          return found;
        }
      }
    } catch(e) { Logger.log('getQuoteForPDF dashboard fallback error: ' + e); }
    return null;
  }

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
        textColor:             String(r[30] || ''),   // AE col 31 - Text Color
        fontStyle:             String(r[31] || ''),   // AF col 32 - Font Style
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
  const isMug       = String(quoteNum).startsWith('MUG-');
  const isSticker   = String(quoteNum).startsWith('STK-');
  const isRiso      = String(quoteNum).startsWith('RG-');
  const isTotebag   = String(quoteNum).startsWith('TB-');
  const isTicket    = String(quoteNum).startsWith('TKT-');
  const isNewsprint = String(quoteNum).startsWith('NL-');
  const isSouvenir  = String(quoteNum).startsWith('SP-');
  const isKeychain  = String(quoteNum).startsWith('KC-');
  const isAcrylicSign = String(quoteNum).startsWith('AS-');
  const isAcrylicPlate = String(quoteNum).startsWith('AP-');

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

  if (isMug) {
    const sheet = ss.getSheetByName(MUG_SHEET);
    if (!sheet) throw new Error('Mug Quotations sheet not found.');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === quoteNum) {
        sheet.getRange(i+1, 22).setValue(status);  // col V = Status
        sheet.getRange(i+1, 23).setValue(session.name + ' — ' + new Date().toLocaleString('en-PH'));  // col W = Approved By
        const color = status === 'Approved' ? '#E6FFF3' : status === 'Rejected' ? '#FFF0F0' : '#FFFFFF';
        sheet.getRange(i+1, 1, 1, 26).setBackground(color);
        return { success: true };
      }
    }
    throw new Error('Mug order not found: ' + quoteNum);
  }

  if (isSticker) {
    const sheet = ss.getSheetByName(STICKER_SHEET);
    if (!sheet) throw new Error('Sticker Quotations sheet not found.');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === quoteNum) {
        sheet.getRange(i+1, 24).setValue(status);  // col X = Status
        sheet.getRange(i+1, 25).setValue(session.name + ' — ' + new Date().toLocaleString('en-PH'));  // col Y = Approved By
        const color = status === 'Approved' ? '#E6FFF3' : status === 'Rejected' ? '#FFF0F0' : '#FFFFFF';
        sheet.getRange(i+1, 1, 1, 27).setBackground(color);
        return { success: true };
      }
    }
    throw new Error('Sticker order not found: ' + quoteNum);
  }

  if (isRiso) {
    const sheet = ss.getSheetByName(RISOGRAPH_SHEET);
    if (!sheet) throw new Error('Risograph Quotations sheet not found.');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === quoteNum) {
        sheet.getRange(i+1, 24).setValue(status);  // col X = Status
        sheet.getRange(i+1, 25).setValue(session.name + ' — ' + new Date().toLocaleString('en-PH'));  // col Y = Approved By
        const color = status === 'Approved' ? '#E6FFF3' : status === 'Rejected' ? '#FFF0F0' : '#FFFFFF';
        sheet.getRange(i+1, 1, 1, 27).setBackground(color);
        return { success: true };
      }
    }
    throw new Error('Risograph order not found: ' + quoteNum);
  }

  if (isTotebag) {
    const sheet = ss.getSheetByName(TOTEBAG_SHEET);
    if (!sheet) throw new Error('Totebag Quotations sheet not found.');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === quoteNum) {
        sheet.getRange(i+1, 21).setValue(status);  // col U = Status
        sheet.getRange(i+1, 22).setValue(session.name + ' — ' + new Date().toLocaleString('en-PH'));  // col V = Approved By
        const color = status === 'Approved' ? '#E6FFF3' : status === 'Rejected' ? '#FFF0F0' : '#FFFFFF';
        sheet.getRange(i+1, 1, 1, 24).setBackground(color);
        return { success: true };
      }
    }
    throw new Error('Totebag order not found: ' + quoteNum);
  }

  if (isTicket) {
    const sheet = ss.getSheetByName(TICKET_SHEET);
    if (!sheet) throw new Error('Ticket Quotations sheet not found.');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === quoteNum) {
        sheet.getRange(i+1, 19).setValue(status);  // col S = Status
        sheet.getRange(i+1, 20).setValue(session.name + ' — ' + new Date().toLocaleString('en-PH'));  // col T = Approved By
        const color = status === 'Approved' ? '#E6FFF3' : status === 'Rejected' ? '#FFF0F0' : '#FFFFFF';
        sheet.getRange(i+1, 1, 1, 22).setBackground(color);
        return { success: true };
      }
    }
    throw new Error('Ticket order not found: ' + quoteNum);
  }

  if (isNewsprint) {
    const sheet = ss.getSheetByName(NEWSPRINT_SHEET);
    if (!sheet) throw new Error('Newsprint Quotations sheet not found.');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === quoteNum) {
        sheet.getRange(i+1, 22).setValue(status);  // col V = Status
        sheet.getRange(i+1, 23).setValue(session.name + ' — ' + new Date().toLocaleString('en-PH'));  // col W = Approved By
        const color = status === 'Approved' ? '#E6FFF3' : status === 'Rejected' ? '#FFF0F0' : '#FFFFFF';
        sheet.getRange(i+1, 1, 1, 25).setBackground(color);
        return { success: true };
      }
    }
    throw new Error('Newsprint order not found: ' + quoteNum);
  }

  if (isSouvenir) {
    const sheet = ss.getSheetByName(SOUVENIR_SHEET);
    if (!sheet) throw new Error('Souvenir Quotations sheet not found.');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === quoteNum) {
        sheet.getRange(i+1, 22).setValue(status);  // col V = Status
        sheet.getRange(i+1, 23).setValue(session.name + ' — ' + new Date().toLocaleString('en-PH'));  // col W = Approved By
        const color = status === 'Approved' ? '#E6FFF3' : status === 'Rejected' ? '#FFF0F0' : '#FFFFFF';
        sheet.getRange(i+1, 1, 1, 25).setBackground(color);
        return { success: true };
      }
    }
    throw new Error('Souvenir order not found: ' + quoteNum);
  }

  if (isKeychain) {
    const sheet = ss.getSheetByName(KEYCHAIN_SHEET);
    if (!sheet) throw new Error('Keychain Quotations sheet not found.');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === quoteNum) {
        sheet.getRange(i+1, 23).setValue(status);  // col W = Status
        sheet.getRange(i+1, 24).setValue(session.name + ' — ' + new Date().toLocaleString('en-PH'));  // col X = Approved By
        const color = status === 'Approved' ? '#E6FFF3' : status === 'Rejected' ? '#FFF0F0' : '#FFFFFF';
        sheet.getRange(i+1, 1, 1, 26).setBackground(color);
        return { success: true };
      }
    }
    throw new Error('Keychain order not found: ' + quoteNum);
  }

  if (isAcrylicSign) {
    const sheet = ss.getSheetByName(ACRYLICSIGN_SHEET);
    if (!sheet) throw new Error('Acrylic Signage Quotations sheet not found.');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === quoteNum) {
        sheet.getRange(i+1, 25).setValue(status);  // col Y = Status
        sheet.getRange(i+1, 26).setValue(session.name + ' — ' + new Date().toLocaleString('en-PH'));  // col Z = Approved By
        const color = status === 'Approved' ? '#E6FFF3' : status === 'Rejected' ? '#FFF0F0' : '#FFFFFF';
        sheet.getRange(i+1, 1, 1, 28).setBackground(color);
        return { success: true };
      }
    }
    throw new Error('Acrylic Signage order not found: ' + quoteNum);
  }

  if (isAcrylicPlate) {
    const sheet = ss.getSheetByName(ACRYLICPLATE_SHEET);
    if (!sheet) throw new Error('Acrylic Plate Quotations sheet not found.');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === quoteNum) {
        sheet.getRange(i+1, 20).setValue(status);  // col T = Status
        sheet.getRange(i+1, 21).setValue(session.name + ' — ' + new Date().toLocaleString('en-PH'));  // col U = Approved By
        const color = status === 'Approved' ? '#E6FFF3' : status === 'Rejected' ? '#FFF0F0' : '#FFFFFF';
        sheet.getRange(i+1, 1, 1, 23).setBackground(color);
        return { success: true };
      }
    }
    throw new Error('Acrylic Plate order not found: ' + quoteNum);
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
      'Tax Type', 'Tax Amount', 'Notes',
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#E8151B').setFontColor('#fff')
      .setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
  }

  lockQuoteNumbering_();
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
    data.notes        || '',          // AE col 31 - Notes / Special Instructions
  ]);

  sheet.getRange(sheet.getLastRow(), 18, 1, 4).setNumberFormat('₱#,##0.00');
  try { notifyQuoteSaved_(quoteNum, 'Tarpaulin', Object.assign({ totalAmount: totalAmt }, data)); } catch(_) {}
  return quoteNum;
}

// ══════════════════════════════════════════════════════════════════
//  TARPAULIN PRICING
// ══════════════════════════════════════════════════════════════════
function getTarpPricing() {
  const defaults = { ratePerSqft: 50, rushFee: 150, designFee: 250 };
  try {
    const ss    = getPriceDbSS_();
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
    const ss    = getPriceDbSS_();
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

  lockQuoteNumbering_();
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
  try { notifyQuoteSaved_(quoteNum, 'Frame', data); } catch(_) {}
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
    const ss = getPriceDbSS_();
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

  lockQuoteNumbering_();
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
  try { notifyQuoteSaved_(quoteNum, 'T-Shirt', data); } catch(_) {}
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
    sticker:  (function(){ try{ return getStickerPricing(); }catch(e){ return null; } })(),
    risograph:(function(){ try{ return getRisographPricing(); }catch(e){ return null; } })(),
    receipt:  (function(){ try{ return getReceiptPricing(); }catch(e){ return null; } })(),
    signage:  (function(){ try{ return getPricing(); }      catch(e){ return null; } })(),
    totebag:  (function(){ try{ return getTotebagPricing();  }catch(e){ return null; } })(),
    ticket:   (function(){ try{ return getTicketPricing();   }catch(e){ return null; } })(),
    newsprint:(function(){ try{ return getNewsprintPricing();}catch(e){ return null; } })(),
    souvenir: (function(){ try{ return getSouvenirPricing(); }catch(e){ return null; } })(),
    keychain: (function(){ try{ return getKeychainPricing(); }catch(e){ return null; } })(),
    acrylicsign:(function(){ try{ return getAcrylicSignPricing(); }catch(e){ return null; } })(),
    acrylicplate:(function(){ try{ return getAcrylicPlatePricing(); }catch(e){ return null; } })(),
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

    lockQuoteNumbering_();
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
      if (data.eyelet) specs += ' | Eyelet: ' + data.eyelet;
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
    } else if (data.productType === 'sticker') {
      specs = 'Sticker · ' + (data.width || '?') + ' × ' + (data.height || '?') + ' ' + (data.unit || 'in') + ' × ' + (data.quantity || 1) + ' pc(s)';
    } else if (data.productType === 'risograph') {
      specs = (data.paperType || '—') + ' · ' + (data.paperSize || '—') + ' · ' + (data.service || '—') + ' · ' + (data.sides || '—') + ' × ' + (data.quantity || 1) + ' ream(s)';
      if (data.sortStaple === 'Yes') specs += ' | Sort & Staple';
    } else if (data.productType === 'signage') {
      specs = (data.signageType || '—') + ' · ' + (data.width || '?') + ' × ' + (data.height || '?') + ' ' + (data.unit || 'ft') + ' × ' + (data.quantity || 1) + ' pc(s)';
      if (data.lighting)  specs += ' | ' + data.lighting;
      if (data.material)  specs += ' | ' + data.material;
      if (data.mounting)  specs += ' | Mount: ' + data.mounting;
      if (data.transport) specs += ' | Transport: ' + data.transport + (data.transportLocation ? ' (' + data.transportLocation + ')' : '');
      if (data.electrical) specs += ' | Electrical: ' + data.electrical;
      if (data.designService === 'Yes') specs += ' | Design';
    } else if (data.productType === 'totebag') {
      specs = 'Tote Bag · ' + (data.totebagSize || '—') + ' · ' + (data.printMethod || 'Sublimation')
            + ' · ' + (data.material || 'Canvas') + ' × ' + (data.quantity || 1) + ' pc(s)';
    } else if (data.productType === 'ticket') {
      specs = 'Tickets · ' + (data.ticketType || '—') + ' × ' + (data.quantity || 1) + ' pc(s)';
    } else if (data.productType === 'newsprint') {
      specs = (data.category || 'Newsprint') + ' · ' + (data.optionLabel || '—')
            + (data.size ? ' · ' + data.size : '') + ' × ' + (data.quantity || 1) + ' copies';
    } else if (data.productType === 'souvenir') {
      specs = 'Souvenir Program · ' + (data.material || '—') + ' · ' + (data.pages || '?') + ' pages ('
            + (data.pageSize || 'A3') + ') × ' + (data.quantity || 1) + ' copies';
    } else if (data.productType === 'keychain') {
      specs = 'Acrylic Keychain · ' + (data.size || '—') + ' · ' + (data.cutType || 'Standard')
            + ' × ' + (data.quantity || 1) + ' pc(s)';
      if (data.designRef) specs += ' | Ref: ' + data.designRef;
    } else if (data.productType === 'acrylicsign') {
      specs = 'Acrylic Signage · ' + (data.signageType || '—') + ' · ' + (data.width || '?') + ' × '
            + (data.height || '?') + ' ' + (data.unit || 'ft') + ' × ' + (data.quantity || 1) + ' pc(s)';
    } else if (data.productType === 'acrylicplate') {
      specs = 'Acrylic Plate Number · ' + (data.plateType || '—') + ' × ' + (data.quantity || 1) + ' pc(s)';
      if (data.plateText) specs += ' | Text: ' + data.plateText;
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

    // ── Decode optional proof-of-payment attachment ──
    let proofBlob = null;
    if (data.proofFileBase64 && data.proofFileName) {
      try {
        const bytes = Utilities.base64Decode(data.proofFileBase64);
        proofBlob = Utilities.newBlob(
          bytes,
          data.proofFileMime || 'application/octet-stream',
          data.proofFileName
        );
      } catch(e) { Logger.log('Proof of payment decode error: ' + e.message); }
    }

    // ── Notify Ormoc Printshoppe of every new customer quote submission ──
    try { notifyCustomerSubmission_(reqNum, type, data, specs, total, proofBlob); } catch(_) {}

    return reqNum;
  } catch(e) {
    throw new Error('Submit failed: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════════
//  NOTIFY ORMOC PRINTSHOPPE — sales-side save (internal quote logged)
// ══════════════════════════════════════════════════════════════════
function notifyQuoteSaved_(quoteNum, productType, data) {
  // Record / update the customer on every save (best-effort, never blocks).
  try { upsertCustomerFromPayload_(data); } catch (e) {}
  try {
    const client  = String(data.clientName || '—');
    const contact = String(data.contact    || '—');
    const email   = String(data.email      || '');
    const dateN   = String(data.dateNeeded || '—');
    const notes   = String(data.notes      || '');
    const total   = parseFloat(data.totalAmount) || 0;
    const stamp   = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Manila', 'yyyy-MM-dd HH:mm');
    const totalPHP = '₱' + total.toLocaleString('en-PH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

    const subject = '📝 New Quote Saved — ' + productType + ' — ' + client + ' (' + quoteNum + ')';
    const body =
      'A new quotation was saved in the system.\n\n' +
      '── Reference ────────────────────────────\n' +
      'Quote #: ' + quoteNum    + '\n' +
      'Product: ' + productType + '\n' +
      'Time:    ' + stamp       + '\n\n' +
      '── Client ───────────────────────────────\n' +
      'Name:    ' + client  + '\n' +
      'Contact: ' + contact + '\n' +
      (email ? 'Email:   ' + email + '\n' : '') +
      'Date Needed: ' + dateN + '\n\n' +
      '── Total ────────────────────────────────\n' +
      'Total: ' + totalPHP + '\n' +
      (notes ? '\nNotes:\n' + notes + '\n' : '') +
      '\nOpen the Dashboard to view full details.';

    MailApp.sendEmail({
      to:      'ormocprintshoppe@gmail.com',
      subject: subject,
      body:    body,
    });
    Logger.log('notifyQuoteSaved_ sent for ' + quoteNum);
  } catch(e) {
    Logger.log('notifyQuoteSaved_ ERROR for ' + quoteNum + ': ' + (e && e.message));
  }
}

// ══════════════════════════════════════════════════════════════════
//  ONE-TIME TEST — run from Apps Script editor to verify MailApp works
// ══════════════════════════════════════════════════════════════════
function testEmailNotif() {
  try {
    MailApp.sendEmail({
      to:      'ormocprintshoppe@gmail.com',
      subject: '🧪 TEST — MailApp from Quotation System',
      body:    'If you received this email, MailApp permissions are working.\n\nSent at ' +
               Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Manila', 'yyyy-MM-dd HH:mm:ss'),
    });
    Logger.log('TEST email sent — remaining daily quota: ' + MailApp.getRemainingDailyQuota());
    return 'OK — test email sent to ormocprintshoppe@gmail.com. Remaining quota: ' + MailApp.getRemainingDailyQuota();
  } catch(e) {
    Logger.log('testEmailNotif ERROR: ' + (e && e.message));
    return 'ERROR: ' + (e && e.message);
  }
}

// ══════════════════════════════════════════════════════════════════
//  NOTIFY ORMOC PRINTSHOPPE — sent after every customer submission
// ══════════════════════════════════════════════════════════════════
function notifyCustomerSubmission_(reqNum, productType, data, specs, total, proofBlob) {
  // Record / update the customer on every portal submission too (best-effort, never blocks).
  try { upsertCustomerFromPayload_(data); } catch (e) {}
  try {
    const client  = String(data.clientName || '—');
    const contact = String(data.contact    || '—');
    const email   = String(data.email      || '');
    const notes   = String(data.notes      || '');
    const dateN   = String(data.dateNeeded || '—');
    const stamp   = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Manila', 'yyyy-MM-dd HH:mm');
    const totalPHP = '₱' + (parseFloat(total) || 0).toLocaleString('en-PH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const hasProof = !!(proofBlob && proofBlob.getName);

    const subject = (hasProof ? '💰 ' : '🆕 ') + 'New Customer Quote — ' + productType + ' — ' + client + ' (' + reqNum + ')';
    const body =
      'A new customer quote request was submitted via the online portal.\n\n' +
      '── Reference ────────────────────────────\n' +
      'Ref #:   ' + reqNum  + '\n' +
      'Product: ' + productType + '\n' +
      'Time:    ' + stamp   + '\n\n' +
      '── Customer ─────────────────────────────\n' +
      'Name:    ' + client  + '\n' +
      'Contact: ' + contact + '\n' +
      (email ? 'Email:   ' + email + '\n' : '') +
      'Date Needed: ' + dateN + '\n\n' +
      '── Specs ────────────────────────────────\n' +
      (specs || '—') + '\n\n' +
      '── Estimate ─────────────────────────────\n' +
      'Total: ' + totalPHP + '\n' +
      (notes ? '\nNotes:\n' + notes + '\n' : '') +
      (hasProof ? '\n📎 Proof of Payment attached: ' + proofBlob.getName() + '\n' : '\n(No proof of payment uploaded.)\n') +
      '\nOpen the Dashboard → Customer Quotes to follow up.';

    const mailOpts = {
      to:      'ormocprintshoppe@gmail.com',
      subject: subject,
      body:    body,
      replyTo: (email && /@/.test(email)) ? email : undefined,
    };
    if (hasProof) mailOpts.attachments = [proofBlob];

    MailApp.sendEmail(mailOpts);
  } catch(e) {
    Logger.log('notifyCustomerSubmission_ error: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════════
//  CUSTOMER DESIGN FILE UPLOAD — emails the file to ormocprintshoppe@gmail.com
// ══════════════════════════════════════════════════════════════════

// Run this ONCE from the Apps Script editor to grant MailApp permission.
// Function dropdown → select __authorizeMailScope → click ▶ Run → approve.
function __authorizeMailScope() {
  MailApp.sendEmail({
    to:      'ormocprintshoppe@gmail.com',
    subject: '✅ Mail permission authorized',
    body:    'This is a one-time test to grant script.send_mail permission. You can ignore this email.',
  });
  return 'OK — sent test email. MailApp scope is now authorized.';
}

function uploadCustomerDesign(data) {
  try {
    if (!data || !data.base64 || !data.filename) {
      return { success: false, message: 'No file received.' };
    }
    const bytes   = Utilities.base64Decode(data.base64);
    const mime    = data.mimeType || 'application/octet-stream';
    const blob    = Utilities.newBlob(bytes, mime, data.filename);
    const product = String(data.productType || 'Quotation').toUpperCase();
    const client  = String(data.clientName  || '—');
    const contact = String(data.contact     || '—');
    const email   = String(data.email       || '—');
    const notes   = String(data.notes       || '');
    const specs   = String(data.specs       || '');
    const stamp   = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Manila', 'yyyy-MM-dd HH:mm');

    const subject = 'Attention Sales New Inquiry ' + product + ' — ' + client;
    const body =
      'A customer uploaded a design file via the ' + product + ' quote portal.\n\n' +
      '── Customer ──────────────────────────────\n' +
      'Name:    ' + client  + '\n' +
      'Contact: ' + contact + '\n' +
      'Email:   ' + email   + '\n' +
      'Product: ' + product + '\n' +
      'Time:    ' + stamp   + '\n' +
      (specs ? '\n── Specifications ────────────────────────\n' + specs + '\n' : '') +
      (notes ? '\nNotes:\n' + notes + '\n' : '') +
      '\nFilename: ' + data.filename + '\n' +
      'See the attached file.';

    MailApp.sendEmail({
      to:          'ormocprintshoppe@gmail.com',
      subject:     subject,
      body:        body,
      attachments: [blob],
      replyTo:     (email && /@/.test(email)) ? email : undefined,
    });
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
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
    const ss    = getPriceDbSS_();
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

  lockQuoteNumbering_();
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
  try { notifyQuoteSaved_(quoteNum, 'Mug', data); } catch(_) {}
  return quoteNum;
}

// ══════════════════════════════════════════════════════════════════
//  GET STICKER PRICING
// ══════════════════════════════════════════════════════════════════
function getStickerPricing() {
  const defaults = {
    materials: {
      'Sticker Paper':  { base: { rate: 80,  unit: 'sheet', range: [80,80]   }, preCut: { rate: 150, unit: 'sheet', range: [125,150] } },
      'Vinyl':          { base: { rate: 80,  unit: 'sqft',  range: [80,80]   }, preCut: { rate: 150, unit: 'sqft',  range: [120,150] } },
      'Clear':          { base: { rate: 80,  unit: 'sqft',  range: [80,80]   }, preCut: { rate: 150, unit: 'sqft',  range: [125,150] } },
      'Frosted':        { base: { rate: 150, unit: 'sheet', range: [150,150] }, preCut: null },
      'Reflectorized':  { base: { rate: 175, unit: 'sheet', range: [175,175] }, preCut: { rate: 225, unit: 'sqft',  range: [225,225] } },
    },
    rushFee:   150,
    designFee: 250,
  };

  // Parse a price cell like "₱80/sheet", "₱125-150/sq.ft." into {rate, unit, range}.
  // Uses the upper bound of any range so we don't underquote.
  function parseCell(raw) {
    const s = String(raw || '').trim();
    if (!s || s === '-' || s === '—') return null;
    const unit = /sq\.?\s*ft|sqft/i.test(s) ? 'sqft'
               : /sheet/i.test(s)            ? 'sheet'
               : null;
    if (!unit) return null;
    const nums = (s.match(/\d+(?:\.\d+)?/g) || []).map(Number).filter(n => n > 0);
    if (!nums.length) return null;
    const lo = Math.min.apply(null, nums);
    const hi = Math.max.apply(null, nums);
    return { rate: hi, unit: unit, range: [lo, hi] };
  }

  try {
    const ss    = getPriceDbSS_();
    const sheet = ss.getSheetByName('Sticker') || ss.getSheetByName('Stickers');
    if (!sheet) return defaults;

    const rows = sheet.getDataRange().getValues();
    if (rows.length < 2) return defaults;

    const materials = {};
    // Skip header row (row 0)
    for (let i = 1; i < rows.length; i++) {
      const name = String(rows[i][0] || '').trim();
      if (!name) continue;
      const base   = parseCell(rows[i][1]);
      const preCut = parseCell(rows[i][2]);
      if (!base && !preCut) continue;
      materials[name] = { base: base, preCut: preCut };
    }

    if (!Object.keys(materials).length) return defaults;
    return { materials: materials, rushFee: defaults.rushFee, designFee: defaults.designFee };
  } catch(e) {
    Logger.log('getStickerPricing error: ' + e);
    return defaults;
  }
}

// ══════════════════════════════════════════════════════════════════
//  SAVE STICKER ORDER
// ══════════════════════════════════════════════════════════════════
function saveStickerOrder(data) {
  const ss    = getMainSS_();
  let   sheet = ss.getSheetByName(STICKER_SHEET);
  if (!sheet) sheet = ss.insertSheet(STICKER_SHEET);

  const headers = [
    'Quote #', 'Date', 'Client Name', 'Contact', 'Email', 'Date Needed',
    'Sticker Type', 'Pre-Cut',
    'Width', 'Height', 'Unit', 'Area (sqft)', 'Quantity',
    'Rate', 'Base Amount',
    'Rush Order', 'Rush Fee', 'Design Service', 'Design Fee',
    'Total Amount',
    'Special Instructions', 'Sales Staff',
    'Payment Term', 'Status', 'Approved By', 'Tax Type', 'Tax Amount',
  ];

  const firstCell = sheet.getLastRow() > 0 ? String(sheet.getRange(1, 1).getValue()) : '';
  if (sheet.getLastRow() === 0 || firstCell.startsWith('STK-')) {
    if (firstCell.startsWith('STK-')) sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#E8151B').setFontColor('#fff')
      .setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
  }

  lockQuoteNumbering_();
  const lastRow  = sheet.getLastRow();
  const quoteNum = 'STK-' + String(lastRow).padStart(4, '0');

  const session   = data.token ? getSessionData_(data.token) : null;
  const staffName = session ? (session.username || session.name) : (data.salesStaff || '');

  const qty       = parseInt(data.quantity)   || 1;
  const ratePerSf = parseFloat(data.ratePerSqft) || 0;
  const sqft      = parseFloat(data.sqft) || 0;
  const baseAmt   = parseFloat(data.baseAmount) || (ratePerSf * sqft * qty);
  const rushFee   = parseFloat(data.rushFee)   || 0;
  const designFee = parseFloat(data.designFee) || 0;
  const totalAmt  = parseFloat(data.totalAmount) > 0 ? parseFloat(data.totalAmount) : (baseAmt + rushFee + designFee);

  sheet.appendRow([
    quoteNum,                                  // A  - Quote #
    new Date(),                                // B  - Date
    data.clientName    || '',                  // C  - Client Name
    data.contact       || '',                  // D  - Contact
    data.email         || '',                  // E  - Email
    data.dateNeeded    || '',                  // F  - Date Needed
    data.stickerType   || '',                  // G  - Sticker Type
    data.layout        || '',                  // H  - Layout
    parseFloat(data.width)  || 0,              // I  - Width
    parseFloat(data.height) || 0,              // J  - Height
    data.unit          || 'in',                // K  - Unit
    parseFloat(sqft.toFixed(4)),               // L  - Area (sqft)
    qty,                                       // M  - Quantity
    parseFloat(ratePerSf.toFixed(2)),          // N  - Rate/sqft
    parseFloat(baseAmt.toFixed(2)),            // O  - Base Amount
    data.rushOrder     || '',                  // P  - Rush Order
    parseFloat(rushFee.toFixed(2)),            // Q  - Rush Fee
    data.designService || '',                  // R  - Design Service
    parseFloat(designFee.toFixed(2)),          // S  - Design Fee
    parseFloat(totalAmt.toFixed(2)),           // T  - Total Amount
    data.notes         || '',                  // U  - Special Instructions
    staffName,                                 // V  - Sales Staff
    '',                                        // W  - Payment Term
    data.status || 'Pending',                  // X  - Status
    '',                                        // Y  - Approved By
    data.taxType       || 'non-vat',           // Z  - Tax Type
    parseFloat(data.taxAmount) || 0,           // AA - Tax Amount
  ]);

  sheet.getRange(sheet.getLastRow(), 14, 1, 7).setNumberFormat('₱#,##0.00');
  try { notifyQuoteSaved_(quoteNum, 'Sticker', data); } catch(_) {}
  return quoteNum;
}

// ══════════════════════════════════════════════════════════════════
//  GET RISOGRAPH PRICING (live from the "Risograph" tab)
// ══════════════════════════════════════════════════════════════════
function getRisographPricing() {
  const defaults = {
    paperTypes: {
      'News Print': {
        'Short': { 'Riso With Paper': { 'Front Only': 340, 'Back to Back': 580 }, 'Riso Only': { 'Front Only': 225, 'Back to Back': 450 } },
        'Long':  { 'Riso With Paper': { 'Front Only': 390, 'Back to Back': 625 }, 'Riso Only': { 'Front Only': 235, 'Back to Back': 470 } },
      },
      'BookPaper Sub 20': {
        'Short': { 'Riso With Paper': { 'Front Only': 450, 'Back to Back': 715 }, 'Riso Only': { 'Front Only': 285, 'Back to Back': 570 } },
        'Long':  { 'Riso With Paper': { 'Front Only': 500, 'Back to Back': 780 }, 'Riso Only': { 'Front Only': 300, 'Back to Back': 600 } },
      },
    },
    sortStapleFee: 200,   // per ream when sort & staple is selected
    rushFee:       150,
    designFee:     250,
    unitLabel:     'ream',
  };

  function parsePrice(raw) {
    if (raw == null || raw === '') return 0;
    if (typeof raw === 'number') return raw;
    const m = String(raw).replace(/[^\d.]/g, '');
    return parseFloat(m) || 0;
  }

  try {
    const ss    = getPriceDbSS_();
    const sheet = ss.getSheetByName('Risograph') || ss.getSheetByName('Riso');
    if (!sheet) return defaults;

    const rows = sheet.getDataRange().getValues();
    if (rows.length < 4) return defaults;

    // Detect columns. Expected layout:
    // Row 0 header A=Paper Type, B-C=Riso With Paper (Front,Back), D-E=Riso Only (Front,Back)
    // Row 1 sub-header: blank, Front only, Back to Back, Front Only, Back to back
    // Row 2+ either a paper-type group header (col A only) or a size row (col A=Short/Long, cols B-E=prices)
    const result = { paperTypes: {}, sortStapleFee: defaults.sortStapleFee, rushFee: defaults.rushFee, designFee: defaults.designFee, unitLabel: defaults.unitLabel };

    let currentPaper = null;
    for (let i = 2; i < rows.length; i++) {
      const r = rows[i];
      const label  = String(r[0] || '').trim();
      if (!label) continue;

      // Note line — "For sorting and stapling ... 200 per ream"
      if (/sort/i.test(label) || /stapl/i.test(label)) {
        const m = (label.match(/(\d+)/g) || []).map(Number).filter(n => n > 0);
        if (m.length) result.sortStapleFee = Math.max.apply(null, m);
        continue;
      }

      // If no prices in columns B-E it's a paper-type header row.
      const b = parsePrice(r[1]);
      const c = parsePrice(r[2]);
      const d2 = parsePrice(r[3]);
      const e = parsePrice(r[4]);
      const anyPrice = (b + c + d2 + e) > 0;

      if (!anyPrice) {
        currentPaper = label;
        result.paperTypes[currentPaper] = result.paperTypes[currentPaper] || {};
        continue;
      }

      // Size row (Short / Long)
      if (!currentPaper) continue;
      result.paperTypes[currentPaper][label] = {
        'Riso With Paper': { 'Front Only': b, 'Back to Back': c },
        'Riso Only':       { 'Front Only': d2, 'Back to Back': e },
      };
    }

    return Object.keys(result.paperTypes).length ? result : defaults;
  } catch(e) {
    Logger.log('getRisographPricing error: ' + e);
    return defaults;
  }
}

// ══════════════════════════════════════════════════════════════════
//  GET TOTE BAG PRICING (live from the "Tote Bag" tab, gid=862123714)
// ══════════════════════════════════════════════════════════════════
//  Sheet shape (flexible):
//    Column A : size label   (e.g. 10x12, 12"x14", 14 × 16)
//    Column B+: tier prices, one per column.
//
//  Tier-qty resolution (in order of preference):
//    1) Header row markers — e.g. "1+", "25 pcs", "50+", "100 pcs"
//    2) Embedded markers   — e.g. "25pcs ₱180" inside the cell
//    3) Default ladder     — [1, 25, 50, 100] mapped to column order
function getTotebagPricing() {
  const DEFAULT_TIER_LADDER = [1, 25, 50, 100];
  const defaults = {
    sizes: {
      '10 × 12': [{ minQty: 1, price: 190 }],
      '12 × 14': [{ minQty: 1, price: 210 }],
      '14 × 16': [{ minQty: 1, price: 240 }],
    },
    material:    'Canvas',
    printMethod: 'Sublimation',
    rushFee:     150,
    designFee:   250,
  };

  function normalizeSize(raw) {
    const s = String(raw || '').trim();
    if (!s) return '';
    const m = s.match(/(\d+(?:\.\d+)?)\s*(?:["”'’]|in|inch|inches)?\s*[x×*by]\s*(\d+(?:\.\d+)?)\s*(?:["”'’]|in|inch|inches)?/i);
    if (!m) return '';
    return m[1] + ' × ' + m[2];
  }

  // Extract a tier-qty marker from a string like "25 pcs", "50+", "100pc".
  function parseQtyMarker(raw) {
    const s = String(raw || '').trim().toLowerCase();
    if (!s) return null;
    const m = s.match(/(\d+)\s*(?:\+|pcs?|pc)/) || s.match(/^(\d+)\s*$/);
    return m ? parseInt(m[1]) : null;
  }

  // Pull a numeric price from a cell (handles "₱180", "180.00", "25pcs ₱180").
  function parsePrice(raw) {
    if (raw == null || raw === '') return 0;
    if (typeof raw === 'number') return raw > 0 ? raw : 0;
    const s = String(raw).trim();
    const m = s.match(/[P₱]\s*(\d+(?:\.\d+)?)/) || s.match(/(\d+(?:\.\d+)?)/);
    return m ? parseFloat(m[1]) || 0 : 0;
  }

  try {
    const ss    = getPriceDbSS_();
    const sheet = ss.getSheetByName('Tote Bag')
               || ss.getSheetByName('TOTE BAG')
               || ss.getSheetByName('ToteBag')
               || ss.getSheetByName('Tote bag')
               || ss.getSheetByName('Totebag')
               || ss.getSheetByName('Tote');
    if (!sheet) return defaults;

    const rows = sheet.getDataRange().getValues();
    if (rows.length < 2) return defaults;

    // 1) Find a size row (column A matches our size pattern). The row BEFORE
    //    the first size row is treated as the header (column B+ may contain
    //    qty markers like "25 pcs" / "50+").
    let firstSizeRow = -1;
    for (let i = 0; i < rows.length; i++) {
      if (normalizeSize(rows[i][0])) { firstSizeRow = i; break; }
    }
    if (firstSizeRow < 0) return defaults;

    // 2) Build a column → tier-qty map.
    const colQty = {};
    if (firstSizeRow > 0) {
      const header = rows[firstSizeRow - 1] || [];
      for (let c = 1; c < header.length; c++) {
        const q = parseQtyMarker(header[c]);
        if (q !== null) colQty[c] = q;
      }
    }

    const sizes = {};
    for (let i = firstSizeRow; i < rows.length; i++) {
      const sizeLabel = normalizeSize(rows[i][0]);
      if (!sizeLabel) continue;

      const tiers = [];
      let priceColIndex = 0;
      for (let c = 1; c < rows[i].length; c++) {
        const cell  = rows[i][c];
        const price = parsePrice(cell);
        if (price <= 0) continue;

        // Tier qty: header marker > embedded cell marker > default ladder
        let qty = colQty[c];
        if (qty == null) qty = parseQtyMarker(cell);
        if (qty == null) qty = DEFAULT_TIER_LADDER[priceColIndex] || DEFAULT_TIER_LADDER[DEFAULT_TIER_LADDER.length - 1];
        tiers.push({ minQty: qty, price: price });
        priceColIndex++;
      }
      if (!tiers.length) continue;
      tiers.sort((a, b) => a.minQty - b.minQty);
      sizes[sizeLabel] = tiers;
    }

    if (!Object.keys(sizes).length) return defaults;
    return {
      sizes:       sizes,
      material:    defaults.material,
      printMethod: defaults.printMethod,
      rushFee:     defaults.rushFee,
      designFee:   defaults.designFee,
    };
  } catch(e) {
    Logger.log('getTotebagPricing error: ' + e);
    return defaults;
  }
}

// ══════════════════════════════════════════════════════════════════
//  SAVE RISOGRAPH ORDER
// ══════════════════════════════════════════════════════════════════
function saveRisographOrder(data) {
  const ss    = getMainSS_();
  let   sheet = ss.getSheetByName(RISOGRAPH_SHEET);
  if (!sheet) sheet = ss.insertSheet(RISOGRAPH_SHEET);

  const headers = [
    'Quote #', 'Date', 'Client Name', 'Contact', 'Email', 'Date Needed',
    'Paper Type', 'Size', 'Service', 'Sides',
    'Quantity', 'Rate', 'Base Amount',
    'Sort & Staple', 'Sort & Staple Fee',
    'Rush Order', 'Rush Fee', 'Design Service', 'Design Fee',
    'Payment Term', 'Total Amount',
    'Special Instructions', 'Sales Staff',
    'Status', 'Approved By', 'Tax Type', 'Tax Amount',
  ];

  const firstCell = sheet.getLastRow() > 0 ? String(sheet.getRange(1, 1).getValue()) : '';
  if (sheet.getLastRow() === 0 || firstCell.startsWith('RG-')) {
    if (firstCell.startsWith('RG-')) sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#E8151B').setFontColor('#fff')
      .setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
  }

  lockQuoteNumbering_();
  const lastRow  = sheet.getLastRow();
  const quoteNum = 'RG-' + String(lastRow).padStart(4, '0');

  const session   = data.token ? getSessionData_(data.token) : null;
  const staffName = session ? (session.username || session.name) : (data.salesStaff || '');

  const qty        = parseInt(data.quantity)   || 1;
  const rate       = parseFloat(data.rate)     || 0;
  const baseAmt    = parseFloat(data.baseAmount) || (rate * qty);
  const ssFee      = parseFloat(data.sortStapleFee) || 0;
  const rushFee    = parseFloat(data.rushFee)   || 0;
  const designFee  = parseFloat(data.designFee) || 0;
  const totalAmt   = parseFloat(data.totalAmount) > 0 ? parseFloat(data.totalAmount) : (baseAmt + ssFee + rushFee + designFee);

  sheet.appendRow([
    quoteNum,                                   // A  - Quote #
    new Date(),                                 // B  - Date
    data.clientName    || '',                   // C  - Client Name
    data.contact       || '',                   // D  - Contact
    data.email         || '',                   // E  - Email
    data.dateNeeded    || '',                   // F  - Date Needed
    data.paperType     || '',                   // G  - Paper Type
    data.paperSize     || '',                   // H  - Size
    data.service       || '',                   // I  - Service
    data.sides         || '',                   // J  - Sides
    qty,                                        // K  - Quantity
    parseFloat(rate.toFixed(2)),                // L  - Rate
    parseFloat(baseAmt.toFixed(2)),             // M  - Base Amount
    data.sortStaple    || '',                   // N  - Sort & Staple
    parseFloat(ssFee.toFixed(2)),               // O  - Sort & Staple Fee
    data.rushOrder     || '',                   // P  - Rush Order
    parseFloat(rushFee.toFixed(2)),             // Q  - Rush Fee
    data.designService || '',                   // R  - Design Service
    parseFloat(designFee.toFixed(2)),           // S  - Design Fee
    '',                                         // T  - Payment Term
    parseFloat(totalAmt.toFixed(2)),            // U  - Total Amount
    data.notes         || '',                   // V  - Special Instructions
    staffName,                                  // W  - Sales Staff
    data.status || 'Pending',                   // X  - Status
    '',                                         // Y  - Approved By
    data.taxType       || 'non-vat',            // Z  - Tax Type
    parseFloat(data.taxAmount) || 0,            // AA - Tax Amount
  ]);

  sheet.getRange(sheet.getLastRow(), 12, 1, 10).setNumberFormat('₱#,##0.00');
  try { notifyQuoteSaved_(quoteNum, 'Risograph', data); } catch(_) {}
  return quoteNum;
}

// ══════════════════════════════════════════════════════════════════
//  SAVE TOTE BAG ORDER
// ══════════════════════════════════════════════════════════════════
function saveTotebagOrder(data) {
  const ss    = getMainSS_();
  let   sheet = ss.getSheetByName(TOTEBAG_SHEET);
  if (!sheet) sheet = ss.insertSheet(TOTEBAG_SHEET);

  const headers = [
    'Quote #', 'Date', 'Client Name', 'Contact', 'Email', 'Date Needed',
    'Size', 'Print Method', 'Material', 'Quantity',
    'Unit Price', 'Base Amount',
    'Rush Order', 'Rush Fee', 'Design Service', 'Design Fee',
    'Payment Term', 'Total Amount',
    'Special Instructions', 'Sales Staff',
    'Status', 'Approved By', 'Tax Type', 'Tax Amount',
  ];

  const firstCell = sheet.getLastRow() > 0 ? String(sheet.getRange(1, 1).getValue()) : '';
  if (sheet.getLastRow() === 0 || firstCell.startsWith('TB-')) {
    if (firstCell.startsWith('TB-')) sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#E8151B').setFontColor('#fff')
      .setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
  }

  lockQuoteNumbering_();
  const lastRow  = sheet.getLastRow();
  const quoteNum = 'TB-' + String(lastRow).padStart(4, '0');

  const session   = data.token ? getSessionData_(data.token) : null;
  const staffName = session ? (session.username || session.name) : (data.salesStaff || '');

  const qty       = parseInt(data.quantity)    || 1;
  const unitP     = parseFloat(data.unitPrice) || 0;
  const baseAmt   = parseFloat(data.baseAmount) || (unitP * qty);
  const rushFee   = parseFloat(data.rushFee)   || 0;
  const designFee = parseFloat(data.designFee) || 0;
  const totalAmt  = parseFloat(data.totalAmount) > 0 ? parseFloat(data.totalAmount) : (baseAmt + rushFee + designFee);

  sheet.appendRow([
    quoteNum,                                   // A  - Quote #
    new Date(),                                 // B  - Date
    data.clientName    || '',                   // C  - Client Name
    data.contact       || '',                   // D  - Contact
    data.email         || '',                   // E  - Email
    data.dateNeeded    || '',                   // F  - Date Needed
    data.totebagSize   || '',                   // G  - Size
    data.printMethod   || 'Sublimation',        // H  - Print Method
    data.material      || 'Canvas',             // I  - Material
    qty,                                        // J  - Quantity
    parseFloat(unitP.toFixed(2)),               // K  - Unit Price
    parseFloat(baseAmt.toFixed(2)),             // L  - Base Amount
    data.rushOrder     || '',                   // M  - Rush Order
    parseFloat(rushFee.toFixed(2)),             // N  - Rush Fee
    data.designService || '',                   // O  - Design Service
    parseFloat(designFee.toFixed(2)),           // P  - Design Fee
    '',                                         // Q  - Payment Term
    parseFloat(totalAmt.toFixed(2)),            // R  - Total Amount
    data.notes         || '',                   // S  - Special Instructions
    staffName,                                  // T  - Sales Staff
    data.status || 'Pending',                   // U  - Status
    '',                                         // V  - Approved By
    data.taxType       || 'non-vat',            // W  - Tax Type
    parseFloat(data.taxAmount) || 0,            // X  - Tax Amount
  ]);

  sheet.getRange(sheet.getLastRow(), 11, 1, 8).setNumberFormat('₱#,##0.00');
  try { notifyQuoteSaved_(quoteNum, 'Tote Bag', data); } catch(_) {}
  return quoteNum;
}

// ══════════════════════════════════════════════════════════════════
//  GET TICKET PRICING (live from the "Tickets" tab)
// ══════════════════════════════════════════════════════════════════
//  Sheet shape (very simple):
//    Column A : ticket type / size label (e.g. "Calling Card Size", "4" x 5"")
//    Column B : per-piece price (e.g. ₱4.50, ₱10, ₱9)
//  No tiers — single price per type. Header row optional.
function getTicketPricing() {
  const defaults = {
    types: {
      'Calling Card Size': { price: 4.50 },
      '4" x 5"':           { price: 10 },
      '3" x 6"':           { price: 9 },
    },
    rushFee:   150,
    designFee: 250,
  };

  function parsePrice(raw) {
    if (raw == null || raw === '') return 0;
    if (typeof raw === 'number') return raw > 0 ? raw : 0;
    const s = String(raw).trim();
    const m = s.match(/[P₱]\s*(\d+(?:\.\d+)?)/) || s.match(/(\d+(?:\.\d+)?)/);
    return m ? parseFloat(m[1]) || 0 : 0;
  }

  function isHeaderRow(label, price) {
    const s = String(label || '').trim().toLowerCase();
    // Treat rows like "Type | Price" or "Size | Price" as headers.
    if (!s) return true;
    if (price > 0) return false;
    return /^(type|size|name|item|product|ticket)$/i.test(s);
  }

  try {
    const ss    = getPriceDbSS_();
    const sheet = ss.getSheetByName('Tickets')
               || ss.getSheetByName('Ticket')
               || ss.getSheetByName('TICKETS');
    if (!sheet) return defaults;

    const rows = sheet.getDataRange().getValues();
    if (rows.length < 1) return defaults;

    const types = {};
    for (let i = 0; i < rows.length; i++) {
      const label = String(rows[i][0] || '').trim();
      const price = parsePrice(rows[i][1]);
      if (!label) continue;
      if (isHeaderRow(label, price)) continue;
      if (price <= 0) continue;
      types[label] = { price: price };
    }

    if (!Object.keys(types).length) return defaults;
    return {
      types:     types,
      rushFee:   defaults.rushFee,
      designFee: defaults.designFee,
    };
  } catch(e) {
    Logger.log('getTicketPricing error: ' + e);
    return defaults;
  }
}

// ══════════════════════════════════════════════════════════════════
//  SAVE TICKET ORDER
// ══════════════════════════════════════════════════════════════════
function saveTicketOrder(data) {
  const ss    = getMainSS_();
  let   sheet = ss.getSheetByName(TICKET_SHEET);
  if (!sheet) sheet = ss.insertSheet(TICKET_SHEET);

  const headers = [
    'Quote #', 'Date', 'Client Name', 'Contact', 'Email', 'Date Needed',
    'Ticket Type', 'Quantity',
    'Unit Price', 'Base Amount',
    'Rush Order', 'Rush Fee', 'Design Service', 'Design Fee',
    'Payment Term', 'Total Amount',
    'Special Instructions', 'Sales Staff',
    'Status', 'Approved By', 'Tax Type', 'Tax Amount',
  ];

  const firstCell = sheet.getLastRow() > 0 ? String(sheet.getRange(1, 1).getValue()) : '';
  if (sheet.getLastRow() === 0 || firstCell.startsWith('TKT-')) {
    if (firstCell.startsWith('TKT-')) sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#E8151B').setFontColor('#fff')
      .setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
  }

  lockQuoteNumbering_();
  const lastRow  = sheet.getLastRow();
  const quoteNum = 'TKT-' + String(lastRow).padStart(4, '0');

  const session   = data.token ? getSessionData_(data.token) : null;
  const staffName = session ? (session.username || session.name) : (data.salesStaff || '');

  const qty       = parseInt(data.quantity)    || 1;
  const unitP     = parseFloat(data.unitPrice) || 0;
  const baseAmt   = parseFloat(data.baseAmount) || (unitP * qty);
  const rushFee   = parseFloat(data.rushFee)   || 0;
  const designFee = parseFloat(data.designFee) || 0;
  const totalAmt  = parseFloat(data.totalAmount) > 0 ? parseFloat(data.totalAmount) : (baseAmt + rushFee + designFee);

  sheet.appendRow([
    quoteNum,                                   // A  - Quote #
    new Date(),                                 // B  - Date
    data.clientName    || '',                   // C  - Client Name
    data.contact       || '',                   // D  - Contact
    data.email         || '',                   // E  - Email
    data.dateNeeded    || '',                   // F  - Date Needed
    data.ticketType    || '',                   // G  - Ticket Type
    qty,                                        // H  - Quantity
    parseFloat(unitP.toFixed(2)),               // I  - Unit Price
    parseFloat(baseAmt.toFixed(2)),             // J  - Base Amount
    data.rushOrder     || '',                   // K  - Rush Order
    parseFloat(rushFee.toFixed(2)),             // L  - Rush Fee
    data.designService || '',                   // M  - Design Service
    parseFloat(designFee.toFixed(2)),           // N  - Design Fee
    '',                                         // O  - Payment Term
    parseFloat(totalAmt.toFixed(2)),            // P  - Total Amount
    data.notes         || '',                   // Q  - Special Instructions
    staffName,                                  // R  - Sales Staff
    data.status || 'Pending',                   // S  - Status
    '',                                         // T  - Approved By
    data.taxType       || 'non-vat',            // U  - Tax Type
    parseFloat(data.taxAmount) || 0,            // V  - Tax Amount
  ]);

  sheet.getRange(sheet.getLastRow(), 9, 1, 8).setNumberFormat('₱#,##0.00');
  try { notifyQuoteSaved_(quoteNum, 'Tickets', data); } catch(_) {}
  return quoteNum;
}

// ══════════════════════════════════════════════════════════════════
//  GET NEWSLETTER / NEWSPAPER PRICING (live from "NewsLetter/NewPaper" tab)
// ══════════════════════════════════════════════════════════════════
//  Sheet shape — two tables side by side:
//    Cols A,B : NewsLetter | Price   (12 pages → 75/copy, etc., then spec rows)
//    Cols C,D : NewsPaper  | Price   (2 pages COLOR → 50/copy, paired with a
//                                     "N pages BW" row underneath, then specs)
function getNewsprintPricing() {
  const defaults = {
    newsletter: {
      options: [
        { label: '12 pages', price: 75 },
        { label: '18 pages', price: 90 },
        { label: '20 pages', price: 110 },
      ],
      size: '11" x 17"', method: 'Offset printing (full color)',
      material: 'Glossy paper', minOrder: 200,
    },
    newspaper: {
      options: [
        { label: '2 pages COLOR + 4 pages BW', price: 50 },
        { label: '2 pages COLOR + 6 pages BW', price: 58 },
        { label: '2 pages COLOR + 8 pages BW', price: 72 },
      ],
      size: '18" x 24"', method: 'Offset printing',
      material: 'Newsprint', minOrder: 200,
    },
    rushFee: 150, designFee: 250,
  };

  function parsePrice(raw) {
    if (raw == null || raw === '') return 0;
    if (typeof raw === 'number') return raw > 0 ? raw : 0;
    const m = String(raw).match(/(\d+(?:\.\d+)?)/);
    return m ? parseFloat(m[1]) || 0 : 0;
  }
  function isPageRow(label) { return /\d+\s*pages?/i.test(String(label || '')); }
  function metaVal(label) {  // strip a leading "size/method/material" keyword, with or without a colon
    return String(label || '')
      .replace(/^\s*(size|method|material|minimum\s*order)\s*:?\s*/i, '')
      .trim();
  }
  function minOrderNum(label) {
    const m = String(label || '').match(/(\d+)/);
    return m ? parseInt(m[1]) : 0;
  }

  try {
    const ss    = getPriceDbSS_();
    const sheet = ss.getSheetByName('NewsLetter/NewPaper')
               || ss.getSheetByName('NewsLetter/NewsPaper')
               || ss.getSheetByName('Newsletter/Newspaper')
               || ss.getSheetByName('NewsLetter')
               || ss.getSheetByName('Newsprint');
    if (!sheet) return defaults;

    const rows = sheet.getDataRange().getValues();
    if (rows.length < 2) return defaults;

    const newsletter = { options: [], size: '', method: '', material: '', minOrder: defaults.newsletter.minOrder };
    const newspaper  = { options: [], size: '', method: '', material: '', minOrder: defaults.newspaper.minOrder };

    // Start at row 1 to skip the header row (NewsLetter | Price | NewsPaper | Price)
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];

      // ── NewsLetter side (cols A=0, B=1) ──
      const aLabel = String(r[0] || '').trim();
      const aPrice = parsePrice(r[1]);
      if (aLabel) {
        const lc = aLabel.toLowerCase();
        if (isPageRow(aLabel) && aPrice > 0) {
          newsletter.options.push({ label: aLabel.replace(/\bpage\b/i, 'pages'), price: aPrice });
        } else if (lc.indexOf('size') === 0)     newsletter.size = metaVal(aLabel);
        else if (lc.indexOf('method') === 0)     newsletter.method = metaVal(aLabel);
        else if (lc.indexOf('material') === 0)   newsletter.material = metaVal(aLabel);
        else if (lc.indexOf('minimum') === 0)    newsletter.minOrder = minOrderNum(aLabel) || newsletter.minOrder;
      }

      // ── NewsPaper side (cols C=2, D=3) ──
      const cLabel = String(r[2] || '').trim();
      const dPrice = parsePrice(r[3]);
      if (cLabel) {
        const lc = cLabel.toLowerCase();
        if (/colou?r/i.test(cLabel) && dPrice > 0) {
          // Pair with the next row's "N pages BW" descriptor if present.
          let label = cLabel;
          const nextC = String((rows[i + 1] || [])[2] || '').trim();
          if (/\bbw\b|black/i.test(nextC) && parsePrice((rows[i + 1] || [])[3]) === 0) {
            label = cLabel + ' + ' + nextC;
          }
          newspaper.options.push({ label: label, price: dPrice });
        } else if (lc.indexOf('size') === 0)     newspaper.size = metaVal(cLabel);
        else if (lc.indexOf('method') === 0)     newspaper.method = metaVal(cLabel);
        else if (lc.indexOf('material') === 0)   newspaper.material = metaVal(cLabel);
        else if (lc.indexOf('minimum') === 0)    newspaper.minOrder = minOrderNum(cLabel) || newspaper.minOrder;
      }
    }

    if (!newsletter.options.length) { newsletter.options = defaults.newsletter.options; }
    if (!newsletter.size)     newsletter.size     = defaults.newsletter.size;
    if (!newsletter.method)   newsletter.method   = defaults.newsletter.method;
    if (!newsletter.material) newsletter.material = defaults.newsletter.material;
    if (!newspaper.options.length)  { newspaper.options = defaults.newspaper.options; }
    if (!newspaper.size)      newspaper.size      = defaults.newspaper.size;
    if (!newspaper.method)    newspaper.method    = defaults.newspaper.method;
    if (!newspaper.material)  newspaper.material  = defaults.newspaper.material;

    return {
      newsletter: newsletter,
      newspaper:  newspaper,
      rushFee:    defaults.rushFee,
      designFee:  defaults.designFee,
    };
  } catch(e) {
    Logger.log('getNewsprintPricing error: ' + e);
    return defaults;
  }
}

// ══════════════════════════════════════════════════════════════════
//  SAVE NEWSLETTER / NEWSPAPER ORDER
// ══════════════════════════════════════════════════════════════════
function saveNewsprintOrder(data) {
  const ss    = getMainSS_();
  let   sheet = ss.getSheetByName(NEWSPRINT_SHEET);
  if (!sheet) sheet = ss.insertSheet(NEWSPRINT_SHEET);

  const headers = [
    'Quote #', 'Date', 'Client Name', 'Contact', 'Email', 'Date Needed',
    'Category', 'Option', 'Size', 'Material', 'Quantity',
    'Unit Price', 'Base Amount',
    'Rush Order', 'Rush Fee', 'Design Service', 'Design Fee',
    'Payment Term', 'Total Amount',
    'Special Instructions', 'Sales Staff',
    'Status', 'Approved By', 'Tax Type', 'Tax Amount',
  ];

  const firstCell = sheet.getLastRow() > 0 ? String(sheet.getRange(1, 1).getValue()) : '';
  if (sheet.getLastRow() === 0 || firstCell.startsWith('NL-')) {
    if (firstCell.startsWith('NL-')) sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#E8151B').setFontColor('#fff')
      .setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
  }

  lockQuoteNumbering_();
  const lastRow  = sheet.getLastRow();
  const quoteNum = 'NL-' + String(lastRow).padStart(4, '0');

  const session   = data.token ? getSessionData_(data.token) : null;
  const staffName = session ? (session.username || session.name) : (data.salesStaff || '');

  const qty       = parseInt(data.quantity)    || 1;
  const unitP     = parseFloat(data.unitPrice) || 0;
  const baseAmt   = parseFloat(data.baseAmount) || (unitP * qty);
  const rushFee   = parseFloat(data.rushFee)   || 0;
  const designFee = parseFloat(data.designFee) || 0;
  const totalAmt  = parseFloat(data.totalAmount) > 0 ? parseFloat(data.totalAmount) : (baseAmt + rushFee + designFee);

  sheet.appendRow([
    quoteNum,                                   // A  - Quote #
    new Date(),                                 // B  - Date
    data.clientName    || '',                   // C  - Client Name
    data.contact       || '',                   // D  - Contact
    data.email         || '',                   // E  - Email
    data.dateNeeded    || '',                   // F  - Date Needed
    data.category      || '',                   // G  - Category (Newsletter/Newspaper)
    data.optionLabel   || '',                   // H  - Option
    data.size          || '',                   // I  - Size
    data.material      || '',                   // J  - Material
    qty,                                        // K  - Quantity
    parseFloat(unitP.toFixed(2)),               // L  - Unit Price
    parseFloat(baseAmt.toFixed(2)),             // M  - Base Amount
    data.rushOrder     || '',                   // N  - Rush Order
    parseFloat(rushFee.toFixed(2)),             // O  - Rush Fee
    data.designService || '',                   // P  - Design Service
    parseFloat(designFee.toFixed(2)),           // Q  - Design Fee
    '',                                         // R  - Payment Term
    parseFloat(totalAmt.toFixed(2)),            // S  - Total Amount
    data.notes         || '',                   // T  - Special Instructions
    staffName,                                  // U  - Sales Staff
    data.status || 'Pending',                   // V  - Status
    '',                                         // W  - Approved By
    data.taxType       || 'non-vat',            // X  - Tax Type
    parseFloat(data.taxAmount) || 0,            // Y  - Tax Amount
  ]);

  sheet.getRange(sheet.getLastRow(), 12, 1, 8).setNumberFormat('₱#,##0.00');
  try { notifyQuoteSaved_(quoteNum, (data.category || 'Newsprint'), data); } catch(_) {}
  return quoteNum;
}

// ══════════════════════════════════════════════════════════════════
//  GET SOUVENIR PROGRAM PRICING (live from "Souvenir Program" tab)
// ══════════════════════════════════════════════════════════════════
//  Sheet shape (single column A):
//    Per page of A3            ← header (page size)
//    ₱80                       ← price per page
//    print method: laser
//    Material: foldcote / mirrorkote / C2S200
function getSouvenirPricing() {
  const defaults = {
    pricePerPage: 80,
    pageSize:     'A3',
    method:       'Laser',
    materials:    ['Foldcote', 'Mirrorkote', 'C2S200'],
    rushFee:      150,
    designFee:    250,
  };

  function parsePrice(raw) {
    if (raw == null || raw === '') return 0;
    if (typeof raw === 'number') return raw > 0 ? raw : 0;
    const m = String(raw).match(/[P₱]?\s*(\d+(?:\.\d+)?)/);
    return m ? parseFloat(m[1]) || 0 : 0;
  }
  function afterColon(s) { return String(s || '').replace(/^[^:]*:\s*/, '').trim(); }

  try {
    const ss    = getPriceDbSS_();
    const sheet = ss.getSheetByName('Souvenir Program')
               || ss.getSheetByName('Souvenir')
               || ss.getSheetByName('Souvenir Programs');
    if (!sheet) return defaults;

    const rows = sheet.getDataRange().getValues();
    if (rows.length < 1) return defaults;

    const result = {
      pricePerPage: 0, pageSize: defaults.pageSize, method: defaults.method,
      materials: [], rushFee: defaults.rushFee, designFee: defaults.designFee,
    };

    // Scan every cell on every row (label may be in col A or B depending on table layout)
    for (let i = 0; i < rows.length; i++) {
      for (let c = 0; c < rows[i].length; c++) {
        const cell = String(rows[i][c] || '').trim();
        if (!cell) continue;
        const lc = cell.toLowerCase();

        if (/per\s*page/i.test(cell)) {
          const m = cell.match(/of\s+(A\d|[\d.]+\s*["x×].*)/i);
          if (m) result.pageSize = m[1].trim();
          // a price may sit in the next column on the same row
          const inlinePrice = parsePrice(rows[i][c + 1]);
          if (inlinePrice > 0 && !result.pricePerPage) result.pricePerPage = inlinePrice;
        } else if (lc.indexOf('method') === 0 || /print\s*method/i.test(cell)) {
          result.method = afterColon(cell) || result.method;
        } else if (lc.indexOf('material') === 0) {
          const list = afterColon(cell).split(/[\/,]/).map(s => s.trim()).filter(Boolean);
          if (list.length) result.materials = list;
        } else if (/^[P₱]?\s*\d+(\.\d+)?$/.test(cell) && !result.pricePerPage) {
          // a lone price cell (e.g. "₱80")
          result.pricePerPage = parsePrice(cell);
        }
      }
    }

    if (!result.pricePerPage)        result.pricePerPage = defaults.pricePerPage;
    if (!result.materials.length)    result.materials    = defaults.materials;
    return result;
  } catch(e) {
    Logger.log('getSouvenirPricing error: ' + e);
    return defaults;
  }
}

// ══════════════════════════════════════════════════════════════════
//  SAVE SOUVENIR PROGRAM ORDER
// ══════════════════════════════════════════════════════════════════
function saveSouvenirOrder(data) {
  const ss    = getMainSS_();
  let   sheet = ss.getSheetByName(SOUVENIR_SHEET);
  if (!sheet) sheet = ss.insertSheet(SOUVENIR_SHEET);

  const headers = [
    'Quote #', 'Date', 'Client Name', 'Contact', 'Email', 'Date Needed',
    'Material', 'Page Size', 'Method', 'Pages/Copy', 'Quantity',
    'Price/Page', 'Base Amount',
    'Rush Order', 'Rush Fee', 'Design Service', 'Design Fee',
    'Payment Term', 'Total Amount',
    'Special Instructions', 'Sales Staff',
    'Status', 'Approved By', 'Tax Type', 'Tax Amount',
  ];

  const firstCell = sheet.getLastRow() > 0 ? String(sheet.getRange(1, 1).getValue()) : '';
  if (sheet.getLastRow() === 0 || firstCell.startsWith('SP-')) {
    if (firstCell.startsWith('SP-')) sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#E8151B').setFontColor('#fff')
      .setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
  }

  lockQuoteNumbering_();
  const lastRow  = sheet.getLastRow();
  const quoteNum = 'SP-' + String(lastRow).padStart(4, '0');

  const session   = data.token ? getSessionData_(data.token) : null;
  const staffName = session ? (session.username || session.name) : (data.salesStaff || '');

  const pages     = parseInt(data.pages)       || 1;
  const qty       = parseInt(data.quantity)    || 1;
  const unitP     = parseFloat(data.pricePerPage) || 0;
  const baseAmt   = parseFloat(data.baseAmount) || (unitP * pages * qty);
  const rushFee   = parseFloat(data.rushFee)   || 0;
  const designFee = parseFloat(data.designFee) || 0;
  const totalAmt  = parseFloat(data.totalAmount) > 0 ? parseFloat(data.totalAmount) : (baseAmt + rushFee + designFee);

  sheet.appendRow([
    quoteNum,                                   // A  - Quote #
    new Date(),                                 // B  - Date
    data.clientName    || '',                   // C  - Client Name
    data.contact       || '',                   // D  - Contact
    data.email         || '',                   // E  - Email
    data.dateNeeded    || '',                   // F  - Date Needed
    data.material      || '',                   // G  - Material
    data.pageSize      || 'A3',                 // H  - Page Size
    data.method        || 'Laser',             // I  - Method
    pages,                                      // J  - Pages/Copy
    qty,                                        // K  - Quantity
    parseFloat(unitP.toFixed(2)),               // L  - Price/Page
    parseFloat(baseAmt.toFixed(2)),             // M  - Base Amount
    data.rushOrder     || '',                   // N  - Rush Order
    parseFloat(rushFee.toFixed(2)),             // O  - Rush Fee
    data.designService || '',                   // P  - Design Service
    parseFloat(designFee.toFixed(2)),           // Q  - Design Fee
    '',                                         // R  - Payment Term
    parseFloat(totalAmt.toFixed(2)),            // S  - Total Amount
    data.notes         || '',                   // T  - Special Instructions
    staffName,                                  // U  - Sales Staff
    data.status || 'Pending',                   // V  - Status
    '',                                         // W  - Approved By
    data.taxType       || 'non-vat',            // X  - Tax Type
    parseFloat(data.taxAmount) || 0,            // Y  - Tax Amount
  ]);

  sheet.getRange(sheet.getLastRow(), 12, 1, 8).setNumberFormat('₱#,##0.00');
  try { notifyQuoteSaved_(quoteNum, 'Souvenir Program', data); } catch(_) {}
  return quoteNum;
}

// ══════════════════════════════════════════════════════════════════
//  GET ACRYLIC KEYCHAIN PRICING (live from "Acrylic Keychains" tab)
// ══════════════════════════════════════════════════════════════════
//  Sheet shape:
//    Rate: ₱25.00 per sq. in.   ·   Min combined: 80 sq.in.
//    Header: Size | Sq. In. | Min. Order | Standard / pc | Die-cut / pc
//    Rows:   2×2" ★ | 4 | 20 pcs | ₱100 | ₱120  ...
//    Rush | 250   ·   Rush more than 250 | 5% ...   ·   Design Fee | 250
function getKeychainPricing() {
  const defaults = {
    ratePerSqIn: 25,
    material: 'Acrylic + Vinyl Sticker',
    minCombinedSqIn: 80,
    diecutMultiplier: 1.2,
    rushFlat: 250,
    rushPct: 0.05,
    designFee: 250,
    sizes: [
      { label: '2×2"', sqin: 4,  minOrder: 20, standard: 100, diecut: 120 },
      { label: '2×3"', sqin: 6,  minOrder: 14, standard: 150, diecut: 180 },
      { label: '3×3"', sqin: 9,  minOrder: 9,  standard: 225, diecut: 270 },
      { label: '2×4"', sqin: 8,  minOrder: 10, standard: 200, diecut: 240 },
      { label: '3×4"', sqin: 12, minOrder: 7,  standard: 300, diecut: 360 },
      { label: '4×4"', sqin: 16, minOrder: 5,  standard: 400, diecut: 480 },
      { label: '2×5"', sqin: 10, minOrder: 8,  standard: 250, diecut: 300 },
      { label: '3×5"', sqin: 15, minOrder: 6,  standard: 375, diecut: 450 },
      { label: '4×5"', sqin: 20, minOrder: 4,  standard: 500, diecut: 600 },
    ],
  };

  function num(raw) {
    if (raw == null || raw === '') return 0;
    if (typeof raw === 'number') return raw;
    const m = String(raw).replace(/,/g, '').match(/(\d+(?:\.\d+)?)/);
    return m ? parseFloat(m[1]) || 0 : 0;
  }
  function isSizeLabel(s) { return /\d+\s*[x×]\s*\d+/i.test(String(s || '')); }
  function cleanSize(s) {
    return String(s || '').replace(/[★*]/g, '').replace(/\s+/g, '').replace(/x/i, '×').trim();
  }

  try {
    const ss    = getPriceDbSS_();
    const sheet = ss.getSheetByName('Acrylic Keychains')
               || ss.getSheetByName('Acrylic Keychain')
               || ss.getSheetByName('Keychain')
               || ss.getSheetByName('Keychains');
    if (!sheet) return defaults;

    const rows = sheet.getDataRange().getValues();
    if (rows.length < 3) return defaults;

    const result = {
      ratePerSqIn: defaults.ratePerSqIn, material: defaults.material,
      minCombinedSqIn: defaults.minCombinedSqIn, diecutMultiplier: defaults.diecutMultiplier,
      rushFlat: defaults.rushFlat, rushPct: defaults.rushPct, designFee: defaults.designFee,
      sizes: [],
    };

    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const joined = r.map(c => String(c || '')).join(' ');

      // Rate per sq.in.
      if (/per\s*sq/i.test(joined)) {
        const m = joined.match(/[P₱]?\s*(\d+(?:\.\d+)?)\s*per\s*sq/i) || joined.match(/rate[^0-9]*(\d+(?:\.\d+)?)/i);
        if (m) result.ratePerSqIn = parseFloat(m[1]) || result.ratePerSqIn;
      }
      // Min combined sq.in.
      if (/min[^.]*combined|combined\s*order/i.test(joined)) {
        const m = joined.match(/(\d+)\s*sq/i);
        if (m) result.minCombinedSqIn = parseInt(m[1]) || result.minCombinedSqIn;
      }

      const a = String(r[0] || '').trim();
      const lcA = a.toLowerCase();

      // Size data row
      if (isSizeLabel(a)) {
        const sqin     = num(r[1]);
        const minOrder = parseInt(num(r[2])) || (sqin ? Math.ceil(result.minCombinedSqIn / sqin) : 1);
        const standard = num(r[3]) || (sqin * result.ratePerSqIn);
        const diecut   = num(r[4]) || Math.round(standard * result.diecutMultiplier / 5) * 5;
        if (standard > 0) {
          result.sizes.push({ label: cleanSize(a), sqin: sqin, minOrder: minOrder, standard: standard, diecut: diecut });
        }
        continue;
      }

      // Fee rows (Rush / Design Fee)
      if (/^rush\b/i.test(lcA) && !/more than/i.test(lcA)) {
        const f = num(r[1]); if (f > 0) result.rushFlat = f;
      } else if (/rush more than|whichever/i.test(joined)) {
        const pm = joined.match(/(\d+(?:\.\d+)?)\s*%/);
        if (pm) result.rushPct = (parseFloat(pm[1]) || 5) / 100;
      } else if (/design\s*fee/i.test(lcA)) {
        const f = num(r[1]); if (f > 0) result.designFee = f;
      }
    }

    if (!result.sizes.length) return defaults;
    return result;
  } catch(e) {
    Logger.log('getKeychainPricing error: ' + e);
    return defaults;
  }
}

// ══════════════════════════════════════════════════════════════════
//  SAVE ACRYLIC KEYCHAIN ORDER
// ══════════════════════════════════════════════════════════════════
function saveKeychainOrder(data) {
  const ss    = getMainSS_();
  let   sheet = ss.getSheetByName(KEYCHAIN_SHEET);
  if (!sheet) sheet = ss.insertSheet(KEYCHAIN_SHEET);

  const headers = [
    'Quote #', 'Date', 'Client Name', 'Contact', 'Email', 'Date Needed',
    'Size', 'Sq. In.', 'Cut Type', 'Material', 'Quantity',
    'Unit Price', 'Base Amount',
    'Rush Order', 'Rush Fee', 'Design Service', 'Design Fee',
    'Design Reference', 'Payment Term', 'Total Amount',
    'Special Instructions', 'Sales Staff',
    'Status', 'Approved By', 'Tax Type', 'Tax Amount',
  ];

  const firstCell = sheet.getLastRow() > 0 ? String(sheet.getRange(1, 1).getValue()) : '';
  if (sheet.getLastRow() === 0 || firstCell.startsWith('KC-')) {
    if (firstCell.startsWith('KC-')) sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#E8151B').setFontColor('#fff')
      .setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
  }

  lockQuoteNumbering_();
  const lastRow  = sheet.getLastRow();
  const quoteNum = 'KC-' + String(lastRow).padStart(4, '0');

  const session   = data.token ? getSessionData_(data.token) : null;
  const staffName = session ? (session.username || session.name) : (data.salesStaff || '');

  const qty       = parseInt(data.quantity)    || 1;
  const unitP     = parseFloat(data.unitPrice) || 0;
  const baseAmt   = parseFloat(data.baseAmount) || (unitP * qty);
  const rushFee   = parseFloat(data.rushFee)   || 0;
  const designFee = parseFloat(data.designFee) || 0;
  const totalAmt  = parseFloat(data.totalAmount) > 0 ? parseFloat(data.totalAmount) : (baseAmt + rushFee + designFee);

  sheet.appendRow([
    quoteNum,                                   // A  - Quote #
    new Date(),                                 // B  - Date
    data.clientName    || '',                   // C  - Client Name
    data.contact       || '',                   // D  - Contact
    data.email         || '',                   // E  - Email
    data.dateNeeded    || '',                   // F  - Date Needed
    data.size          || '',                   // G  - Size
    parseFloat(data.sqin) || 0,                 // H  - Sq. In.
    data.cutType       || 'Standard',          // I  - Cut Type
    data.material      || 'Acrylic + Vinyl Sticker', // J - Material
    qty,                                        // K  - Quantity
    parseFloat(unitP.toFixed(2)),               // L  - Unit Price
    parseFloat(baseAmt.toFixed(2)),             // M  - Base Amount
    data.rushOrder     || '',                   // N  - Rush Order
    parseFloat(rushFee.toFixed(2)),             // O  - Rush Fee
    data.designService || '',                   // P  - Design Service
    parseFloat(designFee.toFixed(2)),           // Q  - Design Fee
    data.designRef     || '',                   // R  - Design Reference
    '',                                         // S  - Payment Term
    parseFloat(totalAmt.toFixed(2)),            // T  - Total Amount
    data.notes         || '',                   // U  - Special Instructions
    staffName,                                  // V  - Sales Staff
    data.status || 'Pending',                   // W  - Status
    '',                                         // X  - Approved By
    data.taxType       || 'non-vat',            // Y  - Tax Type
    parseFloat(data.taxAmount) || 0,            // Z  - Tax Amount
  ]);

  sheet.getRange(sheet.getLastRow(), 12, 1, 9).setNumberFormat('₱#,##0.00');
  try { notifyQuoteSaved_(quoteNum, 'Acrylic Keychain', data); } catch(_) {}
  return quoteNum;
}

// ══════════════════════════════════════════════════════════════════
//  GET ACRYLIC SIGNAGE PRICING (live from "Acrylic Signage and Plate
//  Number" tab — Plate Number rows are intentionally ignored)
// ══════════════════════════════════════════════════════════════════
//  Columns: Signage Type | Price | Minimum Size | Minimum Order | Remarks
//  e.g.  Acrylic Signage | 600 | 1x1 ft | 1 | if smaller than 1x1 ft …
//  Price is treated as ₱ per sq.ft, with the Minimum Size as the billing floor.
function getAcrylicSignPricing() {
  const defaults = {
    types: [
      { name: 'Acrylic Signage', pricePerSqft: 600, minSize: '1x1 ft', minSqft: 1, minOrder: 1,
        remarks: 'If smaller than 1×1 ft, charged the same as 1×1. Customized shape / 1 layer.' },
    ],
    unit: 'ft',
    rushFee: 150,
    designFee: 250,
  };

  function num(raw) {
    if (raw == null || raw === '') return 0;
    if (typeof raw === 'number') return raw;
    const m = String(raw).replace(/,/g, '').match(/(\d+(?:\.\d+)?)/);
    return m ? parseFloat(m[1]) || 0 : 0;
  }
  function minSqftFrom(sizeStr) {
    const m = String(sizeStr || '').match(/(\d+(?:\.\d+)?)\s*[x×]\s*(\d+(?:\.\d+)?)/i);
    if (m) return (parseFloat(m[1]) || 1) * (parseFloat(m[2]) || 1);
    const n = num(sizeStr);
    return n > 0 ? n : 1;
  }

  try {
    const ss    = getPriceDbSS_();
    const sheet = ss.getSheetByName('Acrylic Signage and Plate Number')
               || ss.getSheetByName('Acrylic Signage')
               || ss.getSheetByName('Acrylic Signage and Plate No');
    if (!sheet) return defaults;

    const rows = sheet.getDataRange().getValues();
    if (rows.length < 2) return defaults;

    // Locate the header row (has "Signage Type" + "Price")
    let headerIdx = -1;
    for (let i = 0; i < rows.length; i++) {
      const joined = rows[i].map(c => String(c || '').toLowerCase()).join(' ');
      if (/signage\s*type/.test(joined) && /price/.test(joined)) { headerIdx = i; break; }
    }
    const start = headerIdx >= 0 ? headerIdx + 1 : 1;

    const types = [];
    for (let i = start; i < rows.length; i++) {
      const name = String(rows[i][0] || '').trim();
      if (!name) continue;
      // Stop once the Plate Number section begins — we don't touch plate no.
      if (/plate/i.test(name)) break;
      const price = num(rows[i][1]);
      if (price <= 0) continue;
      const minSize = String(rows[i][2] || '').trim() || '1x1 ft';
      types.push({
        name:         name,
        pricePerSqft: price,
        minSize:      minSize,
        minSqft:      minSqftFrom(minSize),
        minOrder:     parseInt(num(rows[i][3])) || 1,
        remarks:      String(rows[i][4] || '').trim(),
      });
    }

    if (!types.length) return defaults;
    return { types: types, unit: 'ft', rushFee: defaults.rushFee, designFee: defaults.designFee };
  } catch(e) {
    Logger.log('getAcrylicSignPricing error: ' + e);
    return defaults;
  }
}

// ══════════════════════════════════════════════════════════════════
//  SAVE ACRYLIC SIGNAGE ORDER
// ══════════════════════════════════════════════════════════════════
function saveAcrylicSignOrder(data) {
  const ss    = getMainSS_();
  let   sheet = ss.getSheetByName(ACRYLICSIGN_SHEET);
  if (!sheet) sheet = ss.insertSheet(ACRYLICSIGN_SHEET);

  const headers = [
    'Quote #', 'Date', 'Client Name', 'Contact', 'Email', 'Date Needed',
    'Signage Type', 'Width', 'Height', 'Unit', 'Area (sqft)', 'Billed Sqft', 'Quantity',
    'Rate/sqft', 'Unit Price', 'Base Amount',
    'Rush Order', 'Rush Fee', 'Design Service', 'Design Fee',
    'Payment Term', 'Total Amount',
    'Special Instructions', 'Sales Staff',
    'Status', 'Approved By', 'Tax Type', 'Tax Amount',
  ];

  const firstCell = sheet.getLastRow() > 0 ? String(sheet.getRange(1, 1).getValue()) : '';
  if (sheet.getLastRow() === 0 || firstCell.startsWith('AS-')) {
    if (firstCell.startsWith('AS-')) sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#E8151B').setFontColor('#fff')
      .setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
  }

  lockQuoteNumbering_();
  const lastRow  = sheet.getLastRow();
  const quoteNum = 'AS-' + String(lastRow).padStart(4, '0');

  const session   = data.token ? getSessionData_(data.token) : null;
  const staffName = session ? (session.username || session.name) : (data.salesStaff || '');

  const qty       = parseInt(data.quantity)    || 1;
  const rate      = parseFloat(data.ratePerSqft) || 0;
  const unitP     = parseFloat(data.unitPrice) || 0;
  const baseAmt   = parseFloat(data.baseAmount) || (unitP * qty);
  const rushFee   = parseFloat(data.rushFee)   || 0;
  const designFee = parseFloat(data.designFee) || 0;
  const totalAmt  = parseFloat(data.totalAmount) > 0 ? parseFloat(data.totalAmount) : (baseAmt + rushFee + designFee);

  sheet.appendRow([
    quoteNum,                                   // A  - Quote #
    new Date(),                                 // B  - Date
    data.clientName    || '',                   // C  - Client Name
    data.contact       || '',                   // D  - Contact
    data.email         || '',                   // E  - Email
    data.dateNeeded    || '',                   // F  - Date Needed
    data.signageType   || '',                   // G  - Signage Type
    parseFloat(data.width)  || 0,               // H  - Width (entered)
    parseFloat(data.height) || 0,               // I  - Height (entered)
    data.unit          || 'ft',                 // J  - Unit
    parseFloat(data.sqft)   || 0,               // K  - Area (sqft)
    parseFloat(data.billedSqft) || 0,           // L  - Billed Sqft
    qty,                                        // M  - Quantity
    parseFloat(rate.toFixed(2)),                // N  - Rate/sqft
    parseFloat(unitP.toFixed(2)),               // O  - Unit Price
    parseFloat(baseAmt.toFixed(2)),             // P  - Base Amount
    data.rushOrder     || '',                   // Q  - Rush Order
    parseFloat(rushFee.toFixed(2)),             // R  - Rush Fee
    data.designService || '',                   // S  - Design Service
    parseFloat(designFee.toFixed(2)),           // T  - Design Fee
    '',                                         // U  - Payment Term
    parseFloat(totalAmt.toFixed(2)),            // V  - Total Amount
    data.notes         || '',                   // W  - Special Instructions
    staffName,                                  // X  - Sales Staff
    data.status || 'Pending',                   // Y  - Status
    '',                                         // Z  - Approved By
    data.taxType       || 'non-vat',            // AA - Tax Type
    parseFloat(data.taxAmount) || 0,            // AB - Tax Amount
  ]);

  sheet.getRange(sheet.getLastRow(), 14, 1, 9).setNumberFormat('₱#,##0.00');
  try { notifyQuoteSaved_(quoteNum, 'Acrylic Signage', data); } catch(_) {}
  return quoteNum;
}

// ══════════════════════════════════════════════════════════════════
//  GET ACRYLIC PLATE NUMBER PRICING (live from "Acrylic Signage and
//  Plate Number" tab — only the Plate Number rows)
// ══════════════════════════════════════════════════════════════════
//  Rows like:  Acrylic Plate Number - Motorcycle | 200
//              Acrylic Plate Number - Car        | 400
//  Flat price per piece. Rush = ₱250 or 5% whichever higher; Design ₱250.
function getAcrylicPlatePricing() {
  const defaults = {
    types: [
      { name: 'Motorcycle', fullName: 'Acrylic Plate Number - Motorcycle', price: 200 },
      { name: 'Car',        fullName: 'Acrylic Plate Number - Car',        price: 400 },
    ],
    rushFlat: 250, rushPct: 0.05, designFee: 250,
  };

  function num(raw) {
    if (raw == null || raw === '') return 0;
    if (typeof raw === 'number') return raw;
    const m = String(raw).replace(/,/g, '').match(/(\d+(?:\.\d+)?)/);
    return m ? parseFloat(m[1]) || 0 : 0;
  }
  function shortName(full) {
    // "Acrylic Plate Number - Motorcycle" -> "Motorcycle"
    const m = String(full || '').match(/plate\s*number\s*[-–:]?\s*(.+)$/i);
    return m ? m[1].trim() : String(full || '').trim();
  }

  try {
    const ss    = getPriceDbSS_();
    const sheet = ss.getSheetByName('Acrylic Signage and Plate Number')
               || ss.getSheetByName('Acrylic Plate Number')
               || ss.getSheetByName('Acrylic Signage');
    if (!sheet) return defaults;

    const rows = sheet.getDataRange().getValues();
    if (rows.length < 2) return defaults;

    let headerIdx = -1;
    for (let i = 0; i < rows.length; i++) {
      const joined = rows[i].map(c => String(c || '').toLowerCase()).join(' ');
      if (/signage\s*type|plate/.test(joined) && /price/.test(joined)) { headerIdx = i; break; }
    }
    const start = headerIdx >= 0 ? headerIdx + 1 : 1;

    const result = { types: [], rushFlat: defaults.rushFlat, rushPct: defaults.rushPct, designFee: defaults.designFee };
    for (let i = start; i < rows.length; i++) {
      const name = String(rows[i][0] || '').trim();
      if (!name) continue;
      const lc = name.toLowerCase();

      if (/plate/i.test(name)) {
        const price = num(rows[i][1]);
        if (price > 0) result.types.push({ name: shortName(name), fullName: name, price: price });
      } else if (/^rush\b/i.test(lc) && !/more than/i.test(lc)) {
        const f = num(rows[i][1]); if (f > 0) result.rushFlat = f;
      } else if (/rush more than|whichever/i.test(rows[i].map(c => String(c||'')).join(' '))) {
        const pm = rows[i].map(c => String(c||'')).join(' ').match(/(\d+(?:\.\d+)?)\s*%/);
        if (pm) result.rushPct = (parseFloat(pm[1]) || 5) / 100;
      } else if (/design\s*fee/i.test(lc)) {
        const f = num(rows[i][1]); if (f > 0) result.designFee = f;
      }
    }

    if (!result.types.length) return defaults;
    return result;
  } catch(e) {
    Logger.log('getAcrylicPlatePricing error: ' + e);
    return defaults;
  }
}

// ══════════════════════════════════════════════════════════════════
//  SAVE ACRYLIC PLATE NUMBER ORDER
// ══════════════════════════════════════════════════════════════════
function saveAcrylicPlateOrder(data) {
  const ss    = getMainSS_();
  let   sheet = ss.getSheetByName(ACRYLICPLATE_SHEET);
  if (!sheet) sheet = ss.insertSheet(ACRYLICPLATE_SHEET);

  const headers = [
    'Quote #', 'Date', 'Client Name', 'Contact', 'Email', 'Date Needed',
    'Plate Type', 'Plate Text / Name', 'Quantity',
    'Unit Price', 'Base Amount',
    'Rush Order', 'Rush Fee', 'Design Service', 'Design Fee',
    'Payment Term', 'Total Amount',
    'Special Instructions', 'Sales Staff',
    'Status', 'Approved By', 'Tax Type', 'Tax Amount', 'Items JSON',
  ];

  const firstCell = sheet.getLastRow() > 0 ? String(sheet.getRange(1, 1).getValue()) : '';
  if (sheet.getLastRow() === 0 || firstCell.startsWith('AP-')) {
    if (firstCell.startsWith('AP-')) sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#E8151B').setFontColor('#fff')
      .setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
  }

  lockQuoteNumbering_();
  const lastRow  = sheet.getLastRow();
  const quoteNum = 'AP-' + String(lastRow).padStart(4, '0');

  const session   = data.token ? getSessionData_(data.token) : null;
  const staffName = session ? (session.username || session.name) : (data.salesStaff || '');

  const qty       = parseInt(data.quantity)    || 1;
  const unitP     = parseFloat(data.unitPrice) || 0;
  const baseAmt   = parseFloat(data.baseAmount) || (unitP * qty);
  const rushFee   = parseFloat(data.rushFee)   || 0;
  const designFee = parseFloat(data.designFee) || 0;
  const totalAmt  = parseFloat(data.totalAmount) > 0 ? parseFloat(data.totalAmount) : (baseAmt + rushFee + designFee);

  sheet.appendRow([
    quoteNum,                                   // A  - Quote #
    new Date(),                                 // B  - Date
    data.clientName    || '',                   // C  - Client Name
    data.contact       || '',                   // D  - Contact
    data.email         || '',                   // E  - Email
    data.dateNeeded    || '',                   // F  - Date Needed
    data.plateType     || '',                   // G  - Plate Type
    data.plateText     || '',                   // H  - Plate Text / Name
    qty,                                        // I  - Quantity
    parseFloat(unitP.toFixed(2)),               // J  - Unit Price
    parseFloat(baseAmt.toFixed(2)),             // K  - Base Amount
    data.rushOrder     || '',                   // L  - Rush Order
    parseFloat(rushFee.toFixed(2)),             // M  - Rush Fee
    data.designService || '',                   // N  - Design Service
    parseFloat(designFee.toFixed(2)),           // O  - Design Fee
    '',                                         // P  - Payment Term
    parseFloat(totalAmt.toFixed(2)),            // Q  - Total Amount
    data.notes         || '',                   // R  - Special Instructions
    staffName,                                  // S  - Sales Staff
    data.status || 'Pending',                   // T  - Status
    '',                                         // U  - Approved By
    data.taxType       || 'non-vat',            // V  - Tax Type
    parseFloat(data.taxAmount) || 0,            // W  - Tax Amount
    (data.items && data.items.length) ? JSON.stringify(data.items) : '[]',  // X - Items JSON
  ]);

  sheet.getRange(sheet.getLastRow(), 10, 1, 8).setNumberFormat('₱#,##0.00');
  try { notifyQuoteSaved_(quoteNum, 'Acrylic Plate Number', data); } catch(_) {}
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
        notes:        String(row[30] || ''),          // AE = col 31
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

  lockQuoteNumbering_();
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
  try { notifyQuoteSaved_(quoteNum, 'Signage', data); } catch(_) {}
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

  const session   = payload.salesStaff ? getSessionData_(payload.salesStaff) : null;
  const staffName = session ? session.name : '';
  lockQuoteNumbering_();
  const lastRow   = sheet.getLastRow();
  const orderNum  = 'RQ-' + String(lastRow).padStart(4, '0');

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
  try {
    notifyQuoteSaved_(orderNum, 'Receipt', {
      clientName:  (payload.company || ((payload.firstName || '') + ' ' + (payload.lastName || '')).trim() || '—'),
      contact:     payload.mobile     || '',
      company:     payload.company    || '',
      firstName:   payload.firstName  || '',
      lastName:    payload.lastName   || '',
      email:       payload.email      || '',
      dateNeeded:  payload.dateNeeded  || '',
      notes:       payload.notes       || '',
      totalAmount: parseFloat(payload.totalPrice) || 0,
    });
  } catch(_) {}
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
        sellingPrice: row[11],                 // col L = Selling Price/Booklet
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
    const extSS    = getPriceDbSS_();
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
  else if (quoteNum.startsWith('MUG-')) sheetName = MUG_SHEET;
  else if (quoteNum.startsWith('STK-')) sheetName = STICKER_SHEET;
  else if (quoteNum.startsWith('RG-'))  sheetName = RISOGRAPH_SHEET;
  else if (quoteNum.startsWith('TB-'))  sheetName = TOTEBAG_SHEET;
  else if (quoteNum.startsWith('TKT-')) sheetName = TICKET_SHEET;
  else if (quoteNum.startsWith('NL-'))  sheetName = NEWSPRINT_SHEET;
  else if (quoteNum.startsWith('SP-'))  sheetName = SOUVENIR_SHEET;
  else if (quoteNum.startsWith('KC-'))  sheetName = KEYCHAIN_SHEET;
  else if (quoteNum.startsWith('AS-'))  sheetName = ACRYLICSIGN_SHEET;
  else if (quoteNum.startsWith('AP-'))  sheetName = ACRYLICPLATE_SHEET;
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
else if (quoteNum.startsWith('MUG-')) ptCol = 24; // col X
else if (quoteNum.startsWith('STK-')) ptCol = 23; // col W
else if (quoteNum.startsWith('RG-'))  ptCol = 20; // col T
else if (quoteNum.startsWith('TB-'))  ptCol = 17; // col Q
else if (quoteNum.startsWith('TKT-')) ptCol = 15; // col O
else if (quoteNum.startsWith('NL-'))  ptCol = 18; // col R
else if (quoteNum.startsWith('SP-'))  ptCol = 18; // col R
else if (quoteNum.startsWith('KC-'))  ptCol = 19; // col S
else if (quoteNum.startsWith('AS-'))  ptCol = 21; // col U
else if (quoteNum.startsWith('AP-'))  ptCol = 16; // col P

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
    const ss    = getPriceDbSS_();
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
    'Printed Materials Ready', 'Printing Type', 'Printing Fee', 'Rush Order',
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

  lockQuoteNumbering_();
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
  try { notifyQuoteSaved_(quoteNum, 'Bookbinding', data); } catch(_) {}
  return quoteNum;
}
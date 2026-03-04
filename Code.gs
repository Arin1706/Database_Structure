/* ═══════════════════════════════════════════════════════
   DEALPAY — CRM Contact Center
   Google Apps Script Backend (Sheet-Driven)
   ═══════════════════════════════════════════════════════ */

/* ─────────────────────────────────────────────────────
   WEB APP ENTRY
   ───────────────────────────────────────────────────── */
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('DEALPAY — CRM Contact Center')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ─────────────────────────────────────────────────────
   SHEET CONFIG
   ───────────────────────────────────────────────────── */
var SHEETS = {
  Locations:           { color: '#00B4D8', headers: ['Location ID','Location Name','Address','Province','Zone','Contact Person','Contact Phone','Open Date','Status','Rent Cost/Month','Notes'] },
  Projects:            { color: '#00C896', headers: ['Project ID','Location ID','Project Name','Cabinet Type','Total Value (฿)','Collected (฿)','Progress %','Status','Creation Date','Sales Rep','Notes'] },
  Investors:           { color: '#FB923C', headers: ['Customer ID','Location ID','Full Name (TH)','Full Name (EN)','Phone','Email','Line ID','Facebook','Address','ID Card No. (Last 4)','ID Card Photo','Bank Book Photo','Amount Paid (฿)','Risk Score','Debt (฿)','Last Contact','Status','Join Date','Notes'] },
  Transactions:        { color: '#FFC300', headers: ['Transaction ID','Project ID','Customer ID','Location ID','Amount (฿)','Transaction Date','Payment Method','Reference No.','Status','Receipt URL','Notes'] },
  Cases:               { color: '#FF4D4D', headers: ['Case ID','Location ID','Customer ID','Project ID','Case Type','Priority','Status','Sub-Status','Title','Description','Assigned To','Assigned Team','Created Date','Due Date','Resolution Date','Resolution Notes','Estimated Cost (฿)','Actual Cost (฿)','Created By'] },
  CoordCases:          { color: '#A78BFA', headers: ['Work Order ID','Location ID','Location Name','Investor ID','Investor Name','Team','Priority','Title','Description','Photos','Status','Created Date','Created By','Resolved Date','Resolved By','Resolution Notes'] },
  Communications:      { color: '#64748B', headers: ['Comm ID','Case ID','Customer ID','Location ID','Comm Type','Direction','Subject','Message','Contact Result','Date','Sent By','Status','Template Used','Auto-Generated','Attachments'] },
  Invoices:            { color: '#F472B6', headers: ['Invoice ID','Customer ID','Project ID','Location ID','Invoice Date','Due Date','Amount (฿)','Payment Received (฿)','Outstanding (฿)','Status','Payment Terms','Sent Date','Reminder Sent','Reminder Date','Items Description','Notes'] },
  DebtCollection:      { color: '#EF4444', headers: ['Collection ID','Customer ID','Invoice ID','Location ID','Total Debt (฿)','Days Overdue','Collection Status','Escalation Level','Assigned To','Contact Attempts','Last Contact Date','Next Follow-up','Notes'] },
  SatisfactionSurvey:  { color: '#22D3EE', headers: ['Survey ID','Case ID','Customer ID','Location ID','Rating (1-5)','Speed Rating','Quality Rating','Service Rating','Comment','Survey Date','Survey Channel','Follow-up Required','Investor Name','Location Name','Agent','Contact Result'] },
  EmailLog:            { color: '#94A3B8', headers: ['Email ID','Related ID','Related Type','To Email','To Name','Subject','Body Preview','Template','Sent Date','Status','Trigger Event','Auto-Generated'] },
  SystemConfig:        { color: '#475569', headers: ['Config Key','Value','Description'] }
};

/* ─────────────────────────────────────────────────────
   SETUP
   ───────────────────────────────────────────────────── */
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) { ss = SpreadsheetApp.create('DEALPAY — Database'); Logger.log('สร้างใหม่: ' + ss.getUrl()); }
  var names = Object.keys(SHEETS);
  for (var n = 0; n < names.length; n++) {
    var name = names[n], cfg = SHEETS[name];
    var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
    if (!sheet.getRange(1,1).getValue()) {
      sheet.getRange(1,1,1,cfg.headers.length).setValues([cfg.headers]).setFontWeight('bold').setBackground(cfg.color).setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }
    sheet.setTabColor(cfg.color);
  }
  var def = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (def && ss.getSheets().length > 1) { try { ss.deleteSheet(def); } catch(e){} }
  Logger.log('✅ สร้าง ' + names.length + ' Sheets เรียบร้อย');
  return ss.getUrl();
}


/* ═════════════════════════════════════════════════════
   GENERIC HELPERS
   ═════════════════════════════════════════════════════ */
function getSheetData_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  var d = sh.getDataRange().getValues(), h = d[0], out = [];
  for (var i = 1; i < d.length; i++) { var o = {}; for (var j = 0; j < h.length; j++) o[h[j]] = d[i][j]; out.push(o); }
  return out;
}

function appendRow_(name, obj) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet "'+name+'" not found');
  var h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var row = []; for (var i = 0; i < h.length; i++) row.push(obj[h[i]] !== undefined ? obj[h[i]] : '');
  sh.appendRow(row);
  return obj;
}

function updateRow_(name, id, updates) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) return null;
  var d = sh.getDataRange().getValues(), h = d[0];
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][0]) === String(id)) {
      var keys = Object.keys(updates);
      for (var k = 0; k < keys.length; k++) { var ci = h.indexOf(keys[k]); if (ci >= 0) sh.getRange(i+1, ci+1).setValue(updates[keys[k]]); }
      return id;
    }
  }
  return null;
}

function deleteRow_(name, id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) return false;
  var d = sh.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][0]) === String(id)) { sh.deleteRow(i + 1); return true; }
  }
  return false;
}

function countRows_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  return sh ? Math.max(0, sh.getLastRow() - 1) : 0;
}

function safeDate_(v)     { if (!v) return ''; if (v instanceof Date) { try { return Utilities.formatDate(v,'Asia/Bangkok','yyyy-MM-dd'); } catch(e){} } return String(v); }
function safeDateTime_(v) { if (!v) return ''; if (v instanceof Date) { try { return Utilities.formatDate(v,'Asia/Bangkok','yyyy-MM-dd HH:mm'); } catch(e){} } return String(v); }
function safeNum_(v)      { var n = Number(v); return isNaN(n) ? 0 : n; }
function padNum_(n, len)  { var s = String(n); while (s.length < len) s = '0' + s; return s; }


/* ═════════════════════════════════════════════════════
   ID GENERATORS
   ═════════════════════════════════════════════════════ */
function generateLocationId_() {
  var rows = getSheetData_('Locations'), max = 0;
  for (var i = 0; i < rows.length; i++) { var m = String(rows[i]['Location ID']).match(/LOC(\d+)/); if (m) max = Math.max(max, parseInt(m[1])); }
  return 'LOC' + padNum_(max + 1, 3);
}

function generateInvestorId_(nameEn, idLast4) {
  var initials = 'CU';
  if (nameEn && nameEn.length >= 2) {
    var parts = nameEn.toUpperCase().split(/\s+/);
    initials = parts.length > 1 ? parts[0][0] + parts[parts.length-1][0] : parts[0].substring(0,2);
  }
  return initials + (idLast4 || '0000');
}

function generateProjectId_(projectName) {
  var prefix = (projectName || 'PR').toUpperCase().replace(/[^A-Z]/g, '').substring(0, 2);
  if (prefix.length < 2) prefix = 'PR';
  var rows = getSheetData_('Projects'), max = 0;
  for (var i = 0; i < rows.length; i++) {
    var m = String(rows[i]['Project ID']).match(new RegExp('^' + prefix + '(\\d+)'));
    if (m) max = Math.max(max, parseInt(m[1]));
  }
  return prefix + padNum_(max + 1, 4);
}

function generateCaseId_() {
  var y = new Date().getFullYear(), rows = getSheetData_('Cases'), max = 0;
  for (var i = 0; i < rows.length; i++) { var m = String(rows[i]['Case ID']).match(/CS-\d{4}-(\d+)/); if (m) max = Math.max(max, parseInt(m[1])); }
  return 'CS-' + y + '-' + padNum_(max + 1, 6);
}

function generateWOId_() {
  var y = new Date().getFullYear(), rows = getSheetData_('CoordCases'), max = 0;
  for (var i = 0; i < rows.length; i++) { var m = String(rows[i]['Work Order ID']).match(/WO-\d{4}-(\d+)/); if (m) max = Math.max(max, parseInt(m[1])); }
  return 'WO-' + y + '-' + padNum_(max + 1, 6);
}

function generateSurveyId_() { return 'SAT-' + padNum_(countRows_('SatisfactionSurvey') + 1, 3); }
function generateCommId_()   { return countRows_('Communications') + 1; }
function generateInvoiceId_() {
  var y = new Date().getFullYear(), rows = getSheetData_('Invoices'), max = 0;
  for (var i = 0; i < rows.length; i++) { var m = String(rows[i]['Invoice ID']).match(/INV-\d{4}-(\d+)/); if (m) max = Math.max(max, parseInt(m[1])); }
  return 'INV-' + y + '-' + padNum_(max + 1, 3);
}
function generateTxnId_() { return 'TXN-' + padNum_(countRows_('Transactions') + 1, 3); }
function generateCollectionId_() {
  var y = new Date().getFullYear();
  return 'DC-' + y + '-' + padNum_(countRows_('DebtCollection') + 1, 3);
}

function getNewCaseId() { return generateCaseId_(); }
function getNewWOId()   { return generateWOId_(); }


/* ═════════════════════════════════════════════════════
   getAllData() — MAIN DATA LOADER
   ═════════════════════════════════════════════════════ */
function getAllData() {
  try {
  var locations  = getSheetData_('Locations');
  var projects   = getSheetData_('Projects');
  var investors  = getSheetData_('Investors');
  var cases      = getSheetData_('Cases');
  var comms      = getSheetData_('Communications');
  var coordCases = getSheetData_('CoordCases');
  var surveys    = getSheetData_('SatisfactionSurvey');
  var invoices   = getSheetData_('Invoices');
  var debtCol    = getSheetData_('DebtCollection');

  var locsOut = [];
  for (var li = 0; li < locations.length; li++) {
    var loc = locations[li], locId = String(loc['Location ID']);

    var lProj = [];
    for (var pi = 0; pi < projects.length; pi++) {
      if (String(projects[pi]['Location ID']) === locId) {
        lProj.push({ id: String(projects[pi]['Project ID']), name: String(projects[pi]['Project Name']), cabinetType: String(projects[pi]['Cabinet Type']), totalValue: safeNum_(projects[pi]['Total Value (฿)']), collected: safeNum_(projects[pi]['Collected (฿)']), status: String(projects[pi]['Status']) });
      }
    }

    var lInv = [];
    for (var ii = 0; ii < investors.length; ii++) {
      if (String(investors[ii]['Location ID']) === locId) {
        lInv.push({
          customerId: String(investors[ii]['Customer ID']),
          name: String(investors[ii]['Full Name (TH)']),
          nameEn: String(investors[ii]['Full Name (EN)']),
          phone: String(investors[ii]['Phone']),
          email: String(investors[ii]['Email']),
          line: String(investors[ii]['Line ID']),
          facebook: String(investors[ii]['Facebook'] || ''),
          address: String(investors[ii]['Address'] || ''),
          idCardLast4: String(investors[ii]['ID Card No. (Last 4)'] || ''),
          amountPaid: safeNum_(investors[ii]['Amount Paid (฿)']),
          riskScore: safeNum_(investors[ii]['Risk Score']),
          debt: safeNum_(investors[ii]['Debt (฿)']),
          lastContact: safeDate_(investors[ii]['Last Contact']),
          status: String(investors[ii]['Status'] || 'Active'),
          joinDate: safeDate_(investors[ii]['Join Date']),
          notes: String(investors[ii]['Notes'] || '')
        });
      }
    }

    var lCases = [];
    for (var ci = 0; ci < cases.length; ci++) {
      if (String(cases[ci]['Location ID']) === locId) {
        var c = cases[ci];
        lCases.push({ caseId: String(c['Case ID']), type: String(c['Case Type']), title: String(c['Title']), status: String(c['Status']), priority: String(c['Priority']), date: safeDate_(c['Created Date']), assignedTo: String(c['Assigned To']), cost: safeNum_(c['Actual Cost (฿)']) || safeNum_(c['Estimated Cost (฿)']), resolution: String(c['Resolution Notes'] || '') });
      }
    }

    var lComms = [];
    for (var mi = 0; mi < comms.length; mi++) {
      if (String(comms[mi]['Location ID']) === locId) {
        lComms.push({ id: comms[mi]['Comm ID'], type: String(comms[mi]['Comm Type']), dir: String(comms[mi]['Direction']), subject: String(comms[mi]['Subject']), date: safeDate_(comms[mi]['Date']), from: String(comms[mi]['Sent By']) });
      }
    }

    locsOut.push({
      id: locId, name: String(loc['Location Name']),
      address: String(loc['Address']||''), province: String(loc['Province']||''), zone: String(loc['Zone']||''),
      contactPerson: String(loc['Contact Person']||''), contactPhone: String(loc['Contact Phone']||''),
      openDate: safeDate_(loc['Open Date']), status: String(loc['Status']||'Active'), rentCost: safeNum_(loc['Rent Cost/Month']),
      notes: String(loc['Notes']||''),
      projects: lProj, investors: lInv, cases: lCases, comms: lComms
    });
  }

  var coordOut = [];
  for (var wi = 0; wi < coordCases.length; wi++) {
    var wc = coordCases[wi];
    coordOut.push({
      id: String(wc['Work Order ID']), locId: String(wc['Location ID']), locName: String(wc['Location Name']),
      investorId: String(wc['Investor ID']||''), investorName: String(wc['Investor Name']||''),
      team: String(wc['Team']), priority: String(wc['Priority']),
      title: String(wc['Title']), description: String(wc['Description']||''),
      photos: wc['Photos'] ? String(wc['Photos']).split(',').filter(Boolean) : [],
      status: String(wc['Status']), created: safeDateTime_(wc['Created Date']), createdBy: String(wc['Created By']||''),
      resolvedDate: safeDateTime_(wc['Resolved Date']), resolvedBy: String(wc['Resolved By']||''),
      resolutionNotes: String(wc['Resolution Notes']||''),
      emails: [], timeline: []
    });
  }

  var survOut = [];
  for (var si = 0; si < surveys.length; si++) {
    var s = surveys[si], contacted = safeNum_(s['Rating (1-5)']) > 0;
    survOut.push({
      id: String(s['Survey ID']), date: safeDate_(s['Survey Date']),
      investor: String(s['Investor Name']||''), location: String(s['Location Name']||''),
      score: contacted ? safeNum_(s['Rating (1-5)']) : null,
      serviceQuality: String(s['Service Rating']||''), machineCondition: String(s['Quality Rating']||''),
      contactResult: String(s['Contact Result']||''), contacted: contacted,
      agent: String(s['Agent']||'')
    });
  }

  var stats = {
    totalLocations: locations.length,
    totalProjects: projects.length,
    totalInvestors: investors.length,
    totalCases: cases.length,
    openCases: 0,
    totalInvoices: invoices.length,
    overdueInvoices: 0,
    totalDebt: 0
  };
  for (var sc = 0; sc < cases.length; sc++) { if (cases[sc]['Status'] !== 'Resolved' && cases[sc]['Status'] !== 'Closed') stats.openCases++; }
  for (var si2 = 0; si2 < invoices.length; si2++) { if (invoices[si2]['Status'] === 'Overdue') stats.overdueInvoices++; }
  for (var sd = 0; sd < debtCol.length; sd++) { stats.totalDebt += safeNum_(debtCol[sd]['Total Debt (฿)']); }

  return { locations: locsOut, coordCases: coordOut, surveys: survOut, stats: stats };
  } catch(e) {
    Logger.log('❌ getAllData error: ' + e.message);
    throw new Error('getAllData ล้มเหลว: ' + e.message + ' — ลองรัน setupSheets() ก่อน');
  }
}


/* ═════════════════════════════════════════════════════
   LOCATION CRUD
   ═════════════════════════════════════════════════════ */
function addLocation(data) {
  var id = generateLocationId_();
  appendRow_('Locations', {
    'Location ID': id, 'Location Name': data.name || '', 'Address': data.address || '',
    'Province': data.province || '', 'Zone': data.zone || '',
    'Contact Person': data.contactPerson || '', 'Contact Phone': data.contactPhone || '',
    'Open Date': data.openDate || new Date(), 'Status': data.status || 'Active',
    'Rent Cost/Month': data.rentCost || 0, 'Notes': data.notes || ''
  });
  return id;
}

function updateLocation(locId, data) {
  var updates = {};
  if (data.name !== undefined)          updates['Location Name']  = data.name;
  if (data.address !== undefined)       updates['Address']        = data.address;
  if (data.province !== undefined)      updates['Province']       = data.province;
  if (data.zone !== undefined)          updates['Zone']           = data.zone;
  if (data.contactPerson !== undefined) updates['Contact Person'] = data.contactPerson;
  if (data.contactPhone !== undefined)  updates['Contact Phone']  = data.contactPhone;
  if (data.openDate !== undefined)      updates['Open Date']      = data.openDate;
  if (data.status !== undefined)        updates['Status']         = data.status;
  if (data.rentCost !== undefined)      updates['Rent Cost/Month']= data.rentCost;
  if (data.notes !== undefined)         updates['Notes']          = data.notes;
  return updateRow_('Locations', locId, updates);
}

function deleteLocation(locId) { return deleteRow_('Locations', locId); }

function getLocation(locId) {
  var rows = getSheetData_('Locations');
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['Location ID']) === String(locId)) {
      var r = rows[i];
      return { id: String(r['Location ID']), name: String(r['Location Name']), address: String(r['Address']||''), province: String(r['Province']||''), zone: String(r['Zone']||''), contactPerson: String(r['Contact Person']||''), contactPhone: String(r['Contact Phone']||''), openDate: safeDate_(r['Open Date']), status: String(r['Status']||''), rentCost: safeNum_(r['Rent Cost/Month']), notes: String(r['Notes']||'') };
    }
  }
  return null;
}


/* ═════════════════════════════════════════════════════
   INVESTOR CRUD
   ═════════════════════════════════════════════════════ */
function addInvestor(data) {
  var id = generateInvestorId_(data.nameEn, data.idCardLast4);
  var existing = getSheetData_('Investors');
  var baseId = id, suffix = 2;
  while (existing.some(function(r){ return String(r['Customer ID']) === id; })) { id = baseId + suffix; suffix++; }
  appendRow_('Investors', {
    'Customer ID': id, 'Location ID': data.locId || '',
    'Full Name (TH)': data.name || '', 'Full Name (EN)': data.nameEn || '',
    'Phone': data.phone || '', 'Email': data.email || '',
    'Line ID': data.line || '', 'Facebook': data.facebook || '',
    'Address': data.address || '', 'ID Card No. (Last 4)': data.idCardLast4 || '',
    'ID Card Photo': '', 'Bank Book Photo': '',
    'Amount Paid (฿)': data.amountPaid || 0, 'Risk Score': data.riskScore || 0,
    'Debt (฿)': data.debt || 0, 'Last Contact': '',
    'Status': data.status || 'Active', 'Join Date': data.joinDate || new Date(),
    'Notes': data.notes || ''
  });
  return id;
}

function updateInvestor(customerId, data) {
  var updates = {};
  if (data.locId !== undefined)       updates['Location ID']    = data.locId;
  if (data.name !== undefined)        updates['Full Name (TH)'] = data.name;
  if (data.nameEn !== undefined)      updates['Full Name (EN)'] = data.nameEn;
  if (data.phone !== undefined)       updates['Phone']          = data.phone;
  if (data.email !== undefined)       updates['Email']          = data.email;
  if (data.line !== undefined)        updates['Line ID']        = data.line;
  if (data.facebook !== undefined)    updates['Facebook']       = data.facebook;
  if (data.address !== undefined)     updates['Address']        = data.address;
  if (data.idCardLast4 !== undefined) updates['ID Card No. (Last 4)']  = data.idCardLast4;
  if (data.amountPaid !== undefined)  updates['Amount Paid (฿)']    = data.amountPaid;
  if (data.riskScore !== undefined)   updates['Risk Score']     = data.riskScore;
  if (data.debt !== undefined)        updates['Debt (฿)']           = data.debt;
  if (data.lastContact !== undefined) updates['Last Contact']   = data.lastContact;
  if (data.status !== undefined)      updates['Status']         = data.status;
  if (data.notes !== undefined)       updates['Notes']          = data.notes;
  return updateRow_('Investors', customerId, updates);
}

function deleteInvestor(customerId) { return deleteRow_('Investors', customerId); }

function getInvestor(customerId) {
  var rows = getSheetData_('Investors');
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['Customer ID']) === String(customerId)) {
      var r = rows[i];
      return { customerId: String(r['Customer ID']), locId: String(r['Location ID']), name: String(r['Full Name (TH)']||''), nameEn: String(r['Full Name (EN)']||''), phone: String(r['Phone']||''), email: String(r['Email']||''), line: String(r['Line ID']||''), facebook: String(r['Facebook']||''), address: String(r['Address']||''), idCardLast4: String(r['ID Card No. (Last 4)']||''), amountPaid: safeNum_(r['Amount Paid (฿)']), riskScore: safeNum_(r['Risk Score']), debt: safeNum_(r['Debt (฿)']), lastContact: safeDate_(r['Last Contact']), status: String(r['Status']||''), joinDate: safeDate_(r['Join Date']), notes: String(r['Notes']||'') };
    }
  }
  return null;
}


/* ═════════════════════════════════════════════════════
   PROJECT CRUD
   ═════════════════════════════════════════════════════ */
function addProject(data) {
  var id = generateProjectId_(data.name);
  appendRow_('Projects', {
    'Project ID': id, 'Location ID': data.locId||'', 'Project Name': data.name||'',
    'Cabinet Type': data.cabinetType||'', 'Total Value (฿)': data.totalValue||0,
    'Collected (฿)': 0, 'Progress %': 0, 'Status': 'Open',
    'Creation Date': new Date(), 'Sales Rep': data.salesRep||'', 'Notes': data.notes||''
  });
  return id;
}

function updateProject(projectId, data) {
  var updates = {};
  if (data.name !== undefined)        updates['Project Name'] = data.name;
  if (data.cabinetType !== undefined) updates['Cabinet Type']  = data.cabinetType;
  if (data.totalValue !== undefined)  updates['Total Value (฿)']   = data.totalValue;
  if (data.collected !== undefined)   updates['Collected (฿)']     = data.collected;
  if (data.status !== undefined)      updates['Status']        = data.status;
  if (data.salesRep !== undefined)    updates['Sales Rep']     = data.salesRep;
  if (data.notes !== undefined)       updates['Notes']         = data.notes;
  if (data.collected !== undefined || data.totalValue !== undefined) {
    var rows = getSheetData_('Projects');
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i]['Project ID']) === String(projectId)) {
        var tv = data.totalValue !== undefined ? data.totalValue : safeNum_(rows[i]['Total Value (฿)']);
        var cl = data.collected !== undefined ? data.collected : safeNum_(rows[i]['Collected (฿)']);
        updates['Progress %'] = tv > 0 ? Math.round((cl / tv) * 100) : 0;
        break;
      }
    }
  }
  return updateRow_('Projects', projectId, updates);
}


/* ═════════════════════════════════════════════════════
   📧 EMAIL AUTOMATION
   ═════════════════════════════════════════════════════ */
var ADMIN_EMAIL = 'aungelcomes.deal@gmail.com';

function sendAutomationEmail(config) {
  var emailSent = false;
  try {
    MailApp.sendEmail({ to: config.to, subject: config.subject, htmlBody: config.body });
    emailSent = true;
    Logger.log('✅ Email sent → ' + config.to + ' | ' + config.subject);
  } catch (e) {
    Logger.log('❌ MailApp.sendEmail FAILED → ' + config.to + ' | Error: ' + e.toString());
  }
  if (emailSent) {
    try {
      saveEmailLog({ relatedId: config.relatedId || '', relatedType: config.relatedType || '', toEmail: config.to, toName: config.toName || '', subject: config.subject, body: (config.body || '').substring(0, 200), template: config.template || 'Automation', triggerEvent: config.triggerEvent || 'System', autoGenerated: true });
    } catch (logErr) {
      Logger.log('⚠️ Email sent OK but log failed: ' + logErr.message);
    }
  }
  return emailSent;
}

function notifyAdmin_(subject, htmlBody, relatedId, triggerEvent) {
  return sendAutomationEmail({ to: ADMIN_EMAIL, toName: 'Admin', subject: subject, body: htmlBody, relatedId: relatedId || '', relatedType: 'Admin Notification', template: 'Admin Alert', triggerEvent: triggerEvent || 'System' });
}

function emailTemplate_(title, bodyHtml) {
  return '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;padding:20px;background:#f8f9fa">'
    + '<div style="background:#1a1f2e;color:#fff;padding:20px 28px;border-radius:12px 12px 0 0;text-align:center">'
    + '<h1 style="margin:0;font-size:1.5rem">🏢 DEAL<span style="color:#00B4D8">PAY</span></h1></div>'
    + '<div style="background:#fff;padding:24px 28px;border:1px solid #e5e7eb;border-top:none;border-radius:0 0 12px 12px">'
    + '<h2 style="color:#1a1f2e;margin-top:0">' + title + '</h2>'
    + bodyHtml
    + '<hr style="border:none;border-top:1px solid #e5e7eb;margin:20px 0">'
    + '<p style="color:#999;font-size:12px;text-align:center">อีเมลฉบับนี้ส่งอัตโนมัติจากระบบ DEALPAY<br>กรุณาอย่าตอบกลับอีเมลนี้</p>'
    + '</div></body></html>';
}

/* ═════════════════════════════════════════════════════
   HELPER: หา Investor จาก Sheet ด้วย ID
   รองรับทั้ง 'Customer ID' และ 'Investor ID' header
   ═════════════════════════════════════════════════════ */
function findInvestorById_(investorId) {
  if (!investorId) return null;
  var invRows = getSheetData_('Investors');
  for (var i = 0; i < invRows.length; i++) {
    var idInSheet = invRows[i]['Customer ID'] || invRows[i]['Investor ID'];
    if (String(idInSheet).trim() === String(investorId).trim()) {
      return invRows[i];
    }
  }
  return null;
}


/* ═════════════════════════════════════════════════════
   CASE CRUD
   ═════════════════════════════════════════════════════ */
function saveCase(data) {
  var caseId = generateCaseId_();
  appendRow_('Cases', {
    'Case ID': caseId, 'Location ID': data.locId||'', 'Customer ID': data.customerId||'',
    'Project ID': data.projectId||'', 'Case Type': data.caseType||'', 'Priority': data.priority||'Medium',
    'Status': data.status||'Open', 'Sub-Status': data.subStatus||'', 'Title': data.title||'',
    'Description': data.description||'', 'Assigned To': data.assignedTo||'', 'Assigned Team': data.assignedTeam||'',
    'Created Date': new Date(), 'Due Date': '', 'Resolution Date': '', 'Resolution Notes': '',
    'Estimated Cost (฿)': data.estimatedCost||0, 'Actual Cost (฿)': 0, 'Created By': data.createdBy||''
  });
  return caseId;
}

function updateCaseStatus(caseId, newStatus, notes) {
  var u = { 'Status': newStatus };
  if (notes) u['Resolution Notes'] = notes;
  if (newStatus === 'Resolved' || newStatus === 'Closed') u['Resolution Date'] = new Date();
  return updateRow_('Cases', caseId, u);
}


/* ═════════════════════════════════════════════════════
   COORD CASE (Work Order)
   ═════════════════════════════════════════════════════ */
function saveCoordCase(data) {
  var woId = generateWOId_();
  var locName = data.locName || '';
  if (!locName && data.locId) {
    var locs = getSheetData_('Locations');
    for (var i = 0; i < locs.length; i++) { if (String(locs[i]['Location ID']) === String(data.locId)) { locName = String(locs[i]['Location Name']); break; } }
  }
  appendRow_('CoordCases', {
    'Work Order ID': woId, 'Location ID': data.locId||'', 'Location Name': locName,
    'Investor ID': data.investorId||'', 'Investor Name': data.investorName||'',
    'Team': data.team||'repair', 'Priority': data.priority||'Medium',
    'Title': data.title||'', 'Description': data.description||'',
    'Photos': '', 'Status': 'pending', 'Created Date': new Date(), 'Created By': data.createdBy||'',
    'Resolved Date': '', 'Resolved By': '', 'Resolution Notes': ''
  });

  // 🔔 ส่งอีเมลแจ้ง Admin + Investor
  try {
    var teamTh = {repair:'🔧 ทีมซ่อม',location:'🏢 ทีม Location',billing:'💳 ทีมบัญชี',other:'📋 อื่นๆ'}[data.team] || data.team;
    var prioColor = {Critical:'#ef4444',High:'#f59e0b',Medium:'#3b82f6',Low:'#22c55e'}[data.priority] || '#64748b';
    var tableRows = ''
      + '<tr><td style="padding:8px;color:#666;width:140px">Work Order</td><td style="padding:8px;font-weight:700;font-size:1.1em">' + woId + '</td></tr>'
      + '<tr><td style="padding:8px;color:#666">โลเคชั่น</td><td style="padding:8px;font-weight:700">' + locName + '</td></tr>'
      + '<tr><td style="padding:8px;color:#666">หัวข้อ</td><td style="padding:8px;font-weight:700">' + (data.title||'-') + '</td></tr>'
      + '<tr><td style="padding:8px;color:#666">ทีมรับผิดชอบ</td><td style="padding:8px">' + teamTh + '</td></tr>'
      + '<tr><td style="padding:8px;color:#666">ความเร่งด่วน</td><td style="padding:8px"><span style="color:' + prioColor + ';font-weight:700">' + (data.priority||'Medium') + '</span></td></tr>'
      + (data.investorName ? '<tr><td style="padding:8px;color:#666">นักลงทุน</td><td style="padding:8px">' + data.investorName + '</td></tr>' : '')
      + '<tr><td style="padding:8px;color:#666">รายละเอียด</td><td style="padding:8px">' + (data.description||'-') + '</td></tr>'
      + '<tr><td style="padding:8px;color:#666">ผู้สร้าง</td><td style="padding:8px">' + (data.createdBy||'-') + '</td></tr>'
      + '<tr><td style="padding:8px;color:#666">เวลา</td><td style="padding:8px">' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss') + '</td></tr>';

    // ส่ง Admin
    var bodyAdmin = emailTemplate_('🆕 เคสประสานงานใหม่', '<table style="width:100%;border-collapse:collapse;margin:12px 0">' + tableRows + '</table>');
    notifyAdmin_('[DEALPAY] 🆕 เคสใหม่ ' + woId + ' — ' + (data.title||locName), bodyAdmin, woId, 'New Coord Case');

    // ✅ FIX 2: ส่งอีเมลแจ้ง Investor ด้วย (ถ้ามี investorId)
    if (data.investorId) {
      var investor = findInvestorById_(data.investorId);
      if (investor && investor['Email']) {
        var invName = investor['Full Name (TH)'] || investor['Full Name (EN)'] || 'ลูกค้า';
        var bodyInv = emailTemplate_('🆕 แจ้งเปิดเคสประสานงาน',
          '<p>เรียน คุณ' + invName + '</p>'
          + '<p>ระบบ DEALPAY ขอแจ้งให้ทราบว่ามีการเปิดเคสประสานงานใหม่ที่เกี่ยวข้องกับโลเคชั่นของท่าน</p>'
          + '<table style="width:100%;border-collapse:collapse;margin:12px 0">'
          + '<tr><td style="padding:8px;color:#666;width:140px">Work Order</td><td style="padding:8px;font-weight:700">' + woId + '</td></tr>'
          + '<tr><td style="padding:8px;color:#666">โลเคชั่น</td><td style="padding:8px;font-weight:700">' + locName + '</td></tr>'
          + '<tr><td style="padding:8px;color:#666">หัวข้อ</td><td style="padding:8px">' + (data.title||'-') + '</td></tr>'
          + '<tr><td style="padding:8px;color:#666">ทีมรับผิดชอบ</td><td style="padding:8px">' + teamTh + '</td></tr>'
          + '<tr><td style="padding:8px;color:#666">ความเร่งด่วน</td><td style="padding:8px"><span style="color:' + prioColor + ';font-weight:700">' + (data.priority||'Medium') + '</span></td></tr>'
          + '</table>'
          + '<p style="color:#666">ท่านจะได้รับอีเมลแจ้งเตือนอีกครั้งเมื่อสถานะเปลี่ยนแปลง</p>'
          + '<p style="color:#666">ติดต่อทีมงาน: 📞 02-XXX-XXXX | 💬 Line: @dealpay</p>'
        );
        sendAutomationEmail({
          to: investor['Email'], toName: invName,
          subject: '[DEALPAY] แจ้งเปิดเคสใหม่ ' + woId + ' — ' + locName,
          body: bodyInv, relatedId: woId, relatedType: 'WorkOrder → Investor',
          template: 'New Coord Case → Investor', triggerEvent: 'New Coord Case → Investor'
        });
        Logger.log('📧 ส่งเมลแจ้ง Investor สำเร็จ: ' + investor['Email']);
      } else {
        Logger.log('⚠️ ไม่พบ Email Investor สำหรับ ID: ' + data.investorId);
      }
    }
  } catch (emailErr) {
    Logger.log('⚠️ saveCoordCase email failed but Sheet SAVED: ' + emailErr.message);
  }

  return woId;
}

function updateCoordStatus(woId, newStatus, notes, resolvedBy) {
  Logger.log('🔄 เริ่มอัปเดตสถานะ: ' + woId);

  var u = { 'Status': newStatus };
  if (notes) u['Resolution Notes'] = notes;
  if (newStatus === 'resolved') {
    u['Resolved Date'] = new Date();
    u['Resolved By'] = resolvedBy || '';
  }

  // ═══ STEP 1: อัปเดต Sheet ═══
  var result = updateRow_('CoordCases', woId, u);

  // Flexible match ถ้าไม่เจอ
  if (!result) {
    Logger.log('⚠️ updateRow_ ไม่เจอ "' + woId + '" — ลอง flexible match...');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('CoordCases');
    if (sh && sh.getLastRow() > 1) {
      var d = sh.getDataRange().getValues(), h = d[0];
      var target = String(woId).trim();
      for (var i = 1; i < d.length; i++) {
        if (String(d[i][0]).trim() === target) {
          var keys = Object.keys(u);
          for (var k = 0; k < keys.length; k++) { var ci = h.indexOf(keys[k]); if (ci >= 0) sh.getRange(i+1, ci+1).setValue(u[keys[k]]); }
          result = woId;
          Logger.log('✅ Flexible match → row ' + (i+1));
          break;
        }
      }
    }
    if (!result) {
      var allRows = getSheetData_('CoordCases');
      var ids = allRows.map(function(r){ return '"' + String(r['Work Order ID']).trim() + '"'; });
      Logger.log('❌ ไม่เจอ WO: "' + woId + '" ใน ' + allRows.length + ' rows: ' + ids.slice(0,10).join(', '));
    }
  } else {
    Logger.log('✅ Sheet updated OK');
  }

  // ดึงข้อมูล WO สำหรับส่งอีเมล
  var allCoordCases = getSheetData_('CoordCases');
  var woData = null;
  for (var wi = 0; wi < allCoordCases.length; wi++) {
    if (String(allCoordCases[wi]['Work Order ID']).trim() === String(woId).trim()) {
      woData = allCoordCases[wi];
      break;
    }
  }

  if (!woData) {
    Logger.log('❌ ไม่พบข้อมูล Work Order ID: ' + woId + ' สำหรับส่งอีเมล');
    return { saved: !!result, emailed: false, woId: woId, status: newStatus };
  }

  // ═══ STEP 2: ส่งอีเมล ═══
  var emailOk = false;
  try {
    var statusTh = {pending:'รอรับเคส', in_progress:'กำลังดำเนินการ', resolved:'เสร็จสิ้น', cancelled:'ยกเลิก'}[newStatus] || newStatus;
    var statusColor = {pending:'#f59e0b',in_progress:'#3b82f6',resolved:'#22c55e',cancelled:'#ef4444'}[newStatus] || '#64748b';

    // ส่ง Admin
    var bodyAdmin = emailTemplate_('📋 อัปเดตสถานะงาน', ''
      + '<table style="width:100%;border-collapse:collapse;margin:12px 0">'
      + '<tr><td style="padding:8px;color:#666;width:140px">Work Order</td><td style="padding:8px;font-weight:700">' + woId + '</td></tr>'
      + '<tr><td style="padding:8px;color:#666">โลเคชั่น</td><td style="padding:8px;font-weight:700">' + String(woData['Location Name']||'') + '</td></tr>'
      + '<tr><td style="padding:8px;color:#666">หัวข้อ</td><td style="padding:8px">' + String(woData['Title']||'') + '</td></tr>'
      + '<tr><td style="padding:8px;color:#666">สถานะใหม่</td><td style="padding:8px"><span style="background:' + statusColor + ';color:#fff;padding:4px 12px;border-radius:12px;font-weight:700">' + statusTh + '</span></td></tr>'
      + '<tr><td style="padding:8px;color:#666">บันทึก</td><td style="padding:8px">' + (notes || '-') + '</td></tr>'
      + '<tr><td style="padding:8px;color:#666">ผู้อัปเดต</td><td style="padding:8px">' + (resolvedBy || 'System') + '</td></tr>'
      + '<tr><td style="padding:8px;color:#666">เวลา</td><td style="padding:8px">' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss') + '</td></tr>'
      + '</table>'
    );
    var emailAdminOk = sendAutomationEmail({ to: ADMIN_EMAIL, toName: 'Admin', subject: '[DEALPAY] WO ' + woId + ' — ' + statusTh, body: bodyAdmin, relatedId: woId, relatedType: 'WorkOrder', template: 'Coord Status Update', triggerEvent: 'Status → ' + newStatus });

    // ✅ FIX 3: ส่ง Investor — ใช้ findInvestorById_ ที่รองรับทั้ง 'Customer ID' และ 'Investor ID'
    var emailInvestorOk = false;
    var investorIdForLookup = woData['Investor ID'] || woData['investor_id'] || '';
    if (investorIdForLookup) {
      var investor = findInvestorById_(investorIdForLookup);
      if (investor && investor['Email']) {
        var invName = investor['Full Name (TH)'] || investor['Full Name (EN)'] || 'ลูกค้า';
        var bodyInv = emailTemplate_('📋 แจ้งผลการดำเนินงาน',
          '<p>เรียน คุณ' + invName + '</p>'
          + '<p>ขอแจ้งอัปเดตเคส <b>' + woId + '</b> ของท่าน</p>'
          + '<p><b>สถานะปัจจุบัน:</b> <span style="background:' + statusColor + ';color:#fff;padding:4px 12px;border-radius:12px;font-weight:700">' + statusTh + '</span></p>'
          + '<p><b>รายละเอียด:</b> ' + (notes || 'เจ้าหน้าที่กำลังดำเนินการ') + '</p>'
          + '<p style="color:#666">ติดต่อทีมงาน: 📞 02-XXX-XXXX | 💬 Line: @dealpay</p>'
        );
        emailInvestorOk = sendAutomationEmail({ to: investor['Email'], toName: invName, subject: '[DEALPAY] แจ้งสถานะเคส ' + woId, body: bodyInv, relatedId: woId, relatedType: 'WorkOrder → Investor', template: 'Investor Notification', triggerEvent: 'Coord Status → Investor' });
        Logger.log('📧 ส่งเมลหาลูกค้าสำเร็จ: ' + investor['Email']);
      } else {
        Logger.log('⚠️ ไม่พบ Email ลูกค้า Investor ID: ' + investorIdForLookup);
      }
    }
    emailOk = emailAdminOk || emailInvestorOk;
  } catch (emailErr) {
    Logger.log('⚠️ Email error: ' + emailErr.message);
  }

  Logger.log('✅ Done: Sheet=' + (result ? 'OK' : 'MISS') + ' Email=' + (emailOk ? 'OK' : 'FAIL'));
  return { saved: !!result, emailed: emailOk, woId: woId, status: newStatus };
}


/* ═════════════════════════════════════════════════════
   COMMUNICATION
   ═════════════════════════════════════════════════════ */
function saveCommunication(data) {
  var commId = generateCommId_();
  appendRow_('Communications', {
    'Comm ID': commId, 'Case ID': data.caseId||'', 'Customer ID': data.customerId||'',
    'Location ID': data.locId||'', 'Comm Type': data.commType||'Phone Call', 'Direction': data.direction||'Outbound',
    'Subject': data.subject||'', 'Message': data.message||'', 'Contact Result': data.contactResult||'',
    'Date': new Date(), 'Sent By': data.sentBy||'', 'Status': data.commStatus||'Sent',
    'Template Used': data.template||'', 'Auto-Generated': data.autoGenerated||false, 'Attachments': ''
  });
  return commId;
}


/* ═════════════════════════════════════════════════════
   INVOICE
   ═════════════════════════════════════════════════════ */
function addInvoice(data) {
  var id = generateInvoiceId_();
  var amt = safeNum_(data.amount);
  appendRow_('Invoices', {
    'Invoice ID': id, 'Customer ID': data.customerId||'', 'Project ID': data.projectId||'',
    'Location ID': data.locId||'', 'Invoice Date': new Date(), 'Due Date': data.dueDate||'',
    'Amount (฿)': amt, 'Payment Received (฿)': 0, 'Outstanding (฿)': amt,
    'Status': 'Sent', 'Payment Terms': data.paymentTerms||'Net 30',
    'Sent Date': new Date(), 'Reminder Sent': 'No', 'Reminder Date': '',
    'Items Description': data.items||'', 'Notes': data.notes||''
  });

  notifyAdmin_(
    '[DEALPAY] 🧾 Invoice ใหม่ ' + id + ' — ฿' + amt.toLocaleString(),
    emailTemplate_('🧾 Invoice ใหม่', '<table style="width:100%;border-collapse:collapse;margin:12px 0"><tr><td style="padding:8px;color:#666">Invoice ID</td><td style="padding:8px;font-weight:700">' + id + '</td></tr><tr><td style="padding:8px;color:#666">Customer</td><td style="padding:8px">' + (data.customerId||'-') + '</td></tr><tr><td style="padding:8px;color:#666">จำนวนเงิน</td><td style="padding:8px;color:#3b82f6;font-weight:700;font-size:1.1em">฿' + amt.toLocaleString() + '</td></tr><tr><td style="padding:8px;color:#666">กำหนดชำระ</td><td style="padding:8px">' + (data.dueDate||'Net 30') + '</td></tr></table>'),
    id, 'Invoice Created'
  );

  if (data.investorEmail) {
    sendAutomationEmail({ to: data.investorEmail, toName: data.investorName || '', subject: '[DEALPAY] ใบแจ้งหนี้ ' + id + ' — ฿' + amt.toLocaleString(), body: emailTemplate_('🧾 ใบแจ้งหนี้', '<p>เรียน คุณ' + (data.investorName||'ลูกค้า') + '</p><table style="width:100%;border-collapse:collapse;margin:12px 0;background:#f0f9ff;border:1px solid #bae6fd;border-radius:8px"><tr><td style="padding:12px;color:#666">Invoice</td><td style="padding:12px;font-weight:700">' + id + '</td></tr><tr><td style="padding:12px;color:#666">จำนวนเงิน</td><td style="padding:12px;color:#3b82f6;font-weight:700;font-size:1.2em">฿' + amt.toLocaleString() + '</td></tr><tr><td style="padding:12px;color:#666">กำหนดชำระ</td><td style="padding:12px;font-weight:700">' + (data.dueDate||'ภายใน 30 วัน') + '</td></tr></table><p>📞 02-XXX-XXXX | 📧 support@dealpay.com</p>'), relatedId: id, relatedType: 'Invoice → Investor', template: 'Invoice Notification', triggerEvent: 'Invoice Created → Investor' });
  }

  return id;
}

function recordPayment(invoiceId, amount, method, refNo) {
  var rows = getSheetData_('Invoices');
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['Invoice ID']) === String(invoiceId)) {
      var inv = rows[i];
      var newReceived = safeNum_(inv['Payment Received (฿)']) + safeNum_(amount);
      var totalAmt = safeNum_(inv['Amount (฿)']);
      var outstanding = totalAmt - newReceived;
      var newStatus = outstanding <= 0 ? 'Paid' : 'Partially Paid';
      updateRow_('Invoices', invoiceId, { 'Payment Received (฿)': newReceived, 'Outstanding (฿)': Math.max(0, outstanding), 'Status': newStatus });
      appendRow_('Transactions', { 'Transaction ID': generateTxnId_(), 'Project ID': String(inv['Project ID']||''), 'Customer ID': String(inv['Customer ID']||''), 'Location ID': String(inv['Location ID']||''), 'Amount (฿)': amount, 'Transaction Date': new Date(), 'Payment Method': method||'Bank Transfer', 'Reference No.': refNo||'', 'Status': 'Completed', 'Receipt URL': '', 'Notes': 'Payment for ' + invoiceId });
      var statusColor = newStatus === 'Paid' ? '#22c55e' : '#3b82f6';
      notifyAdmin_('[DEALPAY] 💵 รับชำระ ฿' + Number(amount).toLocaleString() + ' — ' + invoiceId + ' (' + newStatus + ')', emailTemplate_('💵 รับชำระเงิน', '<table style="width:100%;border-collapse:collapse;margin:12px 0"><tr><td style="padding:8px;color:#666">Invoice</td><td style="padding:8px;font-weight:700">' + invoiceId + '</td></tr><tr><td style="padding:8px;color:#666">ยอดรับ</td><td style="padding:8px;color:#22c55e;font-weight:700;font-size:1.2em">฿' + Number(amount).toLocaleString() + '</td></tr><tr><td style="padding:8px;color:#666">สถานะ</td><td style="padding:8px"><span style="background:' + statusColor + ';color:#fff;padding:4px 12px;border-radius:12px;font-weight:700">' + newStatus + '</span></td></tr></table>'), invoiceId, 'Payment Received (฿)');
      return { invoiceId: invoiceId, status: newStatus, outstanding: Math.max(0, outstanding) };
    }
  }
  return null;
}


/* ═════════════════════════════════════════════════════
   DEBT COLLECTION
   ═════════════════════════════════════════════════════ */
function addDebtCollection(data) {
  var id = generateCollectionId_();
  appendRow_('DebtCollection', { 'Collection ID': id, 'Customer ID': data.customerId||'', 'Invoice ID': data.invoiceId||'', 'Location ID': data.locId||'', 'Total Debt (฿)': data.totalDebt||0, 'Days Overdue': data.daysOverdue||0, 'Collection Status': 'Not Contacted', 'Escalation Level': 'Level 1', 'Assigned To': data.assignedTo||'', 'Contact Attempts': 0, 'Last Contact Date': '', 'Next Follow-up': data.nextFollowUp||'', 'Notes': data.notes||'' });
  var body = emailTemplate_('💰 รายการทวงหนี้ใหม่', '<table style="width:100%;border-collapse:collapse;margin:12px 0"><tr><td style="padding:8px;color:#666;width:140px">Collection ID</td><td style="padding:8px;font-weight:700">' + id + '</td></tr><tr><td style="padding:8px;color:#666">Invoice</td><td style="padding:8px">' + (data.invoiceId||'-') + '</td></tr><tr><td style="padding:8px;color:#666">ยอดหนี้</td><td style="padding:8px;color:#ef4444;font-weight:700;font-size:1.1em">฿' + Number(data.totalDebt||0).toLocaleString() + '</td></tr><tr><td style="padding:8px;color:#666">เกินกำหนด</td><td style="padding:8px;color:#f59e0b;font-weight:700">' + (data.daysOverdue||0) + ' วัน</td></tr></table>');
  notifyAdmin_('[DEALPAY] 💰 ทวงหนี้ใหม่ ' + id + ' — ฿' + Number(data.totalDebt||0).toLocaleString(), body, id, 'Debt Entry');
  if (data.investorEmail) {
    sendAutomationEmail({ to: data.investorEmail, toName: data.investorName || '', subject: '[DEALPAY] แจ้งเตือนยอดค้างชำระ ' + (data.invoiceId||id), body: emailTemplate_('แจ้งเตือนรายการค้างชำระ', '<p>เรียน คุณ' + (data.investorName||'ลูกค้า') + '</p><p>ระบบตรวจพบยอดค้างชำระ Invoice: <b>' + (data.invoiceId||'-') + '</b></p><table style="width:100%;border-collapse:collapse;margin:12px 0;background:#fff8f0;border:1px solid #fed7aa;border-radius:8px"><tr><td style="padding:12px;color:#666;width:140px">จำนวนเงิน</td><td style="padding:12px;color:#ef4444;font-weight:700;font-size:1.2em">฿' + Number(data.totalDebt||0).toLocaleString() + '</td></tr><tr><td style="padding:12px;color:#666">เกินกำหนด</td><td style="padding:12px;color:#f59e0b;font-weight:700">' + (data.daysOverdue||0) + ' วัน</td></tr></table><p>📞 02-XXX-XXXX | 📧 support@dealpay.com | 💬 Line: @dealpay</p>'), relatedId: id, relatedType: 'Debt → Investor', template: 'Debt Reminder', triggerEvent: 'Debt Entry → Investor' });
  }
  return id;
}

function logDebtContact(collectionId, notes) {
  var rows = getSheetData_('DebtCollection');
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['Collection ID']) === String(collectionId)) {
      var dc = rows[i];
      var attempts = safeNum_(dc['Contact Attempts']) + 1;
      var now = new Date();
      var existing = String(dc['Notes'] || '');
      var newNote = '[' + Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm') + '] ' + notes;
      updateRow_('DebtCollection', collectionId, { 'Contact Attempts': attempts, 'Last Contact Date': now, 'Notes': existing ? existing + '\n' + newNote : newNote });
      notifyAdmin_('[DEALPAY] 📞 ติดตามหนี้ครั้งที่ ' + attempts + ' — ' + collectionId, emailTemplate_('📞 บันทึกการติดตามหนี้', '<table style="width:100%;border-collapse:collapse;margin:12px 0"><tr><td style="padding:8px;color:#666">Collection ID</td><td style="padding:8px;font-weight:700">' + collectionId + '</td></tr><tr><td style="padding:8px;color:#666">ครั้งที่</td><td style="padding:8px;font-weight:700;font-size:1.1em">' + attempts + '</td></tr><tr><td style="padding:8px;color:#666">บันทึก</td><td style="padding:8px">' + notes + '</td></tr></table>'), collectionId, 'Debt Contact #' + attempts);
      return { collectionId: collectionId, attempts: attempts };
    }
  }
  return null;
}


/* ═════════════════════════════════════════════════════
   SURVEY
   ═════════════════════════════════════════════════════ */
function saveSurvey(data) {
  var id = generateSurveyId_();
  appendRow_('SatisfactionSurvey', { 'Survey ID': id, 'Case ID': data.caseId||'', 'Customer ID': data.customerId||'', 'Location ID': data.locId||'', 'Rating (1-5)': data.rating||'', 'Speed Rating': data.speedRating||'', 'Quality Rating': data.qualityRating||'', 'Service Rating': data.serviceRating||'', 'Comment': data.comment||'', 'Survey Date': new Date(), 'Survey Channel': data.channel||'Phone', 'Follow-up Required': data.followUp||false, 'Investor Name': data.investorName||'', 'Location Name': data.locationName||'', 'Agent': data.agent||'', 'Contact Result': data.contactResult||'' });
  return id;
}


/* ═════════════════════════════════════════════════════
   EMAIL LOG
   ═════════════════════════════════════════════════════ */
function saveEmailLog(data) {
  var id = countRows_('EmailLog') + 1;
  appendRow_('EmailLog', { 'Email ID': id, 'Related ID': data.relatedId||'', 'Related Type': data.relatedType||'', 'To Email': data.toEmail||'', 'To Name': data.toName||'', 'Subject': data.subject||'', 'Body Preview': (data.body||'').substring(0,200), 'Template': data.template||'', 'Sent Date': new Date(), 'Status': 'Sent', 'Trigger Event': data.triggerEvent||'', 'Auto-Generated': data.autoGenerated||false });
  return id;
}


/* ═════════════════════════════════════════════════════
   CRM ENTRY
   ═════════════════════════════════════════════════════ */
function saveCRMEntry(data) {
  var caseId = '';
  try {
    caseId = generateCaseId_();
    Logger.log('📝 saveCRMEntry เริ่ม: ' + caseId + ' | ' + (data.contactLabel||''));
    appendRow_('Cases', { 'Case ID': caseId, 'Location ID': data.locId||'', 'Customer ID': data.customerId||'', 'Project ID': '', 'Case Type': data.caseType||'', 'Priority': 'Medium', 'Status': 'Closed', 'Sub-Status': data.subType||'', 'Title': (data.contactLabel||'') + ' — ' + (data.investorName||''), 'Description': data.callNotes||'', 'Assigned To': data.agent||'', 'Assigned Team': 'Customer Support', 'Created Date': new Date(), 'Due Date': '', 'Resolution Date': new Date(), 'Resolution Notes': data.contactResult||'', 'Estimated Cost (฿)': 0, 'Actual Cost (฿)': 0, 'Created By': data.agent||'System' });
    Logger.log('  ✅ Cases Sheet saved');
    var commType = 'Phone Call';
    if (data.contactType === 'email') commType = 'Email';
    else if (data.contactType === 'multichannel') commType = data.formChannel || 'Line';
    appendRow_('Communications', { 'Comm ID': generateCommId_(), 'Case ID': caseId, 'Customer ID': data.customerId||'', 'Location ID': data.locId||'', 'Comm Type': commType, 'Direction': 'Outbound', 'Subject': (data.contactLabel||'') + ' — ' + (data.investorName||''), 'Message': data.callNotes||'', 'Contact Result': data.contactResult||'', 'Date': new Date(), 'Sent By': data.agent||'System', 'Status': 'Completed', 'Template Used': '', 'Auto-Generated': false, 'Attachments': '' });
    Logger.log('  ✅ Communications Sheet saved');
    if (data.contactType === 'satisfaction' && data.overallScore) {
      appendRow_('SatisfactionSurvey', { 'Survey ID': generateSurveyId_(), 'Case ID': caseId, 'Customer ID': data.customerId||'', 'Location ID': data.locId||'', 'Rating (1-5)': safeNum_(data.overallScore), 'Speed Rating': '', 'Quality Rating': data.machineCondition||'', 'Service Rating': data.serviceQuality||'', 'Comment': data.feedback||'', 'Survey Date': new Date(), 'Survey Channel': commType, 'Follow-up Required': (data.followUp && data.followUp !== 'ไม่ต้อง') ? true : false, 'Investor Name': data.investorName||'', 'Location Name': data.locName||'', 'Agent': data.agent||'', 'Contact Result': data.contactResult||'' });
      Logger.log('  ✅ Survey Sheet saved');
    }
  } catch (sheetErr) {
    Logger.log('❌ saveCRMEntry Sheet FAIL: ' + sheetErr.message);
    throw new Error('บันทึก Sheet ไม่สำเร็จ: ' + sheetErr.message);
  }

  var emailLogged = false;
  try {
    if (data.investorEmail) {
      var ts = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm');
      var mainEmailSent = sendAutomationEmail({ to: data.investorEmail, toName: data.investorName||'', subject: '[DEALPAY] ' + (data.contactLabel||'สรุปเคส') + ' (Ref: ' + caseId + ')', body: emailTemplate_('📋 สรุปการติดต่อ', '<p>เรียน คุณ' + (data.investorName||'ลูกค้า') + '</p><p>ทีม DEALPAY ขอสรุปผลการติดต่อดังนี้:</p><table style="width:100%;border-collapse:collapse;margin:12px 0"><tr><td style="padding:8px;color:#666;width:140px">เลขอ้างอิง</td><td style="padding:8px;font-weight:700">' + caseId + '</td></tr><tr><td style="padding:8px;color:#666">ประเภท</td><td style="padding:8px">' + (data.contactLabel||'-') + '</td></tr><tr><td style="padding:8px;color:#666">วันที่</td><td style="padding:8px">' + ts + '</td></tr><tr><td style="padding:8px;color:#666">ผลการติดต่อ</td><td style="padding:8px;font-weight:700">' + (data.contactResult||'-') + '</td></tr></table><p style="color:#666">อ้างอิง <b>' + caseId + '</b> | 📞 02-XXX-XXXX | 💬 @dealpay</p>'), relatedId: caseId, relatedType: 'Case', template: 'CRM Summary', triggerEvent: 'CRM Save' });
      if (mainEmailSent) emailLogged = true;
      if (data.contactFailed) {
        sendAutomationEmail({ to: data.investorEmail, toName: data.investorName||'', subject: '[DEALPAY] แจ้งความพยายามติดต่อ (Ref: ' + caseId + ')', body: emailTemplate_('⚠️ แจ้งความพยายามติดต่อ', '<p>เรียน คุณ' + (data.investorName||'') + '</p><p>ทีมงานพยายามติดต่อท่านแต่ไม่สามารถติดต่อได้</p><p>กรุณาติดต่อกลับ: 📞 02-XXX-XXXX | 💬 @dealpay | อ้างอิง ' + caseId + '</p>'), relatedId: caseId, relatedType: 'Case', template: 'Auto Follow-up', triggerEvent: 'Contact Failed' });
      }
    }
    notifyAdmin_('[DEALPAY] 📋 CC ' + caseId + ' — ' + (data.contactLabel||''), emailTemplate_('📋 CC บันทึกเคส', '<table style="width:100%;border-collapse:collapse;margin:12px 0"><tr><td style="padding:8px;color:#666">Case</td><td style="padding:8px;font-weight:700">' + caseId + '</td></tr><tr><td style="padding:8px;color:#666">ประเภท</td><td style="padding:8px">' + (data.contactLabel||'-') + '</td></tr><tr><td style="padding:8px;color:#666">ลูกค้า</td><td style="padding:8px">' + (data.investorName||'-') + '</td></tr><tr><td style="padding:8px;color:#666">ผลติดต่อ</td><td style="padding:8px;font-weight:700">' + (data.contactResult||'-') + '</td></tr></table>'), caseId, 'CRM Entry');
  } catch (emailErr) {
    Logger.log('⚠️ CRM emails failed but Sheets SAVED OK: ' + emailErr.message);
    emailLogged = false;
  }

  Logger.log('✅ saveCRMEntry เสร็จ: ' + caseId + ' | emailLogged=' + emailLogged);
  return { caseId: caseId, success: true, emailLogged: emailLogged };
}


/* ═════════════════════════════════════════════════════
   BILL ENTRY
   ═════════════════════════════════════════════════════ */
function saveBillEntry(data) {
  var invoiceId = '';
  var amt = safeNum_(data.amount);
  try {
    invoiceId = generateInvoiceId_();
    appendRow_('Invoices', { 'Invoice ID': invoiceId, 'Customer ID': data.customerId||'', 'Project ID': '', 'Location ID': data.locId||'', 'Invoice Date': new Date(), 'Due Date': data.dueDate||'', 'Amount (฿)': amt, 'Payment Received (฿)': 0, 'Outstanding (฿)': amt, 'Status': 'Sent', 'Payment Terms': data.dueDate ? 'Custom' : 'Net 30', 'Sent Date': new Date(), 'Reminder Sent': 'No', 'Reminder Date': '', 'Items Description': data.billNote||'ค่าบริการ/ค่าเช่าเครื่อง', 'Notes': 'CC billing — Case: ' + (data.caseId||'') });
    Logger.log('✅ saveBillEntry Sheet saved: ' + invoiceId);
  } catch (sheetErr) {
    Logger.log('❌ saveBillEntry Sheet FAIL: ' + sheetErr.message);
    throw new Error('บันทึก Invoice ไม่สำเร็จ: ' + sheetErr.message);
  }
  try {
    if (data.investorEmail) {
      sendAutomationEmail({ to: data.investorEmail, toName: data.investorName||'', subject: '[DEALPAY] ใบแจ้งหนี้ ' + invoiceId + ' — ฿' + amt.toLocaleString(), body: emailTemplate_('🧾 ใบแจ้งหนี้', '<p>เรียน คุณ' + (data.investorName||'ลูกค้า') + '</p><div style="background:#f0f9ff;border:2px solid #bae6fd;border-radius:12px;padding:20px;margin:16px 0;text-align:center"><div style="font-size:.85rem;color:#666;margin-bottom:4px">ยอดที่ต้องชำระ</div><div style="font-size:2rem;font-weight:800;color:#0284c7">฿' + amt.toLocaleString() + '</div></div><table style="width:100%;border-collapse:collapse;margin:12px 0"><tr><td style="padding:8px;color:#666">Invoice</td><td style="padding:8px;font-weight:700">' + invoiceId + '</td></tr><tr><td style="padding:8px;color:#666">กำหนดชำระ</td><td style="padding:8px;font-weight:700;color:#f59e0b">' + (data.dueDate||'ภายใน 30 วัน') + '</td></tr></table><p>📞 02-XXX-XXXX | 💬 @dealpay</p>'), relatedId: invoiceId, relatedType: 'Invoice', template: 'Bill Notification', triggerEvent: 'CC Bill Send' });
    }
    notifyAdmin_('[DEALPAY] 💰 บิล ' + invoiceId + ' ฿' + amt.toLocaleString() + ' → ' + (data.investorName||''), emailTemplate_('💰 ส่งบิล', '<table style="width:100%;border-collapse:collapse;margin:12px 0"><tr><td style="padding:8px;color:#666">Invoice</td><td style="padding:8px;font-weight:700">' + invoiceId + '</td></tr><tr><td style="padding:8px;color:#666">ลูกค้า</td><td style="padding:8px">' + (data.investorName||'-') + '</td></tr><tr><td style="padding:8px;color:#666">ยอด</td><td style="padding:8px;color:#0284c7;font-weight:700">฿' + amt.toLocaleString() + '</td></tr></table>'), invoiceId, 'CC Bill Send');
  } catch (emailErr) {
    Logger.log('⚠️ Bill emails failed but Sheet SAVED: ' + emailErr.message);
  }
  return { invoiceId: invoiceId, success: true };
}


/* ═════════════════════════════════════════════════════
   INDIVIDUAL SHEET READERS
   ═════════════════════════════════════════════════════ */
function getLocations()    { return getSheetData_('Locations'); }
function getProjects()     { return getSheetData_('Projects'); }
function getInvestors()    { return getSheetData_('Investors'); }
function getTransactions() { return getSheetData_('Transactions'); }
function getCases()        { return getSheetData_('Cases'); }
function getCoordCases()   { return getSheetData_('CoordCases'); }
function getComms()        { return getSheetData_('Communications'); }
function getInvoices()     { return getSheetData_('Invoices'); }
function getDebt()         { return getSheetData_('DebtCollection'); }
function getSurveys()      { return getSheetData_('SatisfactionSurvey'); }
function getEmailLog()     { return getSheetData_('EmailLog'); }
function getConfig()       { return getSheetData_('SystemConfig'); }


/* ═════════════════════════════════════════════════════
   PING & DEBUG
   ═════════════════════════════════════════════════════ */
function ping() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets().map(function(s){ return s.getName(); });
  var required = ['Locations','Projects','Investors','Cases','Communications','CoordCases','SatisfactionSurvey','Invoices','DebtCollection','Transactions','EmailLog','SystemConfig'];
  var missing = [];
  for (var i = 0; i < required.length; i++) { if (sheets.indexOf(required[i]) === -1) missing.push(required[i]); }
  var emailStatus = '';
  try { var quota = MailApp.getRemainingDailyQuota(); emailStatus = '📧 Email: ✅ พร้อมส่ง (เหลือ ' + quota + ' ฉบับ/วัน)'; }
  catch(e) { emailStatus = '📧 Email: ❌ ไม่พร้อม! รัน authorizeAll() ก่อน (' + e.message + ')'; }
  var coordInfo = '';
  try {
    var coordRows = getSheetData_('CoordCases');
    coordInfo = '📋 CoordCases: ' + coordRows.length + ' rows';
    if (coordRows.length > 0) {
      var ids = coordRows.map(function(r){ return String(r['Work Order ID']); });
      coordInfo += ' [' + ids.slice(0,5).join(', ') + (ids.length > 5 ? ', ...' : '') + ']';
    }
  } catch(e) { coordInfo = '📋 CoordCases: ❌ ' + e.message; }
  if (missing.length > 0) return '⚠️ ขาด Sheet: ' + missing.join(', ') + '\nรัน setupSheets() ก่อน!\n\n' + emailStatus;
  return '✅ Sheets ครบ ' + required.length + ' แผ่น\n' + emailStatus + '\n' + coordInfo;
}

function debugCoordCases() {
  var rows = getSheetData_('CoordCases');
  Logger.log('═══ CoordCases: ' + rows.length + ' rows ═══');
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    Logger.log('  [' + i + '] ID="' + String(r['Work Order ID']) + '" Status="' + String(r['Status']) + '" InvestorID="' + String(r['Investor ID']||'') + '"');
  }
  if (rows.length > 0) {
    var firstId = String(rows[0]['Work Order ID']);
    Logger.log('\n🧪 ทดสอบ findInvestorById_ ด้วย Investor ID: "' + String(rows[0]['Investor ID']||'') + '"');
    var inv = findInvestorById_(String(rows[0]['Investor ID']||''));
    Logger.log(inv ? '✅ พบ Investor: ' + inv['Email'] : '⚠️ ไม่พบ Investor (อาจไม่มี investorId หรือ ID ไม่ตรง)');
  }
  return 'ดูผลใน Execution Log';
}

function authorizeAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('✅ SpreadsheetApp OK — ' + ss.getName());
  var quota = MailApp.getRemainingDailyQuota();
  Logger.log('✅ MailApp OK — โควต้าเหลือ: ' + quota + ' อีเมล/วัน');
  var html = HtmlService.createHtmlOutput('<p>test</p>');
  Logger.log('✅ HtmlService OK');
  Logger.log('═══════════════════════════════════════');
  Logger.log('  ✅ Grant Permission เรียบร้อย!');
  Logger.log('  📧 อีเมลพร้อมส่ง (เหลือ ' + quota + ' ฉบับ/วัน)');
  Logger.log('═══════════════════════════════════════');
  return '✅ ทุก Permission พร้อม! Email quota: ' + quota;
}

function testEmail() {
  var result = sendAutomationEmail({ to: ADMIN_EMAIL, toName: 'Admin', subject: '[DEALPAY] 🧪 ทดสอบระบบอีเมล — ' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm'), body: emailTemplate_('🧪 ทดสอบระบบอีเมล', '<table style="width:100%;border-collapse:collapse;margin:12px 0"><tr><td style="padding:8px;color:#666;width:140px">สถานะ</td><td style="padding:8px;font-weight:700;color:#22c55e">✅ ทำงานปกติ</td></tr><tr><td style="padding:8px;color:#666">เวลา</td><td style="padding:8px">' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss') + '</td></tr><tr><td style="padding:8px;color:#666">โควต้า</td><td style="padding:8px">' + MailApp.getRemainingDailyQuota() + ' อีเมล/วัน</td></tr></table>'), relatedId: 'TEST', relatedType: 'System Test', template: 'Test Email', triggerEvent: 'Manual Test' });
  if (result) { Logger.log('✅ ส่งอีเมลทดสอบสำเร็จ → ' + ADMIN_EMAIL); } else { Logger.log('❌ ส่งอีเมลไม่สำเร็จ — เช็ค Permission!'); }
  return result ? '✅ ส่งอีเมลทดสอบสำเร็จ!' : '❌ ส่งไม่สำเร็จ — รัน authorizeAll() ก่อน';
}

function test_DirectEmail() {
  try {
    var result = sendAutomationEmail({ to: ADMIN_EMAIL, subject: "🧪 TEST 1: Direct Email Connection", body: "<h1>Connection Success!</h1><p>ถ้าเห็นเมลนี้ แสดงว่า Permission และ MailApp ปกติครับ</p>", relatedId: "TEST-001", relatedType: "Testing" });
    Logger.log(result ? "✅ TEST 1 สำเร็จ: ส่งเมลหา " + ADMIN_EMAIL + " แล้ว" : "❌ TEST 1 ล้มเหลว: ส่งเมลไม่ได้");
  } catch (e) { Logger.log("❌ TEST 1 ERROR: " + e.toString()); }
}

function test_UpdateStatusLogic() {
  var testWoId = "WO-2025-000001"; // <--- เปลี่ยนเป็น ID จริงในชีตของคุณ
  Logger.log("🔍 เริ่มทดสอบหา ID: " + testWoId);
  var result = updateCoordStatus(testWoId, "in_progress", "ทดสอบระบบส่งเมลอัตโนมัติ", "System Tester");
  if (result.saved && result.emailed) { Logger.log("✅ TEST 2 สำเร็จ: บันทึก Sheet และส่งเมลเรียบร้อย"); }
  else { Logger.log("⚠️ TEST 2 ผลลัพธ์: Saved=" + result.saved + " | Emailed=" + result.emailed); }
}

function test_SystemReady() {
  Logger.log("📧 Email Quota คงเหลือ: " + MailApp.getRemainingDailyQuota());
  Logger.log("📊 URL Spreadsheet: " + SpreadsheetApp.getActiveSpreadsheet().getUrl());
  var sheets = ["Investors", "CoordCases", "EmailLog"];
  sheets.forEach(function(name) {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
    Logger.log(sh ? "✅ ชีต " + name + ": พร้อมใช้งาน" : "❌ ไม่พบชีต: " + name);
  });
}


/* ═════════════════════════════════════════════════════
   🔧 DEBUG FUNCTIONS — รันใน Apps Script Editor
   ═════════════════════════════════════════════════════ */

/**
 * ทดสอบการบันทึก EmailLog โดยตรง
 * รันใน Apps Script Editor → เลือก debugEmailLog → กด Run
 */
function debugEmailLog() {
  Logger.log('=== DEBUG: EmailLog Test ===');
  
  // 1. เช็ค Sheet EmailLog มีอยู่ไหม
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('EmailLog');
  if (!sh) {
    Logger.log('❌ ERROR: Sheet "EmailLog" ไม่พบ! กรุณารัน setupSheets() ก่อน');
    return;
  }
  Logger.log('✅ Sheet EmailLog พบแล้ว | Rows: ' + sh.getLastRow());
  
  // 2. เช็ค Headers ใน Sheet
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  Logger.log('Headers ใน Sheet: ' + JSON.stringify(headers));
  
  // 3. Headers ที่ Code.gs คาดหวัง
  var expected = ['Email ID','Related ID','Related Type','To Email','To Name','Subject','Body Preview','Template','Sent Date','Status','Trigger Event','Auto-Generated'];
  Logger.log('Headers ที่ต้องการ: ' + JSON.stringify(expected));
  
  // 4. เช็คว่า headers ตรงกันไหม
  var mismatches = [];
  for (var i = 0; i < expected.length; i++) {
    if (headers[i] !== expected[i]) {
      mismatches.push('col ' + (i+1) + ': Sheet="' + headers[i] + '" vs Code="' + expected[i] + '"');
    }
  }
  if (mismatches.length > 0) {
    Logger.log('❌ MISMATCH พบ: ' + mismatches.join(' | '));
  } else {
    Logger.log('✅ Headers ตรงกันทั้งหมด');
  }
  
  // 5. ทดสอบ saveEmailLog โดยตรง
  Logger.log('--- ทดสอบ saveEmailLog ---');
  try {
    saveEmailLog({
      relatedId: 'DEBUG-001',
      relatedType: 'Debug',
      toEmail: 'test@test.com',
      toName: 'Debug Test',
      subject: '[DEBUG] ทดสอบ saveEmailLog',
      body: 'Debug test body',
      template: 'Debug',
      triggerEvent: 'Manual Debug',
      autoGenerated: true
    });
    Logger.log('✅ saveEmailLog สำเร็จ! ตรวจดูใน Sheet EmailLog');
  } catch(e) {
    Logger.log('❌ saveEmailLog FAILED: ' + e.toString());
  }
  
  // 6. ทดสอบ sendAutomationEmail (ไม่ส่งจริง แค่ดู logic)
  Logger.log('--- ทดสอบ appendRow_ EmailLog โดยตรง ---');
  try {
    appendRow_('EmailLog', {
      'Email ID': 9999,
      'Related ID': 'DEBUG-DIRECT',
      'Related Type': 'DirectTest',
      'To Email': 'direct@test.com',
      'To Name': 'Direct Test',
      'Subject': '[DEBUG] Direct appendRow test',
      'Body Preview': 'Direct test',
      'Template': 'DirectDebug',
      'Sent Date': new Date(),
      'Status': 'Sent',
      'Trigger Event': 'DirectDebug',
      'Auto-Generated': true
    });
    Logger.log('✅ appendRow_ EmailLog สำเร็จ');
  } catch(e) {
    Logger.log('❌ appendRow_ EmailLog FAILED: ' + e.toString());
  }
}

/**
 * ทดสอบส่งอีเมลจริง + บันทึก EmailLog
 * รัน debugFullEmailFlow() เพื่อดูว่า MailApp ทำงานได้ไหม
 */
function debugFullEmailFlow() {
  Logger.log('=== DEBUG: Full Email + Log Flow ===');
  var adminEmail = 'aungelcomes.deal@gmail.com';
  
  Logger.log('1. ทดสอบ MailApp.sendEmail...');
  var sent = false;
  try {
    MailApp.sendEmail({
      to: adminEmail,
      subject: '[DEALPAY DEBUG] ทดสอบระบบอีเมล ' + new Date().toISOString(),
      htmlBody: '<h2>🔧 Debug Email</h2><p>ถ้าเห็นอีเมลนี้แสดงว่า MailApp ทำงานปกติ</p><p>เวลา: ' + new Date() + '</p>'
    });
    sent = true;
    Logger.log('✅ MailApp.sendEmail สำเร็จ');
  } catch(e) {
    Logger.log('❌ MailApp.sendEmail FAILED: ' + e.toString());
    Logger.log('   → อาจต้องให้สิทธิ์ MailApp ก่อน ลองรัน authorizeAll()');
    return;
  }
  
  Logger.log('2. ทดสอบ saveEmailLog...');
  try {
    saveEmailLog({
      relatedId: 'DEBUG-FULL-001',
      relatedType: 'DebugFull',
      toEmail: adminEmail,
      toName: 'Admin',
      subject: '[DEALPAY DEBUG] ทดสอบ',
      body: 'Full flow debug test',
      template: 'DebugFull',
      triggerEvent: 'Manual Debug',
      autoGenerated: true
    });
    Logger.log('✅ saveEmailLog สำเร็จ — ตรวจดู Sheet EmailLog');
  } catch(e) {
    Logger.log('❌ saveEmailLog FAILED: ' + e.toString());
  }
  
  Logger.log('=== เสร็จ! เช็ค Sheet EmailLog และ inbox ===');
}

/**
 * แก้ headers ใน EmailLog ให้ตรงกับที่ Code.gs ต้องการ
 * รันเมื่อ debugEmailLog รายงานว่า headers ไม่ตรง
 */
function fixEmailLogHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('EmailLog');
  if (!sh) { Logger.log('❌ ไม่พบ Sheet EmailLog'); return; }
  var correct = ['Email ID','Related ID','Related Type','To Email','To Name','Subject','Body Preview','Template','Sent Date','Status','Trigger Event','Auto-Generated'];
  sh.getRange(1, 1, 1, correct.length).setValues([correct]);
  Logger.log('✅ แก้ headers EmailLog เรียบร้อย: ' + correct.join(', '));
}

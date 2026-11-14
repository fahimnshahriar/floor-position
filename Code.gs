/**
 * Dyeing Floor Tracker – Google Apps Script
 * Sheets used:
 *   "Live"            -> data the operators edit through the web app
 *   "Floor Position"  -> printable report (auto-filled from Live)
 * Run setup() once, then deploy as a Web app.
 */
const LIVE = 'Live';
const REP = 'Floor Position';

const MACHINES = [22]
  .concat(Array.from({length: 21}, (_, i) => i + 1), [23], [101, 102, 103, 104, 105, 106]);

const CAP = {22:1500,1:1500,2:1500,3:1200,4:1200,5:1000,6:1000,7:1000,8:750,9:750,10:750,
  11:500,12:300,13:250,14:250,15:400,16:180,17:500,18:500,19:180,20:100,21:100,23:1000,
  101:1500,102:900,103:750,104:500,105:400,106:200};

const POS = ['MC wash','Scouring','Scouring steam up','Scouring cooling','Scouring addition',
  'Scouring N/hot','Scouring acid/RTU','Enzyme','Enzyme wash','Levelling','Color dosing',
  'Color run time','Salt dosing','Salt run time','Migration','Soda dosing','1st soda dosing',
  '2nd soda dosing','Color steam','Add color dosing','BD wash','Soaping hot ST up',
  'Soaping hot wash','Softener','Fixing'];

const BUYERS = ['AMERICAN Tall','ARDON','AEON','AS COLOR','CAPRICON','CISALFA','COTTON HERITAGE',
  'ICE BERG','B&C','MATALAN','NEW YORKER','LPP','FASHION UK','PROMOSET','DMR','MALFINI',
  'CENTENLINE','ALL GOODS','SALLING','MN DYEING','CAMEL','DUCATI','KAUFMANN','RECOVARY FAB',
  'EL CORTE','MODYF','KSL','HENRY','KARIBAN','NUTMEG','WELLENSTEYN','MAYORAL','P&C','LIMA',
  'ANTHEM','ZXY','COMODO','JEANS','BOSTON','TG','SUM OF IN HOUSE','ALIB','APEX','APS','MANAMI',
  'FORTIS','FOUR H','AKH','HABITUS','HAMS','JM','KA DESIGN','IRISH DESIGN','SUPREME','LIZ',
  'FASHION ASIA','MS FABRICS','PENTAGON','RADIANCE','RAGADI','RHINO','SILKEN','MARB TEX','SQ',
  'AC MILAN','QUATTRO FASHION','GILDEN','HK-DREAM'];

/* ---------- Web app ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Dyeing Floor Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- One-time setup (run manually) ---------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  const n = MACHINES.length;

  // Live sheet
  const live = ss.getSheetByName(LIVE) || ss.insertSheet(LIVE);
  live.clear();
  live.getRange(1, 1, 1, 8).setValues([['MC','CAPACITY','BUYER','COLOR','LOAD TIME','POSITION','UPDATED AT','LOAD ISO']])
    .setFontWeight('bold').setBackground('#d9e2f3');
  live.getRange(2, 3, n, 6).setNumberFormat('@');
  live.getRange(2, 1, n, 2).setValues(MACHINES.map(m => [m, CAP[m]]));
  live.setFrozenRows(1);
  live.hideColumns(8);

  // Report sheet (same layout as the paper form)
  const r = ss.getSheetByName(REP) || ss.insertSheet(REP);
  r.getRange(1, 1, r.getMaxRows(), r.getMaxColumns()).breakApart();
  r.clear();
  r.getRange('A1:F1').merge().setValue('FLOOR POSITION');
  r.getRange('A2:B2').merge().setValue('DATE');
  r.getRange('C2').setFormula('=TODAY()').setNumberFormat('dd/mm/yyyy');
  r.getRange('E2').setValue('TIME');
  r.getRange('F2').setFormula('=NOW()').setNumberFormat('hh:mm');
  r.getRange('A3:F3').setValues([['MC','CAPACITY','BUYER','COLOR','LOAD TIME','POSITION']]);
  const rows = MACHINES.map((m, i) => {
    const k = i + 2;
    return [m, CAP[m]].concat(['C','D','E','F'].map(c => '=IF(Live!' + c + k + '="","",Live!' + c + k + ')'));
  });
  r.getRange(4, 1, n, 6).setFormulas(rows);
  const all = r.getRange(1, 1, 3 + n, 6);
  all.setBorder(true, true, true, true, true, true).setHorizontalAlignment('center').setVerticalAlignment('middle');
  r.getRange('A1:F3').setFontWeight('bold');
  r.getRange('A1:F1').setBackground('#d9e2f3').setFontSize(12);
  [60, 80, 180, 140, 120, 190].forEach((w, i) => r.setColumnWidth(i + 1, w));
  r.setFrozenRows(3);
  r.setActiveSelection('A1');
}

/* ---------- API used by the page ---------- */
function getBoot() {
  return {
    cap: CAP, pos: POS, buyers: BUYERS,
    u1: MACHINES.filter(m => m < 100), u2: MACHINES.filter(m => m >= 100),
    data: readLive_()
  };
}

function getData() { return readLive_(); }

function saveMachine(mc, buyer, color, loadIso, pos) {
  mc = Number(mc);
  buyer = String(buyer || '').trim();
  color = String(color || '').trim();
  const p = POS.filter(x => x.toLowerCase() === String(pos || '').trim().toLowerCase())[0];
  const idx = MACHINES.indexOf(mc);
  if (idx < 0) throw new Error('Unknown machine');
  if (!buyer) throw new Error('Buyer is required');
  if (!p) throw new Error('Pick a position from the list');

  const mon = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const m = /^(\d{4})-(\d\d)-(\d\d)T(\d\d:\d\d)/.exec(loadIso || '');
  const loadTxt = m ? (m[3] + ' ' + mon[Number(m[2]) - 1] + ' ' + m[4]) : '';
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM HH:mm');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    SpreadsheetApp.getActive().getSheetByName(LIVE)
      .getRange(idx + 2, 3, 1, 6).setValues([[buyer, color, loadTxt, p, now, loadIso || '']]);
  } finally { lock.releaseLock(); }
  return readLive_();
}

function clearMachine(mc) {
  const idx = MACHINES.indexOf(Number(mc));
  if (idx < 0) throw new Error('Unknown machine');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    SpreadsheetApp.getActive().getSheetByName(LIVE).getRange(idx + 2, 3, 1, 6).clearContent();
  } finally { lock.releaseLock(); }
  return readLive_();
}

function readLive_() {
  const s = SpreadsheetApp.getActive().getSheetByName(LIVE);
  const v = s.getRange(2, 1, MACHINES.length, 8).getValues();
  const o = {};
  v.forEach(r => {
    if (r[2] || r[5]) o[r[0]] = {buyer: r[2], color: r[3], loadTxt: r[4], pos: r[5], at: r[6], load: r[7]};
  });
  return o;
}

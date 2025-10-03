// === CONFIG ===
const SPREADSHEET_ID = 'PUT_YOUR_SHEET_ID_HERE'; // <-- new target spreadsheet
const SHEET_NAME = 'Orders';                     // <-- sheet/tab name with data

const WEBAPP_EXEC_BASE = 'https://script.google.com/macros/s/AKfycbwF-dEFJO1lJsPplWf7SO5U3JwG9dTrQ4pWBTLuxS8jVokDLyeVumrCIowqkfDqUmMBQQ/exec'; // new exec URL

const LEGACY_DB_SHEET_NAME = 'Base de Datos';
const ORDERS_SHEET_NAME = SHEET_NAME;

// If this script is bound to the sheet, getActive() will work.
// If it's standalone, set SPREADSHEET_ID to the target file ID.
const TARGET_SPREADSHEET_ID = (SPREADSHEET_ID && SPREADSHEET_ID !== 'PUT_YOUR_SHEET_ID_HERE') ? SPREADSHEET_ID : ''; // e.g. '1AbC...'; leave empty if container-bound

function getSpreadsheet_() {
  // Try container-bound first
  try {
    const ss = SpreadsheetApp.getActive();
    if (ss) return ss;
  } catch (e) {}

  // Fallback to explicit ID (standalone)
  if (!TARGET_SPREADSHEET_ID) {
    throw new Error(
      'No TARGET_SPREADSHEET_ID configured and script is not container-bound. ' +
      'Bind the script to the Sheet or set TARGET_SPREADSHEET_ID.'
    );
  }
  try {
    return SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  } catch (err) {
    throw new Error(
      'No tengo permiso para abrir la hoja destino. ' +
      'Comparte el archivo con la cuenta del despliegue (Execute as: Me) y ' +
      'autoriza el proyecto con el scope de spreadsheets.'
    );
  }
}

// Header names
const RECIBO_COL_HEADER = 'Recibo';

// Ensure a "Recibo" column exists and return its 1-based index
function ensureReciboColumn_(sh) {
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, Math.max(1, lastCol)).getValues()[0].map(v => String(v || '').trim());
  let idx = headers.indexOf(RECIBO_COL_HEADER);
  if (idx === -1) {
    const newCol = headers.length + 1;
    sh.getRange(1, newCol).setValue(RECIBO_COL_HEADER);
    idx = newCol - 1; // zero-based
  }
  return idx + 1; // 1-based
}

function buildReceiptUrl_(orderId) {
  if (!orderId) return '';
  return `${WEBAPP_EXEC_BASE}?id=${encodeURIComponent(orderId)}`;
}

function buildPosReceiptUrl_(orderId) {
  if (!orderId) return '';
  return `${WEBAPP_EXEC_BASE}?id=${encodeURIComponent(orderId)}&format=pos58`;
}

// Write a clickable receipt link (RichText) at rowIndex
function writeReceiptLink_(sh, rowIndex, orderId) {
  const col = ensureReciboColumn_(sh);
  const url = buildReceiptUrl_(orderId);
  if (!url) return;
  const rich = SpreadsheetApp.newRichTextValue()
    .setText(`Recibo (${orderId})`)
    .setLinkUrl(url)
    .build();
  sh.getRange(rowIndex, col).setRichTextValue(rich);
}

// ===================== REGION: AGENDA Infra =====================
const AGENDA_SHEETS = {
  paciente: { name: 'SC_Pacientes', headers: ['ID','Nombre','Tel','Email','Notas','CreatedAt'] },
  doctor:   { name: 'SC_Doctores',  headers: ['ID','Nombre','Tel','Email','Notas','CreatedAt'] },
};

function ensureAgenda_() {
  Object.values(AGENDA_SHEETS).forEach(def => {
    const sh = getOrCreateSheet_(def.name);
    ensureHeaders_(sh, def.headers);
  });

  const sh = getDbSheet_();
  const headers = sh.getRange(1,1,1, Math.max(1, sh.getLastColumn())).getValues()[0];
  ['PacienteID','DoctorID'].forEach(col => {
    if (headers.indexOf(col) === -1) {
      sh.getRange(1, headers.length + 1).setValue(col);
      headers.push(col);
    }
  });
}

function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureHeaders_(sh, headers) {
  const lastCol = Math.max(1, sh.getLastColumn());
  const cur = sh.getRange(1,1,1,lastCol).getValues()[0];
  if (cur.join('') === '') {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    return;
  }
  headers.forEach(h => {
    if (cur.indexOf(h) === -1) {
      sh.getRange(1, cur.length + 1).setValue(h);
      cur.push(h);
    }
  });
}

/** Ensure a sheet exists; if created now and headers are provided, write them. 
 *  If the sheet exists, ensure headers array items exist (append to the right if missing).
 *  Returns {sheet, created, headersAdded: string[]}
 */
function ensureSheetWithHeaders_(name, headers) {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(name);
  let created = false;
  if (!sh) {
    sh = ss.insertSheet(name);
    created = true;
  }
  const result = { sheet: sh, created: created, headersAdded: [] };
  if (headers && headers.length) {
    const lastCol = Math.max(1, sh.getLastColumn());
    let cur = sh.getRange(1,1,1,lastCol).getValues()[0].map(v => String(v||'').trim());
    // If the header row is empty, write all at once
    if (cur.join('') === '') {
      sh.getRange(1,1,1,headers.length).setValues([headers]);
      return result;
    }
    // Append missing headers at the end
    headers.forEach(h => {
      if (cur.indexOf(h) === -1) {
        sh.getRange(1, cur.length + 1).setValue(h);
        cur.push(h);
        result.headersAdded.push(h);
      }
    });
  }
  return result;
}

/** Ensure ONLY the doctor agenda sheet exists (no patient sheet needed). */
function ensureDoctorAgenda_() {
  const def = { name: 'SC_Doctores', headers: ['ID','Nombre','Tel','Email','Notas','CreatedAt'] };
  return ensureSheetWithHeaders_(def.name, def.headers);
}

function ensureDoctorDirectory_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName('SC_Doctores');
  if (!sh) sh = ss.insertSheet('SC_Doctores');
  const headers = ['Doctor','Teléfono','Última Orden','Actualizado'];
  const lastCol = Math.max(1, sh.getLastColumn());
  let cur = sh.getRange(1,1,1,lastCol).getValues()[0].map(v => String(v||'').trim());
  if (cur.join('') === '') {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  } else {
    headers.forEach(h => {
      if (!cur.includes(h)) {
        sh.getRange(1, cur.length + 1).setValue(h);
        cur.push(h);
      }
    });
  }
  return sh;
}

function upsertDoctorDirectory_(doctorNombre, doctorTel, orderId) {
  if (!doctorNombre) return;
  const sh = ensureDoctorDirectory_();
  const headers = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0];
  const idx = Object.fromEntries(headers.map((h,i)=>[h,i]));
  const last = sh.getLastRow();
  let foundRow = -1;
  for (let r=2; r<=last; r++) {
    const val = String(sh.getRange(r, idx['Doctor']+1).getValue() || '').trim();
    if (val && val.toLowerCase() === doctorNombre.toLowerCase()) { foundRow = r; break; }
  }
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  if (foundRow === -1) {
    const row = new Array(headers.length).fill('');
    row[idx['Doctor']]        = doctorNombre;
    row[idx['Teléfono']]      = doctorTel || '';
    row[idx['Última Orden']]  = orderId || '';
    row[idx['Actualizado']]   = now;
    sh.appendRow(row);
  } else {
    if (doctorTel) sh.getRange(foundRow, idx['Teléfono']+1).setValue(doctorTel);
    if (orderId)  sh.getRange(foundRow, idx['Última Orden']+1).setValue(orderId);
    sh.getRange(foundRow, idx['Actualizado']+1).setValue(now);
  }
}

function normalizeName_(s) {
  return String(s || '')
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/\s+/g,' ')
    .trim();
}

function getEntitySheet_(tipo) {
  const def = AGENDA_SHEETS[tipo];
  if (!def) throw new Error('Tipo inválido: ' + tipo);
  return getOrCreateSheet_(def.name);
}

function generateEntityId_(tipo) {
  const prefix = (tipo === 'paciente') ? 'PAC' : 'DOC';
  const ymd = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  const seq = Utilities.getUuid().split('-')[0].toUpperCase();
  return `${prefix}-${ymd}-${seq}`;
}

function getEntityByName_(tipo, nombre) {
  const sh = getEntitySheet_(tipo);
  const headers = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0];
  const idx = Object.fromEntries(headers.map((h,i)=>[h,i]));
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  const data = sh.getRange(2,1,lastRow-1, headers.length).getValues();
  const q = normalizeName_(nombre);
  for (let i=0;i<data.length;i++){
    const r = data[i];
    if (normalizeName_(String(r[idx['Nombre']]||'')) === q) {
      return {
        id: r[idx['ID']],
        nombre: r[idx['Nombre']],
        tel: r[idx['Tel']] || '',
        email: r[idx['Email']] || '',
        notas: r[idx['Notas']] || ''
      };
    }
  }
  return null;
}

function upsertEntityByName_(tipo, nombre, tel, email, notas) {
  if (!nombre) throw new Error('Nombre requerido');
  const sh = getEntitySheet_(tipo);
  const headers = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0];
  const idx = Object.fromEntries(headers.map((h,i)=>[h,i]));

  const lastRow = sh.getLastRow();
  const data = lastRow >= 2 ? sh.getRange(2,1,lastRow-1,headers.length).getValues() : [];
  const normTarget = normalizeName_(nombre);

  let rowIndex = -1;
  for (let i=0;i<data.length;i++){
    const row = data[i];
    const name = String(row[idx['Nombre']] || '');
    if (normalizeName_(name) === normTarget) { rowIndex = i+2; break; }
  }

  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  if (rowIndex > -1) {
    const curTel   = String(sh.getRange(rowIndex, idx['Tel']+1).getValue() || '');
    const curEmail = String(sh.getRange(rowIndex, idx['Email']+1).getValue() || '');
    const curNotas = String(sh.getRange(rowIndex, idx['Notas']+1).getValue() || '');
    if (!curTel   && tel)   sh.getRange(rowIndex, idx['Tel']+1).setValue(tel);
    if (!curEmail && email) sh.getRange(rowIndex, idx['Email']+1).setValue(email);
    if (!curNotas && notas) sh.getRange(rowIndex, idx['Notas']+1).setValue(notas);
    const id = sh.getRange(rowIndex, idx['ID']+1).getValue();
    return { id, nombre: nombre.trim(), existed: true };
  } else {
    const id = generateEntityId_(tipo);
    const row = new Array(headers.length).fill('');
    row[idx['ID']]        = id;
    row[idx['Nombre']]    = nombre.trim();
    row[idx['Tel']]       = tel || '';
    row[idx['Email']]     = email || '';
    row[idx['Notas']]     = notas || '';
    row[idx['CreatedAt']] = now;
    sh.appendRow(row);
    return { id, nombre: nombre.trim(), existed: false };
  }
}

function searchEntities_(tipo, query, limit) {
  const q = normalizeName_(query || '');
  const sh = getEntitySheet_(tipo);
  const headers = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0];
  const idx = Object.fromEntries(headers.map((h,i)=>[h,i]));
  const lastRow = sh.getLastRow();
  const data = lastRow >= 2 ? sh.getRange(2,1,lastRow-1,headers.length).getValues() : [];

  const res = [];
  for (let i=0;i<data.length;i++){
    const row = data[i];
    const name = String(row[idx['Nombre']]||'');
    if (!q || normalizeName_(name).indexOf(q) !== -1) {
      res.push({
        id: row[idx['ID']],
        nombre: name,
        tel: row[idx['Tel']] || '',
        email: row[idx['Email']] || '',
        notas: row[idx['Notas']] || ''
      });
      if (limit && res.length >= limit) break;
    }
  }
  return res;
}

function apiSearchAgenda(tipo, query, limit) {
  ensureAgenda_();
  return searchEntities_(tipo, query, limit || 8);
}
function apiGetAgendaByName(tipo, nombre) {
  ensureAgenda_();
  return getEntityByName_(tipo, nombre);
}
function apiEnsureAgenda(tipo, nombre, tel, email, notas) {
  ensureAgenda_();
  return upsertEntityByName_(tipo, nombre, tel, email, notas);
}

function resolveAgendaBeforeSave_(record) {
  ensureAgenda_();

  const pacNombre = (record.pacienteNombre || '').trim();
  const docNombre = (record.doctorNombre || '').trim();

  if (pacNombre) {
    const exists = getEntityByName_('paciente', pacNombre);
    if (exists) {
      record.pacienteId = exists.id;
    } else {
      const pac = upsertEntityByName_('paciente', pacNombre, record.pacienteTel || '', record.pacienteEmail || '', record.pacienteNotas || '');
      record.pacienteId = pac.id;
      record._newPatient = true;
    }
  }
  if (docNombre) {
    const exists = getEntityByName_('doctor', docNombre);
    if (exists) {
      record.doctorId = exists.id;
      record._doctorTel = exists.tel || record.doctorTel || '';
    } else {
      const doc = upsertEntityByName_('doctor', docNombre, record.doctorTel || '', record.doctorEmail || '', record.doctorNotas || '');
      record.doctorId = doc.id;
      record._newDoctor = true;
      record._doctorTel = record.doctorTel || '';
    }
  }
  return record;
}

function menu_BackfillAgendaLinks() {
  ensureAgenda_();
  const sh = getDbSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const headers = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0];
  const idx = Object.fromEntries(headers.map((h,i)=>[h,i]));

  ['ID Orden','Paciente','Doctor','PacienteID','DoctorID'].forEach(n=>{
    if (idx[n] == null) throw new Error('Falta columna: ' + n);
  });

  const rng = sh.getRange(2,1,lastRow-1, headers.length).getValues();
  for (let i=0;i<rng.length;i++){
    const row = rng[i];
    const pacTxt = String(row[idx['Paciente']]||'').trim();
    let docTxt = String(row[idx['Doctor']]||'').trim();
    if (!docTxt) {
      const firstIdx = idx['Nombre'];
      const lastIdx = idx['Apellido'];
      const first = firstIdx != null ? String(row[firstIdx] || '').trim() : '';
      const last = lastIdx != null ? String(row[lastIdx] || '').trim() : '';
      docTxt = [first, last].filter(Boolean).join(' ');
    }
    let pId = row[idx['PacienteID']];
    let dId = row[idx['DoctorID']];

    if (pacTxt && !pId) pId = upsertEntityByName_('paciente', pacTxt).id;
    if (docTxt && !dId) dId = upsertEntityByName_('doctor',   docTxt).id;

    if (pId) sh.getRange(i+2, idx['PacienteID']+1).setValue(pId);
    if (dId) sh.getRange(i+2, idx['DoctorID']+1).setValue(dId);
  }
  SpreadsheetApp.getUi().alert('Backfill de Agenda completado.');
}

// -- Search doctors by name (for autosuggest)
function apiSearchDoctors(query, limit) {
  ensureDoctorAgenda_(); // ensure SC_Doctores exists and has headers
  return searchEntities_('doctor', query, limit || 8);
}

// -- Add/Update a doctor (invoked by "Agregar doctor" button)
function apiAddDoctor(nombre, tel) {
  ensureDoctorAgenda_();
  const res = upsertEntityByName_('doctor', String(nombre || '').trim(), String(tel || '').trim(), '', '');
  return res; // {id, nombre, existed}
}
// ===================== END REGION: AGENDA Infra =====================

// ===================== REGION: Notifications (Feature Flag OFF) =====================
const FEATURE_NOTIFY = false;

function buildOrderNotificationPayload_(record, orderId) {
  return {
    event: 'new_order',
    orderId: orderId,
    orderUrl: (typeof buildReceiptUrl_ === 'function')
      ? buildReceiptUrl_(orderId)
      : ('?id=' + encodeURIComponent(orderId)),
    timestamp: new Date().toISOString(),
    patient: { id: record.pacienteId || '', name: record.pacienteNombre || '' },
    doctor:  { id: record.doctorId || '',  name: record.doctorNombre || '', phone: record._doctorTel || record.doctorTel || '' },
    meta:    { sheetName: getDbSheet_().getName(), user: (Session.getActiveUser().getEmail() || '') }
  };
}

function notifyOnSave_(record, orderId) {
  const payload = buildOrderNotificationPayload_(record, orderId);
  if (!FEATURE_NOTIFY) {
    console.log('[Notify disabled] Would send payload:', JSON.stringify(payload));
    return;
  }
  // When enabling later, replace with UrlFetchApp.fetch(...) to n8n webhook
}
// ===================== END REGION: Notifications =====================

function getDbSheet_() {
  ensureSchemaBaseDeOrdenes();
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(ORDERS_SHEET_NAME);
  if (!sh) {
    throw new Error(`Sheet "${SHEET_NAME}" not found`);
  }
  return sh;
}

function ensureDbSheetSetup_(sheet) {
  ensureSchemaBaseDeOrdenes();
  return sheet || getDbSheet_();
}

function ensureSchemaBaseDeOrdenes() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(ORDERS_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(ORDERS_SHEET_NAME);

  // Single source of truth for headers (order matters)
  const headers = [
    'ID Orden','Timestamp','Nombre','Apellido','Teléfono','Paciente','Fecha Requerida',
    'Tipo de trabajo','Material','Especificación','Dientes Seleccionados','Color Global','Colores por Diente',
    'Terminado','Puentes','Odontograma JSON','Costo','A Cuenta','Total','Garantía','Estado','Notas',
    'Recibo',
    'Doctor','Doctor Tel',
    'PacienteID','DoctorID'
  ];

  const neededCols = headers.length;
  if (sh.getMaxColumns() < neededCols) {
    sh.insertColumnsAfter(sh.getMaxColumns(), neededCols - sh.getMaxColumns());
  }

  const current = sh.getRange(1,1,1,neededCols).getValues()[0];
  let rewrite = false;
  for (let i=0;i<headers.length;i++) {
    if ((current[i] || '').toString().trim() !== headers[i]) { rewrite = true; break; }
  }
  if (rewrite) sh.getRange(1,1,1,neededCols).setValues([headers]);

  // Ensure Recibo column exists and Estado DV/formatting
  ensureReciboColumn_(sh);
  setupEstadoColumn_(sh);
}

/** Run all setup steps needed for a clean install/migration. Idempotent. */
function ensureAllSetup_() {
  // Orders schema, Estado DV/format, Recibo column
  ensureSchemaBaseDeOrdenes();
  // Doctor directory only
  ensureDoctorAgenda_();
}

function verifyInstallation() {
  const logs = [];
  // Orders
  try {
    const ss = getSpreadsheet_();
    let sh = ss.getSheetByName(ORDERS_SHEET_NAME);
    const before = !!sh;
    ensureSchemaBaseDeOrdenes();
    sh = ss.getSheetByName(ORDERS_SHEET_NAME);
    logs.push(`Orders: ${before ? 'OK' : 'CREATED'} (${ORDERS_SHEET_NAME})`);
    // Verify "Recibo" column exists
    const col = ensureReciboColumn_(sh);
    logs.push(` - Recibo column at index ${col}`);
    logs.push(' - Estado validation/formatting: OK');
  } catch (e) {
    logs.push('Orders: ERROR - ' + (e && e.message));
  }

  // SC_Doctores
  try {
    const res = ensureDoctorAgenda_();
    logs.push(`SC_Doctores: ${res.created ? 'CREATED' : 'OK'} (headers added: ${res.headersAdded.join(', ') || 'none'})`);
  } catch (e) {
    logs.push('SC_Doctores: ERROR - ' + (e && e.message));
  }

  SpreadsheetApp.getUi().alert('Verificar instalación\n\n' + logs.join('\n'));
}

function setupEstadoColumn_(sheet) {
  const ESTADO_COL = 21; // U
  const numRows = Math.max(1, sheet.getMaxRows() - 1);
  const rangeCol = sheet.getRange(2, ESTADO_COL, numRows, 1);

  const estados = ['Pendiente', 'En proceso', 'Terminado', 'Entregado'];
  const dv = SpreadsheetApp.newDataValidation()
    .requireValueInList(estados, true)
    .setAllowInvalid(false)
    .build();
  rangeCol.setDataValidation(dv);

  const estadoRange = sheet.getRange(2, ESTADO_COL, numRows, 1);
  const rules = sheet.getConditionalFormatRules()
    .filter(r => !r.getRanges().some(gr => gr.getSheet().getSheetId() === sheet.getSheetId() && gr.getColumn() === ESTADO_COL && gr.getNumColumns() === 1));
  const makeRule = (text, bg) => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(text)
    .setBackground(bg)
    .setRanges([estadoRange])
    .build();
  rules.push(
    makeRule('Pendiente', '#FDE2E2'),
    makeRule('En proceso', '#FFE5CC'),
    makeRule('Terminado', '#FFF4CC'),
    makeRule('Entregado', '#E6F4EA')
  );
  sheet.setConditionalFormatRules(rules);
}

function fixEstadoColumn() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheetByName(LEGACY_DB_SHEET_NAME) || ss.getActiveSheet();
  setupEstadoColumn_(sheet);
}

function columnLetterToNumber_(letter) {
  let number = 0;
  const upper = String(letter || '').toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    const code = upper.charCodeAt(i);
    if (code < 65 || code > 90) continue;
    number = number * 26 + (code - 64);
  }
  return number;
}

function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const id     = String(params.id || '').trim();
  const format = String(params.format || 'a4').toLowerCase();
  if (id) {
    const {record, view} = getRecordAndViewById(id);
    if (!record) {
      return HtmlService.createHtmlOutput(`<p style="font:14px Arial">No se encontró la orden con id <b>${id}</b>.</p>`);
    }
    const templateName = (format === 'pos58') ? 'recibo_pos58' : 'recibo';
    const t = HtmlService.createTemplateFromFile(templateName);
    t.record = record;
    t.view   = view;
    t.format = format;
    t.execBase = WEBAPP_EXEC_BASE; // <-- pass exec base to template
    const out = t.evaluate().setTitle('Recibo de Orden');
    out.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    return out;
  }

  const t = HtmlService.createTemplateFromFile('sidebar');
  t.APP_CONTEXT = 'webapp'; // lets CSS adapt for web vs sidebar
  const out = t.evaluate().setTitle('Smile Center · Orden');
  out.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return out;
}

function mostrarSidebar() {
  const t = HtmlService.createTemplateFromFile('sidebar');
  t.APP_CONTEXT = 'sidebar'; // Sheets sidebar
  const html = t
    .evaluate()
    .setTitle('Smile Center — Orden de Laboratorio')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  SpreadsheetApp.getUi().showSidebar(html);
}

function onOpen() {
  try { ensureAllSetup_(); } catch (e) { console.log('ensureAllSetup_ error:', e && e.message); }
  SpreadsheetApp.getUi()
    .createMenu('Smile Center')
    .addItem('Abrir Orden', 'mostrarSidebar')
    .addSeparator()
    .addItem('Verificar instalación', 'verifyInstallation')
    .addToUi();
}

function getRecordAndViewById(orderId) {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) return {record: null, view: null};

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(String);
  const idxId = headers.indexOf('ID Orden');
  if (idxId < 0) return {record: null, view: null};

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idxId]).trim() === orderId) {
      const rowObj = {};
      headers.forEach((h, i) => rowObj[h] = values[r][i]);

      let selectedTeeth = [];
      let perTooth      = {};
      let bridges       = [];
      let globalShade   = '';

      try { globalShade = String(rowObj['Color Global'] || ''); } catch (e) {}
      try { perTooth = rowObj['Colores por Diente'] ? JSON.parse(rowObj['Colores por Diente']) : {}; } catch (e) {}
      try {
        const rawP = rowObj['Puentes'];
        if (rawP) {
          bridges = Array.isArray(rawP) ? rawP
            : (typeof rawP === 'string' && rawP.trim().startsWith('[')) ? JSON.parse(rawP)
            : parseFlatBridges(rawP);
        }
      } catch (e) {}

      try {
        const rawSel = rowObj['Dientes Seleccionados'];
        if (Array.isArray(rawSel)) {
          selectedTeeth = rawSel;
        } else if (typeof rawSel === 'string') {
          const trimmed = rawSel.trim();
          if (trimmed.startsWith('[')) {
            selectedTeeth = JSON.parse(trimmed);
          } else {
            let toSplit = trimmed;
            if (trimmed.startsWith('(') && trimmed.endsWith(')')) {
              toSplit = trimmed.slice(1, -1);
            }
            selectedTeeth = toSplit.split(',').map(function(s) {
              return s.trim();
            }).filter(function(v) { return v; });
          }
        }
      } catch (e) {}

      const view = { selectedTeeth, perTooth, bridges, globalShade };
      return {record: rowObj, view};
    }
  }
  return {record: null, view: null};
}

function parseFlatBridges(s) {
  return String(s || '')
    .split(',')
    .map(part => part.trim())
    .filter(Boolean)
    .map(run => {
      const [teethPart, colorPart] = run.split(':');
      const dientes = (teethPart || '').split('-').map(t => t.trim()).filter(Boolean);
      const color   = (colorPart || '').trim();
      return dientes.length >= 2 ? { dientes, color } : null;
    })
    .filter(Boolean);
}

function toListString_(value) {
  if (Array.isArray(value)) {
    return value.map(item => String(item || '').trim()).filter(Boolean).join(', ');
  }
  return String(value || '').trim();
}

function normalizeTeethValue_(raw) {
  if (!raw) return raw;
  var s = String(raw).trim();

  if (s.startsWith('(') && s.endsWith(')')) {
    var inside = s.slice(1, -1).trim();
    if (inside.includes(',')) return s;
    if (/^\d+$/.test(inside)) return `(${inside},)`;
    return s;
  }

  if (/^\d+$/.test(s)) return `(${s},)`;

  var parts = s.split(',').map(function(x) { return x.trim(); }).filter(Boolean);
  if (parts.length === 1 && /^\d+$/.test(parts[0])) {
    return `(${parts[0]},)`;
  }

  return s;
}

function safeString_(value) {
  return String(value || '').trim();
}

function safeNumber_(value) {
  if (value === null || typeof value === 'undefined' || value === '') return 0;
  const normalized = String(value).replace(/\s+/g, '').replace(',', '.');
  const num = parseFloat(normalized);
  return Number.isFinite(num) ? num : 0;
}

function safeParseJson_(s, fallback) {
  try { return s ? JSON.parse(String(s)) : fallback; } catch (e) { return fallback; }
}

function toDataUrl_(url) {
  try {
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (res.getResponseCode() !== 200) return '';
    var ct = res.getHeaders()['Content-Type'] || 'image/png';
    var b64 = Utilities.base64Encode(res.getBlob().getBytes());
    return 'data:' + ct + ';base64,' + b64;
  } catch (e) {
    return '';
  }
}

function generateNextId_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return 'ORD-0001';
  }
  const idValues = sh.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(v => String(v || '').trim() !== '');
  let maxNumber = 0;
  idValues.forEach(id => {
    const match = String(id).match(/^ORD-(\d{4,})$/);
    if (match) {
      const numeric = parseInt(match[1], 10);
      if (numeric > maxNumber) maxNumber = numeric;
    }
  });
  const next = (maxNumber + 1).toString().padStart(4, '0');
  return `ORD-${next}`;
}

function saveOrder(record) {
  if (!record) {
    throw new Error('No llegaron datos');
  }

  record = record || {};
  record.pacienteNombre = safeString_(record.pacienteNombre || record.patientName || '');
  record.doctorNombre   = safeString_(record.doctorNombre || '');      // from sidebar
  record.doctorTel      = safeString_(record.doctorTel || '');         // from sidebar
  record.dentistNombre  = safeString_(record.dentistNombre || '');
  record.dentistApellido= safeString_(record.dentistApellido || '');
  record.dentistNumero  = safeString_(record.dentistNumero || '');
  record.pacienteTel = record.pacienteTel || record.patientTel || '';

  // If the optional doctor fields are empty, mirror from required dentist fields
  if (!record.doctorNombre) {
    record.doctorNombre = [record.dentistNombre, record.dentistApellido].filter(Boolean).join(' ');
  }
  if (!record.doctorTel) {
    record.doctorTel = record.dentistNumero;
  }

  const sh = getDbSheet_();
  ensureDbSheetSetup_(sh);

  const material = toListString_(record.material);
  const especificacion = toListString_(record.especificacion);
  const tipoTrabajo = toListString_(record.tipoTrabajo);
  const dientesSeleccionados = normalizeTeethValue_(toListString_(record.selectedTeeth));

  const camposObligatorios = [
    [safeString_(record.dentistNombre), 'Nombre del dentista'],
    [safeString_(record.dentistApellido), 'Apellido del dentista']
  ];

  const faltantes = camposObligatorios.filter(([valor]) => !valor).map(([, etiqueta]) => etiqueta);
  if (faltantes.length) {
    throw new Error(`Campos obligatorios pendientes: ${faltantes.join(', ')}.`);
  }

  const costo = safeNumber_(record.costo);
  const aCuenta = safeNumber_(record.aCuenta);
  if (aCuenta > costo) {
    throw new Error('El anticipo no puede ser mayor que el costo.');
  }

  let fechaRequerida = '';
  if (record.fechaRequerida) {
    const fecha = new Date(record.fechaRequerida);
    if (Number.isNaN(fecha.getTime())) {
      throw new Error('La fecha requerida no es válida.');
    }
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    fecha.setHours(0, 0, 0, 0);
    if (fecha < hoy) {
      throw new Error('La fecha requerida debe ser hoy o una fecha futura.');
    }
    fechaRequerida = fecha;
  }

  const total = (() => {
    if (record.totalCalculado === null || typeof record.totalCalculado === 'undefined' || record.totalCalculado === '') {
      return costo - aCuenta;
    }
    const normalized = String(record.totalCalculado).replace(/\s+/g, '').replace(',', '.');
    const parsed = parseFloat(normalized);
    return Number.isFinite(parsed) ? parsed : (costo - aCuenta);
  })();
  const garantia = safeString_(record.garantia) === 'Sí' ? 'Sí' : 'No';

  // Agenda deshabilitada; no se resuelven entidades automáticamente.

  const timestamp = new Date();
  const orderId = record.id || generateNextId_(sh);
  record.id = orderId;
  const coloresPorDiente = safeString_(record.coloresPorDientePlano) || safeString_(record.coloresPorDiente);
  const puentes = safeString_(record.puentesPlano) || safeString_(record.puentes);
  const odontogramaJson = safeString_(record.odontogramaJson);

  const row = [
    orderId,
    timestamp,
    record.dentistNombre,
    record.dentistApellido,
    record.dentistNumero,
    record.pacienteNombre,
    fechaRequerida,
    tipoTrabajo,
    material,
    especificacion,
    dientesSeleccionados,
    safeString_(record.colorGlobal),
    coloresPorDiente,
    record.terminado ? 'Sí' : 'No',
    puentes,
    odontogramaJson,
    costo,
    aCuenta,
    total,
    garantia,
    'Pendiente',
    safeString_(record.notas),
    '',                         // Recibo (link filled after)
    record.doctorNombre,        // <-- Doctor
    record.doctorTel,           // <-- Doctor Tel
    record.pacienteId || '',
    record.doctorId || ''
  ];

  sh.appendRow(row);

  upsertDoctorDirectory_(record.doctorNombre, record.doctorTel, orderId);

  const newRowIndex = sh.getLastRow();
  writeReceiptLink_(sh, newRowIndex, orderId);

  notifyOnSave_(record, orderId);

  return orderId;
}

function guardarOrdenDesdeSidebar(datos) {
  const orderId = saveOrder(datos);
  const sh = getDbSheet_();
  const folio = Math.max(0, sh.getLastRow() - 1);
  return `Orden ${orderId} guardada con folio #${folio}.`;
}

// ⬇️ NUEVO: render de recibo imprimible
function generarReciboHtml(orderId) {
  const cleanId = String(orderId || '').trim();
  if (!cleanId) {
    return HtmlService.createHtmlOutput("<p style='font-family:sans-serif;color:#b00020;'>Error: El identificador de la orden es obligatorio.</p>");
  }

  const {record, view} = getRecordAndViewById(cleanId);
  if (!record) {
    return HtmlService.createHtmlOutput(`<p style='font-family:sans-serif;color:#b00020;'>Error: No se encontró la orden ${cleanId}.</p>`);
  }

  const template = HtmlService.createTemplateFromFile('recibo');
  template.record = record;
  template.view = view;
  template.format = 'a4';
  return template.evaluate()
    .setTitle(`Recibo ${cleanId}`)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function backfillReceiptLinks() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(ORDERS_SHEET_NAME);
  if (!sh) throw new Error(`No existe la hoja "${ORDERS_SHEET_NAME}"`);

  // ensure "Recibo" header exists or get its index
  const reciboCol = ensureReciboColumn_(sh);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  // Column A = "ID Orden"
  const ids = sh.getRange(2, 1, lastRow - 1, 1).getValues().map(r => String(r[0] || '').trim());
  const existing = sh.getRange(2, reciboCol, lastRow - 1, 1).getRichTextValues();

  ids.forEach((id, i) => {
    if (!id) return;
    const url = buildReceiptUrl_(id);
    const cur = existing[i][0];
    if (cur && cur.getLinkUrl && cur.getLinkUrl() === url) return;
    const rich = SpreadsheetApp.newRichTextValue()
      .setText(`Recibo (${id})`)
      .setLinkUrl(url)
      .build();
    sh.getRange(i + 2, reciboCol).setRichTextValue(rich);
  });
}

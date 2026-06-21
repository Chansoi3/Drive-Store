/**
 * Drive-Store Backend (Google Apps Script)
 *
 * Setup:
 *   1. Open script: https://script.google.com/d/1apw7_QPsqtnwA4EW9VibhtPVmQRFBSPHBg-kVOemhcaE1rFCTDPGLL3A/edit
 *   2. Paste this file over Code.gs.
 *   3. Run `setup` once to create sheets, default stores, branch folders, and the hourly cleanup trigger.
 *   4. Deploy > Manage deployments > pencil > New version > Deploy.
 */

var ROOT_FOLDER_ID = '1w5T4wQOzSEIJmSqdWcNexjXmS0S_aUOz';
var STATUS_SS_ID   = '1YUM50Hltg8fcY30MScA8lBVg-ZgYhCJmn4dzSVMfMKo';
var ADMIN_CODE     = 'admin2026';

var CATEGORIES = ['ราคาสินค้า', 'POP', 'แจ้งเตือน', 'เอกสารบัญชี', 'โปรโมชั่น', 'อื่นๆ'];

var SH_STATUS  = 'DownloadStatus';
var SH_META    = 'FileMeta';
var SH_HISTORY = 'UploadHistory';
var SH_STORES  = 'Stores';

var DEFAULT_STORES = {
  '0730': 'ซอย เพี้ยนพิณ 2',
  '3100': 'เฉลิมพระเกียรติ 14 แยก 34',
  '3108': 'สุภาพงษ์3แยก4',
  '3109': 'ซอย อ่อนนุช 36',
  '3110': 'บางนาตราด กม. 6.5 (ปั๊มน้ำมัน)',
  '3118': 'ซอยเทพารักษ์ 8',
  '3121': 'บางนา - ตราด กม.8 (ปั๊มน้ำมัน)',
  '5209': 'ซอย กุศลศิลป์',
  '5219': 'เฉลิมพระเกียรติ 30 แยก 14',
  '5285': 'แฟลตทหารเรือบางนา',
  '5291': 'บ้านเทพารักษ์',
  '6731': 'ท็อปส์ เดลี่ คาลเท็กซ์ ศรีนครินทร์',
  '6953': 'ท็อปส์ เดลี่ บางจากบางนาตราด ก.ม.7'
};

/* ===========================
   SETUP
   =========================== */
function setup() {
  sheet_(SH_STATUS,  ['branchCode','fileId','downloaded','downloadedAt']);
  sheet_(SH_META,    ['fileId','groupId','category','note','expiresAt','createdAt']);
  sheet_(SH_HISTORY, ['groupId','uploadedAt','filename','mimeType','size','category','note','expiresAt','branches','fileIds']);
  var stores = sheet_(SH_STORES, ['code','name']);
  if (stores.getLastRow() < 2) {
    var rows = Object.keys(DEFAULT_STORES).map(function(k){ return [k, DEFAULT_STORES[k]]; });
    stores.getRange(2, 1, rows.length, 2).setValues(rows);
  }
  Object.keys(DEFAULT_STORES).forEach(function(code){ getOrCreateBranchFolder_(code, DEFAULT_STORES[code]); });
  installTriggers_();
  return { ok: true };
}

function installTriggers_() {
  var fns = ScriptApp.getProjectTriggers().map(function(t){ return t.getHandlerFunction(); });
  if (fns.indexOf('cleanupExpired') === -1) {
    ScriptApp.newTrigger('cleanupExpired').timeBased().everyHours(1).create();
  }
}

/* ===========================
   HTTP
   =========================== */
function doGet(e)  { return route_(e); }
function doPost(e) { return route_(e); }

function route_(e) {
  try {
    var params = (e && e.parameter) || {};
    var body = {};
    if (e && e.postData && e.postData.contents) {
      try { body = JSON.parse(e.postData.contents); } catch (err) { body = {}; }
    }
    var action = params.action || body.action || '';
    var p = function(k){ return body[k] != null ? body[k] : params[k]; };
    var data;
    switch (action) {
      case 'config':         data = apiConfig_();                                                 break;
      case 'stores':         data = apiStores_();                                                 break;
      case 'login':          data = apiLogin_(p('code'));                                         break;
      case 'list':           data = apiList_(p('code'));                                          break;
      case 'markDownloaded': data = apiMark_(p('code'), p('fileId'));                             break;
      case 'upload':         data = apiUpload_(body);                                             break;
      case 'delete':         data = apiDelete_(body);                                             break;
      case 'rename':         data = apiRename_(body);                                             break;
      case 'updateMeta':     data = apiUpdateMeta_(body);                                         break;
      case 'resend':         data = apiResend_(body);                                             break;
      case 'history':        data = apiHistory_(body);                                            break;
      case 'dashboard':      data = apiDashboard_(body);                                          break;
      case 'zip':            data = apiZip_(body);                                                break;
      case 'cleanup':        data = cleanupExpired();                                             break;
      case 'storeSave':      data = apiStoreSave_(body);                                          break;
      case 'storeDelete':    data = apiStoreDelete_(body);                                        break;
      default:               data = { ok: false, error: 'Unknown action: ' + action };
    }
    return json_(data);
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ===========================
   API
   =========================== */
function apiConfig_() { return { ok: true, categories: CATEGORIES }; }

function apiStores_() {
  var s = loadStores_();
  var list = Object.keys(s).map(function(k){ return { code: k, name: s[k] }; });
  list.sort(function(a,b){ return a.code.localeCompare(b.code); });
  return { ok: true, stores: list };
}

function apiLogin_(code) {
  if (!code) return { ok: false, error: 'Missing code' };
  if (code === ADMIN_CODE) return { ok: true, role: 'admin', name: 'Administrator' };
  var stores = loadStores_();
  if (stores[code]) return { ok: true, role: 'branch', code: code, name: stores[code] };
  return { ok: false, error: 'รหัสไม่ถูกต้อง' };
}

function apiList_(code) {
  var stores = loadStores_();
  if (!code || !stores[code]) return { ok: false, error: 'รหัสสาขาไม่ถูกต้อง' };
  var folder = getOrCreateBranchFolder_(code);
  if (!folder) return { ok: false, error: 'ไม่พบโฟลเดอร์สาขา' };
  var iter = folder.getFiles();
  var status = readStatusForBranch_(code);
  var meta = loadMeta_();
  var out = [];
  var now = new Date();
  while (iter.hasNext()) {
    var f = iter.next();
    var id = f.getId();
    var m = meta[id] || {};
    if (m.expiresAt && new Date(m.expiresAt) < now) continue;
    var st = status[id] || {};
    out.push({
      id: id,
      name: f.getName(),
      size: f.getSize(),
      mimeType: f.getMimeType(),
      uploadedAt: f.getDateCreated().toISOString(),
      downloadUrl: 'https://drive.google.com/uc?export=download&id=' + id,
      viewUrl: f.getUrl(),
      downloaded: !!st.downloaded,
      downloadedAt: st.downloadedAt || null,
      category: m.category || '',
      note: m.note || '',
      expiresAt: m.expiresAt || null,
      groupId: m.groupId || ''
    });
  }
  out.sort(function(a,b){ return b.uploadedAt.localeCompare(a.uploadedAt); });
  return { ok: true, branch: { code: code, name: stores[code] }, files: out };
}

function apiMark_(code, fileId) {
  var stores = loadStores_();
  if (!code || !stores[code]) return { ok: false, error: 'รหัสสาขาไม่ถูกต้อง' };
  if (!fileId) return { ok: false, error: 'Missing fileId' };
  setStatus_(code, fileId, true);
  return { ok: true };
}

function apiUpload_(body) {
  if (!body || body.adminCode !== ADMIN_CODE) return { ok: false, error: 'Unauthorized' };
  var name = body.filename || 'file';
  var mime = body.mimeType || 'application/octet-stream';
  var b64 = body.base64 || '';
  var branches = body.branches || [];
  var category = body.category || '';
  var note = body.note || '';
  var expiresAt = body.expiresAt || null;
  if (!b64) return { ok: false, error: 'No file data' };
  if (!branches.length) return { ok: false, error: 'No branches selected' };
  var stores = loadStores_();
  var bytes = Utilities.base64Decode(b64);
  var blob = Utilities.newBlob(bytes, mime, name);
  var groupId = Utilities.getUuid();
  var fileIds = [];
  var validBranches = [];
  for (var i = 0; i < branches.length; i++) {
    var code = branches[i];
    if (!stores[code]) continue;
    var folder = getOrCreateBranchFolder_(code);
    var f = folder.createFile(blob);
    try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    setMeta_(f.getId(), groupId, category, note, expiresAt);
    fileIds.push(f.getId());
    validBranches.push(code);
  }
  var hist = sheet_(SH_HISTORY, []);
  hist.appendRow([
    groupId, new Date(), name, mime, bytes.length,
    category, note, expiresAt ? new Date(expiresAt) : '',
    validBranches.join(','), fileIds.join(',')
  ]);
  return { ok: true, groupId: groupId, fileIds: fileIds, branches: validBranches };
}

function apiDelete_(body) {
  if (body.adminCode !== ADMIN_CODE) return { ok: false, error: 'Unauthorized' };
  if (body.groupId) return deleteGroup_(body.groupId);
  if (!body.fileId) return { ok: false, error: 'Missing target' };
  try { DriveApp.getFileById(body.fileId).setTrashed(true); } catch (e) {}
  removeMetaRow_(body.fileId);
  return { ok: true };
}

function apiRename_(body) {
  if (body.adminCode !== ADMIN_CODE) return { ok: false, error: 'Unauthorized' };
  var newName = body.name;
  if (!newName) return { ok: false, error: 'Missing name' };
  if (body.groupId) {
    var ids = getGroupFileIds_(body.groupId);
    ids.forEach(function(id){ try { DriveApp.getFileById(id).setName(newName); } catch(e) {} });
    updateHistoryRow_(body.groupId, function(row){ row[2] = newName; });
    return { ok: true, renamed: ids.length };
  }
  if (body.fileId) { DriveApp.getFileById(body.fileId).setName(newName); return { ok: true }; }
  return { ok: false, error: 'Missing target' };
}

function apiUpdateMeta_(body) {
  if (body.adminCode !== ADMIN_CODE) return { ok: false, error: 'Unauthorized' };
  if (!body.groupId) return { ok: false, error: 'Missing groupId' };
  var category = body.category || '';
  var note = body.note || '';
  var expiresAt = body.expiresAt || null;
  var ids = getGroupFileIds_(body.groupId);
  ids.forEach(function(id){ setMeta_(id, body.groupId, category, note, expiresAt); });
  updateHistoryRow_(body.groupId, function(row){
    row[5] = category; row[6] = note;
    row[7] = expiresAt ? new Date(expiresAt) : '';
  });
  return { ok: true };
}

function apiResend_(body) {
  if (body.adminCode !== ADMIN_CODE) return { ok: false, error: 'Unauthorized' };
  if (!body.groupId || !(body.branches || []).length)
    return { ok: false, error: 'Missing params' };
  var stores = loadStores_();
  var hist = sheet_(SH_HISTORY, []);
  var last = hist.getLastRow();
  if (last < 2) return { ok: false, error: 'no history' };
  var rng = hist.getRange(2, 1, last-1, 10);
  var rows = rng.getValues();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0] !== body.groupId) continue;
    var existingBranches = String(rows[i][8] || '').split(',').filter(Boolean);
    var existingIds = String(rows[i][9] || '').split(',').filter(Boolean);
    if (!existingIds.length) return { ok: false, error: 'no source' };
    var sourceFile = DriveApp.getFileById(existingIds[0]);
    var newBranches = [], newIds = [];
    for (var b = 0; b < body.branches.length; b++) {
      var code = body.branches[b];
      if (!stores[code] || existingBranches.indexOf(code) !== -1) continue;
      var folder = getOrCreateBranchFolder_(code);
      var copy = sourceFile.makeCopy(sourceFile.getName(), folder);
      try { copy.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
      setMeta_(copy.getId(), body.groupId, rows[i][5], rows[i][6], rows[i][7]);
      newBranches.push(code);
      newIds.push(copy.getId());
    }
    var allB = existingBranches.concat(newBranches);
    var allI = existingIds.concat(newIds);
    hist.getRange(i+2, 9, 1, 2).setValues([[allB.join(','), allI.join(',')]]);
    return { ok: true, added: newBranches };
  }
  return { ok: false, error: 'groupId not found' };
}

function apiHistory_(body) {
  if (body.adminCode !== ADMIN_CODE) return { ok: false, error: 'Unauthorized' };
  var hist = sheet_(SH_HISTORY, []);
  var last = hist.getLastRow();
  if (last < 2) return { ok: true, items: [] };
  var rows = hist.getRange(2, 1, last-1, 10).getValues();
  var items = rows.map(function(r){
    return {
      groupId: r[0],
      uploadedAt: r[1] ? new Date(r[1]).toISOString() : null,
      filename: r[2],
      mimeType: r[3],
      size: r[4],
      category: r[5],
      note: r[6],
      expiresAt: r[7] ? new Date(r[7]).toISOString() : null,
      branches: String(r[8] || '').split(',').filter(Boolean),
      fileIds: String(r[9] || '').split(',').filter(Boolean)
    };
  });
  items.sort(function(a,b){ return (b.uploadedAt || '').localeCompare(a.uploadedAt || ''); });
  return { ok: true, items: items };
}

function apiDashboard_(body) {
  if (body.adminCode !== ADMIN_CODE) return { ok: false, error: 'Unauthorized' };
  var h = apiHistory_(body);
  if (!h.ok) return h;
  var stores = loadStores_();
  var storeKeys = Object.keys(stores).sort();
  var sh = sheet_(SH_STATUS, []);
  var last = sh.getLastRow();
  var status = {};
  if (last >= 2) {
    var rows = sh.getRange(2, 1, last-1, 4).getValues();
    rows.forEach(function(r){ status[r[0]+'|'+r[1]] = !!r[2]; });
  }
  var matrix = h.items.map(function(it){
    var cells = {};
    storeKeys.forEach(function(code){ cells[code] = { status: 'na' }; });
    for (var i = 0; i < it.branches.length; i++) {
      var code = it.branches[i], id = it.fileIds[i];
      cells[code] = { fileId: id, status: status[code+'|'+id] ? 'done' : 'pending' };
    }
    return {
      groupId: it.groupId, uploadedAt: it.uploadedAt,
      filename: it.filename, category: it.category, cells: cells
    };
  });
  return {
    ok: true,
    stores: storeKeys.map(function(k){ return { code: k, name: stores[k] }; }),
    rows: matrix
  };
}

function apiZip_(body) {
  var code = body.code;
  var fileIds = body.fileIds || [];
  var stores = loadStores_();
  if (!code || !stores[code]) return { ok: false, error: 'Invalid branch' };
  if (!fileIds.length) return { ok: false, error: 'No files' };
  var blobs = [];
  var total = 0;
  for (var i = 0; i < fileIds.length; i++) {
    var f = DriveApp.getFileById(fileIds[i]);
    var blob = f.getBlob();
    total += blob.getBytes().length;
    if (total > 80 * 1024 * 1024) return { ok: false, error: 'รวมเกิน 80 MB' };
    blobs.push(blob);
  }
  var zipBlob = Utilities.zip(blobs, code + '-bundle-' + new Date().getTime() + '.zip');
  var folder = getOrCreateBranchFolder_(code);
  var zipFile = folder.createFile(zipBlob);
  try { zipFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  // expire the zip in 1 hour
  setMeta_(zipFile.getId(), '', 'อื่นๆ', '(bundle ชั่วคราว)', new Date(Date.now() + 60*60*1000).toISOString());
  return { ok: true, url: 'https://drive.google.com/uc?export=download&id=' + zipFile.getId(), fileId: zipFile.getId() };
}

function cleanupExpired() {
  var meta = sheet_(SH_META, []);
  var last = meta.getLastRow();
  if (last < 2) return { ok: true, removed: 0 };
  var rows = meta.getRange(2, 1, last-1, 6).getValues();
  var now = new Date();
  var removed = 0;
  for (var i = rows.length - 1; i >= 0; i--) {
    var exp = rows[i][4];
    if (exp && new Date(exp) < now) {
      try { DriveApp.getFileById(String(rows[i][0])).setTrashed(true); removed++; } catch (e) {}
      meta.deleteRow(i + 2);
    }
  }
  return { ok: true, removed: removed };
}

function apiStoreSave_(body) {
  if (body.adminCode !== ADMIN_CODE) return { ok: false, error: 'Unauthorized' };
  var code = String(body.code || '').trim();
  var name = String(body.name || '').trim();
  if (!code || !name) return { ok: false, error: 'Missing code or name' };
  var sh = sheet_(SH_STORES, ['code','name']);
  var last = sh.getLastRow();
  if (last >= 2) {
    var rng = sh.getRange(2, 1, last-1, 2).getValues();
    for (var i = 0; i < rng.length; i++) {
      if (String(rng[i][0]) === code) {
        sh.getRange(i+2, 2).setValue(name);
        getOrCreateBranchFolder_(code, name);
        return { ok: true, updated: true };
      }
    }
  }
  sh.appendRow([code, name]);
  getOrCreateBranchFolder_(code, name);
  return { ok: true, created: true };
}

function apiStoreDelete_(body) {
  if (body.adminCode !== ADMIN_CODE) return { ok: false, error: 'Unauthorized' };
  var code = String(body.code || '');
  if (!code) return { ok: false, error: 'Missing code' };
  var sh = sheet_(SH_STORES, ['code','name']);
  var last = sh.getLastRow();
  if (last < 2) return { ok: false, error: 'Empty' };
  var rng = sh.getRange(2, 1, last-1, 2).getValues();
  for (var i = 0; i < rng.length; i++) {
    if (String(rng[i][0]) === code) { sh.deleteRow(i+2); return { ok: true }; }
  }
  return { ok: false, error: 'Not found' };
}

/* ===========================
   Drive helpers
   =========================== */
function getRoot_() { return DriveApp.getFolderById(ROOT_FOLDER_ID); }

function getOrCreateBranchFolder_(code, name) {
  var n = name || loadStores_()[code];
  if (!n) return null;
  var folderName = code + ' ' + n;
  var root = getRoot_();
  var it = root.getFoldersByName(folderName);
  if (it.hasNext()) return it.next();
  return root.createFolder(folderName);
}

/* ===========================
   Sheet helpers
   =========================== */
function sheet_(name, headers) {
  var ss = SpreadsheetApp.openById(STATUS_SS_ID);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers && headers.length) { sh.appendRow(headers); sh.setFrozenRows(1); }
  } else if (sh.getLastRow() === 0 && headers && headers.length) {
    sh.appendRow(headers); sh.setFrozenRows(1);
  }
  return sh;
}

function loadStores_() {
  var sh = sheet_(SH_STORES, ['code','name']);
  var last = sh.getLastRow();
  if (last < 2) return Object.assign({}, DEFAULT_STORES);
  var rows = sh.getRange(2, 1, last-1, 2).getValues();
  var out = {};
  rows.forEach(function(r){ if (r[0]) out[String(r[0])] = String(r[1] || ''); });
  return out;
}

function loadMeta_() {
  var sh = sheet_(SH_META, ['fileId','groupId','category','note','expiresAt','createdAt']);
  var last = sh.getLastRow();
  if (last < 2) return {};
  var rows = sh.getRange(2, 1, last-1, 6).getValues();
  var out = {};
  rows.forEach(function(r){
    if (r[0]) out[String(r[0])] = {
      groupId: r[1] || '',
      category: r[2] || '',
      note: r[3] || '',
      expiresAt: r[4] ? new Date(r[4]).toISOString() : null,
      createdAt: r[5] ? new Date(r[5]).toISOString() : null
    };
  });
  return out;
}

function setMeta_(fileId, groupId, category, note, expiresAt) {
  var sh = sheet_(SH_META, ['fileId','groupId','category','note','expiresAt','createdAt']);
  var last = sh.getLastRow();
  var exp = expiresAt ? new Date(expiresAt) : '';
  var now = new Date();
  if (last >= 2) {
    var rng = sh.getRange(2, 1, last-1, 1).getValues();
    for (var i = 0; i < rng.length; i++) {
      if (String(rng[i][0]) === fileId) {
        sh.getRange(i+2, 1, 1, 6).setValues([[fileId, groupId, category, note, exp, now]]);
        return;
      }
    }
  }
  sh.appendRow([fileId, groupId, category || '', note || '', exp, now]);
}

function removeMetaRow_(fileId) {
  var sh = sheet_(SH_META, []);
  var last = sh.getLastRow();
  if (last < 2) return;
  var rng = sh.getRange(2, 1, last-1, 1).getValues();
  for (var i = 0; i < rng.length; i++) {
    if (String(rng[i][0]) === fileId) { sh.deleteRow(i+2); return; }
  }
}

function readStatusForBranch_(code) {
  var sh = sheet_(SH_STATUS, ['branchCode','fileId','downloaded','downloadedAt']);
  var last = sh.getLastRow();
  if (last < 2) return {};
  var rows = sh.getRange(2, 1, last-1, 4).getValues();
  var out = {};
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0] === code) out[rows[i][1]] = { downloaded: !!rows[i][2], downloadedAt: rows[i][3] || null };
  }
  return out;
}

function setStatus_(code, fileId, downloaded) {
  var sh = sheet_(SH_STATUS, ['branchCode','fileId','downloaded','downloadedAt']);
  var last = sh.getLastRow();
  var when = new Date().toISOString();
  if (last >= 2) {
    var rng = sh.getRange(2, 1, last-1, 2).getValues();
    for (var i = 0; i < rng.length; i++) {
      if (rng[i][0] === code && rng[i][1] === fileId) {
        sh.getRange(i+2, 3, 1, 2).setValues([[downloaded, when]]);
        return;
      }
    }
  }
  sh.appendRow([code, fileId, downloaded, when]);
}

function getGroupFileIds_(groupId) {
  var hist = sheet_(SH_HISTORY, []);
  var last = hist.getLastRow();
  if (last < 2) return [];
  var rows = hist.getRange(2, 1, last-1, 10).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0] === groupId) return String(rows[i][9] || '').split(',').filter(Boolean);
  }
  return [];
}

function updateHistoryRow_(groupId, mutator) {
  var hist = sheet_(SH_HISTORY, []);
  var last = hist.getLastRow();
  if (last < 2) return;
  var rng = hist.getRange(2, 1, last-1, 10);
  var rows = rng.getValues();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0] === groupId) { mutator(rows[i]); hist.getRange(i+2, 1, 1, 10).setValues([rows[i]]); return; }
  }
}

function deleteGroup_(groupId) {
  var hist = sheet_(SH_HISTORY, []);
  var last = hist.getLastRow();
  if (last < 2) return { ok: false, error: 'empty' };
  var rows = hist.getRange(2, 1, last-1, 10).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0] === groupId) {
      var ids = String(rows[i][9] || '').split(',').filter(Boolean);
      ids.forEach(function(id){
        try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {}
        removeMetaRow_(id);
      });
      hist.deleteRow(i+2);
      return { ok: true, removed: ids.length };
    }
  }
  return { ok: false, error: 'not found' };
}

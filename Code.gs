/**
 * Drive-Store Backend (Google Apps Script)
 *
 * Deploy:
 *   1. Open script: https://script.google.com/d/1apw7_QPsqtnwA4EW9VibhtPVmQRFBSPHBg-kVOemhcaE1rFCTDPGLL3A/edit
 *   2. Replace Code.gs with this file.
 *   3. Deploy > New deployment > Type: Web app
 *        - Execute as: Me
 *        - Who has access: Anyone
 *   4. Copy the /exec URL and paste it into index.html as API_URL.
 */

var ROOT_FOLDER_ID = '1w5T4wQOzSEIJmSqdWcNexjXmS0S_aUOz';
var ADMIN_CODE     = 'admin2026';
var STATUS_SHEET_NAME = 'DownloadStatus';

var STORES = {
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

/* ---------- One-time setup ----------
 * เปิด Apps Script editor → เลือกฟังก์ชัน setup → กด Run ครั้งเดียว
 * จะสร้าง Sheet "Drive-Store Status" และโฟลเดอร์ทั้ง 13 สาขาให้ครบ
 */
function setup() {
  var sh = getStatusSheet_();
  var ss = sh.getParent();
  var folders = [];
  Object.keys(STORES).forEach(function (code) {
    var f = getOrCreateBranchFolder_(code);
    folders.push(code + ' → ' + f.getName());
  });
  Logger.log('Sheet: ' + ss.getUrl());
  Logger.log('Folders created/verified:\n' + folders.join('\n'));
  return { sheetUrl: ss.getUrl(), folders: folders };
}

/* ---------- HTTP ---------- */

function doGet(e) {
  return route_(e);
}

function doPost(e) {
  return route_(e);
}

function route_(e) {
  try {
    var params = (e && e.parameter) || {};
    var body = {};
    if (e && e.postData && e.postData.contents) {
      try { body = JSON.parse(e.postData.contents); } catch (err) { body = {}; }
    }
    var action = params.action || body.action || '';
    var data;
    switch (action) {
      case 'stores':         data = apiStores_();                              break;
      case 'login':          data = apiLogin_(params.code || body.code);        break;
      case 'list':           data = apiList_(params.code || body.code);         break;
      case 'markDownloaded': data = apiMark_(params.code || body.code,
                                             params.fileId || body.fileId);    break;
      case 'upload':         data = apiUpload_(body);                           break;
      default:               data = { ok: false, error: 'Unknown action' };
    }
    return json_(data);
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- API ---------- */

function apiStores_() {
  var list = [];
  Object.keys(STORES).forEach(function (k) { list.push({ code: k, name: STORES[k] }); });
  return { ok: true, stores: list };
}

function apiLogin_(code) {
  if (!code) return { ok: false, error: 'Missing code' };
  if (code === ADMIN_CODE) return { ok: true, role: 'admin', name: 'Administrator' };
  if (STORES[code]) return { ok: true, role: 'branch', code: code, name: STORES[code] };
  return { ok: false, error: 'รหัสไม่ถูกต้อง' };
}

function apiList_(code) {
  if (!code || !STORES[code]) return { ok: false, error: 'รหัสสาขาไม่ถูกต้อง' };
  var folder = getOrCreateBranchFolder_(code);
  var iter = folder.getFiles();
  var status = readStatusForBranch_(code);
  var out = [];
  while (iter.hasNext()) {
    var f = iter.next();
    var id = f.getId();
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
      downloadedAt: st.downloadedAt || null
    });
  }
  out.sort(function (a, b) { return b.uploadedAt.localeCompare(a.uploadedAt); });
  return { ok: true, branch: { code: code, name: STORES[code] }, files: out };
}

function apiMark_(code, fileId) {
  if (!code || !STORES[code]) return { ok: false, error: 'รหัสสาขาไม่ถูกต้อง' };
  if (!fileId) return { ok: false, error: 'Missing fileId' };
  setStatus_(code, fileId, true);
  return { ok: true };
}

function apiUpload_(body) {
  if (!body || body.adminCode !== ADMIN_CODE) return { ok: false, error: 'Unauthorized' };
  var name = body.filename || 'file';
  var mime = body.mimeType || 'application/octet-stream';
  var b64  = body.base64 || '';
  var branches = body.branches || [];
  if (!b64) return { ok: false, error: 'No file data' };
  if (!branches.length) return { ok: false, error: 'No branches selected' };

  var bytes = Utilities.base64Decode(b64);
  var blob  = Utilities.newBlob(bytes, mime, name);
  var created = [];
  for (var i = 0; i < branches.length; i++) {
    var code = branches[i];
    if (!STORES[code]) continue;
    var folder = getOrCreateBranchFolder_(code);
    var f = folder.createFile(blob);
    created.push({ branch: code, fileId: f.getId(), name: f.getName() });
  }
  return { ok: true, created: created };
}

/* ---------- Drive helpers ---------- */

function getRoot_() {
  return DriveApp.getFolderById(ROOT_FOLDER_ID);
}

function getOrCreateBranchFolder_(code) {
  var folderName = code + ' ' + STORES[code];
  var root = getRoot_();
  var it = root.getFoldersByName(folderName);
  if (it.hasNext()) return it.next();
  return root.createFolder(folderName);
}

/* ---------- Status sheet helpers ---------- */

function getStatusSheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('STATUS_SS_ID');
  var ss;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('Drive-Store Status');
    props.setProperty('STATUS_SS_ID', ss.getId());
    try { DriveApp.getFileById(ss.getId()).moveTo(getRoot_()); } catch (e) {}
  }
  var sh = ss.getSheetByName(STATUS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(STATUS_SHEET_NAME);
    sh.appendRow(['branchCode', 'fileId', 'downloaded', 'downloadedAt']);
  }
  return sh;
}

function readStatusForBranch_(code) {
  var sh = getStatusSheet_();
  var last = sh.getLastRow();
  if (last < 2) return {};
  var rows = sh.getRange(2, 1, last - 1, 4).getValues();
  var out = {};
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0] === code) {
      out[rows[i][1]] = { downloaded: !!rows[i][2], downloadedAt: rows[i][3] || null };
    }
  }
  return out;
}

function setStatus_(code, fileId, downloaded) {
  var sh = getStatusSheet_();
  var last = sh.getLastRow();
  var when = new Date().toISOString();
  if (last >= 2) {
    var rng = sh.getRange(2, 1, last - 1, 2).getValues();
    for (var i = 0; i < rng.length; i++) {
      if (rng[i][0] === code && rng[i][1] === fileId) {
        sh.getRange(i + 2, 3, 1, 2).setValues([[downloaded, when]]);
        return;
      }
    }
  }
  sh.appendRow([code, fileId, downloaded, when]);
}

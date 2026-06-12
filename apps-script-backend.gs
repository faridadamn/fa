const CONFIG = {
  spreadsheetId: '', // kosongkan kalau script terikat langsung ke Google Sheet
  sessionTtlHours: 24 * 14,
  sheets: {
    products: 'Products',
    articles: 'Articles',
    transactions: 'Transactions',
    admins: 'Admins',
    profile: 'Profile',
    payment: 'Payment'
  }
};

const HEADERS = {
  products: ['id','storeId','name','cat','desc','price','img','link','createdAt','updatedAt'],
  articles: ['id','storeId','title','topic','excerpt','body','img','status','createdAt','updatedAt'],
  transactions: ['transactionId','id','storeId','productId','productName','price','buyerName','buyerWa','buyerEmail','note','status','createdAt','updatedAt'],
  admins: ['adminId','storeId','name','email','wa','ig','linkedin','role','passwordHash','sessionToken','sessionExpires','createdAt','updatedAt'],
  profile: ['storeId','name','tagline','ig','wa','linkedin','avatar','updatedAt'],
  payment: ['storeId','bank','ewallet','qris','updatedAt']
};

function doGet(e) {
  const p = e.parameter || {};
  const action = p.action || 'list';
  const resource = p.resource || 'articles';
  const storeId = normalizeStoreId_(p.storeId || 'farid-adam');
  const isPublic = String(p.public || '') === '1';
  const callback = p.callback;

  try {
    if (action === 'register') return output_({ ok:true, session: registerAdmin_(p, storeId) }, callback);
    if (action === 'login') return output_({ ok:true, session: loginAdmin_(p, storeId) }, callback);
    if (action === 'session') return output_({ ok:true, session: requireAdmin_(storeId, p.adminId, p.sessionToken) }, callback);

    if (action !== 'list') return output_({ ok:false, error:'Unsupported action' }, callback);
    if (!HEADERS[resource]) throw new Error('Unknown resource: ' + resource);
    if (!isPublic) requireAdmin_(storeId, p.adminId, p.sessionToken);
    if (isPublic && (resource === 'admins' || resource === 'transactions')) throw new Error('Resource is not public');

    const data = listResource_(resource, storeId, isPublic);
    return output_({ ok:true, [resource]: data }, callback);
  } catch (err) {
    return output_({ ok:false, error:String(err && err.message || err) }, callback);
  }
}

function doPost(e) {
  try {
    const data = JSON.parse((e.postData && e.postData.contents) || '{}');
    const action = data.action;
    const resource = data.resource;
    const storeId = normalizeStoreId_(data.storeId || 'farid-adam');

    if (!resource || !HEADERS[resource]) throw new Error('Unknown resource: ' + resource);
    if (!(resource === 'transactions' && action === 'create')) {
      requireAdmin_(storeId, data.adminId, data.sessionToken);
    }

    if (resource === 'products' && action === 'replace') {
      replaceStoreRows_(resource, storeId, data.products || []);
    } else if (resource === 'articles' && action === 'upsert') {
      upsertRow_(resource, storeId, data.article || {}, 'id');
    } else if (resource === 'articles' && action === 'delete') {
      deleteRow_(resource, storeId, (data.article || {}).id, 'id');
    } else if (resource === 'transactions' && action === 'create') {
      appendRow_(resource, storeId, Object.assign({ status:'paid_unverified' }, data.transaction || {}));
    } else if (resource === 'admins' && action === 'upsert') {
      upsertAdminProfile_(storeId, data.admin || {});
    } else if (resource === 'profile' && action === 'upsert') {
      upsertRow_(resource, storeId, data.profile || {}, 'storeId');
    } else if (resource === 'payment' && action === 'upsert') {
      upsertRow_(resource, storeId, data.payment || {}, 'storeId');
    } else {
      throw new Error('Unsupported action/resource: ' + action + '/' + resource);
    }

    return output_({ ok:true });
  } catch (err) {
    return output_({ ok:false, error:String(err && err.message || err) });
  }
}

function registerAdmin_(p, storeId) {
  const email = normalizeEmail_(p.email);
  const passwordHash = String(p.passwordHash || '');
  const name = String(p.adminName || p.name || '').trim();
  const storeName = String(p.storeName || storeId).trim();
  if (!storeId || !email || !passwordHash || !name) throw new Error('Data register belum lengkap');
  if (findAdminByEmail_(storeId, email)) throw new Error('Admin untuk toko ini sudah terdaftar');

  const now = new Date().toISOString();
  const sessionToken = makeToken_();
  const sessionExpires = expiresAt_();
  const admin = {
    adminId: makeAdminId_(storeId, email),
    storeId,
    name,
    email,
    wa: String(p.wa || '').trim(),
    ig: '',
    linkedin: '',
    role: 'owner',
    passwordHash,
    sessionToken,
    sessionExpires,
    createdAt: now,
    updatedAt: now
  };
  appendRow_('admins', storeId, admin);
  upsertRow_('profile', storeId, { storeId, name: storeName, wa: admin.wa, updatedAt: now }, 'storeId');
  return publicSession_(admin, storeName);
}

function loginAdmin_(p, storeId) {
  const email = normalizeEmail_(p.email);
  const passwordHash = String(p.passwordHash || '');
  const found = findAdminByEmail_(storeId, email);
  if (!found || String(found.admin.passwordHash || '') !== passwordHash) throw new Error('Store ID, email, atau password salah');

  const sessionToken = makeToken_();
  const sessionExpires = expiresAt_();
  const updated = Object.assign({}, found.admin, { sessionToken, sessionExpires, updatedAt: new Date().toISOString() });
  writeRow_('admins', found.rowIndex, updated);
  return publicSession_(updated, getStoreName_(storeId));
}

function requireAdmin_(storeId, adminId, sessionToken) {
  const found = findAdminById_(storeId, adminId);
  if (!found) throw new Error('Sesi admin tidak ditemukan');
  const admin = found.admin;
  if (!sessionToken || String(admin.sessionToken || '') !== String(sessionToken)) throw new Error('Sesi admin tidak valid');
  if (admin.sessionExpires && new Date(admin.sessionExpires).getTime() < Date.now()) throw new Error('Sesi admin sudah berakhir');
  return publicSession_(admin, getStoreName_(storeId));
}

function ss_() {
  return CONFIG.spreadsheetId
    ? SpreadsheetApp.openById(CONFIG.spreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(resource) {
  const ss = ss_();
  const name = CONFIG.sheets[resource];
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  ensureHeaders_(sh, HEADERS[resource]);
  return sh;
}

function ensureHeaders_(sh, headers) {
  const lastCol = Math.max(sh.getLastColumn(), headers.length);
  const current = sh.getRange(1, 1, 1, lastCol).getValues()[0].filter(Boolean);
  if (!current.length) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }
  const merged = current.slice();
  headers.forEach(h => { if (merged.indexOf(h) === -1) merged.push(h); });
  sh.getRange(1, 1, 1, merged.length).setValues([merged]);
}

function listResource_(resource, storeId, isPublic) {
  const sh = sheet_(resource);
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return resource === 'profile' || resource === 'payment' ? {} : [];

  const headers = values[0];
  const rows = values.slice(1).filter(r => String(r[headers.indexOf('storeId')] || '') === String(storeId));
  let objects = rows.map(r => rowToObject_(headers, r));
  if (isPublic) objects = objects.map(compactPublicRow_);

  if (resource === 'profile' || resource === 'payment') return objects[0] || {};
  return objects;
}

function replaceStoreRows_(resource, storeId, items) {
  const sh = sheet_(resource);
  const headers = sh.getDataRange().getValues()[0];
  deleteStoreRows_(sh, headers, storeId);
  (items || []).forEach(item => appendRow_(resource, storeId, item));
}

function appendRow_(resource, storeId, item) {
  const sh = sheet_(resource);
  const headers = sh.getDataRange().getValues()[0];
  const now = new Date().toISOString();
  const rowObj = Object.assign({}, item, {
    storeId,
    createdAt: item.createdAt || now,
    updatedAt: item.updatedAt || now
  });
  sh.appendRow(headers.map(h => rowObj[h] == null ? '' : rowObj[h]));
}

function upsertRow_(resource, storeId, item, keyField) {
  const sh = sheet_(resource);
  const headers = sh.getDataRange().getValues()[0];
  const key = item[keyField] || (keyField === 'storeId' ? storeId : '');
  if (!key) throw new Error('Missing key: ' + keyField);

  const values = sh.getDataRange().getValues();
  const storeIdx = headers.indexOf('storeId');
  const keyIdx = headers.indexOf(keyField);
  let rowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][storeIdx]) === String(storeId) && String(values[i][keyIdx]) === String(key)) {
      rowIndex = i + 1;
      break;
    }
  }

  const existing = rowIndex > -1 ? rowToObject_(headers, sh.getRange(rowIndex, 1, 1, headers.length).getValues()[0]) : {};
  const now = new Date().toISOString();
  const rowObj = Object.assign({}, existing, item, { storeId, [keyField]: key, updatedAt: now });
  if (!rowObj.createdAt && headers.indexOf('createdAt') > -1) rowObj.createdAt = now;
  const row = headers.map(h => rowObj[h] == null ? '' : rowObj[h]);

  if (rowIndex > -1) sh.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
  else sh.appendRow(row);
}

function writeRow_(resource, rowIndex, item) {
  const sh = sheet_(resource);
  const headers = sh.getDataRange().getValues()[0];
  sh.getRange(rowIndex, 1, 1, headers.length).setValues([headers.map(h => item[h] == null ? '' : item[h])]);
}

function upsertAdminProfile_(storeId, item) {
  const found = findAdminById_(storeId, item.adminId);
  if (!found) throw new Error('Admin tidak ditemukan');
  const existing = found.admin;
  const updated = Object.assign({}, existing, item, {
    storeId,
    adminId: existing.adminId,
    email: existing.email,
    passwordHash: existing.passwordHash,
    sessionToken: existing.sessionToken,
    sessionExpires: existing.sessionExpires,
    role: existing.role || item.role || 'owner',
    updatedAt: new Date().toISOString()
  });
  writeRow_('admins', found.rowIndex, updated);
}

function deleteRow_(resource, storeId, key, keyField) {
  if (!key) return;
  const sh = sheet_(resource);
  const headers = sh.getDataRange().getValues()[0];
  const values = sh.getDataRange().getValues();
  const storeIdx = headers.indexOf('storeId');
  const keyIdx = headers.indexOf(keyField);
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][storeIdx]) === String(storeId) && String(values[i][keyIdx]) === String(key)) sh.deleteRow(i + 1);
  }
}

function deleteStoreRows_(sh, headers, storeId) {
  const values = sh.getDataRange().getValues();
  const storeIdx = headers.indexOf('storeId');
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][storeIdx]) === String(storeId)) sh.deleteRow(i + 1);
  }
}

function findAdminByEmail_(storeId, email) {
  return findAdmin_(storeId, 'email', normalizeEmail_(email));
}

function findAdminById_(storeId, adminId) {
  return findAdmin_(storeId, 'adminId', String(adminId || ''));
}

function findAdmin_(storeId, field, value) {
  if (!value) return null;
  const sh = sheet_('admins');
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return null;
  const headers = values[0];
  const storeIdx = headers.indexOf('storeId');
  const fieldIdx = headers.indexOf(field);
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][storeIdx]) === String(storeId) && String(values[i][fieldIdx]) === String(value)) {
      return { rowIndex: i + 1, admin: rowToObject_(headers, values[i]) };
    }
  }
  return null;
}

function getStoreName_(storeId) {
  const profile = listResource_('profile', storeId, false);
  return profile.name || storeId;
}

function publicSession_(admin, storeName) {
  return {
    storeId: admin.storeId,
    storeName: storeName || admin.storeId,
    adminId: admin.adminId,
    name: admin.name,
    email: admin.email,
    role: admin.role || 'admin',
    sessionToken: admin.sessionToken,
    sessionExpires: admin.sessionExpires
  };
}

function rowToObject_(headers, row) {
  const obj = {};
  headers.forEach((h, i) => obj[h] = row[i]);
  return obj;
}

function compactPublicRow_(obj) {
  const out = Object.assign({}, obj);
  ['passwordHash','sessionToken','sessionExpires'].forEach(k => delete out[k]);
  if (out.img && String(out.img).startsWith('data:') && String(out.img).length > 5000) out.img = '';
  if (out.avatar && String(out.avatar).startsWith('data:') && String(out.avatar).length > 5000) out.avatar = '';
  return out;
}

function normalizeStoreId_(value) {
  return String(value || '').toLowerCase().trim().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 48);
}

function normalizeEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function makeAdminId_(storeId, email) {
  return `${storeId}-${Utilities.base64EncodeWebSafe(email).replace(/=+$/,'').slice(0, 16)}`;
}

function makeToken_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function expiresAt_() {
  return new Date(Date.now() + CONFIG.sessionTtlHours * 60 * 60 * 1000).toISOString();
}

function output_(payload, callback) {
  const body = callback ? `${callback}(${JSON.stringify(payload)})` : JSON.stringify(payload);
  return ContentService
    .createTextOutput(body)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

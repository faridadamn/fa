const SHEET_NAME = 'Articles';
const HEADERS = ['id', 'title', 'topic', 'excerpt', 'body', 'img', 'status', 'createdAt', 'updatedAt'];

function doGet(e) {
  const action = (e.parameter.action || 'list').toLowerCase();
  const callback = e.parameter.callback;
  const payload = action === 'list'
    ? { ok: true, articles: listArticles_() }
    : { ok: false, error: 'Unknown action' };
  return output_(payload, callback);
}

function doPost(e) {
  const data = JSON.parse((e.postData && e.postData.contents) || '{}');
  const action = (data.action || '').toLowerCase();
  if (data.resource !== 'articles') return output_({ ok: false, error: 'Unknown resource' });

  if (action === 'upsert') {
    upsertArticle_(data.article || {});
    return output_({ ok: true });
  }

  if (action === 'delete') {
    deleteArticle_((data.article || {}).id);
    return output_({ ok: true });
  }

  return output_({ ok: false, error: 'Unknown action' });
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (firstRow.join('') === '') sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  return sheet;
}

function listArticles_() {
  const sheet = sheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values[0];
  return values.slice(1).filter(row => row[0]).map(row => {
    const article = {};
    headers.forEach((key, i) => article[key] = row[i]);
    return article;
  });
}

function upsertArticle_(article) {
  if (!article.id) throw new Error('Missing article id');
  const sheet = sheet_();
  const values = sheet.getDataRange().getValues();
  const ids = values.map(row => row[0]);
  const rowIndex = ids.indexOf(article.id) + 1;
  const row = HEADERS.map(key => article[key] || '');
  if (rowIndex > 1) {
    sheet.getRange(rowIndex, 1, 1, HEADERS.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}

function deleteArticle_(id) {
  if (!id) return;
  const sheet = sheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) {
    if (values[i][0] === id) sheet.deleteRow(i + 1);
  }
}

function output_(payload, callback) {
  const body = callback ? `${callback}(${JSON.stringify(payload)})` : JSON.stringify(payload);
  const mime = callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON;
  return ContentService.createTextOutput(body).setMimeType(mime);
}

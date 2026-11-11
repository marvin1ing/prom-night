/**
 * 高三舞会之夜 — school photo import helper.
 *
 * Runs in the SCHOOL Google account, the one that can open the yearbook photo folder.
 * It reads each photo, makes an 800px copy and sends it to the prom website's backend,
 * where organizers pair photos with students by hand. Nothing is changed in the school folder.
 *
 * 1. Paste the import code from the organizer desk (学校照片配对 → 显示学校账号导入码) into KEY below.
 * 2. Pick importPhotos at the top and press Run. Approve access the first time.
 * 3. If the log says some photos are left, press Run again until it says 全部完成.
 */
const KEY = '在这里粘贴导入码';
const API = 'https://script.google.com/macros/s/AKfycbzN2PEzNCUdXDhXviwFMf1hc6SS7hoJBg-_w_EDueIZU8TgaVSr7l2KXAZYr7c-wd6Ecg/exec';
const FOLDER_ID = '1EWx_OiRHOdgF5KlZwU2nbJ2YvpMdK1GN';
const SIZE = 800;
const BUDGET_MS = 4.5 * 60 * 1000;

function importPhotos() {
  if (!KEY || /粘贴/.test(KEY)) throw new Error('先把导入码粘贴到第 12 行的 KEY 里。');
  const start = Date.now();
  const known = {};
  call_({ action: 'schoolKnown', key: KEY }).ids.forEach(id => { known[id] = true; });
  let batch = [], added = 0, left = 0, failed = 0;
  const flush = () => {
    if (!batch.length) return;
    added += call_({ action: 'schoolImport', key: KEY, items: batch }).added;
    batch = [];
  };
  const walk = (folder, cls) => {
    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (!/^image\//.test(f.getMimeType()) || known[f.getId()]) continue;
      if (Date.now() - start > BUDGET_MS) { left++; continue; }
      let data;
      try { data = small_(f); } catch (e) { failed++; Logger.log('跳过 ' + cls + '/' + f.getName() + '：' + e.message); continue; }
      batch.push({ sourceId: f.getId(), class: cls, file: f.getName(), data });
      known[f.getId()] = true;
      if (batch.length >= 4) flush();
    }
    const subs = folder.getFolders();
    while (subs.hasNext()) { const s = subs.next(); walk(s, s.getName()); }
  };
  const root = DriveApp.getFolderById(FOLDER_ID);
  walk(root, root.getName());
  flush();
  Logger.log('这次传了 ' + added + ' 张。' + (failed ? '有 ' + failed + ' 张读不到。' : '') + (left ? '还剩 ' + left + ' 张，请再按一次 Run。' : '全部完成。'));
}

/** An 800px JPEG of the photo, as a data URL. */
function small_(f) {
  const auth = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  const meta = JSON.parse(UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + f.getId() + '?fields=thumbnailLink&supportsAllDrives=true', { headers: auth }).getContentText());
  let blob = null;
  if (meta.thumbnailLink) {
    const r = UrlFetchApp.fetch(meta.thumbnailLink.replace(/=s\d+$/, '=s' + SIZE), { headers: auth, muteHttpExceptions: true });
    if (r.getResponseCode() === 200) blob = r.getBlob();
  }
  if (!blob) blob = f.getThumbnail();
  if (!blob) throw new Error('no thumbnail');
  const type = (blob.getContentType() || 'image/jpeg').split(';')[0];
  return 'data:' + (/^image\/(jpeg|png|webp)$/.test(type) ? type : 'image/jpeg') + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

function call_(body) {
  const r = UrlFetchApp.fetch(API, { method: 'post', contentType: 'text/plain; charset=utf-8', payload: JSON.stringify(body), muteHttpExceptions: true, followRedirects: true });
  const out = JSON.parse(r.getContentText());
  if (!out.ok) throw new Error(out.error === 'bad_key' ? '导入码不对，请从后台重新复制。' : out.error === 'import_off' ? '后台还没生成导入码：先在后台点“显示学校账号导入码”。' : '后台出错：' + out.error);
  return out;
}

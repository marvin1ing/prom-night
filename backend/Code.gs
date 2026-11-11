/**
 * 高三舞会之夜 — Google Apps Script backend.
 *
 * Lives inside a Google Sheet (Extensions → Apps Script). Stores people, ballots and photo records in the
 * sheet, and photos in Drive. The website calls it with POST requests whose body is JSON text.
 *
 * Privacy: ballots are stored without any code or person id. The People tab only records THAT someone voted.
 * The admin password lives in Script Properties, never in this file.
 */

const SHOTS_PER_ROLL = 15;
const DEFAULT_DEVELOP_AT = '2026-11-29T09:00:00+08:00';
const PEOPLE_COLS = ['id', 'code', 'name', 'en', 'class', 'no', 'role', 'list', 'photoId', 'photoSource', 'voted', 'votedAt', 'shots', 'updatedAt'];
const ALBUM_COLS = ['fileId', 'ownerId', 'at', 'film', 'deleted'];
const BALLOT_COLS = ['at', 'king1', 'king2', 'queen1', 'queen2'];
const CODE_ALPHABET = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';

/* ---------------- setup (run from the script editor) ---------------- */
/*
 * Run setup() once. It creates (or reuses) the spreadsheet, the Drive folders and the People list.
 * Script properties (Project Settings → Script properties):
 *   ADMIN_PASSWORD           organizer password, set by hand, never in code
 *   SCHOOL_PHOTO_FOLDER_ID   optional, the school's photo folder, then run importSchoolPhotos()
 */
function ss_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SHEET_ID');
  if (!id) {
    const ss = SpreadsheetApp.create('高三舞会之夜 后台');
    id = ss.getId();
    props.setProperty('SHEET_ID', id);
  }
  return SpreadsheetApp.openById(id);
}
function say_(msg) { Logger.log(msg); return msg; }

function setup() {
  const ss = ss_();
  let people = ss.getSheetByName('People');
  if (!people) {
    if (typeof ROSTER_DATA === 'undefined') return say_('没有找到名单（Roster.gs）。');
    const data = [ROSTER_DATA.cols].concat(ROSTER_DATA.rows);
    people = ss.insertSheet('People');
    const head = data[0].map(String);
    const out = [PEOPLE_COLS];
    for (let r = 1; r < data.length; r++) {
      if (!data[r][head.indexOf('code')]) continue;
      out.push(PEOPLE_COLS.map(c => (head.indexOf(c) >= 0 ? data[r][head.indexOf(c)] : (c === 'shots' ? 0 : ''))));
    }
    people.getRange('B:B').setNumberFormat('@');
    people.getRange(1, 1, out.length, PEOPLE_COLS.length).setValues(out);
    people.setFrozenRows(1);
  }
  ensureSheet_('Ballots', BALLOT_COLS);
  ensureSheet_('Album', ALBUM_COLS);
  ensureSheet_('Log', ['at', 'who', 'action', 'detail']);
  const settings = ensureSheet_('Settings', ['key', 'value']);
  const defaults = { voting: 'closed', headshots: 'open', developAt: DEFAULT_DEVELOP_AT, developed: 'auto', title: '高三舞会之夜' };
  const have = settings.getDataRange().getValues().map(r => r[0]);
  Object.keys(defaults).forEach(k => { if (have.indexOf(k) < 0) settings.appendRow([k, defaults[k]]); });
  ss.getSheets().forEach(sh => { if (['People', 'Ballots', 'Album', 'Log', 'Settings'].indexOf(sh.getName()) < 0 && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh); });

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ROOT_FOLDER_ID')) {
    const root = DriveApp.createFolder('Prom Night 2026');
    props.setProperties({ ROOT_FOLDER_ID: root.getId(), HEADSHOT_FOLDER_ID: root.createFolder('头像').getId(), ALBUM_FOLDER_ID: root.createFolder('相册').getId() });
    DriveApp.getFileById(ss.getId()).moveTo(root);
  }
  clearCache_();
  return say_('初始化完成：People 有 ' + (people.getLastRow() - 1) + ' 人。表格：' + ss.getUrl());
}

/** Copies photos from the school folder (Script property SCHOOL_PHOTO_FOLDER_ID) for students without their own photo. Matches file names to 中文名 or 英文名. */
function importSchoolPhotos() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('SCHOOL_PHOTO_FOLDER_ID') || '';
  const m = raw.match(/folders\/([A-Za-z0-9_-]+)/) || raw.match(/^([A-Za-z0-9_-]{20,})$/);
  if (!m) return say_('先在 Script properties 里填 SCHOOL_PHOTO_FOLDER_ID（文件夹链接或 ID）。');
  const sh = ss_().getSheetByName('People');
  const rows = sh.getDataRange().getValues();
  const H = rows[0];
  const col = n => H.indexOf(n);
  const norm = s => String(s || '').toLowerCase().replace(/\.[a-z0-9]+$/, '').replace(/[\s_\-.()（）]+/g, '');
  const index = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][col('role')] !== 'student') continue;
    [rows[i][col('name')], rows[i][col('en')]].forEach(n => { if (n) index[norm(n)] = i; });
  }
  const headFolder = DriveApp.getFolderById(props.getProperty('HEADSHOT_FOLDER_ID'));
  const files = DriveApp.getFolderById(m[1]).getFiles();
  const unmatched = [];
  let done = 0, skipped = 0;
  while (files.hasNext()) {
    const f = files.next();
    if (!/^image\//.test(f.getMimeType())) continue;
    const key = norm(f.getName());
    let i = index[key];
    if (i === undefined) {
      const hit = Object.keys(index).filter(k => key.indexOf(k) >= 0 && k.length >= 2).sort((a, b) => b.length - a.length)[0];
      i = hit ? index[hit] : undefined;
    }
    if (i === undefined) { unmatched.push(f.getName()); continue; }
    if (rows[i][col('photoSource')] === 'self') { skipped++; continue; }
    const copy = f.makeCopy(rows[i][col('id')] + '_school', headFolder);
    copy.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    sh.getRange(i + 1, col('photoId') + 1, 1, 2).setValues([[copy.getId(), 'school']]);
    rows[i][col('photoSource')] = 'school';
    done++;
  }
  clearCache_();
  return say_('导入了 ' + done + ' 张，跳过 ' + skipped + ' 张（本人已上传）。对不上名字的文件：' + (unmatched.join('、') || '无'));
}

/** Adds extra guest codes. Edit the two values below, run once, then read the codes in the People sheet. */
function addGuestCodes(unit, n) {
  unit = unit || '校友会'; n = n || 1;
  const sh = ss_().getSheetByName('People');
  const added = [];
  for (let k = 1; k <= n; k++) {
    const code = newCode_(sh);
    const id = 'G' + Utilities.getUuid().slice(0, 6);
    sh.appendRow(PEOPLE_COLS.map(c => ({ id, code, name: unit + ' +' + k, class: '外聘&赞助', no: k, role: 'guest', shots: 0 }[c] || '')));
    added.push(unit + ' +' + k + '：' + code);
  }
  clearCache_();
  return say_('新增：' + added.join('，'));
}

/* ---------------- web API ---------------- */

function doGet() {
  return json_({ ok: true, service: 'prom-night', time: new Date().toISOString() });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  try {
    const fn = ACTIONS[req.action];
    if (!fn) return json_({ ok: false, error: 'unknown_action' });
    return json_(Object.assign({ ok: true }, fn(req)));
  } catch (err) {
    const known = String(err && err.message || err);
    return json_({ ok: false, error: /^[a-z_]+$/.test(known) ? known : 'server_error', detail: /^[a-z_]+$/.test(known) ? undefined : known });
  }
}

const ACTIONS = {
  status: () => ({ settings: publicSettings_(), turnout: turnout_() }),

  login: req => {
    const p = person_(req.code);
    return { me: meView_(p), settings: publicSettings_(), roster: rosterPublic_(), turnout: turnout_() };
  },

  roster: req => { person_(req.code); return { roster: rosterPublic_() }; },

  headshot: req => {
    const p = person_(req.code);
    if (p.role !== 'student') throw new Error('not_student');
    const s = settings_();
    if (s.headshots !== 'open' || s.voting !== 'closed') throw new Error('headshots_closed');
    const blob = imageBlob_(req.data, p.id + '_' + Date.now() + '.jpg');
    return withLock_(() => {
      const fresh = person_(req.code);
      const folder = DriveApp.getFolderById(prop_('HEADSHOT_FOLDER_ID'));
      const f = folder.createFile(blob);
      f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      if (fresh.photoId) trash_(fresh.photoId);
      setPerson_(fresh.row, { photoId: f.getId(), photoSource: 'self', updatedAt: new Date() });
      clearCache_();
      log_(fresh.id, 'headshot', f.getId());
      return { photo: photoUrl_(f.getId(), 600) };
    });
  },

  vote: req => {
    const p = person_(req.code);
    if (settings_().voting !== 'open') throw new Error('voting_closed');
    const king = (req.king || []).map(String), queen = (req.queen || []).map(String);
    const roster = rosterPublic_();
    const ok = (ids, list) => ids.length === 2 && ids[0] !== ids[1] && ids.every(id => roster.some(r => r.id === id && r.list === list));
    if (!ok(king, 'king') || !ok(queen, 'queen')) throw new Error('bad_ballot');
    return withLock_(() => {
      const fresh = person_(req.code);
      if (fresh.voted) throw new Error('already_voted');
      ss_().getSheetByName('Ballots').appendRow([new Date(), king[0], king[1], queen[0], queen[1]]);
      setPerson_(fresh.row, { voted: true, votedAt: new Date() });
      CacheService.getScriptCache().remove('turnout');
      return { voted: true, turnout: turnout_() };
    });
  },

  shot: req => {
    const p = person_(req.code);
    if ((Number(p.shots) || 0) >= SHOTS_PER_ROLL) throw new Error('roll_finished');
    const blob = imageBlob_(req.data, 'shot_' + Date.now() + '_' + p.id + '.jpg');
    return withLock_(() => {
      const fresh = person_(req.code);
      const n = Number(fresh.shots) || 0;
      if (n >= SHOTS_PER_ROLL) throw new Error('roll_finished');
      const f = DriveApp.getFolderById(prop_('ALBUM_FOLDER_ID')).createFile(blob);
      f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      ss_().getSheetByName('Album').appendRow([f.getId(), fresh.id, new Date(req.at || Date.now()), String(req.film || ''), '']);
      setPerson_(fresh.row, { shots: n + 1 });
      return { shots: n + 1 };
    });
  },

  album: req => {
    const s = settings_();
    let me = null;
    if (req.code) { try { me = person_(req.code); } catch (err) { me = null; } }
    const rows = albumRows_().filter(r => !r.deleted);
    if (!isDeveloped_(s)) return { developed: false, count: rows.length, developAt: s.developAt };
    return {
      developed: true,
      photos: rows.sort((a, b) => a.at - b.at).map(r => ({ id: r.fileId, at: r.at.toISOString(), mine: !!(me && r.ownerId === me.id), thumb: photoUrl_(r.fileId, 700), full: photoUrl_(r.fileId, 1600) })),
    };
  },

  deleteShot: req => {
    const p = person_(req.code);
    return withLock_(() => {
      const r = albumRows_().find(x => x.fileId === req.fileId && !x.deleted);
      if (!r) throw new Error('not_found');
      if (r.ownerId !== p.id) throw new Error('not_yours');
      markShotDeleted_(r, p.id);
      return { deleted: true };
    });
  },

  admin: req => {
    checkPassword_(req.pw);
    const op = ADMIN[req.op];
    if (!op) throw new Error('unknown_op');
    return op(req);
  },
};

const ADMIN = {
  summary: () => {
    const all = peopleRows_();
    const students = all.filter(p => p.role === 'student');
    return {
      settings: Object.assign(settings_(), { developedNow: isDeveloped_(settings_()) }),
      people: { total: all.length, students: students.length, voted: all.filter(p => p.voted).length },
      headshots: { self: students.filter(p => p.photoSource === 'self').length, school: students.filter(p => p.photoSource === 'school').length, none: students.filter(p => !p.photoId).length },
      album: { photos: albumRows_().filter(r => !r.deleted).length, cameras: all.filter(p => Number(p.shots) > 0).length },
      results: results_(),
      roster: rosterPublic_(),
    };
  },
  setSetting: req => {
    const allowed = { voting: ['closed', 'open', 'done'], headshots: ['open', 'closed'], developed: ['auto', 'yes'] };
    if (!allowed[req.key] || allowed[req.key].indexOf(req.value) < 0) throw new Error('bad_setting');
    writeSetting_(req.key, req.value);
    log_('admin', 'setting', req.key + '=' + req.value);
    return { settings: settings_() };
  },
  find: req => {
    const q = String(req.q || '').toLowerCase().trim();
    if (!q) return { people: [] };
    return {
      people: peopleRows_().filter(p => (p.name + ' ' + p.en + ' ' + p.class).toLowerCase().indexOf(q) >= 0).slice(0, 20)
        .map(p => ({ id: p.id, name: p.name, en: p.en, class: p.class, role: p.role, code: p.code, voted: !!p.voted, shots: Number(p.shots) || 0, photo: p.photoId ? photoUrl_(p.photoId, 300) : '', photoSource: p.photoSource })),
    };
  },
  reissue: req => withLock_(() => {
    const p = peopleRows_().find(x => x.id === req.id);
    if (!p) throw new Error('not_found');
    const sh = ss_().getSheetByName('People');
    const code = newCode_(sh);
    setPerson_(p.row, { code, updatedAt: new Date() });
    clearCache_();
    log_('admin', 'reissue', p.id);
    return { id: p.id, code };
  }),
  deleteHeadshot: req => withLock_(() => {
    const p = peopleRows_().find(x => x.id === req.id);
    if (!p) throw new Error('not_found');
    if (p.photoId) trash_(p.photoId);
    setPerson_(p.row, { photoId: '', photoSource: '', updatedAt: new Date() });
    clearCache_();
    log_('admin', 'deleteHeadshot', p.id);
    return { deleted: true };
  }),
  deleteShot: req => withLock_(() => {
    const r = albumRows_().find(x => x.fileId === req.fileId && !x.deleted);
    if (!r) throw new Error('not_found');
    markShotDeleted_(r, 'admin');
    return { deleted: true };
  }),
  resetTestData: req => {
    if (req.confirm !== 'RESET') throw new Error('confirm_needed');
    return withLock_(() => {
      const ss = ss_();
      ['Ballots', 'Album'].forEach(n => { const s = ss.getSheetByName(n); if (s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1); });
      const sh = ss.getSheetByName('People');
      const n = sh.getLastRow() - 1;
      if (n > 0) {
        const c = PEOPLE_COLS.indexOf('voted') + 1;
        sh.getRange(2, c, n, 3).setValues(Array.from({ length: n }, () => ['', '', 0]));
      }
      const files = DriveApp.getFolderById(prop_('ALBUM_FOLDER_ID')).getFiles();
      while (files.hasNext()) files.next().setTrashed(true);
      writeSetting_('voting', 'closed');
      writeSetting_('developed', 'auto');
      clearCache_();
      CacheService.getScriptCache().remove('turnout');
      log_('admin', 'resetTestData', '');
      return { reset: true };
    });
  },
};

/* ---------------- helpers ---------------- */

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function prop_(k) { const v = PropertiesService.getScriptProperties().getProperty(k); if (!v) throw new Error('not_set_up'); return v; }
function headerIndex_(sheet, name) { const r = sheet.getLastRow() ? sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0] : []; return r.map(String).indexOf(name); }
function ensureSheet_(name, cols) {
  const ss = ss_();
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.getRange(1, 1, 1, cols.length).setValues([cols]); s.setFrozenRows(1); }
  return s;
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('busy');
  try { return fn(); } finally { lock.releaseLock(); }
}
function log_(who, action, detail) {
  try { ss_().getSheetByName('Log').appendRow([new Date(), who, action, detail]); } catch (err) { /* logging must never break a request */ }
}
function trash_(id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (err) { /* already gone */ } }
function photoUrl_(id, w) { return 'https://drive.google.com/thumbnail?id=' + id + '&sz=w' + w; }

function imageBlob_(data, name) {
  const m = String(data || '').match(/^data:image\/(jpeg|png|webp);base64,(.+)$/);
  if (!m) throw new Error('bad_image');
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > 4 * 1024 * 1024) throw new Error('image_too_big');
  return Utilities.newBlob(bytes, 'image/' + m[1], name);
}

function peopleRows_() {
  const sh = ss_().getSheetByName('People');
  const v = sh.getDataRange().getValues();
  const H = v[0];
  return v.slice(1).map((r, i) => { const o = { row: i + 2 }; H.forEach((h, j) => { o[h] = r[j]; }); o.code = String(o.code || '').toUpperCase(); return o; });
}
function person_(code) {
  const c = String(code || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (c.length !== 8) throw new Error('bad_code');
  const p = peopleRows_().find(x => x.code === c);
  if (!p) throw new Error('bad_code');
  return p;
}
function setPerson_(row, patch) {
  const sh = ss_().getSheetByName('People');
  Object.keys(patch).forEach(k => {
    const c = PEOPLE_COLS.indexOf(k);
    if (c >= 0) sh.getRange(row, c + 1).setValue(patch[k]);
  });
}
function newCode_(sh) {
  const have = {};
  sh.getRange(2, 2, Math.max(1, sh.getLastRow() - 1), 1).getValues().forEach(r => { have[String(r[0]).toUpperCase()] = true; });
  for (;;) {
    let c = '';
    for (let i = 0; i < 8; i++) c += CODE_ALPHABET[Math.floor(Math.random() * CODE_ALPHABET.length)];
    if (!have[c]) return c;
  }
}
function meView_(p) {
  return { id: p.id, name: p.name, en: p.en, class: p.class, role: p.role, list: p.list, photo: p.photoId ? photoUrl_(p.photoId, 600) : '', photoSource: p.photoSource, voted: !!p.voted, shots: Number(p.shots) || 0, shotsMax: SHOTS_PER_ROLL };
}
function rosterPublic_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('roster');
  if (hit) return JSON.parse(hit);
  const list = peopleRows_().filter(p => p.role === 'student')
    .map(p => ({ id: p.id, name: p.name, en: p.en, class: p.class, list: p.list, photo: p.photoId ? photoUrl_(p.photoId, 300) : '' }));
  cache.put('roster', JSON.stringify(list), 120);
  return list;
}
function clearCache_() { CacheService.getScriptCache().remove('roster'); }
function settings_() {
  const v = ss_().getSheetByName('Settings').getDataRange().getValues();
  const o = {};
  v.slice(1).forEach(r => { o[r[0]] = r[1] instanceof Date ? r[1].toISOString() : String(r[1]); });
  return o;
}
function writeSetting_(key, value) {
  const sh = ss_().getSheetByName('Settings');
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] === key) { sh.getRange(i + 1, 2).setValue(value); return; }
  sh.appendRow([key, value]);
}
function publicSettings_() {
  const s = settings_();
  return { voting: s.voting, headshots: s.headshots, developed: isDeveloped_(s), developAt: s.developAt, title: s.title, votingHint: s.votingHint, shotsMax: SHOTS_PER_ROLL };
}
function isDeveloped_(s) { return s.developed === 'yes' || Date.now() >= new Date(s.developAt || DEFAULT_DEVELOP_AT).getTime(); }
function turnout_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('turnout');
  if (hit) return JSON.parse(hit);
  const all = peopleRows_();
  const t = { voted: all.filter(p => p.voted).length, eligible: all.length };
  cache.put('turnout', JSON.stringify(t), 15);
  return t;
}
function albumRows_() {
  const v = ss_().getSheetByName('Album').getDataRange().getValues();
  return v.slice(1).map((r, i) => ({ row: i + 2, fileId: r[0], ownerId: r[1], at: r[2] instanceof Date ? r[2] : new Date(r[2]), film: r[3], deleted: !!r[4] }));
}
function markShotDeleted_(r, who) {
  ss_().getSheetByName('Album').getRange(r.row, 5).setValue(new Date());
  trash_(r.fileId);
  log_(who, 'deleteShot', r.fileId);
}
function results_() {
  const v = ss_().getSheetByName('Ballots').getDataRange().getValues().slice(1);
  const t = { king: {}, queen: {} };
  v.forEach(r => {
    [r[1], r[2]].forEach(id => { if (id) t.king[id] = (t.king[id] || 0) + 1; });
    [r[3], r[4]].forEach(id => { if (id) t.queen[id] = (t.queen[id] || 0) + 1; });
  });
  return { ballots: v.length, king: t.king, queen: t.queen };
}
function checkPassword_(pw) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pwfails') || 0);
  if (fails >= 10) throw new Error('too_many_tries');
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!real) throw new Error('no_password_set');
  if (String(pw || '') !== real) { cache.put('pwfails', String(fails + 1), 600); throw new Error('bad_password'); }
}

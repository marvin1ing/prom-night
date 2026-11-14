/**
 * 《光序•耀光》谢师宴暨高三毕业舞会 — Google Apps Script backend.
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
const SCHOOL_COLS = ['sourceId', 'fileId', 'class', 'name', 'personId', 'skip'];

/* ---------------- setup (run from the script editor) ---------------- */
/*
 * Run setup() once. It creates (or reuses) the spreadsheet, the Drive folders and the People list.
 * Script properties (Project Settings → Script properties):
 *   ADMIN_PASSWORD           organizer password, set by hand, never in code
 *   SCHOOL_PHOTO_FOLDER_ID   set from the organizer desk; school photos are paired by hand there
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
  const defaults = { voting: 'closed', headshots: 'open', developAt: DEFAULT_DEVELOP_AT, developed: 'auto', title: '《光序•耀光》谢师宴暨高三毕业舞会' };
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

/** Adds people from ROSTER_DATA who are not in the People sheet yet (by id). Never changes existing rows. */
function syncRoster() {
  if (typeof ROSTER_DATA === 'undefined') return say_('没有找到名单。');
  const sh = ss_().getSheetByName('People');
  const have = {};
  sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 2).getValues().forEach(r => { have[String(r[0])] = true; have[String(r[1]).toUpperCase()] = true; });
  const head = ROSTER_DATA.cols;
  const added = [];
  ROSTER_DATA.rows.forEach(row => {
    const o = {}; head.forEach((c, i) => { o[c] = row[i]; });
    if (!o.code || have[String(o.id)] || have[String(o.code).toUpperCase()]) return;
    sh.appendRow(PEOPLE_COLS.map(c => (c in o ? o[c] : (c === 'shots' ? 0 : ''))));
    have[String(o.id)] = true;
    added.push(o.name);
  });
  sh.getRange('B:B').setNumberFormat('@');
  clearCache_();
  return say_('新增 ' + added.length + ' 人：' + added.join('，') + '。People 现在有 ' + (sh.getLastRow() - 1) + ' 人。');
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

  /** Used by the helper script in the school account: which school photos are already here. */
  schoolKnown: req => { checkImportKey_(req.key); return { ids: schoolRows_().map(r => String(r.sourceId)) }; },

  /** Used by the helper script in the school account: receives a few resized school photos. */
  schoolImport: req => {
    checkImportKey_(req.key);
    const items = (req.items || []).slice(0, 6);
    const blobs = items.map(it => ({ it, blob: imageBlob_(it.data, String(it.class || '') + '_' + String(it.file || 'photo.jpg')) }));
    return withLock_(() => {
      const props = PropertiesService.getScriptProperties();
      let target = props.getProperty('SCHOOL_COPY_FOLDER_ID');
      if (!target) { target = DriveApp.getFolderById(prop_('ROOT_FOLDER_ID')).createFolder('学校照片').getId(); props.setProperty('SCHOOL_COPY_FOLDER_ID', target); }
      const dest = DriveApp.getFolderById(target);
      const sh = schoolSheet_();
      const have = {};
      schoolRows_().forEach(r => { have[String(r.sourceId)] = true; });
      let added = 0;
      blobs.forEach(({ it, blob }) => {
        const src = String(it.sourceId || '');
        if (!src || have[src]) return;
        const f = dest.createFile(blob);
        f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        sh.appendRow([src, f.getId(), String(it.class || ''), String(it.file || ''), '', '']);
        have[src] = true;
        added++;
      });
      return { added, staged: Object.keys(have).length };
    });
  },

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
      if (fresh.photoId && fresh.photoSource === 'self') trash_(fresh.photoId);
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
      school: { folder: PropertiesService.getScriptProperties().getProperty('SCHOOL_PHOTO_FOLDER_ID') || '', staged: schoolRows_().length },
      sources: students.reduce((o, p) => { o[p.id] = p.photoSource || ''; return o; }, {}),
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
    if (p.photoId && p.photoSource === 'self') trash_(p.photoId);
    const school = schoolRows_().find(r => r.personId === p.id && !r.skip);
    if (p.photoSource === 'school' && school && school.fileId === p.photoId) unassignSchool_(school);
    const fallback = schoolRows_().find(r => r.personId === p.id && !r.skip);
    setPerson_(p.row, fallback ? { photoId: fallback.fileId, photoSource: 'school', updatedAt: new Date() } : { photoId: '', photoSource: '', updatedAt: new Date() });
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
  setSchoolFolder: req => {
    const raw = String(req.url || '').trim();
    const m = raw.match(/folders\/([A-Za-z0-9_-]+)/) || raw.match(/^([A-Za-z0-9_-]{20,})$/);
    if (!m) throw new Error('bad_folder_link');
    try { DriveApp.getFolderById(m[1]).getName(); } catch (e) { throw new Error('no_folder_access'); }
    PropertiesService.getScriptProperties().setProperty('SCHOOL_PHOTO_FOLDER_ID', m[1]);
    return { folder: m[1] };
  },
  stageSchool: () => stageSchoolPhotos_(70000),
  /** Import code for the helper script in the school account. Created once, then reused. */
  importKey: () => {
    const props = PropertiesService.getScriptProperties();
    let k = props.getProperty('SCHOOL_IMPORT_KEY');
    if (!k) { k = Utilities.getUuid().replace(/-/g, '').slice(0, 20).toUpperCase(); props.setProperty('SCHOOL_IMPORT_KEY', k); }
    return { key: k };
  },
  schoolPhotos: () => ({ photos: schoolRows_().map(r => ({ id: r.fileId, class: r.class, file: r.name, personId: r.personId, skip: r.skip, thumb: photoUrl_(r.fileId, 500), big: photoUrl_(r.fileId, 1000) })) }),
  assignSchool: req => withLock_(() => {
    const rows = schoolRows_();
    const row = rows.find(r => r.fileId === req.fileId);
    if (!row) throw new Error('not_found');
    if (row.personId) unassignSchool_(row);
    if (req.skip) { schoolSheet_().getRange(row.row, 6).setValue('yes'); clearCache_(); return { ok: true }; }
    schoolSheet_().getRange(row.row, 6).setValue('');
    if (!req.personId) { clearCache_(); return { ok: true }; }
    const p = peopleRows_().find(x => x.id === req.personId && x.role === 'student');
    if (!p) throw new Error('not_found');
    schoolRows_().filter(r => r.personId === p.id).forEach(r => unassignSchool_(r));
    schoolSheet_().getRange(row.row, 5).setValue(p.id);
    if (p.photoSource !== 'self') setPerson_(p.row, { photoId: row.fileId, photoSource: 'school', updatedAt: new Date() });
    clearCache_();
    log_('admin', 'assignSchool', row.fileId + '→' + p.id);
    return { ok: true };
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
function schoolSheet_() { return ensureSheet_('SchoolPhotos', SCHOOL_COLS); }
function schoolRows_() {
  const v = schoolSheet_().getDataRange().getValues();
  return v.slice(1).map((r, i) => ({ row: i + 2, sourceId: r[0], fileId: r[1], class: r[2], name: r[3], personId: r[4], skip: !!r[5] }));
}
/** Clears one pairing; if that person was showing this school photo, they go back to no photo. */
function unassignSchool_(row) {
  if (!row.personId) return;
  const p = peopleRows_().find(x => x.id === row.personId);
  if (p && p.photoSource === 'school' && p.photoId === row.fileId) setPerson_(p.row, { photoId: '', photoSource: '', updatedAt: new Date() });
  schoolSheet_().getRange(row.row, 5).setValue('');
}
/** Copies school photos (class subfolders) into 学校照片 so the website can show them. Resumable: call again until remaining is false. */
function stageSchoolPhotos_(budgetMs) {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SCHOOL_PHOTO_FOLDER_ID');
  if (!id) throw new Error('no_school_folder');
  let root;
  try { root = DriveApp.getFolderById(id); root.getName(); } catch (e) { throw new Error('no_folder_access'); }
  let target = props.getProperty('SCHOOL_COPY_FOLDER_ID');
  if (!target) { target = DriveApp.getFolderById(prop_('ROOT_FOLDER_ID')).createFolder('学校照片').getId(); props.setProperty('SCHOOL_COPY_FOLDER_ID', target); }
  const dest = DriveApp.getFolderById(target);
  const sh = schoolSheet_();
  const have = {};
  schoolRows_().forEach(r => { have[r.sourceId] = true; });
  const start = Date.now();
  let added = 0, remaining = false;
  const walk = (folder, cls) => {
    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (!/^image\//.test(f.getMimeType()) || have[f.getId()]) continue;
      if (Date.now() - start > budgetMs) { remaining = true; return; }
      const c = f.makeCopy(cls + '_' + f.getName(), dest);
      c.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      sh.appendRow([f.getId(), c.getId(), cls, f.getName(), '', '']);
      have[f.getId()] = true;
      added++;
    }
    const subs = folder.getFolders();
    while (subs.hasNext() && !remaining) { const sub = subs.next(); walk(sub, sub.getName()); }
  };
  walk(root, root.getName());
  return { added, remaining, staged: Object.keys(have).length };
}

/** The helper script in the school account sends this key (script property SCHOOL_IMPORT_KEY). */
function checkImportKey_(key) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('keyfails') || 0);
  if (fails >= 10) throw new Error('too_many_tries');
  const real = PropertiesService.getScriptProperties().getProperty('SCHOOL_IMPORT_KEY');
  if (!real) throw new Error('import_off');
  if (String(key || '') !== real) { cache.put('keyfails', String(fails + 1), 600); throw new Error('bad_key'); }
}

function checkPassword_(pw) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pwfails') || 0);
  if (fails >= 10) throw new Error('too_many_tries');
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!real) throw new Error('no_password_set');
  if (String(pw || '') !== real) { cache.put('pwfails', String(fails + 1), 600); throw new Error('bad_password'); }
}

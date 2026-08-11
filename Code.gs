/**
 * ระบบติดตามงานตัดกล่องตัวอย่าง — Backend
 * วางไฟล์นี้ใน Google Apps Script ที่ผูกกับ Google Sheet
 *
 * ขั้นตอนใช้งานครั้งแรก: รันฟังก์ชัน setup() หนึ่งครั้ง แล้วดูรหัสผ่านใน Log
 */

const CFG = {
  JOBS: 'Jobs',
  USERS: 'Users',
  LOG: 'Log',
  TZ: 'Asia/Bangkok',
  TOKEN_TTL_H: 12,
};

const STATUS = {
  WAIT_CUT: 'รอตัด',
  CUT_DONE: 'ตัดเสร็จ',
  SHIPPED: 'ส่งแล้ว',
  CANCELLED: 'ยกเลิก',
};

const ROLE = { CUTTER: 'cutter', SHIPPING: 'shipping', VIEWER: 'viewer' };

const JOB_COLS = [
  'id', 'customer', 'sales', 'jobName', 'qty', 'destination',
  'contactName', 'contactPhone', 'dueDate', 'status', 'note',
  'createdBy', 'createdAt', 'cutBy', 'cutAt',
  'vehicle', 'shipBy', 'shipAt', 'updatedAt',
];

const USER_COLS = ['username', 'name', 'role', 'salt', 'hash', 'active'];
const LOG_COLS = ['time', 'username', 'action', 'jobId', 'detail'];

/* ============================================================
 * ติดตั้งครั้งแรก
 * ========================================================== */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(CFG.TZ);

  ensureSheet(CFG.JOBS, JOB_COLS);
  ensureSheet(CFG.USERS, USER_COLS);
  ensureSheet(CFG.LOG, LOG_COLS);

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SECRET')) {
    props.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  }

  const users = sheet(CFG.USERS);
  if (users.getLastRow() > 1) {
    Logger.log('มีผู้ใช้อยู่แล้ว ข้ามการสร้างบัญชีเริ่มต้น');
    return;
  }

  const seed = [
    ['cutter', 'ช่างตัด', ROLE.CUTTER],
    ['ship', 'จัดส่ง', ROLE.SHIPPING],
    ['ae', 'AE', ROLE.VIEWER],
    ['artwork', 'Artwork', ROLE.VIEWER],
  ];

  const lines = ['===== รหัสผ่านเริ่มต้น (บันทึกไว้แล้วเปลี่ยนทีหลัง) ====='];
  seed.forEach(function (u) {
    const pw = randomPassword();
    const salt = Utilities.getUuid();
    users.appendRow([u[0], u[1], u[2], salt, hashPassword(pw, salt), 'yes']);
    lines.push(u[0] + '  /  ' + pw + '   (' + u[1] + ')');
  });
  Logger.log(lines.join('\n'));
}

function ensureSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#f1f3f5');
    sh.setFrozenRows(1);
  }
  // บังคับให้ทุกคอลัมน์เป็นข้อความล้วน ไม่ให้ Sheets แปลงวันที่/เวลาเอง
  sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
  return sh;
}

function sheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function randomPassword() {
  const chars = 'abcdefghjkmnpqrstuvwxyz23456789';
  let out = '';
  for (let i = 0; i < 10; i++) out += chars.charAt(Math.floor(Math.random() * chars.length));
  return out;
}

/* ============================================================
 * Auth
 * ========================================================== */

function hashPassword(password, salt) {
  const raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, salt + '::' + password, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(raw);
}

function secret() {
  return PropertiesService.getScriptProperties().getProperty('SECRET');
}

function sign(payload) {
  const raw = Utilities.computeHmacSha256Signature(payload, secret());
  return Utilities.base64EncodeWebSafe(raw);
}

function makeToken(username) {
  const exp = Date.now() + CFG.TOKEN_TTL_H * 3600 * 1000;
  const payload = username + '|' + exp;
  return Utilities.base64EncodeWebSafe(payload + '|' + sign(payload));
}

function readToken(token) {
  if (!token) return null;
  let decoded;
  try {
    decoded = Utilities.newBlob(Utilities.base64DecodeWebSafe(token)).getDataAsString();
  } catch (err) {
    return null;
  }
  const parts = decoded.split('|');
  if (parts.length !== 3) return null;
  const payload = parts[0] + '|' + parts[1];
  if (sign(payload) !== parts[2]) return null;
  if (Date.now() > Number(parts[1])) return null;
  return parts[0];
}

function findUser(username) {
  const rows = readSheet(CFG.USERS, USER_COLS);
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i].username).toLowerCase() === String(username).toLowerCase()) {
      rows[i]._row = i + 2;
      return rows[i];
    }
  }
  return null;
}

/** คืน object ผู้ใช้ หรือโยน error ถ้า token ใช้ไม่ได้ */
function requireUser(req) {
  const username = readToken(req.token);
  if (!username) throw new Error('AUTH');
  const user = findUser(username);
  if (!user || String(user.active).toLowerCase() !== 'yes') throw new Error('AUTH');
  return user;
}

function requireRole(user, roles) {
  if (roles.indexOf(user.role) === -1) throw new Error('คุณไม่มีสิทธิ์ทำรายการนี้');
}

/* ============================================================
 * Router
 * ========================================================== */

function doGet() {
  return ContentService.createTextOutput('ระบบติดตามงานตัดกล่องตัวอย่าง — API พร้อมใช้งาน');
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const handler = HANDLERS[req.action];
    if (!handler) throw new Error('ไม่รู้จักคำสั่ง: ' + req.action);
    out = handler(req);
    if (out.ok === undefined) out.ok = true;
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

const HANDLERS = {
  login: function (req) {
    const user = findUser(req.username);
    const bad = { ok: false, error: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
    if (!user || String(user.active).toLowerCase() !== 'yes') return bad;
    if (hashPassword(req.password || '', user.salt) !== user.hash) return bad;
    writeLog(user.username, 'login', '', '');
    return {
      token: makeToken(user.username),
      user: { username: user.username, name: user.name, role: user.role },
    };
  },

  me: function (req) {
    const user = requireUser(req);
    return { user: { username: user.username, name: user.name, role: user.role } };
  },

  listJobs: function (req) {
    requireUser(req);
    return { jobs: readSheet(CFG.JOBS, JOB_COLS), today: today() };
  },

  createJob: function (req) {
    const user = requireUser(req);
    requireRole(user, [ROLE.CUTTER]);
    const j = req.job || {};
    if (!j.customer) throw new Error('กรุณากรอกชื่อลูกค้า');
    if (!j.destination) throw new Error('กรุณากรอกสถานที่จัดส่ง');

    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const sh = sheet(CFG.JOBS);
      const now = nowStr();
      const row = {
        id: nextJobId(sh),
        customer: j.customer,
        sales: j.sales || '',
        jobName: j.jobName || '',
        qty: j.qty || '',
        destination: j.destination,
        // ผู้ติดต่อ/เบอร์โทร เป็นหน้าที่ของฝ่ายจัดส่ง ไม่รับค่าจากช่างตัด
        contactName: '',
        contactPhone: '',
        dueDate: j.dueDate || '',
        status: STATUS.WAIT_CUT,
        note: j.note || '',
        createdBy: user.name,
        createdAt: now,
        cutBy: '', cutAt: '', vehicle: '', shipBy: '', shipAt: '',
        updatedAt: now,
      };
      sh.appendRow(JOB_COLS.map(function (c) { return row[c]; }));
      writeLog(user.username, 'สร้างงาน', row.id, row.customer);
      return { job: row };
    } finally {
      lock.releaseLock();
    }
  },

  updateJob: function (req) {
    const user = requireUser(req);
    const target = locateJob(req.id);
    const j = req.job || {};

    // ช่างตัด = ข้อมูลงานและปลายทาง / จัดส่ง = ข้อมูลการขนส่ง
    // ผู้ติดต่อ เบอร์โทร และทะเบียนรถ เป็นของฝ่ายจัดส่งเท่านั้น
    const allowed = user.role === ROLE.CUTTER
      ? ['customer', 'sales', 'jobName', 'qty', 'destination', 'dueDate', 'note']
      : user.role === ROLE.SHIPPING
        ? ['destination', 'contactName', 'contactPhone', 'vehicle', 'note']
        : [];
    if (!allowed.length) throw new Error('คุณไม่มีสิทธิ์แก้ไขข้อมูล');

    const patch = {};
    allowed.forEach(function (f) {
      if (j[f] !== undefined) patch[f] = j[f];
    });
    patch.updatedAt = nowStr();
    applyPatch(target, patch);
    writeLog(user.username, 'แก้ไขงาน', req.id, Object.keys(patch).join(','));
    return { job: readJobRow(target.row) };
  },

  setStatus: function (req) {
    const user = requireUser(req);
    const target = locateJob(req.id);
    const current = target.data.status;
    const next = req.status;
    const now = nowStr();
    const patch = { status: next, updatedAt: now };

    if (next === STATUS.CUT_DONE) {
      requireRole(user, [ROLE.CUTTER]);
      if (current !== STATUS.WAIT_CUT) throw new Error('งานนี้ไม่ได้อยู่ในสถานะรอตัด');
      patch.cutBy = user.name;
      patch.cutAt = now;
    } else if (next === STATUS.SHIPPED) {
      requireRole(user, [ROLE.SHIPPING]);
      if (current !== STATUS.CUT_DONE) throw new Error('ต้องตัดเสร็จก่อนจึงจะส่งได้');
      const vehicle = req.vehicle || target.data.vehicle;
      if (!vehicle) throw new Error('กรุณาระบุรถที่ใช้จัดส่ง');
      patch.vehicle = vehicle;
      if (req.contactName) patch.contactName = req.contactName;
      if (req.contactPhone) patch.contactPhone = req.contactPhone;
      patch.shipBy = user.name;
      patch.shipAt = now;
    } else if (next === STATUS.WAIT_CUT) {
      requireRole(user, [ROLE.CUTTER]);
      patch.cutBy = ''; patch.cutAt = '';
    } else if (next === STATUS.CANCELLED) {
      requireRole(user, [ROLE.CUTTER]);
    } else {
      throw new Error('สถานะไม่ถูกต้อง');
    }

    applyPatch(target, patch);
    writeLog(user.username, 'เปลี่ยนสถานะ → ' + next, req.id, '');
    return { job: readJobRow(target.row) };
  },

  changePassword: function (req) {
    const user = requireUser(req);
    if (hashPassword(req.oldPassword || '', user.salt) !== user.hash) {
      throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
    }
    if (!req.newPassword || req.newPassword.length < 6) {
      throw new Error('รหัสผ่านใหม่ต้องยาวอย่างน้อย 6 ตัวอักษร');
    }
    const salt = Utilities.getUuid();
    const sh = sheet(CFG.USERS);
    sh.getRange(user._row, USER_COLS.indexOf('salt') + 1).setValue(salt);
    sh.getRange(user._row, USER_COLS.indexOf('hash') + 1)
      .setValue(hashPassword(req.newPassword, salt));
    writeLog(user.username, 'เปลี่ยนรหัสผ่าน', '', '');
    return {};
  },
};

/* ============================================================
 * Sheet helpers
 * ========================================================== */

function readSheet(name, cols) {
  const sh = sheet(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, cols.length).getValues();
  return values.map(function (row) {
    const obj = {};
    cols.forEach(function (c, i) { obj[c] = cellToString(row[i]); });
    return obj;
  }).filter(function (o) { return o.id !== '' || name !== CFG.JOBS; });
}

function cellToString(v) {
  if (v instanceof Date) {
    // ถ้ามีเวลาติดมาด้วยให้เก็บเวลาไว้ ไม่งั้นคืนเฉพาะวันที่
    const hasTime = v.getHours() || v.getMinutes() || v.getSeconds();
    return Utilities.formatDate(v, CFG.TZ, hasTime ? 'yyyy-MM-dd HH:mm' : 'yyyy-MM-dd');
  }
  return v === null || v === undefined ? '' : String(v);
}

function locateJob(id) {
  const sh = sheet(CFG.JOBS);
  const last = sh.getLastRow();
  if (last < 2) throw new Error('ไม่พบงานนี้');
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      return { row: i + 2, data: readJobRow(i + 2) };
    }
  }
  throw new Error('ไม่พบงาน ' + id);
}

function readJobRow(row) {
  const values = sheet(CFG.JOBS).getRange(row, 1, 1, JOB_COLS.length).getValues()[0];
  const obj = {};
  JOB_COLS.forEach(function (c, i) { obj[c] = cellToString(values[i]); });
  return obj;
}

function applyPatch(target, patch) {
  const sh = sheet(CFG.JOBS);
  Object.keys(patch).forEach(function (key) {
    const idx = JOB_COLS.indexOf(key);
    if (idx >= 0) sh.getRange(target.row, idx + 1).setValue(patch[key]);
  });
}

function nextJobId(sh) {
  const prefix = 'SB-' + Utilities.formatDate(new Date(), CFG.TZ, 'yyMM') + '-';
  const last = sh.getLastRow();
  let max = 0;
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
      const id = String(r[0]);
      if (id.indexOf(prefix) === 0) {
        max = Math.max(max, Number(id.slice(prefix.length)) || 0);
      }
    });
  }
  return prefix + ('00' + (max + 1)).slice(-3);
}

function writeLog(username, action, jobId, detail) {
  try {
    sheet(CFG.LOG).appendRow([nowStr(), username, action, jobId, detail]);
  } catch (err) {
    // ไม่ให้ log ที่พังทำให้งานหลักล้มเหลว
  }
}

function nowStr() {
  return Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd HH:mm');
}

function today() {
  return Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd');
}

/* ============================================================
 * เครื่องมือสำหรับผู้ดูแล — รันจากเมนู Apps Script ได้เลย
 * ========================================================== */

/** เปลี่ยนรหัสผ่านให้ผู้ใช้คนใดคนหนึ่ง (แก้ค่าสองบรรทัดนี้แล้วกดรัน) */
function adminResetPassword() {
  const username = 'cutter';
  const newPassword = 'เปลี่ยนตรงนี้';

  const user = findUser(username);
  if (!user) throw new Error('ไม่พบผู้ใช้ ' + username);
  const salt = Utilities.getUuid();
  const sh = sheet(CFG.USERS);
  sh.getRange(user._row, USER_COLS.indexOf('salt') + 1).setValue(salt);
  sh.getRange(user._row, USER_COLS.indexOf('hash') + 1)
    .setValue(hashPassword(newPassword, salt));
  Logger.log('ตั้งรหัสผ่านใหม่ให้ ' + username + ' เรียบร้อย');
}

/**
 * ล้างข้อมูลงานทั้งหมด เริ่มนับรหัสงานใหม่ตั้งแต่ 001
 * บัญชีผู้ใช้และรหัสผ่านไม่ถูกแตะต้อง
 *
 * ⚠️ ลบแล้วกู้จากในระบบไม่ได้ — ถ้าจะกู้ต้องใช้ File → Version history ของ Google Sheet
 * วิธีใช้: เปลี่ยน CONFIRM เป็น true แล้วกด Run
 */
function adminClearJobs() {
  const CONFIRM = false;    // ← เปลี่ยนเป็น true เพื่อยืนยันว่าจะลบจริง
  const CLEAR_LOG = true;   // ล้างชีต Log ด้วยหรือไม่

  if (!CONFIRM) {
    throw new Error('ยังไม่ได้ยืนยัน — เปลี่ยน CONFIRM เป็น true ก่อนแล้วค่อยกด Run');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const jobs = sheet(CFG.JOBS);
    const jobRows = jobs.getLastRow() - 1;
    if (jobRows > 0) jobs.deleteRows(2, jobRows);

    let logRows = 0;
    if (CLEAR_LOG) {
      const log = sheet(CFG.LOG);
      logRows = log.getLastRow() - 1;
      if (logRows > 0) log.deleteRows(2, logRows);
    }

    Logger.log('ล้างข้อมูลเรียบร้อย — ลบงาน ' + Math.max(jobRows, 0) + ' รายการ, ' +
      'ลบ log ' + Math.max(logRows, 0) + ' บรรทัด\n' +
      'รหัสงานถัดไปจะเริ่มที่ ' + nextJobId(jobs) + '\n' +
      'อย่าลืมเปลี่ยน CONFIRM กลับเป็น false');
  } finally {
    lock.releaseLock();
  }
}

/** เพิ่มผู้ใช้ใหม่ (แก้ค่าแล้วกดรัน) */
function adminAddUser() {
  const username = 'newuser';
  const name = 'ชื่อที่แสดง';
  const role = ROLE.VIEWER; // ROLE.CUTTER | ROLE.SHIPPING | ROLE.VIEWER
  const password = randomPassword();

  if (findUser(username)) throw new Error('มีชื่อผู้ใช้นี้อยู่แล้ว');
  const salt = Utilities.getUuid();
  sheet(CFG.USERS).appendRow([username, name, role, salt, hashPassword(password, salt), 'yes']);
  Logger.log('สร้างผู้ใช้ ' + username + ' รหัสผ่าน: ' + password);
}

/**
 * ระบบติดตามงานตัดกล่องตัวอย่าง — Backend
 * วางไฟล์นี้ใน Google Apps Script ที่ผูกกับ Google Sheet
 *
 * ขั้นตอนใช้งานครั้งแรก: รันฟังก์ชัน setup() หนึ่งครั้ง แล้วดูรหัสผ่านใน Log
 *
 * หมายเหตุ: ไฟล์นี้ใช้ MailApp ส่งอีเมลแจ้งเตือน ถ้าเพิ่งวางโค้ดใหม่ทับของเก่า
 * ต้องรันฟังก์ชันจากหน้า Editor หนึ่งครั้งเพื่อกด Allow ให้สิทธิ์ส่งเมลก่อน
 */

const CFG = {
  JOBS: 'Jobs',
  USERS: 'Users',
  AES: 'AEs',
  LOG: 'Log',
  TZ: 'Asia/Bangkok',
  TOKEN_TTL_H: 12,
  WEB_URL: 'https://tcrbsamplecutter-sys.github.io/tracking-status/',
};

const STATUS = {
  PENDING_SHIP: 'รอส่ง',
  SHIPPED: 'ส่งแล้ว',
  CANCELLED: 'ยกเลิก',
};

const ROLE = { CUTTER: 'cutter', SHIPPING: 'shipping', VIEWER: 'viewer' };

const JOB_COLS = [
  'id', 'customer', 'fileCode', 'rscTele', 'dieCut', 'accessory',
  'flute', 'qty', 'aeName', 'aeEmail', 'dueDate', 'destination',
  'status', 'note',
  'contactName', 'contactPhone', 'vehicle',
  'createdBy', 'createdAt', 'shipBy', 'shipAt', 'updatedAt',
];

const USER_COLS = ['username', 'name', 'role', 'salt', 'hash', 'active'];
const AE_COLS = ['name', 'email', 'active'];
const LOG_COLS = ['time', 'username', 'action', 'jobId', 'detail'];

/* ============================================================
 * ติดตั้งครั้งแรก
 * ========================================================== */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(CFG.TZ);

  ensureSheet(CFG.JOBS, JOB_COLS);
  ensureSheet(CFG.USERS, USER_COLS);
  ensureSheet(CFG.AES, AE_COLS);
  ensureSheet(CFG.LOG, LOG_COLS);

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SECRET')) {
    props.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  }
  if (!props.getProperty('NOTIFY_EMAIL')) {
    props.setProperty('NOTIFY_EMAIL', 'nawinsan@scg.com');
  }

  const notes = [
    '',
    '===== สิ่งที่ต้องทำต่อ =====',
    '1) เปิดชีต "AEs" แล้วกรอกชื่อ AE กับอีเมลให้ครบ (คอลัมน์ active ใส่ yes)',
    '   ถ้าชีตนี้ว่าง ช่างตัดจะบันทึกงานไม่ได้ เพราะดรอปดาวน์ AE ไม่มีตัวเลือก',
    '2) อีเมลที่รับแจ้งเตือน "รอส่ง" ตอนนี้คือ: ' + notifyEmail(),
    '   แก้ได้ที่ Project Settings → Script Properties → NOTIFY_EMAIL (ไม่ต้อง Deploy ใหม่)',
    '3) โควตาส่งเมลที่เหลือวันนี้: ' + mailQuota() + ' ฉบับ',
  ];

  const users = sheet(CFG.USERS);
  if (users.getLastRow() > 1) {
    Logger.log(['มีผู้ใช้อยู่แล้ว ข้ามการสร้างบัญชีเริ่มต้น'].concat(notes).join('\n'));
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
  Logger.log(lines.concat(notes).join('\n'));
}

/**
 * สร้างชีตถ้ายังไม่มี และเขียนหัวตารางใหม่เมื่อไม่ตรงกับโค้ด
 * (จำเป็นเวลาปรับโครงคอลัมน์ เพราะกด Save แล้วรัน setup() ซ้ำต้องอัพเดตหัวตารางให้ด้วย)
 */
function ensureSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  const width = sh.getLastColumn();
  const current = (sh.getLastRow() === 0 || width === 0) ? []
    : sh.getRange(1, 1, 1, width).getValues()[0].map(function (v) { return String(v); });

  const same = current.length === headers.length &&
    headers.every(function (h, i) { return current[i] === h; });

  if (!same) {
    if (sh.getLastRow() > 1) {
      Logger.log('⚠️ ชีต ' + name + ' มีข้อมูล ' + (sh.getLastRow() - 1) +
        ' แถว แต่หัวตารางไม่ตรงกับโค้ด — เขียนหัวตารางใหม่แล้ว ' +
        'กรุณาตรวจว่าข้อมูลเดิมยังอยู่ตรงคอลัมน์');
    }
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
 * AE — อีเมลถูก resolve ที่ฝั่งนี้เท่านั้น ไม่รับค่าจากหน้าเว็บ
 * เพื่อไม่ให้ยิงคำสั่งตรงแล้วสั่งส่งเมลไปที่อื่นได้ และกันอีเมลพิมพ์ผิด
 * ========================================================== */

function activeAes() {
  return readSheet(CFG.AES, AE_COLS)
    .filter(function (a) {
      return String(a.name).trim() !== '' && String(a.active).toLowerCase() !== 'no';
    })
    .map(function (a) {
      return { name: String(a.name).trim(), email: String(a.email).trim() };
    });
}

function findAe(name) {
  const target = String(name || '').trim();
  if (!target) return null;
  const list = activeAes();
  for (let i = 0; i < list.length; i++) {
    if (list[i].name === target) return list[i];
  }
  return null;
}

/** หา AE หรือโยน error ที่บอกสาเหตุชัดเจน */
function requireAe(name) {
  if (!String(name || '').trim()) throw new Error('กรุณาเลือก AE');
  const ae = findAe(name);
  if (!ae) throw new Error('ไม่พบ AE ชื่อ "' + name + '" ในชีต AEs');
  if (!ae.email) throw new Error('AE "' + ae.name + '" ยังไม่มีอีเมลในชีต AEs');
  return ae;
}

/** ทุกแถวรวมที่ปิดใช้งานแล้ว — ใช้ในหน้าตั้งค่าซึ่งต้องเห็นครบเพื่อเปิดกลับได้ */
function allAes() {
  return readSheet(CFG.AES, AE_COLS)
    .filter(function (a) { return String(a.name).trim() !== ''; })
    .map(function (a) {
      return {
        name: String(a.name).trim(),
        email: String(a.email).trim(),
        active: String(a.active).toLowerCase() !== 'no',
      };
    });
}

/** เลขแถวจริงในชีตของ AE ชื่อนี้ (0 = ไม่พบ) — ข้อมูลเริ่มแถว 2 เพราะแถว 1 เป็นหัวตาราง */
function findAeRow(name) {
  const target = String(name || '').trim();
  if (!target) return 0;
  const rows = readSheet(CFG.AES, AE_COLS);
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i].name).trim() === target) return i + 2;
  }
  return 0;
}

/**
 * แยกอีเมลที่คั่นด้วยจุลภาค ตรวจรูปแบบ แล้วคืนสตริงที่จัดรูปแล้ว
 * ใช้ร่วมกันทั้งอีเมลแจ้งเตือนและอีเมล AE จะได้ตรวจด้วยเกณฑ์เดียวกัน
 */
function parseEmails(text) {
  const list = String(text || '').split(',')
    .map(function (s) { return s.trim(); })
    .filter(function (s) { return s !== ''; });

  if (!list.length) throw new Error('กรุณากรอกอีเมล');

  const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  list.forEach(function (e) {
    if (!re.test(e)) throw new Error('รูปแบบอีเมลไม่ถูกต้อง: ' + e);
  });
  return list.join(', ');
}

/* ============================================================
 * อีเมลแจ้งเตือน
 * ========================================================== */

function notifyEmail() {
  return PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL') || '';
}

function mailQuota() {
  try {
    return MailApp.getRemainingDailyQuota();
  } catch (err) {
    return 'อ่านไม่ได้ (ยังไม่ได้ให้สิทธิ์ส่งเมล)';
  }
}

/** ส่งเมลแบบไม่ให้ความล้มเหลวลามไปทำให้งานหลักพัง — แนวเดียวกับ writeLog() */
function sendMailSafe(to, subject, body) {
  if (!to) return;
  try {
    MailApp.sendEmail({
      to: to, subject: subject, body: body,
      name: 'ระบบงานตัดกล่องตัวอย่าง',
    });
  } catch (err) {
    writeLog('system', 'ส่งเมลไม่สำเร็จ', '', to + ' — ' + String(err.message || err));
  }
}

/** รายละเอียดงานแบบข้อความ ตัดบรรทัดที่ไม่มีค่าออก */
function jobLines(job) {
  return [
    ['รหัสงาน', job.id],
    ['ลูกค้า', job.customer],
    ['รหัสไฟล์', job.fileCode],
    ['RSC / Tele', job.rscTele],
    ['DieCut', job.dieCut],
    ['Accessory', job.accessory],
    ['ลอน', job.flute],
    ['จำนวน / ใบ', job.qty],
    ['AE', job.aeName],
    ['กำหนดส่ง', job.dueDate],
    ['ปลายทาง', job.destination],
    ['หมายเหตุ', job.note],
  ].filter(function (r) {
    return String(r[1] === undefined || r[1] === null ? '' : r[1]) !== '';
  }).map(function (r) {
    return r[0] + ': ' + r[1];
  }).join('\n');
}

/** แจ้งฝ่ายจัดส่งว่ามีงานรอส่ง */
function notifyPendingShip(job) {
  const body = [
    'มีงานตัดกล่องตัวอย่างรอจัดส่ง',
    '',
    jobLines(job),
    '',
    'บันทึกโดย: ' + job.createdBy + '  (' + job.createdAt + ')',
    '',
    'เปิดดูรายการทั้งหมด: ' + CFG.WEB_URL,
  ].join('\n');

  sendMailSafe(notifyEmail(), '[รอส่ง] ' + job.id + ' ' + job.customer, body);
}

/** แจ้ง AE ว่างานของตัวเองส่งออกแล้ว */
function notifyShipped(job) {
  const ship = [['ทะเบียนรถ', job.vehicle], ['ผู้ติดต่อ', job.contactName],
                ['เบอร์ติดต่อ', job.contactPhone], ['ส่งโดย', job.shipBy],
                ['เวลาส่ง', job.shipAt]]
    .filter(function (r) { return String(r[1] || '') !== ''; })
    .map(function (r) { return r[0] + ': ' + r[1]; }).join('\n');

  const body = [
    'งานตัดกล่องตัวอย่างที่คุณสั่งไว้ จัดส่งเรียบร้อยแล้ว',
    '',
    jobLines(job),
    '',
    '--- ข้อมูลการจัดส่ง ---',
    ship,
    '',
    'เปิดดูรายการทั้งหมด: ' + CFG.WEB_URL,
  ].join('\n');

  sendMailSafe(job.aeEmail, '[ส่งแล้ว] ' + job.id + ' ' + job.customer, body);
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
    return {
      jobs: readSheet(CFG.JOBS, JOB_COLS),
      aes: activeAes(),
      today: today(),
    };
  },

  createJob: function (req) {
    const user = requireUser(req);
    requireRole(user, [ROLE.CUTTER]);
    const j = req.job || {};
    if (!j.customer) throw new Error('กรุณากรอกชื่อลูกค้า');
    if (!j.destination) throw new Error('กรุณากรอกสถานที่จัดส่ง');

    // อีเมลมาจากชีต AEs เท่านั้น ค่า aeEmail ที่ส่งมาจากหน้าเว็บถูกทิ้งทั้งหมด
    const ae = requireAe(j.aeName);

    let row;
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const sh = sheet(CFG.JOBS);
      const now = nowStr();
      row = {
        id: nextJobId(sh),
        customer: j.customer,
        fileCode: j.fileCode || '',
        rscTele: j.rscTele || '',
        dieCut: j.dieCut || '',
        accessory: j.accessory || '',
        flute: j.flute || '',
        qty: j.qty || '',
        aeName: ae.name,
        aeEmail: ae.email,
        dueDate: j.dueDate || '',
        destination: j.destination,
        // งานถูกบันทึกหลังตัดเสร็จแล้ว จึงเข้าสถานะรอส่งทันที
        status: STATUS.PENDING_SHIP,
        note: j.note || '',
        // ผู้ติดต่อ/เบอร์โทร/ทะเบียนรถ เป็นหน้าที่ของฝ่ายจัดส่ง ไม่รับค่าจากช่างตัด
        contactName: '',
        contactPhone: '',
        vehicle: '',
        createdBy: user.name,
        createdAt: now,
        shipBy: '', shipAt: '',
        updatedAt: now,
      };
      sh.appendRow(JOB_COLS.map(function (c) { return row[c]; }));
      writeLog(user.username, 'บันทึกงาน', row.id, row.customer);
    } finally {
      lock.releaseLock();
    }

    // ส่งเมลหลังปล่อย lock แล้ว เพราะ MailApp ใช้เวลาราว 1-2 วินาที
    notifyPendingShip(row);
    return { job: row };
  },

  updateJob: function (req) {
    const user = requireUser(req);
    const target = locateJob(req.id);
    const j = req.job || {};

    // ช่างตัด = ข้อมูลงานและปลายทาง / จัดส่ง = ข้อมูลการขนส่ง
    // ผู้ติดต่อ เบอร์โทร และทะเบียนรถ เป็นของฝ่ายจัดส่งเท่านั้น
    const allowed = user.role === ROLE.CUTTER
      ? ['customer', 'fileCode', 'rscTele', 'dieCut', 'accessory', 'flute',
         'qty', 'aeName', 'dueDate', 'destination', 'note']
      : user.role === ROLE.SHIPPING
        ? ['destination', 'contactName', 'contactPhone', 'vehicle', 'note']
        : [];
    if (!allowed.length) throw new Error('คุณไม่มีสิทธิ์แก้ไขข้อมูล');

    const patch = {};
    allowed.forEach(function (f) {
      if (j[f] !== undefined) patch[f] = j[f];
    });

    // เปลี่ยน AE แล้วต้องไปหาอีเมลใหม่จากชีต AEs ไม่เชื่อค่าจากหน้าเว็บ
    if (patch.aeName !== undefined && patch.aeName !== target.data.aeName) {
      const ae = requireAe(patch.aeName);
      patch.aeName = ae.name;
      patch.aeEmail = ae.email;
    }

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
    let mailAe = false;

    if (next === STATUS.SHIPPED) {
      requireRole(user, [ROLE.SHIPPING]);
      if (current !== STATUS.PENDING_SHIP) throw new Error('งานนี้ไม่ได้อยู่ในสถานะรอส่ง');
      const vehicle = req.vehicle || target.data.vehicle;
      if (!vehicle) throw new Error('กรุณาระบุรถที่ใช้จัดส่ง');
      patch.vehicle = vehicle;
      if (req.contactName) patch.contactName = req.contactName;
      if (req.contactPhone) patch.contactPhone = req.contactPhone;
      patch.shipBy = user.name;
      patch.shipAt = now;
      mailAe = true;
    } else if (next === STATUS.PENDING_SHIP) {
      // ใช้เปิดงานที่ยกเลิกไปแล้วกลับมา
      requireRole(user, [ROLE.CUTTER]);
      patch.shipBy = ''; patch.shipAt = '';
    } else if (next === STATUS.CANCELLED) {
      requireRole(user, [ROLE.CUTTER]);
    } else {
      throw new Error('สถานะไม่ถูกต้อง');
    }

    applyPatch(target, patch);
    writeLog(user.username, 'เปลี่ยนสถานะ → ' + next, req.id, '');

    const job = readJobRow(target.row);
    if (mailAe) notifyShipped(job);
    return { job: job };
  },

  /* ---------- หน้าตั้งค่า — เฉพาะช่างตัด ---------- */

  getSettings: function (req) {
    const user = requireUser(req);
    requireRole(user, [ROLE.CUTTER]);
    return { notifyEmail: notifyEmail(), aes: allAes() };
  },

  setNotifyEmail: function (req) {
    const user = requireUser(req);
    requireRole(user, [ROLE.CUTTER]);

    const next = parseEmails(req.email);
    const prev = notifyEmail();
    PropertiesService.getScriptProperties().setProperty('NOTIFY_EMAIL', next);

    // เปลี่ยนปลายทางอีเมล = เปลี่ยนทางเดินข้อมูลลูกค้า ต้องสาวกลับได้ว่าใครเปลี่ยน
    writeLog(user.username, 'เปลี่ยนอีเมลแจ้งเตือน', '', (prev || '(ว่าง)') + ' → ' + next);
    return { notifyEmail: next };
  },

  /** originalName ว่าง = เพิ่มใหม่ · มีค่า = แก้แถวเดิม (รองรับการเปลี่ยนชื่อ) */
  saveAe: function (req) {
    const user = requireUser(req);
    requireRole(user, [ROLE.CUTTER]);

    const name = String(req.name || '').trim();
    if (!name) throw new Error('กรุณากรอกชื่อ AE');
    const email = parseEmails(req.email);
    const original = String(req.originalName || '').trim();
    const active = req.active === false ? 'no' : 'yes';

    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const sh = sheet(CFG.AES);
      const row = original ? findAeRow(original) : 0;
      if (original && !row) throw new Error('ไม่พบ AE ชื่อ "' + original + '"');

      const clash = findAeRow(name);
      if (clash && clash !== row) throw new Error('มีชื่อ "' + name + '" อยู่แล้ว');

      if (!row) {
        sh.appendRow([name, email, active]);
        writeLog(user.username, 'เพิ่ม AE', '', name + ' / ' + email);
      } else {
        const before = sh.getRange(row, 1, 1, AE_COLS.length).getValues()[0];
        sh.getRange(row, 1, 1, AE_COLS.length).setValues([[name, email, active]]);
        writeLog(user.username, 'แก้ AE', '',
          before[0] + ' / ' + before[1] + '  →  ' + name + ' / ' + email);
      }
    } finally {
      lock.releaseLock();
    }
    return { aes: allAes() };
  },

  /**
   * เปิด/ปิดการใช้งาน AE — ไม่มีการลบแถว เพราะงานเก่าเก็บชื่อ AE ไว้เป็น snapshot
   * ลบทิ้งแล้วงานเก่าจะกลายเป็นชื่อลอยที่สาวกลับไม่ได้
   */
  setAeActive: function (req) {
    const user = requireUser(req);
    requireRole(user, [ROLE.CUTTER]);

    const name = String(req.name || '').trim();
    const row = findAeRow(name);
    if (!row) throw new Error('ไม่พบ AE ชื่อ "' + name + '"');

    const active = req.active ? 'yes' : 'no';
    sheet(CFG.AES).getRange(row, AE_COLS.indexOf('active') + 1).setValue(active);
    writeLog(user.username, 'ตั้งสถานะ AE', '',
      name + ' → ' + (active === 'yes' ? 'ใช้งาน' : 'ปิดใช้งาน'));
    return { aes: allAes() };
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
 * บัญชีผู้ใช้ รายชื่อ AE และรหัสผ่านไม่ถูกแตะต้อง
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

/**
 * ตรวจว่าการตั้งค่าอีเมลพร้อมใช้งานไหม — รันได้ทุกเมื่อ ไม่แก้ข้อมูลอะไร
 * ใช้เช็คหลังวางโค้ดใหม่ว่าให้สิทธิ์ MailApp แล้วและชีต AEs กรอกครบ
 */
function adminCheckMailSetup() {
  const aes = activeAes();
  const missing = aes.filter(function (a) { return !a.email; });

  const lines = [
    '===== ตรวจการตั้งค่าอีเมล =====',
    'ผู้รับแจ้งเตือน "รอส่ง" (NOTIFY_EMAIL): ' + (notifyEmail() || '⚠️ ยังไม่ได้ตั้งค่า'),
    'โควตาส่งเมลที่เหลือวันนี้: ' + mailQuota(),
    'จำนวน AE ที่ใช้งานได้: ' + aes.length + (aes.length ? '' : '  ⚠️ ชีต AEs ว่าง บันทึกงานไม่ได้'),
  ];

  aes.forEach(function (a) {
    lines.push('  · ' + a.name + '  →  ' + (a.email || '⚠️ ไม่มีอีเมล'));
  });
  if (missing.length) {
    lines.push('⚠️ มี AE ' + missing.length + ' คนที่ไม่มีอีเมล จะบันทึกงานให้คนนั้นไม่ได้');
  }

  Logger.log(lines.join('\n'));
}

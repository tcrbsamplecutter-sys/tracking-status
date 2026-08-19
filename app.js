/* ============================================================
 * ระบบติดตามงานตัดกล่องตัวอย่าง — Frontend
 * ========================================================== */

const STATUS = {
  PENDING: 'รอส่ง',
  SHIP: 'ส่งแล้ว',
  CANCEL: 'ยกเลิก',
};

const STATUS_CLASS = {
  'รอส่ง': 's-cut',
  'ส่งแล้ว': 's-ship',
  'ยกเลิก': 's-cancel',
};

const ROLE_LABEL = { cutter: 'ช่างตัด', shipping: 'จัดส่ง', viewer: 'ดูอย่างเดียว' };

const VIEW_TITLE = { jobs: 'รายการงาน', dashboard: 'Dashboard', settings: 'ตั้งค่า' };

/**
 * ฟิลด์ที่แต่ละบทบาทแก้ไขได้
 * ช่างตัด = ข้อมูลงานและปลายทาง / จัดส่ง = ข้อมูลการขนส่ง
 * ผู้ติดต่อ เบอร์โทร ทะเบียนรถ เป็นของฝ่ายจัดส่งเท่านั้น
 * (Apps Script ตรวจซ้ำอีกชั้น การซ่อนตรงนี้เป็นแค่เรื่องความสะดวก)
 */
const EDITABLE = {
  cutter: ['customer', 'fileCode', 'rscTele', 'dieCut', 'accessory', 'flute',
           'qty', 'aeName', 'dueDate', 'destination', 'note'],
  shipping: ['destination', 'contactName', 'contactPhone', 'vehicle', 'note'],
  viewer: [],
};

const state = {
  token: localStorage.getItem('token') || '',
  user: JSON.parse(localStorage.getItem('user') || 'null'),
  jobs: [],
  aes: [],
  today: '',
  filter: 'all',
  search: '',
  view: 'jobs',
  editingId: null,
  dashSearch: '',
  dashBasis: 'createdAt',
  settings: {
    notifyEmail: '', ccEmail: '',
    ccOnPending: true, ccOnShipped: true,
    aes: [],
  },
};

const EMPTY_SETTINGS = JSON.parse(JSON.stringify(state.settings));

const $ = function (sel) { return document.querySelector(sel); };
const $$ = function (sel) { return Array.prototype.slice.call(document.querySelectorAll(sel)); };

/* ============================================================
 * API
 * ========================================================== */

async function api(action, payload) {
  const body = JSON.stringify(Object.assign({ action: action, token: state.token }, payload || {}));

  let res;
  try {
    // ไม่ตั้ง Content-Type เพื่อให้เป็น simple request (เลี่ยง CORS preflight ของ Apps Script)
    res = await fetch(CONFIG.API_URL, { method: 'POST', body: body, redirect: 'follow' });
  } catch (err) {
    throw new Error('เชื่อมต่อเซิร์ฟเวอร์ไม่ได้ กรุณาตรวจสอบอินเทอร์เน็ตหรือ URL ใน config.js');
  }

  let data;
  try {
    data = await res.json();
  } catch (err) {
    throw new Error('เซิร์ฟเวอร์ตอบกลับผิดรูปแบบ — ตรวจสอบว่า Deploy Apps Script เป็นแบบ "ทุกคน" แล้ว');
  }

  if (!data.ok) {
    if (data.error === 'AUTH') {
      logout();
      throw new Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
    }
    throw new Error(data.error || 'เกิดข้อผิดพลาด');
  }
  return data;
}

/* ============================================================
 * Utilities
 * ========================================================== */

function esc(v) {
  return String(v === undefined || v === null ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function toast(msg, isError) {
  const el = $('#toast');
  el.textContent = msg;
  el.className = 'toast' + (isError ? ' err' : '');
  clearTimeout(el._timer);
  el._timer = setTimeout(function () { el.classList.add('hidden'); }, 2600);
}

function thaiDate(iso) {
  if (!iso) return '';
  const parts = String(iso).slice(0, 10).split('-');
  if (parts.length !== 3) return iso;
  return parts[2] + '/' + parts[1] + '/' + (Number(parts[0]) + 543).toString().slice(-2);
}

function isToday(stamp) {
  return String(stamp || '').slice(0, 10) === state.today;
}

function dueClass(job) {
  if (!job.dueDate || job.status === STATUS.SHIP || job.status === STATUS.CANCEL) return '';
  const due = String(job.dueDate).slice(0, 10);
  if (due < state.today) return 'due-late';
  if (due === state.today) return 'due-warn';
  return '';
}

function can(action) {
  const role = state.user ? state.user.role : '';
  if (action === 'create') return role === 'cutter';
  if (action === 'markShip') return role === 'shipping';
  if (action === 'cancel') return role === 'cutter';
  if (action === 'manage') return role === 'cutter';
  return false;
}

/** ใช้ร่วมกันทั้งหน้ารายการงานและ Dashboard จะได้ค้นด้วยเกณฑ์เดียวกัน */
function matchesSearch(job, term) {
  if (!term) return true;
  return [job.id, job.customer, job.fileCode, job.aeName, job.destination,
          job.contactName, job.vehicle].join(' ').toLowerCase().indexOf(term) !== -1;
}

/* ============================================================
 * Login
 * ========================================================== */

$('#loginForm').addEventListener('submit', async function (e) {
  e.preventDefault();
  const btn = $('#loginBtn');
  const err = $('#loginError');
  err.classList.add('hidden');
  btn.disabled = true;
  btn.textContent = 'กำลังเข้าสู่ระบบ...';

  try {
    const data = await api('login', {
      username: $('#username').value.trim(),
      password: $('#password').value,
    });
    state.token = data.token;
    state.user = data.user;
    localStorage.setItem('token', data.token);
    localStorage.setItem('user', JSON.stringify(data.user));
    $('#password').value = '';
    await enterApp();
  } catch (ex) {
    err.textContent = ex.message;
    err.classList.remove('hidden');
  } finally {
    btn.disabled = false;
    btn.textContent = 'เข้าสู่ระบบ';
  }
});

function logout() {
  state.token = '';
  state.user = null;
  state.jobs = [];
  state.aes = [];
  state.settings = JSON.parse(JSON.stringify(EMPTY_SETTINGS));
  localStorage.removeItem('token');
  localStorage.removeItem('user');
  $('#appView').classList.add('hidden');
  $('#loginView').classList.remove('hidden');
  closeModal();
}

async function enterApp() {
  $('#loginView').classList.add('hidden');
  $('#appView').classList.remove('hidden');
  $('#userName').textContent = state.user.name;
  $('#userRole').textContent = ROLE_LABEL[state.user.role] || state.user.role;

  // เมนูตั้งค่าเห็นเฉพาะช่างตัด (เซิร์ฟเวอร์ตรวจซ้ำอีกชั้นอยู่แล้ว)
  $$('[data-cutter-only]').forEach(function (el) {
    el.classList.toggle('hidden', !can('manage'));
  });
  if (state.view === 'settings' && !can('manage')) state.view = 'jobs';

  setView(state.view);
  await loadJobs();
}

function setView(view) {
  state.view = view;
  $$('[data-view]').forEach(function (b) {
    b.classList.toggle('active', b.dataset.view === view);
  });
  $('#jobsView').classList.toggle('hidden', view !== 'jobs');
  $('#dashboardView').classList.toggle('hidden', view !== 'dashboard');
  $('#settingsView').classList.toggle('hidden', view !== 'settings');
  $('#pageTitle').textContent = VIEW_TITLE[view] || '';
  $('#newJobBtn').classList.toggle('hidden', !can('create') || view !== 'jobs');

  if (view === 'settings') loadSettings();
}

/* ============================================================
 * โหลดและแสดงผล
 * ========================================================== */

async function loadJobs() {
  try {
    const data = await api('listJobs');
    state.jobs = data.jobs;
    state.aes = data.aes || [];
    state.today = data.today;
    render();
  } catch (ex) {
    toast(ex.message, true);
  }
}

function render() {
  renderStats();
  renderJobList();
  renderDashboard();
}

function renderStats() {
  const jobs = state.jobs;
  const count = function (fn) { return jobs.filter(fn).length; };

  const pendingShip = count(function (j) { return j.status === STATUS.PENDING; });
  const shippedToday = count(function (j) {
    return j.status === STATUS.SHIP && isToday(j.shipAt);
  });
  const late = count(function (j) {
    return j.dueDate && String(j.dueDate).slice(0, 10) < state.today
      && j.status !== STATUS.SHIP && j.status !== STATUS.CANCEL;
  });
  const loggedToday = count(function (j) { return isToday(j.createdAt); });

  $('#stats').innerHTML = [
    { cls: 'blue', num: pendingShip, label: 'รอจัดส่ง' },
    { cls: 'green', num: shippedToday, label: 'ส่งแล้ววันนี้' },
    { cls: 'red', num: late, label: 'เลยกำหนด' },
    { cls: 'amber', num: loggedToday, label: 'บันทึกวันนี้' },
  ].map(function (s) {
    return '<div class="stat ' + s.cls + '">' +
      '<div class="stat-num">' + s.num + '</div>' +
      '<div class="stat-label">' + s.label + '</div></div>';
  }).join('');
}

function matchesFilter(job) {
  if (state.filter !== 'all' && job.status !== state.filter) return false;
  return matchesSearch(job, state.search);
}

function jobCard(job, compact) {
  const cls = STATUS_CLASS[job.status] || 's-cancel';
  const bits = [];

  if (job.fileCode) bits.push('<span>รหัสไฟล์: <b>' + esc(job.fileCode) + '</b></span>');
  if (job.flute) bits.push('<span>ลอน: <b>' + esc(job.flute) + '</b></span>');
  if (job.qty) bits.push('<span>จำนวน: <b>' + esc(job.qty) + '</b></span>');

  const counts = [['RSC', job.rscTele], ['DieCut', job.dieCut], ['Acc', job.accessory]]
    .filter(function (c) { return String(c[1] || '') !== ''; })
    .map(function (c) { return c[0] + ' ' + esc(c[1]); });
  if (counts.length) bits.push('<span>' + counts.join(' · ') + '</span>');

  if (job.aeName) bits.push('<span>AE: <b>' + esc(job.aeName) + '</b></span>');
  if (!compact && job.destination) {
    bits.push('<span>ส่งที่: <b>' + esc(job.destination) + '</b></span>');
  }
  if (job.dueDate) {
    bits.push('<span class="' + dueClass(job) + '">กำหนดส่ง: <b>' +
      thaiDate(job.dueDate) + '</b></span>');
  }
  if (job.status === STATUS.SHIP && job.vehicle) {
    bits.push('<span>รถ: <b>' + esc(job.vehicle) + '</b></span>');
  }

  return '<div class="job-card ' + cls + '" data-id="' + esc(job.id) + '">' +
    '<div class="job-top"><div>' +
      '<div class="job-customer">' + esc(job.customer) + '</div>' +
      '<div class="job-id">' + esc(job.id) + '</div>' +
    '</div>' +
    '<span class="badge ' + cls + '">' + esc(job.status) + '</span>' +
    '</div>' +
    '<div class="job-meta">' + bits.join('') + '</div></div>';
}

function renderJobList() {
  const list = state.jobs.filter(matchesFilter).slice().reverse();
  $('#jobList').innerHTML = list.map(function (j) { return jobCard(j, false); }).join('');
  $('#emptyState').classList.toggle('hidden', list.length > 0);
}

/* ============================================================
 * Dashboard
 * ========================================================== */

/**
 * จัดกลุ่มงานตามวัน เรียงวันใหม่สุดขึ้นก่อน
 * งานที่ไม่มีค่าวันที่ในโหมดนั้นถูกแยกไว้ใน noDate เพื่อไม่ให้หายไปจากจอเงียบๆ
 * (สำคัญมากในโหมด "ยึดวันจัดส่ง" เพราะงานที่ยังไม่ส่งไม่มี shipAt)
 */
function groupByDay(jobs, basis) {
  const map = {};
  const noDate = [];

  jobs.forEach(function (j) {
    const day = String(j[basis] || '').slice(0, 10);
    if (!day) { noDate.push(j); return; }
    if (!map[day]) map[day] = [];
    map[day].push(j);
  });

  const groups = Object.keys(map).sort().reverse().map(function (day) {
    return { day: day, jobs: map[day] };
  });
  return { groups: groups, noDate: noDate };
}

function dayCounts(jobs) {
  const c = { pend: 0, ship: 0, cancel: 0 };
  jobs.forEach(function (j) {
    if (j.status === STATUS.PENDING) c.pend++;
    else if (j.status === STATUS.SHIP) c.ship++;
    else if (j.status === STATUS.CANCEL) c.cancel++;
  });
  return c;
}

function subBlock(title, cls, jobs) {
  if (!jobs.length) return '';
  return '<div class="day-sub">' +
    '<div class="sub-title ' + cls + '">' + title + ' (' + jobs.length + ')</div>' +
    '<div class="job-list compact">' +
      jobs.map(function (j) { return jobCard(j, true); }).join('') +
    '</div></div>';
}

function dayBlock(label, jobs, splitByStatus) {
  const c = dayCounts(jobs);
  const counts = [];
  if (c.pend) counts.push('<span class="c-pend">● รอส่ง ' + c.pend + '</span>');
  if (c.ship) counts.push('<span class="c-ship">✓ ส่งแล้ว ' + c.ship + '</span>');
  if (c.cancel) counts.push('<span class="c-cancel">ยกเลิก ' + c.cancel + '</span>');

  const byStatus = function (s) {
    return jobs.filter(function (j) { return j.status === s; });
  };

  // แยกรอส่ง/ส่งแล้วเฉพาะตอนที่กลุ่มนั้นปนกันจริง
  // โหมดยึดวันจัดส่ง ทุกงานในกลุ่มรายวันคือส่งแล้วทั้งหมด แยกไปก็ไม่ได้ความ
  const body = splitByStatus
    ? subBlock('● รอส่ง', 'pend', byStatus(STATUS.PENDING)) +
      subBlock('✓ ส่งแล้ว', 'ship', byStatus(STATUS.SHIP)) +
      subBlock('ยกเลิก', 'cancel', byStatus(STATUS.CANCEL))
    : '<div class="job-list compact">' +
        jobs.map(function (j) { return jobCard(j, true); }).join('') + '</div>';

  return '<section class="day-group">' +
    '<div class="day-head">' +
      '<span class="day-date">' + esc(label) + '</span>' +
      '<span class="day-total">รวม ' + jobs.length + ' รายการ</span>' +
      '<span class="day-counts">' + counts.join('') + '</span>' +
    '</div>' + body + '</section>';
}

function renderDashboard() {
  const basis = state.dashBasis;
  const list = state.jobs.filter(function (j) { return matchesSearch(j, state.dashSearch); });
  const g = groupByDay(list, basis);
  const html = [];

  if (g.noDate.length) {
    html.push(dayBlock(basis === 'shipAt' ? 'ยังไม่ได้ส่ง' : 'ไม่ระบุวันที่', g.noDate, true));
  }
  g.groups.forEach(function (grp) {
    html.push(dayBlock(thaiDate(grp.day), grp.jobs, basis === 'createdAt'));
  });

  $('#dayGroups').innerHTML = html.join('');
  $('#dashEmpty').classList.toggle('hidden', html.length > 0);
}

/* ============================================================
 * หน้าตั้งค่า
 * ========================================================== */

async function loadSettings() {
  try {
    const data = await api('getSettings');
    state.settings = {
      notifyEmail: data.notifyEmail || '',
      ccEmail: data.ccEmail || '',
      ccOnPending: data.ccOnPending !== false,
      ccOnShipped: data.ccOnShipped !== false,
      aes: data.aes || [],
    };
    syncActiveAes();
    renderSettings();
  } catch (ex) {
    toast(ex.message, true);
  }
}

/** ให้ดรอปดาวน์ AE ในฟอร์มงานอัพเดตตามทันทีโดยไม่ต้องรีเฟรชหรือ login ใหม่ */
function syncActiveAes() {
  state.aes = state.settings.aes
    .filter(function (a) { return a.active; })
    .map(function (a) { return { name: a.name, email: a.email }; });
}

function aeRow(ae) {
  const isNew = !ae;
  const name = isNew ? '' : ae.name;
  const email = isNew ? '' : ae.email;
  const checked = isNew || ae.active ? ' checked' : '';

  return '<div class="ae-row" data-original="' + esc(name) + '">' +
    '<input class="ae-name" type="text" value="' + esc(name) + '" placeholder="ชื่อ AE">' +
    '<input class="ae-email" type="text" value="' + esc(email) + '" ' +
      'placeholder="name@scg.com" autocapitalize="none" spellcheck="false">' +
    '<label class="ae-active"><input type="checkbox"' + checked + '></label>' +
    '<button class="btn ae-save">บันทึก</button>' +
  '</div>';
}

function renderSettings() {
  const s = state.settings;
  $('#notifyEmail').value = s.notifyEmail;
  $('#ccEmail').value = s.ccEmail;
  $('#ccOnPending').checked = s.ccOnPending;
  $('#ccOnShipped').checked = s.ccOnShipped;

  ['#notifyError', '#ccError', '#aeError', '#clearError'].forEach(function (sel) {
    $(sel).classList.add('hidden');
  });

  $('#aeList').innerHTML = s.aes.length
    ? s.aes.map(aeRow).join('')
    : '<div class="empty">ยังไม่มีรายชื่อ AE — กด "เพิ่ม AE ใหม่" ด้านล่าง</div>';
}

/** แสดงข้อความผิดพลาดในกล่อง .alert ที่ระบุ — ใช้ร่วมทั้งหน้าตั้งค่าและบล็อกดึงข้อมูล */
function showError(sel, msg) {
  const el = $(sel);
  el.textContent = msg;
  el.classList.remove('hidden');
}

/** อีเมลแจ้งเตือนกับอีเมล CC ใช้ขั้นตอนบันทึกเหมือนกัน ต่างแค่ key กับช่องที่อ่าน */
async function saveEmailSetting(key, field, errSel, btnSel, okMsg) {
  const btn = $(btnSel);
  $(errSel).classList.add('hidden');
  btn.disabled = true;
  try {
    const data = await api('setEmailSetting', { key: key, email: $(field).value.trim() });
    $(field).value = data.value;
    if (key === 'CC_EMAIL') state.settings.ccEmail = data.value;
    else state.settings.notifyEmail = data.value;
    toast(okMsg);
  } catch (ex) {
    showError(errSel, ex.message);
  } finally {
    btn.disabled = false;
  }
}

$('#saveNotifyBtn').addEventListener('click', function () {
  saveEmailSetting('NOTIFY_EMAIL', '#notifyEmail', '#notifyError',
    '#saveNotifyBtn', 'บันทึกอีเมลแจ้งเตือนแล้ว');
});

$('#saveCcBtn').addEventListener('click', function () {
  saveEmailSetting('CC_EMAIL', '#ccEmail', '#ccError',
    '#saveCcBtn', 'บันทึกอีเมล CC แล้ว');
});

/** สวิตช์ CC — ยิงทันทีที่ติ๊ก ถ้าพลาดให้ดีดกลับสถานะเดิม */
async function saveCcFlag(key, box) {
  $('#ccError').classList.add('hidden');
  try {
    await api('setFlagSetting', { key: key, on: box.checked });
    if (key === 'CC_ON_PENDING') state.settings.ccOnPending = box.checked;
    else state.settings.ccOnShipped = box.checked;
    toast(box.checked ? 'เปิด CC แล้ว' : 'ปิด CC แล้ว');
  } catch (ex) {
    box.checked = !box.checked;
    showError('#ccError', ex.message);
  }
}

$('#ccOnPending').addEventListener('change', function (e) {
  saveCcFlag('CC_ON_PENDING', e.target);
});
$('#ccOnShipped').addEventListener('change', function (e) {
  saveCcFlag('CC_ON_SHIPPED', e.target);
});

$('#checkQuotaBtn').addEventListener('click', async function () {
  const btn = $('#checkQuotaBtn');
  const out = $('#quotaResult');
  btn.disabled = true;
  out.className = 'quota-result';
  out.textContent = 'กำลังตรวจ...';
  try {
    const data = await api('getMailQuota');
    const n = Number(data.quota);
    if (isNaN(n)) {
      out.textContent = String(data.quota);
      out.classList.add('low');
    } else {
      out.textContent = 'วันนี้ส่งได้อีก ' + n + ' ฉบับ';
      out.classList.add(n > 40 ? 'ok' : n > 15 ? 'warn' : 'low');
    }
  } catch (ex) {
    out.className = 'quota-result low';
    out.textContent = ex.message;
  } finally {
    btn.disabled = false;
  }
});

$('#clearHistoryBtn').addEventListener('click', async function () {
  const code = $('#adminCode').value.trim();
  $('#clearError').classList.add('hidden');
  if (!code) return showError('#clearError', 'กรุณากรอกรหัสยืนยัน');

  // ด่านที่สอง — บอกจำนวนงานที่จะหายจริงก่อนให้ยืนยัน
  const warn = 'ยืนยันลบงานทั้งหมด ' + state.jobs.length + ' รายการ และบันทึก Log ทั้งหมด?\n\n' +
    'กู้คืนจากในระบบไม่ได้ ต้องกู้จาก File → Version history ของ Google Sheet เท่านั้น';
  if (!confirm(warn)) return;

  const btn = $('#clearHistoryBtn');
  const label = btn.textContent;
  btn.disabled = true;
  btn.textContent = 'กำลังลบ...';
  try {
    const data = await api('clearHistory', { code: code });
    $('#adminCode').value = '';
    toast('ล้างข้อมูลแล้ว — ลบงาน ' + data.deletedJobs + ' รายการ · รหัสถัดไป ' + data.nextId);
    await loadJobs();
  } catch (ex) {
    showError('#clearError', ex.message);
  } finally {
    btn.disabled = false;
    btn.textContent = label;
  }
});

$('#addAeBtn').addEventListener('click', function () {
  const list = $('#aeList');
  if (!state.settings.aes.length) list.innerHTML = '';
  list.insertAdjacentHTML('beforeend', aeRow(null));
  list.lastElementChild.querySelector('.ae-name').focus();
});

$('#aeList').addEventListener('click', async function (e) {
  const btn = e.target.closest('.ae-save');
  if (!btn) return;

  const row = btn.closest('.ae-row');
  $('#aeError').classList.add('hidden');
  btn.disabled = true;
  const original = btn.textContent;
  btn.textContent = '...';

  try {
    const data = await api('saveAe', {
      originalName: row.dataset.original,
      name: row.querySelector('.ae-name').value.trim(),
      email: row.querySelector('.ae-email').value.trim(),
      active: row.querySelector('input[type="checkbox"]').checked,
    });
    state.settings.aes = data.aes;
    syncActiveAes();
    renderSettings();
    toast('บันทึกรายชื่อ AE แล้ว');
  } catch (ex) {
    showError('#aeError', ex.message);
    btn.disabled = false;
    btn.textContent = original;
  }
});

$('#aeList').addEventListener('change', async function (e) {
  const box = e.target;
  if (box.type !== 'checkbox') return;

  const row = box.closest('.ae-row');
  // แถวที่เพิ่งกด "เพิ่ม AE ใหม่" ยังไม่มีในชีต ค่อยบันทึกพร้อมทั้งแถวตอนกดปุ่ม
  if (!row.dataset.original) return;

  $('#aeError').classList.add('hidden');
  try {
    const data = await api('setAeActive', {
      name: row.dataset.original,
      active: box.checked,
    });
    state.settings.aes = data.aes;
    syncActiveAes();
    toast(box.checked ? 'เปิดใช้งาน AE แล้ว' : 'ปิดใช้งาน AE แล้ว');
  } catch (ex) {
    box.checked = !box.checked;
    showError('#aeError', ex.message);
  }
});

/* ============================================================
 * Modal งาน
 * ========================================================== */

/**
 * เติมตัวเลือก AE จากชีต AEs
 * ถ้า AE ของงานเดิมถูกปิดใช้งานไปแล้ว ยังต้องคงชื่อไว้ให้เห็นว่าเป็นใคร
 */
function fillAeOptions(selected) {
  const list = state.aes.slice();
  if (selected && !list.some(function (a) { return a.name === selected; })) {
    list.push({ name: selected, email: '' });
  }
  $('#f_aeName').innerHTML = ['<option value="">— เลือก AE —</option>'].concat(
    list.map(function (a) {
      return '<option value="' + esc(a.name) + '">' + esc(a.name) + '</option>';
    })
  ).join('');
}

function openModal(job) {
  const form = $('#jobForm');
  const role = state.user.role;
  const editable = EDITABLE[role] || [];
  const isNew = !job;

  state.editingId = isNew ? null : job.id;
  form.reset();
  $('#modalError').classList.add('hidden');
  $('#modalTitle').textContent = isNew ? 'บันทึกงานใหม่' : 'รายละเอียดงาน';

  // ต้องเติม option ก่อนเซ็ตค่า ไม่งั้น select หาค่าที่จะเลือกไม่เจอ
  fillAeOptions(isNew ? '' : job.aeName);

  Array.prototype.forEach.call(form.elements, function (el) {
    if (!el.name) return;
    if (!isNew) el.value = job[el.name] || '';
    el.disabled = editable.indexOf(el.name) === -1;
  });

  // ตอนบันทึกงานใหม่ยังไม่มีข้อมูลจัดส่ง และช่างตัดก็ไม่มีสิทธิ์กรอก จึงซ่อนทั้งบล็อก
  $('#shippingBlock').classList.toggle('hidden', isNew);

  // บล็อกดึงข้อมูลใช้เฉพาะตอนบันทึกงานใหม่ · ล้างทุกครั้งไม่ให้ค่าเก่าค้าง
  $('#lookupBox').classList.toggle('hidden', !isNew || !can('create'));
  $('#requestId').value = '';
  $('#lookupOk').classList.add('hidden');
  $('#lookupError').classList.add('hidden');

  const meta = $('#modalMeta');
  if (isNew) {
    meta.classList.add('hidden');
  } else {
    const lines = [
      '<div>รหัสงาน <b>' + esc(job.id) + '</b> · สถานะ <b>' + esc(job.status) + '</b></div>',
      '<div>บันทึกโดย <b>' + esc(job.createdBy) + '</b> เมื่อ ' + esc(job.createdAt) + '</div>',
    ];
    if (job.aeEmail) {
      lines.push('<div>แจ้ง AE ที่ <b>' + esc(job.aeEmail) + '</b></div>');
    }
    if (job.shipAt) {
      lines.push('<div>จัดส่งโดย <b>' + esc(job.shipBy) + '</b> เมื่อ ' + esc(job.shipAt) + '</div>');
    }
    meta.innerHTML = lines.join('');
    meta.classList.remove('hidden');
  }

  renderModalActions(job);
  $('#modalBackdrop').classList.remove('hidden');
}

function renderModalActions(job) {
  const role = state.user.role;
  const buttons = [];

  if (!job) {
    buttons.push('<button class="btn" data-act="close">ยกเลิก</button>');
    buttons.push('<button class="btn btn-primary" data-act="create">บันทึกงาน</button>');
  } else {
    buttons.push('<button class="btn" data-act="close">ปิด</button>');

    if (job.status === STATUS.PENDING && can('cancel')) {
      buttons.push('<button class="btn btn-danger" data-act="cancel">ยกเลิกงาน</button>');
    }
    if (job.status === STATUS.CANCEL && can('cancel')) {
      buttons.push('<button class="btn" data-act="reopen">เปิดงานใหม่</button>');
    }
    if ((EDITABLE[role] || []).length) {
      buttons.push('<button class="btn" data-act="save">บันทึกการแก้ไข</button>');
    }
    if (job.status === STATUS.PENDING && can('markShip')) {
      buttons.push('<button class="btn btn-green" data-act="ship">ส่งแล้ว</button>');
    }
  }

  $('#modalActions').innerHTML = buttons.join('');
}

function closeModal() {
  $('#modalBackdrop').classList.add('hidden');
  state.editingId = null;
}

function formData() {
  const out = {};
  Array.prototype.forEach.call($('#jobForm').elements, function (el) {
    if (el.name) out[el.name] = el.value.trim();
  });
  return out;
}

function modalError(msg) {
  const el = $('#modalError');
  el.textContent = msg;
  el.classList.remove('hidden');
}

/**
 * ดึงข้อมูลจากระบบ Presales มาเติมฟอร์ม
 * ล้มเหลวแล้วต้องไม่กระทบฟอร์ม — ช่างตัดกรอกเองต่อได้เสมอ
 */
async function lookupRequest() {
  const btn = $('#lookupBtn');
  const id = $('#requestId').value.trim();

  $('#lookupOk').classList.add('hidden');
  $('#lookupError').classList.add('hidden');
  if (!id) return showError('#lookupError', 'กรุณากรอก request_id');

  const label = btn.textContent;
  btn.disabled = true;
  btn.textContent = 'กำลังดึง...';

  try {
    const data = await api('lookupRequest', { requestId: id });

    // อัพเดตรายชื่อ AE ก่อนเติม option ไม่งั้น select เลือกคนที่เพิ่งถูกเพิ่มไม่เจอ
    if (data.aes) state.aes = data.aes;

    const filled = [];

    if (data.customer) {
      $('#f_customer').value = data.customer;
      filled.push('ลูกค้า <b>' + esc(data.customer) + '</b>');
    }

    const code = [data.productCode, data.salesText]
      .filter(function (v) { return String(v || '').trim() !== ''; })
      .join(' | ');
    if (code) {
      $('#f_fileCode').value = code;
      filled.push('รหัสสินค้า <b>' + esc(code) + '</b>');
    }

    if (data.aeName) {
      fillAeOptions(data.aeName);
      $('#f_aeName').value = data.aeName;
      filled.push('AE <b>' + esc(data.aeName) + '</b> · ' + esc(data.aeEmail));
    }

    const ok = $('#lookupOk');
    ok.innerHTML = filled.length
      ? 'เติมให้แล้ว — ' + filled.join('<br>') + '<br>ที่เหลือกรอกเองด้านล่าง'
      : 'พบรายการนี้ แต่ไม่มีข้อมูลที่เติมได้ กรุณากรอกเอง';
    ok.classList.remove('hidden');
  } catch (ex) {
    showError('#lookupError', ex.message);
  } finally {
    btn.disabled = false;
    btn.textContent = label;
  }
}

$('#lookupBtn').addEventListener('click', lookupRequest);

// กด Enter ในช่อง request_id ให้ดึงเลย ไม่ใช่ไปกดปุ่มบันทึกงาน
$('#requestId').addEventListener('keydown', function (e) {
  if (e.key !== 'Enter') return;
  e.preventDefault();
  lookupRequest();
});

$('#modalActions').addEventListener('click', async function (e) {
  const btn = e.target.closest('button[data-act]');
  if (!btn) return;
  const act = btn.dataset.act;

  if (act === 'close') { closeModal(); return; }

  const job = formData();
  $('#modalError').classList.add('hidden');

  if (act === 'create' || act === 'save') {
    if (!job.customer) return modalError('กรุณากรอกชื่อลูกค้า');
    if (!job.destination) return modalError('กรุณากรอกสถานที่จัดส่ง');
    // ช่างตัดต้องเลือก AE เพราะระบบใช้อีเมลของ AE ส่งแจ้งตอนจัดส่งเสร็จ
    if (can('create') && !job.aeName) return modalError('กรุณาเลือก AE');
  }
  if (act === 'ship' && !job.vehicle) {
    return modalError('กรุณาระบุทะเบียนรถก่อนกดส่งแล้ว');
  }
  if (act === 'cancel' && !confirm('ยืนยันยกเลิกงานนี้?')) return;

  const buttons = $$('#modalActions button');
  buttons.forEach(function (b) { b.disabled = true; });
  const originalText = btn.textContent;
  btn.textContent = 'กำลังบันทึก...';

  try {
    if (act === 'create') {
      await api('createJob', { job: job });
      toast('บันทึกงานเรียบร้อย ส่งเมลแจ้งฝ่ายจัดส่งแล้ว');
    } else if (act === 'save') {
      await api('updateJob', { id: state.editingId, job: job });
      toast('บันทึกเรียบร้อย');
    } else if (act === 'ship') {
      await api('setStatus', {
        id: state.editingId, status: STATUS.SHIP,
        vehicle: job.vehicle, contactName: job.contactName, contactPhone: job.contactPhone,
      });
      toast('บันทึกว่าจัดส่งแล้ว ส่งเมลแจ้ง AE แล้ว');
    } else if (act === 'cancel') {
      await api('setStatus', { id: state.editingId, status: STATUS.CANCEL });
      toast('ยกเลิกงานแล้ว');
    } else if (act === 'reopen') {
      await api('setStatus', { id: state.editingId, status: STATUS.PENDING });
      toast('เปลี่ยนกลับเป็นรอส่งแล้ว');
    }
    closeModal();
    await loadJobs();
  } catch (ex) {
    modalError(ex.message);
  } finally {
    buttons.forEach(function (b) { b.disabled = false; });
    btn.textContent = originalText;
  }
});

/* ============================================================
 * Event bindings
 * ========================================================== */

document.addEventListener('click', function (e) {
  const nav = e.target.closest('[data-view]');
  if (nav) { setView(nav.dataset.view); return; }

  const card = e.target.closest('.job-card');
  if (card) {
    const job = state.jobs.find(function (j) { return String(j.id) === card.dataset.id; });
    if (job) openModal(job);
  }
});

$('#newJobBtn').addEventListener('click', function () {
  if (!state.aes.length) {
    toast('ยังไม่มีรายชื่อ AE ในระบบ — เพิ่มได้ที่หน้าตั้งค่า', true);
    return;
  }
  openModal(null);
});
$('#modalClose').addEventListener('click', closeModal);
$('#modalBackdrop').addEventListener('click', function (e) {
  if (e.target === e.currentTarget) closeModal();
});
$('#jobForm').addEventListener('submit', function (e) { e.preventDefault(); });

$('#refreshBtn').addEventListener('click', async function () {
  await loadJobs();
  toast('อัพเดทข้อมูลแล้ว');
});
$('#logoutBtn').addEventListener('click', logout);

$('#searchInput').addEventListener('input', function (e) {
  state.search = e.target.value.trim().toLowerCase();
  renderJobList();
});

$('#statusChips').addEventListener('click', function (e) {
  const chip = e.target.closest('.chip');
  if (!chip) return;
  $$('#statusChips .chip').forEach(function (c) { c.classList.remove('active'); });
  chip.classList.add('active');
  state.filter = chip.dataset.status;
  renderJobList();
});

/* ---------- Dashboard ---------- */
$('#dashSearch').addEventListener('input', function (e) {
  state.dashSearch = e.target.value.trim().toLowerCase();
  renderDashboard();
});

$('#basisChips').addEventListener('click', function (e) {
  const chip = e.target.closest('.chip');
  if (!chip) return;
  $$('#basisChips .chip').forEach(function (c) { c.classList.remove('active'); });
  chip.classList.add('active');
  state.dashBasis = chip.dataset.basis;
  renderDashboard();
});

/* ---------- เปลี่ยนรหัสผ่าน ---------- */
$('#changePwBtn').addEventListener('click', function () {
  $('#pwForm').reset();
  $('#pwError').classList.add('hidden');
  $('#pwBackdrop').classList.remove('hidden');
});
const closePw = function () { $('#pwBackdrop').classList.add('hidden'); };
$('#pwClose').addEventListener('click', closePw);
$('#pwCancel').addEventListener('click', closePw);
$('#pwForm').addEventListener('submit', function (e) { e.preventDefault(); });

$('#pwSave').addEventListener('click', async function () {
  const form = $('#pwForm');
  if (!form.reportValidity()) return;
  const btn = $('#pwSave');
  btn.disabled = true;
  try {
    await api('changePassword', {
      oldPassword: form.elements.oldPassword.value,
      newPassword: form.elements.newPassword.value,
    });
    closePw();
    toast('เปลี่ยนรหัสผ่านเรียบร้อย');
  } catch (ex) {
    const el = $('#pwError');
    el.textContent = ex.message;
    el.classList.remove('hidden');
  } finally {
    btn.disabled = false;
  }
});

/* ---------- รีเฟรชอัตโนมัติเมื่อกลับมาที่หน้าจอ ---------- */
document.addEventListener('visibilitychange', function () {
  if (!document.hidden && state.token) loadJobs();
});

/* ============================================================
 * เริ่มทำงาน
 * ========================================================== */

(async function init() {
  if (!state.token || !state.user) return;
  try {
    const data = await api('me');
    state.user = data.user;
    localStorage.setItem('user', JSON.stringify(data.user));
    await enterApp();
  } catch (ex) {
    logout();
  }
})();

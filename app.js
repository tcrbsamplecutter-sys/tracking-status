/* ============================================================
 * ระบบติดตามงานตัดกล่องตัวอย่าง — Frontend
 * ========================================================== */

const STATUS = {
  WAIT: 'รอตัด',
  CUT: 'ตัดเสร็จ',
  SHIP: 'ส่งแล้ว',
  CANCEL: 'ยกเลิก',
};

const STATUS_CLASS = {
  'รอตัด': 's-wait',
  'ตัดเสร็จ': 's-cut',
  'ส่งแล้ว': 's-ship',
  'ยกเลิก': 's-cancel',
};

const STATUS_LABEL = {
  'รอตัด': 'รอตัด',
  'ตัดเสร็จ': 'รอส่ง',
  'ส่งแล้ว': 'ส่งแล้ว',
  'ยกเลิก': 'ยกเลิก',
};

const ROLE_LABEL = { cutter: 'ช่างตัด', shipping: 'จัดส่ง', viewer: 'ดูอย่างเดียว' };

const VIEW_TITLE = { jobs: 'รายการงาน', summary: 'สรุปวันนี้' };

/**
 * ฟิลด์ที่แต่ละบทบาทแก้ไขได้
 * ช่างตัด = ข้อมูลงานและปลายทาง / จัดส่ง = ข้อมูลการขนส่ง
 * ผู้ติดต่อ เบอร์โทร ทะเบียนรถ เป็นของฝ่ายจัดส่งเท่านั้น
 * (Apps Script ตรวจซ้ำอีกชั้น การซ่อนตรงนี้เป็นแค่เรื่องความสะดวก)
 */
const EDITABLE = {
  cutter: ['customer', 'sales', 'jobName', 'qty', 'destination', 'dueDate', 'note'],
  shipping: ['destination', 'contactName', 'contactPhone', 'vehicle', 'note'],
  viewer: [],
};

const state = {
  token: localStorage.getItem('token') || '',
  user: JSON.parse(localStorage.getItem('user') || 'null'),
  jobs: [],
  today: '',
  filter: 'all',
  search: '',
  view: 'jobs',
  editingId: null,
};

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
  if (action === 'markCut') return role === 'cutter';
  if (action === 'markShip') return role === 'shipping';
  if (action === 'cancel') return role === 'cutter';
  return false;
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
  setView(state.view);
  await loadJobs();
}

function setView(view) {
  state.view = view;
  $$('[data-view]').forEach(function (b) {
    b.classList.toggle('active', b.dataset.view === view);
  });
  $('#jobsView').classList.toggle('hidden', view !== 'jobs');
  $('#summaryView').classList.toggle('hidden', view !== 'summary');
  $('#pageTitle').textContent = VIEW_TITLE[view] || '';
  $('#newJobBtn').classList.toggle('hidden', !can('create') || view !== 'jobs');
}

/* ============================================================
 * โหลดและแสดงผล
 * ========================================================== */

async function loadJobs() {
  try {
    const data = await api('listJobs');
    state.jobs = data.jobs;
    state.today = data.today;
    render();
  } catch (ex) {
    toast(ex.message, true);
  }
}

function render() {
  renderStats();
  renderJobList();
  renderSummary();
}

function renderStats() {
  const jobs = state.jobs;
  const count = function (fn) { return jobs.filter(fn).length; };

  const waiting = count(function (j) { return j.status === STATUS.WAIT; });
  const pendingShip = count(function (j) { return j.status === STATUS.CUT; });
  const shippedToday = count(function (j) {
    return j.status === STATUS.SHIP && String(j.shipAt).slice(0, 10) === state.today;
  });
  const late = count(function (j) {
    return j.dueDate && String(j.dueDate).slice(0, 10) < state.today
      && j.status !== STATUS.SHIP && j.status !== STATUS.CANCEL;
  });

  $('#stats').innerHTML = [
    { cls: 'amber', num: waiting, label: 'รอตัด' },
    { cls: 'blue', num: pendingShip, label: 'รอจัดส่ง' },
    { cls: 'green', num: shippedToday, label: 'ส่งแล้ววันนี้' },
    { cls: 'red', num: late, label: 'เลยกำหนด' },
  ].map(function (s) {
    return '<div class="stat ' + s.cls + '">' +
      '<div class="stat-num">' + s.num + '</div>' +
      '<div class="stat-label">' + s.label + '</div></div>';
  }).join('');
}

function matchesFilter(job) {
  if (state.filter !== 'all' && job.status !== state.filter) return false;
  if (!state.search) return true;
  const hay = [job.id, job.customer, job.sales, job.jobName, job.destination,
               job.contactName, job.vehicle].join(' ').toLowerCase();
  return hay.indexOf(state.search) !== -1;
}

function jobCard(job, compact) {
  const cls = STATUS_CLASS[job.status] || 's-cancel';
  const bits = [];

  if (job.jobName) bits.push('<span>งาน: <b>' + esc(job.jobName) + '</b></span>');
  if (job.qty) bits.push('<span>จำนวน: <b>' + esc(job.qty) + '</b></span>');
  if (job.sales) bits.push('<span>Sales: <b>' + esc(job.sales) + '</b></span>');
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
    '<span class="badge ' + cls + '">' + (STATUS_LABEL[job.status] || job.status) + '</span>' +
    '</div>' +
    '<div class="job-meta">' + bits.join('') + '</div></div>';
}

function renderJobList() {
  const list = state.jobs.filter(matchesFilter).slice().reverse();
  $('#jobList').innerHTML = list.map(function (j) { return jobCard(j, false); }).join('');
  $('#emptyState').classList.toggle('hidden', list.length > 0);
}

function renderSummary() {
  $('#todayLabel').textContent = thaiDate(state.today);

  const fill = function (sel, jobs, emptyText) {
    $(sel).innerHTML = jobs.length
      ? jobs.map(function (j) { return jobCard(j, true); }).join('')
      : '<div class="empty">' + emptyText + '</div>';
  };

  fill('#pendingShipList',
    state.jobs.filter(function (j) { return j.status === STATUS.CUT; }),
    'ไม่มีงานค้างรอส่ง');

  fill('#cutTodayList',
    state.jobs.filter(function (j) { return String(j.cutAt).slice(0, 10) === state.today; }),
    'วันนี้ยังไม่มีงานที่ตัดเสร็จ');

  fill('#shipTodayList',
    state.jobs.filter(function (j) { return String(j.shipAt).slice(0, 10) === state.today; }),
    'วันนี้ยังไม่มีงานที่จัดส่ง');
}

/* ============================================================
 * Modal งาน
 * ========================================================== */

function openModal(job) {
  const form = $('#jobForm');
  const role = state.user.role;
  const editable = EDITABLE[role] || [];
  const isNew = !job;

  state.editingId = isNew ? null : job.id;
  form.reset();
  $('#modalError').classList.add('hidden');
  $('#modalTitle').textContent = isNew ? 'สร้างงานใหม่' : 'รายละเอียดงาน';

  Array.prototype.forEach.call(form.elements, function (el) {
    if (!el.name) return;
    if (!isNew) el.value = job[el.name] || '';
    el.disabled = editable.indexOf(el.name) === -1;
  });

  // ตอนสร้างงานใหม่ยังไม่มีข้อมูลจัดส่ง และช่างตัดก็ไม่มีสิทธิ์กรอก จึงซ่อนทั้งบล็อก
  $('#shippingBlock').classList.toggle('hidden', isNew);

  const meta = $('#modalMeta');
  if (isNew) {
    meta.classList.add('hidden');
  } else {
    const lines = [
      '<div>รหัสงาน <b>' + esc(job.id) + '</b> · สถานะ <b>' +
        (STATUS_LABEL[job.status] || job.status) + '</b></div>',
      '<div>สร้างโดย <b>' + esc(job.createdBy) + '</b> เมื่อ ' + esc(job.createdAt) + '</div>',
    ];
    if (job.cutAt) lines.push('<div>ตัดเสร็จโดย <b>' + esc(job.cutBy) + '</b> เมื่อ ' + esc(job.cutAt) + '</div>');
    if (job.shipAt) lines.push('<div>จัดส่งโดย <b>' + esc(job.shipBy) + '</b> เมื่อ ' + esc(job.shipAt) + '</div>');
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
    buttons.push('<button class="btn btn-primary" data-act="create">สร้างงาน</button>');
  } else {
    buttons.push('<button class="btn" data-act="close">ปิด</button>');

    if (job.status === STATUS.WAIT && can('cancel')) {
      buttons.push('<button class="btn btn-danger" data-act="cancel">ยกเลิกงาน</button>');
    }
    if (job.status === STATUS.CANCEL && can('markCut')) {
      buttons.push('<button class="btn" data-act="reopen">เปิดงานใหม่</button>');
    }
    if (job.status === STATUS.CUT && can('markCut')) {
      buttons.push('<button class="btn" data-act="reopen">กลับไปรอตัด</button>');
    }
    if ((EDITABLE[role] || []).length) {
      buttons.push('<button class="btn" data-act="save">บันทึกการแก้ไข</button>');
    }
    if (job.status === STATUS.WAIT && can('markCut')) {
      buttons.push('<button class="btn btn-green" data-act="cut">ตัดเสร็จแล้ว</button>');
    }
    if (job.status === STATUS.CUT && can('markShip')) {
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
      toast('สร้างงานเรียบร้อย');
    } else if (act === 'save') {
      await api('updateJob', { id: state.editingId, job: job });
      toast('บันทึกเรียบร้อย');
    } else if (act === 'cut') {
      await api('setStatus', { id: state.editingId, status: STATUS.CUT });
      toast('บันทึกว่าตัดเสร็จแล้ว');
    } else if (act === 'ship') {
      await api('setStatus', {
        id: state.editingId, status: STATUS.SHIP,
        vehicle: job.vehicle, contactName: job.contactName, contactPhone: job.contactPhone,
      });
      toast('บันทึกว่าจัดส่งแล้ว');
    } else if (act === 'cancel') {
      await api('setStatus', { id: state.editingId, status: STATUS.CANCEL });
      toast('ยกเลิกงานแล้ว');
    } else if (act === 'reopen') {
      await api('setStatus', { id: state.editingId, status: STATUS.WAIT });
      toast('เปลี่ยนกลับเป็นรอตัดแล้ว');
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

$('#newJobBtn').addEventListener('click', function () { openModal(null); });
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

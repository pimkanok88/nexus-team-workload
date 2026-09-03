/**
 * Team Workload + OT Web App v5.4 Performance
 * Weekly Consult workload + Subtasks + member capacity/status + workload-adjusted target.
 * NEW SYSTEM - standalone Apps Script.
 * Does not read or depend on any old spreadsheet.
 *
 * First run:
 *   1) Run setupNewSystem()
 *   2) Authorize
 *   3) Deploy > New deployment > Web app
 */

const APP = {
  DB_PROP: 'DB_SPREADSHEET_ID',
  BACKUP_FOLDER_PROP: 'BACKUP_FOLDER_ID',
  NETLIFY_API_SECRET_PROP: 'NETLIFY_API_SECRET',
  TZ: 'Asia/Bangkok',
  SHEETS: {
    MEMBERS: 'Members',
    SETTINGS: 'Settings',
    TASKS: 'Tasks',
    TASK_MEMBERS: 'TaskMembers',
    TASK_UPDATES: 'TaskUpdates',
    SUBTASKS: 'Subtasks',
    CONSULT: 'ConsultLog',
    OT_SCHEDULE: 'OT_Schedule',
    OT_ENTRIES: 'OT_Entries',
    AUDIT: 'AuditLog',
    BACKUP: 'BackupLog'
  }
};

const HEADERS = {
  Members: ['member_id', 'name', 'active', 'sort_order', 'work_status', 'status_note', 'status_updated_at'],
  Settings: ['key', 'value', 'description'],
  Tasks: [
    'task_id', 'category', 'job', 'description', 'holder_name',
    'start_date', 'due_date', 'deadline', 'weight', 'status',
    'created_at', 'created_by', 'updated_at'
  ],
  TaskMembers: ['task_id', 'member_name', 'share_weight'],
  TaskUpdates: [
    'update_id', 'task_id', 'update_date', 'member_name',
    'progress', 'note', 'created_at', 'created_by'
  ],
  Subtasks: [
    'subtask_id', 'task_id', 'title', 'description', 'due_date',
    'status', 'sort_order', 'created_at', 'created_by', 'updated_at'
  ],
  ConsultLog: [
    'consult_id', 'consult_date', 'member_name', 'consult_count',
    'source_task_id', 'created_at', 'created_by', 'updated_at'
  ],
  OT_Schedule: [
    'schedule_id', 'work_date', 'start_time', 'end_time',
    'break_hours', 'planned_hours', 'is_weekend', 'active',
    'note', 'created_at', 'updated_at'
  ],
  OT_Entries: [
    'entry_id', 'work_date', 'member_name', 'start_time', 'end_time',
    'break_hours', 'hours', 'roles', 'rate_per_hour', 'amount',
    'note', 'created_at', 'created_by', 'updated_at'
  ],
  AuditLog: [
    'log_id', 'timestamp', 'user', 'action', 'module',
    'record_id', 'before_json', 'after_json'
  ],
  BackupLog: ['backup_id', 'timestamp', 'google_copy_name', 'xlsx_name', 'status', 'message']
};

/* =========================
   SETUP
========================= */

function setupNewSystem() {
  const props = PropertiesService.getScriptProperties();
  const oldId = props.getProperty(APP.DB_PROP);

  if (oldId) {
    throw new Error(
      'Script นี้เคย setup แล้ว และมี DB_SPREADSHEET_ID อยู่แล้ว\n' +
      'หากต้องการสร้างฐานใหม่จริง ๆ ให้รัน resetSystemPointer() ก่อน แล้วค่อยรัน setupNewSystem()'
    );
  }

  const ss = SpreadsheetApp.create('Team Workload & OT - NEW DATABASE');
  ss.setSpreadsheetTimeZone(APP.TZ);

  const defaultSheet = ss.getSheets()[0];
  defaultSheet.setName(APP.SHEETS.MEMBERS);

  const sheetNames = Object.values(APP.SHEETS);
  sheetNames.forEach((name) => {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });

  Object.entries(HEADERS).forEach(([sheetName, headers]) => {
    const sh = ss.getSheetByName(sheetName);
    sh.clear();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#E8EEF9');
    sh.autoResizeColumns(1, headers.length);
  });

  const members = ['Bus', 'Tem', 'Phai', 'Kanok', 'Ninii', 'Por', 'Kob'];
  const memberRows = members.map((name, i) => [
    `M${String(i + 1).padStart(2, '0')}`, name, true, i + 1, 1, 'ปกติ', nowString_()
  ]);
  ss.getSheetByName(APP.SHEETS.MEMBERS)
    .getRange(2, 1, memberRows.length, memberRows[0].length)
    .setValues(memberRows);

  const settings = [
    ['CONSULT_REFERENCE', 36, 'สัดส่วน consult สำหรับใช้อ้างอิง'],
    ['OT_RATE', 100, 'อัตรา OT บาทต่อชั่วโมง'],
    ['WEEKDAY_START', '16:00', 'เวลาเริ่มต้นปกติวันธรรมดา'],
    ['WEEKDAY_END', '18:00', 'เวลาสิ้นสุดปกติวันธรรมดา'],
    ['WEEKDAY_BREAK', 0, 'เวลาพักวันธรรมดา'],
    ['WEEKEND_START', '08:00', 'เวลาเริ่มต้นปกติวันเสาร์-อาทิตย์'],
    ['WEEKEND_END', '16:00', 'เวลาสิ้นสุดปกติวันเสาร์-อาทิตย์'],
    ['WEEKEND_BREAK', 1, 'เวลาพักวันเสาร์-อาทิตย์'],
    ['APP_NAME', 'Team Workload & OT', 'ชื่อระบบ']
  ];
  ss.getSheetByName(APP.SHEETS.SETTINGS)
    .getRange(2, 1, settings.length, settings[0].length)
    .setValues(settings);

  props.setProperty(APP.DB_PROP, ss.getId());

  const folder = DriveApp.createFolder('TeamWorkloadOT_Backup');
  props.setProperty(APP.BACKUP_FOLDER_PROP, folder.getId());

  ensureWeeklyBackupTrigger_();

  appendAudit_('SYSTEM_SETUP', 'SYSTEM', ss.getId(), null, {
    spreadsheetUrl: ss.getUrl(),
    backupFolderId: folder.getId()
  });

  return {
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    backupFolderId: folder.getId(),
    backupFolderUrl: folder.getUrl()
  };
}

function resetSystemPointer() {
  // IMPORTANT: This does not delete any existing files.
  // It only clears script properties so setupNewSystem() can create a NEW database.
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(APP.DB_PROP);
  props.deleteProperty(APP.BACKUP_FOLDER_PROP);

  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === 'weeklyBackup') ScriptApp.deleteTrigger(t);
  });

  return { ok: true };
}

function ensureWeeklyBackupTrigger_() {
  const exists = ScriptApp.getProjectTriggers()
    .some((t) => t.getHandlerFunction() === 'weeklyBackup');

  if (!exists) {
    ScriptApp.newTrigger('weeklyBackup')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.FRIDAY)
      .atHour(23)
      .create();
  }
}

/* =========================
   WEB APP
========================= */

let REQUEST_ACTOR_ = '';

function doGet() {
  return jsonOutput_({
    ok: true,
    service: 'Team Workload API',
    message: 'API พร้อมใช้งาน — ให้เปิดหน้าเว็บจาก Netlify',
    timestamp: nowString_()
  });
}

/**
 * HTTP API endpoint used by the Netlify proxy.
 * The browser never receives NETLIFY_API_SECRET; Netlify adds it server-side.
 */
function doPost(e) {
  REQUEST_ACTOR_ = '';
  try {
    const raw = e && e.postData ? String(e.postData.contents || '') : '';
    if (!raw) throw new Error('Request body ว่าง');

    const body = JSON.parse(raw);
    const expectedSecret = PropertiesService.getScriptProperties()
      .getProperty(APP.NETLIFY_API_SECRET_PROP);

    if (!expectedSecret) {
      throw new Error(
        'ยังไม่ได้ตั้ง NETLIFY_API_SECRET ใน Apps Script กรุณารัน generateNetlifyApiSecret() ก่อน'
      );
    }

    if (String(body.secret || '') !== String(expectedSecret)) {
      return jsonOutput_({ ok: false, error: 'Unauthorized API request' });
    }

    REQUEST_ACTOR_ = sanitizeActor_(body.actor);
    const result = apiCall(body.fn, body.args);

    return jsonOutput_({
      ok: true,
      data: result
    });
  } catch (err) {
    console.error(err);
    return jsonOutput_({
      ok: false,
      error: err && err.message ? err.message : String(err)
    });
  } finally {
    REQUEST_ACTOR_ = '';
  }
}

/**
 * Run ONCE and copy the returned secret to Netlify:
 * APPS_SCRIPT_API_SECRET = <returned value>
 */
function generateNetlifyApiSecret() {
  const secret = [
    Utilities.getUuid().replace(/-/g, ''),
    Utilities.getUuid().replace(/-/g, '')
  ].join('');
  PropertiesService.getScriptProperties()
    .setProperty(APP.NETLIFY_API_SECRET_PROP, secret);
  console.log('APPS_SCRIPT_API_SECRET=' + secret);
  return secret;
}

/**
 * Optional: set your own secret manually.
 */
function setNetlifyApiSecret(secret) {
  secret = String(secret || '').trim();
  if (secret.length < 24) {
    throw new Error('Secret ควรยาวอย่างน้อย 24 ตัวอักษร');
  }
  PropertiesService.getScriptProperties()
    .setProperty(APP.NETLIFY_API_SECRET_PROP, secret);
  return { ok: true };
}

/**
 * Only needed if this Code.gs is placed in a NEW Apps Script project.
 * If you are replacing Code.gs in the existing project, DB_SPREADSHEET_ID
 * is already present and you do NOT need to run this.
 */
function connectExistingSpreadsheet(spreadsheetId) {
  spreadsheetId = String(spreadsheetId || '').trim();
  if (!spreadsheetId) throw new Error('กรุณาระบุ Spreadsheet ID');

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const required = [
    APP.SHEETS.MEMBERS,
    APP.SHEETS.SETTINGS,
    APP.SHEETS.TASKS,
    APP.SHEETS.TASK_MEMBERS,
    APP.SHEETS.CONSULT
  ];
  const missing = required.filter(name => !ss.getSheetByName(name));
  if (missing.length) {
    throw new Error('Spreadsheet นี้ขาด Sheet: ' + missing.join(', '));
  }

  PropertiesService.getScriptProperties().setProperty(APP.DB_PROP, spreadsheetId);
  DB_HANDLE_CACHE_ = null;

  return {
    ok: true,
    spreadsheetId,
    name: ss.getName(),
    url: ss.getUrl()
  };
}

function getNetlifyApiStatus() {
  const props = PropertiesService.getScriptProperties();
  const dbId = props.getProperty(APP.DB_PROP) || '';
  return {
    databaseConnected: !!dbId,
    spreadsheetId: dbId,
    apiSecretConfigured: !!props.getProperty(APP.NETLIFY_API_SECRET_PROP),
    webAppUrl: ScriptApp.getService().getUrl() || ''
  };
}

/**
 * RPC gateway shared by the Netlify HTTP API and the old Apps Script UI.
 */
function apiCall(fn, args) {
  args = Array.isArray(args) ? args : [];

  switch (String(fn || '')) {
    case 'getBootstrapData':
      return getBootstrapData();
    case 'getTasks':
      return getTasks();
    case 'getAuditLogs':
      return getAuditLogs(args[0]);
    case 'saveMemberStatus':
      return saveMemberStatus(args[0]);
    case 'createTask':
      return createTask(args[0]);
    case 'updateTaskStatus':
      return updateTaskStatus(args[0], args[1]);
    case 'getTaskDetail':
      return getTaskDetail(args[0]);
    case 'updateTask':
      return updateTask(args[0]);
    case 'getConsultData':
      return getConsultData(args[0]);
    case 'getConsultWeek':
      return getConsultWeek(args[0]);
    case 'saveConsultWeek':
      return saveConsultWeek(args[0]);
    case 'addSubtask':
      return addSubtask(args[0]);
    case 'updateSubtask':
      return updateSubtask(args[0]);
    case 'deleteSubtask':
      return deleteSubtask(args[0]);
    case 'generateOtMonth':
      return generateOtMonth(args[0], args[1]);
    case 'getOtMonth':
      return getOtMonth(args[0], args[1]);
    case 'saveOtDay':
      return saveOtDay(args[0]);
    default:
      throw new Error('ไม่พบ API function: ' + fn);
  }
}

function getBootstrapData() {
  // PERFORMANCE: normal page load is read-only and only reads the five sheets
  // required by Dashboard / Members / Consult. Heavy pages (Tasks detail,
  // OT and Audit) are loaded lazily when the user opens them.
  const core = readCoreData_();
  const now = new Date();
  const consultRange = currentWorkWeekRange_();
  const members = membersFromRows_(core.members);
  const settings = settingsFromRows_(core.settings);

  return {
    appName: settings.APP_NAME || 'Team Workload & OT',
    members,
    settings,
    currentYear: Number(Utilities.formatDate(now, APP.TZ, 'yyyy')),
    currentMonth: Number(Utilities.formatDate(now, APP.TZ, 'M')),
    dashboard: buildDashboardData_(members, settings, core.tasks, core.taskMembers, core.consult, consultRange),
    consult: buildConsultData_(consultRange, members, core.consult),
    // These pages are intentionally lazy-loaded by Index.html.
    tasks: [],
    audit: [],
    lazy: { tasks: true, ot: true, audit: true }
  };
}

function readCoreData_() {
  return {
    members: getSheetObjects_(APP.SHEETS.MEMBERS),
    settings: getSheetObjects_(APP.SHEETS.SETTINGS),
    tasks: getSheetObjects_(APP.SHEETS.TASKS),
    taskMembers: getSheetObjects_(APP.SHEETS.TASK_MEMBERS),
    consult: getSheetObjects_(APP.SHEETS.CONSULT)
  };
}

function ensureV5Schema_() {
  const ss = getDb_();
  Object.entries(HEADERS).forEach(([sheetName, headers]) => {
    let sh = ss.getSheetByName(sheetName);
    if (!sh) sh = ss.insertSheet(sheetName);

    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else {
      const currentHeaders = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String);
      headers.forEach((field) => {
        if (!currentHeaders.includes(field)) {
          const col = sh.getLastColumn() + 1;
          sh.getRange(1, col).setValue(field);
          currentHeaders.push(field);
        }
      });
    }

    sh.setFrozenRows(1);
    if (sh.getLastColumn() > 0) {
      sh.getRange(1, 1, 1, sh.getLastColumn())
        .setFontWeight('bold')
        .setBackground('#E8EEF9');
    }
  });

  ensureMemberStatusSchema_();
  // IMPORTANT: ConsultLog is manual-entry data only.
  // Do NOT auto-migrate legacy Consult tasks during page load/refresh.
  // Refresh must never create or modify Consult values.
}

function migrateLegacyConsultTasks_() {
  const ss = getDb_();
  const taskSh = ss.getSheetByName(APP.SHEETS.TASKS);
  const memberSh = ss.getSheetByName(APP.SHEETS.TASK_MEMBERS);
  const consultSh = ss.getSheetByName(APP.SHEETS.CONSULT);
  if (!taskSh || !memberSh || !consultSh || taskSh.getLastRow() < 2) return 0;

  const tasks = getSheetObjects_(APP.SHEETS.TASKS)
    .filter((t) => String(t.category || '').toLowerCase() === 'consult');
  if (!tasks.length) return 0;

  const taskMembers = getSheetObjects_(APP.SHEETS.TASK_MEMBERS);
  const existing = new Set(
    getSheetObjects_(APP.SHEETS.CONSULT)
      .filter((x) => x.source_task_id)
      .map((x) => `${String(x.source_task_id)}|${String(x.member_name)}`)
  );

  const byTask = {};
  taskMembers.forEach((tm) => {
    if (!byTask[tm.task_id]) byTask[tm.task_id] = [];
    byTask[tm.task_id].push(tm);
  });

  let created = 0;
  const now = nowString_();
  tasks.forEach((task) => {
    const taskId = String(task.task_id);
    const rawDate = dateOnly_(task.start_date);
    const consultDate = /^\d{4}-\d{2}-\d{2}$/.test(rawDate) ? rawDate : todayString_();
    (byTask[taskId] || []).forEach((tm) => {
      const member = String(tm.member_name || '');
      const key = `${taskId}|${member}`;
      if (!member || existing.has(key)) return;
      const count = num_(tm.share_weight);
      if (count <= 0) return;

      consultSh.appendRow([
        makeId_('CON'), consultDate, member, count, taskId,
        now, 'legacy-migration', now
      ]);
      existing.add(key);
      created++;
    });
  });

  if (created) {
    appendAudit_('MIGRATE_LEGACY', 'CONSULT', 'legacy-tasks', null, {created});
  }
  return created;
}

/**
 * ONE-TIME CLEANUP for v5.3 and earlier.
 * Removes Consult rows that were automatically created from legacy Tasks.
 * It does NOT delete manually entered Consult rows (source_task_id blank).
 *
 * Run this function manually ONCE from Apps Script if you already saw
 * unexpected Consult values after Refresh.
 */
function cleanupLegacyConsultAutoRows() {
  ensureV5Schema_();
  const sh = getDb_().getSheetByName(APP.SHEETS.CONSULT);
  if (!sh || sh.getLastRow() < 2) {
    return {ok: true, deleted: 0};
  }

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(String);
  const idx = headerIndex_(headers);
  const removed = [];

  for (let i = values.length - 1; i >= 1; i--) {
    const sourceTaskId = String(values[i][idx.source_task_id] || '').trim();
    const createdBy = String(values[i][idx.created_by] || '').trim();
    const isLegacyAutoRow = !!sourceTaskId || createdBy === 'legacy-migration';
    if (!isLegacyAutoRow) continue;

    removed.push(rowToObject_(headers, values[i]));
    sh.deleteRow(i + 1);
  }

  appendAudit_(
    'CLEANUP_LEGACY_AUTO',
    'CONSULT',
    'legacy-auto-rows',
    removed,
    {deleted: removed.length}
  );

  return {
    ok: true,
    deleted: removed.length,
    message: `ลบ Consult ที่ระบบ migration สร้างอัตโนมัติแล้ว ${removed.length} รายการ`
  };
}

/* =========================
   MEMBER CAPACITY / STATUS
========================= */

function ensureMemberStatusSchema_() {
  const sh = getDb_().getSheetByName(APP.SHEETS.MEMBERS);
  if (!sh) throw new Error('ไม่พบ sheet Members');

  const lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const required = ['work_status', 'status_note', 'status_updated_at'];

  required.forEach((field) => {
    if (!headers.includes(field)) {
      const col = sh.getLastColumn() + 1;
      sh.getRange(1, col).setValue(field).setFontWeight('bold').setBackground('#E8EEF9');
      headers.push(field);
    }
  });

  const idx = headerIndex_(headers);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  const now = nowString_();
  const statusCol = idx.work_status + 1;
  const noteCol = idx.status_note + 1;
  const updatedCol = idx.status_updated_at + 1;

  values.forEach((row, i) => {
    const rowNum = i + 2;
    if (row[idx.work_status] === '' || row[idx.work_status] === null || row[idx.work_status] === undefined) {
      sh.getRange(rowNum, statusCol).setValue(1);
    }
    if (row[idx.status_note] === '' || row[idx.status_note] === null || row[idx.status_note] === undefined) {
      sh.getRange(rowNum, noteCol).setValue('ปกติ');
    }
    if (row[idx.status_updated_at] === '' || row[idx.status_updated_at] === null || row[idx.status_updated_at] === undefined) {
      sh.getRange(rowNum, updatedCol).setValue(now);
    }
  });

  const rule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 1)
    .setAllowInvalid(false)
    .setHelpText('กรอกค่าระหว่าง 0 ถึง 1 เช่น 1, 0.5, 0')
    .build();
  sh.getRange(2, statusCol, Math.max(1, lastRow - 1), 1).setDataValidation(rule);
}

function saveMemberStatus(payload) {
  if (!payload || !payload.memberId) throw new Error('กรุณาระบุสมาชิก');
  const workStatus = Number(payload.workStatus);
  if (!Number.isFinite(workStatus) || workStatus < 0 || workStatus > 1) {
    throw new Error('Status ต้องเป็นตัวเลขระหว่าง 0 ถึง 1');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = getDb_().getSheetByName(APP.SHEETS.MEMBERS);
    const values = sh.getDataRange().getValues();
    const headers = values[0].map(String);
    const idx = headerIndex_(headers);

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idx.member_id]) !== String(payload.memberId)) continue;

      const before = rowToObject_(headers, values[i]);
      const rowNum = i + 1;
      const note = String(payload.statusNote || '').trim() || defaultMemberStatusNote_(workStatus);
      const now = nowString_();

      values[i][idx.work_status] = workStatus;
      values[i][idx.status_note] = note;
      values[i][idx.status_updated_at] = now;
      sh.getRange(rowNum, 1, 1, headers.length).setValues([values[i].slice(0, headers.length)]);

      appendAudit_('UPDATE_STATUS', 'MEMBER', payload.memberId, before, {
        member_id: payload.memberId,
        work_status: workStatus,
        status_note: note,
        status_updated_at: now
      });

      return {
        ok: true,
        members: getMembers_(),
        dashboard: getDashboardData()
      };
    }
    throw new Error('ไม่พบสมาชิก: ' + payload.memberId);
  } finally {
    lock.releaseLock();
  }
}

function defaultMemberStatusNote_(workStatus) {
  if (workStatus === 0) return 'ไม่รับงาน / ลา';
  if (workStatus < 1) return 'รับงานบางส่วน';
  return 'ปกติ';
}

/* =========================
   PAGE 1: WORKLOAD
========================= */

function getDashboardData() {
  const members = getMembers_();
  const settings = getSettings_();
  const tasks = getSheetObjects_(APP.SHEETS.TASKS);
  const taskMembers = getSheetObjects_(APP.SHEETS.TASK_MEMBERS);
  const consultRows = getSheetObjects_(APP.SHEETS.CONSULT);
  return buildDashboardData_(members, settings, tasks, taskMembers, consultRows, currentWorkWeekRange_());
}

function buildDashboardData_(members, settings, tasks, taskMembers, consultRows, consultRange) {
  const consultEntries = (consultRows || [])
    .filter((x) => dateInRange_(dateOnly_(x.consult_date), consultRange.startDate, consultRange.endDate));

  // Consult is stored separately. Legacy Consult tasks are deliberately excluded.
  const activeTasks = (tasks || []).filter((t) =>
    String(t.category || '').toLowerCase() !== 'consult' &&
    !['Complete', 'Cancelled'].includes(String(t.status))
  );

  const tmByTask = {};
  (taskMembers || []).forEach((tm) => {
    if (!tmByTask[tm.task_id]) tmByTask[tm.task_id] = [];
    tmByTask[tm.task_id].push(tm);
  });

  const byMember = {};
  (members || []).forEach((m) => {
    byMember[m.name] = {
      member: m.name,
      workStatus: m.workStatus,
      statusNote: m.statusNote,
      consult: 0,
      other: 0,
      workload: 0,
      activeTasks: 0,
      targetWorkload: 0,
      remainingConsult: 0
    };
  });

  let consultActual = 0;
  consultEntries.forEach((entry) => {
    const name = String(entry.member_name || '');
    const count = num_(entry.consult_count); // 1 Consult = workload weight 1
    consultActual += count;
    if (byMember[name]) byMember[name].consult += count;
  });

  let otherActual = 0;
  activeTasks.forEach((task) => {
    const assigned = tmByTask[task.task_id] || [];
    assigned.forEach((tm) => {
      const name = String(tm.member_name);
      if (!byMember[name]) return;
      const share = num_(tm.share_weight);
      byMember[name].other += share;
      byMember[name].workload += share;
      byMember[name].activeTasks += 1;
      otherActual += share;
    });
  });

  Object.values(byMember).forEach((x) => { x.workload += x.consult; });

  const total = consultActual + otherActual;
  const totalCapacity = (members || []).reduce((sum, m) => sum + num_(m.workStatus), 0);
  const average = totalCapacity > 0 ? total / totalCapacity : 0;

  Object.values(byMember).forEach((x) => {
    x.consult = round2_(x.consult);
    x.other = round2_(x.other);
    x.workload = round2_(x.workload);
    x.targetWorkload = round2_(average * num_(x.workStatus));
    x.remainingConsult = round2_(x.targetWorkload - x.workload);
  });

  return {
    consultReference: num_(settings.CONSULT_REFERENCE),
    consultActual: round2_(consultActual),
    consultPeriod: consultRange,
    otherActual: round2_(otherActual),
    total: round2_(total),
    totalCapacity: round2_(totalCapacity),
    average: round2_(average),
    members: Object.values(byMember).sort((a, b) => b.remainingConsult - a.remainingConsult)
  };
}

function createTask(payload) {
  validateTask_(payload);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const now = nowString_();
    const user = getUser_();
    const taskId = makeId_('TSK');
    const members = [...new Set((payload.members || []).map(String))];
    const weight = num_(payload.weight);
    const share = members.length ? weight / members.length : 0;

    const taskRow = [
      taskId,
      String(payload.category || 'Other').toLowerCase() === 'consult' ? 'Other' : (payload.category || 'Other'),
      payload.job || '',
      payload.description || '',
      payload.holder_name || '',
      payload.start_date || '',
      payload.due_date || '',
      payload.deadline || '',
      weight,
      payload.status || 'Inprocess',
      now,
      user,
      now
    ];

    appendRow_(APP.SHEETS.TASKS, taskRow);

    appendRows_(APP.SHEETS.TASK_MEMBERS, members.map((member) => [
      taskId, member, round4_(share)
    ]));

    appendAudit_('CREATE', 'TASK', taskId, null, payload);

    return {
      ok: true,
      taskId,
      dashboard: getDashboardData(),
      tasks: getTasks()
    };
  } finally {
    lock.releaseLock();
  }
}


function getTasks() {
  const rows = {
    tasks: getSheetObjects_(APP.SHEETS.TASKS),
    taskMembers: getSheetObjects_(APP.SHEETS.TASK_MEMBERS),
    taskUpdates: getSheetObjects_(APP.SHEETS.TASK_UPDATES),
    subtasks: getSheetObjects_(APP.SHEETS.SUBTASKS)
  };
  return buildTasks_(rows.tasks, rows.taskMembers, rows.taskUpdates, rows.subtasks);
}

function buildTasks_(tasks, taskMembers, taskUpdates, subtasks) {
  const subtaskMap = {};
  (subtasks || []).forEach((x) => {
    if (!subtaskMap[x.task_id]) subtaskMap[x.task_id] = [];
    subtaskMap[x.task_id].push(normalizeSubtask_(x));
  });

  const memberMap = {};
  (taskMembers || []).forEach((x) => {
    if (!memberMap[x.task_id]) memberMap[x.task_id] = [];
    memberMap[x.task_id].push({
      member: String(x.member_name),
      shareWeight: num_(x.share_weight)
    });
  });

  const updateMap = {};
  (taskUpdates || []).forEach((x) => {
    if (!updateMap[x.task_id]) updateMap[x.task_id] = [];
    updateMap[x.task_id].push(normalizeTaskUpdate_(x));
  });

  Object.keys(updateMap).forEach((taskId) => {
    updateMap[taskId].sort(compareTaskUpdatesDesc_);
  });

  return (tasks || [])
    .filter((t) => String(t.category || '').toLowerCase() !== 'consult')
    .map((t) => {
      const history = updateMap[t.task_id] || [];
      const latest = history.length ? history[0] : null;

      return {
        taskId: String(t.task_id),
        category: String(t.category || ''),
        job: String(t.job || ''),
        description: String(t.description || ''),
        holderName: String(t.holder_name || ''),
        startDate: dateOnly_(t.start_date),
        dueDate: dateOnly_(t.due_date),
        deadline: dateOnly_(t.deadline),
        weight: num_(t.weight),
        status: String(t.status || ''),
        createdAt: String(t.created_at || ''),
        updatedAt: String(t.updated_at || ''),
        members: memberMap[t.task_id] || [],
        subtaskCount: (subtaskMap[t.task_id] || []).length,
        subtaskDone: (subtaskMap[t.task_id] || []).filter((x) => x.status === 'Complete').length,
        subtaskProgress: subtaskProgress_(subtaskMap[t.task_id] || []),
        updateCount: history.length,
        latestUpdate: latest
      };
    })
    .sort((a, b) => taskActivityKey_(b).localeCompare(taskActivityKey_(a)));
}

function getTaskDetail(taskId) {
  // Read each related sheet once. Older code reread TaskUpdates/Subtasks after getTasks().
  const tasks = getSheetObjects_(APP.SHEETS.TASKS);
  const taskMembers = getSheetObjects_(APP.SHEETS.TASK_MEMBERS);
  const taskUpdates = getSheetObjects_(APP.SHEETS.TASK_UPDATES);
  const subtasks = getSheetObjects_(APP.SHEETS.SUBTASKS);

  const task = buildTasks_(tasks, taskMembers, taskUpdates, subtasks)
    .find((x) => String(x.taskId) === String(taskId));
  if (!task) throw new Error('ไม่พบ task_id: ' + taskId);

  const updates = taskUpdates
    .filter((x) => String(x.task_id) === String(taskId))
    .map(normalizeTaskUpdate_)
    .sort(compareTaskUpdatesDesc_);

  const subtaskRows = subtasks
    .filter((x) => String(x.task_id) === String(taskId))
    .map(normalizeSubtask_)
    .sort((a, b) => {
      const aOrder = num_(a.sortOrder);
      const bOrder = num_(b.sortOrder);
      if (aOrder !== bOrder) return aOrder - bOrder;
      return String(a.createdAt || '').localeCompare(String(b.createdAt || ''));
    });

  return { task, updates, subtasks: subtaskRows };
}

function getTaskUpdateHistory(taskId) {
  return getSheetObjects_(APP.SHEETS.TASK_UPDATES)
    .filter((x) => String(x.task_id) === String(taskId))
    .map(normalizeTaskUpdate_)
    .sort(compareTaskUpdatesDesc_);
}

function updateTask(payload) {
  if (!payload || !payload.taskId) throw new Error('กรุณาระบุงาน');
  validateTask_(payload);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getDb_();
    const taskSh = ss.getSheetByName(APP.SHEETS.TASKS);
    const values = taskSh.getDataRange().getValues();
    if (!values.length) throw new Error('ไม่พบข้อมูล Tasks');

    const headers = values[0].map(String);
    const idx = headerIndex_(headers);
    const rowIndex = values.findIndex((row, i) =>
      i > 0 && String(row[idx.task_id]) === String(payload.taskId)
    );

    if (rowIndex < 1) throw new Error('ไม่พบ task_id: ' + payload.taskId);

    const rowNum = rowIndex + 1;
    const beforeTask = rowToObject_(headers, values[rowIndex]);
    const now = nowString_();

    const updates = {
      category: String(payload.category || 'Other').toLowerCase() === 'consult' ? 'Other' : (payload.category || 'Other'),
      job: payload.job || '',
      description: payload.description || '',
      holder_name: payload.holder_name || '',
      start_date: payload.start_date || '',
      due_date: payload.due_date || '',
      deadline: payload.deadline || '',
      weight: num_(payload.weight),
      status: payload.status || 'Inprocess',
      updated_at: now
    };

    Object.entries(updates).forEach(([field, value]) => {
      if (idx[field] !== undefined) {
        taskSh.getRange(rowNum, idx[field] + 1).setValue(value);
      }
    });

    // Replace current assignees for this task.
    const tmSh = ss.getSheetByName(APP.SHEETS.TASK_MEMBERS);
    const tmValues = tmSh.getDataRange().getValues();
    const tmHeaders = tmValues[0].map(String);
    const tmIdx = headerIndex_(tmHeaders);

    for (let i = tmValues.length - 1; i >= 1; i--) {
      if (String(tmValues[i][tmIdx.task_id]) === String(payload.taskId)) {
        tmSh.deleteRow(i + 1);
      }
    }

    const members = [...new Set((payload.members || []).map(String).filter(Boolean))];
    const share = members.length ? num_(payload.weight) / members.length : 0;
    members.forEach((member) => {
      tmSh.appendRow([
        payload.taskId,
        member,
        round4_(share)
      ]);
    });

    const afterTask = {
      ...updates,
      task_id: payload.taskId,
      members: members
    };

    appendAudit_('UPDATE', 'TASK', payload.taskId, beforeTask, afterTask);

    // Optional latest-update note. Existing TaskUpdates schema is reused.
    const latestUpdate = payload.latestUpdate || {};
    const hasProgress = latestUpdate.progress !== '' &&
      latestUpdate.progress !== null &&
      latestUpdate.progress !== undefined;
    const hasNote = String(latestUpdate.note || '').trim() !== '';

    if (hasProgress || hasNote) {
      addTaskUpdate_({
        taskId: payload.taskId,
        updateDate: latestUpdate.updateDate || todayString_(),
        memberName: latestUpdate.memberName || '',
        progress: hasProgress ? latestUpdate.progress : '',
        note: latestUpdate.note || ''
      });
    }

    return {
      ok: true,
      dashboard: getDashboardData(),
      tasks: getTasks(),
      detail: getTaskDetail(payload.taskId)
    };
  } finally {
    lock.releaseLock();
  }
}

function updateTaskStatus(taskId, newStatus) {
  const sh = getDb_().getSheetByName(APP.SHEETS.TASKS);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idxTask = headers.indexOf('task_id');
  const idxStatus = headers.indexOf('status');
  const idxUpdated = headers.indexOf('updated_at');

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idxTask]) === String(taskId)) {
      const before = { status: values[i][idxStatus] };
      sh.getRange(i + 1, idxStatus + 1).setValue(newStatus);
      sh.getRange(i + 1, idxUpdated + 1).setValue(nowString_());
      appendAudit_('UPDATE_STATUS', 'TASK', taskId, before, { status: newStatus });
      return {
        ok: true,
        dashboard: getDashboardData(),
        tasks: getTasks()
      };
    }
  }
  throw new Error('ไม่พบ task_id: ' + taskId);
}


function addTaskUpdate(payload) {
  if (!payload || !payload.taskId) throw new Error('กรุณาระบุงาน');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const update = addTaskUpdate_(payload);
    return {
      ok: true,
      update,
      tasks: getTasks(),
      detail: getTaskDetail(payload.taskId)
    };
  } finally {
    lock.releaseLock();
  }
}

function addTaskUpdate_(payload) {
  const updateId = makeId_('UPD');
  const now = nowString_();
  const progress =
    payload.progress === '' || payload.progress === null || payload.progress === undefined
      ? ''
      : num_(payload.progress);

  appendRow_(APP.SHEETS.TASK_UPDATES, [
    updateId,
    payload.taskId,
    payload.updateDate || todayString_(),
    payload.memberName || '',
    progress,
    payload.note || '',
    now,
    getUser_()
  ]);

  appendAudit_('ADD_UPDATE', 'TASK_UPDATE', updateId, null, payload);

  return {
    updateId,
    taskId: String(payload.taskId),
    updateDate: payload.updateDate || todayString_(),
    memberName: payload.memberName || '',
    progress,
    note: payload.note || '',
    createdAt: now,
    createdBy: getUser_()
  };
}

function normalizeTaskUpdate_(x) {
  return {
    updateId: String(x.update_id || ''),
    taskId: String(x.task_id || ''),
    updateDate: dateOnly_(x.update_date),
    memberName: String(x.member_name || ''),
    progress:
      x.progress === '' || x.progress === null || x.progress === undefined
        ? ''
        : num_(x.progress),
    note: String(x.note || ''),
    createdAt: String(x.created_at || ''),
    createdBy: String(x.created_by || '')
  };
}

function compareTaskUpdatesDesc_(a, b) {
  return taskUpdateSortKey_(b).localeCompare(taskUpdateSortKey_(a));
}

function taskUpdateSortKey_(x) {
  return String(x.updateDate || x.createdAt || '');
}

function taskActivityKey_(task) {
  const latest = task.latestUpdate
    ? taskUpdateSortKey_(task.latestUpdate)
    : '';
  return String(latest || task.updatedAt || task.createdAt || '');
}

function dateOnly_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, APP.TZ, 'yyyy-MM-dd');
  }

  const text = String(value).trim();
  const iso = text.match(/^(\d{4}-\d{2}-\d{2})/);
  return iso ? iso[1] : text;
}


/* =========================
   SUBTASKS
========================= */

function getSubtasks(taskId) {
  return getSheetObjects_(APP.SHEETS.SUBTASKS)
    .filter((x) => String(x.task_id) === String(taskId))
    .map(normalizeSubtask_)
    .sort((a, b) => {
      const aOrder = num_(a.sortOrder);
      const bOrder = num_(b.sortOrder);
      if (aOrder !== bOrder) return aOrder - bOrder;
      return String(a.createdAt || '').localeCompare(String(b.createdAt || ''));
    });
}

function addSubtask(payload) {
  if (!payload || !payload.taskId) throw new Error('กรุณาระบุ Task');
  if (!String(payload.title || '').trim()) throw new Error('กรุณาระบุชื่อ Subtask');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const current = getSubtasks(payload.taskId);
    const nextOrder = current.length
      ? Math.max.apply(null, current.map((x) => num_(x.sortOrder))) + 1
      : 1;
    const now = nowString_();
    const subtaskId = makeId_('SUB');

    appendRow_(APP.SHEETS.SUBTASKS, [
      subtaskId,
      payload.taskId,
      String(payload.title || '').trim(),
      String(payload.description || '').trim(),
      payload.dueDate || '',
      payload.status || 'not start',
      nextOrder,
      now,
      getUser_(),
      now
    ]);

    appendAudit_('CREATE', 'SUBTASK', subtaskId, null, payload);
    return {
      ok: true,
      detail: getTaskDetail(payload.taskId),
      tasks: getTasks()
    };
  } finally {
    lock.releaseLock();
  }
}

function updateSubtask(payload) {
  if (!payload || !payload.subtaskId) throw new Error('กรุณาระบุ Subtask');
  if (!String(payload.title || '').trim()) throw new Error('กรุณาระบุชื่อ Subtask');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = getDb_().getSheetByName(APP.SHEETS.SUBTASKS);
    const values = sh.getDataRange().getValues();
    const headers = values[0].map(String);
    const idx = headerIndex_(headers);

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idx.subtask_id]) !== String(payload.subtaskId)) continue;
      const before = rowToObject_(headers, values[i]);
      const rowNum = i + 1;
      const updates = {
        title: String(payload.title || '').trim(),
        description: String(payload.description || '').trim(),
        due_date: payload.dueDate || '',
        status: payload.status || 'not start',
        updated_at: nowString_()
      };
      Object.entries(updates).forEach(([field, value]) => {
        sh.getRange(rowNum, idx[field] + 1).setValue(value);
      });
      appendAudit_('UPDATE', 'SUBTASK', payload.subtaskId, before, updates);
      return {
        ok: true,
        detail: getTaskDetail(before.task_id),
        tasks: getTasks()
      };
    }
    throw new Error('ไม่พบ Subtask: ' + payload.subtaskId);
  } finally {
    lock.releaseLock();
  }
}

function deleteSubtask(subtaskId) {
  if (!subtaskId) throw new Error('กรุณาระบุ Subtask');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = getDb_().getSheetByName(APP.SHEETS.SUBTASKS);
    const values = sh.getDataRange().getValues();
    const headers = values[0].map(String);
    const idx = headerIndex_(headers);

    for (let i = values.length - 1; i >= 1; i--) {
      if (String(values[i][idx.subtask_id]) !== String(subtaskId)) continue;
      const before = rowToObject_(headers, values[i]);
      sh.deleteRow(i + 1);
      appendAudit_('DELETE', 'SUBTASK', subtaskId, before, null);
      return {
        ok: true,
        detail: getTaskDetail(before.task_id),
        tasks: getTasks()
      };
    }
    throw new Error('ไม่พบ Subtask: ' + subtaskId);
  } finally {
    lock.releaseLock();
  }
}

function normalizeSubtask_(x) {
  return {
    subtaskId: String(x.subtask_id || ''),
    taskId: String(x.task_id || ''),
    title: String(x.title || ''),
    description: String(x.description || ''),
    dueDate: dateOnly_(x.due_date),
    status: String(x.status || 'not start'),
    sortOrder: num_(x.sort_order),
    createdAt: String(x.created_at || ''),
    createdBy: String(x.created_by || ''),
    updatedAt: String(x.updated_at || '')
  };
}

function subtaskProgress_(items) {
  if (!items || !items.length) return 0;
  const done = items.filter((x) => String(x.status) === 'Complete').length;
  return round2_((done / items.length) * 100);
}

/* =========================
   CONSULT
========================= */

function getConsultData(filters) {
  const range = normalizeConsultRange_(filters || {});
  const members = getMembers_();
  const consultRows = getSheetObjects_(APP.SHEETS.CONSULT);
  return buildConsultData_(filters || range, members, consultRows);
}

function buildConsultData_(filters, members, consultRows) {
  const range = normalizeConsultRange_(filters || {});
  const memberFilter = String((filters || {}).memberName || '');
  const rows = (consultRows || [])
    .map(normalizeConsultEntry_)
    .filter((x) => dateInRange_(x.consultDate, range.startDate, range.endDate))
    .filter((x) => !memberFilter || x.memberName === memberFilter)
    .sort((a, b) => {
      const d = String(b.consultDate).localeCompare(String(a.consultDate));
      return d || String(a.memberName).localeCompare(String(b.memberName), 'th');
    });

  const summaryMap = {};
  (members || []).forEach((m) => {
    summaryMap[m.name] = {member: m.name, count: 0};
  });
  rows.forEach((x) => {
    if (!summaryMap[x.memberName]) summaryMap[x.memberName] = {member: x.memberName, count: 0};
    summaryMap[x.memberName].count += x.count;
  });

  const summary = Object.values(summaryMap)
    .map((x) => ({member: x.member, count: round2_(x.count)}))
    .sort((a, b) => b.count - a.count || a.member.localeCompare(b.member, 'th'));

  return {
    startDate: range.startDate,
    endDate: range.endDate,
    memberName: memberFilter,
    entries: rows,
    summary,
    total: round2_(rows.reduce((sum, x) => sum + x.count, 0)),
    activeMembers: summary.filter((x) => x.count > 0).length
  };
}

function saveConsultWeek(payload) {
  if (!payload || !/^\d{4}-\d{2}-\d{2}$/.test(String(payload.date || ''))) {
    throw new Error('กรุณาเลือกวันที่ในสัปดาห์ Consult');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const range = weekRangeForDate_(String(payload.date));
    const sh = getDb_().getSheetByName(APP.SHEETS.CONSULT);
    const values = sh.getDataRange().getValues();
    const headers = values[0].map(String);
    const idx = headerIndex_(headers);
    const before = [];
    const keptRows = [];

    for (let i = 1; i < values.length; i++) {
      const rowDate = dateOnly_(values[i][idx.consult_date]);
      if (dateInRange_(rowDate, range.startDate, range.endDate)) {
        before.push(rowToObject_(headers, values[i]));
      } else if (values[i].some((v) => v !== '')) {
        keptRows.push(values[i].slice(0, headers.length));
      }
    }

    const now = nowString_();
    const user = getUser_();
    const after = [];
    const newRows = [];

    (payload.counts || []).forEach((item) => {
      const member = String(item.memberName || '').trim();
      const raw = Number(item.count);
      if (!member || !Number.isFinite(raw) || raw <= 0) return;
      if (Math.floor(raw) !== raw) throw new Error('จำนวน Consult ต้องเป็นจำนวนเต็ม');

      const row = [makeId_('CON'), range.startDate, member, raw, '', now, user, now];
      newRows.push(row);
      after.push({
        consult_date: range.startDate,
        week_end: range.endDate,
        member_name: member,
        consult_count: raw,
        workload_weight: raw
      });
    });

    // Batch replace the selected week instead of deleteRow()/appendRow() repeatedly.
    const outputRows = keptRows.concat(newRows);
    const existingDataRows = Math.max(0, sh.getLastRow() - 1);
    if (existingDataRows > 0) {
      sh.getRange(2, 1, existingDataRows, headers.length).clearContent();
    }
    if (outputRows.length) {
      sh.getRange(2, 1, outputRows.length, headers.length).setValues(outputRows);
    }

    appendAudit_('REPLACE_WEEK', 'CONSULT', `${range.startDate}_${range.endDate}`, before, after);

    // Read the post-save core data once and reuse it for both Consult and Dashboard.
    const core = readCoreData_();
    const members = membersFromRows_(core.members);
    const settings = settingsFromRows_(core.settings);
    return {
      ok: true,
      week: range,
      consult: buildConsultData_(range, members, core.consult),
      dashboard: buildDashboardData_(members, settings, core.tasks, core.taskMembers, core.consult, currentWorkWeekRange_())
    };
  } finally {
    lock.releaseLock();
  }
}

// Backward-compatible alias for an older cached frontend.
function saveConsultDay(payload) {
  return saveConsultWeek(payload);
}

function getConsultWeek(date) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(date || ''))) {
    throw new Error('กรุณาเลือกวันที่ในสัปดาห์ Consult');
  }
  const range = weekRangeForDate_(String(date));
  const data = getConsultData(range);
  data.week = range;
  return data;
}

function normalizeConsultEntry_(x) {
  return {
    consultId: String(x.consult_id || ''),
    consultDate: dateOnly_(x.consult_date),
    memberName: String(x.member_name || ''),
    count: num_(x.consult_count),
    sourceTaskId: String(x.source_task_id || ''),
    createdAt: String(x.created_at || ''),
    createdBy: String(x.created_by || ''),
    updatedAt: String(x.updated_at || '')
  };
}

function normalizeConsultRange_(filters) {
  if (filters && filters.startDate && filters.endDate) {
    return {startDate: String(filters.startDate), endDate: String(filters.endDate)};
  }
  return currentWorkWeekRange_();
}

function currentWorkWeekRange_() {
  return weekRangeForDate_(todayString_());
}

function weekRangeForDate_(date) {
  const parts = String(date).split('-').map(Number);
  if (parts.length !== 3 || !parts[0] || !parts[1] || !parts[2]) {
    throw new Error('รูปแบบวันที่ไม่ถูกต้อง');
  }

  // Use midday to reduce date rollover edge cases.
  const d = new Date(parts[0], parts[1] - 1, parts[2], 12, 0, 0);
  const day = d.getDay(); // Sun=0, Mon=1 ... Sat=6
  const deltaToMonday = day === 0 ? -6 : 1 - day;

  const monday = new Date(d);
  monday.setDate(d.getDate() + deltaToMonday);

  const friday = new Date(monday);
  friday.setDate(monday.getDate() + 4);

  return {
    startDate: Utilities.formatDate(monday, APP.TZ, 'yyyy-MM-dd'),
    endDate: Utilities.formatDate(friday, APP.TZ, 'yyyy-MM-dd')
  };
}

function dateInRange_(date, startDate, endDate) {
  if (!date || !/^\d{4}-\d{2}-\d{2}$/.test(String(date))) return false;
  return String(date) >= String(startDate) && String(date) <= String(endDate);
}

/* =========================
   PAGE 2: OT
========================= */

function generateOtMonth(year, month) {
  year = Number(year);
  month = Number(month);
  if (!year || month < 1 || month > 12) throw new Error('เดือน/ปีไม่ถูกต้อง');

  const settings = getSettings_();
  const existing = getSheetObjects_(APP.SHEETS.OT_SCHEDULE);
  const existingDates = new Set(existing.map((x) => normalizeDateOnly_(x.work_date)).filter(Boolean));

  const lastDay = new Date(year, month, 0).getDate();
  const now = nowString_();
  const rowsToCreate = [];

  for (let day = 1; day <= lastDay; day++) {
    const jsDate = new Date(year, month - 1, day);
    const dateStr = Utilities.formatDate(jsDate, APP.TZ, 'yyyy-MM-dd');
    if (existingDates.has(dateStr)) continue;

    const dow = jsDate.getDay();
    const isWeekend = dow === 0 || dow === 6;
    const start = isWeekend ? String(settings.WEEKEND_START || '08:00') : String(settings.WEEKDAY_START || '16:00');
    const end = isWeekend ? String(settings.WEEKEND_END || '16:00') : String(settings.WEEKDAY_END || '18:00');
    const breakHours = isWeekend ? num_(settings.WEEKEND_BREAK) : num_(settings.WEEKDAY_BREAK);
    const plannedHours = calcHours_(start, end, breakHours);

    rowsToCreate.push([
      makeId_('SCH'), dateStr, start, end, breakHours, plannedHours,
      isWeekend, true, '', now, now
    ]);
  }

  // One batch write for the whole month instead of appendRow() 28-31 times.
  appendRows_(APP.SHEETS.OT_SCHEDULE, rowsToCreate);
  const created = rowsToCreate.length;

  appendAudit_('GENERATE_MONTH', 'OT_SCHEDULE', `${year}-${month}`, null, { year, month, created });
  return getOtMonth(year, month);
}

function getOtMonth(year, month) {
  year = Number(year);
  month = Number(month);
  if (!year || month < 1 || month > 12) throw new Error('เดือน/ปีไม่ถูกต้อง');

  const monthKey = `${year}-${String(month).padStart(2, '0')}`;
  const settings = getSettings_();
  const members = getMembers_();

  // Normalize persisted schedule values because Google Sheets can convert
  // date/time-looking strings into Date objects.
  const storedSchedule = getSheetObjects_(APP.SHEETS.OT_SCHEDULE)
    .map((x) => {
      const date = normalizeDateOnly_(x.work_date);
      const note = String(x.note || '');
      const rawStart = normalizeTime_(x.start_time);
      const rawEnd = normalizeTime_(x.end_time);
      const parts = date ? date.split('-').map(Number) : [];
      const calculatedWeekend = parts.length === 3
        ? [0, 6].includes(new Date(parts[0], parts[1] - 1, parts[2]).getDay())
        : false;
      const isWeekend = x.is_weekend === '' || x.is_weekend === null || x.is_weekend === undefined
        ? calculatedWeekend
        : bool_(x.is_weekend);
      const legacyPlaceholder = /^Legacy placeholder/i.test(note) && !rawStart && !rawEnd;
      const defaultStart = isWeekend
        ? normalizeTime_(settings.WEEKEND_START || '08:00')
        : normalizeTime_(settings.WEEKDAY_START || '16:00');
      const defaultEnd = isWeekend
        ? normalizeTime_(settings.WEEKEND_END || '16:00')
        : normalizeTime_(settings.WEEKDAY_END || '18:00');
      const defaultBreak = isWeekend
        ? num_(settings.WEEKEND_BREAK)
        : num_(settings.WEEKDAY_BREAK);
      const start = rawStart || (legacyPlaceholder ? defaultStart : '');
      const end = rawEnd || (legacyPlaceholder ? defaultEnd : '');
      const breakHours = legacyPlaceholder ? defaultBreak : num_(x.break_hours);
      const active = legacyPlaceholder ? true : bool_(x.active);
      const plannedHours = legacyPlaceholder
        ? calcHours_(start, end, breakHours)
        : num_(x.planned_hours);

      return {
        scheduleId: String(x.schedule_id || ''),
        date,
        start,
        end,
        breakHours,
        plannedHours,
        isWeekend,
        active,
        note: legacyPlaceholder ? '' : note,
        saved: !legacyPlaceholder,
        legacyPlaceholder
      };
    })
    .filter((x) => x.date && x.date.startsWith(monthKey));

  const scheduleMap = {};
  storedSchedule.forEach((x) => { scheduleMap[x.date] = x; });

  // Always return every calendar day. A missing row gets a virtual default
  // so the web UI never shows a blank/missing date.
  const lastDay = new Date(year, month, 0).getDate();
  const schedule = [];
  for (let day = 1; day <= lastDay; day++) {
    const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    if (scheduleMap[dateStr]) {
      schedule.push(scheduleMap[dateStr]);
      continue;
    }

    const jsDate = new Date(year, month - 1, day);
    const isWeekend = [0, 6].includes(jsDate.getDay());
    const start = isWeekend
      ? normalizeTime_(settings.WEEKEND_START || '08:00')
      : normalizeTime_(settings.WEEKDAY_START || '16:00');
    const end = isWeekend
      ? normalizeTime_(settings.WEEKEND_END || '16:00')
      : normalizeTime_(settings.WEEKDAY_END || '18:00');
    const breakHours = isWeekend
      ? num_(settings.WEEKEND_BREAK)
      : num_(settings.WEEKDAY_BREAK);

    schedule.push({
      scheduleId: '',
      date: dateStr,
      start,
      end,
      breakHours,
      plannedHours: calcHours_(start, end, breakHours),
      isWeekend,
      active: true,
      note: '',
      saved: false
    });
  }

  const entries = getSheetObjects_(APP.SHEETS.OT_ENTRIES)
    .map((x) => ({
      entryId: String(x.entry_id || ''),
      date: normalizeDateOnly_(x.work_date),
      member: String(x.member_name || ''),
      start: normalizeTime_(x.start_time),
      end: normalizeTime_(x.end_time),
      breakHours: num_(x.break_hours),
      hours: num_(x.hours),
      roles: String(x.roles || ''),
      rate: num_(x.rate_per_hour),
      amount: num_(x.amount),
      note: String(x.note || '')
    }))
    .filter((x) => x.date && x.date.startsWith(monthKey));

  const entriesByDate = {};
  entries.forEach((e) => {
    if (!entriesByDate[e.date]) entriesByDate[e.date] = [];
    entriesByDate[e.date].push(e);
  });

  const summaryMap = {};
  members.forEach((m) => {
    summaryMap[m.name] = {
      member: m.name,
      hours: 0,
      amount: 0,
      weekdayDays: new Set(),
      weekendDays: new Set()
    };
  });

  let actualHours = 0;
  let actualPayout = 0;

  entries.forEach((e) => {
    if (!summaryMap[e.member]) {
      summaryMap[e.member] = {
        member: e.member,
        hours: 0,
        amount: 0,
        weekdayDays: new Set(),
        weekendDays: new Set()
      };
    }
    summaryMap[e.member].hours += e.hours;
    summaryMap[e.member].amount += e.amount;
    actualHours += e.hours;
    actualPayout += e.amount;

    const parts = e.date.split('-').map(Number);
    const d = new Date(parts[0], parts[1] - 1, parts[2]);
    const dow = d.getDay();
    if (dow === 0 || dow === 6) summaryMap[e.member].weekendDays.add(e.date);
    else summaryMap[e.member].weekdayDays.add(e.date);
  });

  const plannedHours = schedule
    .filter((x) => x.active)
    .reduce((sum, x) => sum + num_(x.plannedHours), 0);

  const rate = num_(settings.OT_RATE || 100);

  const summary = Object.values(summaryMap).map((x) => ({
    member: x.member,
    hours: round2_(x.hours),
    amount: round2_(x.amount),
    weekdayDays: x.weekdayDays.size,
    weekendDays: x.weekendDays.size
  }));

  return {
    year,
    month,
    rate,
    schedule: schedule.map((day) => ({
      ...day,
      entries: entriesByDate[day.date] || []
    })),
    orphanEntries: entries.filter((e) => !scheduleMap[e.date] && !schedule.some((d) => d.date === e.date)),
    summary,
    totals: {
      plannedHours: round2_(plannedHours),
      plannedBudget: round2_(plannedHours * rate),
      actualHours: round2_(actualHours),
      actualPayout: round2_(actualPayout)
    }
  };
}

function saveOtScheduleDay(payload) {
  if (!payload || !payload.date) throw new Error('กรุณาระบุวันที่');

  const start = payload.start || '';
  const end = payload.end || '';
  const breakHours = num_(payload.breakHours);
  const plannedHours = payload.active === false ? 0 : calcHours_(start, end, breakHours);

  const sh = getDb_().getSheetByName(APP.SHEETS.OT_SCHEDULE);
  const data = sh.getDataRange().getValues();
  const headers = data[0].map(String);
  const idx = headerIndex_(headers);

  let rowIndex = -1;
  let before = null;
  for (let i = 1; i < data.length; i++) {
    if (normalizeDateOnly_(data[i][idx.work_date]) === normalizeDateOnly_(payload.date)) {
      rowIndex = i;
      before = rowToObject_(headers, data[i]);
      break;
    }
  }

  const now = nowString_();
  const safeDate = normalizeDateOnly_(payload.date);
  const parts = safeDate.split('-').map(Number);
  const jsDate = new Date(parts[0], parts[1] - 1, parts[2]);
  const isWeekend = [0, 6].includes(jsDate.getDay());

  if (rowIndex === -1) {
    appendRow_(APP.SHEETS.OT_SCHEDULE, [
      makeId_('SCH'), safeDate, start, end, breakHours, plannedHours,
      isWeekend, payload.active !== false, payload.note || '', now, now
    ]);
  } else {
    // One row write instead of 7 separate setValue() calls.
    const row = data[rowIndex].slice(0, headers.length);
    row[idx.start_time] = start;
    row[idx.end_time] = end;
    row[idx.break_hours] = breakHours;
    row[idx.planned_hours] = plannedHours;
    row[idx.is_weekend] = isWeekend;
    row[idx.active] = payload.active !== false;
    row[idx.note] = payload.note || '';
    row[idx.updated_at] = now;
    sh.getRange(rowIndex + 1, 1, 1, headers.length).setValues([row]);
  }

  appendAudit_('UPSERT', 'OT_SCHEDULE', safeDate, before, payload);
  return { ok: true };
}

function saveOtDay(payload) {
  if (!payload || !payload.date) throw new Error('กรุณาระบุวันที่');
  payload.date = normalizeDateOnly_(payload.date);
  if (!payload.date) throw new Error('รูปแบบวันที่ OT ไม่ถูกต้อง');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    if (payload.schedule) {
      saveOtScheduleDay({
        date: payload.date,
        start: payload.schedule.start,
        end: payload.schedule.end,
        breakHours: payload.schedule.breakHours,
        active: payload.schedule.active,
        note: payload.schedule.note
      });
    }

    const sh = getDb_().getSheetByName(APP.SHEETS.OT_ENTRIES);
    const data = sh.getDataRange().getValues();
    const headers = data[0].map(String);
    const idx = headerIndex_(headers);

    const oldRows = [];
    const keptRows = [];
    for (let i = 1; i < data.length; i++) {
      if (normalizeDateOnly_(data[i][idx.work_date]) === payload.date) {
        oldRows.push(rowToObject_(headers, data[i]));
      } else if (data[i].some((v) => v !== '')) {
        keptRows.push(data[i].slice(0, headers.length));
      }
    }

    const rateDefault = num_(getSettings_().OT_RATE || 100);
    const now = nowString_();
    const user = getUser_();
    const newRows = [];

    (payload.entries || []).forEach((entry) => {
      if (!entry.member) return;
      const start = entry.start || '';
      const end = entry.end || '';
      const breakHours = num_(entry.breakHours);
      const hours = calcHours_(start, end, breakHours);
      const rate = entry.rate !== undefined && entry.rate !== '' ? num_(entry.rate) : rateDefault;
      const amount = round2_(hours * rate);

      newRows.push([
        makeId_('OTE'), payload.date, entry.member, start, end, breakHours, hours,
        Array.isArray(entry.roles) ? entry.roles.join(', ') : (entry.roles || ''),
        rate, amount, entry.note || '', now, user, now
      ]);
    });

    // Replace all OT entry rows in one batch. This is much faster than
    // deleteRow()/appendRow() for every member.
    const outputRows = keptRows.concat(newRows);
    const existingDataRows = Math.max(0, sh.getLastRow() - 1);
    if (existingDataRows > 0) {
      sh.getRange(2, 1, existingDataRows, headers.length).clearContent();
    }
    if (outputRows.length) {
      sh.getRange(2, 1, outputRows.length, headers.length).setValues(outputRows);
    }

    appendAudit_('REPLACE_DAY', 'OT_ENTRY', payload.date, oldRows, payload.entries || []);

    const [year, month] = payload.date.split('-').map(Number);
    return { ok: true, monthData: getOtMonth(year, month) };
  } finally {
    lock.releaseLock();
  }
}

/* =========================
   AUDIT + BACKUP
========================= */

function getAuditLogs(limit) {
  // PERFORMANCE: read only the last N physical rows instead of getDataRange()
  // over the entire AuditLog sheet.
  const sh = getDb_().getSheetByName(APP.SHEETS.AUDIT);
  if (!sh || sh.getLastRow() < 2) return [];

  const n = Math.max(1, Number(limit || 100));
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const startRow = Math.max(2, lastRow - n + 1);
  const values = sh.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();

  return values
    .filter((row) => row.some((v) => v !== ''))
    .map((row) => rowToObject_(headers, row))
    .reverse()
    .map((x) => ({
      logId: String(x.log_id || ''),
      timestamp: String(x.timestamp || ''),
      user: String(x.user || ''),
      action: String(x.action || ''),
      module: String(x.module || ''),
      recordId: String(x.record_id || ''),
      before: String(x.before_json || ''),
      after: String(x.after_json || '')
    }));
}

function weeklyBackup() {
  const ss = getDb_();
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty(APP.BACKUP_FOLDER_PROP);

  if (!folderId) {
    const folder = DriveApp.createFolder('TeamWorkloadOT_Backup');
    folderId = folder.getId();
    props.setProperty(APP.BACKUP_FOLDER_PROP, folderId);
  }

  const folder = DriveApp.getFolderById(folderId);
  const stamp = Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM-dd_HHmmss');
  const baseName = `TeamWorkloadOT_${stamp}`;

  let googleCopyName = '';
  let xlsxName = '';
  let status = 'SUCCESS';
  let message = '';

  try {
    const srcFile = DriveApp.getFileById(ss.getId());
    const copy = srcFile.makeCopy(baseName, folder);
    googleCopyName = copy.getName();

    try {
      const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=xlsx`;
      const response = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      });

      if (response.getResponseCode() >= 200 && response.getResponseCode() < 300) {
        const blob = response.getBlob().setName(baseName + '.xlsx');
        const xlsx = folder.createFile(blob);
        xlsxName = xlsx.getName();
      } else {
        status = 'PARTIAL';
        message = `Google Sheet copy สำเร็จ แต่ export XLSX ไม่สำเร็จ: HTTP ${response.getResponseCode()}`;
      }
    } catch (xlsxErr) {
      status = 'PARTIAL';
      message = 'Google Sheet copy สำเร็จ แต่ export XLSX ไม่สำเร็จ: ' + xlsxErr.message;
    }
  } catch (err) {
    status = 'FAILED';
    message = err.message;
  }

  appendRow_(APP.SHEETS.BACKUP, [
    makeId_('BKP'),
    nowString_(),
    googleCopyName,
    xlsxName,
    status,
    message
  ]);

  appendAudit_('BACKUP', 'SYSTEM', baseName, null, {
    googleCopyName,
    xlsxName,
    status,
    message
  });

  return { googleCopyName, xlsxName, status, message };
}

/* =========================
   HELPERS
========================= */

let DB_HANDLE_CACHE_ = null;

function getDb_() {
  // Reuse the Spreadsheet handle within the same Apps Script execution.
  if (DB_HANDLE_CACHE_) return DB_HANDLE_CACHE_;
  const id = PropertiesService.getScriptProperties().getProperty(APP.DB_PROP);
  if (!id) {
    throw new Error('ยังไม่ได้สร้างฐานข้อมูลใหม่ กรุณารัน setupNewSystem() ก่อน');
  }
  DB_HANDLE_CACHE_ = SpreadsheetApp.openById(id);
  return DB_HANDLE_CACHE_;
}

function getMembers_() {
  // Schema upgrades are setup/upgrade work, not normal read work.
  return membersFromRows_(getSheetObjects_(APP.SHEETS.MEMBERS));
}

function membersFromRows_(rows) {
  return (rows || [])
    .filter((x) => bool_(x.active))
    .sort((a, b) => num_(a.sort_order) - num_(b.sort_order))
    .map((x) => ({
      id: String(x.member_id),
      name: String(x.name),
      workStatus: memberWorkStatus_(x.work_status),
      statusNote: String(x.status_note || defaultMemberStatusNote_(memberWorkStatus_(x.work_status))),
      statusUpdatedAt: String(x.status_updated_at || '')
    }));
}

function memberWorkStatus_(value) {
  if (value === '' || value === null || value === undefined) return 1;
  const n = Number(value);
  if (!Number.isFinite(n)) return 1;
  return Math.max(0, Math.min(1, n));
}

function getSettings_() {
  return settingsFromRows_(getSheetObjects_(APP.SHEETS.SETTINGS));
}

function settingsFromRows_(rows) {
  const out = {};
  (rows || []).forEach((x) => { out[String(x.key)] = x.value; });
  return out;
}

function getSheetObjects_(sheetName) {
  const sh = getDb_().getSheetByName(sheetName);
  if (!sh) throw new Error('ไม่พบ sheet: ' + sheetName);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  // Explicit used range avoids scanning accidental formatting outside the table.
  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(String);
  return values.slice(1)
    .filter((row) => row.some((v) => v !== ''))
    .map((row) => rowToObject_(headers, row));
}

function rowToObject_(headers, row) {
  const obj = {};
  headers.forEach((h, i) => {
    let v = row[i];
    if (v instanceof Date) {
      v = Utilities.formatDate(v, APP.TZ, 'yyyy-MM-dd HH:mm:ss');
    }
    obj[h] = v;
  });
  return obj;
}

function appendRow_(sheetName, row) {
  appendRows_(sheetName, [row]);
}

function appendRows_(sheetName, rows) {
  if (!Array.isArray(rows) || !rows.length) return;
  const sh = getDb_().getSheetByName(sheetName);
  if (!sh) throw new Error('ไม่พบ sheet: ' + sheetName);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function appendAudit_(action, module, recordId, beforeData, afterData) {
  try {
    appendRow_(APP.SHEETS.AUDIT, [
      makeId_('LOG'),
      nowString_(),
      getUser_(),
      action,
      module,
      recordId || '',
      safeJson_(beforeData),
      safeJson_(afterData)
    ]);
  } catch (err) {
    console.error('Audit failed:', err);
  }
}

function validateTask_(payload) {
  if (!payload) throw new Error('ไม่พบข้อมูล');
  if (!String(payload.job || '').trim()) throw new Error('กรุณาระบุชื่องาน');
  if (!Array.isArray(payload.members) || !payload.members.length) {
    throw new Error('กรุณาเลือกผู้รับผิดชอบอย่างน้อย 1 คน');
  }
  if (num_(payload.weight) < 0) throw new Error('Weight ต้องไม่ติดลบ');
}

function calcHours_(start, end, breakHours) {
  if (!start || !end) return 0;
  const s = timeToMinutes_(start);
  const e = timeToMinutes_(end);
  if (e < s) throw new Error(`เวลาสิ้นสุดต้องมากกว่าเวลาเริ่ม: ${start} - ${end}`);

  const hours = ((e - s) / 60) - num_(breakHours);
  if (hours < 0) throw new Error('ชั่วโมงหลังหักพักติดลบ');
  return round2_(hours);
}

function timeToMinutes_(value) {
  const parts = String(value).split(':');
  if (parts.length < 2) throw new Error('รูปแบบเวลาต้องเป็น HH:mm');
  return Number(parts[0]) * 60 + Number(parts[1]);
}

function headerIndex_(headers) {
  const out = {};
  headers.forEach((h, i) => out[String(h)] = i);
  return out;
}

function makeId_(prefix) {
  const stamp = Utilities.formatDate(new Date(), APP.TZ, 'yyyyMMddHHmmss');
  const rand = Math.floor(Math.random() * 9000 + 1000);
  return `${prefix}_${stamp}_${rand}`;
}

function getUser_() {
  if (REQUEST_ACTOR_) return REQUEST_ACTOR_;

  const email = Session.getActiveUser().getEmail();
  if (email) return email;

  const tempKey = Session.getTemporaryActiveUserKey();
  return tempKey ? 'user:' + tempKey : 'Netlify API';
}

function sanitizeActor_(value) {
  const actor = String(value || '').trim().replace(/[\r\n\t]+/g, ' ');
  return actor.slice(0, 80);
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function nowString_() {
  return Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM-dd HH:mm:ss');
}

function todayString_() {
  return Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM-dd');
}

function safeJson_(value) {
  if (value === null || value === undefined) return '';
  try {
    const text = JSON.stringify(value);
    return text.length > 45000 ? text.slice(0, 45000) : text;
  } catch (err) {
    return String(value);
  }
}

function normalizeDateOnly_(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) return Utilities.formatDate(value, APP.TZ, 'yyyy-MM-dd');
  const text = String(value).trim();
  const iso = text.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
  const slash = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (slash) {
    const d = String(Number(slash[1])).padStart(2, '0');
    const m = String(Number(slash[2])).padStart(2, '0');
    return `${slash[3]}-${m}-${d}`;
  }
  return text.slice(0, 10);
}

function normalizeTime_(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) return Utilities.formatDate(value, APP.TZ, 'HH:mm');
  const text = String(value).trim();
  const m = text.match(/(?:^|\s)(\d{1,2}):(\d{2})(?::\d{2})?(?:\s|$)/);
  if (m) return `${String(Number(m[1])).padStart(2, '0')}:${m[2]}`;
  const simple = text.match(/^(\d{1,2})(?:\.(\d+))?$/);
  if (simple) {
    const h = Number(simple[1]);
    const mins = simple[2] ? Math.round(Number('0.' + simple[2]) * 60) : 0;
    return `${String(h).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
  }
  return text.slice(-5);
}

function num_(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}

function bool_(v) {
  if (typeof v === 'boolean') return v;
  return ['true', '1', 'yes', 'y'].includes(String(v).toLowerCase());
}

function round2_(n) {
  return Math.round((Number(n) + Number.EPSILON) * 100) / 100;
}

function round4_(n) {
  return Math.round((Number(n) + Number.EPSILON) * 10000) / 10000;
}

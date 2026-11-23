/**
 * 수행평가 제출함 — Google Apps Script 백엔드
 *
 * 구조
 *   내 드라이브/수행평가 제출함/          ← 루트 폴더 (setup()이 자동 생성)
 *     ├─ 제출기록 (스프레드시트)          ← 모든 제출 로그
 *     ├─ 1학기 국어 독서감상문/           ← 수행평가 1개 = 하위 폴더 1개
 *     │    ├─ 10101_홍길동.pdf
 *     │    └─ 10102_김철수.hwp
 *     └─ 2학기 과학 탐구보고서/
 *
 * 관리자 페이지에서 수행평가를 만들면 기준 폴더 안에 하위 폴더가 생기고 바로 학생 화면에 나타납니다.
 * 드라이브에서 직접 만든 폴더는 관리자 페이지에서 '다시 열기'를 눌러야 학생에게 보입니다.
 */

/***** 설정 — 필요하면 바꾸세요 *****/
const CONFIG = {
  ROOT_FOLDER_ID: '1KRtCgOHfaD10E8AZOPcWtkLdrgvC_wgb',                // 기준 폴더 ID (비우면 setup()이 '수행평가 제출함' 폴더를 새로 만듦)
  ROOT_FOLDER_NAME: '수행평가 제출함',
  MAX_TOTAL_MB: 25,                    // 한 번에 올릴 수 있는 총 용량 (Apps Script 요청 한도 때문에 30MB 이하 권장)
  MAX_FILES: 5,                        // 한 번에 올릴 수 있는 파일 수
  STUDENT_ID_PATTERN: /^\d{4,6}$/,     // 학번 형식 (예: 10101, 30215)
  NAME_PATTERN: /^[가-힣a-zA-Z]{1,20}$/,
  OVERWRITE_ON_RESUBMIT: true,         // 재제출 시 같은 학번_이름의 이전 파일을 휴지통으로 이동
  LOG_SHEET_NAME: '수행평가 제출기록',
};

const PROPS = PropertiesService.getScriptProperties();

/**
 * 처음 한 번 실행하세요. 루트 폴더·제출기록 시트·관리자 비밀번호를 만듭니다.
 * 실행 후 [실행 로그]에서 관리자 비밀번호를 확인하세요.
 */
function setup() {
  const root = getRoot_(true);
  if (!PROPS.getProperty('ADMIN_PASSWORD')) {
    PROPS.setProperty('ADMIN_PASSWORD', Utilities.getUuid().replace(/-/g, '').slice(0, 10));
  }
  getLogSheet_();
  Logger.log('루트 폴더: ' + root.getUrl());
  Logger.log('관리자 비밀번호: ' + PROPS.getProperty('ADMIN_PASSWORD'));
  Logger.log('비밀번호 변경: [프로젝트 설정 > 스크립트 속성]의 ADMIN_PASSWORD 값을 수정하세요.');
}

/***** 웹 앱 진입점 *****/
function doGet(e) {
  return respond_(() => {
    const action = (e && e.parameter && e.parameter.action) || 'list';
    if (action === 'list') return { assignments: listAssignments_(false) };
    if (action === 'ping') return { pong: true };
    throw new Error('알 수 없는 요청입니다.');
  });
}

function doPost(e) {
  return respond_(() => {
    if (!e || !e.postData || !e.postData.contents) throw new Error('요청 내용이 비어 있습니다.');
    const d = JSON.parse(e.postData.contents);
    switch (d.action) {
      case 'submit':       return submit_(d);
      case 'admin.login':  requireAdmin_(d); return {};
      case 'admin.list':   requireAdmin_(d); return { assignments: listAssignments_(true) };
      case 'admin.create': requireAdmin_(d); return { assignment: createAssignment_(d) };
      case 'admin.update': requireAdmin_(d); return { assignment: updateAssignment_(d) };
      case 'admin.files':  requireAdmin_(d); return { files: listSubmissions_(d.assignmentId) };
      default: throw new Error('알 수 없는 요청입니다.');
    }
  });
}

function respond_(fn) {
  let out;
  try {
    out = Object.assign({ ok: true }, fn());
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/***** 학생 제출 *****/
function submit_(d) {
  const studentId = String(d.studentId || '').trim();
  const name = String(d.name || '').replace(/\s+/g, '');
  const files = Array.isArray(d.files) ? d.files : [];

  if (!CONFIG.STUDENT_ID_PATTERN.test(studentId)) throw new Error('학번 형식이 올바르지 않습니다.');
  if (!CONFIG.NAME_PATTERN.test(name)) throw new Error('이름은 한글 또는 영문으로만 입력하세요.');
  if (files.length === 0) throw new Error('제출할 파일을 첨부하세요.');
  if (files.length > CONFIG.MAX_FILES) throw new Error('파일은 최대 ' + CONFIG.MAX_FILES + '개까지 올릴 수 있습니다.');

  const totalBytes = files.reduce((s, f) => s + Math.floor(String(f.data || '').length * 3 / 4), 0);
  if (totalBytes > CONFIG.MAX_TOTAL_MB * 1024 * 1024) throw new Error('총 용량이 ' + CONFIG.MAX_TOTAL_MB + 'MB를 넘습니다.');

  const folder = getAssignmentFolder_(d.assignmentId);
  const meta = readMeta_(folder);
  if (!isAccepting_(meta)) throw new Error('제출이 마감된 수행평가입니다.');

  const base = studentId + '_' + name;
  const saved = [];
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (CONFIG.OVERWRITE_ON_RESUBMIT) {
      const it = folder.getFiles();
      while (it.hasNext()) {
        const f = it.next();
        const n = f.getName();
        if (n === base || n.indexOf(base + '.') === 0 || n.indexOf(base + '_') === 0) f.setTrashed(true);
      }
    }
    files.forEach((f, i) => {
      const ext = extOf_(f.name);
      const fileName = (files.length > 1 ? base + '_' + (i + 1) : base) + ext;
      const bytes = Utilities.base64Decode(String(f.data || ''));
      if (bytes.length === 0) throw new Error('빈 파일은 제출할 수 없습니다: ' + f.name);
      const blob = Utilities.newBlob(bytes, f.mimeType || 'application/octet-stream', fileName);
      const created = folder.createFile(blob);
      saved.push({ name: fileName, url: created.getUrl() });
    });
  } finally {
    lock.releaseLock();
  }

  try {
    const sheet = getLogSheet_();
    const now = new Date();
    saved.forEach(s => sheet.appendRow([now, folder.getName(), "'" + studentId, name, s.name, s.url]));
  } catch (err) {
    console.warn('로그 기록 실패: ' + err);
  }

  return { saved: saved.map(s => s.name), assignment: folder.getName() };
}

function extOf_(fileName) {
  const m = String(fileName || '').match(/\.([A-Za-z0-9]{1,10})$/);
  return m ? '.' + m[1].toLowerCase() : '';
}

/***** 수행평가(폴더) 관리 *****/
function listAssignments_(isAdmin) {
  const root = getRoot_(false);
  const it = root.getFolders();
  const list = [];
  while (it.hasNext()) {
    const f = it.next();
    const item = toItem_(f, readMeta_(f));
    if (isAdmin) {
      item.url = f.getUrl();
      item.fileCount = countFiles_(f);
      list.push(item);
    } else if (item.accepting) {
      list.push(item);
    }
  }
  list.sort((a, b) => b.created.localeCompare(a.created));
  return list;
}

function createAssignment_(d) {
  const name = String(d.name || '').trim();
  if (!name) throw new Error('수행평가 이름을 입력하세요.');
  if (/[\\/]/.test(name)) throw new Error('이름에 / 또는 \\ 는 쓸 수 없습니다.');
  const root = getRoot_(false);
  if (root.getFoldersByName(name).hasNext()) throw new Error('같은 이름의 수행평가가 이미 있습니다.');
  const folder = root.createFolder(name);
  const meta = { open: true, deadline: d.deadline || '', note: String(d.note || '') };
  writeMeta_(folder, meta);
  return toItem_(folder, meta);
}

function updateAssignment_(d) {
  const folder = getAssignmentFolder_(d.assignmentId);
  const meta = readMeta_(folder);
  if ('open' in d) meta.open = !!d.open;
  if ('deadline' in d) meta.deadline = d.deadline || '';
  if ('note' in d) meta.note = String(d.note || '');
  if (d.name && String(d.name).trim() !== folder.getName()) folder.setName(String(d.name).trim());
  writeMeta_(folder, meta);
  return toItem_(folder, meta);
}

function listSubmissions_(assignmentId) {
  const folder = getAssignmentFolder_(assignmentId);
  const it = folder.getFiles();
  const out = [];
  while (it.hasNext()) {
    const f = it.next();
    const m = f.getName().match(/^(\d+)_([^_.]+)/);
    out.push({
      fileName: f.getName(),
      studentId: m ? m[1] : '',
      studentName: m ? m[2] : '',
      updated: f.getLastUpdated().toISOString(),
      size: f.getSize(),
      url: f.getUrl(),
    });
  }
  out.sort((a, b) => a.studentId.localeCompare(b.studentId) || a.fileName.localeCompare(b.fileName));
  return out;
}

function toItem_(folder, meta) {
  return {
    id: folder.getId(),
    name: folder.getName(),
    open: meta.open,
    deadline: meta.deadline,
    note: meta.note,
    accepting: isAccepting_(meta),
    created: folder.getDateCreated().toISOString(),
  };
}

function isAccepting_(meta) {
  if (!meta.open) return false;
  if (meta.deadline && new Date(meta.deadline).getTime() < Date.now()) return false;
  return true;
}

// 폴더 설명란에 JSON으로 상태를 저장합니다.
// 앱으로 만들지 않은 폴더(설명이 없거나 JSON이 아님)는 '마감' 상태로 취급해 학생에게 보이지 않습니다.
// 드라이브에서 직접 만든 폴더는 관리자 페이지에서 '다시 열기'를 누르면 제출을 받기 시작합니다.
function readMeta_(folder) {
  const meta = { open: false, deadline: '', note: '' };
  const desc = folder.getDescription() || '';
  if (!desc) return meta;
  try {
    const parsed = JSON.parse(desc);
    if (parsed && typeof parsed === 'object') return Object.assign(meta, parsed);
  } catch (e) { /* 일반 설명 문구 */ }
  return meta;
}

function writeMeta_(folder, meta) {
  folder.setDescription(JSON.stringify({ open: meta.open, deadline: meta.deadline, note: meta.note }));
}

function countFiles_(folder) {
  let n = 0;
  const it = folder.getFiles();
  while (it.hasNext()) { it.next(); n++; }
  return n;
}

/***** 공통 *****/
function getRoot_(create) {
  const id = CONFIG.ROOT_FOLDER_ID || PROPS.getProperty('ROOT_FOLDER_ID');
  if (id) {
    try {
      const f = DriveApp.getFolderById(id);
      if (!f.isTrashed()) return f;
    } catch (e) { /* 폴더가 지워졌으면 아래에서 새로 만듦 */ }
  }
  if (!create) throw new Error('서버 설정이 끝나지 않았습니다. 선생님께 setup() 실행을 요청하세요.');
  const folder = DriveApp.createFolder(CONFIG.ROOT_FOLDER_NAME);
  PROPS.setProperty('ROOT_FOLDER_ID', folder.getId());
  return folder;
}

function getAssignmentFolder_(id) {
  if (!id) throw new Error('수행평가를 선택하세요.');
  let folder;
  try {
    folder = DriveApp.getFolderById(String(id));
  } catch (e) {
    throw new Error('수행평가를 찾을 수 없습니다.');
  }
  const rootId = getRoot_(false).getId();
  let inRoot = false;
  const parents = folder.getParents();
  while (parents.hasNext()) if (parents.next().getId() === rootId) inRoot = true;
  if (!inRoot || folder.isTrashed()) throw new Error('수행평가를 찾을 수 없습니다.');
  return folder;
}

function requireAdmin_(d) {
  const pw = PROPS.getProperty('ADMIN_PASSWORD');
  if (!pw || String(d.password || '') !== pw) {
    Utilities.sleep(800);
    throw new Error('관리자 비밀번호가 틀렸습니다.');
  }
}

function getLogSheet_() {
  const id = PROPS.getProperty('LOG_SHEET_ID');
  let ss = null;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create(CONFIG.LOG_SHEET_NAME);
    DriveApp.getFileById(ss.getId()).moveTo(getRoot_(true));
    PROPS.setProperty('LOG_SHEET_ID', ss.getId());
    const sh = ss.getSheets()[0];
    sh.appendRow(['제출시각', '수행평가', '학번', '이름', '파일명', '파일 링크']);
    sh.setFrozenRows(1);
  }
  return ss.getSheets()[0];
}

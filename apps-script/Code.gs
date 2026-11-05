/**
 * 2학년 8반 사이트 — 구글 시트 연결 코드
 * (칭찬 기록부 + 기말고사 목표 트리)
 *
 * 처음 설치
 * 1. 구글 시트에서 [확장 프로그램] → [Apps Script]를 엽니다.
 * 2. 이 파일 내용을 모두 붙여 넣고 저장합니다.
 * 3. 함수 목록에서 setup 을 골라 [실행]합니다. (권한 허용 필요)
 * 4. [배포] → [새 배포] → 웹 앱 (실행: 나 / 액세스: 모든 사용자)
 *
 * 이미 칭찬 기록부를 쓰고 있다면 (업데이트)
 * 1. 기존 코드를 모두 지우고 이 파일 내용으로 바꾼 뒤 저장합니다.
 * 2. setup 을 한 번 더 실행합니다. (기존 칭찬 기록은 그대로 남습니다)
 * 3. [배포] → [배포 관리] → 연필(수정) → 버전: 새 버전 → [배포]
 *    → 웹 앱 주소는 바뀌지 않습니다.
 *
 * '설정' 시트
 *   B2  학생에게 칭찬 누적 기록 보이기 (체크박스)
 *   B3  학급 이름
 *   B4  교사용 비밀번호 (목표 트리의 '선생님 모드'에서 사용)
 *   B5  기말고사 첫날 (목표 트리에 D-day 표시)
 */

const RECORD_SHEET = '칭찬기록';
const SETTING_SHEET = '설정';
const SUMMARY_SHEET = '학생별 합계';
const GOAL_SHEET = '기말목표';
const HEADERS = ['작성 시각', '날짜', '번호', '칭찬받은 학생', '칭찬 영역', '구체적인 행동', '쓴 사람', '기록 ID'];
const GOAL_HEADERS = ['작성 시각', '번호', '이름', '장식', '나의 목표', '실천 방법', '도장', '도장 찍은 시각', '목표 ID', '수정키'];
const CATS = ['학업·탐구', '배려·나눔', '봉사', '리더십', '질서·생활', '성실·책임', '기타'];
const DECOS = ['별', '하트', '방울', '종', '리본', '눈송이', '선물상자', '사탕지팡이', '달', '도토리'];

/* 기말목표 시트의 열 번호 (1부터) */
const G = { time: 1, num: 2, name: 3, deco: 4, goal: 5, plan: 6, stamp: 7, stampAt: 8, id: 9, token: 10 };

/** 처음 한 번 실행 (다시 실행해도 기존 기록은 지워지지 않습니다). */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const rec = ss.getSheetByName(RECORD_SHEET) || ss.insertSheet(RECORD_SHEET, 0);
  if (rec.getLastRow() === 0) rec.appendRow(HEADERS);
  rec.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#e6edf6');
  rec.setFrozenRows(1);
  rec.setColumnWidth(6, 420);
  rec.getRange('F:F').setWrap(true);
  rec.hideColumns(8);

  const set = ss.getSheetByName(SETTING_SHEET) || ss.insertSheet(SETTING_SHEET);
  set.getRange('A1:B1').setValues([['항목', '값']]).setFontWeight('bold');
  set.getRange('A2').setValue('학생에게 칭찬 누적 기록 보이기');
  if (set.getRange('B2').getValue() === '') set.getRange('B2').insertCheckboxes().setValue(false);
  set.getRange('A3').setValue('학급 이름');
  if (set.getRange('B3').getValue() === '') set.getRange('B3').setValue('2학년 8반');
  set.getRange('A4').setValue('교사용 비밀번호 (목표 트리)');
  if (String(set.getRange('B4').getValue()).trim() === '') {
    set.getRange('B4').setNumberFormat('@').setValue(String(Math.floor(100000 + Math.random() * 900000)));
  }
  set.getRange('A5').setValue('기말고사 첫날 (D-day 표시)');
  if (set.getRange('B5').getValue() === '') set.getRange('B5').setNumberFormat('yyyy-mm-dd');
  set.setColumnWidth(1, 240);

  const sum = ss.getSheetByName(SUMMARY_SHEET) || ss.insertSheet(SUMMARY_SHEET);
  sum.clear();
  sum.getRange('A1').setFormula(
    "=IFERROR(QUERY({'" + RECORD_SHEET + "'!C2:E}, " +
    "\"select Col1, Col2, count(Col2) where Col2 <> '' group by Col1, Col2 pivot Col3 order by Col1 label Col1 '번호', Col2 '이름'\", 0), \"아직 기록이 없습니다\")"
  );

  const goal = ss.getSheetByName(GOAL_SHEET) || ss.insertSheet(GOAL_SHEET);
  if (goal.getLastRow() === 0) goal.appendRow(GOAL_HEADERS);
  goal.getRange(1, 1, 1, GOAL_HEADERS.length).setFontWeight('bold').setBackground('#fdf1e7');
  goal.setFrozenRows(1);
  goal.setColumnWidth(G.goal, 360);
  goal.setColumnWidth(G.plan, 300);
  goal.getRange(1, G.goal, goal.getMaxRows(), 2).setWrap(true);
  goal.hideColumns(G.id, 2);

  const bad = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (bad && bad.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(bad);
}

/* ───────────── 읽기 ───────────── */

function doGet(e) {
  const type = e && e.parameter ? e.parameter.type : '';
  if (type === 'goals') return json(publicGoals());
  return json(praiseData());
}

function praiseData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const set = ss.getSheetByName(SETTING_SHEET);
  const show = set ? set.getRange('B2').getValue() === true : false;
  const out = { ok: true, showRecords: show, className: className(), records: [], total: 0 };
  const rec = ss.getSheetByName(RECORD_SHEET);
  if (rec && rec.getLastRow() > 1) {
    const rows = rec.getRange(2, 1, rec.getLastRow() - 1, HEADERS.length).getValues();
    out.total = rows.length;
    if (show) {
      out.records = rows.map(r => ({
        date: fmtDate(r[1]), num: r[2], name: r[3], category: r[4], content: r[5], writer: r[6], id: r[7]
      }));
    }
  }
  return out;
}

/** 학생에게는 이름·장식·도장 여부만 보냅니다. 목표 내용은 절대 보내지 않습니다. */
function publicGoals() {
  return {
    ok: true,
    className: className(),
    examDate: examDate(),
    goals: goalRows().map(g => ({ id: g.id, num: g.num, name: g.name, deco: g.deco, stamped: g.stamped }))
  };
}

/* ───────────── 쓰기 ───────────── */

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    switch (d.action) {
      case 'goal': return json(saveGoal(d));
      case 'myGoal': return json(myGoal(d));
      case 'teacher': return json(teacherGoals(d));
      case 'stamp': return json(stampGoal(d));
      default: return json(savePraise(d)); // 칭찬 기록부 (action 없음)
    }
  } catch (err) {
    return json({ ok: false, error: '처리하지 못했습니다. 잠시 후 다시 시도해 주세요.' });
  }
}

function savePraise(d) {
  const name = clip(d.name, 20), content = clip(d.content, 1000), writer = clip(d.writer, 20);
  if (!name || !content || !writer) return { ok: false, error: '친구 이름, 행동, 쓴 사람을 모두 적어 주세요.' };
  const num = Number(d.num);
  const category = CATS.indexOf(d.category) >= 0 ? d.category : '기타';
  const date = /^\d{4}-\d{2}-\d{2}$/.test(d.date) ? d.date : fmtDate(new Date());
  const id = Utilities.getUuid();
  return withLock(() => {
    const rec = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RECORD_SHEET);
    if (!rec) return { ok: false, error: '선생님이 아직 시트 준비(setup)를 하지 않았어요.' };
    rec.appendRow([new Date(), date, num > 0 && num < 100 ? num : '', safe(name), category, safe(content), safe(writer), id]);
    return { ok: true, id: id };
  });
}

/** 목표 새로 달기 또는 (수정키가 맞으면) 고치기 */
function saveGoal(d) {
  const name = clip(d.name, 20), goal = clip(d.goal, 500), plan = clip(d.plan, 500);
  if (!name || !goal) return { ok: false, error: '이름과 나의 목표를 적어 주세요.' };
  const num = Number(d.num);
  const deco = DECOS.indexOf(d.deco) >= 0 ? d.deco : DECOS[0];
  return withLock(() => {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GOAL_SHEET);
    if (!sh) return { ok: false, error: '선생님이 아직 시트 준비(setup)를 하지 않았어요.' };
    if (d.id) {
      const row = findRow(sh, d.id);
      if (!row || String(sh.getRange(row, G.token).getValue()) !== String(d.token || '')) {
        return { ok: false, error: '이 목표를 고칠 수 없어요. 처음 쓴 기기에서만 고칠 수 있어요.' };
      }
      if (sh.getRange(row, G.stamp).getValue() === true) {
        return { ok: false, error: '도장을 받은 목표는 고칠 수 없어요.' };
      }
      sh.getRange(row, G.num, 1, 5).setValues([[num > 0 && num < 100 ? num : '', safe(name), deco, safe(goal), safe(plan)]]);
      return { ok: true, id: d.id, token: d.token };
    }
    const id = Utilities.getUuid();
    const token = Utilities.getUuid();
    sh.appendRow([new Date(), num > 0 && num < 100 ? num : '', safe(name), deco, safe(goal), safe(plan), false, '', id, token]);
    sh.getRange(sh.getLastRow(), G.stamp).insertCheckboxes().setValue(false);
    return { ok: true, id: id, token: token };
  });
}

/** 학생 본인 목표 다시 보기 (수정키가 맞을 때만) */
function myGoal(d) {
  const g = goalRows().filter(x => x.id === d.id && x.token === String(d.token || ''))[0];
  if (!g) return { ok: false };
  return { ok: true, goal: { id: g.id, num: g.num, name: g.name, deco: g.deco, goal: g.goal, plan: g.plan, stamped: g.stamped } };
}

function teacherGoals(d) {
  if (!checkKey(d.key)) return { ok: false, error: '비밀번호가 맞지 않아요.' };
  return {
    ok: true,
    goals: goalRows().map(g => ({
      id: g.id, num: g.num, name: g.name, deco: g.deco, goal: g.goal, plan: g.plan,
      stamped: g.stamped, stampAt: g.stampAt, time: g.time
    }))
  };
}

function stampGoal(d) {
  if (!checkKey(d.key)) return { ok: false, error: '비밀번호가 맞지 않아요.' };
  return withLock(() => {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GOAL_SHEET);
    const row = sh && findRow(sh, d.id);
    if (!row) return { ok: false, error: '목표를 찾지 못했어요.' };
    const on = d.stamped !== false;
    sh.getRange(row, G.stamp).setValue(on);
    sh.getRange(row, G.stampAt).setValue(on ? new Date() : '');
    return { ok: true, id: d.id, stamped: on };
  });
}

/* ───────────── 도우미 ───────────── */

function goalRows() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GOAL_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, GOAL_HEADERS.length).getValues()
    .filter(r => r[G.id - 1] && r[G.name - 1])
    .map(r => ({
      time: r[G.time - 1] instanceof Date ? r[G.time - 1].getTime() : 0,
      num: r[G.num - 1], name: String(r[G.name - 1]), deco: String(r[G.deco - 1]),
      goal: String(r[G.goal - 1]), plan: String(r[G.plan - 1]),
      stamped: r[G.stamp - 1] === true, stampAt: fmtDate(r[G.stampAt - 1]),
      id: String(r[G.id - 1]), token: String(r[G.token - 1])
    }));
}

function findRow(sh, id) {
  if (!id || sh.getLastRow() < 2) return 0;
  const ids = sh.getRange(2, G.id, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

function checkKey(key) {
  const set = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTING_SHEET);
  const real = set ? String(set.getRange('B4').getValue()).trim() : '';
  const ok = real !== '' && String(key || '').trim() === real;
  if (!ok) Utilities.sleep(800); // 비밀번호 마구 넣어 보기 방지
  return ok;
}

function className() {
  const set = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTING_SHEET);
  return set ? String(set.getRange('B3').getValue() || '') : '';
}

function examDate() {
  const set = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTING_SHEET);
  return set ? fmtDate(set.getRange('B5').getValue()) : '';
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function clip(v, n) { return String(v == null ? '' : v).trim().slice(0, n); }
/** '=', '+', '-', '@'로 시작하는 글이 시트에서 수식으로 실행되지 않게 막습니다. */
function safe(v) { return /^[=+\-@]/.test(v) ? "'" + v : v; }
function fmtDate(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '');
}

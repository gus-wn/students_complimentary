/**
 * 우리반 칭찬 기록부 — 구글 시트 연결 코드
 *
 * 사용법
 * 1. 새 구글 시트에서 [확장 프로그램] → [Apps Script]를 엽니다.
 * 2. 이 파일 내용을 모두 붙여 넣고 저장합니다.
 * 3. 위쪽 함수 목록에서 setup 을 골라 [실행]합니다. (처음 한 번, 권한 허용 필요)
 * 4. [배포] → [새 배포] → 유형: 웹 앱
 *      - 다음 사용자 인증 정보로 실행: 나
 *      - 액세스 권한이 있는 사용자: 모든 사용자
 *    → 배포 후 나오는 "웹 앱 URL"을 index.html 의 SCRIPT_URL 에 넣습니다.
 *
 * 누적 기록 공개 여부는 '설정' 시트의 B2 체크박스로 조정합니다.
 */

const RECORD_SHEET = '칭찬기록';
const SETTING_SHEET = '설정';
const SUMMARY_SHEET = '학생별 합계';
const HEADERS = ['작성 시각', '날짜', '번호', '칭찬받은 학생', '칭찬 영역', '구체적인 행동', '쓴 사람', '기록 ID'];
const CATS = ['학업·탐구', '배려·나눔', '봉사', '리더십', '질서·생활', '성실·책임', '기타'];

/** 처음 한 번 실행: 시트 세 개를 만들고 서식을 맞춥니다. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let rec = ss.getSheetByName(RECORD_SHEET) || ss.insertSheet(RECORD_SHEET, 0);
  if (rec.getLastRow() === 0) rec.appendRow(HEADERS);
  rec.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#e6edf6');
  rec.setFrozenRows(1);
  rec.setColumnWidth(6, 420);
  rec.getRange('F:F').setWrap(true);
  rec.hideColumns(8);

  let set = ss.getSheetByName(SETTING_SHEET) || ss.insertSheet(SETTING_SHEET);
  set.getRange('A1:B1').setValues([['항목', '값']]).setFontWeight('bold');
  set.getRange('A2').setValue('학생에게 누적 기록 보이기');
  if (set.getRange('B2').getValue() === '') {
    set.getRange('B2').insertCheckboxes().setValue(false);
  }
  set.getRange('A3').setValue('학급 이름');
  set.setColumnWidth(1, 220);

  let sum = ss.getSheetByName(SUMMARY_SHEET) || ss.insertSheet(SUMMARY_SHEET);
  sum.clear();
  sum.getRange('A1').setFormula(
    "=IFERROR(QUERY({'" + RECORD_SHEET + "'!C2:E}, " +
    "\"select Col1, Col2, count(Col2) where Col2 <> '' group by Col1, Col2 pivot Col3 order by Col1 label Col1 '번호', Col2 '이름'\", 0), \"아직 기록이 없습니다\")"
  );

  const bad = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (bad && bad.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(bad);
}

/** 페이지가 기록과 공개 설정을 읽어 갑니다. */
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const set = ss.getSheetByName(SETTING_SHEET);
  const show = set ? set.getRange('B2').getValue() === true : false;
  const className = set ? String(set.getRange('B3').getValue() || '') : '';
  const out = { ok: true, showRecords: show, className: className, records: [], total: 0 };

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
  return json(out);
}

/** 학생이 쓴 칭찬을 한 줄 추가합니다. */
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const name = clip(d.name, 20), content = clip(d.content, 1000), writer = clip(d.writer, 20);
    if (!name || !content || !writer) return json({ ok: false, error: '친구 이름, 행동, 쓴 사람을 모두 적어 주세요.' });
    const num = Number(d.num);
    const category = CATS.indexOf(d.category) >= 0 ? d.category : '기타';
    const date = /^\d{4}-\d{2}-\d{2}$/.test(d.date) ? d.date : fmtDate(new Date());
    const id = Utilities.getUuid();

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const rec = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RECORD_SHEET);
      if (!rec) return json({ ok: false, error: '선생님이 아직 시트 준비(setup)를 하지 않았어요.' });
      rec.appendRow([new Date(), date, num > 0 && num < 100 ? num : '', safe(name), category, safe(content), safe(writer), id]);
    } finally {
      lock.releaseLock();
    }
    return json({ ok: true, id: id });
  } catch (err) {
    return json({ ok: false, error: '저장하지 못했습니다. 잠시 후 다시 시도해 주세요.' });
  }
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

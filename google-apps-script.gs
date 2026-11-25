/**
 * 영우 ♥ 수정 청첩장 - 참석 여부 집계용 Google Apps Script
 * hanuri1655@gmail.com 계정의 구글 스프레드시트에 붙여 넣어 사용합니다.
 */
const SHEET_RESP = '응답';
const SHEET_SUM  = '집계';
const NOTIFY_EMAIL = '';           // 응답 올 때마다 메일 받고 싶으면 'hanuri1655@gmail.com' 입력
const HEADERS = ['제출시각', '구분', '참석여부', '이름', '관계', '인원'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getRespSheet_();
    const p = e.parameter || {};
    const clean = v => {
      v = String(v || '').trim().slice(0, 40);
      return /^[=+\-@]/.test(v) ? "'" + v : v;   // 수식 삽입 방지
    };
    const side = ['신랑측', '신부측'].includes(p.side) ? p.side : '';
    const attend = ['참석', '불참'].includes(p.attend) ? p.attend : '';
    const rel = ['가족', '친구', '직장동료', '지인'].includes(p.relation) ? p.relation : '';
    const count = attend === '참석' ? Math.max(1, Math.min(10, parseInt(p.count, 10) || 1)) : 0;
    if (!side || !attend || !rel || !clean(p.name)) throw new Error('invalid');

    sh.appendRow([new Date(), side, attend, clean(p.name), rel, count]);

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail(NOTIFY_EMAIL, `[청첩장] ${clean(p.name)} (${side}) ${attend} ${count}명`,
        `${side} / ${rel} / ${clean(p.name)} / ${attend} / ${count}명`);
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return json_({ ok: true, msg: 'RSVP endpoint is running' }); }

/** 처음 한 번 실행: 시트와 집계표를 만들어 줍니다. */
function setup() {
  getRespSheet_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sum = ss.getSheetByName(SHEET_SUM) || ss.insertSheet(SHEET_SUM);
  sum.clear();
  const R = `'${SHEET_RESP}'!`;
  const rows = [
    ['', '신랑측', '신부측', '합계'],
    ['참석 인원(명)',
      `=SUMIFS(${R}F:F,${R}B:B,"신랑측",${R}C:C,"참석")`,
      `=SUMIFS(${R}F:F,${R}B:B,"신부측",${R}C:C,"참석")`, '=B2+C2'],
    ['참석 응답(건)',
      `=COUNTIFS(${R}B:B,"신랑측",${R}C:C,"참석")`,
      `=COUNTIFS(${R}B:B,"신부측",${R}C:C,"참석")`, '=B3+C3'],
    ['불참 응답(건)',
      `=COUNTIFS(${R}B:B,"신랑측",${R}C:C,"불참")`,
      `=COUNTIFS(${R}B:B,"신부측",${R}C:C,"불참")`, '=B4+C4'],
    ['', '', '', ''],
    ['관계별 참석 인원', '신랑측', '신부측', '합계'],
  ];
  ['가족', '친구', '직장동료', '지인'].forEach((r, i) => {
    const n = 7 + i;
    rows.push([r,
      `=SUMIFS(${R}F:F,${R}B:B,"신랑측",${R}C:C,"참석",${R}E:E,"${r}")`,
      `=SUMIFS(${R}F:F,${R}B:B,"신부측",${R}C:C,"참석",${R}E:E,"${r}")`, `=B${n}+C${n}`]);
  });
  sum.getRange(1, 1, rows.length, 4).setValues(rows);
  sum.getRange('A1:D1').setFontWeight('bold').setBackground('#F6D3D3');
  sum.getRange('A6:D6').setFontWeight('bold').setBackground('#F6D3D3');
  sum.getRange('A1:A10').setFontWeight('bold');
  sum.setColumnWidth(1, 140);
}

function getRespSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_RESP);
  if (!sh) {
    sh = ss.insertSheet(SHEET_RESP, 0);
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#F9F4DE');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

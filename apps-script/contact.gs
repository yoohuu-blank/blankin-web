/**
 * 블랭크 홈페이지 상담 신청 접수 (Google Apps Script)
 *
 * 홈페이지 '상담하기' 폼에서 보낸 내용을 이 스프레드시트의 '상담 신청' 탭에 한 줄씩 쌓고,
 * NOTIFY_TO 주소로 알림 메일을 보냅니다.
 *
 * 설치
 * 1. 상담 신청을 받을 구글 시트에서 [확장 프로그램 → Apps Script]를 엽니다.
 * 2. 이 파일 내용을 통째로 붙여 넣고 저장합니다.
 * 3. [배포 → 새 배포 → 유형: 웹 앱], 실행 사용자: 나, 액세스 권한: 모든 사용자 → 배포.
 *    (처음 한 번 권한 승인 창이 뜹니다.)
 * 4. 나오는 웹 앱 URL(…/exec)을 index.html 의 CONTACT_URL 에 넣습니다.
 *
 * 코드를 고친 뒤에는 [배포 → 배포 관리 → 수정(연필) → 버전: 새 버전]으로 다시 배포해야 반영됩니다.
 * (이렇게 하면 URL은 그대로 유지됩니다.)
 */

// 알림 받을 주소. 여러 명이면 쉼표로 구분. 비워 두면 배포한 사람에게 갑니다.
const NOTIFY_TO = '';
const SHEET_NAME = '상담 신청';
const HEADERS = ['접수 시각', '이름 / 단체명', '연락처 또는 이메일', '문의 내용', '보낸 페이지', '처리 상태', '메모'];

function doPost(e) {
  const p = (e && e.parameter) || {};

  // 봇이 채우는 숨은 칸. 채워져 있으면 조용히 성공처럼 응답하고 버립니다.
  if (p.website) return json_({ ok: true });

  const name = clean_(p.name, 100);
  const contact = clean_(p.contact, 200);
  const msg = clean_(p.msg, 5000);
  const page = clean_(p.page, 300);
  if (!name || !contact) return json_({ ok: false, error: 'missing' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet_().appendRow([new Date(), safe_(name), safe_(contact), safe_(msg), safe_(page), '접수', '']);
  } finally {
    lock.releaseLock();
  }

  try {
    const to = NOTIFY_TO || Session.getEffectiveUser().getEmail();
    const mail = {
      to: to,
      subject: '[블랭크 상담 신청] ' + name,
      body: [
        '홈페이지로 상담 신청이 들어왔습니다.',
        '',
        '이름 / 단체명: ' + name,
        '연락처 또는 이메일: ' + contact,
        '',
        '문의 내용:',
        msg || '(비어 있음)',
        '',
        '전체 목록: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl(),
      ].join('\n'),
    };
    if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(contact)) mail.replyTo = contact;
    MailApp.sendEmail(mail);
  } catch (err) {
    // 메일이 실패해도 시트에는 저장됐으므로 접수는 성공으로 처리합니다.
    console.error(err);
  }

  return json_({ ok: true });
}

// 주소를 브라우저로 열었을 때 작동 여부만 확인용으로 보여 줍니다.
function doGet() {
  return json_({ ok: true, service: 'blank-contact' });
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME, 0);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sh;
}

function clean_(v, max) {
  return String(v == null ? '' : v).trim().slice(0, max);
}

// 시트 수식으로 해석되지 않도록 = + - @ 로 시작하는 값 앞에 ' 를 붙입니다.
function safe_(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 2026 조혈모세포 기증자 감사의 날 — RSVP 접수 스크립트
 *
 * 사용법
 *  1. 구글 스프레드시트 "2026 감사의 날 RSVP 회신 명단" 열기
 *  2. 확장 프로그램 → Apps Script → 기본 코드(Code.gs) 전부 지우고 이 파일 내용 붙여넣기 → 저장(💾)
 *  3. 배포 → 새 배포 → 유형 선택(⚙) → 웹 앱
 *       - 설명: RSVP
 *       - 다음 사용자 인증 정보로 실행: **나**
 *       - 액세스 권한이 있는 사용자: **모든 사용자**
 *     → 배포 → 액세스 승인(계정 선택 → 고급 → 안전하지 않은 페이지로 이동 → 허용)
 *  4. 나오는 "웹 앱 URL"(https://script.google.com/macros/s/…/exec) 을 복사해서 전달
 *
 * 시트는 본인만 열람 가능(기본 비공개). 초청장 페이지는 이 URL로 POST만 하므로
 * 수신자는 시트를 볼 수 없습니다.
 */

var SHEET_NAME = 'RSVP';            // 응답이 쌓일 시트 탭 이름 (없으면 첫 번째 탭 사용)
var TOKEN = 'hsc-thanksday-2026';   // 페이지의 RSVP_TOKEN 과 같아야 함 (스팸 방지용)
var NOTIFY_EMAIL = true;            // 회신이 들어올 때마다 본인 메일로 알림 (false 로 끄기)

var HEADERS = ['접수일시', '참석여부', '성함', '연락처', '동반인원', '동반자 성함', '총인원', '비고'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var data = JSON.parse(e.postData.contents || '{}');
    if (data.token !== TOKEN) return respond({ ok: false, error: 'bad token' });

    var name = String(data.name || '').trim();
    var phone = String(data.phone || '').trim();
    if (!name || !phone) return respond({ ok: false, error: 'missing fields' });

    var attend = data.attend !== false;
    // 동반자는 1인까지만 허용 (페이지에서도 제한하지만 서버에서 한 번 더 자름)
    var companions = (data.companions || []).map(function (c) { return String(c || '').trim(); }).filter(Boolean).slice(0, 1);
    var total = attend ? 1 + companions.length : 0;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    if (sh.getLastRow() === 0) sh.appendRow(HEADERS);

    sh.appendRow([
      new Date(),
      attend ? '참석' : '불참',
      name,
      phone,
      attend ? companions.length : 0,
      companions.join(', '),
      total,
      String(data.memo || '').trim()
    ]);
    var row = sh.getLastRow();
    sh.getRange(row, 1).setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange(row, 4).setNumberFormat('@');   // 연락처 앞 0 유지

    if (NOTIFY_EMAIL) {
      try {
        MailApp.sendEmail(
          Session.getEffectiveUser().getEmail(),
          '[RSVP] ' + name + ' — ' + (attend ? '참석 (' + total + '명)' : '불참'),
          '성함: ' + name + '\n연락처: ' + phone + '\n참석: ' + (attend ? '참석' : '불참') +
          '\n동반자: ' + (companions.join(', ') || '없음') + '\n총인원: ' + total +
          '\n\n시트: ' + ss.getUrl()
        );
      } catch (mailErr) { /* 메일 실패해도 접수는 유지 */ }
    }
    return respond({ ok: true, row: row });
  } catch (err) {
    return respond({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('RSVP endpoint OK').setMimeType(ContentService.MimeType.TEXT);
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

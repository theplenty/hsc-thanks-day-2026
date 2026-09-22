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
 * 시트는 본인만 열람 가능(기본 비공개). 초청장 페이지는 이 URL로 POST/GET만 하므로
 * 수신자는 시트를 볼 수 없습니다.
 *
 * 중복 방지: 연락처(숫자만 비교)를 키로 사용합니다. 같은 연락처로 다시 보내면
 * 새 행을 추가하지 않고 기존 행을 갱신하고 "수정일시"를 기록합니다.
 */

var SHEET_NAME = 'RSVP';            // 응답이 쌓일 시트 탭 이름 (없으면 첫 번째 탭 사용)
var TOKEN = 'hsc-thanksday-2026';   // 페이지의 RSVP_TOKEN 과 같아야 함 (스팸 방지용)
var NOTIFY_EMAIL = true;            // 회신이 들어올 때마다 본인 메일로 알림 (false 로 끄기)

var HEADERS = ['접수일시', '참석여부', '성함', '연락처', '동반인원', '동반자 성함', '동반자 연락처', '총인원', '비고', '수정일시'];
var COL = { date: 1, attend: 2, name: 3, phone: 4, cCount: 5, cName: 6, cPhone: 7, total: 8, memo: 9, updated: 10 };

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  // 헤더가 없거나 다르면 1행을 표준 헤더로 맞춤
  var cur = sh.getLastRow() ? sh.getRange(1, 1, 1, HEADERS.length).getValues()[0].join('|') : '';
  if (cur !== HEADERS.join('|')) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function digits(v) { return String(v || '').replace(/\D/g, ''); }

// 같은 연락처(숫자만 비교)로 이미 접수된 행 번호를 찾음. 없으면 0
function findRowByPhone(sh, phone) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var key = digits(phone);
  var vals = sh.getRange(2, COL.phone, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (digits(vals[i][0]) === key) return i + 2;
  }
  return 0;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var data = JSON.parse(e.postData.contents || '{}');
    if (data.token !== TOKEN) return respond({ ok: false, error: 'bad token' });

    var name = String(data.name || '').trim();
    var phone = String(data.phone || '').trim();
    if (!name || digits(phone).length < 9) return respond({ ok: false, error: 'missing fields' });

    var attend = data.attend !== false;
    // 동반자는 1인까지만 (페이지에서도 제한하지만 서버에서 한 번 더 자름)
    var comp = (data.companions || []).slice(0, 1).map(function (c) {
      if (typeof c === 'string') return { name: c.trim(), phone: '' };
      return { name: String(c.name || '').trim(), phone: String(c.phone || '').trim() };
    }).filter(function (c) { return c.name; });
    if (!attend) comp = [];
    var c = comp[0] || { name: '', phone: '' };
    var total = attend ? 1 + comp.length : 0;

    var sh = getSheet();
    var existing = findRowByPhone(sh, phone);
    var nowDate = new Date();
    var row;

    if (existing) {
      // 기존 회신 갱신 (접수일시는 유지, 수정일시 기록)
      row = existing;
      sh.getRange(row, COL.attend, 1, HEADERS.length - 1).setValues([[
        attend ? '참석' : '불참', name, phone, comp.length, c.name, c.phone, total, String(data.memo || '').trim(), nowDate
      ]]);
    } else {
      sh.appendRow([nowDate, attend ? '참석' : '불참', name, phone, comp.length, c.name, c.phone, total, String(data.memo || '').trim(), '']);
      row = sh.getLastRow();
    }
    sh.getRange(row, COL.date).setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange(row, COL.updated).setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange(row, COL.phone).setNumberFormat('@');    // 연락처 앞 0 유지
    sh.getRange(row, COL.cPhone).setNumberFormat('@');

    if (NOTIFY_EMAIL) {
      try {
        MailApp.sendEmail(
          Session.getEffectiveUser().getEmail(),
          '[RSVP' + (existing ? ' 수정' : '') + '] ' + name + ' — ' + (attend ? '참석 (' + total + '명)' : '불참'),
          '성함: ' + name + '\n연락처: ' + phone + '\n참석: ' + (attend ? '참석' : '불참') +
          '\n동반자: ' + (c.name ? c.name + ' (' + c.phone + ')' : '없음') + '\n총인원: ' + total +
          (existing ? '\n\n※ 같은 연락처의 기존 회신을 갱신했습니다. (행 ' + row + ')' : '') +
          '\n\n시트: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
        );
      } catch (mailErr) { /* 메일 실패해도 접수는 유지 */ }
    }
    return respond({ ok: true, row: row, updated: !!existing });
  } catch (err) {
    return respond({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// GET ?phone=01012345678&token=… → 해당 연락처의 회신 여부 조회 (이름은 가운데 글자 마스킹)
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!p.phone) return ContentService.createTextOutput('RSVP endpoint OK').setMimeType(ContentService.MimeType.TEXT);
  if (p.token !== TOKEN) return respond({ ok: false, error: 'bad token' });

  var sh = getSheet();
  var row = findRowByPhone(sh, p.phone);
  if (!row) return respond({ ok: true, found: false });

  var v = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];
  var name = String(v[COL.name - 1]);
  var masked = name.length >= 2 ? name[0] + '*'.repeat(name.length - 2) + name[name.length - 1] : name;
  var date = v[COL.updated - 1] instanceof Date ? v[COL.updated - 1] : v[COL.date - 1];
  return respond({
    ok: true, found: true,
    name: masked,
    attend: v[COL.attend - 1] === '참석',
    total: Number(v[COL.total - 1]) || 0,
    date: Utilities.formatDate(new Date(date), 'Asia/Seoul', 'M월 d일')
  });
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

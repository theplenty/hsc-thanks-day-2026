/**
 * 2026 조혈모세포 기증자 감사의 날 — RSVP 접수 스크립트  (v4)
 *
 * 처음 설치
 *  1. 구글 스프레드시트 "2026 감사의 날 RSVP 회신 명단" 열기
 *  2. 확장 프로그램 → Apps Script → 기본 코드(Code.gs) 전부 지우고 이 파일 내용 붙여넣기 → 저장(💾)
 *  3. 배포 → 새 배포 → 유형 선택(⚙) → 웹 앱
 *       - 다음 사용자 인증 정보로 실행: **나** / 액세스 권한이 있는 사용자: **모든 사용자**
 *     → 배포 → 액세스 승인
 *  4. 나오는 "웹 앱 URL"(https://script.google.com/macros/s/…/exec) 을 복사해서 전달
 *
 * 코드만 바꿀 때 (URL 유지)
 *  코드 교체 → 저장 → 배포 → 배포 관리 → ✎(수정) → 버전: "새 버전" → 배포
 *
 * 시트는 본인만 열람 가능(기본 비공개). 초청장 페이지는 이 URL로 POST/GET만 하므로
 * 수신자는 시트를 볼 수 없습니다.
 *
 * 중복 처리: 시트는 "추가만" 합니다(기존 행을 절대 수정·삭제하지 않음).
 * 같은 이름 + 휴대전화 뒤 4자리가 이미 있으면 새 행을 추가하지 않고 "이미 접수됨"으로 응답합니다.
 * 변경이 필요한 경우 담당자가 시트에서 직접 수정합니다.
 */

var SHEET_NAME = 'RSVP';            // 응답이 쌓일 시트 탭 이름 (없으면 첫 번째 탭 사용)
var TOKEN = 'hsc-thanksday-2026';   // 페이지의 RSVP_TOKEN 과 같아야 함 (스팸 방지용)
var NOTIFY_EMAIL = true;            // 회신이 들어올 때마다 본인 메일로 알림 (false 로 끄기)

var HEADERS = ['접수일시', '이름', '휴대전화 뒤4자리', '참석여부', '가족동반', '총인원', '비고'];
var COL = { date: 1, name: 2, last4: 3, attend: 4, family: 5, total: 6, memo: 7 };

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  // 헤더가 없거나 다르면 1행을 표준 헤더로 맞춤 (데이터 행은 건드리지 않음)
  var cur = sh.getLastRow() ? sh.getRange(1, 1, 1, HEADERS.length).getValues()[0].join('|') : '';
  if (cur !== HEADERS.join('|')) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  // 뒤4자리 열은 항상 텍스트 서식 (0001 의 앞 0 보존)
  sh.getRange(2, COL.last4, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('@');
  return sh;
}

function normName(s) { return String(s || '').replace(/\s+/g, '').trim(); }
// 숫자로 저장된 경우(1 → "0001")도 같은 값으로 취급
function normLast4(s) { var d = String(s || '').replace(/\D/g, ''); return d ? ('0000' + d).slice(-4) : ''; }

// 같은 이름 + 뒤4자리 행 번호 (없으면 0)
function findRow(sh, name, last4) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var vals = sh.getRange(2, COL.name, last - 1, 2).getValues();
  var n = normName(name), l = normLast4(last4);
  if (!n || !l) return 0;
  for (var i = 0; i < vals.length; i++) {
    if (normName(vals[i][0]) === n && normLast4(vals[i][1]) === l) return i + 2;
  }
  return 0;
}

function rowInfo(sh, row) {
  var v = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];
  var name = String(v[COL.name - 1]);
  var masked = name.length >= 2 ? name[0] + '*'.repeat(name.length - 2) + name[name.length - 1] : name;
  return {
    name: masked,
    attend: v[COL.attend - 1] === '참석',
    family: v[COL.family - 1] === '가족 1인 동반',
    total: Number(v[COL.total - 1]) || 0,
    date: Utilities.formatDate(new Date(v[COL.date - 1]), 'Asia/Seoul', 'M월 d일')
  };
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var data = JSON.parse(e.postData.contents || '{}');
    if (data.token !== TOKEN) return respond({ ok: false, error: 'bad token' });

    var name = String(data.name || '').trim();
    var last4 = normLast4(data.last4);
    if (!name || last4.length !== 4) return respond({ ok: false, error: 'missing fields' });

    var attend = data.attend !== false;
    var family = attend && data.family === true;
    var total = attend ? (family ? 2 : 1) : 0;

    var sh = getSheet();
    var existing = findRow(sh, name, last4);
    if (existing) {
      // 이미 접수됨 → 추가하지 않음
      var info = rowInfo(sh, existing);
      info.ok = false; info.duplicate = true;
      return respond(info);
    }

    var row = sh.getLastRow() + 1;
    sh.getRange(row, COL.last4).setNumberFormat('@');   // 값 쓰기 전에 텍스트 서식 → "0001" 그대로 저장
    sh.getRange(row, 1, 1, HEADERS.length).setValues([[
      new Date(), name, last4, attend ? '참석' : '불참',
      attend ? (family ? '가족 1인 동반' : '동반하지 않음') : '', total, String(data.memo || '').trim()
    ]]);
    sh.getRange(row, COL.date).setNumberFormat('yyyy-mm-dd hh:mm');

    if (NOTIFY_EMAIL) {
      try {
        MailApp.sendEmail(
          Session.getEffectiveUser().getEmail(),
          '[RSVP] ' + name + ' — ' + (attend ? '참석 (' + total + '명)' : '불참'),
          '이름: ' + name + '\n뒤 4자리: ' + last4 + '\n참석: ' + (attend ? '참석' : '불참') +
          '\n가족 동반: ' + (attend ? (family ? '가족 1인 동반' : '동반하지 않음') : '-') + '\n총인원: ' + total +
          '\n\n시트: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
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

// GET ?name=홍길동&last4=1234&token=… → 회신 여부 조회 (이름은 가운데 마스킹)
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!p.name && !p.last4) return ContentService.createTextOutput('RSVP endpoint OK').setMimeType(ContentService.MimeType.TEXT);
  if (p.token !== TOKEN) return respond({ ok: false, error: 'bad token' });

  var sh = getSheet();
  var row = findRow(sh, p.name, p.last4);
  if (!row) return respond({ ok: true, found: false });
  var info = rowInfo(sh, row);
  info.ok = true; info.found = true;
  return respond(info);
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ── 관리용: 이름이 "테스트" 인 행 전부 삭제. 편집기에서 함수 선택 → ▶ 실행
function deleteTestRows() {
  var sh = getSheet();
  var last = sh.getLastRow();
  var removed = 0;
  for (var r = last; r >= 2; r--) {
    if (normName(sh.getRange(r, COL.name).getValue()) === '테스트') { sh.deleteRow(r); removed++; }
  }
  Logger.log('삭제한 테스트 행: ' + removed);
}

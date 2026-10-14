/**
 * 영업과 커피 주문 - Google Apps Script 백엔드
 *
 * 설치 방법은 coffee/README.md 참고.
 * 1) 구글 시트 새로 만들기 → 확장 프로그램 → Apps Script
 * 2) 이 파일 내용을 통째로 붙여넣기 (ADMIN_PIN 원하는 번호로 변경)
 * 3) 배포 → 새 배포 → 유형: 웹 앱 / 실행: 나 / 액세스: 모든 사용자 → 배포
 * 4) 나온 웹 앱 URL(…/exec)을 coffee/index.html 의 SCRIPT_URL 에 붙여넣기
 */

const ADMIN_PIN = '1234'; // 주문 전체 초기화용 비밀번호 (꼭 바꾸세요)
const SHEET_NAME = 'orders';
const HEADER = ['round', 'name', 'menu', 'temp', 'memo', 'updatedAt'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADER);
    // 날짜/숫자로 자동 변환되지 않도록 텍스트 서식 지정
    sheet.getRange('A:E').setNumberFormat('@');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// 시트가 값을 날짜로 바꿔 저장했더라도 문자열로 맞춰 비교
function str_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v == null ? '' : v).trim();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const round = str_(e.parameter.round);
  const values = getSheet_().getDataRange().getValues().slice(1);
  const orders = values
    .filter(function (r) { return str_(r[0]) === round; })
    .map(function (r) {
      return {
        name: str_(r[1]),
        menu: str_(r[2]),
        temp: str_(r[3]),
        memo: str_(r[4]),
        updatedAt: r[5] instanceof Date ? r[5].toISOString() : str_(r[5]),
      };
    });
  return json_({ ok: true, orders: orders });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sheet = getSheet_();
    const values = sheet.getDataRange().getValues();
    const round = str_(d.round).slice(0, 30);
    const name = str_(d.name).slice(0, 20);

    if (!round) return json_({ ok: false, error: 'round 없음' });

    if (d.action === 'clear') {
      if (str_(d.pin) !== ADMIN_PIN) return json_({ ok: false, error: '비밀번호가 틀렸어요' });
      for (let i = values.length - 1; i >= 1; i--) {
        if (str_(values[i][0]) === round) sheet.deleteRow(i + 1);
      }
      return json_({ ok: true });
    }

    if (!name) return json_({ ok: false, error: '이름 없음' });

    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (str_(values[i][0]) === round && str_(values[i][1]) === name) { rowIndex = i; break; }
    }

    if (d.action === 'cancel') {
      if (rowIndex > 0) sheet.deleteRow(rowIndex + 1);
      return json_({ ok: true });
    }

    const menu = str_(d.menu).slice(0, 40);
    if (!menu) return json_({ ok: false, error: '메뉴 없음' });
    const row = [round, name, menu, str_(d.temp).slice(0, 10), str_(d.memo).slice(0, 60), new Date()];

    // 같은 사람이 다시 주문하면 덮어쓰기 (1인 1잔)
    if (rowIndex > 0) {
      sheet.getRange(rowIndex + 1, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 폐수배출시설 운영일지 - Google Apps Script
 *
 * 설정:
 * 1. 폐수 스프레드시트 열기 (이미 생성됨)
 * 2. 확장 프로그램 → Apps Script
 * 3. 이 코드 전체 붙여넣기
 * 4. 배포 → 새 배포 → 웹 앱
 *    - 실행 계정: 본인
 *    - 액세스: 모든 사용자
 * 5. URL 복사 → HTML 생성기에 한 번만 입력
 */

const SHEET_NAME = '폐수배출';
const HEADER = ['날짜','요일','쉬는날','1호기용수','2호기용수','1호기계량기','2호기계량기','필터교체','필터교체량','위탁량','확인서번호','처리업소명'];

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADER);
    return jsonRes({ success: true, data: [] });
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return jsonRes({ success: true, data: [] });

  const header = data[0].map(h => String(h).trim());
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    header.forEach((h, j) => {
      let v = data[i][j];
      if (v instanceof Date) {
        v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      obj[h] = v != null ? String(v).trim() : '';
    });
    if (obj['날짜']) rows.push(obj);
  }
  return jsonRes({ success: true, data: rows });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADER);
    }

    const rows = body.rows || [];
    const existing = sheet.getDataRange().getValues();

    // 날짜 정규화 함수
    function norm(ds) {
      ds = String(ds).trim();
      if (ds.indexOf('-') > -1) {
        var p = ds.split('-');
        return p[0].slice(-2) + '/' + ('0'+p[1]).slice(-2) + '/' + ('0'+p[2]).slice(-2);
      }
      if (ds.indexOf('/') > -1) {
        var p = ds.split('/');
        return p[0] + '/' + ('0'+p[1]).slice(-2) + '/' + ('0'+p[2]).slice(-2);
      }
      return ds;
    }

    // 기존 날짜 → 행번호 매핑
    var dateMap = {};
    for (var i = 1; i < existing.length; i++) {
      if (existing[i][0]) {
        dateMap[norm(existing[i][0])] = i + 1;
      }
    }

    var updated = 0, added = 0;
    for (var ri = 0; ri < rows.length; ri++) {
      var row = rows[ri];
      var dateStr = row['날짜'] || '';
      if (!dateStr) continue;
      var n = norm(dateStr);
      var values = HEADER.map(function(h) { return row[h] || ''; });

      if (dateMap[n]) {
        sheet.getRange(dateMap[n], 1, 1, HEADER.length).setValues([values]);
        updated++;
      } else {
        sheet.appendRow(values);
        dateMap[n] = existing.length + added + 1;
        added++;
      }
    }
    return jsonRes({ success: true, updated: updated, added: added });
  } catch (err) {
    return jsonRes({ success: false, error: err.message });
  }
}

function jsonRes(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * SOLTRI 피복 관리 — 레거시 백필 핸들러
 *
 * Apps Script 편집기(https://script.google.com/)에서:
 *  1. 기존 프로젝트 열기
 *  2. 이 파일 내용을 기존 코드 파일(Code.gs 등) 뒤에 append
 *  3. doPost() 안의 switch(action) 분기에 distributeBackfillBatch / updateEmpStatus 케이스 추가
 *  4. 새 버전으로 배포(Deploy → Manage deployments → Edit → New version → Deploy)
 *
 * 두 개의 액션 추가:
 *   - distributeBackfillBatch: 과거 지급 이력 일괄 기록 (재고 차감 없음)
 *   - updateEmpStatus: 직원 재직/퇴사 상태 변경
 */

// ═══════════════════════════════════════════════════════
// 설정: 시트 이름·컬럼 위치를 실제 구조에 맞게 조정하세요
// ═══════════════════════════════════════════════════════
var BACKFILL_CONFIG = {
  // 분배이력 시트 이름 (기존 distribute 액션이 쓰는 시트와 동일)
  historySheetName: '분배이력',

  // 기존 분배이력 컬럼 순서 (실제 시트 헤더에 맞게 수정)
  // 예: [날짜, 시간, 레코드ID, 사번, 이름, 부서, 품목ID, 품목명, 사이즈, 수량, 구분, 비고]
  historyColumns: ['date', 'time', 'recordId', 'empId', 'name', 'dept',
                   'itemId', 'itemName', 'size', 'qty', 'type', 'note'],

  // 직원 시트
  empSheetName: '직원',
  empKeyCol: 'empId',        // 조회 키 컬럼
  empNameCol: 'name',         // 이름 컬럼 (empId 없을 때 fallback)
  empStatusCol: 'status',     // 상태 컬럼 (신설 필요 - active/retired)
  empRetiredDateCol: 'retiredDate',  // 퇴사일 컬럼 (신설 필요, 선택)
};


// ═══════════════════════════════════════════════════════
// distributeBackfillBatch — 과거 지급 일괄 기록
// ═══════════════════════════════════════════════════════
/**
 * 요청 포맷:
 *   {
 *     action: 'distributeBackfillBatch',
 *     records: [
 *       { id, empName, dept, date, time, items:[{name,size,qty}], note, retired }
 *     ]
 *   }
 *
 * 응답:
 *   { ok: true, success: [id...], failed: [{id, error}...] }
 */
function distributeBackfillBatch(data) {
  var records = data.records || [];
  if (!records.length) return {ok: false, error: 'no records'};

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var histSheet = ss.getSheetByName(BACKFILL_CONFIG.historySheetName);
  if (!histSheet) return {ok: false, error: 'history sheet not found: ' + BACKFILL_CONFIG.historySheetName};

  // 직원 이름 → empId 매핑 (한 번만 로드)
  var empMap = _loadEmpMap(ss);

  // 품목 이름+사이즈 → itemId 매핑
  var itemMap = _loadItemMap(ss);

  var success = [];
  var failed = [];
  var rowsToAppend = [];

  for (var i = 0; i < records.length; i++) {
    var r = records[i];
    try {
      var empId = empMap[_normName(r.empName)] || '';
      // 각 item마다 한 행씩 append
      (r.items || []).forEach(function(it) {
        var itemKey = it.name + '|' + (it.size || '');
        var itemId = itemMap[itemKey] || '';
        var row = BACKFILL_CONFIG.historyColumns.map(function(col) {
          switch (col) {
            case 'date': return r.date;
            case 'time': return r.time || '00:00';
            case 'recordId': return r.id;
            case 'empId': return empId;
            case 'name': return r.empName;
            case 'dept': return r.dept || '';
            case 'itemId': return itemId;
            case 'itemName': return it.name;
            case 'size': return it.size || '';
            case 'qty': return it.qty || 0;
            case 'type': return 'backfill';
            case 'note': return r.note || (r.retired ? 'legacy(퇴사자)' : 'legacy');
            default: return '';
          }
        });
        rowsToAppend.push(row);
      });
      success.push(r.id);
    } catch (e) {
      failed.push({id: r.id, error: String(e.message || e)});
    }
  }

  // 배치 append (성능: 한 번에 기록)
  if (rowsToAppend.length > 0) {
    var lastRow = histSheet.getLastRow();
    histSheet
      .getRange(lastRow + 1, 1, rowsToAppend.length, BACKFILL_CONFIG.historyColumns.length)
      .setValues(rowsToAppend);
  }

  return {ok: true, success: success, failed: failed, appended: rowsToAppend.length};
}


// ═══════════════════════════════════════════════════════
// updateEmpStatus — 직원 재직/퇴사 상태 변경
// ═══════════════════════════════════════════════════════
/**
 * 요청:
 *   { action: 'updateEmpStatus', empId: '...', status: 'active' | 'retired', retiredDate: 'YYYY-MM-DD' }
 *
 * 응답:
 *   { ok: true } | { ok: false, error: '...' }
 */
function updateEmpStatus(data) {
  var empId = data.empId;
  var empName = data.empName;  // empId 없으면 이름으로 찾기
  var status = data.status;    // 'active' | 'retired'
  var retiredDate = data.retiredDate || '';

  if (!status) return {ok: false, error: 'status required'};

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var empSheet = ss.getSheetByName(BACKFILL_CONFIG.empSheetName);
  if (!empSheet) return {ok: false, error: 'emp sheet not found'};

  var headers = empSheet.getRange(1, 1, 1, empSheet.getLastColumn()).getValues()[0];
  var idCol = headers.indexOf(BACKFILL_CONFIG.empKeyCol);
  var nameCol = headers.indexOf(BACKFILL_CONFIG.empNameCol);
  var statusCol = headers.indexOf(BACKFILL_CONFIG.empStatusCol);
  var retDateCol = headers.indexOf(BACKFILL_CONFIG.empRetiredDateCol);

  if (statusCol < 0) {
    // status 컬럼이 없으면 헤더에 추가
    statusCol = headers.length;
    empSheet.getRange(1, statusCol + 1).setValue(BACKFILL_CONFIG.empStatusCol);
  }

  var data2 = empSheet.getDataRange().getValues();
  for (var i = 1; i < data2.length; i++) {
    var match = (empId && idCol >= 0 && data2[i][idCol] === empId) ||
                (empName && nameCol >= 0 && data2[i][nameCol] === empName);
    if (match) {
      empSheet.getRange(i + 1, statusCol + 1).setValue(status);
      if (retDateCol >= 0 && status === 'retired') {
        empSheet.getRange(i + 1, retDateCol + 1).setValue(retiredDate);
      }
      return {ok: true, updated: data2[i][nameCol] || empId};
    }
  }
  return {ok: false, error: 'employee not found'};
}


// ═══════════════════════════════════════════════════════
// 헬퍼
// ═══════════════════════════════════════════════════════
function _loadEmpMap(ss) {
  var sheet = ss.getSheetByName(BACKFILL_CONFIG.empSheetName);
  if (!sheet) return {};
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return {};
  var headers = values[0];
  var idCol = headers.indexOf(BACKFILL_CONFIG.empKeyCol);
  var nameCol = headers.indexOf(BACKFILL_CONFIG.empNameCol);
  if (nameCol < 0) return {};
  var map = {};
  for (var i = 1; i < values.length; i++) {
    var name = _normName(values[i][nameCol]);
    if (name) map[name] = idCol >= 0 ? values[i][idCol] : '';
  }
  return map;
}

function _loadItemMap(ss) {
  var sheet = ss.getSheetByName('품목');
  if (!sheet) return {};
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return {};
  var headers = values[0];
  var idCol = headers.indexOf('itemId');
  var nameCol = headers.indexOf('name');
  var sizeCol = headers.indexOf('size');
  if (nameCol < 0) return {};
  var map = {};
  for (var i = 1; i < values.length; i++) {
    var name = values[i][nameCol];
    var size = sizeCol >= 0 ? (values[i][sizeCol] || '') : '';
    var key = name + '|' + (size === 'FREE' ? '' : size);
    if (name) map[key] = idCol >= 0 ? values[i][idCol] : '';
  }
  return map;
}

function _normName(s) {
  return (s || '').toString().trim();
}


// ═══════════════════════════════════════════════════════
// doPost() 통합 — 기존 코드의 switch문에 아래 두 case를 추가하세요
// ═══════════════════════════════════════════════════════
/*
   doPost(e) 안의 switch (action) { ... } 블록에 다음을 추가:

      case 'distributeBackfillBatch':
        result = distributeBackfillBatch(data);
        break;

      case 'updateEmpStatus':
        result = updateEmpStatus(data);
        break;

   그리고 배포:
      Deploy → Manage deployments → (활성 배포 선택) → Edit(연필 아이콘)
      → Version: New version → Description: "레거시 백필 핸들러 추가"
      → Deploy

   배포 URL은 동일하게 유지됩니다.
*/

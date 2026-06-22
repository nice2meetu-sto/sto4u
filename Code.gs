/**
 * 서울관광재단 인력현황 관리 시스템 — Google Apps Script (GET 전용 JSON API)
 *
 * 역할: 같은 스프레드시트의 3개 탭(기초명부 / 인사발령 / 정원)을 읽어
 *       index.html 이 fetch 로 호출하는 GET JSON API 를 제공한다.
 *
 * 배포: [배포] → [새 배포] → 유형 "웹 앱"
 *       - 실행 계정: 나(본인)
 *       - 액세스 권한: 모든 사용자 (Anyone)
 *       그 후 발급된 웹앱 URL 을 index.html 상단 SCRIPT_URL 에 붙여넣는다.
 *
 * 제약: POST 사용 안 함(CORS preflight 회피). 모든 호출은 GET + 쿼리스트링.
 *       날짜는 ISO 문자열(yyyy-MM-dd)로 직렬화해 반환(클라이언트 파싱 일관성).
 */

// 탭 이름 (스프레드시트 탭명과 일치해야 함)
var SHEET_BASE   = "기초명부";
var SHEET_ORDERS = "인사발령";
var SHEET_QUOTA  = "정원";
var SHEET_DEPT   = "부서마스터";

// 타임존 (날짜 직렬화 기준)
var TZ = "Asia/Seoul";

/**
 * 진입점.
 *   ?action=data      → { base:[...], orders:[...], quota:[...] } 반환 (기본값)
 *   ?action=addOrder  → 인사발령 탭에 1행 append (2차 범위)
 */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || "data";
  try {
    if (action === "data") {
      return jsonOut(getData());
    } else if (action === "addOrder") {
      return jsonOut(addOrder(e.parameter));
    }
    return jsonOut({ error: "unknown action: " + action });
  } catch (err) {
    return jsonOut({ error: String(err && err.message ? err.message : err) });
  }
}

/** 세 탭을 읽어 JSON 객체로 반환 */
function getData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    base:   readSheet(ss, SHEET_BASE),
    orders: readSheet(ss, SHEET_ORDERS),
    quota:  readSheet(ss, SHEET_QUOTA),
    dept:   readSheet(ss, SHEET_DEPT),
    syncedAt: Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ssXXX")
  };
}

/**
 * 한 시트를 [{헤더:값, ...}, ...] 배열로 변환.
 * 1행=헤더, 2행부터 데이터. Date 셀은 yyyy-MM-dd 문자열로 직렬화.
 * 완전 빈 행은 건너뛴다. (탭이 없으면 빈 배열)
 */
function readSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0].map(function (h) { return String(h).trim(); });

  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var obj = {};
    var hasValue = false;
    for (var c = 0; c < headers.length; c++) {
      var key = headers[c];
      if (!key) continue;
      var v = row[c];
      if (v instanceof Date) {
        v = Utilities.formatDate(v, TZ, "yyyy-MM-dd");
      } else if (typeof v === "string") {
        v = v.trim();
      }
      if (v !== "" && v !== null && v !== undefined) hasValue = true;
      obj[key] = v;
    }
    if (hasValue) out.push(obj);
  }
  return out;
}

/**
 * (2차 범위) 인사발령 탭에 1행 추가.
 * 쿼리스트링 파라미터 키를 헤더와 매칭해 append.
 * 예) ?action=addOrder&발령일=2024-08-01&사원코드=E010&변동구분=입사&...
 */
function addOrder(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_ORDERS);
  if (!sheet) return { ok: false, error: "시트 없음: " + SHEET_ORDERS };

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h).trim(); });

  var rowData = headers.map(function (h) {
    return (h && params[h] !== undefined) ? params[h] : "";
  });
  sheet.appendRow(rowData);
  return { ok: true, appended: rowData, row: sheet.getLastRow() };
}

/** ContentService 로 JSON 출력 */
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

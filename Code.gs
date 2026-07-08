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
var SHEET_EMAP   = "기타_사원코드";

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
    } else if (action === "addBase") {
      return jsonOut(addBase(e.parameter));     // 입사: 기초명부에 1행 추가
    } else if (action === "setLeave") {
      return jsonOut(setLeave(e.parameter));    // 퇴사: 기초명부의 퇴사일 수정
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
    emap:   readSheet(ss, SHEET_EMAP),
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

/**
 * (입사) 기초명부 탭에 1행 추가. 쿼리 파라미터 키를 헤더와 매칭해 append.
 * 예) ?action=addBase&입사일=2026-07-01&사원코드=E200&사원명=홍길동&부서코드=D202&...
 */
function addBase(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_BASE);
  if (!sheet) return { ok: false, error: "시트 없음: " + SHEET_BASE };
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h).trim(); });
  var rowData = headers.map(function (h) {
    return (h && params[h] !== undefined) ? params[h] : "";
  });
  sheet.appendRow(rowData);
  return { ok: true, appended: rowData, row: sheet.getLastRow() };
}

/**
 * (퇴사) 기초명부에서 사원코드(+선택적으로 입사일)로 행을 찾아 퇴사일을 기록.
 * 예) ?action=setLeave&사원코드=E200&입사일=2020-01-01&퇴사일=2026-07-31
 */
function setLeave(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_BASE);
  if (!sheet) return { ok: false, error: "시트 없음: " + SHEET_BASE };
  var code  = String(params["사원코드"] || "").trim();
  var leave = String(params["퇴사일"]   || "").trim();
  if (!code)  return { ok: false, error: "사원코드가 필요합니다." };
  if (!leave) return { ok: false, error: "퇴사일이 필요합니다." };

  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0].map(function (h) { return String(h).trim(); });
  function colOf(re) { for (var c = 0; c < headers.length; c++) { if (re.test(headers[c])) return c; } return -1; }
  var cCode  = colOf(/^(사원코드|사번)$/);
  var cLeave = colOf(/^퇴사일$/);
  if (cCode  < 0) return { ok: false, error: "기초명부에 '사원코드' 열이 없습니다." };
  if (cLeave < 0) return { ok: false, error: "기초명부에 '퇴사일' 열이 없습니다." };

  // 사원코드(사번) 기준으로만 찾는다. 상태(재직/입사예정)·입사일은 보지 않음.
  // 비교는 숫자·문자·공백 차이를 흡수하기 위해 문자열 trim 후 대조.
  var target = -1, matched = 0;
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][cCode]).trim() === code) { matched++; if (target < 0) target = r; }
  }
  if (target < 0) return { ok: false, error: "해당 사원코드 행을 찾지 못했습니다: " + code };
  sheet.getRange(target + 1, cLeave + 1).setValue(leave);
  return { ok: true, 사원코드: code, 퇴사일: leave, row: target + 1, matched: matched };
}

/** ContentService 로 JSON 출력 */
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

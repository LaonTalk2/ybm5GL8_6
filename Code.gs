/* ============================================================
   Lesson 8 Wrap Up 앱 - 반 전체 공유 저장소 (Google Apps Script)
   ------------------------------------------------------------
   사용법:
   1) 이 스크립트를 넣을 스프레드시트를 여세요.
      (스프레드시트가 없다면 새로 만들어도 됩니다. 시트 자체는
       스크립트가 알아서 "Store" 탭을 만들어 사용합니다.)
   2) 상단 메뉴 [확장 프로그램] > [Apps Script] 클릭
   3) 기본 코드(Code.gs)를 전부 지우고 이 파일 내용을 붙여넣기
   4) 저장 후 우측 상단 [배포] > [새 배포]
      - 유형: 웹 앱
      - 실행 계정: 나(본인)
      - 액세스 권한이 있는 사용자: 전체
      - [배포] 클릭 → 처음 한 번은 권한 승인 필요
   5) 나온 "웹 앱 URL" (…/exec 로 끝남)을 복사해서
      HTML 파일의 GAS_URL 자리에 붙여넣으세요.
   ============================================================ */

var SHEET_NAME = 'Store';

function getSheet_(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh){
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['key', 'value', 'updatedAt']);
    sh.getRange('B:B').setNumberFormat('@'); // value 열은 항상 텍스트로
  }
  return sh;
}

function findRow_(sh, key){
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return -1;
  var keys = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < keys.length; i++){
    if (String(keys[i][0]) === key) return i + 2; // 실제 시트 행 번호(1-index, 헤더 포함)
  }
  return -1;
}

function jsonOut_(obj){
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e){
  var action = e.parameter.action;
  var sh = getSheet_();

  if (action === 'get'){
    var key = e.parameter.key || '';
    var row = findRow_(sh, key);
    if (row === -1) return jsonOut_({ value: null });
    var value = sh.getRange(row, 2).getValue();
    return jsonOut_({ value: String(value) });
  }

  if (action === 'list'){
    var prefix = e.parameter.prefix || '';
    var lastRow = sh.getLastRow();
    var keys = [];
    if (lastRow >= 2){
      var data = sh.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < data.length; i++){
        var k = String(data[i][0]);
        if (k.indexOf(prefix) === 0) keys.push(k);
      }
    }
    return jsonOut_({ keys: keys });
  }

  return jsonOut_({ error: 'unknown action' });
}

function doPost(e){
  var action = e.parameter.action;
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err){ body = {}; }
  var sh = getSheet_();

  if (action === 'delete'){
    var delRow = findRow_(sh, body.key);
    if (delRow !== -1) sh.deleteRow(delRow);
    return jsonOut_({ ok: true });
  }

  // action === 'set' (기본값)
  var key = body.key;
  var value = (typeof body.value === 'string') ? body.value : JSON.stringify(body.value);
  var now = new Date().toISOString();
  var row = findRow_(sh, key);

  if (row === -1){
    sh.appendRow([key, value, now]);
    sh.getRange(sh.getLastRow(), 2).setNumberFormat('@').setValue(value);
  } else {
    sh.getRange(row, 2, 1, 2).setNumberFormat('@').setValues([[value, now]]);
  }
  return jsonOut_({ ok: true });
}

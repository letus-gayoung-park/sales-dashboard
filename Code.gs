// Google Apps Script - 2026 영업 대시보드 데이터 API
// 배포: 실행 > 웹 앱으로 배포 > 전체 공개(모든 사용자) > 새 버전으로 배포

function doGet(e) {
  try {
    // ── 1. RAW 시트 데이터 (파이프라인/수주율 기준) ──
    var rawSS = SpreadsheetApp.openById('1_h1USjyp3f0_angKLEyalu5P9j6FisdPK0rXMhClmHQ');
    var rawSheet = rawSS.getSheets()[0];
    var rawValues = rawSheet.getDataRange().getValues();
    var headers = rawValues[0];
    var rawData = [];
    for (var i = 1; i < rawValues.length; i++) {
      var row = rawValues[i];
      if (!row[1]) continue; // 수요처명 없는 행 스킵
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        if (headers[j]) obj[String(headers[j])] = row[j];
      }
      rawData.push(obj);
    }

    // ── 2. 매출시트 월별 합계 (정기청소 포함 확정 매출 기준) ──
    var salesSS = SpreadsheetApp.openById('1U0wMXkHHJMyeRSFKp4YePwFlr6NssIFMY_Wwo5Vp4Nw');
    var salesSheet = salesSS.getSheetByName('매출(26년)');
    var salesValues = salesSheet.getDataRange().getValues();
    var salesHeaders = salesValues[0];

    var monthIdx = -1, amountIdx = -1;
    for (var k = 0; k < salesHeaders.length; k++) {
      if (String(salesHeaders[k]).indexOf('실행월')  !== -1) monthIdx  = k;
      if (String(salesHeaders[k]).indexOf('매출금액') !== -1) amountIdx = k;
    }

    var MONTHS = ['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월'];
    var monthlySales = [0,0,0,0,0,0,0,0,0,0,0,0];

    for (var r = 1; r < salesValues.length; r++) {
      var sRow = salesValues[r];
      var month = String(sRow[monthIdx] || '').trim();
      var mIdx = MONTHS.indexOf(month);
      if (mIdx === -1) continue; // 소계/합계 행 제외
      var amount = parseFloat(String(sRow[amountIdx] || '0').replace(/[^0-9.]/g, '')) || 0;
      monthlySales[mIdx] += amount;
    }

    var result = {
      rawData: rawData,
      monthlySales: monthlySales
    };

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

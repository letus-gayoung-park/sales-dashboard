// Google Apps Script - 2026 영업 대시보드 데이터 API
// 배포: 실행 > 웹 앱으로 배포 > 전체 공개(모든 사용자) > 새 버전으로 배포

// ── 매월 1일 자동 실행: 정기청소 합계를 RAW 시트에 추가 ──
// 트리거 설정: Apps Script 편집기 > 왼쪽 시계 아이콘(트리거) > 추가
//   함수: addRegularMonthly / 이벤트: 시간 기반 / 월별 타이머 / 매월 1일 오전 9시
function addRegularMonthly() {
  var SALES_ID = '1U0wMXkHHJMyeRSFKp4YePwFlr6NssIFMY_Wwo5Vp4Nw';
  var RAW_ID   = '1_h1USjyp3f0_angKLEyalu5P9j6FisdPK0rXMhClmHQ';

  var MONTH_LAST_DAY = {
    '1월':'2026-01-31','2월':'2026-02-28','3월':'2026-03-31',
    '4월':'2026-04-30','5월':'2026-05-31','6월':'2026-06-30',
    '7월':'2026-07-31','8월':'2026-08-31','9월':'2026-09-30',
    '10월':'2026-10-31','11월':'2026-11-30','12월':'2026-12-31'
  };

  // 1. 매출시트 정기 항목 월별 합계
  var salesSheet = SpreadsheetApp.openById(SALES_ID).getSheetByName('매출(26년)');
  var salesData  = salesSheet.getDataRange().getValues();
  var sh = salesData[0];
  var monthIdx = -1, gubunIdx = -1, amountIdx = -1;
  for (var k = 0; k < sh.length; k++) {
    if (String(sh[k]).indexOf('실행월')   !== -1) monthIdx  = k;
    if (String(sh[k]).indexOf('구분')     !== -1) gubunIdx  = k;
    if (String(sh[k]).indexOf('매출금액') !== -1) amountIdx = k;
  }
  var monthlyRegular = {};
  for (var i = 1; i < salesData.length; i++) {
    var r = salesData[i];
    var m = String(r[monthIdx]  || '').trim();
    var g = String(r[gubunIdx]  || '').trim();
    if (!MONTH_LAST_DAY[m] || g !== '정기') continue;
    monthlyRegular[m] = (monthlyRegular[m] || 0) +
      (parseFloat(String(r[amountIdx] || '0').replace(/[^0-9.]/g, '')) || 0);
  }

  // 2. RAW에서 이미 추가된 정기청소 행 확인
  var rawSheet  = SpreadsheetApp.openById(RAW_ID).getSheets()[0];
  var rawData   = rawSheet.getDataRange().getValues();
  var rh = rawData[0];
  var nameIdx = -1, svcIdx = -1, winIdx = -1, execIdx = -1, priceIdx = -1;
  for (var k = 0; k < rh.length; k++) {
    if (String(rh[k]).indexOf('수요처명')        !== -1) nameIdx  = k;
    if (String(rh[k]).indexOf('서비스 유형(대)') !== -1) svcIdx   = k;
    if (String(rh[k]).indexOf('수주여부')        !== -1) winIdx   = k;
    if (String(rh[k]).indexOf('서비스 실행 일자') !== -1) execIdx  = k;
    if (String(rh[k]).indexOf('견적가')          !== -1) priceIdx = k;
  }
  var alreadyAdded = {};
  for (var i = 1; i < rawData.length; i++) {
    var name = String(rawData[i][nameIdx] || '').trim();
    if (name.indexOf('정기청소_') === 0 && name.indexOf('합계') !== -1)
      alreadyAdded[name.replace('정기청소_','').replace('합계','')] = true;
  }

  // 3. 이번 달 이전 + 아직 미추가 월만 대상
  var prevMonth = new Date(); prevMonth.setDate(0); // 전달 마지막 날
  var prevMonthNum = prevMonth.getMonth() + 1; // 1~12

  var toAdd = Object.keys(monthlyRegular).filter(function(m) {
    var mNum = parseInt(m.replace('월',''));
    return !alreadyAdded[m] && mNum <= prevMonthNum;
  }).sort(function(a, b) {
    return parseInt(a.replace('월','')) - parseInt(b.replace('월',''));
  });

  if (toAdd.length === 0) { Logger.log('추가할 정기청소 행 없음'); return; }

  // 4. RAW 맨 아래에 행 추가
  var numCols = rh.length;
  var newRows = toAdd.map(function(month) {
    var row = new Array(numCols).fill('');
    row[nameIdx]  = '정기청소_' + month + '합계';
    row[svcIdx]   = '클리닝';
    row[winIdx]   = 'O';
    row[execIdx]  = MONTH_LAST_DAY[month];
    row[priceIdx] = monthlyRegular[month];
    return row;
  });

  rawSheet.getRange(rawSheet.getLastRow() + 1, 1, newRows.length, numCols).setValues(newRows);
  Logger.log(toAdd.length + '개 월 추가: ' + toAdd.join(', '));
}

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

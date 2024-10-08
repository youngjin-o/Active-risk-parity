function updateStockData() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dashboardSheet = spreadsheet.getSheetByName("대시보드");
  var alertSheet = spreadsheet.getSheetByName("알림기록");
  var investmentRecordSheet = spreadsheet.getSheetByName("투자기록");
  var volatilityAdjustmentSheet = spreadsheet.getSheetByName("변동성-레버리지조정");

  if (!dashboardSheet || !alertSheet || !investmentRecordSheet || !volatilityAdjustmentSheet) {
    throw new Error("필요한 시트를 찾을 수 없습니다.");
  }

  var lastColumn = dashboardSheet.getLastColumn();
  var lastRow = dashboardSheet.getLastRow();
  
  var headerRow = 2;  // 티커가 있는 행
  var tickers = dashboardSheet.getRange(headerRow, 2, 1, lastColumn - 1).getValues()[0];
  
  var rowLabels = dashboardSheet.getRange(headerRow, 1, lastRow - headerRow + 1, 1).getValues().map(row => row[0]);

  var currentPriceRow = rowLabels.indexOf("현재가격") + headerRow;
  var highPriceRow = rowLabels.indexOf("고점") + headerRow;
  var highPriceUpdateDateRow = rowLabels.indexOf("고점갱신일") + headerRow;
  var currentToPeakRatioRow = rowLabels.indexOf("고점대비 현재가 비율") + headerRow;
  var decline1Row = rowLabels.indexOf("하락률 1단계") + headerRow;
  var decline2Row = rowLabels.indexOf("하락률 2단계") + headerRow;
  var decline3Row = rowLabels.indexOf("하락률 3단계") + headerRow;
  var decline1DateRow = rowLabels.indexOf("1단계 도달 날짜") + headerRow;
  var decline2DateRow = rowLabels.indexOf("2단계 도달 날짜") + headerRow;
  var decline3DateRow = rowLabels.indexOf("3단계 도달 날짜") + headerRow;

  var today = new Date();
  var dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");

  var isFirstDayOfMonth = today.getDate() === 1;
  var events = []; // 발생한 모든 이벤트를 저장할 배열  
  var eventOccurred = false;
  var eventTicker = null;
  var eventType = null;

  Logger.log("오늘 날짜: " + dateString);
  Logger.log("월 첫 날 여부: " + isFirstDayOfMonth);

  var currentPrices = [];
  var highPrices = [];
  var declineRatios = [];
 
  for (var col = 2; col <= lastColumn; col++) {
    var ticker = tickers[col - 2];
    if (!ticker) continue;

    var currentPrice;
    if (ticker === "전국아파트실거래가") {
      var realEstateData = getRealEstatePrice();
      if (realEstateData) {
        currentPrice = realEstateData.value;
        dashboardSheet.getRange(currentPriceRow, col).setValue(currentPrice);
      } else {
        Logger.log("부동산 가격을 받아오지 못했습니다.");
        continue;
      }
    } else {
      currentPrice = getStockPrice(dashboardSheet, currentPriceRow, col, ticker);
    }

    if (currentPrice === null) {
      // 가격을 받아오지 못했을 때의 처리
      Logger.log(`${ticker}의 가격을 받아오지 못했습니다. 이 종목은 건너뜁니다.`);
      continue;
    }
    var highPrice = dashboardSheet.getRange(highPriceRow, col).getValue();
    var decline1 = dashboardSheet.getRange(decline1Row, col).getValue();
    var decline2 = dashboardSheet.getRange(decline2Row, col).getValue();
    var decline3 = dashboardSheet.getRange(decline3Row, col).getValue();

    Logger.log("티커: " + ticker + ", 현재 가격: " + currentPrice + ", 고점: " + highPrice);

    currentPrices.push(currentPrice);
    highPrices.push(highPrice);

    if (isNaN(currentPrice)) continue;

    // 고점이 설정되지 않았다면 현재 가격으로 초기화합니다
    if (!highPrice || isNaN(highPrice)) {
      highPrice = currentPrice;
      dashboardSheet.getRange(highPriceRow, col).setValue(highPrice);
      dashboardSheet.getRange(highPriceUpdateDateRow, col).setValue(dateString);
    }

    // 고점 업데이트 로직
    if (currentPrice > highPrice) {
      highPrice = currentPrice;
      dashboardSheet.getRange(highPriceRow, col).setValue(highPrice);
      dashboardSheet.getRange(highPriceUpdateDateRow, col).setValue(dateString);
      
      // 1단계 하락 날짜가 있는지 확인
      var decline1Date = dashboardSheet.getRange(decline1DateRow, col).getValue();
      if (decline1Date) {
        // 1~3단계 하락 날짜를 모두 지웁니다
        dashboardSheet.getRange(decline1DateRow, col).clearContent();
        dashboardSheet.getRange(decline2DateRow, col).clearContent();
        dashboardSheet.getRange(decline3DateRow, col).clearContent();
        
        recordAlert(alertSheet, ticker, "고점 갱신", dateString, currentPrice, null, highPrice);
        sendEmail(ticker, "고점 갱신", dateString, currentPrice, highPrice);
      }
    }

    // 하락 비율 계산 (마이너스 값으로 표현)
    var declineRatio = currentPrice / highPrice - 1;
    declineRatios.push(declineRatio);
    dashboardSheet.getRange(currentToPeakRatioRow, col).setValue(declineRatio);

    // 하락 단계 확인
    var stageResult = checkDeclineStage(dashboardSheet, alertSheet, col, declineRatio, decline1, decline2, decline3, 
                                        decline1DateRow, decline2DateRow, decline3DateRow, 
                                        dateString, ticker, currentPrice, highPrice);
    
    if (stageResult.eventOccurred) {
      events.push({
        ticker: ticker,
        type: stageResult.eventType
      });               
    }
  }

  // 매월 첫날이거나 이벤트가 발생했을 때 투자 기록 업데이트
  if (isFirstDayOfMonth) {
    events.push({ type: "매달기록" });
  }

  if (events.length > 0) {
    Logger.log("발생한 이벤트: " + JSON.stringify(events));

    var exchangeRate = getExchangeRate();
    if (exchangeRate === null) {
      Logger.log("환율을 받아오지 못했습니다. 스크립트를 종료합니다.");
      return;
    }

    var tickerClassifications = getTickerClassifications();

    // 각 이벤트에 대해 순차적으로 투자 기록 업데이트
    for (var i = 0; i < events.length; i++) {
      var event = events[i];
      Logger.log(`이벤트 처리 중 (${i + 1}/${events.length}): ${event.type} ${event.ticker || ''}`);

      updateInvestmentRecord(
        investmentRecordSheet, 
        volatilityAdjustmentSheet, 
        tickers, 
        currentPrices, 
        highPrices, 
        declineRatios, 
        event.type, 
        event.ticker, 
        exchangeRate, 
        tickerClassifications
      );

      // 다음 이벤트 처리 전 잠시 대기 (선택사항)
      if (i < events.length - 1) {
        Utilities.sleep(1000); // 1초 대기
      }
    }
  } else {
    Logger.log("투자 기록 업데이트 조건 미충족");
  }
}

function checkDeclineStage(sheet, alertSheet, col, declineRatio, decline1, decline2, decline3, 
                           decline1DateRow, decline2DateRow, decline3DateRow, 
                           dateString, ticker, currentPrice, highPrice) {
  var eventOccurred = false;
  var eventType = null;

  if (declineRatio <= decline3 && !sheet.getRange(decline3DateRow, col).getValue()) {
    sheet.getRange(decline3DateRow, col).setValue(dateString);
    recordAlert(alertSheet, ticker, "3단계 하락", dateString, currentPrice, declineRatio, highPrice);
    sendEmail(ticker, "3단계 하락", dateString, currentPrice, highPrice, declineRatio);
    eventOccurred = true;
    eventType = "3단계 하락";
  } else if (declineRatio <= decline2 && !sheet.getRange(decline2DateRow, col).getValue()) {
    sheet.getRange(decline2DateRow, col).setValue(dateString);
    recordAlert(alertSheet, ticker, "2단계 하락", dateString, currentPrice, declineRatio, highPrice);
    sendEmail(ticker, "2단계 하락", dateString, currentPrice, highPrice, declineRatio);
    eventOccurred = true;
    eventType = "2단계 하락";
  } else if (declineRatio <= decline1 && !sheet.getRange(decline1DateRow, col).getValue()) {
    sheet.getRange(decline1DateRow, col).setValue(dateString);
    recordAlert(alertSheet, ticker, "1단계 하락", dateString, currentPrice, declineRatio, highPrice);
    sendEmail(ticker, "1단계 하락", dateString, currentPrice, highPrice, declineRatio);
    eventOccurred = true;
    eventType = "1단계 하락";
  }

  return { eventOccurred: eventOccurred, eventType: eventType };
}

function getNewRatioFromVolatilitySheet(sheet, ticker, event) {
  // 헤더 행 찾기
  var headerRow = sheet.getRange("A:A").getValues().flat().indexOf("티커") + 1;
  
  // 티커 열 찾기
  var tickerColumn = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf(ticker) + 1;
  
  if (tickerColumn === 0) {
    throw new Error(`티커 ${ticker}를 찾을 수 없습니다.`);
  }
  
  // 이벤트에 따른 행 결정
  var ratioRow;
  switch (event) {
    case "1단계 하락":
      ratioRow = sheet.getRange("A:A").getValues().flat().indexOf("진입비율 1단계") + 1;
      break;
    case "2단계 하락":
      ratioRow = sheet.getRange("A:A").getValues().flat().indexOf("진입비율 2단계") + 1;
      break;
    case "3단계 하락":
      ratioRow = sheet.getRange("A:A").getValues().flat().indexOf("진입비율 3단계") + 1;
      break;
    default:
      throw new Error(`알 수 없는 이벤트: ${event}`);
  }
  
  // 진입비율 가져오기
  var entryRatio = sheet.getRange(ratioRow, tickerColumn).getValue();
  
  // 레버리지비율 가져오기
  var leverageRow = sheet.getRange("A:A").getValues().flat().indexOf("레버리지 비율") + 1;
  var leverageRatio = sheet.getRange(leverageRow, tickerColumn).getValue();
  
  // 자산배수 찾기
  var assetMultiplierRow = sheet.getRange("A:A").getValues().flat().indexOf("자산배수") + 1;
  var assetMultiplier = sheet.getRange(assetMultiplierRow, 2).getValue();
  
  // 새로운 비율 계산
  var newRatio = entryRatio * leverageRatio * assetMultiplier;
  
  // 백분율을 소수점으로 변환 (예: 5% -> 0.05)
  return newRatio / 100;
}

function recordAlert(alertSheet, ticker, event, date, currentPrice, relevantValue, highPrice) {
  var headers = alertSheet.getRange(2, 1, 1, alertSheet.getLastColumn()).getValues()[0];
  var tickerColumn = headers.indexOf(ticker);
  
  if (tickerColumn === -1) {
    throw new Error("알림 기록 시트에서 티커 " + ticker + "를 찾을 수 없습니다.");
  }
  
  var dateColumn = tickerColumn+1;
  var eventColumn = tickerColumn + 2;
  
  var lastRow = alertSheet.getLastRow();
  var newRow = 3; // 데이터는 3행부터 시작

  // 해당 티커의 마지막 행을 찾습니다
  for (var i = 3; i <= lastRow; i++) {
    if (alertSheet.getRange(i, dateColumn).getValue() === "") {
      newRow = i;
      break;
    }
    newRow = i + 1;
  }
  
  var eventMessage = event + " (" + currentPrice + "), 고점: " + highPrice;
  if (event !== "고점 갱신") {
    eventMessage += " - " + (relevantValue * 100).toFixed(2) + "%";
  }
  
  alertSheet.getRange(newRow, dateColumn).setValue(date);
  alertSheet.getRange(newRow, eventColumn).setValue(eventMessage);
}

function sendEmail(ticker, event, date, currentPrice, highPrice, declineRatio) {
  var recipient = "dhdudwls66@gmail.com";  // 받는 사람의 이메일 주소로 변경하세요
  var subject = ticker + " - " + event + " 알림";
  var body = "종목: " + ticker + "\n" +
             "이벤트: " + event + "\n" +
             "날짜: " + date + "\n" +
             "현재가: " + currentPrice + "\n" +
             "고점: " + highPrice + "\n";
  
  if (event === "고점 갱신") {
    body += "새로운 고점: " + highPrice;
  } else {
    body += "하락률: " + (declineRatio * 100).toFixed(2) + "%";
  }

  MailApp.sendEmail(recipient, subject, body);
}

function createDailyTrigger() {
  ScriptApp.newTrigger('updateStockData')
      .timeBased()
      .everyDays(1)
      .atHour(10)
      .create();
}

function updateInvestmentRecord(investmentRecordSheet, volatilityAdjustmentSheet, tickers, currentPrices, highPrices, declineRatios, event, triggerTicker, exchangeRate, tickerClassifications) {
  try {
    var lastRow = investmentRecordSheet.getLastRow();
    var newRowNumber = lastRow + 1;
    var today = new Date();
    var dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");

    // 부동산 데이터 제외
    var filteredTickers = [];
    var filteredCurrentPrices = [];
    var filteredHighPrices = [];
    var filteredDeclineRatios = [];
    var filteredTickerClassifications = {};

    for (var i = 0; i < tickers.length; i++) {
      if (tickers[i] !== "전국아파트실거래가") {
        filteredTickers.push(tickers[i]);
        filteredCurrentPrices.push(currentPrices[i]);
        if (highPrices && highPrices[i]) filteredHighPrices.push(highPrices[i]);
        if (declineRatios && declineRatios[i]) filteredDeclineRatios.push(declineRatios[i]);
        filteredTickerClassifications[tickers[i]] = tickerClassifications[tickers[i]];
      }
    }

    // 이전 행의 데이터 가져오기
    var previousRowData = getPreviousRowData(investmentRecordSheet, lastRow, filteredTickers.length);

    var newRowData;
    if (event === "매달기록") {
      newRowData = calculateMonthlyRecord(dateString, previousRowData, filteredTickers, filteredCurrentPrices, newRowNumber, exchangeRate, filteredTickerClassifications);
    } else {
      newRowData = calculateDeclineEvent(dateString, previousRowData, filteredTickers, filteredCurrentPrices, event, triggerTicker, volatilityAdjustmentSheet, newRowNumber, exchangeRate, filteredTickerClassifications);
    }

    // 새로운 행 추가
    investmentRecordSheet.getRange(newRowNumber, 1, 1, newRowData.length).setValues([newRowData]);
    Logger.log("투자 기록이 성공적으로 업데이트되었습니다.");
  } catch (error) {
    Logger.log("투자 기록 업데이트 중 오류 발생: " + error.message);
    sendErrorEmail("투자 기록 업데이트 오류", error.message);
  }
}

function sendErrorEmail(subject, errorMessage) {
  var recipient = "dhdudwls66@gmail.com";  // 오류 알림을 받을 이메일 주소
  var body = "투자 기록 업데이트 중 다음과 같은 오류가 발생했습니다:\n\n" + errorMessage;
  MailApp.sendEmail(recipient, subject, body);
}

function calculateMonthlyRecord(dateString, previousRowData, tickers, currentPrices, newRowNumber, exchangeRate, tickerClassifications) {
  Logger.log("calculateMonthlyRecord 시작");
  Logger.log("이전 데이터: " + JSON.stringify(previousRowData));
  Logger.log("현재 가격: " + JSON.stringify(currentPrices));
  Logger.log("환율: " + exchangeRate);
  Logger.log("티커 분류: " + JSON.stringify(tickerClassifications));

  var totalInvestment = previousRowData.totalInvestment;
  var currentTotalValue = 0;
  var tickerData = [];

  tickers.forEach((ticker, index) => {
    var classification = tickerClassifications[ticker];
    var originalPrice = currentPrices[index];
    var calculationPrice = calculatePriceInKRW(ticker, originalPrice, exchangeRate, classification);
    var shares = processShares(previousRowData.shares[index], classification);
    
    Logger.log(`${ticker}: 분류=${classification}, 원래 가격=${originalPrice}, 계산용 가격=${calculationPrice}, 주식 수=${shares}`);
    
    var value = calculationPrice * shares;
    currentTotalValue += value;
    
    // 비율 계산
    var ratio = (value / totalInvestment) * 100; // 퍼센트로 표시
    
    tickerData.push(originalPrice, shares, ratio.toFixed(2) + '%'); // 소수점 둘째자리까지 표시
  });

  var totalRatio = (currentTotalValue / totalInvestment) * 100;

  Logger.log("계산 결과: 총 투자금=" + totalInvestment + ", 현재 총 가치=" + currentTotalValue);
  Logger.log("티커 데이터: " + JSON.stringify(tickerData));

  return [dateString, totalInvestment, currentTotalValue, totalRatio.toFixed(2) + '%', "매달기록", exchangeRate]
    .concat(tickerData);
}

function calculateDeclineEvent(dateString, previousRowData, tickers, currentPrices, event, triggerTicker, volatilityAdjustmentSheet, newRowNumber, exchangeRate, tickerClassifications) {
  Logger.log("calculateDeclineEvent 시작");
  Logger.log("이전 데이터: " + JSON.stringify(previousRowData));
  Logger.log("현재 가격: " + JSON.stringify(currentPrices));
  Logger.log("환율: " + exchangeRate);
  Logger.log("티커 분류: " + JSON.stringify(tickerClassifications));
  
  var totalInvestment = previousRowData.totalInvestment;
  var currentTotalValue = 0;
  var tickerData = [];
  var adjustedRatios = {};
  
  // 변동성-레버리지조정 시트에서 새로운 비율 가져오기
  var newRatio = getNewRatioFromVolatilitySheet(volatilityAdjustmentSheet, triggerTicker, event);
  adjustedRatios[triggerTicker] = newRatio;
  Logger.log("새로운 비율: " + JSON.stringify(adjustedRatios));
  
  // 초기 ratio 계산
  var initialRatios = {};
  tickers.forEach((ticker, index) => {
    var classification = tickerClassifications[ticker];
    var originalPrice = currentPrices[index];
    var calculationPrice = calculatePriceInKRW(ticker, originalPrice, exchangeRate, classification);
    var shares = processShares(previousRowData.shares[index], classification);
    var currentValue = calculationPrice * shares;
    initialRatios[ticker] = currentValue / totalInvestment;
  });
  Logger.log("초기 비율: " + JSON.stringify(initialRatios));
  
  // 새로운 비율 합 계산
  var newTotalRatio = Object.values(initialRatios).reduce((sum, ratio) => sum + ratio, 0) - initialRatios[triggerTicker] + newRatio;
  
  // 초과 비율 계산 (1을 초과할 때만)
  var excessRatio = Math.max(0, newTotalRatio - 1);
  Logger.log("새로운 총 비율: " + newTotalRatio + ", 초과 비율: " + excessRatio);
  
  tickers.forEach((ticker, index) => {
    var classification = tickerClassifications[ticker];
    var originalPrice = currentPrices[index];
    var calculationPrice = calculatePriceInKRW(ticker, originalPrice, exchangeRate, classification);
    
    var ratio;
    if (ticker === triggerTicker) {
      ratio = adjustedRatios[ticker];
    } else if (excessRatio > 0) {
      // 다른 종목들의 비율 조정 로직 (초과 비율이 있을 때만)
      var reductionFactor = (excessRatio / (newTotalRatio - newRatio)) * initialRatios[ticker];
      ratio = initialRatios[ticker] - reductionFactor;
    } else {
      ratio = initialRatios[ticker];
    }
    
    // 새로운 주식 수 계산
    var newShares = processShares((ratio * totalInvestment) / calculationPrice, classification);
    var value = calculationPrice * newShares;
    currentTotalValue += value;
    
    // 비율을 퍼센트로 계산
    var assetRatio = (value / totalInvestment) * 100;
    tickerData.push(originalPrice, newShares, assetRatio.toFixed(2) + '%');
    Logger.log(ticker + " 데이터: 원래 가격=" + originalPrice + ", 계산용 가격=" + calculationPrice + ", 주식 수=" + newShares + ", 비율=" + assetRatio.toFixed(2) + '%');
  });
  
  var totalRatio = (currentTotalValue / totalInvestment) * 100;
  var eventName = event.includes("단계 하락") ? `${event} (${triggerTicker})` : event;
  Logger.log("계산 결과: 총 투자금=" + totalInvestment + ", 현재 총 가치=" + currentTotalValue + ", 총 비율=" + totalRatio.toFixed(2) + '%');
  
  return [dateString, totalInvestment, currentTotalValue, totalRatio.toFixed(2) + '%', eventName, exchangeRate]
    .concat(tickerData);
}

function getPreviousRowData(sheet, lastRow, tickerCount) {
  try {
    if (lastRow < 2) {
      throw new Error("이전 데이터가 없습니다.");
    }
    var previousRowData = sheet.getRange(lastRow, 1, 1, 6 + tickerCount * 3).getValues()[0];
    Logger.log("이전 행 데이터: " + JSON.stringify(previousRowData));
    
    if (previousRowData[1] === "" || isNaN(previousRowData[1])) {
      throw new Error("총 투자금 데이터가 유효하지 않습니다.");
    }

    return {
      totalInvestment: previousRowData[1],
      shares: previousRowData.slice(6).filter((_, i) => i % 3 === 1)
    };
  } catch (error) {
    Logger.log("이전 행 데이터 가져오기 중 오류 발생: " + error.message);
    // 기본값 반환 또는 오류 처리 로직 추가
    return {
      totalInvestment: 10000, // 기본 투자금 설정
      shares: new Array(tickerCount).fill(0)
    };
  }
}

function testInvestmentRecord() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var investmentRecordSheet = spreadsheet.getSheetByName("투자기록");
  var volatilityAdjustmentSheet = spreadsheet.getSheetByName("변동성-레버리지조정");
  
  var tickers = ["NASDAQ:GOOGL", "NASDAQ:META", "NYSE:BRK.B", "NASDAQ:ADBE", "INDEXNASDAQ:.IXIC", "KRX:005935", "KOSDAQ:058470", "KOSDAQ:074600", "KRX:097955", "KOSDAQ:140860", "KRX:298020", "KRX:KOSPI", "BTCUSD", "ETHUSD", "NYSEARCA:EDV"];
  var currentPrices = [157.36, 511.76, 476.83, 571.04, 17136.3, 57500, 183100, 25400, 137000, 175100, 288500, 2605.78, 56699.04, 2366.580385, 79.52];
  var event = "1단계 하락"; //매달기록 or 0단계 하락
  var triggerTicker = "BTCUSD"; //

  updateInvestmentRecord(investmentRecordSheet, volatilityAdjustmentSheet, tickers, currentPrices, [], [], event, triggerTicker);
  
  Logger.log("테스트 완료");
}

function getColumnName(index) {
  let columnName = '';
  while (index > 0) {
    index--;
    columnName = String.fromCharCode(65 + (index % 26)) + columnName;
    index = Math.floor(index / 26);
  }
  return columnName;
}

function getExchangeRate(maxRetries = 20, retryDelay = 2000) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dashboardSheet = spreadsheet.getSheetByName("대시보드");
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    var exchangeRate = dashboardSheet.getRange("B15").getValue();
    
    if (exchangeRate !== "#ERROR!" && !isNaN(exchangeRate) && exchangeRate !== "") {
      Logger.log(`환율 성공적으로 받아옴: ${exchangeRate} (시도 ${attempt}번째)`);
      return exchangeRate;
    }
    
    Logger.log(`환율 받아오기 실패 (시도 ${attempt}/${maxRetries}). 현재 값: ${exchangeRate}`);
    
    if (attempt < maxRetries) {
      Logger.log(`${retryDelay/1000}초 후 다시 시도합니다...`);
      Utilities.sleep(retryDelay);
    }
  }
  
  Logger.log(`환율을 ${maxRetries}번 시도 후에도 받아오지 못했습니다. 오류 처리가 필요합니다.`);
  return null;
}

function getTickerClassifications() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dashboardSheet = spreadsheet.getSheetByName("대시보드");
  
  var tickerRow = dashboardSheet.getRange("A:A").getValues().flat().indexOf("티커") + 1;
  var classificationRow = dashboardSheet.getRange("A:A").getValues().flat().indexOf("분류") + 1;
  
  var tickers = dashboardSheet.getRange(tickerRow, 2, 1, dashboardSheet.getLastColumn() - 1).getValues()[0];
  var classifications = dashboardSheet.getRange(classificationRow, 2, 1, dashboardSheet.getLastColumn() - 1).getValues()[0];
  
  var tickerClassifications = {};
  tickers.forEach((ticker, index) => {
    if (ticker) {
      tickerClassifications[ticker] = classifications[index];
    }
  });
  
  return tickerClassifications;
}

function calculatePriceInKRW(ticker, price, exchangeRate, classification) {
  Logger.log(`calculatePriceInKRW: ticker=${ticker}, price=${price}, exchangeRate=${exchangeRate}, classification=${classification}`);
  
  if (classification === "한국주식" || classification === "한국ETF" || classification === "한국아파트") {
    Logger.log(`${ticker}: 한국 주식/ETF, 원래 가격 반환`);
    return price;
  } else if (classification === "미국주식" || classification === "미국ETF" || classification === "코인" || classification === "미국채권") {
    Logger.log(`${ticker}: 미국 주식/ETF/코인/채권, 환율 적용`);
    return price * exchangeRate;
  } else {
    Logger.log(`${ticker}: 알 수 없는 분류, 기본적으로 환율 적용`);
    return price * exchangeRate;
  }
}

function processShares(shares, classification) {
  if (classification === "코인") {
    return parseFloat(parseFloat(shares).toFixed(4));
  } else {
    return Math.floor(parseFloat(shares));
  }
}

function getStockPrice(sheet, row, col, ticker, maxRetries = 20, retryDelay = 2000) {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    var price = sheet.getRange(row, col).getValue();
    
    if (price !== "#ERROR!" && !isNaN(price) && price !== "") {
      Logger.log(`${ticker} 가격 성공적으로 받아옴: ${price} (시도 ${attempt}번째)`);
      return price;
    }
    
    Logger.log(`${ticker} 가격 받아오기 실패 (시도 ${attempt}/${maxRetries}). 현재 값: ${price}`);
    
    if (attempt < maxRetries) {
      Logger.log(`${retryDelay/1000}초 후 다시 시도합니다...`);
      Utilities.sleep(retryDelay);
    }
  }
  
  Logger.log(`${ticker} 가격을 ${maxRetries}번 시도 후에도 받아오지 못했습니다. 오류 처리가 필요합니다.`);
  return null;
}

function getRealEstatePrice() {
  var apiKey = "2e4f2ef8f5964d9f80eba4b0fad27d04"; // 실제 API 키로 교체해야 합니다
  var baseUrl = "https://www.reb.or.kr/r-one/openapi/SttsApiTblData.do";

  // 현재 날짜 가져오기
  var currentDate = new Date();
  
  // 최대 12개월 전까지 시도
  for (var i = 1; i <= 12; i++) {
    // i개월 전 날짜 계산
    var targetDate = new Date(currentDate.getFullYear(), currentDate.getMonth() - i, 1);
    var yearMonth = Utilities.formatDate(targetDate, "GMT+9", "yyyyMM");

    var params = {
      KEY: apiKey,
      Type: "xml",
      pIndex: 1,
      pSize: 100,
      STATBL_ID: "A_2024_00178",
      DTACYCLE_CD: "MM",
      WRTTIME_IDTFR_ID: yearMonth,
      CLS_ID: "500001"
    };

    var url = baseUrl + "?" + Object.keys(params).map(key => key + "=" + encodeURIComponent(params[key])).join("&");

    try {
      var response = UrlFetchApp.fetch(url);
      var xmlContent = response.getContentText();
      Logger.log("API 응답: " + xmlContent); // 응답 내용 로깅

      var document = XmlService.parse(xmlContent);
      var root = document.getRootElement();
      
      // 안전한 방식으로 RESULT 요소 찾기
      var resultElement = root.getChild('RESULT') || 
                          (root.getChild('head') ? root.getChild('head').getChild('RESULT') : null);

      if (resultElement === null) {
        Logger.log("RESULT 요소를 찾을 수 없습니다. XML 구조: " + xmlContent);
        continue; // 다음 월로 넘어감
      }

      var resultCode = resultElement.getChildText('CODE');
      var resultMessage = resultElement.getChildText('MESSAGE');

      if (resultCode === 'INFO-000') {
        // 성공적으로 데이터를 찾음
        var rowElement = root.getChild('row');
        if (rowElement) {
          var data = {
            date: rowElement.getChildText('WRTTIME_DESC'),
            value: parseFloat(rowElement.getChildText('DTA_VAL'))
          };
          Logger.log("성공적으로 데이터를 찾았습니다: " + JSON.stringify(data));
          return data;
        } else {
          Logger.log("row 요소를 찾을 수 없습니다.");
          continue;
        }
      } else if (resultCode === 'INFO-200') {
        // 데이터가 없음, 다음 달로 넘어감
        Logger.log(yearMonth + "에 대한 데이터가 없습니다. 이전 달을 확인합니다.");
        continue;
      } else {
        // 다른 오류
        Logger.log("API 오류: " + resultMessage);
        continue;
      }
    } catch (error) {
      Logger.log("API 호출 중 오류 발생: " + error.toString());
      continue; // 다음 월로 넘어감
    }
  }
  
  // 12개월 동안 데이터를 찾지 못함
  Logger.log("최근 12개월 동안 데이터를 찾지 못했습니다.");
  return null;
}

function testRealEstateAPI() {
  var priceData = getRealEstatePrice();
  if (priceData !== null) {
    Logger.log("가져온 지가지수 데이터: " + JSON.stringify(priceData));
  } else {
    Logger.log("지가지수 데이터를 가져오는데 실패했습니다.");
  }
}

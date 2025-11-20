// Google Apps Script 코드
// 이 코드를 Google 스프레드시트의 Apps Script 편집기에 붙여넣으세요
// 배포: Apps Script 편집기 > 배포 > 새 배포 > 유형: 웹앱 > 액세스 권한: 모든 사용자

/**
 * 스프레드시트를 열 때 자동으로 실행되는 함수
 * 커스텀 메뉴를 추가하여 관리 기능을 제공합니다
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🍱 점심주문 관리')
    .addItem('🧹 중복 주문 데이터 정리', 'cleanupDuplicateOrders')
    .addItem('📊 오늘 주문 통계 보기', 'showTodayStats')
    .addToUi();
}

/**
 * POST 요청 처리 함수
 * 웹 앱에서 주문 데이터를 저장하거나 설정을 업데이트할 때 호출됩니다
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // 설정 저장 요청 처리
    if (data.action === 'saveSettings') {
      const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('설정');

      // 기존 설정 데이터 삭제 (헤더 제외)
      if (settingsSheet.getLastRow() > 1) {
        settingsSheet.deleteRows(2, settingsSheet.getLastRow() - 1);
      }

      // 새로운 설정 데이터 저장
      const settings = data.settings;
      settingsSheet.appendRow(['deadline', settings.deadline]);
      settingsSheet.appendRow(['reminderMinutes', settings.reminderMinutes]);
      settingsSheet.appendRow(['menuList', JSON.stringify(settings.menuList)]);
      settingsSheet.appendRow(['sideMenuList', JSON.stringify(settings.sideMenuList)]);
      settingsSheet.appendRow(['employees', JSON.stringify(settings.employees)]);
      settingsSheet.appendRow(['googleSheetUrl', settings.googleSheetUrl]);

      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: '설정이 저장되었습니다'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 주문 저장 요청 처리
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문내역');
    const allData = sheet.getDataRange().getValues();
    let rowToUpdate = -1;

    // ⭐ 중복 체크: 같은 날짜 + 같은 사용자의 기존 주문이 있는지 확인
    // 있으면 해당 행을 업데이트, 없으면 새로 추가
    for (let i = 1; i < allData.length; i++) {
      // 날짜 값 처리 (Date 객체를 문자열로 변환)
      let dateValue = allData[i][0];
      let dateStr = '';

      if (dateValue && dateValue.getTime) {
        // Date 객체인 경우 문자열로 변환
        dateStr = Utilities.formatDate(dateValue, 'Europe/Moscow', 'yyyy-MM-dd');
      } else if (typeof dateValue === 'string' && dateValue.trim()) {
        // 이미 문자열인 경우
        dateStr = dateValue.trim().split('T')[0];
      }

      // 날짜와 사용자명이 모두 일치하는지 확인
      if (dateStr === data.date && allData[i][1] === data.user) {
        rowToUpdate = i + 1; // 스프레드시트 행 번호 (1-based index)
        Logger.log(`기존 주문 발견: ${data.user} (행 ${rowToUpdate}) - 업데이트 예정`);
        break;
      }
    }

    // 저장할 데이터 행 구성
    const row = [
      data.date,                      // 날짜
      data.user,                      // 사용자명
      data.menu,                      // 메뉴
      data.time,                      // 주문 시간
      data.isGuest ? '손님' : '직원', // 구분
      new Date().toISOString()        // 최종 수정 시간 (타임스탬프)
    ];

    // 기존 주문이 있으면 업데이트, 없으면 새로 추가
    if (rowToUpdate > 0) {
      // ✅ 업데이트: 기존 행의 데이터를 덮어씀
      sheet.getRange(rowToUpdate, 1, 1, row.length).setValues([row]);
      Logger.log(`주문 업데이트: ${data.user} - ${data.menu} (행 ${rowToUpdate})`);
    } else {
      // ✅ 새로 추가: 맨 아래에 새 행 추가
      sheet.appendRow(row);
      Logger.log(`새 주문 추가: ${data.user} - ${data.menu}`);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: '주문이 저장되었습니다'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log('오류 발생: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * GET 요청 처리 함수
 * 웹 앱에서 주문 데이터나 설정을 불러올 때 호출됩니다
 */
function doGet(e) {
  try {
    const action = e.parameter.action;

    // 설정 가져오기
    if (action === 'getSettings') {
      const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('설정');
      const settingsData = settingsSheet.getDataRange().getValues();
      const settings = {};

      // 설정 데이터 파싱
      for (let i = 1; i < settingsData.length; i++) {
        const key = settingsData[i][0];
        let value = settingsData[i][1];

        // JSON 문자열은 객체로 변환
        if (key === 'menuList' || key === 'sideMenuList' || key === 'employees') {
          value = JSON.parse(value);
        } else if (key === 'reminderMinutes') {
          value = parseInt(value);
        }

        settings[key] = value;
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        settings: settings
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 주문 가져오기 (오늘 날짜의 주문만)
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문내역');

    // 모스크바 시간대로 오늘 날짜 가져오기
    const today = Utilities.formatDate(new Date(), 'Europe/Moscow', 'yyyy-MM-dd');

    const allData = sheet.getDataRange().getValues();
    const userOrders = {}; // 사용자별 최신 주문을 저장할 객체

    // ⭐ 중복 필터링: 같은 사용자의 주문이 여러 개 있으면 가장 최근 것만 반환
    for (let i = 1; i < allData.length; i++) {
      // 날짜 값 처리 (Date 객체 또는 문자열)
      let dateValue = allData[i][0];
      let dateStr = '';

      if (dateValue && dateValue.getTime) {
        // Date 객체인 경우
        dateStr = Utilities.formatDate(dateValue, 'Europe/Moscow', 'yyyy-MM-dd');
      } else if (typeof dateValue === 'string' && dateValue.trim()) {
        // 문자열인 경우
        dateStr = dateValue.trim().split('T')[0]; // ISO 형식이면 날짜 부분만 추출
      }

      // 오늘 날짜의 주문만 처리
      if (dateStr === today) {
        const userName = allData[i][1];
        const orderData = {
          date: dateStr,
          user: userName,
          menu: allData[i][2],
          time: allData[i][3],
          isGuest: allData[i][4] === '손님',
          rowIndex: i // 행 번호 저장 (나중에 최신 것 찾기 위해)
        };

        // 같은 사용자의 주문이 이미 있으면, 더 최근(더 큰 rowIndex)인 것으로 교체
        if (!userOrders[userName] || userOrders[userName].rowIndex < orderData.rowIndex) {
          userOrders[userName] = orderData;
        }
      }
    }

    // Map을 배열로 변환 (rowIndex 제거)
    const todayOrders = Object.values(userOrders).map(order => ({
      date: order.date,
      user: order.user,
      menu: order.menu,
      time: order.time,
      isGuest: order.isGuest
    }));

    Logger.log(`오늘 주문 ${todayOrders.length}건 반환`);

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      orders: todayOrders
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log('오류 발생: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString(),
      orders: []
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 🧹 중복 주문 데이터 정리 함수
 * 같은 날짜, 같은 사용자의 중복 주문이 있으면 가장 최근 것만 남기고 나머지 삭제
 *
 * 사용 방법:
 * 1. 스프레드시트 메뉴 > 점심주문 관리 > 중복 주문 데이터 정리
 * 2. 또는 Apps Script 편집기에서 직접 실행
 */
function cleanupDuplicateOrders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문내역');
  const allData = sheet.getDataRange().getValues();

  // 날짜+사용자를 키로 하는 Map (가장 최근 주문만 보관)
  const uniqueOrders = {};
  const rowsToDelete = []; // 삭제할 행 번호 목록

  // 헤더를 제외하고 모든 데이터 검사
  for (let i = 1; i < allData.length; i++) {
    const rowNum = i + 1; // 스프레드시트 행 번호 (1-based)

    // 날짜 값 처리
    let dateValue = allData[i][0];
    let dateStr = '';

    if (dateValue && dateValue.getTime) {
      dateStr = Utilities.formatDate(dateValue, 'Europe/Moscow', 'yyyy-MM-dd');
    } else if (typeof dateValue === 'string' && dateValue.trim()) {
      dateStr = dateValue.trim().split('T')[0];
    } else {
      continue; // 날짜가 없으면 스킵
    }

    const userName = allData[i][1];
    const key = `${dateStr}|${userName}`; // 날짜+사용자를 조합한 고유 키

    if (uniqueOrders[key]) {
      // 이미 같은 키가 있으면 중복
      // 더 최근 것을 남기기 위해 현재 행과 기존 행 중 나중 것을 선택
      const existingRow = uniqueOrders[key];

      // 타임스탬프 비교 (6번째 컬럼)
      const currentTimestamp = allData[i][5];
      const existingTimestamp = allData[existingRow - 1][5];

      if (currentTimestamp > existingTimestamp) {
        // 현재 것이 더 최근이면, 기존 것을 삭제 목록에 추가
        rowsToDelete.push(existingRow);
        uniqueOrders[key] = rowNum; // 현재 행으로 교체
      } else {
        // 기존 것이 더 최근이면, 현재 것을 삭제 목록에 추가
        rowsToDelete.push(rowNum);
      }
    } else {
      // 처음 나온 키면 저장
      uniqueOrders[key] = rowNum;
    }
  }

  // 삭제할 행이 있으면 역순으로 삭제 (뒤에서부터 삭제해야 인덱스가 안 깨짐)
  if (rowsToDelete.length > 0) {
    rowsToDelete.sort((a, b) => b - a); // 내림차순 정렬

    for (let i = 0; i < rowsToDelete.length; i++) {
      sheet.deleteRow(rowsToDelete[i]);
      Logger.log(`중복 행 삭제: ${rowsToDelete[i]}`);
    }

    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '✅ 중복 데이터 정리 완료',
      `${rowsToDelete.length}개의 중복 주문이 삭제되었습니다.\n\n` +
      `같은 날짜, 같은 사용자의 주문 중 가장 최근 것만 남겼습니다.`,
      ui.ButtonSet.OK
    );
  } else {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '✅ 중복 데이터 없음',
      '중복된 주문 데이터가 없습니다.',
      ui.ButtonSet.OK
    );
  }

  Logger.log(`중복 정리 완료: ${rowsToDelete.length}개 행 삭제`);
}

/**
 * 📊 오늘 주문 통계 표시 함수
 * 오늘 날짜의 주문 통계를 대화상자로 보여줍니다
 */
function showTodayStats() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문내역');
  const today = Utilities.formatDate(new Date(), 'Europe/Moscow', 'yyyy-MM-dd');
  const allData = sheet.getDataRange().getValues();

  let totalOrders = 0;
  let employeeOrders = 0;
  let guestOrders = 0;
  const menuCount = {};

  for (let i = 1; i < allData.length; i++) {
    let dateValue = allData[i][0];
    let dateStr = '';

    if (dateValue && dateValue.getTime) {
      dateStr = Utilities.formatDate(dateValue, 'Europe/Moscow', 'yyyy-MM-dd');
    } else if (typeof dateValue === 'string' && dateValue.trim()) {
      dateStr = dateValue.trim().split('T')[0];
    }

    if (dateStr === today) {
      totalOrders++;

      if (allData[i][4] === '손님') {
        guestOrders++;
      } else {
        employeeOrders++;
      }

      const menu = allData[i][2];
      menuCount[menu] = (menuCount[menu] || 0) + 1;
    }
  }

  let message = `📅 날짜: ${today}\n\n`;
  message += `📊 전체 주문: ${totalOrders}건\n`;
  message += `👥 직원: ${employeeOrders}건\n`;
  message += `🎯 손님: ${guestOrders}건\n\n`;
  message += `🍱 메뉴별 주문:\n`;

  // 메뉴별 통계를 주문 수 내림차순으로 정렬
  const sortedMenus = Object.entries(menuCount).sort((a, b) => b[1] - a[1]);
  for (const [menu, count] of sortedMenus) {
    message += `  • ${menu}: ${count}건\n`;
  }

  const ui = SpreadsheetApp.getUi();
  ui.alert('📊 오늘 주문 통계', message, ui.ButtonSet.OK);
}

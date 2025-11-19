// Google Apps Script 코드
// 이 코드를 Google 스프레드시트의 Apps Script 편집기에 붙여넣으세요

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // 설정 저장
    if (data.action === 'saveSettings') {
      const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('설정');

      if (settingsSheet.getLastRow() > 1) {
        settingsSheet.deleteRows(2, settingsSheet.getLastRow() - 1);
      }

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

    // 주문 저장
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문내역');
    const today = new Date().toISOString().split('T')[0];
    const allData = sheet.getDataRange().getValues();
    let rowToUpdate = -1;

    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === data.date && allData[i][1] === data.user) {
        rowToUpdate = i + 1;
        break;
      }
    }

    const row = [
      data.date,
      data.user,
      data.menu,
      data.time,
      data.isGuest ? '손님' : '직원',
      new Date().toISOString()
    ];

    if (rowToUpdate > 0) {
      sheet.getRange(rowToUpdate, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: '주문이 저장되었습니다'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    const action = e.parameter.action;

    // 설정 가져오기
    if (action === 'getSettings') {
      const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('설정');
      const settingsData = settingsSheet.getDataRange().getValues();
      const settings = {};

      for (let i = 1; i < settingsData.length; i++) {
        const key = settingsData[i][0];
        let value = settingsData[i][1];

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

    // 주문 가져오기
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문내역');

    // 모스크바 시간대로 오늘 날짜 가져오기 (Apps Script 시간대 설정 활용)
    const today = Utilities.formatDate(new Date(), 'Europe/Moscow', 'yyyy-MM-dd');

    const allData = sheet.getDataRange().getValues();
    const userOrders = {}; // 사용자별 최신 주문을 저장할 객체

    for (let i = 1; i < allData.length; i++) {
      // 날짜를 문자열로 변환해서 비교
      let dateValue = allData[i][0];
      let dateStr = '';

      // Date 객체 체크 (Apps Script에서 더 안전한 방법)
      if (dateValue && dateValue.getTime) {
        dateStr = Utilities.formatDate(dateValue, 'Europe/Moscow', 'yyyy-MM-dd');
      } else if (typeof dateValue === 'string' && dateValue.trim()) {
        // 이미 문자열이면 그대로 사용
        dateStr = dateValue.trim().split('T')[0]; // ISO 형식인 경우 날짜 부분만 추출
      }

      if (dateStr === today) {
        const userName = allData[i][1];
        const orderData = {
          date: dateStr,
          user: userName,
          menu: allData[i][2],
          time: allData[i][3],
          isGuest: allData[i][4] === '손님',
          rowIndex: i // 행 인덱스 저장 (나중에 가장 최근 것 찾기 위해)
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

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      orders: todayOrders
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString(),
      orders: []
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

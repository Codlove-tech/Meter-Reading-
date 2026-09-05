const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID_HERE';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(getSheetName(data.date)) || createMonthlySheet(data.date);
    
    // Find the row for the current date
    const dateRow = findDateRow(sheet, data.date);
    
    // Map data to columns based on your format
    updateMeterReadings(sheet, dateRow, data);
    
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getSheetName(dateStr) {
  const date = new Date(dateStr);
  const monthNames = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 
                      'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER'];
  return `${monthNames[date.getMonth()]} ${date.getFullYear()}`;
}

function createMonthlySheet(sheetName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.insertSheet(sheetName);
  
  // Create headers based on your format
  const headers = [
    'DATE',
    'MAIN INCOMER - PRESENT', 'MAIN INCOMER - PREVIOUS', 'MAIN INCOMER - CONSUMPTION',
    'FEWA MAIN - PRESENT', 'FEWA MAIN - PREVIOUS', 'FEWA MAIN - CONSUMPTION',
    'ELECTRICITY BUDGET', 'CHILLER SET POINT',
    // AL HAMRA Meters 1-9
    'AL HAMRA 1 - PRESENT', 'AL HAMRA 1 - PREVIOUS', 'AL HAMRA 1 - CONSUMPTION',
    'AL HAMRA 2 - PRESENT', 'AL HAMRA 2 - PREVIOUS', 'AL HAMRA 2 - CONSUMPTION',
    'AL HAMRA 3 - PRESENT', 'AL HAMRA 3 - PREVIOUS', 'AL HAMRA 3 - CONSUMPTION',
    'AL HAMRA 4 - PRESENT', 'AL HAMRA 4 - PREVIOUS', 'AL HAMRA 4 - CONSUMPTION',
    'AL HAMRA 5 - PRESENT', 'AL HAMRA 5 - PREVIOUS', 'AL HAMRA 5 - CONSUMPTION',
    'AL HAMRA 6 - PRESENT', 'AL HAMRA 6 - PREVIOUS', 'AL HAMRA 6 - CONSUMPTION',
    'AL HAMRA 7 - PRESENT', 'AL HAMRA 7 - PREVIOUS', 'AL HAMRA 7 - CONSUMPTION',
    'AL HAMRA 8 - PRESENT', 'AL HAMRA 8 - PREVIOUS', 'AL HAMRA 8 - CONSUMPTION',
    'AL HAMRA 9 - PRESENT', 'AL HAMRA 9 - PREVIOUS', 'AL HAMRA 9 - CONSUMPTION',
    // Water
    'UTICO WATER - PRESENT', 'UTICO WATER - PREVIOUS', 'UTICO WATER - CONSUMPTION',
    'WATER BUDGET', 'OCCUPANCY ACTUAL', 'OCCUPANCY BUDGET',
    // Cooling Tower
    'COOLING TOWER - PRESENT', 'COOLING TOWER - PREVIOUS', 'COOLING TOWER - CONSUMPTION',
    'CHILLER SET POINT (COOLING)', 'BOND CONSTRUCTION - PRESENT', 'BOND CONSTRUCTION - PREVIOUS', 'BOND CONSTRUCTION - CONSUMPTION',
    // Pools
    'MAIN POOL - PRESENT', 'MAIN POOL - PREVIOUS', 'MAIN POOL - CONSUMPTION',
    'SPLASH POOL - PRESENT', 'SPLASH POOL - PREVIOUS', 'SPLASH POOL - CONSUMPTION',
    'SUNSET POOL - PRESENT', 'SUNSET POOL - PREVIOUS', 'SUNSET POOL - CONSUMPTION',
    // Irrigation
    'MAIN IRRIGATION - PRESENT', 'MAIN IRRIGATION - PREVIOUS', 'MAIN IRRIGATION - CONSUMPTION',
    'BOARDWALK IRRIGATION - PRESENT', 'BOARDWALK IRRIGATION - PREVIOUS', 'BOARDWALK IRRIGATION - CONSUMPTION',
    'AL HAMRA RESIDENCE - PRESENT', 'AL HAMRA RESIDENCE - PREVIOUS', 'AL HAMRA RESIDENCE - CONSUMPTION',
    // LPG
    'BOILER - PRESENT', 'BOILER - PREVIOUS', 'BOILER - CONSUMPTION',
    'KITCHEN - PRESENT', 'KITCHEN - PREVIOUS', 'KITCHEN - CONSUMPTION',
    'CONVERSION FACTOR', 'PRESSURE FACTOR',
    'TANK 1 GAUGE', 'TANK 2 GAUGE', 'TANK BALANCE', 'DELIVERY RECEIVED',
    'BOILER LPG BUDGET', 'CLARIFIER SET POINT'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.getRange(1, 1, 1, headers.length).setBackground('#1e3c72');
  sheet.getRange(1, 1, 1, headers.length).setFontColor('#ffffff');
  
  // Create date rows for the entire month
  const date = new Date(sheetName);
  const daysInMonth = new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();
  
  for (let day = 1; day <= daysInMonth; day++) {
    const rowDate = new Date(date.getFullYear(), date.getMonth(), day);
    sheet.getRange(day + 1, 1).setValue(rowDate);
    sheet.getRange(day + 1, 1).setNumberFormat('dd mm yyyy');
  }
  
  return sheet;
}

function findDateRow(sheet, dateStr) {
  const data = sheet.getDataRange().getValues();
  const targetDate = new Date(dateStr);
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] instanceof Date) {
      const rowDate = data[i][0];
      if (rowDate.getDate() === targetDate.getDate() && 
          rowDate.getMonth() === targetDate.getMonth() && 
          rowDate.getFullYear() === targetDate.getFullYear()) {
        return i + 1; // 1-based row index
      }
    }
  }
  return -1;
}

function updateMeterReadings(sheet, row, data) {
  if (row === -1) return;
  
  const updates = [];
  
  // Electrical meters
  updates.push([2, data.electrical?.mainIncomer?.present || '']);
  updates.push([3, data.electrical?.mainIncomer?.previous || '']);
  updates.push([4, data.electrical?.mainIncomer?.consumption || '']);
  updates.push([5, data.electrical?.fewaMain?.present || '']);
  updates.push([6, data.electrical?.fewaMain?.previous || '']);
  updates.push([7, data.electrical?.fewaMain?.consumption || '']);
  updates.push([8, data.electrical?.budget || '']);
  updates.push([9, data.electrical?.chillerSetPoint || '']);
  
  // AL HAMRA meters (columns 10-36)
  if (data.electrical?.alHamra) {
    data.electrical.alHamra.forEach((meter, index) => {
      const baseCol = 10 + (index * 3);
      updates.push([baseCol, meter.present || '']);
      updates.push([baseCol + 1, meter.previous || '']);
      updates.push([baseCol + 2, meter.consumption || '']);
    });
  }
  
  // Water meters (columns 37-42)
  updates.push([37, data.water?.utico?.present || '']);
  updates.push([38, data.water?.utico?.previous || '']);
  updates.push([39, data.water?.utico?.consumption || '']);
  updates.push([40, data.water?.budget || '']);
  updates.push([41, data.water?.occupancy?.actual || '']);
  updates.push([42, data.water?.occupancy?.budget || '']);
  
  // Cooling tower (columns 43-49)
  updates.push([43, data.coolingTower?.meter?.present || '']);
  updates.push([44, data.coolingTower?.meter?.previous || '']);
  updates.push([45, data.coolingTower?.meter?.consumption || '']);
  updates.push([46, data.coolingTower?.chillerSetPoint || '']);
  updates.push([47, data.coolingTower?.bondConstruction?.present || '']);
  updates.push([48, data.coolingTower?.bondConstruction?.previous || '']);
  updates.push([49, data.coolingTower?.bondConstruction?.consumption || '']);
  
  // Pools (columns 50-58)
  updates.push([50, data.pools?.main?.present || '']);
  updates.push([51, data.pools?.main?.previous || '']);
  updates.push([52, data.pools?.main?.consumption || '']);
  updates.push([53, data.pools?.splash?.present || '']);
  updates.push([54, data.pools?.splash?.previous || '']);
  updates.push([55, data.pools?.splash?.consumption || '']);
  updates.push([56, data.pools?.sunset?.present || '']);
  updates.push([57, data.pools?.sunset?.previous || '']);
  updates.push([58, data.pools?.sunset?.consumption || '']);
  
  // Irrigation (columns 59-67)
  updates.push([59, data.irrigation?.main?.present || '']);
  updates.push([60, data.irrigation?.main?.previous || '']);
  updates.push([61, data.irrigation?.main?.consumption || '']);
  updates.push([62, data.irrigation?.boardwalk?.present || '']);
  updates.push([63, data.irrigation?.boardwalk?.previous || '']);
  updates.push([64, data.irrigation?.boardwalk?.consumption || '']);
  updates.push([65, data.irrigation?.residence?.present || '']);
  updates.push([66, data.irrigation?.residence?.previous || '']);
  updates.push([67, data.irrigation?.residence?.consumption || '']);
  
  // LPG (columns 68-78)
  updates.push([68, data.lpg?.boiler?.present || '']);
  updates.push([69, data.lpg?.boiler?.previous || '']);
  updates.push([70, data.lpg?.boiler?.consumption || '']);
  updates.push([71, data.lpg?.kitchen?.present || '']);
  updates.push([72, data.lpg?.kitchen?.previous || '']);
  updates.push([73, data.lpg?.kitchen?.consumption || '']);
  updates.push([74, data.lpg?.conversionFactor || '']);
  updates.push([75, data.lpg?.pressureFactor || '']);
  updates.push([76, data.lpg?.tank1Gauge || '']);
  updates.push([77, data.lpg?.tank2Gauge || '']);
  updates.push([78, data.lpg?.tankBalance || '']);
  updates.push([79, data.lpg?.deliveryReceived || '']);
  updates.push([80, data.lpg?.boilerBudget || '']);
  updates.push([81, data.lpg?.clarifierSetPoint || '']);
  
  // Apply all updates
  updates.forEach(([col, value]) => {
    if (value !== '') {
      sheet.getRange(row, col).setValue(value);
    }
  });
}

function doGet(e) {
  if (e.parameter.action === 'getHistory') {
    return getHistory();
  }
  
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'OK' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getHistory() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheets = ss.getSheets();
    const history = [];
    
    sheets.forEach(sheet => {
      const data = sheet.getDataRange().getValues();
      if (data.length > 1) {
        const headers = data[0];
        for (let i = 1; i < data.length; i++) {
          if (data[i][0] && data[i][2]) { // Has date and main incomer reading
            history.push({
              date: data[i][0],
              electrical: {
                mainIncomer: {
                  present: data[i][1],
                  previous: data[i][2],
                  consumption: data[i][3]
                },
                fewaMain: {
                  present: data[i][4],
                  previous: data[i][5],
                  consumption: data[i][6]
                }
              },
              water: {
                utico: {
                  present: data[i][36],
                  previous: data[i][37],
                  consumption: data[i][38]
                }
              },
              coolingTower: {
                meter: {
                  present: data[i][42],
                  previous: data[i][43],
                  consumption: data[i][44]
                }
              }
            });
          }
        }
      }
    });
    
    return ContentService
      .createTextOutput(JSON.stringify(history))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
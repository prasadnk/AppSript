function onOpen() {
    SpreadsheetApp.getUi()
        .createAddonMenu()
        .addItem('Download Data', 'showSidebar')
        .addToUi();
  }
  
  function showSidebar() {
    SpreadsheetApp.getUi()
        .showSidebar(HtmlService.createHtmlOutputFromFile('index'));
  }
  
  const AAPI_URL_BASE = 'https://api.smartsheet.com/2.0/sheets/';
  
  function getCredentials() {
    const userProperties = PropertiesService.getUserProperties();
    let apiAccessToken = userProperties.getProperty('SMARTSHEET_API_ACCESS_TOKEN');
    let sheetId = userProperties.getProperty('SMARTSHEET_ID');
  
    if (!apiAccessToken || !sheetId) {
      apiAccessToken = Browser.inputBox('Enter Your Smartsheet API Access Token Here:');
      sheetId = Browser.inputBox('Enter Your Smartsheet ID Here:');
      userProperties.setProperty('SMARTSHEET_API_ACCESS_TOKEN', apiAccessToken);
      userProperties.setProperty('SMARTSHEET_ID', sheetId);
    }
  
    return { apiAccessToken, sheetId };
  }
  
  function fetchFromSmartsheet(apiAccessToken, sheetId) {
    const url = `${AAPI_URL_BASE}${sheetId}`;
    const options = {
      'method': 'get',
      'headers': {
        'Authorization': `Bearer ${apiAccessToken}`,
      },
      'muteHttpExceptions': true
    };
  
    try {
      const response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() === 200) {
        return JSON.parse(response.getContentText());
      } else {
        throw new Error(`Error ${response.getResponseCode()}: ${response.getContentText()}`);
      }
    } catch (error) {
      Logger.log(error.message);
      throw new Error('Failed to fetch data from Smartsheet');
    }
  }
  
  function updateSpreadsheet(data) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(data.name);
  
    if (!sheet) {
      sheet = spreadsheet.insertSheet(data.name);
    } else {
      sheet.clear();
    }
  
    const headers = ['Updated As Of'].concat(data.columns.map(column => column.title));
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
    const currentTime = new Date().toLocaleString();
    const rowsData = data.rows.map(row => {
      const rowData = row.cells.map(cell => cell.value || '');
      rowData.unshift(currentTime);
      return rowData;
    });
  
    sheet.getRange(2, 1, rowsData.length, headers.length).setValues(rowsData);
  }
  
  function getDataFromSmartsheet() {
    const { apiAccessToken, sheetId } = getCredentials();
    const data = fetchFromSmartsheet(apiAccessToken, sheetId);
    updateSpreadsheet(data);
  }
  
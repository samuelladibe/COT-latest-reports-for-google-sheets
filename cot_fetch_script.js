// Configuration using CFTC contract market codes
const COT_CONFIG = {
  'JAPANESE YEN': {
    cftc_code: '097741',
    display_name: 'JAPANESE YEN - CHICAGO MERCANTILE EXCHANGE'
  },
  'GOLD': {
    cftc_code: '088691',
    display_name: 'GOLD - COMMODITY EXCHANGE INC.'
  },
  'S&P 500': {
    cftc_code: '13874V',
    display_name: 'S&P 500 - CHICAGO MERCANTILE EXCHANGE'
  },
  'EUR': {
    cftc_code: '099741',
    display_name: 'EURO FX - CHICAGO MERCANTILE EXCHANGE'
  }
};

/**
 * Main function to update COT data for all configured assets
 * Iterates through COT_CONFIG and fetches/updates data for each asset
 */
function updateAllCOTData() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    for (const [assetName, config] of Object.entries(COT_CONFIG)) {
      Logger.log(`Processing ${assetName}...`);
      
      const cotData = fetchCOTDataByCode(config);
      if (cotData) {
        updateAssetSheet(spreadsheet, assetName, cotData);
        Logger.log(`✓ ${assetName} updated successfully`);
      } else {
        Logger.log(`✗ No data found for ${assetName} (Code: ${config.cftc_code})`);
      }
      
      Utilities.sleep(1000);
    }
    
    Logger.log('COT data update completed successfully');
  } catch (error) {
    Logger.log(`Error in updateAllCOTData: ${error.toString()}`);
  }
}

/**
 * Fetches COT data from CFTC API using the contract market code
 * @param {Object} config - Asset configuration containing cftc_code and display_name
 * @returns {Object|null} COT data object or null if fetch fails
 */
function fetchCOTDataByCode(config) {
  try {
    // Fetch data using the specific CFTC contract code
    const apiUrl = `https://publicreporting.cftc.gov/resource/6dca-aqww.json?cftc_contract_market_code=${config.cftc_code}&$limit=1&$order=report_date_as_yyyy_mm_dd DESC`; //latest report for each asset
    
    Logger.log(`Fetching data for ${config.display_name} (Code: ${config.cftc_code})`);
    
    const response = UrlFetchApp.fetch(apiUrl, {
      muteHttpExceptions: true
    });
    
    Logger.log(`API Response Code: ${response.getResponseCode()}`);
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      
      if (data && data.length > 0) {
        Logger.log(`✓ Found data for ${config.display_name}`);
        return parseCOTAPIResponse(data[0], config.display_name);
      } else {
        Logger.log(`✗ No data returned for code ${config.cftc_code}`);
        return null;
      }
    } else {
      Logger.log(`API Error: ${response.getResponseCode()} - ${response.getContentText().substring(0, 200)}`);
      return null;
    }
    
  } catch (error) {
    Logger.log(`Error fetching data for ${config.display_name}: ${error.toString()}`);
    return null;
  }
}

/**
 * Parses raw API response into standardized COT data object
 * @param {Object} apiData - Raw data from CFTC API
 * @param {string} displayName - Display name for the asset
 * @returns {Object} Standardized COT data object
 */
function parseCOTAPIResponse(apiData, displayName) {
  if (!apiData) return null;
  
  // Extract just the date part from the ISO string
  const reportDateString = apiData.report_date_as_yyyy_mm_dd ? 
    apiData.report_date_as_yyyy_mm_dd.split('T')[0] : '';
  
  // Convert the report date string to a Date object
  const reportDate = parseReportDate(apiData.report_date_as_yyyy_mm_dd);
  
  return {
    nonCommercialLong: parseInt(apiData.noncomm_positions_long_all || '0'),
    nonCommercialShort: parseInt(apiData.noncomm_positions_short_all || '0'),
    commercialLong: parseInt(apiData.comm_positions_long_all || '0'),
    commercialShort: parseInt(apiData.comm_positions_short_all || '0'),
    openInterest: parseInt(apiData.open_interest_all || '0'),
    timestamp: reportDate, // Use the report date instead of current date
    reportDate: reportDateString, // Store the clean "YYYY-MM-DD" string
    marketName: displayName,
    cftcCode: apiData.cftc_contract_market_code
  };
}

/**
 * Parses report date string from API into Date object
 * Extracts YYYY-MM-DD from ISO format and creates Date object
 * @param {string} dateString - Date string from API (ISO format)
 * @returns {Date} Date object for the report date
 */
function parseReportDate(dateString) {
  try {
    if (dateString) {
      // Extract just the date part (YYYY-MM-DD) from the ISO string
      const datePart = dateString.split('T')[0];
      if (datePart.match(/^\d{4}-\d{2}-\d{2}$/)) {
        return new Date(datePart + 'T00:00:00'); // Set to start of day
      }
    }
    // Fallback to current date if parsing fails
    return new Date();
  } catch (error) {
    Logger.log(`Error parsing date ${dateString}: ${error.toString()}`);
    return new Date();
  }
}

/**
 * Updates the asset sheet with new COT data
 * Creates sheet if it doesn't exist, checks for duplicates before adding data
 * @param {Spreadsheet} spreadsheet - Google Sheets object
 * @param {string} assetName - Name of the asset (sheet name)
 * @param {Object} cotData - COT data to add/update
 */
function updateAssetSheet(spreadsheet, assetName, cotData) {
  try {
    let sheet = spreadsheet.getSheetByName(assetName);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(assetName);
      initializeAssetSheet(sheet);
    }
    
    // Check if we already have data for this report date
    const existingRow = findExistingReportDate(sheet, cotData.reportDate);
    
    if (existingRow) {
      Logger.log(`ℹ Data for report date ${cotData.reportDate} already exists at row ${existingRow}. Updating...`);
      updateExistingRow(sheet, existingRow, cotData);
    } else {
      // Add new row for new report date
      const lastRow = sheet.getLastRow();
      const updateRow = lastRow > 0 ? lastRow + 1 : 2;
      
      addNewRow(sheet, updateRow, cotData);
    }
    
    Logger.log(`Updated ${assetName} sheet successfully`);
    
  } catch (error) {
    Logger.log(`Error updating sheet for ${assetName}: ${error.toString()}`);
  }
}

/**
 * Finds if data for a specific report date already exists in the sheet
 * @param {Sheet} sheet - Google Sheet object
 * @param {string} reportDate - Report date in YYYY-MM-DD format
 * @returns {number|null} Row number if exists, null if not found
 */
function findExistingReportDate(sheet, reportDate) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null; // No data rows yet
    
    const dateRange = sheet.getRange(2, 1, lastRow - 1, 1);
    const dates = dateRange.getValues();
    
    for (let i = 0; i < dates.length; i++) {
      const cellDate = dates[i][0];
      if (cellDate instanceof Date) {
        const cellDateString = Utilities.formatDate(cellDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        if (cellDateString === reportDate) {
          return i + 2; // +2 because we start at row 2 and i is 0-based
        }
      }
    }
    
    return null;
  } catch (error) {
    Logger.log(`Error finding existing report date: ${error.toString()}`);
    return null;
  }
}

/**
 * Adds a new row of COT data to the sheet
 * @param {Sheet} sheet - Google Sheet object
 * @param {number} row - Row number to insert data
 * @param {Object} cotData - COT data to add
 */
function addNewRow(sheet, row, cotData) {
  const dataRow = [
    cotData.reportDate, // This is now the report date, not current date
    cotData.nonCommercialLong,
    cotData.nonCommercialShort,
    cotData.commercialLong,
    cotData.commercialShort,
    cotData.openInterest,
    cotData.nonCommercialLong - cotData.nonCommercialShort,
    cotData.commercialLong - cotData.commercialShort,
    calculateNetPositionPercentage(cotData)
  ];
  
  sheet.getRange(row, 1, 1, dataRow.length).setValues([dataRow]);
  formatDataRow(sheet, row);
  
  Logger.log(`✓ Added new data for report date: ${cotData.reportDate}`);
}

/**
 * Updates an existing row with new COT data
 * @param {Sheet} sheet - Google Sheet object
 * @param {number} row - Row number to update
 * @param {Object} cotData - COT data to update with
 */
function updateExistingRow(sheet, row, cotData) {
  const dataRow = [
    cotData.reportDate,
    cotData.nonCommercialLong,
    cotData.nonCommercialShort,
    cotData.commercialLong,
    cotData.commercialShort,
    cotData.openInterest,
    cotData.nonCommercialLong - cotData.nonCommercialShort,
    cotData.commercialLong - cotData.commercialShort,
    calculateNetPositionPercentage(cotData)
  ];
  
  sheet.getRange(row, 1, 1, dataRow.length).setValues([dataRow]);
  formatDataRow(sheet, row);
  
  Logger.log(`✓ Updated existing data for report date: ${cotData.reportDate}`);
}

/**
 * Initializes a new asset sheet with headers and formatting
 * @param {Sheet} sheet - Google Sheet object to initialize
 */
function initializeAssetSheet(sheet) {
  const headers = [
    'Report Date',
    'Non-Commercial Long',
    'Non-Commercial Short',
    'Commercial Long',
    'Commercial Short',
    'Open Interest',
    'Net Non-Commercial',
    'Net Commercial',
    'Net Position %'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4a86e8')
    .setFontColor('white');
  
  sheet.setColumnWidth(1, 120); // Report Date
  for (let i = 2; i <= headers.length; i++) {
    sheet.setColumnWidth(i, 120);
  }
  
  sheet.setFrozenRows(1);
}

/**
 * Applies formatting to a data row in the sheet
 * @param {Sheet} sheet - Google Sheet object
 * @param {number} row - Row number to format
 */
function formatDataRow(sheet, row) {
  const range = sheet.getRange(row, 1, 1, 9);
  
  // Format report date as yyyy-mm-dd
  sheet.getRange(row, 1).setNumberFormat('yyyy-mm-dd');
  
  // Format numbers with commas
  const numberRange = sheet.getRange(row, 2, 1, 8);
  numberRange.setNumberFormat('#,##0');
  
  // Format percentage
  sheet.getRange(row, 9).setNumberFormat('0.00%');
  
  // Add alternating row colors
  if (row % 2 === 0) {
    range.setBackground('#f3f3f3');
  }
}

/**
 * Calculates net position percentage for non-commercial traders
 * @param {Object} cotData - COT data object
 * @returns {number} Net position as percentage of open interest
 */
function calculateNetPositionPercentage(cotData) {
  const netNonCommercial = cotData.nonCommercialLong - cotData.nonCommercialShort;
  const totalOpenInterest = cotData.openInterest;
  return totalOpenInterest === 0 ? 0 : netNonCommercial / totalOpenInterest;
}

/**
 * Verifies that all configured CFTC contract codes are valid and return data
 * Tests each asset in COT_CONFIG to ensure API connectivity
 */
function verifyAllCodes() {
  Logger.log('Verifying all CFTC contract codes...');
  
  for (const [assetName, config] of Object.entries(COT_CONFIG)) {
    Logger.log(`\nVerifying ${assetName}:`);
    Logger.log(`- Code: ${config.cftc_code}`);
    Logger.log(`- Display: ${config.display_name}`);
    
    const cotData = fetchCOTDataByCode(config);
    if (cotData) {
      Logger.log(`✓ SUCCESS - Found data for report date: ${cotData.reportDate}`);
      Logger.log(`  NonComm Long: ${cotData.nonCommercialLong.toLocaleString()}`);
      Logger.log(`  NonComm Short: ${cotData.nonCommercialShort.toLocaleString()}`);
    } else {
      Logger.log(`✗ FAILED - No data found`);
    }
    
    Utilities.sleep(500);
  }
}

/**
 * Sets up automatic daily triggers to update COT data
 * Runs every day at 10 AM to fetch latest CFTC reports
 */
function setupTriggers() {
  ScriptApp.newTrigger('updateAllCOTData')
    .timeBased()
    .everyDays(1)
    .atHour(10)
    .create();
}

// --------------------------------------------------------------------------------------------------------------//

/**
 * Adds a new asset to the COT_CONFIG for tracking
 * @param {string} assetName - Name of the asset
 * @param {string} cftcCode - CFTC contract market code
 * @param {string} displayName - Display name for the asset
 */
function addNewAsset(assetName, cftcCode, displayName) {
  if (!COT_CONFIG[assetName]) {
    COT_CONFIG[assetName] = {
      cftc_code: cftcCode,
      display_name: displayName
    };
    Logger.log(`✓ Added ${assetName} with code ${cftcCode}`);
  } else {
    Logger.log(`ℹ ${assetName} already exists in config`);
  }
}

/**
 * Searches for Bitcoin COT data across CFTC datasets
 * Attempts to find Bitcoin futures data in main and alternative datasets
 */
function findBitcoinData() {
  Logger.log('Searching for Bitcoin data...');
  
  // Try the main dataset first
  const apiUrl = 'https://publicreporting.cftc.gov/resource/6dca-aqww.json?cftc_contract_market_code=133741&$limit=5';  // change cftc_contract_market_code here for other assets
  
  const response = UrlFetchApp.fetch(apiUrl, {
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() === 200) {
    const data = JSON.parse(response.getContentText());
    if (data && data.length > 0) {
      Logger.log('✓ Bitcoin found in main dataset:');
      data.forEach(item => {
        Logger.log(`- ${item.market_and_exchange_names} (${item.report_date_as_iso})`);
      });
    } else {
      Logger.log('✗ Bitcoin not found in main dataset with code 133741');
      
      // Search by name as fallback
      searchContractByName('BITCOIN');
    }
  }
}

/**
 * Searches for contracts by name in the CFTC API
 * Useful for discovering contract codes for new assets
 * @param {string} contractName - Name of contract to search for
 */
function searchContractByName(contractName) {
  Logger.log(`Searching for contracts containing: ${contractName}`);
  
  const apiUrl = `https://publicreporting.cftc.gov/resource/6dca-aqww.json?$where=market_and_exchange_names like '%${encodeURIComponent(contractName)}%'&$limit=10`;
  
  try {
    const response = UrlFetchApp.fetch(apiUrl, {
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      if (data && data.length > 0) {
        Logger.log(`Found ${data.length} contracts:`);
        data.forEach(item => {
          Logger.log(`- ${item.market_and_exchange_names} (Code: ${item.cftc_contract_market_code})`);
        });
      } else {
        Logger.log('No contracts found with that name');
      }
    }
  } catch (error) {
    Logger.log(`Search error: ${error.toString()}`);
  }
}

/**
 * Tests a specific asset to verify data fetching works
 * @param {string} assetName - Name of asset to test
 */
function testAsset(assetName) {
  const config = COT_CONFIG[assetName];
  if (!config) {
    Logger.log(`✗ ${assetName} not found in COT_CONFIG`);
    return;
  }
  
  Logger.log(`Testing ${assetName}...`);
  const cotData = fetchCOTDataByCode(config);
  
  if (cotData) {
    Logger.log(`✓ SUCCESS - ${config.display_name}`);
    Logger.log(`  Non-Commercial Long: ${cotData.nonCommercialLong.toLocaleString()}`);
    Logger.log(`  Non-Commercial Short: ${cotData.nonCommercialShort.toLocaleString()}`);
    Logger.log(`  Commercial Long: ${cotData.commercialLong.toLocaleString()}`);
    Logger.log(`  Commercial Short: ${cotData.commercialShort.toLocaleString()}`);
    Logger.log(`  Open Interest: ${cotData.openInterest.toLocaleString()}`);
    Logger.log(`  Net Position: ${(cotData.nonCommercialLong - cotData.nonCommercialShort).toLocaleString()}`);
  } else {
    Logger.log(`✗ FAILED - No data found`);
  }
}

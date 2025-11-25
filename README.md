## COT Reports Automation for Google Sheets
Automatically fetch and update CFTC Commitment of Traders (COT) reports data directly into Google Sheets using Google Apps Script.

## 🚀 Features
- **Automated Updates**: Daily automatic updates via triggers
- **Multiple Assets**: Support for JPY, Gold, S&P 500, EUR, and more
- **Accurate Dating**: Uses actual CFTC report dates instead of current date
- **Duplicate Prevention**: Prevents duplicate entries for the same report period
- **Professional Formatting**: Clean, readable output with proper formatting

## 🛠 Setup Instructions
### 1. Create Google Sheet
- Create a new Google Sheet named "COT Reports Dashboard"
- Or use any existing Google Sheet
### 2. Open Apps Script Editor
- Go to `Extensions > Apps Script`
- Delete the default code and paste the entire script
### 3. Configure Script Properties
Run this function once to set up required URLs:
```Javascript
function setupProperties() {
  PropertiesService.getScriptProperties().setProperties({
    'cot_curr_url': 'https://www.cftc.gov/dea/futures/deacmesf.htm',
    'cot_com_url': 'https://www.cftc.gov/dea/futures/deacmxsf.htm'
  });
}
```
### 4. Authorize the Script
- Save the script and run any function
- Authorize the script when prompted

## 📊 Configure Assets
The script comes pre-configured with these assets:

| Asset          |	  CFTC      |              Code	Display Name              |
| ---------------|--------------|---------------------------------------------|
| Japanese Yen	 |   097741	    |  JAPANESE YEN - CHICAGO MERCANTILE EXCHANGE |
| Gold	         |   088691	    |  GOLD - COMMODITY EXCHANGE INC.             |
| S&P 500	       |   13874V     |  S&P 500 - CHICAGO MERCANTILE EXCHANGE      |
| Euro FX	       |   099741	    |  EURO FX - CHICAGO MERCANTILE EXCHANGE      |

## Adding New Assets
Use the `addNewAsset()` function:
```Javascript
addNewAsset('BITCOIN', '133741', 'BITCOIN - CHICAGO MERCANTILE EXCHANGE');
}
```
Or manually add to `COT_CONFIG`:
```Javascript
const COT_CONFIG = {
  // ... existing assets ...
  'NEW_ASSET': {
    cftc_code: 'CFTC_CODE_HERE',
    display_name: 'ASSET NAME - EXCHANGE'
  }
};
```

## 🎯 How to Use
### Initial Setup
1. Run `setupProperties()` to configure URLs
2. Run `verifyAllCodes()` to test all asset connections
3. Run `testCOTUpdate()` for initial data population

### Daily Operation
- The script will run automatically at 10 AM daily via trigger
- Or manually run `updateAllCOTData()` anytime

### Testing
- `testAsset('GOLD')` - Test specific asset
- `verifyAllCodes()` - Verify all assets work
- `testCOTUpdate()` - Full update test

## ⚙️ Functions Overview
### Core Functions
- `updateAllCOTData()` - Main function to update all assets
- `fetchCOTDataByCode()` - Fetches data from CFTC API
- `updateAssetSheet()` - Updates Google Sheets with new data

### Utility Functions
- `setupTriggers()` - Sets up automatic daily updates. Advice : weekly script execution frequency as the COT results are not published every day.
- `verifyAllCodes()` - Tests all asset configurations
- `addNewAsset()` - Adds new assets to track

### Test Functions
- `testAsset()` - Tests specific asset
- `findBitcoinData()` - Searches for Bitcoin data
- `searchContractByName()` - Finds contracts by name

## 📈 Data Columns
Each asset sheet contains:
1. **Report Date** - Actual CFTC report date (YYYY-MM-DD
2. **Non-Commercial Long** - Long positions by speculators
3. **Non-Commercial Short** - Short positions by speculators
4. **Commercial Long** - Long positions by institutions
5. **Commercial Short** - Short positions by institutions
6. **Open Interest** - Total open contracts
7. **Net Non-Commercial** - Net position (Long - Short)
8. **Net Commercial** - Net position (Long - Short)
9. **Net Position %** - Net position as % of open interest

## 🔧 Troubleshooting
### Common Issues
1. **403 Errors**: The script uses CFTC API, not web scraping
2. **No Data Found**: Verify CFTC codes are correct using `verifyAllCodes()`
3. **Authorization Errors**: Re-run authorization for the script
4. **Duplicate Data**: Script automatically prevents duplicates

### Debugging
- Run `verifyAllCodes()` to test connectivity
- Check Apps Script logs for detailed error messages
- Use `testAsset('GOLD')` to isolate issues

## 📝 Notes
- CFTC reports are typically released weekly on Fridays
- The script fetches the most recent available report
- Data is stored with actual report dates for accurate historical tracking
- The script handles date formatting and prevents duplicate entries

##🔒 Permissions Required
- Google Sheets: Read/write access. Advice: Simply create your and then apply the script to the file via extensions
- URL Fetch: Access CFTC API
- Script Properties: Store configuration

## 📄 License
This project is open source and available under the MIT License.

## 🤝 Contributing
Feel free to submit issues and enhancement requests!
__________________________________________________________________________
Happy Trading! 📊🚀

/*
  A script that is used to call the quality scores at a Campaign, Account, and AdGroup level.
  Uses http://www.freeadwordsscripts.com/search?q=quality+score for examples.
  In addition to calling the quality scores, refers to already created spreadsheet with info of URL to output data to.
  
  Created by Tanner Leach w/ Chip Oglesby for FastPivot
*/
var SIG_FIGS = 10000; //this will give you 4 decimal places of accuracy
var APPEND = true; //set this to false to overwrite your data daily
var SPREADSHEET_URL = " "; 
var ACCOUNTS = ['354-254-9708'];
function main() {

  var accountSelector = MccApp.accounts();
  if (ACCOUNTS.length > 0) {
    accountSelector = accountSelector.withIds(ACCOUNTS);
  }
  
  var accountIterator = accountSelector.get();
  
  while (accountIterator.hasNext()) {
    var account = accountIterator.next();
    var account_id = account.getCustomerId();
    MccApp.select(account); //<--- makes the account selected ~~~ This is important
    SPREADSHEET_URL = selectSheet(account_id);
    
    var date_str = Utilities.formatDate(new Date(),AdWordsApp.currentAccount().getTimeZone(),'yyyy-MM-dd');
    //var account_id = AdWordsApp.currentAccount().getCustomerId();
    
    var kw_iter = AdWordsApp.keywords()
    .withCondition("Status = ACTIVE")
    .forDateRange("LAST_30_DAYS")
    .withCondition("Impressions > 0")
    .orderBy("Impressions DESC")
    .withLimit(50000)
    .get();
    
    //This is where i am going to store all my data
    var account_score_map = {};
    var camp_score_map = {};
    var ag_score_map = {};
    
    while(kw_iter.hasNext()) {
      var kw = kw_iter.next();
      var kw_stats = kw.getStatsFor("LAST_30_DAYS");
      var imps = kw_stats.getImpressions();
      var qs = kw.getQualityScore();
      var camp_name = kw.getCampaign().getName();
      var ag_name = kw.getAdGroup().getName();
      var imps_weighted_qs = (qs*imps);
      _loadEntityMap(account_score_map,account_id,imps_weighted_qs,imps);
      _loadEntityMap(camp_score_map,camp_name,imps_weighted_qs,imps);
      _loadEntityMap(ag_score_map,camp_name + "~~!~~" + ag_name,imps_weighted_qs,imps);
    }
    
    //Make sure the spreadsheet has all the sheets
    _addSpreadsheetSheets(SPREADSHEET_URL,['Account','Campaign','AdGroup']);
    
    //Load Account level QS
    Logger.log(SPREADSHEET_URL);
    
    var sheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL).getSheetByName('Account');
    _addHeadingsIfNeeded(sheet,['Date','Account','QS']);
    var e = account_score_map[account_id];
    try{
      sheet.appendRow([date_str,account_id,Math.round(e.imps_weighted_qs / e.tot_imps * SIG_FIGS)/SIG_FIGS]);
    }
    catch (e){
      sheet.insertRowAfter(sheet.getLastRow());
      sheet.appendRow([date_str,account_id,Math.round(e.imps_weighted_qs / e.tot_imps * SIG_FIGS)/SIG_FIGS]);
    }
      
    
    //Load Campaign level QS
    sheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL).getSheetByName('Campaign');
    _addHeadingsIfNeeded(sheet,['Date','Account','Campaign','QS']);
    var to_write = [];
    for(var i in camp_score_map) {
      var e = camp_score_map[i];
      to_write.push([date_str,
                     account_id,
                     i,
                     Math.round(e.imps_weighted_qs / e.tot_imps * SIG_FIGS)/SIG_FIGS]);
    }
    _writeDataToSheet(sheet,to_write);
    
    //Load Campaign + AdGroup level QS
    sheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL).getSheetByName('AdGroup');
    _addHeadingsIfNeeded(sheet,['Date','Account','Campaign','AdGroup','QS']);
    to_write = [];
    for(var i in ag_score_map) {
      var e = ag_score_map[i];
      to_write.push([date_str,
                     account_id,
                     i.split('~~!~~')[0],
                     i.split('~~!~~')[1],
                     Math.round(e.imps_weighted_qs / e.tot_imps * SIG_FIGS)/SIG_FIGS]);
    }
    _writeDataToSheet(sheet,to_write);
    Logger.log('Done processing %s', account.getCustomerId());
  }
}

// Super fast spreadsheet insertion
function _writeDataToSheet(sheet,to_write) {
  var last_row = sheet.getLastRow();
  var numRows = sheet.getMaxRows();
  if((numRows-last_row) < to_write.length) {
    sheet.insertRows(numRows,to_write.length-numRows+last_row); 
  }
  var range = sheet.getRange(last_row,1,to_write.length,to_write[0].length);
  range.setValues(to_write);
}
 
// Helper function to add the sheets  to the spreadsheet if needed
function _addSpreadsheetSheets(url,sheet_names) {
  var spreadsheet = SpreadsheetApp.openByUrl(url);
  var all_sheets = spreadsheet.getSheets();
  var all_sheet_names = [];
  for(var i in all_sheets) {
    all_sheet_names.push(all_sheets[i].getName());
  }
   
  for(var i in sheet_names) {
    var name = sheet_names[i];
    if(all_sheet_names.indexOf(name) == -1) {
      spreadsheet.insertSheet(name);
    } else {
      if(!APPEND) {
        spreadsheet.getSheetByName(name).clear();
      }
    }
  }
}
 
// Helper function to load the map for storing data
function _loadEntityMap(map,key,imps_weighted_qs,imps) {
  if(!map[key]) {
    map[key] = { imps_weighted_qs : imps_weighted_qs, tot_imps : imps };
  } else {
    map[key].imps_weighted_qs += imps_weighted_qs;
    map[key].tot_imps += imps;
  }
}
  
//Helper function to add headers to sheet if needed
function _addHeadingsIfNeeded(sheet,headings) {
  if(sheet.getRange('A1:A1').getValues()[0][0] == "") {
    sheet.clear();
    sheet.appendRow(headings);
  }
}
function selectSheet(currentAcc) {
  
  var spread = SpreadsheetApp.openByUrl("https://docs.google.com/a/fastpivot.com/spreadsheet/ccc?key=0Aou6MRJm21cydERGSDdmM2xqTTlQUDE1WGJWSWJjM3c#gid=0"); 
  var sheet = spread.getSheets()[0];
  var clientData = spread.getRangeByName("clientData");
  
  var clients = getRowsData(sheet, clientData);
  
  for(var i = 0; clients[i]; i++) {
    if (currentAcc = clients[i].clientId){
      SPREADSHEET_URL = clients[i].spreadsheetUrl;
      return SPREADSHEET_URL;
    }
  }
}

function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getLastColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), normalizeHeaders(headers));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

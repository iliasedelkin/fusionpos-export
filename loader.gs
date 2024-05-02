let sheetForExport = 'Экспорт'

// Default values for debugging
let defaultStartDate = '2024-01-01'

date = new Date()
date.setDate(date.getDate() - 1);
let defaultEndDate = date.toISOString().split('T')[0]

// Menu initialisation
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Загрузка данных')
    .addItem('Загрузить транзакции на экспорт', 'loadTransactionsForExport')
    .addToUi();
}

// Retrieve secret from the Script Properties or throw an error
function getScriptSecret(key) {
  let secret = PropertiesService.getScriptProperties().getProperty(key)
  if (!secret) throw Error(`Secret ${key} is empty`)
  return secret
}


function loadTransactionsForExport() {

const tok = getScriptSecret("Fusion_key")

let pointDict = getDict(tok, dictType = 'point')

let menuDict = getDict(tok, dictType = 'menu', query1 = 'name', query2 = 'category_id')
// Logger.log(menuDict)
let categoryDict =  getDict(tok, dictType = 'menucategory', query1 = 'category_name')
// let menuDict = getDict(dictType = 'menu', query2 = 'category_id')
// let categoryDict =  getDict(dictType = 'menucategory', query1 = 'category_name')
// Logger.log(categoryDict)

var dateStart = getDate('start')
var dateEnd = getDate('end')

if (dateStart !== 'too_soon') {
    
    Logger.log(`Fetching transactions from ${dateStart} to ${dateEnd}`)

    var metaJson = executeRequest(tok, dateStart, dateEnd, 1)
    var pageCount = metaJson.data._meta.pageCount
    var arrayToWrite = []

    Logger.log(`Pages to export: ${pageCount}`)

    if (pageCount !== 0) {

      for (i=1; i<=pageCount; i++) {
        json = executeRequest(tok, dateStart, dateEnd, i)
        var arrayForPage = prepareArray(json, pointDict, menuDict, categoryDict)
        // Logger.log(arrayForPage)
        arrayToWrite = arrayToWrite.concat(arrayForPage)
        // Logger.log(arrayToWrite);
      }

      // Logger.log(arrayToWrite)
      // Logger.log(arrayToWrite.length !== 0)

      if (arrayToWrite.length !== 0) {
      clearSheet()
      writeDataToSheet(arrayToWrite)
      Logger.log(`Written data from ${dateStart} to ${dateEnd}`) ;
      }
    
    } else {

      Logger.log (`No new transactions found`)

    }

  } else {

    Logger.log(`Start date invalid`)

  }
}

function getDict(tok, dictType, query1 = 'name', query2 = 0) {

  // Configure request
  var pageNum = 1

  var url = `https://mirvr.fusion24.ru:8443/v1/${dictType}`
  + '?page=' + pageNum

  var dict = {}

  try {
    // call the API
    var response = UrlFetchApp.fetch(url, {
      method: "GET", 
      headers: {Authorization:'Bearer '+tok}
    });

    var data = response.getContentText();
    var json = JSON.parse(data);
    var dictPageCount = json.data._meta.pageCount
    var arrayData = json.data.items
    if (query2 === 0) {
      dict = Object.assign({}, ...arrayData.map((x) => ({[x.id]: x[query1]})))
    } else {
      dict = Object.assign({}, ...arrayData.map((x) => ({[x.id]: [x[query1], x[query2]]})))
    }
    pageNum += 1

    for (let i = pageNum; i <=  dictPageCount; i++) {
      url = `https://mirvr.fusion24.ru:8443/v1/${dictType}`
      + '?page=' + pageNum
      try {
        // call the API
        var response = UrlFetchApp.fetch(url, {
          method: "GET", 
          headers: {Authorization:'Bearer '+tok}
        });

        var data = response.getContentText();
        var json = JSON.parse(data);
        var dictPageCount = json.data._meta.pageCount
        var arrayData = json.data.items
        if (query2 === 0) {
          dictToAdd = Object.assign({}, ...arrayData.map((x) => ({[x.id]: x[query1]})))
        } else {
          dictToAdd = Object.assign({}, ...arrayData.map((x) => ({[x.id]: [x[query1], x[query2]]})))
        }
        dict = Object.assign({}, dict, dictToAdd)
        pageNum += 1

      } catch (error) {
        // deal with any errors
        Logger.log(error);
      };
    }
  } catch (error) {
    // deal with any errors
    Logger.log(error);
  };
  
  // Logger.log(dict)
  return dict;
}

// Use this func to parse string to Datetime object
function isoToDate(dateStr){// argument = date string iso format
  var str = dateStr.replace(/-/,'/').replace(/-/,'/').replace(/T/,' ').replace(/\+/,' \+').replace(/Z/,' +00');
  return new Date(str);
}

function executeRequest(tok, dateStart = defaultStartDate, dateEnd = defaultEndDate, pageNum = 1) {

  // Configure request
  var filter = `{\"date_start\":\"${dateStart} 0:0:0\",\"date_end\":\"${dateEnd} 23:59:59\"}`;
  var fields = 'id,order_number,id_point,open_date,close_date,status,id_client,discount,id_discount,discount_name,type_payment,cost_price,waiterName,is_fiskal,comment,total_money,content';

  var url = 'https://mirvr.fusion24.ru:8443/v1/order'
  + '?sort=!open_date'
  + '&expand=content'
  + '&filter=' + encodeURIComponent(filter)
  + '&fields=' + fields
  + '&page=' + pageNum

  try {
    // call the API
    Logger.log(url)
    var response = UrlFetchApp.fetch(url, {
    method: "GET", 
    headers: {Authorization:'Bearer '+tok}
    }
  );

  var data = response.getContentText();
  var json = JSON.parse(data);
  }

  catch (error) {
  // deal with any errors
  Logger.log(error);
  };
  
  return json
}

function prepareArray(json, pointDict, menuDict, categoryDict) {
  // get data array
  Logger.log('Num items fetched: ' +json.data.items.length)
  var arrayForPage = []
  if (json.data.items.length !== 0) {
    var arrayData = json.data.items; 
    // Add the arrayProperties to the array 
    arrayData.forEach(function(el) {
      // IMPORTANT: If the order does not have any items or has status 'returned', it is omit
      // Logger.log(el)
      // Logger.log(el.content)
      // Logger.log(el.content.length)
      if (el.content.length != 0 && el.status != 'returned') {

        arrayToAppend = [el.id,        
                            el.order_number,
                            pointDict[el.id_point],
                            el.open_date,
                            el.close_date,
                            el.discount_name,
                            el.status,
                            el.waiterName,
                            el.comment] 
        
        el.content.forEach(function(item) {

          arrayForTransaction = [...arrayToAppend]

          arrayForTransaction.push(menuDict[item.menu_id][0])
          arrayForTransaction.push(categoryDict[menuDict[item.menu_id][1]])
          arrayForTransaction.push(item.menu_price)
          arrayForTransaction.push(item.discount_price)
          arrayForTransaction.push(item.discount_value)
          arrayForTransaction.push(item.menu_count)
          arrayForTransaction.push(item.cost_price)
          arrayForTransaction.push(item.discount_price * item.menu_count)

          arrayForPage.push(arrayForTransaction)

        })

      } 
    })

    // Logger.log(arrayForPage)
    return arrayForPage
    

  } else {

    return arrayForPage

  }
}

function writeDataToSheet(arrayToWrite) {

  // select the output sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = sheetForExport
  var sheet = ss.getSheetByName(`${sheetName}`);   

  // get last non-empty row to append data below it
  var lastRow = sheet.getLastRow();        
  
  // calculate the number of rows and columns needed
  var numRows = arrayToWrite.length;
  var numCols = arrayToWrite[0].length;
  
  // output the numbers to the sheet
  sheet.getRange(lastRow + 1,1,numRows,numCols).setValues(arrayToWrite)

}

function getDate(type = 'start') {
  // select the sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetName = sheetForExport
  let sheet = ss.getSheetByName(`${sheetName}`) 
  
  // select the column
  var columnNum
  if (type === 'start') {
    columnNum = 22
  } else if (type === 'end') {
    columnNum = 24
  }

  // get the date
  var cellValue = sheet.getRange(1, columnNum).getValue()
  // Logger.log(cellValue);

  // TODO: improve timeZone unification, current method of setting time to 12 is suboptimal
  cellValue.setHours(12)
  monthIndex = cellValue.getMonth()
  dateIndex = cellValue.getDate()
  // Logger.log(cellValue)
  // Logger.log(cellValue.getMonth())
  // Logger.log(cellValue.getDate());

  try {

    // Add one day to the date value
    let date = new Date()
    // TODO: this day-adding syntax must be improved for end-of-year / end-of-month cases
    date.setHours(12)
    date.setMonth(monthIndex, dateIndex) 
    // Logger.log(date);


    // Compare the date value with today. If the proposed request date is today, it is too soon to parse data
    today = new Date()
    // Logger.log(today.getTime())
    // Logger.log(date.getTime())
    // Logger.log(today.getTime() >= date.getTime())

    // Check if the read date is not later than current date
    if (today.getTime() >= date.getTime()) {
      // NOTE: the toISOString conversion may shift the date if the time of datetime object is very early in the day
      date.setHours(12)
      dateStr = date.toISOString().split('T')[0]
      // Logger.log(dateStr)

      return dateStr

    } else {
      // If end date is today or later, return yesterday
      if (type = 'end') {
        today.setHours(12)
        // Logger.log(today)
        // TODO: Check if this date subtraction works fine on the first day of the month
        today.setDate(today.getDate() - 1)
        // Logger.log(today)
        dateStr = today.toISOString().split('T')[0]
        // Logger.log(dateStr)

        return dateStr
      } else {
        // If start date is today or later, do not return any date
        // Logger.log('too_soon')
        return 'too_soon'
      }
    }
  }

  // If the date cannot be interpreted (probably because the sheet is empty) return default start date
  catch {
    Logger.log('error interpering the date')
    return defaultStartDate

  }
}

function clearSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(`${sheetForExport}`);   
  // Clear everything but the first row
  sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).clear();

}

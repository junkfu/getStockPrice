const privateKey = ''; //自行輸入api key
function onOpen() {
  var menu = SpreadsheetApp.getUi().createMenu('自定義程式');
      menu.addItem('取得價格', 'getPrice');
      menu.addToUi();  
}

function doPost(e){
  
  var postData = JSON.parse(e.postData.contents);
  var apikey = postData['apikey'];
  var output = ContentService.createTextOutput();

  if(apikey ==  privateKey){

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // 獲取當前活動工作表
    var stocks = sheet.getDataRange().getValues();
    var msg = '';
    for(var i in stocks){
      if(i==0)continue;
      var row =  stocks[i];
      msg = msg + row[2] + ' : ' + row[4] + '   '+(row[5] * 100).toFixed(2) + '%\r\n'; 
    }

      output.setContent(msg);
      output.setMimeType(ContentService.MimeType.JSON);
  }else {
    output.setContent(JSON.stringify({ "status": "apikey error" }));
      output.setMimeType(ContentService.MimeType.JSON);
  }

   return output;

}


//取得股價資料並插入表格內
function getPrice(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // 獲取當前活動工作表
  var stocks = sheet.getDataRange().getValues();
  //Logger.log(stocks);return;

  for (var i in stocks){
    
    if(i==0) continue;
    var row = stocks[i];
    var stork_string = row[0] + '_' + row[1] + '.tw';
    var price_data = callAPI(stork_string);
    var name = price_data['msgArray'][0]['n'];
    var open_price = price_data['msgArray'][0]['y'];
    var now_price = price_data['msgArray'][0]['z'];
    //Logger.log(open_price + '...' + now_price);
    sheet.getRange( parseInt(i)+1,3).setValue(name);
    sheet.getRange( parseInt(i)+1,4).setValue(open_price);
    sheet.getRange( parseInt(i)+1,5).setValue(now_price);

  }
}

//呼叫證交所API
function callAPI(stock_string) {
  var url = 'https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch='+stock_string+'&json=1&delay=100'; // 將這裡的URL替換為您想要呼叫的API的URL
  var response = UrlFetchApp.fetch(url); // 發送GET請求
  var json = response.getContentText(); // 獲取返回的內容
  var data = JSON.parse(json); // 解析JSON格式的數據
  Utilities.sleep(3000);// API有限制5秒內不得超過3次，故sleep
  return data;
  Logger.log(JSON.stringify(data)); // 在日誌中輸出獲取到的數據
}
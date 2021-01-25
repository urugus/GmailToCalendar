const CALENDAR_ID = '';
const SHEET_ID = '1h2BO91iDwKwATLXMIactVUBU2Hc0TadREQy8t5bPWwc';
const SHEET_NAME = 'Setting';

function gmailToCalendar(query) {
  const threads = GmailApp.search(query);

}

function arrayToDict(keys_array, values_array){
  //arrayはいずれも一次元配列
  dict = {}
  for(var i in keys_array){
    key = keys_array[i]
    value = Utilities.formatString(values_array[i]);
    if(key != ""){
      dict[key] = value
    }
  }
  return dict
}


function main(){

  const setting_sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const columns_array = setting_sheet.getRange('1:1').getValues();
  const data_array = setting_sheet.getDataRange().getValues();
  // 先頭行（項目名）を削除
  data_array.shift();
  // Logger.log(data_array);

  for(var i in data_array){
    var data_dict = arrayToDict(columns_array[0],data_array[i]);
    Logger.log(data_dict);
    const threads = GmailApp.search(data_dict['Search']);
    for(var j in threads){
      var thread = threads[j];
      var messages = GmailApp.getMessagesForThread(thread);
      var message = messages[0];
      Logger.log(setting_sheet.getRange("C2").getValue() == "A(.*?)\\d\\d\\d");
      var subject_reg = new RegExp("A(.*?)\\d\\d\\d");
      // Logger.log(reg);
      var subject = message.getBody().match(subject_reg);
      Logger.log(subject);
    }
  }

}

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
    var key = keys_array[i];
    var value = Utilities.formatString(values_array[i]);
    if(key != ""){
      dict[key] = value;
    }
  }
  return dict;
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
      var cal_subject = message.getBody().match(new RegExp(data_dict['Title']))[0];
      var cal_body = message.getBody().match(new RegExp(data_dict['Body'],'g'))[0];
      var start_time = message.getBody().match(new RegExp(data_dict['Date'],'g'))[0].replace(/-/g, '/').replace(/\([\s\S]\)/g, '');
      var time = Number(message.getBody().match(new RegExp(data_dict['Time'],'g'))[0].replace(/[^0-9]/g, ''));
      var cal_start_time = new Date(start_time);
      Logger.log(cal_start_time);
      var cal_end_time = new Date(cal_start_time.setMinutes(cal_start_time.getMinutes()+time));
      Logger.log(cal_end_time);
      var cal_option = {description: cal_body};
      // CalendarApp.createEvent(cal_subject, cal_start_time, cal_end_time,cal_option);
    }
  }

}

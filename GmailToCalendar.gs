const CALENDAR_ID = '';
const SHEET_ID = '1h2BO91iDwKwATLXMIactVUBU2Hc0TadREQy8t5bPWwc';
const SHEET_NAME = 'Setting';

const CALENDAR_COLOR = {
  'PALE_BLUE':'1',
  'PALE_GREEN':'2',
  'MAUVE':'3',
  'PALE_RED':'4',
  'YELLOW':'5',
  'ORANGE':'6',
  'CYAN':'7',
  'GRAY':'8',
  'BLUE':'9',
  'GREEN':'10',
  'RED':'11',
}


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
    // Logger.log(data_dict);
    const gmail_label_name = data_dict['GmailLabel'];
    const threads = GmailApp.search('label:'+ gmail_label_name +" "+ data_dict['Search']);
    for(var j in threads){
      var thread = threads[j];
      var messages = GmailApp.getMessagesForThread(thread);
      var message = messages[0];
      var cal_subject = message.getBody().match(new RegExp(data_dict['Title']))[0];
      var cal_body = message.getBody().match(new RegExp(data_dict['Body'],'g'))[0];
      var start_time = message.getBody().match(new RegExp(data_dict['Date'],'g'))[0].replace(/-/g, '/').replace(/\([\s\S]\)/g, '');
      var time = Number(message.getBody().match(new RegExp(data_dict['Time'],'g'))[0].replace(/[^0-9]/g, ''));
      var cal_start_time = new Date(start_time);
      var cal_end_time = new Date(start_time);
      cal_end_time.setMinutes(cal_start_time.getMinutes()+time);
      var cal_location = message.getBody().match(new RegExp(data_dict['Location'],'g'))[0];
      var cal_visibility = data_dict['Visibility '];
      var cal_color = CALENDAR_COLOR[data_dict['Color']];
      var cal_option = {
        description: cal_body,
        location: cal_location
      };
      // 公開設定によって分岐
      if(cal_visibility == 'Close'){
        var visibility = CalendarApp.Visibility.PRIVATE;
      }else{
        var visibility = CalendarApp.Visibility.PUBLIC;
      }
      CalendarApp.createEvent(cal_subject, cal_start_time, cal_end_time　,cal_option).setVisibility(visibility).setColor(cal_color);
      // カレンダー予定作成後、メールのラベルを外す（完了フラグとして）
      var gmail_label = GmailApp.getUserLabelByName(gmail_label_name);
      thread.removeLabel(gmail_label);
    }
  }

}

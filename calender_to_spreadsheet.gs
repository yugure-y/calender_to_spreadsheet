function calender_to_spreadsheet() {
  var sheet = SpreadsheetApp.getActiveSheet(); // シートを取得
  var no = 1; //No

  var pattern_leader = new RegExp('責任者.*(参加者)*', 'u');
  var pattern_member = new RegExp('参加者.*(業務内容)*', 'u');
  var pattern_detail = new RegExp('業務内容.*(備考)*', 'u');
  var pattern_extra = new RegExp('備考.*', 'u');
 
  // 期間を指定する
  var startDate = new Date('2022/04/01 00:00:00'); // 取得開始日
  var endDate = new Date('2023/03/31 00:00:00'); // 取得終了日
  
  // アカウントに紐付けられているすべてのカレンダーを取得
  var calendars = CalendarApp.getAllCalendars();
  for (var i in calendars) {
    // 対象オブジェクト毎の指定期間内の予定を取得
    var events = calendars[i].getEvents(startDate,endDate);

    // カレンダー毎のイベントを繰り返し最終行に追加
    events.forEach( function(evt) {
      var desc = String(evt.getDescription())
      if (desc.match(pattern_leader) != null) {
        var leader = desc.match(pattern_leader)[0]
        leader = leader.replace('<br>','').replace('参加者','').replace('責任者：','').replace('責任者:','')
      }
      if (desc.match(pattern_member) != null) {
        var member = desc.match(pattern_member)[0]
        member = member.replace('<br>','').replace('業務内容','').replace('参加者：','').replace('参加者:','')
      }
      if (desc.match(pattern_detail) != null) {
        var detail = desc.match(pattern_detail)[0]
        detail = detail.replace('<br>','').replace('備考','').replace('業務内容：','').replace('業務内容:','')
      }
      if (desc.match(pattern_extra) != null) {
        var extra = desc.match(pattern_extra)[0]
        extra = extra.replace('<br>','').replace('<u>','').replace('備考：','').replace('備考:','')
      }
  
      sheet.appendRow(
      [
        no, //No
        calendars[i].getName(),
        evt.getTitle(), // イベントタイトル
        evt.getStartTime(), // イベントの開始時刻
        evt.getEndTime(), // イベントの終了時刻
        leader,
        member,
        detail,
        extra
      ]
      );
      no++;
    });
  }
}

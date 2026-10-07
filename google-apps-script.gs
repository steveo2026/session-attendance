function doGet(){return ContentService.createTextOutput('Session Attendance Google Sheets endpoint is ready.').setMimeType(ContentService.MimeType.TEXT);}
function doPost(e){
  try{
    var data=JSON.parse(e.postData.contents||'{}');
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sh=ss.getSheetByName('Attendance')||ss.insertSheet('Attendance');
    var headers=['Player','Session Date','Attendance','Session Cost (£)','Amount Owed (£)','Session Note','Player Note'];
    sh.clearContents();sh.getRange(1,1,1,headers.length).setValues([headers]);
    var rows=[];var players=data.players||[],sessions=data.sessions||[];
    sessions.sort(function(a,b){return String(a.date).localeCompare(String(b.date));}).forEach(function(s){
      var a=s.attendance||{};
      players.forEach(function(p){var x=a[p.id]||{status:'no',note:''};var cost=Number(s.cost||0);rows.push([p.name,s.date,x.status==='yes'?'Attended':'Not attended',cost,x.status==='yes'?cost:0,s.note||'',x.note||'']);});
    });
    if(rows.length)sh.getRange(2,1,rows.length,headers.length).setValues(rows);
    sh.autoResizeColumns(1,headers.length);
    var sum=ss.getSheetByName('Player Summary')||ss.insertSheet('Player Summary');
    sum.clearContents();sum.getRange(1,1,1,3).setValues([['Player','Sessions Attended','Total Owed (£)']]);
    var summary=players.map(function(p){var n=0,total=0;sessions.forEach(function(s){var x=(s.attendance||{})[p.id];if(x&&x.status==='yes'){n++;total+=Number(s.cost||0);}});return [p.name,n,total];});
    if(summary.length)sum.getRange(2,1,summary.length,3).setValues(summary);sum.autoResizeColumns(1,3);
    return ContentService.createTextOutput(JSON.stringify({ok:true,rows:rows.length})).setMimeType(ContentService.MimeType.JSON);
  }catch(err){return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)})).setMimeType(ContentService.MimeType.JSON);}
}

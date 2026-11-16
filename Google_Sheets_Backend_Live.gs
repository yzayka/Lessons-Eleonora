/* English Studio answer storage. Paste into a spreadsheet-bound Apps Script project. */
function setup() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!ss) throw new Error('Open Apps Script from a Google Sheet: Extensions > Apps Script.');
  var p=PropertiesService.getScriptProperties();
  p.setProperty('SHEET_ID',ss.getId());
  if(!p.getProperty('ACCESS_KEY')) p.setProperty('ACCESS_KEY',Utilities.getUuid());
  ensureSheet_(ss,'Snapshots',['Client ID','Student','Lesson','Updated','State JSON']);
  ensureSheet_(ss,'Answers',['Client ID','Student','Lesson','Question ID','Question','Answer','Checked','Model shown','Updated']);
  Logger.log('Access key: '+p.getProperty('ACCESS_KEY'));
  Logger.log('Spreadsheet: '+ss.getUrl());
}
function ensureSheet_(ss,name,headers){var s=ss.getSheetByName(name)||ss.insertSheet(name);if(s.getLastRow()===0){s.appendRow(headers);s.setFrozenRows(1);s.getRange(1,1,1,headers.length).setBackground('#192743').setFontColor('#ffffff').setFontWeight('bold');}return s;}
function safe_(v){var s=String(v==null?'':v);return /^[=+@\-]/.test(s)?"'"+s:s;}
function output_(data,cb){var text=JSON.stringify(data);if(cb&&/^ppack_[a-zA-Z0-9_]+$/.test(cb))return ContentService.createTextOutput(cb+'('+text+');').setMimeType(ContentService.MimeType.JAVASCRIPT);return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);}
function doGet(e){var p=PropertiesService.getScriptProperties();if(e.parameter.key!==p.getProperty('ACCESS_KEY'))return output_({ok:false,error:'Invalid access key'},e.parameter.callback);if(e.parameter.action==='read'){
 var student=String(e.parameter.student||''),lesson=String(e.parameter.lesson||'Present Perfect Simple');
 if(!student||student.length>120)return output_({ok:false,error:'Missing student'},e.parameter.callback);
 var ss=SpreadsheetApp.openById(p.getProperty('SHEET_ID')),sh=ss.getSheetByName('Snapshots');
 if(!sh||sh.getLastRow()<2)return output_({ok:true,found:false},e.parameter.callback);
 var rows=sh.getDataRange().getValues(),latest=null,stamp=0;
 for(var i=1;i<rows.length;i++){
  if(String(rows[i][1])!==student||String(rows[i][2])!==lesson)continue;
  var time=new Date(rows[i][3]).getTime();
  if(!latest||time>stamp){latest=rows[i];stamp=time;}
 }
 if(!latest)return output_({ok:true,found:false},e.parameter.callback);
 var snapshot;try{snapshot=JSON.parse(latest[4]);}catch(err){return output_({ok:false,error:'Invalid stored snapshot'},e.parameter.callback);}
 // Do not return private cloud credentials saved by older clients.
 delete snapshot.cloudConfig;
 return output_({ok:true,found:true,updatedAt:stamp,state:snapshot},e.parameter.callback);
}
var id=e.parameter.requestId;if(!/^[\w-]{1,100}$/.test(id||''))return output_({ok:false,error:'Invalid request'},e.parameter.callback);var stored=CacheService.getScriptCache().get('ack:'+id);return output_(stored?JSON.parse(stored):{pending:true},e.parameter.callback);}
function doPost(e){var data,id,lock=LockService.getScriptLock();
 try{
  data=JSON.parse(e.postData.contents);id=data.requestId;
  if(data.key!==PropertiesService.getScriptProperties().getProperty('ACCESS_KEY'))return output_({ok:false,error:'Invalid access key'});
  if(!/^[\w-]{1,100}$/.test(id||'')||!data.clientId||!data.student)throw new Error('Missing identity');
  var snapshot=JSON.stringify(data.state);if(snapshot.length>45000)throw new Error('Answer snapshot exceeds the sheet cell limit. Export your local copy.');
  lock.waitLock(25000);
  var ss=SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  var sh=ensureSheet_(ss,'Snapshots',['Client ID','Student','Lesson','Updated','State JSON']);
  var rows=sh.getDataRange().getValues(),idx=-1;
  for(var i=1;i<rows.length;i++)if(String(rows[i][0])===data.clientId){idx=i+1;break;}
  var old=idx>0?new Date(rows[idx-1][3]).getTime():0;
  if(old<=data.updatedAt){
   sh.getRange(idx>0?idx:sh.getLastRow()+1,1,1,5).setValues([[safe_(data.clientId),safe_(data.student),safe_(data.lesson),new Date(data.updatedAt),snapshot]]);
   var ans=ensureSheet_(ss,'Answers',['Client ID','Student','Lesson','Question ID','Question','Answer','Checked','Model shown','Updated']);
   var existing=ans.getDataRange().getValues(),map={};for(var j=1;j<existing.length;j++)map[String(existing[j][0])+'|'+String(existing[j][3])]=j+1;
   (data.entries||[]).forEach(function(q){var record=[safe_(data.clientId),safe_(data.student),safe_(data.lesson),safe_(q.id),safe_(q.prompt),safe_(q.answer),!!q.checked,!!q.revealed,new Date(data.updatedAt)];var row=map[data.clientId+'|'+q.id]||ans.getLastRow()+1;ans.getRange(row,1,1,record.length).setValues([record]);});
  }
  CacheService.getScriptCache().put('ack:'+id,JSON.stringify({ok:true,requestId:id,updatedAt:data.updatedAt}),600);
  return output_({ok:true});
 }catch(err){if(id&&data&&data.key===PropertiesService.getScriptProperties().getProperty('ACCESS_KEY'))CacheService.getScriptCache().put('ack:'+id,JSON.stringify({ok:false,error:String(err.message)}),600);return output_({ok:false,error:String(err.message)});
 }finally{if(lock.hasLock())lock.releaseLock();}
}
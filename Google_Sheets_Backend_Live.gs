/* Present Perfect: Google Sheets is the single source of truth.
   Spreadsheet-bound Apps Script. Run setup() once, then deploy Web app:
   Execute as Me; Who has access Anyone. Update existing deployment with New version.
   Public lesson link must carry the access token; do not put the token in HTML.
*/
var LESSON_='Present Perfect Simple';
function setup(){
 var ss=SpreadsheetApp.getActiveSpreadsheet();
 if(!ss)throw Error('Open from a Google Sheet: Extensions > Apps Script.');
 var props=PropertiesService.getScriptProperties();
 props.setProperty('SHEET_ID',ss.getId());
 if(!props.getProperty('ACCESS_KEY'))props.setProperty('ACCESS_KEY',Utilities.getUuid());
 sheet_('Snapshots',['Client ID','Student','Lesson','Updated','State JSON']);
 sheet_('Answers',['Client ID','Student','Lesson','Question ID','Question','Answer','Checked','Model shown','Updated']);
 Logger.log('Setup complete. Keep the access key private.');
}
function sheet_(name,headers){
 var id=PropertiesService.getScriptProperties().getProperty('SHEET_ID');
 var ss=id?SpreadsheetApp.openById(id):SpreadsheetApp.getActiveSpreadsheet();
 var sh=ss.getSheetByName(name)||ss.insertSheet(name);
 if(sh.getLastRow()===0){sh.appendRow(headers);sh.setFrozenRows(1);}
 return sh;
}
function out_(payload,callback){
 var json=JSON.stringify(payload);
 if(callback && /^ppack_[a-zA-Z0-9_]+$/.test(callback))
  return ContentService.createTextOutput(callback+'('+json+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
 return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
function authorized_(key){
 var known=PropertiesService.getScriptProperties().getProperty('ACCESS_KEY');
 return !!known && String(key||'')===known;
}
function read_(student,lesson){
 var sh=sheet_('Snapshots',['Client ID','Student','Lesson','Updated','State JSON']);
 var rows=sh.getDataRange().getValues(),best=null,when=-1,row=-1;
 for(var i=1;i<rows.length;i++){
  if(String(rows[i][1])!==student||String(rows[i][2])!==lesson)continue;
  var timestamp=new Date(rows[i][3]).getTime();
  if(!best||timestamp>when){best=rows[i];when=timestamp;row=i+1;}
 }
 if(!best)return {found:false,revision:0,state:{answers:{},open:{},ticks:{},checked:{},revealed:{},done:[]},row:-1};
 var data=JSON.parse(best[4]||'{}');
 delete data.cloudConfig;
 return {found:true,revision:Math.max(0,Number(data._revision)||0),updatedAt:when,state:data,row:row};
}
function doGet(e){
 var p=(e&&e.parameter)||{},cb=p.callback;
 try{
  if(!authorized_(p.key))return out_({ok:false,error:'Invalid access key'},cb);
  if(p.action==='read'){
   var student=String(p.student||'').slice(0,120),lesson=String(p.lesson||LESSON_).slice(0,120);
   if(!student)return out_({ok:false,error:'Missing student'},cb);
   var data=read_(student,lesson);
   return out_({ok:true,found:data.found,revision:data.revision,updatedAt:data.updatedAt||0,state:data.state},cb);
  }
  var id=String(p.requestId||'');
  if(!/^[\w-]{1,100}$/.test(id))return out_({ok:false,error:'Invalid request'},cb);
  var result=CacheService.getScriptCache().get('ack:'+id);
  return out_(result?JSON.parse(result):{pending:true},cb);
 }catch(err){return out_({ok:false,error:String(err.message)},cb);}
}
function cell_(value){
 var s=String(value==null?'':value);
 return /^[=+@\-]/.test(s)?"'"+s:s;
}
function doPost(e){
 var data,requestId='',lock=LockService.getScriptLock();
 try{
  data=JSON.parse(e.postData.contents);
  requestId=String(data.requestId||'');
  if(!authorized_(data.key))return out_({ok:false,error:'Invalid access key'});
  if(!/^[\w-]{1,100}$/.test(requestId)||!data.student||!data.patch)throw Error('Invalid request');
  var student=String(data.student).slice(0,120),lesson=String(data.lesson||LESSON_).slice(0,120);
  lock.waitLock(25000);
  var existing=read_(student,lesson),state=existing.state;
  var patch=data.patch;
  ['answers','open','ticks','checked','revealed'].forEach(function(group){
   if(!patch[group]||typeof patch[group]!=='object')return;
   if(!state[group]||typeof state[group]!=='object')state[group]={};
   Object.keys(patch[group]).forEach(function(id){
    if(id==='__proto__'||id==='constructor'||id==='prototype')return;
    if(patch[group][id]===null)delete state[group][id];
    else state[group][id]=patch[group][id];
   });
  });
  if(Array.isArray(patch.done))state.done=patch.done.slice(0,100);
  if(typeof patch.storyTopic==='string')state.storyTopic=patch.storyTopic;
  state._revision=existing.revision+1;
  state.updatedAt=Date.now();
  delete state.cloudConfig;
  var json=JSON.stringify(state);
  if(json.length>45000)throw Error('Snapshot exceeds Sheets cell limit');
  var sh=sheet_('Snapshots',['Client ID','Student','Lesson','Updated','State JSON']);
  var target=existing.row>0?existing.row:sh.getLastRow()+1;
  var ownerId=existing.row>0?sh.getRange(target,1).getValue():'shared-'+student;
  sh.getRange(target,1,1,5).setValues([[cell_(ownerId),cell_(student),cell_(lesson),new Date(state.updatedAt),json]]);
  var ans=sheet_('Answers',['Client ID','Student','Lesson','Question ID','Question','Answer','Checked','Model shown','Updated']);
  var rows=ans.getDataRange().getValues(),index={};
  for(var i=1;i<rows.length;i++)
   if(String(rows[i][1])===student&&String(rows[i][2])===lesson)index[String(rows[i][3])]=i+1;
  var updated={};
  ['answers','open'].forEach(function(group){
   Object.keys(patch[group]||{}).forEach(function(k){updated[(group==='open'?'open:':'')+k]=true;});
  });
  Object.keys(updated).forEach(function(id){
   var open=id.indexOf('open:')===0,k=open?id.slice(5):id;
   var value=open?(state.open||{})[k]:(state.answers||{})[k];
   var row=index[id]||ans.getLastRow()+1;
   var record=[cell_(ownerId),cell_(student),cell_(lesson),cell_(id),cell_(id),cell_(value),!!(state.checked||{})[k],!!(state.revealed||{})[k],new Date(state.updatedAt)];
   ans.getRange(row,1,1,9).setValues([record]);
   index[id]=row;
  });
  var ack={ok:true,revision:state._revision,requestId:requestId};
  CacheService.getScriptCache().put('ack:'+requestId,JSON.stringify(ack),600);
  return out_(ack);
 }catch(err){
  var fail={ok:false,error:String(err.message)};
  if(requestId&&data&&authorized_(data.key))CacheService.getScriptCache().put('ack:'+requestId,JSON.stringify(fail),600);
  return out_(fail);
 }finally{if(lock.hasLock())lock.releaseLock();}
}

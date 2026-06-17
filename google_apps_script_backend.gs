
const SHEET_ID = "1R-cjiNMoNWv-BhokRXeeBCmHtssr4-LMQwYGTd39g-0";
const FOLDER_ID = "13lzW_JzwCiLaYXDGWV77Muo90FjNV5au";

function getSheet_(){return SpreadsheetApp.openById(SHEET_ID).getSheets()[0];}
function out_(obj, cb){
  var s = cb ? cb + "(" + JSON.stringify(obj) + ");" : JSON.stringify(obj);
  return ContentService.createTextOutput(s).setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}
function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
function headers_(sheet){return sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(),1)).getValues()[0].map(String);}
function col_(h,names){for(var i=0;i<names.length;i++){var p=h.indexOf(names[i]);if(p>=0)return p+1;}return -1;}
function ensureCols_(){
  var sheet=getSheet_();
  if(sheet.getLastRow()===0)return;
  var h=headers_(sheet);
  ["Transcript","Transcript Status","Transcribed By","Transcription Date","Code","Theme","Auto Suggested","Sample Quote","Interpretation","Researcher Review"].forEach(function(n){
    if(h.indexOf(n)<0){sheet.getRange(1,sheet.getLastColumn()+1).setValue(n);h.push(n);}
  });
}
function ensureSubmissionHeaders_(sheet){
  var hs=["Submitted At","Started At","Duration Seconds","Project ID","Project Title","Institution","Researcher","Researcher Contact","Survey Title","Respondent ID","Enumerator ID","District","Latitude","Longitude","GPS Accuracy","GPS Timestamp","Consent","Question ID","Question Text","Question Audio","Answer File","Drive URL","File ID","Client Timestamp","Notes","Transcript","Transcript Status","Transcribed By","Transcription Date","Code","Theme","Auto Suggested","Sample Quote","Interpretation","Researcher Review"];
  if(sheet.getLastRow()===0){sheet.appendRow(hs);return;}
  var h=headers_(sheet);
  hs.forEach(function(n){if(h.indexOf(n)<0){sheet.getRange(1,sheet.getLastColumn()+1).setValue(n);h.push(n);}});
}
function doGet(e){
  var cb=e&&e.parameter&&(e.parameter.callback||e.parameter.cb);
  try{
    var action=e&&e.parameter&&e.parameter.action;
    if(action==="stats")return out_(stats_(),cb);
    if(action==="transcriptionStatus")return out_({status:"success",version:"V20.2",engine:"OpenAI Whisper",apiKeyConfigured:!!PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY")},cb);
    return out_({status:"success",message:"Audio Research Backend V20.2 running"},cb);
  }catch(err){return out_({status:"error",message:String(err)},cb);}
}
function doPost(e){
  try{
    var data=JSON.parse(e.postData.contents);
    if(data.action==="transcribeAudio")return transcribeAudio_(data);
    if(data.action==="updateTranscript")return updateTranscript_(data);
    var folder=DriveApp.getFolderById(FOLDER_ID), sheet=getSheet_();
    ensureSubmissionHeaders_(sheet);
    var h=headers_(sheet), map={}; h.forEach(function(x,i){map[x]=i+1;});
    var fileMap={};
    (data.files||[]).forEach(function(file){
      var blob=Utilities.newBlob(Utilities.base64Decode(file.base64),file.mimeType||"audio/webm",file.fileName);
      var saved=folder.createFile(blob);
      fileMap[file.fileName]={url:saved.getUrl(),id:saved.getId()};
    });
    (data.metadata||[]).forEach(function(row){
      var f=fileMap[row.answerFileName]||{}, arr=new Array(sheet.getLastColumn()).fill("");
      function set(k,v){if(map[k])arr[map[k]-1]=v;}
      set("Submitted At",data.submittedAt); set("Started At",data.startedAt||""); set("Duration Seconds",data.durationSeconds||"");
      set("Project ID",data.projectId||""); set("Project Title",data.projectTitle||""); set("Institution",data.institution||"");
      set("Researcher",data.researcherName||""); set("Researcher Contact",data.researcherContact||""); set("Survey Title",data.surveyTitle||"");
      set("Respondent ID",data.respondentId||""); set("Enumerator ID",data.enumeratorId||row.enumeratorId||""); set("District",data.district||row.district||"");
      set("Latitude",data.latitude||row.latitude||""); set("Longitude",data.longitude||row.longitude||""); set("GPS Accuracy",data.gpsAccuracy||row.gpsAccuracy||"");
      set("GPS Timestamp",data.gpsTimestamp||row.gpsTimestamp||""); set("Consent",data.consent||"");
      set("Question ID",row.questionId); set("Question Text",row.questionText); set("Question Audio",row.questionAudio);
      set("Answer File",row.answerFileName); set("Drive URL",f.url||""); set("File ID",f.id||"");
      set("Client Timestamp",row.timestamp); set("Notes",data.notes||row.notes||""); set("Transcript Status","Pending");
      sheet.appendRow(arr);
    });
    return json_({status:"success",rows:(data.metadata||[]).length});
  }catch(err){return json_({status:"error",message:String(err)});}
}
function stats_(){
  var sheet=getSheet_(); ensureSubmissionHeaders_(sheet);
  var values=sheet.getDataRange().getValues(), h=values[0].map(String);
  function ix(n){return h.indexOf(n);}
  var ri=ix("Respondent ID"), qi=ix("Question ID"), ei=ix("Enumerator ID"), di=ix("District"), lati=ix("Latitude"), lngi=ix("Longitude"), afi=ix("Answer File"), dui=ix("Drive URL"), fidi=ix("File ID"), tri=ix("Transcript"), tsi=ix("Transcript Status"), thi=ix("Theme"), coi=ix("Code"), qui=ix("Sample Quote"), ini=ix("Interpretation");
  var respondents={},questions={},progress={},gps={},enums={},dists={},byEnum={},byDist={},manifest=[],themeMap={},codeMap={},answers=0,transcribed=0,pending=0,failed=0;
  for(var r=1;r<values.length;r++){
    var row=values[r], resp=ri>=0?String(row[ri]||"").trim():"", q=qi>=0?String(row[qi]||"").trim():"", en=ei>=0?String(row[ei]||"").trim():"", dist=di>=0?String(row[di]||"").trim():"";
    if(resp)respondents[resp]=true; if(q)questions[q]=true;
    if(resp&&q){answers++; if(!progress[resp])progress[resp]={}; progress[resp][q]=true;}
    if(lati>=0&&lngi>=0&&row[lati]&&row[lngi]&&resp)gps[resp]=true;
    if(en){enums[en]=true;if(resp){if(!byEnum[en])byEnum[en]={};byEnum[en][resp]=true;}}
    if(dist){dists[dist]=true;if(resp){if(!byDist[dist])byDist[dist]={};byDist[dist][resp]=true;}}
    var transcript=tri>=0?String(row[tri]||"").trim():"", st=tsi>=0?String(row[tsi]||"").trim():""; if(!st)st=transcript?"Transcribed":"Pending";
    if(afi>=0&&row[afi]){
      if(transcript||st==="Transcribed")transcribed++; else if(st==="Failed"||st==="Error")failed++; else pending++;
      manifest.push({respondentId:resp,questionId:q,answerFile:String(row[afi]||""),driveUrl:dui>=0?String(row[dui]||""):"",fileId:fidi>=0?String(row[fidi]||""):"",transcriptStatus:st});
    }
    var th=thi>=0?String(row[thi]||"").trim():"", co=coi>=0?String(row[coi]||"").trim():"";
    if(th){if(!themeMap[th])themeMap[th]={theme:th,mentions:0,sampleQuote:"",interpretation:""};themeMap[th].mentions++;if(!themeMap[th].sampleQuote&&qui>=0)themeMap[th].sampleQuote=String(row[qui]||"");if(!themeMap[th].interpretation&&ini>=0)themeMap[th].interpretation=String(row[ini]||"");}
    if(co)codeMap[co]=(codeMap[co]||0)+1;
  }
  var respondentCount=Object.keys(respondents).length, questionCount=Object.keys(questions).length, fully=Object.keys(progress).filter(function(k){return Object.keys(progress[k]).length>=questionCount;}).length;
  var prog={}; Object.keys(progress).forEach(function(k){var n=Object.keys(progress[k]).length, key=n+" of "+questionCount; prog[key]=(prog[key]||0)+1;});
  var enumOut={}; Object.keys(byEnum).forEach(function(k){enumOut[k]=Object.keys(byEnum[k]).length;});
  var distOut={}; Object.keys(byDist).forEach(function(k){distOut[k]=Object.keys(byDist[k]).length;});
  return {status:"success",respondents:respondentCount,answers:answers,questions:questionCount,completion:respondentCount?Math.round(fully/respondentCount*100):0,gpsCaptured:Object.keys(gps).length,fullyAnswered:fully,enumerators:Object.keys(enums).length,districts:Object.keys(dists).length,byEnumerator:enumOut,byDistrict:distOut,progressSummary:prog,audioManifest:manifest,transcription:{audioRecords:manifest.length,transcribed:transcribed,pending:pending,failed:failed,coverage:manifest.length?Math.round(transcribed/manifest.length*100):0},themeSummary:Object.keys(themeMap).map(function(k){return themeMap[k];}),codeSummary:Object.keys(codeMap).map(function(k){return {code:k,mentions:codeMap[k]};}),updatedAt:new Date().toISOString()};
}
function updateTranscript_(data){
  ensureCols_(); var sheet=getSheet_(), h=headers_(sheet);
  var respCol=col_(h,["Respondent ID","Respondent"]), qCol=col_(h,["Question ID","Question"]), fileCol=col_(h,["File ID"]), answerCol=col_(h,["Answer File"]), trCol=col_(h,["Transcript"]), stCol=col_(h,["Transcript Status"]), byCol=col_(h,["Transcribed By"]), dtCol=col_(h,["Transcription Date"]);
  var tf=String(data.fileId||"").trim(), ta=String(data.answerFile||"").trim(), tr=String(data.respondentId||"").trim(), tq=String(data.questionId||"").trim(), updated=0, matchedBy="";
  for(var r=2;r<=sheet.getLastRow();r++){
    var rf=fileCol>0?String(sheet.getRange(r,fileCol).getValue()||"").trim():"", ra=answerCol>0?String(sheet.getRange(r,answerCol).getValue()||"").trim():"", rr=respCol>0?String(sheet.getRange(r,respCol).getValue()||"").trim():"", rq=qCol>0?String(sheet.getRange(r,qCol).getValue()||"").trim():"";
    var match=false; if(tf&&rf&&tf===rf){match=true;matchedBy="File ID";} else if(ta&&ra&&ta===ra){match=true;matchedBy="Answer File";} else if(tr&&tq&&rr===tr&&String(rq)===String(tq)){match=true;matchedBy="Respondent ID + Question ID";}
    if(match){sheet.getRange(r,trCol).setValue(data.transcript||""); if(stCol>0)sheet.getRange(r,stCol).setValue("Transcribed"); if(byCol>0)sheet.getRange(r,byCol).setValue(data.transcribedBy||"OpenAI Whisper"); if(dtCol>0)sheet.getRange(r,dtCol).setValue(data.transcriptionDate||new Date().toISOString()); updated++;}
  }
  return json_({status:"success",updated:updated,matchedBy:matchedBy});
}
function transcribeAudio_(data){
  try{
    var key=PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
    if(!key)return json_({status:"error",message:"OPENAI_API_KEY not configured"});
    if(!data.fileId)return json_({status:"error",message:"fileId missing"});
    var file=DriveApp.getFileById(data.fileId), blob=file.getBlob(), boundary="----ARDC"+Date.now();
    var part1="--"+boundary+"\r\nContent-Disposition: form-data; name=\"model\"\r\n\r\nwhisper-1\r\n--"+boundary+"\r\nContent-Disposition: form-data; name=\"file\"; filename=\""+file.getName()+"\"\r\nContent-Type: "+blob.getContentType()+"\r\n\r\n", part2="\r\n--"+boundary+"--";
    var payload=Utilities.newBlob(part1).getBytes().concat(blob.getBytes()).concat(Utilities.newBlob(part2).getBytes());
    var res=UrlFetchApp.fetch("https://api.openai.com/v1/audio/transcriptions",{method:"post",headers:{Authorization:"Bearer "+key},contentType:"multipart/form-data; boundary="+boundary,payload:payload,muteHttpExceptions:true});
    if(res.getResponseCode()<200||res.getResponseCode()>=300)return json_({status:"error",message:"OpenAI error: "+res.getContentText()});
    var transcript=JSON.parse(res.getContentText()).text||"";
    var save=JSON.parse(updateTranscript_({respondentId:data.respondentId,questionId:data.questionId,fileId:data.fileId,answerFile:data.answerFile,transcript:transcript,transcribedBy:"OpenAI Whisper",transcriptionDate:new Date().toISOString()}).getContent());
    if(!save.updated)return json_({status:"error",message:"Transcript created but no sheet row updated",transcript:transcript,updated:0});
    return json_({status:"success",transcript:transcript,updated:save.updated,matchedBy:save.matchedBy});
  }catch(err){return json_({status:"error",message:String(err)});}
}

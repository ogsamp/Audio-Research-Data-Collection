const SHEET_ID="1R-cjiNMoNWv-BhokRXeeBCmHtssr4-LMQwYGTd39g-0";
const FOLDER_ID="13lzW_JzwCiLaYXDGWV77Muo90FjNV5au";

function getSheet_(){return SpreadsheetApp.openById(SHEET_ID).getSheets()[0];}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function out_(o,cb){var s=cb?cb+"("+JSON.stringify(o)+");":JSON.stringify(o);return ContentService.createTextOutput(s).setMimeType(cb?ContentService.MimeType.JAVASCRIPT:ContentService.MimeType.JSON);}
function headers_(sh){return sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getValues()[0].map(String);}
function col_(h,n){for(var i=0;i<n.length;i++){var p=h.indexOf(n[i]);if(p>=0)return p+1;}return -1;}
function autoOn_(){return String(PropertiesService.getScriptProperties().getProperty("AUTO_TRANSCRIBE_ON_SUBMIT")||"Yes").toLowerCase()!=="no";}

function ensureSubmissionHeaders_(sh){
  var hs=["Submitted At","Started At","Duration Seconds","Project ID","Project Title","Institution","Researcher","Researcher Contact","Survey Title","Respondent ID","Enumerator ID","District","Latitude","Longitude","GPS Accuracy","GPS Timestamp","Consent","Question ID","Question Text","Question Audio","Answer File","Drive URL","File ID","Client Timestamp","Notes","Transcript","Transcript Status","Transcribed By","Transcription Date","Code","Theme","Auto Suggested","Sample Quote","Interpretation","Researcher Review"];
  if(sh.getLastRow()===0){sh.appendRow(hs);return;}
  var h=headers_(sh);hs.forEach(function(x){if(h.indexOf(x)<0){sh.getRange(1,sh.getLastColumn()+1).setValue(x);h.push(x);}});
}
function ensureCols_(){var sh=getSheet_();if(sh.getLastRow()===0)return;ensureSubmissionHeaders_(sh);}

function doGet(e){
  var cb=e&&e.parameter&&(e.parameter.callback||e.parameter.cb);
  try{
    var a=e&&e.parameter&&e.parameter.action;
    if(a==="stats")return out_(stats_(),cb);
    if(a==="transcriptionStatus")return out_({status:"success",version:"V21.0",engine:"OpenAI Whisper",apiKeyConfigured:!!PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY"),autoTranscribeOnSubmit:autoOn_()},cb);
    return out_({status:"success",message:"Backend V21.0 running",autoTranscribeOnSubmit:autoOn_()},cb);
  }catch(err){return out_({status:"error",message:String(err)},cb);}
}

function doPost(e){
  try{
    var data=JSON.parse(e.postData.contents);
    if(data.action==="transcribeAudio")return transcribeAudio_(data);
    if(data.action==="updateTranscript")return updateTranscript_(data);
    if(data.action==="transcribePendingBatch")return transcribePendingBatch_();

    var folder=DriveApp.getFolderById(FOLDER_ID), sh=getSheet_();
    ensureSubmissionHeaders_(sh);
    var h=headers_(sh), map={};h.forEach(function(x,i){map[x]=i+1;});
    var fileMap={};
    (data.files||[]).forEach(function(file){
      var blob=Utilities.newBlob(Utilities.base64Decode(file.base64),file.mimeType||"audio/webm",file.fileName);
      var saved=folder.createFile(blob);
      fileMap[file.fileName]={url:saved.getUrl(),id:saved.getId()};
    });
    var results=[], automatic=autoOn_();
    (data.metadata||[]).forEach(function(row){
      var f=fileMap[row.answerFileName]||{}, arr=new Array(sh.getLastColumn()).fill("");
      function set(k,v){if(map[k])arr[map[k]-1]=v;}
      set("Submitted At",data.submittedAt);set("Started At",data.startedAt||"");set("Duration Seconds",data.durationSeconds||"");
      set("Project ID",data.projectId||"");set("Project Title",data.projectTitle||"");set("Institution",data.institution||"");
      set("Researcher",data.researcherName||"");set("Researcher Contact",data.researcherContact||"");set("Survey Title",data.surveyTitle||"");
      set("Respondent ID",data.respondentId||"");set("Enumerator ID",data.enumeratorId||row.enumeratorId||"");set("District",data.district||row.district||"");
      set("Latitude",data.latitude||row.latitude||"");set("Longitude",data.longitude||row.longitude||"");set("GPS Accuracy",data.gpsAccuracy||row.gpsAccuracy||"");set("GPS Timestamp",data.gpsTimestamp||row.gpsTimestamp||"");
      set("Consent",data.consent||"");set("Question ID",row.questionId);set("Question Text",row.questionText);set("Question Audio",row.questionAudio);
      set("Answer File",row.answerFileName);set("Drive URL",f.url||"");set("File ID",f.id||"");set("Client Timestamp",row.timestamp);set("Notes",data.notes||row.notes||"");
      set("Transcript Status",automatic?"Auto Transcribing":"Pending");
      sh.appendRow(arr);
      if(automatic && f.id){
        try{
          var txt=transcribeFileToText_(f.id);
          var save=JSON.parse(updateTranscript_({respondentId:data.respondentId,questionId:row.questionId,fileId:f.id,answerFile:row.answerFileName,transcript:txt,transcribedBy:"OpenAI Whisper Auto",transcriptionDate:new Date().toISOString(),transcriptStatus:"Transcribed"}).getContent());
          results.push({file:row.answerFileName,status:"Transcribed",updated:save.updated});
        }catch(err){
          markStatus_(f.id,row.answerFileName,data.respondentId,row.questionId,"Failed / Needs Review",String(err));
          results.push({file:row.answerFileName,status:"Failed",message:String(err)});
        }
      }
    });
    return json_({status:"success",rows:(data.metadata||[]).length,autoTranscribeOnSubmit:automatic,autoTranscription:results});
  }catch(err){return json_({status:"error",message:String(err)});}
}

function transcribeFileToText_(fileId){
  var key=PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if(!key)throw new Error("OPENAI_API_KEY not configured");
  var file=DriveApp.getFileById(fileId), blob=file.getBlob(), boundary="----ARDC"+Date.now();
  var p1="--"+boundary+"\r\nContent-Disposition: form-data; name=\"model\"\r\n\r\nwhisper-1\r\n--"+boundary+"\r\nContent-Disposition: form-data; name=\"file\"; filename=\""+file.getName()+"\"\r\nContent-Type: "+blob.getContentType()+"\r\n\r\n";
  var p2="\r\n--"+boundary+"--";
  var payload=Utilities.newBlob(p1).getBytes().concat(blob.getBytes()).concat(Utilities.newBlob(p2).getBytes());
  var res=UrlFetchApp.fetch("https://api.openai.com/v1/audio/transcriptions",{method:"post",headers:{Authorization:"Bearer "+key},contentType:"multipart/form-data; boundary="+boundary,payload:payload,muteHttpExceptions:true});
  if(res.getResponseCode()<200||res.getResponseCode()>=300)throw new Error("OpenAI error: "+res.getContentText());
  return JSON.parse(res.getContentText()).text||"";
}

function updateTranscript_(d){
  ensureCols_();var sh=getSheet_(), h=headers_(sh);
  var respCol=col_(h,["Respondent ID","Respondent"]), qCol=col_(h,["Question ID","Question"]), fileCol=col_(h,["File ID"]), ansCol=col_(h,["Answer File"]), trCol=col_(h,["Transcript"]), stCol=col_(h,["Transcript Status"]), byCol=col_(h,["Transcribed By"]), dtCol=col_(h,["Transcription Date"]);
  var tf=String(d.fileId||"").trim(), ta=String(d.answerFile||"").trim(), tr=String(d.respondentId||"").trim(), tq=String(d.questionId||"").trim(), updated=0, matchedBy="";
  for(var r=2;r<=sh.getLastRow();r++){
    var rf=fileCol>0?String(sh.getRange(r,fileCol).getValue()||"").trim():"", ra=ansCol>0?String(sh.getRange(r,ansCol).getValue()||"").trim():"", rr=respCol>0?String(sh.getRange(r,respCol).getValue()||"").trim():"", rq=qCol>0?String(sh.getRange(r,qCol).getValue()||"").trim():"";
    var m=false;if(tf&&rf&&tf===rf){m=true;matchedBy="File ID";}else if(ta&&ra&&ta===ra){m=true;matchedBy="Answer File";}else if(tr&&tq&&rr===tr&&String(rq)===String(tq)){m=true;matchedBy="Respondent ID + Question ID";}
    if(m){sh.getRange(r,trCol).setValue(d.transcript||"");if(stCol>0)sh.getRange(r,stCol).setValue(d.transcriptStatus||"Transcribed");if(byCol>0)sh.getRange(r,byCol).setValue(d.transcribedBy||"OpenAI Whisper");if(dtCol>0)sh.getRange(r,dtCol).setValue(d.transcriptionDate||new Date().toISOString());updated++;}
  }
  return json_({status:"success",updated:updated,matchedBy:matchedBy});
}

function transcribeAudio_(d){
  try{
    var txt=transcribeFileToText_(d.fileId);
    var save=JSON.parse(updateTranscript_({respondentId:d.respondentId,questionId:d.questionId,fileId:d.fileId,answerFile:d.answerFile,transcript:txt,transcribedBy:"OpenAI Whisper Manual",transcriptionDate:new Date().toISOString(),transcriptStatus:"Transcribed"}).getContent());
    if(!save.updated)return json_({status:"error",message:"Transcript created but no sheet row updated",transcript:txt,updated:0});
    return json_({status:"success",transcript:txt,updated:save.updated,matchedBy:save.matchedBy});
  }catch(err){return json_({status:"error",message:String(err)});}
}

function markStatus_(fileId,answerFile,respondentId,questionId,status,note){
  ensureCols_();var sh=getSheet_(), h=headers_(sh);
  var fileCol=col_(h,["File ID"]), ansCol=col_(h,["Answer File"]), respCol=col_(h,["Respondent ID"]), qCol=col_(h,["Question ID"]), stCol=col_(h,["Transcript Status"]), notesCol=col_(h,["Notes"]);
  for(var r=2;r<=sh.getLastRow();r++){
    var rf=fileCol>0?String(sh.getRange(r,fileCol).getValue()||"").trim():"", ra=ansCol>0?String(sh.getRange(r,ansCol).getValue()||"").trim():"", rr=respCol>0?String(sh.getRange(r,respCol).getValue()||"").trim():"", rq=qCol>0?String(sh.getRange(r,qCol).getValue()||"").trim():"";
    if((fileId&&rf===fileId)||(answerFile&&ra===answerFile)||(respondentId&&questionId&&rr===String(respondentId)&&String(rq)===String(questionId))){
      if(stCol>0)sh.getRange(r,stCol).setValue(status);
      if(notesCol>0&&note)sh.getRange(r,notesCol).setValue(String(note).substring(0,500));
    }
  }
}

function transcribePendingBatch_(){
  ensureCols_();var sh=getSheet_(), h=headers_(sh);
  var respCol=col_(h,["Respondent ID"]), qCol=col_(h,["Question ID"]), fileCol=col_(h,["File ID"]), ansCol=col_(h,["Answer File"]), trCol=col_(h,["Transcript"]), stCol=col_(h,["Transcript Status"]);
  var limit=Number(PropertiesService.getScriptProperties().getProperty("BATCH_TRANSCRIBE_LIMIT")||"10"), done=0, results=[];
  for(var r=2;r<=sh.getLastRow()&&done<limit;r++){
    var transcript=trCol>0?String(sh.getRange(r,trCol).getValue()||"").trim():"", fileId=fileCol>0?String(sh.getRange(r,fileCol).getValue()||"").trim():"";
    if(fileId&&!transcript){
      try{
        if(stCol>0)sh.getRange(r,stCol).setValue("Auto Transcribing");
        var txt=transcribeFileToText_(fileId);
        var save=JSON.parse(updateTranscript_({respondentId:respCol>0?sh.getRange(r,respCol).getValue():"",questionId:qCol>0?sh.getRange(r,qCol).getValue():"",fileId:fileId,answerFile:ansCol>0?sh.getRange(r,ansCol).getValue():"",transcript:txt,transcribedBy:"OpenAI Whisper Batch",transcriptionDate:new Date().toISOString(),transcriptStatus:"Transcribed"}).getContent());
        results.push({row:r,status:"Transcribed",updated:save.updated});
      }catch(err){if(stCol>0)sh.getRange(r,stCol).setValue("Failed / Needs Review");results.push({row:r,status:"Failed",message:String(err)});}
      done++;
    }
  }
  return json_({status:"success",processed:done,results:results});
}

function authorizeUrlFetch(){UrlFetchApp.fetch("https://www.google.com");}

function stats_(){
  var sh=getSheet_();ensureSubmissionHeaders_(sh);var v=sh.getDataRange().getValues(),h=v[0].map(String);
  function ix(n){return h.indexOf(n);}
  var ri=ix("Respondent ID"),qi=ix("Question ID"),ei=ix("Enumerator ID"),di=ix("District"),lati=ix("Latitude"),lngi=ix("Longitude"),afi=ix("Answer File"),dui=ix("Drive URL"),fidi=ix("File ID"),tri=ix("Transcript"),tsi=ix("Transcript Status"),thi=ix("Theme"),coi=ix("Code"),qui=ix("Sample Quote"),ini=ix("Interpretation");
  var respondents={},questions={},progress={},gps={},enums={},dists={},byEnum={},byDist={},manifest=[],themeMap={},codeMap={},answers=0,transcribed=0,pending=0,failed=0;
  for(var r=1;r<v.length;r++){
    var row=v[r],resp=ri>=0?String(row[ri]||"").trim():"",q=qi>=0?String(row[qi]||"").trim():"",en=ei>=0?String(row[ei]||"").trim():"",dist=di>=0?String(row[di]||"").trim():"";
    if(resp)respondents[resp]=true;if(q)questions[q]=true;if(resp&&q){answers++;if(!progress[resp])progress[resp]={};progress[resp][q]=true;}
    if(lati>=0&&lngi>=0&&row[lati]&&row[lngi]&&resp)gps[resp]=true;
    if(en){enums[en]=true;if(resp){if(!byEnum[en])byEnum[en]={};byEnum[en][resp]=true;}}
    if(dist){dists[dist]=true;if(resp){if(!byDist[dist])byDist[dist]={};byDist[dist][resp]=true;}}
    var transcript=tri>=0?String(row[tri]||"").trim():"",st=tsi>=0?String(row[tsi]||"").trim():"";if(!st)st=transcript?"Transcribed":"Pending";
    if(afi>=0&&row[afi]){if(transcript||st==="Transcribed")transcribed++;else if(String(st).indexOf("Failed")>=0||st==="Error")failed++;else pending++;manifest.push({respondentId:resp,questionId:q,answerFile:String(row[afi]||""),driveUrl:dui>=0?String(row[dui]||""):"",fileId:fidi>=0?String(row[fidi]||""):"",transcriptStatus:st});}
    var th=thi>=0?String(row[thi]||"").trim():"",co=coi>=0?String(row[coi]||"").trim():"";if(th){if(!themeMap[th])themeMap[th]={theme:th,mentions:0,sampleQuote:"",interpretation:""};themeMap[th].mentions++;if(!themeMap[th].sampleQuote&&qui>=0)themeMap[th].sampleQuote=String(row[qui]||"");if(!themeMap[th].interpretation&&ini>=0)themeMap[th].interpretation=String(row[ini]||"");}if(co)codeMap[co]=(codeMap[co]||0)+1;
  }
  var rc=Object.keys(respondents).length,qc=Object.keys(questions).length,fully=Object.keys(progress).filter(function(k){return Object.keys(progress[k]).length>=qc;}).length,prog={};Object.keys(progress).forEach(function(k){var n=Object.keys(progress[k]).length,key=n+" of "+qc;prog[key]=(prog[key]||0)+1;});
  var enumOut={};Object.keys(byEnum).forEach(function(k){enumOut[k]=Object.keys(byEnum[k]).length;});var distOut={};Object.keys(byDist).forEach(function(k){distOut[k]=Object.keys(byDist[k]).length;});
  return {status:"success",respondents:rc,answers:answers,questions:qc,completion:rc?Math.round(fully/rc*100):0,gpsCaptured:Object.keys(gps).length,fullyAnswered:fully,enumerators:Object.keys(enums).length,districts:Object.keys(dists).length,byEnumerator:enumOut,byDistrict:distOut,progressSummary:prog,audioManifest:manifest,transcription:{audioRecords:manifest.length,transcribed:transcribed,pending:pending,failed:failed,coverage:manifest.length?Math.round(transcribed/manifest.length*100):0},themeSummary:Object.keys(themeMap).map(function(k){return themeMap[k];}),codeSummary:Object.keys(codeMap).map(function(k){return {code:k,mentions:codeMap[k]};}),updatedAt:new Date().toISOString()};
}

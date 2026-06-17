const SHEET_ID="1R-cjiNMoNWv-BhokRXeeBCmHtssr4-LMQwYGTd39g-0";
const FOLDER_ID="13lzW_JzwCiLaYXDGWV77Muo90FjNV5au";
function sh_(){return SpreadsheetApp.openById(SHEET_ID).getSheets()[0];}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function out_(o,cb){var s=cb?cb+"("+JSON.stringify(o)+");":JSON.stringify(o);return ContentService.createTextOutput(s).setMimeType(cb?ContentService.MimeType.JAVASCRIPT:ContentService.MimeType.JSON);}
function hdr_(sh){return sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getValues()[0].map(String);}
function col_(h,n){for(var i=0;i<n.length;i++){var p=h.indexOf(n[i]);if(p>=0)return p+1;}return -1;}
function autoTx_(){return String(PropertiesService.getScriptProperties().getProperty("AUTO_TRANSCRIBE_ON_SUBMIT")||"Yes").toLowerCase()!=="no";}
function autoCode_(){return String(PropertiesService.getScriptProperties().getProperty("AUTO_CODE_ON_TRANSCRIBE")||"Yes").toLowerCase()!=="no";}
function ensure_(sh){
 var hs=["Submitted At","Started At","Duration Seconds","Project ID","Project Title","Institution","Researcher","Researcher Contact","Survey Title","Respondent ID","Enumerator ID","District","Latitude","Longitude","GPS Accuracy","GPS Timestamp","Consent","Question ID","Question Text","Question Audio","Answer File","Drive URL","File ID","Client Timestamp","Notes","Transcript","Transcript Status","Transcribed By","Transcription Date","Code","Theme","Auto Suggested","Sample Quote","Interpretation","Researcher Review"];
 if(sh.getLastRow()===0){sh.appendRow(hs);return;}
 var h=hdr_(sh);hs.forEach(function(x){if(h.indexOf(x)<0){sh.getRange(1,sh.getLastColumn()+1).setValue(x);h.push(x);}});
}
function doGet(e){
 var cb=e&&e.parameter&&(e.parameter.callback||e.parameter.cb), a=e&&e.parameter&&e.parameter.action;
 try{
  if(a==="stats")return out_(stats_(),cb);
  if(a==="transcriptionStatus")return out_({status:"success",version:"V22.0",engine:"OpenAI Whisper",apiKeyConfigured:!!PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY"),autoTranscribeOnSubmit:autoTx_(),autoCodeOnTranscribe:autoCode_()},cb);
  if(a==="codePendingBatch")return out_(codePendingBatchObj_(),cb);
  return out_({status:"success",message:"Backend V22.0 running",autoTranscribeOnSubmit:autoTx_(),autoCodeOnTranscribe:autoCode_()},cb);
 }catch(err){return out_({status:"error",message:String(err)},cb);}
}
function doPost(e){
 try{
  var d=JSON.parse(e.postData.contents);
  if(d.action==="transcribeAudio")return transcribeAudio_(d);
  if(d.action==="updateTranscript")return updateTranscript_(d);
  if(d.action==="codePendingBatch")return json_(codePendingBatchObj_());
  var folder=DriveApp.getFolderById(FOLDER_ID), sh=sh_(); ensure_(sh);
  var h=hdr_(sh), map={}; h.forEach(function(x,i){map[x]=i+1;});
  var fm={}; (d.files||[]).forEach(function(f){var b=Utilities.newBlob(Utilities.base64Decode(f.base64),f.mimeType||"audio/webm",f.fileName);var saved=folder.createFile(b);fm[f.fileName]={url:saved.getUrl(),id:saved.getId()};});
  var results=[], automatic=autoTx_();
  (d.metadata||[]).forEach(function(row){
   var f=fm[row.answerFileName]||{}, arr=new Array(sh.getLastColumn()).fill("");
   function set(k,v){if(map[k])arr[map[k]-1]=v;}
   set("Submitted At",d.submittedAt); set("Started At",d.startedAt||""); set("Duration Seconds",d.durationSeconds||""); set("Project ID",d.projectId||""); set("Project Title",d.projectTitle||""); set("Institution",d.institution||"");
   set("Researcher",d.researcherName||""); set("Researcher Contact",d.researcherContact||""); set("Survey Title",d.surveyTitle||""); set("Respondent ID",d.respondentId||""); set("Enumerator ID",d.enumeratorId||row.enumeratorId||""); set("District",d.district||row.district||"");
   set("Latitude",d.latitude||row.latitude||""); set("Longitude",d.longitude||row.longitude||""); set("GPS Accuracy",d.gpsAccuracy||row.gpsAccuracy||""); set("GPS Timestamp",d.gpsTimestamp||row.gpsTimestamp||""); set("Consent",d.consent||"");
   set("Question ID",row.questionId); set("Question Text",row.questionText); set("Question Audio",row.questionAudio); set("Answer File",row.answerFileName); set("Drive URL",f.url||""); set("File ID",f.id||""); set("Client Timestamp",row.timestamp); set("Notes",d.notes||row.notes||""); set("Transcript Status",automatic?"Auto Transcribing":"Pending");
   sh.appendRow(arr);
   if(automatic&&f.id){try{var txt=transcribeFile_(f.id), c=autoCode_()?analyse_(txt):{}; var save=JSON.parse(updateTranscript_({respondentId:d.respondentId,questionId:row.questionId,fileId:f.id,answerFile:row.answerFileName,transcript:txt,transcribedBy:"OpenAI Whisper Auto",transcriptionDate:new Date().toISOString(),transcriptStatus:"Transcribed",code:c.code||"",theme:c.theme||"",autoSuggested:c.autoSuggested||"",sampleQuote:c.sampleQuote||"",interpretation:c.interpretation||"",researcherReview:"Pending Researcher Review"}).getContent());results.push({file:row.answerFileName,status:"Transcribed and Coded",updated:save.updated});}catch(err){mark_(f.id,row.answerFileName,d.respondentId,row.questionId,"Failed / Needs Review",String(err));results.push({file:row.answerFileName,status:"Failed",message:String(err)});}}
  });
  return json_({status:"success",rows:(d.metadata||[]).length,autoTranscribeOnSubmit:automatic,autoCodeOnTranscribe:autoCode_(),autoTranscription:results});
 }catch(err){return json_({status:"error",message:String(err)});}
}
function analyse_(txt){
 txt=String(txt||"").trim(); var low=txt.toLowerCase(), code=[], theme=[];
 if(/transport|road|distance|travel|far|bus|taxi/.test(low)){code.push("Transport and Distance");theme.push("Access Barriers");}
 if(/cost|money|fee|expensive|afford|price|payment/.test(low)){code.push("Cost Constraint");theme.push("Financial Barriers");}
 if(/skill|practical|training|competence|hands-on|practice/.test(low)){code.push("Practical Skills");theme.push("Skills Development");}
 if(/ict|computer|digital|system|online|internet|software|platform/.test(low)){code.push("ICT Competence");theme.push("Digital Readiness");}
 if(/attachment|industry|workplace|employer|internship/.test(low)){code.push("Industrial Attachment");theme.push("Work-Based Learning");}
 if(/delay|slow|late|time|waiting/.test(low)){code.push("Service Delay");theme.push("System Efficiency");}
 if(/certificate|result|assessment|examination|exam|curriculum|programme|program/.test(low)){code.push("Assessment / Curriculum Service");theme.push("Assessment and Certification");}
 if(/clear|loud|hear|listen|audio|sound|cheers|test/.test(low)){code.push("Audio Quality");theme.push("Data Collection Quality");}
 if(!code.length){code.push("General Response");theme.push("General Findings");}
 code=[...new Set(code)]; theme=[...new Set(theme)];
 var quote=(txt.split(/[.!?]\s+/).filter(function(s){return s.trim().length>5;})[0]||txt.substring(0,160)).substring(0,250);
 return {code:code.join("; "),theme:theme.join("; "),autoSuggested:"Yes",sampleQuote:quote,interpretation:"The transcript indicates issues related to "+theme.join(", ")+". The researcher should compare this response with other interviews before final reporting."};
}
function transcribeFile_(fid){
 var key=PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY"); if(!key)throw new Error("OPENAI_API_KEY not configured");
 var file=DriveApp.getFileById(fid), blob=file.getBlob(), bd="----ARDC"+Date.now();
 var p1="--"+bd+"\r\nContent-Disposition: form-data; name=\"model\"\r\n\r\nwhisper-1\r\n--"+bd+"\r\nContent-Disposition: form-data; name=\"file\"; filename=\""+file.getName()+"\"\r\nContent-Type: "+blob.getContentType()+"\r\n\r\n";
 var payload=Utilities.newBlob(p1).getBytes().concat(blob.getBytes()).concat(Utilities.newBlob("\r\n--"+bd+"--").getBytes());
 var res=UrlFetchApp.fetch("https://api.openai.com/v1/audio/transcriptions",{method:"post",headers:{Authorization:"Bearer "+key},contentType:"multipart/form-data; boundary="+bd,payload:payload,muteHttpExceptions:true});
 if(res.getResponseCode()<200||res.getResponseCode()>=300)throw new Error("OpenAI error: "+res.getContentText());
 return JSON.parse(res.getContentText()).text||"";
}
function updateTranscript_(d){
 var sh=sh_(); ensure_(sh); var h=hdr_(sh);
 var resp=col_(h,["Respondent ID","Respondent"]), q=col_(h,["Question ID","Question"]), fid=col_(h,["File ID"]), ans=col_(h,["Answer File"]), trc=col_(h,["Transcript"]), st=col_(h,["Transcript Status"]), by=col_(h,["Transcribed By"]), dt=col_(h,["Transcription Date"]), code=col_(h,["Code"]), theme=col_(h,["Theme"]), auto=col_(h,["Auto Suggested"]), quote=col_(h,["Sample Quote"]), interp=col_(h,["Interpretation"]), rev=col_(h,["Researcher Review"]);
 var tf=String(d.fileId||"").trim(), ta=String(d.answerFile||"").trim(), tr=String(d.respondentId||"").trim(), tq=String(d.questionId||"").trim(), updated=0, matchedBy="";
 for(var r=2;r<=sh.getLastRow();r++){
  var rf=fid>0?String(sh.getRange(r,fid).getValue()||"").trim():"", ra=ans>0?String(sh.getRange(r,ans).getValue()||"").trim():"", rr=resp>0?String(sh.getRange(r,resp).getValue()||"").trim():"", rq=q>0?String(sh.getRange(r,q).getValue()||"").trim():"";
  var m=false; if(tf&&rf&&tf===rf){m=true;matchedBy="File ID";}else if(ta&&ra&&ta===ra){m=true;matchedBy="Answer File";}else if(tr&&tq&&rr===tr&&String(rq)===String(tq)){m=true;matchedBy="Respondent ID + Question ID";}
  if(m){if(trc>0)sh.getRange(r,trc).setValue(d.transcript||""); if(st>0)sh.getRange(r,st).setValue(d.transcriptStatus||"Transcribed"); if(by>0)sh.getRange(r,by).setValue(d.transcribedBy||"OpenAI Whisper"); if(dt>0)sh.getRange(r,dt).setValue(d.transcriptionDate||new Date().toISOString()); if(code>0&&d.code!==undefined)sh.getRange(r,code).setValue(d.code); if(theme>0&&d.theme!==undefined)sh.getRange(r,theme).setValue(d.theme); if(auto>0&&d.autoSuggested!==undefined)sh.getRange(r,auto).setValue(d.autoSuggested); if(quote>0&&d.sampleQuote!==undefined)sh.getRange(r,quote).setValue(d.sampleQuote); if(interp>0&&d.interpretation!==undefined)sh.getRange(r,interp).setValue(d.interpretation); if(rev>0&&d.researcherReview!==undefined)sh.getRange(r,rev).setValue(d.researcherReview); updated++;}
 }
 return json_({status:"success",updated:updated,matchedBy:matchedBy});
}
function transcribeAudio_(d){try{var txt=transcribeFile_(d.fileId), c=autoCode_()?analyse_(txt):{};var save=JSON.parse(updateTranscript_({respondentId:d.respondentId,questionId:d.questionId,fileId:d.fileId,answerFile:d.answerFile,transcript:txt,transcribedBy:"OpenAI Whisper Manual",transcriptionDate:new Date().toISOString(),transcriptStatus:"Transcribed",code:c.code||"",theme:c.theme||"",autoSuggested:c.autoSuggested||"",sampleQuote:c.sampleQuote||"",interpretation:c.interpretation||"",researcherReview:"Pending Researcher Review"}).getContent());return json_({status:"success",transcript:txt,updated:save.updated,matchedBy:save.matchedBy});}catch(err){return json_({status:"error",message:String(err)});}}
function mark_(fid,ans,resp,q,status,note){var sh=sh_();ensure_(sh);var h=hdr_(sh),fc=col_(h,["File ID"]),ac=col_(h,["Answer File"]),rc=col_(h,["Respondent ID"]),qc=col_(h,["Question ID"]),sc=col_(h,["Transcript Status"]),nc=col_(h,["Notes"]);for(var r=2;r<=sh.getLastRow();r++){if((fc>0&&String(sh.getRange(r,fc).getValue())===String(fid))||(ac>0&&String(sh.getRange(r,ac).getValue())===String(ans))){if(sc>0)sh.getRange(r,sc).setValue(status); if(nc>0)sh.getRange(r,nc).setValue(String(note).substring(0,500));}}}
function codePendingBatchObj_(){var sh=sh_();ensure_(sh);var h=hdr_(sh),tc=col_(h,["Transcript"]),cc=col_(h,["Code"]),thc=col_(h,["Theme"]),ac=col_(h,["Auto Suggested"]),qc=col_(h,["Sample Quote"]),ic=col_(h,["Interpretation"]),rc=col_(h,["Researcher Review"]),done=0;for(var r=2;r<=sh.getLastRow();r++){var txt=tc>0?String(sh.getRange(r,tc).getValue()||"").trim():"", code=cc>0?String(sh.getRange(r,cc).getValue()||"").trim():"";if(txt&&!code){var c=analyse_(txt); if(cc>0)sh.getRange(r,cc).setValue(c.code); if(thc>0)sh.getRange(r,thc).setValue(c.theme); if(ac>0)sh.getRange(r,ac).setValue("Yes"); if(qc>0)sh.getRange(r,qc).setValue(c.sampleQuote); if(ic>0)sh.getRange(r,ic).setValue(c.interpretation); if(rc>0)sh.getRange(r,rc).setValue("Pending Researcher Review"); done++;}}return {status:"success",coded:done};}
function authorizeUrlFetch(){UrlFetchApp.fetch("https://www.google.com");}
function stats_(){var sh=sh_();ensure_(sh);var v=sh.getDataRange().getValues(),h=v[0].map(String);function ix(n){return h.indexOf(n);}var ri=ix("Respondent ID"),qi=ix("Question ID"),afi=ix("Answer File"),tri=ix("Transcript"),tsi=ix("Transcript Status"),thi=ix("Theme"),coi=ix("Code"),dui=ix("Drive URL"),fidi=ix("File ID"),resp={},qs={},manifest=[],themeMap={},codeMap={},ans=0,tx=0,p=0,f=0;for(var r=1;r<v.length;r++){var row=v[r],rp=ri>=0?String(row[ri]||"").trim():"",q=qi>=0?String(row[qi]||"").trim():"";if(rp)resp[rp]=1;if(q)qs[q]=1;if(rp&&q)ans++;var transcript=tri>=0?String(row[tri]||"").trim():"",st=tsi>=0?String(row[tsi]||"").trim():"";if(!st)st=transcript?"Transcribed":"Pending";if(afi>=0&&row[afi]){if(transcript||st==="Transcribed")tx++;else if(String(st).indexOf("Failed")>=0)f++;else p++;manifest.push({respondentId:rp,questionId:q,answerFile:String(row[afi]||""),driveUrl:dui>=0?String(row[dui]||""):"",fileId:fidi>=0?String(row[fidi]||""):"",transcriptStatus:st});}var th=thi>=0?String(row[thi]||"").trim():"",co=coi>=0?String(row[coi]||"").trim():"";if(th){if(!themeMap[th])themeMap[th]={theme:th,mentions:0};themeMap[th].mentions++;}if(co)codeMap[co]=(codeMap[co]||0)+1;}return {status:"success",respondents:Object.keys(resp).length,answers:ans,questions:Object.keys(qs).length,completion:0,gpsCaptured:0,fullyAnswered:0,enumerators:0,districts:0,byEnumerator:{},byDistrict:{},progressSummary:{},audioManifest:manifest,transcription:{audioRecords:manifest.length,transcribed:tx,pending:p,failed:f,coverage:manifest.length?Math.round(tx/manifest.length*100):0},themeSummary:Object.keys(themeMap).map(k=>themeMap[k]),codeSummary:Object.keys(codeMap).map(k=>({code:k,mentions:codeMap[k]})),updatedAt:new Date().toISOString()};}

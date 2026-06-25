
const SHEET_ID="1R-cjiNMoNWv-BhokRXeeBCmHtssr4-LMQwYGTd39g-0";
const FOLDER_ID="13lzW_JzwCiLaYXDGWV77Muo90FjNV5au";
function sh_(){return SpreadsheetApp.openById(SHEET_ID).getSheets()[0];}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function out_(o,cb){var s=cb?cb+"("+JSON.stringify(o)+");":JSON.stringify(o);return ContentService.createTextOutput(s).setMimeType(cb?ContentService.MimeType.JAVASCRIPT:ContentService.MimeType.JSON);}
function hdr_(sh){return sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getValues()[0].map(String);}
function col_(h,n){for(var i=0;i<n.length;i++){var p=h.indexOf(n[i]);if(p>=0)return p+1;}return -1;}
function autoTx_(){return String(PropertiesService.getScriptProperties().getProperty("AUTO_TRANSCRIBE_ON_SUBMIT")||"Yes").toLowerCase()!=="no";}
function autoCode_(){return String(PropertiesService.getScriptProperties().getProperty("AUTO_CODE_ON_TRANSCRIBE")||"Yes").toLowerCase()!=="no";}
function ensure_(sh){var hs=["Submitted At","Started At","Duration Seconds","Project ID","Project Title","Institution","Researcher","Researcher Contact","Survey Title","Respondent ID","Enumerator ID","District","Latitude","Longitude","GPS Accuracy","GPS Timestamp","Consent","Question ID","Question Text","Question Audio","Answer File","Drive URL","File ID","Client Timestamp","Notes","Transcript","Transcript Status","Transcribed By","Transcription Date","Code","Theme","Auto Suggested","Sample Quote","Interpretation","Researcher Review"];if(sh.getLastRow()===0){sh.appendRow(hs);return;}var h=hdr_(sh);hs.forEach(function(x){if(h.indexOf(x)<0){sh.getRange(1,sh.getLastColumn()+1).setValue(x);h.push(x);}});}
function doGet(e){var cb=e&&e.parameter&&(e.parameter.callback||e.parameter.cb),a=e&&e.parameter&&e.parameter.action;try{if(a==="stats")return out_(stats_(),cb);if(a==="transcriptionStatus")return out_({status:"success",version:"V24.1",engine:"OpenAI Whisper",apiKeyConfigured:!!PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY"),autoTranscribeOnSubmit:autoTx_(),autoCodeOnTranscribe:autoCode_()},cb);if(a==="codePendingBatch")return out_(codePendingBatchObj_(),cb);return out_({status:"success",message:"Backend V24.1 running",autoTranscribeOnSubmit:autoTx_(),autoCodeOnTranscribe:autoCode_()},cb);}catch(err){return out_({status:"error",message:String(err)},cb);}}
function doPost(e){try{var d=JSON.parse(e.postData.contents);if(d.action==="transcribeAudio")return transcribeAudio_(d);if(d.action==="updateTranscript")return updateTranscript_(d);if(d.action==="codePendingBatch")return json_(codePendingBatchObj_());var folder=DriveApp.getFolderById(FOLDER_ID),sh=sh_();ensure_(sh);var h=hdr_(sh),map={};h.forEach(function(x,i){map[x]=i+1;});var fm={};(d.files||[]).forEach(function(f){var b=Utilities.newBlob(Utilities.base64Decode(f.base64),f.mimeType||"audio/webm",f.fileName);var saved=folder.createFile(b);fm[f.fileName]={url:saved.getUrl(),id:saved.getId()};});var results=[],automatic=autoTx_();(d.metadata||[]).forEach(function(row){var f=fm[row.answerFileName]||{},arr=new Array(sh.getLastColumn()).fill("");function set(k,v){if(map[k])arr[map[k]-1]=v;}set("Submitted At",d.submittedAt);set("Started At",d.startedAt||"");set("Duration Seconds",d.durationSeconds||"");set("Project ID",d.projectId||"");set("Project Title",d.projectTitle||"");set("Institution",d.institution||"");set("Researcher",d.researcherName||"");set("Researcher Contact",d.researcherContact||"");set("Survey Title",d.surveyTitle||"");set("Respondent ID",d.respondentId||"");set("Enumerator ID",d.enumeratorId||row.enumeratorId||"");set("District",d.district||row.district||"");set("Latitude",d.latitude||row.latitude||"");set("Longitude",d.longitude||row.longitude||"");set("GPS Accuracy",d.gpsAccuracy||row.gpsAccuracy||"");set("GPS Timestamp",d.gpsTimestamp||row.gpsTimestamp||"");set("Consent",d.consent||"");set("Question ID",row.questionId);set("Question Text",row.questionText);set("Question Audio",row.questionAudio);set("Answer File",row.answerFileName);set("Drive URL",f.url||"");set("File ID",f.id||"");set("Client Timestamp",row.timestamp);set("Notes",d.notes||row.notes||"");set("Transcript Status",automatic?"Auto Transcribing":"Pending");sh.appendRow(arr);if(automatic&&f.id){try{var txt=transcribeFile_(f.id),c=autoCode_()?analyse_(txt):{};var save=JSON.parse(updateTranscript_({respondentId:d.respondentId,questionId:row.questionId,fileId:f.id,answerFile:row.answerFileName,transcript:txt,transcribedBy:"OpenAI Whisper Auto",transcriptionDate:new Date().toISOString(),transcriptStatus:"Transcribed",code:c.code||"",theme:c.theme||"",autoSuggested:c.autoSuggested||"",sampleQuote:c.sampleQuote||"",interpretation:c.interpretation||"",researcherReview:"Pending Researcher Review"}).getContent());results.push({file:row.answerFileName,status:"Transcribed and Coded",updated:save.updated});}catch(err){results.push({file:row.answerFileName,status:"Failed",message:String(err)});}}});return json_({status:"success",rows:(d.metadata||[]).length,autoTranscribeOnSubmit:automatic,autoCodeOnTranscribe:autoCode_(),autoTranscription:results});}catch(err){return json_({status:"error",message:String(err)});}}
function analyse_(txt){txt=String(txt||"").trim();var low=txt.toLowerCase(),code=[],theme=[];if(/transport|road|distance|travel|far|bus|taxi/.test(low)){code.push("Transport and Distance");theme.push("Access Barriers");}if(/cost|money|fee|expensive|afford|price|payment/.test(low)){code.push("Cost Constraint");theme.push("Financial Barriers");}if(/skill|practical|training|competence|hands-on|practice/.test(low)){code.push("Practical Skills");theme.push("Skills Development");}if(/ict|computer|digital|system|online|internet|software|platform/.test(low)){code.push("ICT Competence");theme.push("Digital Readiness");}if(/attachment|industry|workplace|employer|internship/.test(low)){code.push("Industrial Attachment");theme.push("Work-Based Learning");}if(/delay|slow|late|time|waiting/.test(low)){code.push("Service Delay");theme.push("System Efficiency");}if(/certificate|result|assessment|examination|exam|curriculum|programme|program/.test(low)){code.push("Assessment / Curriculum Service");theme.push("Assessment and Certification");}if(/clear|loud|hear|listen|audio|sound|cheers|test/.test(low)){code.push("Audio Quality");theme.push("Data Collection Quality");}if(!code.length){code.push("General Response");theme.push("General Findings");}code=[...new Set(code)];theme=[...new Set(theme)];var quote=(txt.split(/[.!?]\s+/).filter(function(s){return s.trim().length>5;})[0]||txt.substring(0,160)).substring(0,250);return {code:code.join("; "),theme:theme.join("; "),autoSuggested:"Yes",sampleQuote:quote,interpretation:"The transcript indicates issues related to "+theme.join(", ")+". The researcher should compare this response with other interviews before final reporting."};}
function transcribeFile_(fid){var key=PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");if(!key)throw new Error("OPENAI_API_KEY not configured");var file=DriveApp.getFileById(fid),blob=file.getBlob(),bd="----ARDC"+Date.now();var p1="--"+bd+"\r\nContent-Disposition: form-data; name=\"model\"\r\n\r\nwhisper-1\r\n--"+bd+"\r\nContent-Disposition: form-data; name=\"file\"; filename=\""+file.getName()+"\"\r\nContent-Type: "+blob.getContentType()+"\r\n\r\n";var payload=Utilities.newBlob(p1).getBytes().concat(blob.getBytes()).concat(Utilities.newBlob("\r\n--"+bd+"--").getBytes());var res=UrlFetchApp.fetch("https://api.openai.com/v1/audio/transcriptions",{method:"post",headers:{Authorization:"Bearer "+key},contentType:"multipart/form-data; boundary="+bd,payload:payload,muteHttpExceptions:true});if(res.getResponseCode()<200||res.getResponseCode()>=300)throw new Error("OpenAI error: "+res.getContentText());return JSON.parse(res.getContentText()).text||"";}
function updateTranscript_(d){var sh=sh_();ensure_(sh);var h=hdr_(sh);var resp=col_(h,["Respondent ID","Respondent"]),q=col_(h,["Question ID","Question"]),fid=col_(h,["File ID"]),ans=col_(h,["Answer File"]),trc=col_(h,["Transcript"]),st=col_(h,["Transcript Status"]),by=col_(h,["Transcribed By"]),dt=col_(h,["Transcription Date"]),code=col_(h,["Code"]),theme=col_(h,["Theme"]),auto=col_(h,["Auto Suggested"]),quote=col_(h,["Sample Quote"]),interp=col_(h,["Interpretation"]),rev=col_(h,["Researcher Review"]);var tf=String(d.fileId||"").trim(),ta=String(d.answerFile||"").trim(),tr=String(d.respondentId||"").trim(),tq=String(d.questionId||"").trim(),updated=0,matchedBy="";for(var r=2;r<=sh.getLastRow();r++){var rf=fid>0?String(sh.getRange(r,fid).getValue()||"").trim():"",ra=ans>0?String(sh.getRange(r,ans).getValue()||"").trim():"",rr=resp>0?String(sh.getRange(r,resp).getValue()||"").trim():"",rq=q>0?String(sh.getRange(r,q).getValue()||"").trim():"";var m=false;if(tf&&rf&&tf===rf){m=true;matchedBy="File ID";}else if(ta&&ra&&ta===ra){m=true;matchedBy="Answer File";}else if(tr&&tq&&rr===tr&&String(rq)===String(tq)){m=true;matchedBy="Respondent ID + Question ID";}if(m){if(trc>0)sh.getRange(r,trc).setValue(d.transcript||"");if(st>0)sh.getRange(r,st).setValue(d.transcriptStatus||"Transcribed");if(by>0)sh.getRange(r,by).setValue(d.transcribedBy||"OpenAI Whisper");if(dt>0)sh.getRange(r,dt).setValue(d.transcriptionDate||new Date().toISOString());if(code>0&&d.code!==undefined)sh.getRange(r,code).setValue(d.code);if(theme>0&&d.theme!==undefined)sh.getRange(r,theme).setValue(d.theme);if(auto>0&&d.autoSuggested!==undefined)sh.getRange(r,auto).setValue(d.autoSuggested);if(quote>0&&d.sampleQuote!==undefined)sh.getRange(r,quote).setValue(d.sampleQuote);if(interp>0&&d.interpretation!==undefined)sh.getRange(r,interp).setValue(d.interpretation);if(rev>0&&d.researcherReview!==undefined)sh.getRange(r,rev).setValue(d.researcherReview);updated++;}}return json_({status:"success",updated:updated,matchedBy:matchedBy});}
function transcribeAudio_(d){try{var txt=transcribeFile_(d.fileId),c=autoCode_()?analyse_(txt):{};var save=JSON.parse(updateTranscript_({respondentId:d.respondentId,questionId:d.questionId,fileId:d.fileId,answerFile:d.answerFile,transcript:txt,transcribedBy:"OpenAI Whisper Manual",transcriptionDate:new Date().toISOString(),transcriptStatus:"Transcribed",code:c.code||"",theme:c.theme||"",autoSuggested:c.autoSuggested||"",sampleQuote:c.sampleQuote||"",interpretation:c.interpretation||"",researcherReview:"Pending Researcher Review"}).getContent());return json_({status:"success",transcript:txt,updated:save.updated,matchedBy:save.matchedBy});}catch(err){return json_({status:"error",message:String(err)});}}
function codePendingBatchObj_(){var sh=sh_();ensure_(sh);var h=hdr_(sh),tc=col_(h,["Transcript"]),cc=col_(h,["Code"]),thc=col_(h,["Theme"]),ac=col_(h,["Auto Suggested"]),qc=col_(h,["Sample Quote"]),ic=col_(h,["Interpretation"]),rc=col_(h,["Researcher Review"]),done=0;for(var r=2;r<=sh.getLastRow();r++){var txt=tc>0?String(sh.getRange(r,tc).getValue()||"").trim():"",code=cc>0?String(sh.getRange(r,cc).getValue()||"").trim():"";if(txt&&!code){var c=analyse_(txt);if(cc>0)sh.getRange(r,cc).setValue(c.code);if(thc>0)sh.getRange(r,thc).setValue(c.theme);if(ac>0)sh.getRange(r,ac).setValue("Yes");if(qc>0)sh.getRange(r,qc).setValue(c.sampleQuote);if(ic>0)sh.getRange(r,ic).setValue(c.interpretation);if(rc>0)sh.getRange(r,rc).setValue("Pending Researcher Review");done++;}}return {status:"success",coded:done};}
function authorizeUrlFetch(){UrlFetchApp.fetch("https://www.google.com");}
function norm_(v){return String(v==null?"":v).trim();}
function goodCoord_(v){var n=parseFloat(v);return !isNaN(n)&&n!==0&&Math.abs(n)<=180;}
function stats_(){var sh=sh_();ensure_(sh);var v=sh.getDataRange().getValues(),h=v[0].map(String);function ix(names){for(var i=0;i<names.length;i++){var p=h.indexOf(names[i]);if(p>=0)return p;}return -1;}var ri=ix(["Respondent ID","Respondent"]),qi=ix(["Question ID","Question"]),ei=ix(["Enumerator ID","Enumerator"]),di=ix(["District","District / Location","Location"]),lati=ix(["Latitude","Lat"]),lngi=ix(["Longitude","Lng","Long"]),afi=ix(["Answer File","Audio File"]),dui=ix(["Drive URL","Drive Url"]),fidi=ix(["File ID"]),tri=ix(["Transcript"]),tsi=ix(["Transcript Status"]),thi=ix(["Theme"]),coi=ix(["Code"]),qui=ix(["Sample Quote"]),ini=ix(["Interpretation"]);var respondents={},questions={},progress={},gps={},enumR={},distR={},answers=0,manifest=[],themeMap={},codeMap={},tx=0,pending=0,failed=0;for(var r=1;r<v.length;r++){var row=v[r],resp=ri>=0?norm_(row[ri]):"",q=qi>=0?norm_(row[qi]):"",en=ei>=0?norm_(row[ei]):"",dist=di>=0?norm_(row[di]):"",lat=lati>=0?norm_(row[lati]):"",lng=lngi>=0?norm_(row[lngi]):"",ans=afi>=0?norm_(row[afi]):"",transcript=tri>=0?norm_(row[tri]):"",st=tsi>=0?norm_(row[tsi]):"";if(!st)st=transcript?"Transcribed":"Pending";if(resp)respondents[resp]=1;if(q)questions[q]=1;if(resp&&q){answers++;if(!progress[resp])progress[resp]={};progress[resp][q]=1;}if(resp&&goodCoord_(lat)&&goodCoord_(lng))gps[resp]=1;if(resp&&en){if(!enumR[en])enumR[en]={};enumR[en][resp]=1;}if(resp&&dist){if(!distR[dist])distR[dist]={};distR[dist][resp]=1;}if(ans){if(transcript||st==="Transcribed")tx++;else if(String(st).indexOf("Failed")>=0||String(st).indexOf("Review")>=0)failed++;else pending++;manifest.push({respondentId:resp,questionId:q,answerFile:ans,driveUrl:dui>=0?norm_(row[dui]):"",fileId:fidi>=0?norm_(row[fidi]):"",transcriptStatus:st});}var th=thi>=0?norm_(row[thi]):"",co=coi>=0?norm_(row[coi]):"",quote=qui>=0?norm_(row[qui]):"",interp=ini>=0?norm_(row[ini]):"";if(th){if(!themeMap[th])themeMap[th]={theme:th,mentions:0,sampleQuote:"",interpretation:""};themeMap[th].mentions++;if(!themeMap[th].sampleQuote&&quote)themeMap[th].sampleQuote=quote;if(!themeMap[th].interpretation&&interp)themeMap[th].interpretation=interp;}if(co)codeMap[co]=(codeMap[co]||0)+1;}var rc=Object.keys(respondents).length,qc=Object.keys(questions).length,fully=0,prog={};Object.keys(progress).forEach(function(k){var n=Object.keys(progress[k]).length;if(qc>0&&n>=qc)fully++;var key=n+" of "+qc;prog[key]=(prog[key]||0)+1;});var enumOut={};Object.keys(enumR).forEach(function(k){enumOut[k]=Object.keys(enumR[k]).length;});var distOut={};Object.keys(distR).forEach(function(k){distOut[k]=Object.keys(distR[k]).length;});return {status:"success",respondents:rc,answers:answers,savedAnswers:answers,questions:qc,completion:rc?Math.round(fully/rc*100):0,gpsCaptured:Object.keys(gps).length,fullyAnswered:fully,enumerators:Object.keys(enumR).length,districts:Object.keys(distR).length,byEnumerator:enumOut,byDistrict:distOut,progressSummary:prog,audioManifest:manifest,transcription:{audioRecords:manifest.length,transcribed:tx,pending:pending,failed:failed,coverage:manifest.length?Math.round(tx/manifest.length*100):0},themeSummary:Object.keys(themeMap).map(function(k){return themeMap[k];}).sort(function(a,b){return b.mentions-a.mentions;}),codeSummary:Object.keys(codeMap).map(function(k){return {code:k,mentions:codeMap[k]};}).sort(function(a,b){return b.mentions-a.mentions;}),updatedAt:new Date().toISOString()};}



function v23Idx_(h,n){for(var i=0;i<n.length;i++){var p=h.indexOf(n[i]);if(p>=0)return p;}return -1;}
function v23Val_(v){return String(v==null?"":v).trim();}
function v23Sent_(t){var l=String(t||"").toLowerCase();var p=/good|excellent|clear|well|efficient|happy|satisfied|improved|success|proceed|thank|better/.test(l);var n=/bad|poor|delay|problem|failed|difficult|expensive|not working|cannot|complain|challenge|late|unclear/.test(l);if(p&&!n)return "Positive";if(n&&!p)return "Negative";return "Neutral";}
function v23Lang_(t){t=String(t||"");var l=t.toLowerCase();if(/[\u0600-\u06FF]/.test(t))return "Arabic";if(/\b(kwa|hivyo|sana|wajaguna|basi)\b/.test(l))return "Swahili/Luganda";if(/[áéíóúñ¿¡]/i.test(t))return "Spanish";if(/\b(le|la|les|avec|pour|bonjour)\b/.test(l))return "French";return "English/Other";}
function v23Qual_(t){t=String(t||"").trim();var w=t?t.split(/\s+/).length:0,s=40;if(w>5)s+=15;if(w>15)s+=15;if(w>40)s+=15;if(/[.!?]/.test(t))s+=5;if(!/test|cheers|clear|hear/i.test(t)||w>10)s+=10;return Math.min(100,s);}
function v23EnsureCols_(){var sh=sh_();ensure_(sh);var h=hdr_(sh);["Sentiment","Language","Quality Score","Research Memo"].forEach(function(x){if(h.indexOf(x)<0){sh.getRange(1,sh.getLastColumn()+1).setValue(x);h.push(x);}});}
function researchIntelligence_(){v23EnsureCols_();var sh=sh_(),v=sh.getDataRange().getValues(),h=v[0].map(String);var ri=v23Idx_(h,["Respondent ID","Respondent"]),qi=v23Idx_(h,["Question ID","Question"]),ei=v23Idx_(h,["Enumerator ID","Enumerator"]),di=v23Idx_(h,["District","District / Location","Location"]),tri=v23Idx_(h,["Transcript"]),ci=v23Idx_(h,["Code"]),thi=v23Idx_(h,["Theme"]),qui=v23Idx_(h,["Sample Quote"]),ini=v23Idx_(h,["Interpretation"]),si=v23Idx_(h,["Sentiment"]),li=v23Idx_(h,["Language"]),qsi=v23Idx_(h,["Quality Score"]);var rec=[],themes={},codes={},sent={Positive:0,Neutral:0,Negative:0},langs={},sum=0,qn=0,resps={},answers=0;for(var r=1;r<v.length;r++){var row=v[r],txt=tri>=0?v23Val_(row[tri]):"",resp=ri>=0?v23Val_(row[ri]):"",q=qi>=0?v23Val_(row[qi]):"";if(resp)resps[resp]=1;if(resp&&q)answers++;if(!txt)continue;var se=v23Sent_(txt),la=v23Lang_(txt),qu=v23Qual_(txt);if(si>=0&&!v23Val_(row[si]))sh.getRange(r+1,si+1).setValue(se);if(li>=0&&!v23Val_(row[li]))sh.getRange(r+1,li+1).setValue(la);if(qsi>=0&&!v23Val_(row[qsi]))sh.getRange(r+1,qsi+1).setValue(qu);sent[se]=(sent[se]||0)+1;langs[la]=(langs[la]||0)+1;sum+=qu;qn++;var theme=thi>=0?v23Val_(row[thi]):"",code=ci>=0?v23Val_(row[ci]):"",quote=qui>=0?v23Val_(row[qui]):"",interp=ini>=0?v23Val_(row[ini]):"";if(theme){if(!themes[theme])themes[theme]={theme:theme,mentions:0,sampleQuote:"",interpretation:""};themes[theme].mentions++;if(!themes[theme].sampleQuote&&quote)themes[theme].sampleQuote=quote;if(!themes[theme].interpretation&&interp)themes[theme].interpretation=interp;}if(code)codes[code]=(codes[code]||0)+1;rec.push({respondentId:resp,questionId:q,enumerator:ei>=0?v23Val_(row[ei]):"",district:di>=0?v23Val_(row[di]):"",transcript:txt,code:code,theme:theme,sentiment:se,language:la,qualityScore:qu});}return {status:"success",version:"V24.1",respondents:Object.keys(resps).length,answers:answers,records:rec,sentimentCounts:sent,languageCounts:langs,averageQuality:qn?Math.round(sum/qn):0,themeSummary:Object.keys(themes).map(k=>themes[k]).sort((a,b)=>b.mentions-a.mentions),codeSummary:Object.keys(codes).map(k=>({code:k,mentions:codes[k]})).sort((a,b)=>b.mentions-a.mentions)};}
function researchReport_(){var i=researchIntelligence_(),r=[];r.push("AI RESEARCH REPORT");r.push("Generated by Audio Research Tool V24.1 Research Intelligence Edition");r.push("");r.push("1. EXECUTIVE SUMMARY");r.push("Total respondents: "+i.respondents);r.push("Total recorded answers: "+i.answers);r.push("Number of coded records: "+i.records.length);r.push("Average interview quality score: "+i.averageQuality+"%");r.push("");r.push("2. SENTIMENT SUMMARY");r.push("Positive: "+(i.sentimentCounts.Positive||0));r.push("Neutral: "+(i.sentimentCounts.Neutral||0));r.push("Negative: "+(i.sentimentCounts.Negative||0));r.push("");r.push("3. LANGUAGE SUMMARY");Object.keys(i.languageCounts||{}).forEach(k=>r.push(k+": "+i.languageCounts[k]));r.push("");r.push("4. MAJOR THEMES");(i.themeSummary||[]).forEach(function(t,n){r.push((n+1)+". "+t.theme+" ("+t.mentions+" mentions)");if(t.sampleQuote)r.push("Sample quote: "+t.sampleQuote);if(t.interpretation)r.push("Interpretation: "+t.interpretation);r.push("");});r.push("5. KEY CODES");(i.codeSummary||[]).forEach((c,n)=>r.push((n+1)+". "+c.code+" - "+c.mentions+" mentions"));r.push("");r.push("6. CONCLUSIONS");r.push("AI-generated findings should be reviewed and validated by the researcher before final reporting.");r.push("");r.push("7. RECOMMENDATIONS");r.push("- Review AI-coded themes and adjust where necessary.");r.push("- Use sample quotes to support qualitative findings.");r.push("- Compare findings by district, enumerator and respondent category where available.");return {status:"success",report:r.join("\n"),intelligence:i};}
var v23OldDoGet=doGet;doGet=function(e){var cb=e&&e.parameter&&(e.parameter.callback||e.parameter.cb),a=e&&e.parameter&&e.parameter.action;try{if(a==="researchIntelligence")return out_(researchIntelligence_(),cb);if(a==="researchReport")return out_(researchReport_(),cb);}catch(err){return out_({status:"error",message:String(err)},cb);}return v23OldDoGet(e);};


/* V24.1 Populated Board-Ready Report Fix */

function researchReport_(){
  var intel = researchIntelligence_();
  var themes = intel.themeSummary || [];
  var codes = intel.codeSummary || [];
  var sent = intel.sentimentCounts || {};
  var langs = intel.languageCounts || {};
  var report = [];

  report.push("BOARD-READY AUDIO RESEARCH REPORT");
  report.push("");
  report.push("Project ID: UVTAB001.6.26");
  report.push("Project Title: DEVELOPMENT OF AUDIO RESEARCH DATA COLLECTION TOOL");
  report.push("Institution: UVTAB");
  report.push("Researcher: Sam Patrick Ogwang");
  report.push("Report Generated: " + new Date().toLocaleString());
  report.push("System Version: Audio Research Tool V24.1");
  report.push("");

  report.push("1. EXECUTIVE SUMMARY");
  report.push("This report presents automatically generated research intelligence from the Audio Research Data Collection Tool. The system collected audio responses, transcribed them using OpenAI Whisper, coded the transcripts, generated themes, extracted sample quotations, produced preliminary interpretations, and summarized the findings for researcher review.");
  report.push("");
  report.push("A total of " + (intel.respondents || 0) + " respondent(s) and " + (intel.answers || 0) + " recorded answer(s) were processed. The system identified " + themes.length + " major theme(s), with an average interview quality score of " + (intel.averageQuality || 0) + "%.");
  report.push("Sentiment distribution: Positive = " + (sent.Positive || 0) + ", Neutral = " + (sent.Neutral || 0) + ", Negative = " + (sent.Negative || 0) + ".");
  report.push("");

  report.push("2. METHODOLOGY");
  report.push("The data were collected using an audio-based research data collection platform. Respondents answered structured audio questions. Responses were recorded, uploaded to Google Drive, logged in Google Sheets, automatically transcribed, and processed through an AI-assisted coding and thematic analysis workflow.");
  report.push("The AI-generated findings are preliminary and should be validated by the researcher before final approval, publication, or submission to management/Board.");
  report.push("");

  report.push("3. KEY FINDINGS AND THEMATIC ANALYSIS");
  if(themes.length){
    themes.forEach(function(t, i){
      report.push("3." + (i+1) + " Theme: " + (t.theme || "Unnamed Theme"));
      report.push("Frequency: " + (t.mentions || 0) + " mention(s)");
      if(t.sampleQuote) report.push("Sample Quote: \"" + t.sampleQuote + "\"");
      if(t.interpretation) report.push("Interpretation: " + t.interpretation);
      report.push("");
    });
  } else {
    report.push("No themes were available at the time of report generation.");
    report.push("");
  }

  report.push("4. SENTIMENT ANALYSIS");
  report.push("Positive responses: " + (sent.Positive || 0));
  report.push("Neutral responses: " + (sent.Neutral || 0));
  report.push("Negative responses: " + (sent.Negative || 0));
  report.push("");

  report.push("5. LANGUAGE SUMMARY");
  var langKeys = Object.keys(langs);
  if(langKeys.length){ langKeys.forEach(function(k){ report.push(k + ": " + langs[k] + " response(s)"); }); }
  else { report.push("No language summary was available."); }
  report.push("");

  report.push("6. CODE SUMMARY");
  if(codes.length){ codes.forEach(function(c, i){ report.push((i+1) + ". " + c.code + " - " + c.mentions + " mention(s)"); }); }
  else { report.push("No code summary was available."); }
  report.push("");

  report.push("7. INTERVIEW QUALITY");
  report.push("Average quality score: " + (intel.averageQuality || 0) + "%");
  report.push("The quality score is based on transcript completeness, response length, and clarity indicators. Low scores should be reviewed against the original audio files.");
  report.push("");

  report.push("8. CONCLUSIONS");
  if(themes.length){
    report.push("The strongest emerging themes are: " + themes.slice(0,5).map(function(t){return t.theme;}).join(", ") + ".");
    report.push("These findings demonstrate that the Audio Research Tool can support rapid qualitative data collection, transcription, coding, and synthesis for institutional research and monitoring.");
  } else {
    report.push("More coded data are required for substantive conclusions.");
  }
  report.push("");

  report.push("9. RECOMMENDATIONS");
  report.push("1. Review all AI-generated codes, themes, sample quotes and interpretations before final approval.");
  report.push("2. Use the tool for structured qualitative research, tracer studies, stakeholder interviews and field monitoring.");
  report.push("3. Develop future versions with direct Word/PDF export, advanced charts and theme matrix analysis.");
  report.push("4. Check low quality score responses against the source audio before inclusion in final reporting.");
  report.push("");

  report.push("10. APPENDIX: TOP TRANSCRIPT RECORDS");
  (intel.records || []).slice(0,20).forEach(function(r, i){
    report.push((i+1) + ". Respondent: " + (r.respondentId || "") + " | Question: " + (r.questionId || "") + " | Theme: " + (r.theme || "") + " | Sentiment: " + (r.sentiment || ""));
    report.push("Transcript: " + (r.transcript || "").substring(0,500));
    report.push("");
  });

  return {status:"success", report:report.join("\n"), intelligence:intel};
}
var v232OldDoGet = doGet;
doGet = function(e){
  var cb=e&&e.parameter&&(e.parameter.callback||e.parameter.cb), a=e&&e.parameter&&e.parameter.action;
  try{ if(a==="researchReport")return out_(researchReport_(), cb); }
  catch(err){ return out_({status:"error",message:String(err)}, cb); }
  return v232OldDoGet(e);
};



/* ===================== V24.1 Local Whisper Transcription Mode Backend ===================== */
function transcriptionEngine_(){
  return String(PropertiesService.getScriptProperties().getProperty("TRANSCRIPTION_ENGINE") || "OPENAI").toUpperCase();
}
function localWebhookUrl_(){
  return String(PropertiesService.getScriptProperties().getProperty("LOCAL_TRANSCRIPTION_WEBHOOK_URL") || "").trim();
}
var v24OriginalTranscribeFile_ = transcribeFile_;
transcribeFile_ = function(fid){
  var engine = transcriptionEngine_();
  if(engine === "LOCAL_MANIFEST"){
    throw new Error("Local manifest mode enabled. Audio was saved. Download the local transcription manifest and transcribe offline.");
  }
  if(engine === "LOCAL_WEBHOOK"){
    var hook = localWebhookUrl_();
    if(!hook) throw new Error("LOCAL_TRANSCRIPTION_WEBHOOK_URL not configured.");
    var file = DriveApp.getFileById(fid);
    var payload = {
      fileId: fid,
      fileName: file.getName(),
      driveUrl: file.getUrl(),
      note: "Local server should download or receive this file and return JSON {status:'success', transcript:'...'}"
    };
    var res = UrlFetchApp.fetch(hook, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    if(res.getResponseCode()<200 || res.getResponseCode()>=300){
      throw new Error("Local webhook error: " + res.getContentText());
    }
    var obj = JSON.parse(res.getContentText());
    if(obj.status !== "success") throw new Error(obj.message || "Local webhook failed.");
    return obj.transcript || "";
  }
  return v24OriginalTranscribeFile_(fid);
};

var v24OldDoGet = doGet;
doGet = function(e){
  var cb=e&&e.parameter&&(e.parameter.callback||e.parameter.cb), a=e&&e.parameter&&e.parameter.action;
  try{
    if(a==="transcriptionStatus"){
      return out_({
        status:"success",
        version:"V24.1",
        engine:"Local Whisper / OpenAI",
        transcriptionEngine:transcriptionEngine_(),
        apiKeyConfigured:!!PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY"),
        autoTranscribeOnSubmit:autoTx_ ? autoTx_() : true,
        autoCodeOnTranscribe:autoCode_ ? autoCode_() : true,
        localWebhookConfigured:!!localWebhookUrl_(),
        localWebhookUrl:localWebhookUrl_()? "Configured" : "Not configured"
      }, cb);
    }
  }catch(err){return out_({status:"error",message:String(err)}, cb);}
  return v24OldDoGet(e);
};
/* =================== End V24.1 Local Whisper Backend =================== */


/* V24.1 Local AI Backend Status */
function v241Engine_(){return String(PropertiesService.getScriptProperties().getProperty("TRANSCRIPTION_ENGINE") || "OPENAI").toUpperCase();}
var v241OldDoGet = doGet;
doGet = function(e){
  var cb=e&&e.parameter&&(e.parameter.callback||e.parameter.cb), a=e&&e.parameter&&e.parameter.action;
  try{
    if(a==="transcriptionStatus"){
      return out_({status:"success",version:"V24.1",engine:"Local AI / OpenAI",transcriptionEngine:v241Engine_(),apiKeyConfigured:!!PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY"),autoTranscribeOnSubmit:(typeof autoTx_==="function"?autoTx_():false),autoCodeOnTranscribe:(typeof autoCode_==="function"?autoCode_():true),recommendation:"For free transcription set TRANSCRIPTION_ENGINE=LOCAL_MANIFEST and AUTO_TRANSCRIBE_ON_SUBMIT=No"}, cb);
    }
  }catch(err){return out_({status:"error",message:String(err)}, cb);}
  return v241OldDoGet(e);
};

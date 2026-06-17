const SHEET_ID = "1R-cjiNMoNWv-BhokRXeeBCmHtssr4-LMQwYGTd39g-0";
const FOLDER_ID = "13lzW_JzwCiLaYXDGWV77Muo90FjNV5au";

function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
function getSheet_(){return SpreadsheetApp.openById(SHEET_ID).getSheets()[0];}

function doGet(e){
  try{
    const action=e&&e.parameter&&e.parameter.action;
    if(action==="transcriptionStatus"){
      return json_({status:"success",version:"V20.1",engine:"OpenAI Whisper",apiKeyConfigured:!!PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY")});
    }
    if(action==="stats")return getStats_();
    return json_({status:"success",message:"Audio Research Backend V20.1 running"});
  }catch(err){return json_({status:"error",message:String(err)});}
}

function getHeaders_(sheet){
  if(sheet.getLastRow()===0)return [];
  return sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(),1)).getValues()[0].map(String);
}
function col_(headers,names){
  for(let i=0;i<names.length;i++){let p=headers.indexOf(names[i]);if(p>=0)return p+1;}
  return -1;
}
function ensureColumns_(){
  const sheet=getSheet_();
  if(sheet.getLastRow()===0)return;
  const need=["Transcript","Transcript Status","Transcribed By","Transcription Date","Code","Theme","Auto Suggested","Sample Quote","Interpretation","Researcher Review"];
  const h=getHeaders_(sheet);
  need.forEach(n=>{if(h.indexOf(n)<0){sheet.getRange(1,sheet.getLastColumn()+1).setValue(n);h.push(n);}});
}

function doPost(e){
  try{
    const data=JSON.parse(e.postData.contents);
    if(data.action==="transcribeAudio")return transcribeAudio_(data);
    if(data.action==="updateTranscript")return updateTranscript_(data);
    return json_({status:"error",message:"Unknown V20.1 action. Merge this into your full submission backend if needed."});
  }catch(err){return json_({status:"error",message:String(err)});}
}

function updateTranscript_(data){
  ensureColumns_();
  const sheet=getSheet_();
  const h=getHeaders_(sheet);
  const respCol=col_(h,["Respondent ID","Respondent"]);
  const qCol=col_(h,["Question ID","Question"]);
  const fileCol=col_(h,["File ID"]);
  const answerCol=col_(h,["Answer File"]);
  const trCol=col_(h,["Transcript"]);
  const stCol=col_(h,["Transcript Status"]);
  const byCol=col_(h,["Transcribed By"]);
  const dtCol=col_(h,["Transcription Date"]);
  if(trCol<0)return json_({status:"error",message:"Transcript column missing"});
  const targetFileId=String(data.fileId||"").trim();
  const targetResp=String(data.respondentId||"").trim();
  const targetQ=String(data.questionId||"").trim();
  const targetAnswer=String(data.answerFile||"").trim();
  let updated=0, matchedBy="";
  for(let r=2;r<=sheet.getLastRow();r++){
    const rowFile=fileCol>0?String(sheet.getRange(r,fileCol).getValue()||"").trim():"";
    const rowResp=respCol>0?String(sheet.getRange(r,respCol).getValue()||"").trim():"";
    const rowQ=qCol>0?String(sheet.getRange(r,qCol).getValue()||"").trim():"";
    const rowAnswer=answerCol>0?String(sheet.getRange(r,answerCol).getValue()||"").trim():"";
    let match=false;
    if(targetFileId&&rowFile&&targetFileId===rowFile){match=true;matchedBy="File ID";}
    else if(targetAnswer&&rowAnswer&&targetAnswer===rowAnswer){match=true;matchedBy="Answer File";}
    else if(targetResp&&targetQ&&rowResp===targetResp&&String(rowQ)===String(targetQ)){match=true;matchedBy="Respondent ID + Question ID";}
    if(match){
      sheet.getRange(r,trCol).setValue(data.transcript||"");
      if(stCol>0)sheet.getRange(r,stCol).setValue(data.transcriptStatus||"Transcribed");
      if(byCol>0)sheet.getRange(r,byCol).setValue(data.transcribedBy||"OpenAI Whisper");
      if(dtCol>0)sheet.getRange(r,dtCol).setValue(data.transcriptionDate||new Date().toISOString());
      updated++;
    }
  }
  return json_({status:"success",updated:updated,matchedBy:matchedBy});
}

function transcribeAudio_(data){
  try{
    const key=PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
    if(!key)return json_({status:"error",message:"OPENAI_API_KEY not configured in Apps Script Properties"});
    if(!data.fileId)return json_({status:"error",message:"fileId missing from selected audio record"});
    const file=DriveApp.getFileById(data.fileId);
    const blob=file.getBlob();
    const boundary="----ARDC"+new Date().getTime();
    const part1="--"+boundary+"\r\nContent-Disposition: form-data; name=\"model\"\r\n\r\nwhisper-1\r\n--"+boundary+"\r\nContent-Disposition: form-data; name=\"file\"; filename=\""+file.getName()+"\"\r\nContent-Type: "+blob.getContentType()+"\r\n\r\n";
    const part2="\r\n--"+boundary+"--";
    const payload=Utilities.newBlob(part1).getBytes().concat(blob.getBytes()).concat(Utilities.newBlob(part2).getBytes());
    const res=UrlFetchApp.fetch("https://api.openai.com/v1/audio/transcriptions",{method:"post",headers:{Authorization:"Bearer "+key},contentType:"multipart/form-data; boundary="+boundary,payload:payload,muteHttpExceptions:true});
    const code=res.getResponseCode(), body=res.getContentText();
    if(code<200||code>=300)return json_({status:"error",message:"OpenAI error: "+body});
    const transcript=JSON.parse(body).text||"";
    const save=JSON.parse(updateTranscript_({respondentId:data.respondentId,questionId:data.questionId,fileId:data.fileId,answerFile:data.answerFile,transcript:transcript,transcribedBy:"OpenAI Whisper",transcriptionDate:new Date().toISOString(),transcriptStatus:"Transcribed"}).getContent());
    if(!save.updated||Number(save.updated)<1)return json_({status:"error",message:"Transcript was created but no Google Sheet row was updated. Check File ID/Answer File matching.",transcript:transcript,updated:0});
    return json_({status:"success",transcript:transcript,updated:save.updated,matchedBy:save.matchedBy});
  }catch(err){return json_({status:"error",message:String(err)});}
}

function getStats_(){
  ensureColumns_();
  const sheet=getSheet_();
  const values=sheet.getDataRange().getValues();
  if(values.length<2)return json_({status:"success",audioManifest:[],answers:0,transcription:{audioRecords:0,transcribed:0,pending:0,failed:0,coverage:0},updatedAt:new Date().toISOString()});
  const h=values[0].map(String);
  const resp=col_(h,["Respondent ID"])-1, q=col_(h,["Question ID"])-1, af=col_(h,["Answer File"])-1, du=col_(h,["Drive URL"])-1, fid=col_(h,["File ID"])-1, tr=col_(h,["Transcript"])-1, ts=col_(h,["Transcript Status"])-1;
  const manifest=[];let transcribed=0,pending=0,failed=0;
  for(let r=1;r<values.length;r++){
    if(af>=0&&values[r][af]){
      const transcript=tr>=0?String(values[r][tr]||"").trim():"";
      let status=ts>=0?String(values[r][ts]||"").trim():"";
      if(!status)status=transcript?"Transcribed":"Pending";
      if(status==="Transcribed"||transcript)transcribed++;else if(status==="Failed"||status==="Error")failed++;else pending++;
      manifest.push({respondentId:resp>=0?String(values[r][resp]||""):"",questionId:q>=0?String(values[r][q]||""):"",answerFile:String(values[r][af]||""),driveUrl:du>=0?String(values[r][du]||""):"",fileId:fid>=0?String(values[r][fid]||""):"",transcriptStatus:status});
    }
  }
  return json_({status:"success",audioManifest:manifest,answers:manifest.length,transcription:{audioRecords:manifest.length,transcribed:transcribed,pending:pending,failed:failed,coverage:manifest.length?Math.round(transcribed/manifest.length*100):0},updatedAt:new Date().toISOString()});
}

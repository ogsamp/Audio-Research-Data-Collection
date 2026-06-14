const SHEET_ID = "1R-cjiNMoNWv-BhokRXeeBCmHtssr4-LMQwYGTd39g-0";
const FOLDER_ID = "13lzW_JzwCiLaYXDGWV77Muo90FjNV5au";

function getSheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
}

function jsonOutput_(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + json + ");").setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function normalize_(s) {
  return String(s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function findColumn_(headers, possibleNames) {
  const normalizedHeaders = headers.map(normalize_);
  for (let i = 0; i < possibleNames.length; i++) {
    const target = normalize_(possibleNames[i]);
    const idx = normalizedHeaders.indexOf(target);
    if (idx >= 0) return idx;
  }
  for (let h = 0; h < normalizedHeaders.length; h++) {
    for (let i = 0; i < possibleNames.length; i++) {
      const target = normalize_(possibleNames[i]);
      if (target && normalizedHeaders[h].indexOf(target) >= 0) return h;
    }
  }
  return -1;
}

function doGet(e) {
  const callback = e && e.parameter && e.parameter.callback;
  try {
    const action = e && e.parameter && e.parameter.action;
    if (action === "stats") return getStats_(callback);
    return jsonOutput_({status:"success", message:"Audio Research Backend V10.4 is running"}, callback);
  } catch (err) {
    return jsonOutput_({status:"error", message:String(err)}, callback);
  }
}

function getStats_(callback) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return jsonOutput_({status:"success", respondents:0, answers:0, questions:0, completion:0, gpsCaptured:0, enumerators:0, districts:0, byEnumerator:{}, byDistrict:{}}, callback);
  }

  const headers = values[0].map(String);
  const respondentIdx = findColumn_(headers, ["Respondent ID","RespondentID","Respondent","Participant ID","ParticipantID","Interviewee ID"]);
  const questionIdx = findColumn_(headers, ["Question ID","QuestionID","Question No","Question Number","QID"]);
  const enumIdx = findColumn_(headers, ["Enumerator ID","EnumeratorID","Enumerator","Interviewer ID","Interviewer","Field Officer"]);
  const districtIdx = findColumn_(headers, ["District","District / Location","District Location","Location","Area","Site","Region"]);
  const latIdx = findColumn_(headers, ["Latitude","Lat","GPS Latitude","GPS Lat"]);
  const longIdx = findColumn_(headers, ["Longitude","Long","Lng","GPS Longitude","GPS Long","GPS Lng"]);
  const answerIdx = findColumn_(headers, ["Answer File","Answer Audio","Answer","Audio Answer","Answer File Name"]);
  const driveIdx = findColumn_(headers, ["Drive URL","Drive Link","Google Drive URL","File URL","Audio URL"]);

  const respondents = {}, questions = {}, byEnumerator = {}, byDistrict = {};
  let answers = 0, gpsCaptured = 0;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const resp = respondentIdx >= 0 ? String(row[respondentIdx] || "").trim() : "";
    const q = questionIdx >= 0 ? String(row[questionIdx] || "").trim() : "";
    const en = enumIdx >= 0 ? String(row[enumIdx] || "").trim() : "";
    const dist = districtIdx >= 0 ? String(row[districtIdx] || "").trim() : "";
    const lat = latIdx >= 0 ? String(row[latIdx] || "").trim() : "";
    const lng = longIdx >= 0 ? String(row[longIdx] || "").trim() : "";
    const ans = answerIdx >= 0 ? String(row[answerIdx] || "").trim() : "";
    const drive = driveIdx >= 0 ? String(row[driveIdx] || "").trim() : "";

    if (resp) respondents[resp] = true;
    if (q) questions[q] = true;
    if ((resp && q) || ans || drive) answers++;
    if (lat || lng) gpsCaptured++;
    if (en) byEnumerator[en] = (byEnumerator[en] || 0) + 1;
    if (dist) byDistrict[dist] = (byDistrict[dist] || 0) + 1;
  }

  const respondentCount = Object.keys(respondents).length;
  const questionCount = Object.keys(questions).length;
  const expected = respondentCount * questionCount;
  const completion = expected ? Math.round((answers / expected) * 100) : 0;

  return jsonOutput_({
    status:"success",
    respondents:respondentCount,
    answers:answers,
    questions:questionCount,
    completion:completion,
    gpsCaptured:gpsCaptured,
    enumerators:Object.keys(byEnumerator).length,
    districts:Object.keys(byDistrict).length,
    byEnumerator:byEnumerator,
    byDistrict:byDistrict,
    detectedHeaders:{
      respondent: respondentIdx >= 0 ? headers[respondentIdx] : "",
      question: questionIdx >= 0 ? headers[questionIdx] : "",
      enumerator: enumIdx >= 0 ? headers[enumIdx] : "",
      district: districtIdx >= 0 ? headers[districtIdx] : "",
      latitude: latIdx >= 0 ? headers[latIdx] : "",
      longitude: longIdx >= 0 ? headers[longIdx] : ""
    },
    updatedAt:new Date().toISOString()
  }, callback);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const sheet = getSheet_();

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Submitted At","Started At","Duration Seconds","Project ID","Project Title","Institution","Researcher","Researcher Contact","Survey Title",
        "Respondent ID","Enumerator ID","District","Latitude","Longitude","GPS Accuracy","GPS Timestamp","Consent","Question ID","Question Text",
        "Question Audio","Answer File","Drive URL","File ID","Client Timestamp","Notes"
      ]);
    }

    const fileMap = {};
    (data.files || []).forEach(file => {
      const bytes = Utilities.base64Decode(file.base64);
      const blob = Utilities.newBlob(bytes, file.mimeType || "audio/webm", file.fileName);
      const saved = folder.createFile(blob);
      fileMap[file.fileName] = {url:saved.getUrl(), id:saved.getId()};
    });

    (data.metadata || []).forEach(row => {
      const f = fileMap[row.answerFileName] || {};
      sheet.appendRow([
        data.submittedAt, data.startedAt || "", data.durationSeconds || "", data.projectId || "", data.projectTitle || "", data.institution || "",
        data.researcherName || "", data.researcherContact || "", data.surveyTitle || "", data.respondentId || "", data.enumeratorId || row.enumeratorId || "",
        data.district || row.district || "", data.latitude || row.latitude || "", data.longitude || row.longitude || "", data.gpsAccuracy || row.gpsAccuracy || "",
        data.gpsTimestamp || row.gpsTimestamp || "", data.consent || "", row.questionId, row.questionText, row.questionAudio, row.answerFileName,
        f.url || "", f.id || "", row.timestamp, data.notes || row.notes || ""
      ]);
    });

    return jsonOutput_({status:"success", rows:(data.metadata || []).length}, null);
  } catch (err) {
    return jsonOutput_({status:"error", message:String(err)}, null);
  }
}

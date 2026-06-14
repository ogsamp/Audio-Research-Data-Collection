const SHEET_ID = "PASTE_YOUR_GOOGLE_SHEET_ID_HERE";
const FOLDER_ID = "PASTE_YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE";

function getSheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
}

function doGet(e) {
  try {
    const action = e && e.parameter && e.parameter.action;
    if (action === "stats") {
      return getStats_();
    }
    return ContentService
      .createTextOutput(JSON.stringify({status:"success", message:"Audio Research Backend V10.2 is running"}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({status:"error", message:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getStats_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status:"success",
        respondents:0,
        answers:0,
        questions:0,
        completion:0,
        gpsCaptured:0,
        byEnumerator:{},
        byDistrict:{}
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const headers = values[0].map(String);
  function idx(name) { return headers.indexOf(name); }

  const respondentIdx = idx("Respondent ID");
  const questionIdx = idx("Question ID");
  const enumIdx = idx("Enumerator ID");
  const districtIdx = idx("District");
  const latIdx = idx("Latitude");
  const longIdx = idx("Longitude");

  const respondents = {};
  const questions = {};
  const byEnumerator = {};
  const byDistrict = {};
  let answers = 0;
  let gpsCaptured = 0;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const resp = respondentIdx >= 0 ? String(row[respondentIdx] || "").trim() : "";
    const q = questionIdx >= 0 ? String(row[questionIdx] || "").trim() : "";
    const en = enumIdx >= 0 ? String(row[enumIdx] || "").trim() : "";
    const dist = districtIdx >= 0 ? String(row[districtIdx] || "").trim() : "";
    const lat = latIdx >= 0 ? String(row[latIdx] || "").trim() : "";
    const lng = longIdx >= 0 ? String(row[longIdx] || "").trim() : "";

    if (resp) respondents[resp] = true;
    if (q) questions[q] = true;
    if (resp && q) answers++;
    if (lat || lng) gpsCaptured++;

    if (en) byEnumerator[en] = (byEnumerator[en] || 0) + 1;
    if (dist) byDistrict[dist] = (byDistrict[dist] || 0) + 1;
  }

  const respondentCount = Object.keys(respondents).length;
  const questionCount = Object.keys(questions).length;
  const expected = respondentCount * questionCount;
  const completion = expected ? Math.round((answers / expected) * 100) : 0;

  return ContentService
    .createTextOutput(JSON.stringify({
      status:"success",
      respondents:respondentCount,
      answers:answers,
      questions:questionCount,
      completion:completion,
      gpsCaptured:gpsCaptured,
      byEnumerator:byEnumerator,
      byDistrict:byDistrict,
      updatedAt:new Date().toISOString()
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const sheet = getSheet_();

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Submitted At","Started At","Duration Seconds",
        "Project ID","Project Title","Institution","Researcher","Researcher Contact","Survey Title",
        "Respondent ID","Enumerator ID","District","Latitude","Longitude","GPS Accuracy","GPS Timestamp",
        "Consent","Question ID","Question Text","Question Audio","Answer File","Drive URL","File ID",
        "Client Timestamp","Notes"
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
        data.submittedAt,
        data.startedAt || "",
        data.durationSeconds || "",
        data.projectId || "",
        data.projectTitle || "",
        data.institution || "",
        data.researcherName || "",
        data.researcherContact || "",
        data.surveyTitle || "",
        data.respondentId || "",
        data.enumeratorId || row.enumeratorId || "",
        data.district || row.district || "",
        data.latitude || row.latitude || "",
        data.longitude || row.longitude || "",
        data.gpsAccuracy || row.gpsAccuracy || "",
        data.gpsTimestamp || row.gpsTimestamp || "",
        data.consent || "",
        row.questionId,
        row.questionText,
        row.questionAudio,
        row.answerFileName,
        f.url || "",
        f.id || "",
        row.timestamp,
        data.notes || row.notes || ""
      ]);
    });

    return ContentService
      .createTextOutput(JSON.stringify({status:"success", rows:(data.metadata || []).length}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({status:"error", message:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

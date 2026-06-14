const SHEET_ID = "PASTE_YOUR_GOOGLE_SHEET_ID_HERE";
const FOLDER_ID = "PASTE_YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];

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

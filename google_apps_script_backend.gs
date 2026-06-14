const SHEET_ID = "1R-cjiNMoNWv-BhokRXeeBCmHtssr4-LMQwYGTd39g-0";
const FOLDER_ID = "13lzW_JzwCiLaYXDGWV77Muo90FjNV5au";

const HEADERS = [
  "Submitted At","Started At","Duration Seconds",
  "Project ID","Project Title","Institution","Researcher","Researcher Contact","Survey Title",
  "Respondent ID","Enumerator ID","District",
  "Latitude","Longitude","GPS Accuracy","GPS Timestamp",
  "Consent","Question ID","Question Text","Question Audio",
  "Answer File","Drive URL","File ID",
  "Client Timestamp","Notes",
  "Transcript","Code","Theme","Sample Quote","Interpretation"
];

function getSheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
}

function jsonOutput_(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
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

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    return HEADERS.slice();
  }
  let lastCol = Math.max(sheet.getLastColumn(), HEADERS.length);
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const norm = headers.map(normalize_);
  const hasRespondent = norm.indexOf(normalize_("Respondent ID")) >= 0;
  const hasQuestion = norm.indexOf(normalize_("Question ID")) >= 0;

  if (!headers.join("").trim() || !hasRespondent || !hasQuestion) {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    return HEADERS.slice();
  }

  HEADERS.forEach(h => {
    if (norm.indexOf(normalize_(h)) < 0) {
      headers.push(h);
      sheet.getRange(1, headers.length).setValue(h);
      norm.push(normalize_(h));
    }
  });
  return headers;
}

function rowFromObject_(headers, obj) {
  return headers.map(h => obj[h] !== undefined ? obj[h] : "");
}

function doGet(e) {
  const callback = e && e.parameter && e.parameter.callback;
  try {
    const action = e && e.parameter && e.parameter.action;
    if (action === "stats") return getStats_(callback, e);
    if (action === "setupColumns") return setupThemeColumns_(callback);
    return jsonOutput_({status:"success", message:"Audio Research Backend V10.7 is running"}, callback);
  } catch (err) {
    return jsonOutput_({status:"error", message:String(err)}, callback);
  }
}

function setupThemeColumns_(callback) {
  const sheet = getSheet_();
  const headers = ensureHeaders_(sheet);
  return jsonOutput_({status:"success", message:"Theme analysis columns are ready", headers:headers}, callback);
}

function getStats_(callback, e) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return jsonOutput_({
      status:"success", respondents:0, answers:0, questions:0, completion:0,
      gpsCaptured:0, fullyAnswered:0, byQuestionsAnswered:{},
      enumerators:0, districts:0, byEnumerator:{}, byDistrict:{}, themes:{}, invalidRowsIgnored:0
    }, callback);
  }

  const headers = values[0].map(String);
  const respondentIdx = findColumn_(headers, ["Respondent ID","RespondentID","Respondent","Participant ID","ParticipantID","Interviewee ID"]);
  const questionIdx   = findColumn_(headers, ["Question ID","QuestionID","Question No","Question Number","QID"]);
  const enumIdx       = findColumn_(headers, ["Enumerator ID","EnumeratorID","Enumerator","Interviewer ID","Interviewer","Field Officer"]);
  const districtIdx   = findColumn_(headers, ["District","District / Location","District Location","Location","Area","Site","Region"]);
  const latIdx        = findColumn_(headers, ["Latitude","Lat","GPS Latitude","GPS Lat"]);
  const longIdx       = findColumn_(headers, ["Longitude","Long","Lng","GPS Longitude","GPS Long","GPS Lng"]);
  const answerIdx     = findColumn_(headers, ["Answer File","Answer Audio","Answer","Audio Answer","Answer File Name"]);
  const driveIdx      = findColumn_(headers, ["Drive URL","Drive Link","Google Drive URL","File URL","Audio URL"]);
  const transcriptIdx = findColumn_(headers, ["Transcript","Transcription","Text Transcript","Response Transcript"]);
  const codeIdx       = findColumn_(headers, ["Code","Codes","Coding","Qualitative Code"]);
  const themeIdx      = findColumn_(headers, ["Theme","Themes","Category","Finding Theme"]);
  const quoteIdx      = findColumn_(headers, ["Sample Quote","Quote","Representative Quote","Illustrative Quote"]);
  const interpIdx     = findColumn_(headers, ["Interpretation","Meaning","Finding","Explanation"]);

  const expectedQuestions = Number((e && e.parameter && e.parameter.expectedQuestions) || 0);
  const respondentQuestions = {};
  const respondentHasGps = {};
  const byEnumerator = {};
  const byDistrict = {};
  const themes = {};
  let answers = 0;
  let invalidRowsIgnored = 0;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const resp = respondentIdx >= 0 ? String(row[respondentIdx] || "").trim() : "";
    const qRaw = questionIdx >= 0 ? String(row[questionIdx] || "").trim() : "";
    const en = enumIdx >= 0 ? String(row[enumIdx] || "").trim() : "";
    const dist = districtIdx >= 0 ? String(row[districtIdx] || "").trim() : "";
    const lat = latIdx >= 0 ? String(row[latIdx] || "").trim() : "";
    const lng = longIdx >= 0 ? String(row[longIdx] || "").trim() : "";
    const ans = answerIdx >= 0 ? String(row[answerIdx] || "").trim() : "";
    const drive = driveIdx >= 0 ? String(row[driveIdx] || "").trim() : "";
    const transcript = transcriptIdx >= 0 ? String(row[transcriptIdx] || "").trim() : "";
    const code = codeIdx >= 0 ? String(row[codeIdx] || "").trim() : "";
    const theme = themeIdx >= 0 ? String(row[themeIdx] || "").trim() : "";
    const quote = quoteIdx >= 0 ? String(row[quoteIdx] || "").trim() : "";
    const interpretation = interpIdx >= 0 ? String(row[interpIdx] || "").trim() : "";

    const match = qRaw.match(/\d+/);
    const qNum = match ? Number(match[0]) : 0;
    const validQuestionRange = qNum > 0 && (!expectedQuestions || qNum <= expectedQuestions);
    const hasAnswerEvidence = !!(ans || drive);
    const valid = !!resp && validQuestionRange && hasAnswerEvidence;

    if (!valid) {
      invalidRowsIgnored++;
      continue;
    }

    if (!respondentQuestions[resp]) respondentQuestions[resp] = {};
    respondentQuestions[resp][qNum] = true;
    answers++;

    if (lat || lng) respondentHasGps[resp] = true;

    if (en) {
      if (!byEnumerator[en]) byEnumerator[en] = {};
      byEnumerator[en][resp] = true;
    }
    if (dist) {
      if (!byDistrict[dist]) byDistrict[dist] = {};
      byDistrict[dist][resp] = true;
    }

    // Theme summary: count coded rows. Prefer Theme column; fall back to Code.
    const themeKey = theme || code;
    if (themeKey) {
      if (!themes[themeKey]) {
        themes[themeKey] = {mentions:0, sampleQuote:"", interpretation:"", codes:{}};
      }
      themes[themeKey].mentions++;
      if (!themes[themeKey].sampleQuote) {
        themes[themeKey].sampleQuote = quote || transcript.substring(0, 180);
      }
      if (!themes[themeKey].interpretation) {
        themes[themeKey].interpretation = interpretation || "";
      }
      if (code) themes[themeKey].codes[code] = true;
    }
  }

  Object.keys(themes).forEach(k => {
    themes[k].codes = Object.keys(themes[k].codes).join("; ");
  });

  const respondentIds = Object.keys(respondentQuestions);
  const respondentCount = respondentIds.length;
  const questionCount = expectedQuestions || Math.max.apply(null, respondentIds.map(id => Object.keys(respondentQuestions[id]).length).concat([0]));
  const expected = respondentCount * questionCount;
  const completion = expected ? Math.round((answers / expected) * 100) : 0;

  let fullyAnswered = 0;
  const byQuestionsAnswered = {};
  respondentIds.forEach(id => {
    const count = Object.keys(respondentQuestions[id]).length;
    byQuestionsAnswered[count + " of " + questionCount] = (byQuestionsAnswered[count + " of " + questionCount] || 0) + 1;
    if (count >= questionCount) fullyAnswered++;
  });

  const enumCounts = {};
  Object.keys(byEnumerator).forEach(k => enumCounts[k] = Object.keys(byEnumerator[k]).length);

  const districtCounts = {};
  Object.keys(byDistrict).forEach(k => districtCounts[k] = Object.keys(byDistrict[k]).length);

  return jsonOutput_({
    status:"success",
    respondents:respondentCount,
    answers:answers,
    questions:questionCount,
    completion:completion,
    gpsCaptured:Object.keys(respondentHasGps).length,
    fullyAnswered:fullyAnswered,
    byQuestionsAnswered:byQuestionsAnswered,
    enumerators:Object.keys(enumCounts).length,
    districts:Object.keys(districtCounts).length,
    byEnumerator:enumCounts,
    byDistrict:districtCounts,
    themes:themes,
    invalidRowsIgnored:invalidRowsIgnored,
    detectedHeaders:{
      respondent: respondentIdx >= 0 ? headers[respondentIdx] : "",
      question: questionIdx >= 0 ? headers[questionIdx] : "",
      enumerator: enumIdx >= 0 ? headers[enumIdx] : "",
      district: districtIdx >= 0 ? headers[districtIdx] : "",
      latitude: latIdx >= 0 ? headers[latIdx] : "",
      longitude: longIdx >= 0 ? headers[longIdx] : "",
      transcript: transcriptIdx >= 0 ? headers[transcriptIdx] : "",
      code: codeIdx >= 0 ? headers[codeIdx] : "",
      theme: themeIdx >= 0 ? headers[themeIdx] : ""
    },
    updatedAt:new Date().toISOString()
  }, callback);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const sheet = getSheet_();
    const headers = ensureHeaders_(sheet);

    const fileMap = {};
    (data.files || []).forEach(file => {
      const bytes = Utilities.base64Decode(file.base64);
      const blob = Utilities.newBlob(bytes, file.mimeType || "audio/webm", file.fileName);
      const saved = folder.createFile(blob);
      fileMap[file.fileName] = {url:saved.getUrl(), id:saved.getId()};
    });

    (data.metadata || []).forEach(row => {
      const f = fileMap[row.answerFileName] || {};
      const record = {
        "Submitted At": data.submittedAt || "",
        "Started At": data.startedAt || "",
        "Duration Seconds": data.durationSeconds || "",
        "Project ID": data.projectId || "",
        "Project Title": data.projectTitle || "",
        "Institution": data.institution || "",
        "Researcher": data.researcherName || "",
        "Researcher Contact": data.researcherContact || "",
        "Survey Title": data.surveyTitle || "",
        "Respondent ID": data.respondentId || row.respondentId || "",
        "Enumerator ID": data.enumeratorId || row.enumeratorId || "",
        "District": data.district || row.district || "",
        "Latitude": data.latitude || row.latitude || "",
        "Longitude": data.longitude || row.longitude || "",
        "GPS Accuracy": data.gpsAccuracy || row.gpsAccuracy || "",
        "GPS Timestamp": data.gpsTimestamp || row.gpsTimestamp || "",
        "Consent": data.consent || "",
        "Question ID": row.questionId || "",
        "Question Text": row.questionText || "",
        "Question Audio": row.questionAudio || "",
        "Answer File": row.answerFileName || "",
        "Drive URL": f.url || "",
        "File ID": f.id || "",
        "Client Timestamp": row.timestamp || "",
        "Notes": data.notes || row.notes || "",
        "Transcript": "",
        "Code": "",
        "Theme": "",
        "Sample Quote": "",
        "Interpretation": ""
      };
      sheet.appendRow(rowFromObject_(headers, record));
    });

    return jsonOutput_({status:"success", rows:(data.metadata || []).length}, null);
  } catch (err) {
    return jsonOutput_({status:"error", message:String(err)}, null);
  }
}

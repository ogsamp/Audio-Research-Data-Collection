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
  "Transcript","Transcript Status","Transcribed By","Transcription Date",
  "Code","Theme","Auto Suggested Theme","Sample Quote","Interpretation","Researcher Review"
];

const THEME_RULES = [
  {theme:"Practical Skills", keywords:["practice","practical","hands-on","hands on","machine","equipment","maintenance","workshop","skill","skills","competence","doing","perform"]},
  {theme:"ICT Competence", keywords:["computer","ict","digital","system","online","software","internet","data","technology","typing","spreadsheet","database"]},
  {theme:"Industrial Attachment", keywords:["attachment","industry","workplace","company","factory","employer","supervisor","field placement","internship"]},
  {theme:"Assessment Quality", keywords:["assessment","exam","test","marking","moderation","verification","certification","quality assurance","standard"]},
  {theme:"Training Relevance", keywords:["curriculum","training","relevant","course","learning","module","lesson","content","trade"]},
  {theme:"Challenges", keywords:["challenge","challenges","problem","difficulty","lack","shortage","delay","expensive","cost","limited","barrier"]},
  {theme:"Support Needed", keywords:["support","need","needs","recommend","improve","provide","facilitate","fund","equipment","training materials"]},
  {theme:"Employment and Livelihoods", keywords:["job","employment","income","business","self employment","enterprise","livelihood","market"]},
  {theme:"Access and Inclusion", keywords:["access","inclusion","gender","disability","youth","rural","equity","opportunity","marginalized"]},
  {theme:"Quality Assurance", keywords:["quality","assurance","monitoring","audit","standard","compliance","verification","improvement"]},
  {theme:"Soft Skills", keywords:["communication","teamwork","discipline","attitude","confidence","leadership","customer care","time management"]},
  {theme:"Resources and Equipment", keywords:["tools","equipment","materials","machines","workshop","laboratory","facility","facilities","resources"]}
];

const STOP_WORDS = {
  "the":true,"and":true,"for":true,"that":true,"this":true,"with":true,"have":true,"has":true,"was":true,"were":true,"are":true,"from":true,
  "you":true,"your":true,"they":true,"their":true,"them":true,"our":true,"ours":true,"but":true,"not":true,"can":true,"will":true,
  "had":true,"been":true,"into":true,"about":true,"also":true,"very":true,"there":true,"here":true,"when":true,"what":true,"which":true,
  "then":true,"than":true,"because":true,"would":true,"could":true,"should":true,"these":true,"those":true,"each":true,"where":true,
  "how":true,"why":true,"who":true,"his":true,"her":true,"him":true,"she":true,"he":true,"it":true,"is":true,"to":true,"of":true,
  "in":true,"on":true,"a":true,"an":true,"as":true,"by":true,"or":true,"at":true,"we":true,"i":true,"my":true
};

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

function suggestThemes_(text) {
  const t = String(text || "").toLowerCase();
  const found = [];
  THEME_RULES.forEach(rule => {
    let score = 0;
    rule.keywords.forEach(k => {
      if (t.indexOf(k.toLowerCase()) >= 0) score++;
    });
    if (score > 0) found.push({theme:rule.theme, score:score});
  });
  found.sort((a,b)=>b.score-a.score);
  return found.slice(0,3).map(x=>x.theme);
}

function addWords_(freq, text) {
  const words = String(text || "").toLowerCase().replace(/[^a-z0-9\s]/g, " ").split(/\s+/);
  words.forEach(w => {
    if (!w || w.length < 3 || STOP_WORDS[w]) return;
    freq[w] = (freq[w] || 0) + 1;
  });
}

function sortedFrequency_(freq) {
  const out = {};
  Object.entries(freq).sort((a,b)=>b[1]-a[1]).slice(0,100).forEach(([k,v]) => out[k]=v);
  return out;
}

function splitThemes_(s) {
  return String(s || "").split(/[;,|]/).map(x=>x.trim()).filter(Boolean);
}

function doGet(e) {
  const callback = e && e.parameter && e.parameter.callback;
  try {
    const action = e && e.parameter && e.parameter.action;
    if (action === "stats") return getStats_(callback, e);
    if (action === "setupColumns") return setupColumns_(callback);
    if (action === "suggestThemes") return writeSuggestedThemes_(callback);
    return jsonOutput_({status:"success", message:"Audio Research Backend V15.0 is running"}, callback);
  } catch (err) {
    return jsonOutput_({status:"error", message:String(err)}, callback);
  }
}

function setupColumns_(callback) {
  const sheet = getSheet_();
  const headers = ensureHeaders_(sheet);
  return jsonOutput_({status:"success", message:"V15.0 GPS map and advanced qualitative analysis columns are ready", headers:headers}, callback);
}

function writeSuggestedThemes_(callback) {
  const sheet = getSheet_();
  const headers = ensureHeaders_(sheet);
  const values = sheet.getDataRange().getValues();

  const transcriptIdx = findColumn_(headers, ["Transcript","Transcription","Text Transcript","Response Transcript"]);
  const autoThemeIdx = findColumn_(headers, ["Auto Suggested Theme","Suggested Theme","AI Theme"]);
  const themeIdx = findColumn_(headers, ["Theme","Themes","Category","Finding Theme"]);
  const transcriptStatusIdx = findColumn_(headers, ["Transcript Status","Transcription Status"]);
  const transcriptionDateIdx = findColumn_(headers, ["Transcription Date"]);

  if (transcriptIdx < 0 || autoThemeIdx < 0) {
    return jsonOutput_({status:"error", message:"Transcript or Auto Suggested Theme column not found"}, callback);
  }

  let updated = 0;
  for (let r = 1; r < values.length; r++) {
    const transcript = String(values[r][transcriptIdx] || "").trim();
    if (!transcript) continue;
    const suggestions = suggestThemes_(transcript);
    if (!suggestions.length) continue;
    const suggestedText = suggestions.join("; ");
    sheet.getRange(r + 1, autoThemeIdx + 1).setValue(suggestedText);
    if (themeIdx >= 0 && !String(values[r][themeIdx] || "").trim()) {
      sheet.getRange(r + 1, themeIdx + 1).setValue(suggestedText);
    }
    if (transcriptStatusIdx >= 0) sheet.getRange(r + 1, transcriptStatusIdx + 1).setValue("Transcribed");
    if (transcriptionDateIdx >= 0 && !String(values[r][transcriptionDateIdx] || "").trim()) {
      sheet.getRange(r + 1, transcriptionDateIdx + 1).setValue(new Date());
    }
    updated++;
  }

  return jsonOutput_({status:"success", updated:updated, message:"Multi-theme suggestions written"}, callback);
}

function getStats_(callback, e) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return jsonOutput_({
      status:"success", respondents:0, answers:0, questions:0, completion:0,
      gpsCaptured:0, fullyAnswered:0, byQuestionsAnswered:{}, enumerators:0, districts:0,
      byEnumerator:{}, byDistrict:{}, themes:{}, suggestedThemes:{}, themeMatrix:{}, wordFrequency:{}, audioManifest:[],
      transcription:{audioRecords:0, transcribed:0, pending:0, coverage:0}, invalidRowsIgnored:0
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
  const fileIdIdx     = findColumn_(headers, ["File ID","FileID","Drive File ID"]);
  const transcriptIdx = findColumn_(headers, ["Transcript","Transcription","Text Transcript","Response Transcript"]);
  const transcriptStatusIdx = findColumn_(headers, ["Transcript Status","Transcription Status"]);
  const codeIdx       = findColumn_(headers, ["Code","Codes","Coding","Qualitative Code"]);
  const themeIdx      = findColumn_(headers, ["Theme","Themes","Category","Finding Theme"]);
  const autoThemeIdx  = findColumn_(headers, ["Auto Suggested Theme","Suggested Theme","AI Theme"]);
  const quoteIdx      = findColumn_(headers, ["Sample Quote","Quote","Representative Quote","Illustrative Quote"]);
  const interpIdx     = findColumn_(headers, ["Interpretation","Meaning","Finding","Explanation"]);
  const qTextIdx      = findColumn_(headers, ["Question Text","Question"]);
  const gpsAccIdx     = findColumn_(headers, ["GPS Accuracy","Accuracy","GPSAccuracy"]);
  const gpsTimeIdx    = findColumn_(headers, ["GPS Timestamp","Location Timestamp","GPSTimestamp"]);
  const submittedIdx  = findColumn_(headers, ["Submitted At","SubmittedAt","Timestamp"]);

  const expectedQuestions = Number((e && e.parameter && e.parameter.expectedQuestions) || 0);
  const respondentQuestions = {};
  const respondentHasGps = {};
  const byEnumerator = {};
  const byDistrict = {};
  const themes = {};
  const suggestedThemes = {};
  const themeMatrix = {};
  const wordFrequency = {};
  const audioManifest = [];
  const gpsByRespondent = {};
  let answers = 0;
  let transcripts = 0;
  let invalidRowsIgnored = 0;

  function addTheme(container, key, quote, interpretation, code) {
    if (!key) return;
    if (!container[key]) container[key] = {mentions:0, sampleQuote:"", interpretation:"", codes:{}};
    container[key].mentions++;
    if (!container[key].sampleQuote) container[key].sampleQuote = quote || "";
    if (!container[key].interpretation) container[key].interpretation = interpretation || "";
    if (code) container[key].codes[code] = true;
  }

  function addMatrix(theme, resp, qNum, code) {
    if (!theme) return;
    if (!themeMatrix[theme]) themeMatrix[theme] = {mentions:0, respondents:{}, questions:{}, codes:{}};
    themeMatrix[theme].mentions++;
    if (resp) themeMatrix[theme].respondents[resp] = true;
    if (qNum) themeMatrix[theme].questions[qNum] = true;
    if (code) themeMatrix[theme].codes[code] = true;
  }

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const resp = respondentIdx >= 0 ? String(row[respondentIdx] || "").trim() : "";
    const qRaw = questionIdx >= 0 ? String(row[questionIdx] || "").trim() : "";
    const qText = qTextIdx >= 0 ? String(row[qTextIdx] || "").trim() : "";
    const en = enumIdx >= 0 ? String(row[enumIdx] || "").trim() : "";
    const dist = districtIdx >= 0 ? String(row[districtIdx] || "").trim() : "";
    const lat = latIdx >= 0 ? String(row[latIdx] || "").trim() : "";
    const lng = longIdx >= 0 ? String(row[longIdx] || "").trim() : "";
    const ans = answerIdx >= 0 ? String(row[answerIdx] || "").trim() : "";
    const drive = driveIdx >= 0 ? String(row[driveIdx] || "").trim() : "";
    const fileId = fileIdIdx >= 0 ? String(row[fileIdIdx] || "").trim() : "";
    const transcript = transcriptIdx >= 0 ? String(row[transcriptIdx] || "").trim() : "";
    const transcriptStatus = transcriptStatusIdx >= 0 ? String(row[transcriptStatusIdx] || "").trim() : "";
    const code = codeIdx >= 0 ? String(row[codeIdx] || "").trim() : "";
    const theme = themeIdx >= 0 ? String(row[themeIdx] || "").trim() : "";
    const autoTheme = autoThemeIdx >= 0 ? String(row[autoThemeIdx] || "").trim() : "";
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

    if (lat || lng) {
      respondentHasGps[resp] = true;
      if (!gpsByRespondent[resp]) {
        gpsByRespondent[resp] = {
          respondentId:resp,
          enumerator:en,
          district:dist,
          latitude:lat,
          longitude:lng,
          gpsAccuracy:gpsAccIdx >= 0 ? String(row[gpsAccIdx] || "") : "",
          gpsTimestamp:gpsTimeIdx >= 0 ? String(row[gpsTimeIdx] || "") : "",
          submittedAt:submittedIdx >= 0 ? String(row[submittedIdx] || "") : "",
          questionsAnswered:0
        };
      }
    }

    if (en) {
      if (!byEnumerator[en]) byEnumerator[en] = {};
      byEnumerator[en][resp] = true;
    }
    if (dist) {
      if (!byDistrict[dist]) byDistrict[dist] = {};
      byDistrict[dist][resp] = true;
    }

    if (transcript) {
      transcripts++;
      addWords_(wordFrequency, transcript);
    }

    audioManifest.push({
      respondentId:resp,
      questionId:qNum,
      questionText:qText,
      answerFile:ans,
      driveUrl:drive,
      fileId:fileId,
      transcriptStatus: transcript ? "Transcribed" : (transcriptStatus || "Pending")
    });

    const manualThemes = splitThemes_(theme || code);
    manualThemes.forEach(t => {
      addTheme(themes, t, quote || transcript.substring(0, 180), interpretation, code);
      addMatrix(t, resp, qNum, code);
    });

    const autoThemes = splitThemes_(autoTheme);
    const suggestions = autoThemes.length ? autoThemes : suggestThemes_(transcript);
    suggestions.forEach(t => {
      addTheme(suggestedThemes, t, quote || transcript.substring(0, 180), "Auto-suggested from transcript keywords. Researcher validation required.", code);
      addMatrix(t, resp, qNum, code);
    });
  }

  Object.keys(themes).forEach(k => themes[k].codes = Object.keys(themes[k].codes).join("; "));
  Object.keys(suggestedThemes).forEach(k => suggestedThemes[k].codes = Object.keys(suggestedThemes[k].codes).join("; "));

  const flatMatrix = {};
  Object.keys(themeMatrix).forEach(k => {
    flatMatrix[k] = {
      mentions: themeMatrix[k].mentions,
      respondents: Object.keys(themeMatrix[k].respondents).length,
      questions: Object.keys(themeMatrix[k].questions).length,
      codes: Object.keys(themeMatrix[k].codes).join("; ")
    };
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

  Object.keys(gpsByRespondent).forEach(id => {
    gpsByRespondent[id].questionsAnswered = respondentQuestions[id] ? Object.keys(respondentQuestions[id]).length : 0;
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
    suggestedThemes:suggestedThemes,
    themeMatrix:flatMatrix,
    wordFrequency:sortedFrequency_(wordFrequency),
    audioManifest:audioManifest,
    gpsRecords:Object.values(gpsByRespondent),
    transcription:{
      audioRecords:answers,
      transcribed:transcripts,
      pending:Math.max(answers - transcripts, 0),
      coverage: answers ? Math.round((transcripts / answers) * 100) : 0
    },
    invalidRowsIgnored:invalidRowsIgnored,
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
        "Transcript Status": "Pending",
        "Transcribed By": "",
        "Transcription Date": "",
        "Code": "",
        "Theme": "",
        "Auto Suggested Theme": "",
        "Sample Quote": "",
        "Interpretation": "",
        "Researcher Review": ""
      };
      sheet.appendRow(rowFromObject_(headers, record));
    });

    return jsonOutput_({status:"success", rows:(data.metadata || []).length}, null);
  } catch (err) {
    return jsonOutput_({status:"error", message:String(err)}, null);
  }
}

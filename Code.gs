const SHEET_NAME = "GameData";

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ok:true, message:"Math Quest API is running"}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        "Timestamp","Player","Level","Question",
        "Correct Answer","User Answer","Correct",
        "Score Earned","Total Score","Streak","Hint Used"
      ]);
    }

    sheet.appendRow([
      new Date(),
      data.player || "Guest",
      data.level || "",
      data.question || "",
      data.correctAnswer ?? "",
      data.userAnswer ?? "",
      data.correct ? "TRUE" : "FALSE",
      data.scoreEarned ?? 0,
      data.totalScore ?? 0,
      data.streak ?? 0,
      data.hintUsed ? "YES" : "NO"
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ok:true}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ok:false,error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 版本四：Google 帳號與各自存檔
const FILE_NAME = "幾A幾B_個人遊戲存檔";

function doGet(e) {
  const action = e.parameter.action;
  
  if (action === "load") {
    return ContentService.createTextOutput(JSON.stringify(loadPersonalData()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Invalid action" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === "save") {
      savePersonalData(data.games);
      return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 取得或建立當前使用者的個人試算表
function getOrCreateUserSpreadsheet() {
  const files = DriveApp.getFilesByName(FILE_NAME);
  let ss;
  
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    // 建立新試算表於使用者個人的 Google 雲端硬碟
    ss = SpreadsheetApp.create(FILE_NAME);
    let sheet = ss.getActiveSheet();
    sheet.setName("GameData");
    sheet.appendRow(["gameId", "secretAnswer", "status", "guessHistoryJSON", "updatedAt"]);
    sheet.getRange("B:B").setNumberFormat("@"); // 保留 4 位數字首位 0
  }
  
  let sheet = ss.getSheetByName("GameData");
  if (!sheet) {
    sheet = ss.insertSheet("GameData");
    sheet.appendRow(["gameId", "secretAnswer", "status", "guessHistoryJSON", "updatedAt"]);
    sheet.getRange("B:B").setNumberFormat("@");
  }
  return sheet;
}

function loadPersonalData() {
  try {
    const sheet = getOrCreateUserSpreadsheet();
    const data = sheet.getDataRange().getValues();
    const games = [];
    
    for (let i = 1; i < data.length; i++) {
      games.push({
        id: data[i][0],
        secretAnswer: String(data[i][1]).padStart(4, '0'), // 確保 0 開頭不丟失
        status: data[i][2],
        history: JSON.parse(data[i][3] || "[]")
      });
    }
    const userEmail = Session.getActiveUser().getEmail();
    return { status: "success", games: games, userEmail: userEmail };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function savePersonalData(games) {
  const sheet = getOrCreateUserSpreadsheet();
  const data = sheet.getDataRange().getValues();
  
  games.forEach(game => {
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(game.id)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    const formattedAnswer = "'" + String(game.secretAnswer).padStart(4, '0');
    const historyJSON = JSON.stringify(game.history);
    const now = new Date();
    
    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 2).setValue(formattedAnswer);
      sheet.getRange(rowIndex, 3).setValue(game.status);
      sheet.getRange(rowIndex, 4).setValue(historyJSON);
      sheet.getRange(rowIndex, 5).setValue(now);
    } else {
      sheet.appendRow([game.id, formattedAnswer, game.status, historyJSON, now]);
    }
  });
}
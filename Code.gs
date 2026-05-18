const SHEET_HEADERS = ["timestamp", "senderName", "userId", "text", "sourceType"];

function doPost(e) {
  const data = JSON.parse(e.postData.contents || "{}");

  if (Array.isArray(data.events)) {
    handleLineWebhook(data.events);
    return ContentService.createTextOutput("ok");
  }

  handlePagesMessage(data);
  return ContentService.createTextOutput("ok");
}

function doGet(e) {
  const callback = sanitizeCallback(e.parameter.callback || "callback");
  const limit = Math.min(Number(e.parameter.limit || 50), 50);
  const messages = getLatestMessages(limit);
  const body = callback + "(" + JSON.stringify({ messages: messages }) + ");";

  return ContentService
    .createTextOutput(body)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function handleLineWebhook(events) {
  events.forEach(function (event) {
    if (!event.message || event.message.type !== "text") {
      return;
    }

    const userId = event.source && event.source.userId ? event.source.userId : "";
    const sourceType = event.source && event.source.type ? event.source.type : "";
    const senderName = getDisplayName(event.source, userId);
    const timestamp = new Date(event.timestamp || Date.now()).toISOString();

    appendMessage({
      timestamp: timestamp,
      senderName: senderName,
      userId: userId,
      text: event.message.text,
      sourceType: sourceType
    });
  });
}

function handlePagesMessage(data) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("CHANNEL_ACCESS_TOKEN");
  const groupId = props.getProperty("GROUP_ID");
  const text = String(data.text || data.body || "");

  appendMessage({
    timestamp: new Date().toISOString(),
    senderName: String(data.senderName || data.sender || "やんさん"),
    userId: "github-pages",
    text: text,
    sourceType: "githubPages"
  });

  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
    method: "post",
    headers: {
      Authorization: "Bearer " + token,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify({
      to: groupId,
      messages: [{ type: "text", text: text }]
    }),
    muteHttpExceptions: true
  });
}

function appendMessage(message) {
  const sheet = getSheet();
  sheet.appendRow([
    message.timestamp,
    message.senderName,
    message.userId,
    message.text,
    message.sourceType
  ]);
}

function getLatestMessages(limit) {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  const count = Math.min(limit, lastRow - 1);
  const startRow = lastRow - count + 1;
  const values = sheet.getRange(startRow, 1, count, SHEET_HEADERS.length).getValues();

  return values.map(function (row) {
    return {
      timestamp: row[0],
      senderName: row[1],
      userId: row[2],
      text: row[3],
      sourceType: row[4]
    };
  });
}

function getDisplayName(source, userId) {
  if (!userId) {
    return "";
  }
  if (!source) {
    return userId;
  }

  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("CHANNEL_ACCESS_TOKEN");
  let url = "";

  if (source.type === "group" && source.groupId) {
    url = "https://api.line.me/v2/bot/group/" + source.groupId + "/member/" + userId;
  } else if (source.type === "room" && source.roomId) {
    url = "https://api.line.me/v2/bot/room/" + source.roomId + "/member/" + userId;
  } else {
    url = "https://api.line.me/v2/bot/profile/" + userId;
  }

  const response = UrlFetchApp.fetch(url, {
    method: "get",
    headers: {
      Authorization: "Bearer " + token
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    return userId;
  }

  const profile = JSON.parse(response.getContentText() || "{}");
  return profile.displayName || userId;
}

function getSheet() {
  const sheetId = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  const sheet = SpreadsheetApp.openById(sheetId).getSheets()[0];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(SHEET_HEADERS);
  }

  return sheet;
}

function sanitizeCallback(callback) {
  if (/^[A-Za-z_$][0-9A-Za-z_$]*(\.[A-Za-z_$][0-9A-Za-z_$]*)*$/.test(callback)) {
    return callback;
  }
  return "callback";
}

function keepWarm() {
  const token = PropertiesService
    .getScriptProperties()
    .getProperty("CHANNEL_ACCESS_TOKEN");

  UrlFetchApp.fetch("https://api.line.me/v2/bot/info", {
    method: "get",
    headers: {
      Authorization: "Bearer " + token
    },
    muteHttpExceptions: true
  });
}

function testLineToken() {
  const token = PropertiesService
    .getScriptProperties()
    .getProperty("CHANNEL_ACCESS_TOKEN");

  const response = UrlFetchApp.fetch("https://api.line.me/v2/bot/info", {
    method: "get",
    headers: {
      Authorization: "Bearer " + token
    },
    muteHttpExceptions: true
  });

  console.log(response.getResponseCode());
  console.log(response.getContentText());
}

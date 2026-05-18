const MESSAGE_SHEET_NAME = "messages";
const GROUP_SHEET_NAME = "groups";
const SHEET_HEADERS = ["timestamp", "senderName", "userId", "text", "sourceType", "groupId"];
const GROUP_HEADERS = ["groupId", "groupName", "lastSeenAt"];

function doPost(e) {
  const data = JSON.parse(e.postData.contents || "{}");

  if (Array.isArray(data.events)) {
    handleLineWebhook(data.events);
    return ContentService.createTextOutput("ok");
  }

  if (data.action === "deleteGroup") {
    deleteGroup(String(data.groupId || ""));
    return ContentService.createTextOutput("ok");
  }

  handlePagesMessage(data);
  return ContentService.createTextOutput("ok");
}

function doGet(e) {
  const callback = sanitizeCallback(e.parameter.callback || "callback");
  const action = e.parameter.action || "messages";
  const defaultGroupId = PropertiesService.getScriptProperties().getProperty("GROUP_ID") || "";
  const groupId = String(e.parameter.groupId || defaultGroupId);
  const payload = action === "groups"
    ? { groups: getGroups() }
    : { messages: getLatestMessages(Math.min(Number(e.parameter.limit || 50), 50), groupId) };
  const body = callback + "(" + JSON.stringify(payload) + ");";

  return ContentService
    .createTextOutput(body)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function handleLineWebhook(events) {
  events.forEach(function (event) {
    if (event.source && event.source.groupId) {
      upsertGroup(event.source.groupId);
    }

    if (!event.message || event.message.type !== "text") {
      return;
    }

    const userId = event.source && event.source.userId ? event.source.userId : "";
    const sourceType = event.source && event.source.type ? event.source.type : "";
    const groupId = event.source && event.source.groupId ? event.source.groupId : "";
    const senderName = getDisplayName(event.source, userId);
    const timestamp = new Date(event.timestamp || Date.now()).toISOString();

    appendMessage({
      timestamp: timestamp,
      senderName: senderName,
      userId: userId,
      text: event.message.text,
      sourceType: sourceType,
      groupId: groupId
    });
  });
}

function handlePagesMessage(data) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("CHANNEL_ACCESS_TOKEN");
  const groupId = String(data.groupId || props.getProperty("GROUP_ID") || "");
  const text = String(data.text || data.body || "");

  appendMessage({
    timestamp: new Date().toISOString(),
    senderName: String(data.senderName || data.sender || "やんさん"),
    userId: "github-pages",
    text: text,
    sourceType: "githubPages",
    groupId: groupId
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
    message.sourceType,
    message.groupId
  ]);
}

function getLatestMessages(limit, groupId) {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  const values = sheet.getRange(2, 1, lastRow - 1, SHEET_HEADERS.length).getValues();
  const matched = [];

  for (let i = values.length - 1; i >= 0 && matched.length < limit; i -= 1) {
    const row = values[i];
    if (String(row[5] || "") !== groupId) {
      continue;
    }
    matched.unshift({
      timestamp: row[0],
      senderName: row[1],
      userId: row[2],
      text: row[3],
      sourceType: row[4],
      groupId: row[5]
    });
  }

  return matched;
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
  const spreadsheet = SpreadsheetApp.openById(sheetId);
  const sheet = spreadsheet.getSheetByName(MESSAGE_SHEET_NAME) || spreadsheet.getSheets()[0];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(SHEET_HEADERS);
  } else {
    ensureHeaders(sheet, SHEET_HEADERS);
  }

  return sheet;
}

function getGroupSheet() {
  const sheetId = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  const spreadsheet = SpreadsheetApp.openById(sheetId);
  const sheet = spreadsheet.getSheetByName(GROUP_SHEET_NAME) || spreadsheet.insertSheet(GROUP_SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(GROUP_HEADERS);
  } else {
    ensureHeaders(sheet, GROUP_HEADERS);
  }

  return sheet;
}

function ensureHeaders(sheet, headers) {
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  let changed = false;

  headers.forEach(function (header, index) {
    if (current[index] !== header) {
      current[index] = header;
      changed = true;
    }
  });

  if (changed) {
    sheet.getRange(1, 1, 1, headers.length).setValues([current]);
  }
}

function upsertGroup(groupId) {
  const sheet = getGroupSheet();
  const lastRow = sheet.getLastRow();
  const groupName = getGroupName(groupId);
  const lastSeenAt = new Date().toISOString();

  if (lastRow > 1) {
    const values = sheet.getRange(2, 1, lastRow - 1, GROUP_HEADERS.length).getValues();
    for (let i = 0; i < values.length; i += 1) {
      if (values[i][0] === groupId) {
        sheet.getRange(i + 2, 2, 1, 2).setValues([[groupName || values[i][1] || groupId, lastSeenAt]]);
        return;
      }
    }
  }

  sheet.appendRow([groupId, groupName || groupId, lastSeenAt]);
}

function getGroups() {
  const sheet = getGroupSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  return sheet.getRange(2, 1, lastRow - 1, GROUP_HEADERS.length).getValues()
    .map(function (row) {
      return {
        groupId: row[0],
        groupName: row[1],
        lastSeenAt: row[2]
      };
    })
    .sort(function (a, b) {
      return new Date(b.lastSeenAt).getTime() - new Date(a.lastSeenAt).getTime();
    });
}

function deleteGroup(groupId) {
  if (!groupId) {
    return;
  }

  const sheet = getGroupSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, GROUP_HEADERS.length).getValues();
  for (let i = values.length - 1; i >= 0; i -= 1) {
    if (values[i][0] === groupId) {
      sheet.deleteRow(i + 2);
    }
  }
}

function getGroupName(groupId) {
  const token = PropertiesService.getScriptProperties().getProperty("CHANNEL_ACCESS_TOKEN");
  const response = UrlFetchApp.fetch("https://api.line.me/v2/bot/group/" + groupId + "/summary", {
    method: "get",
    headers: {
      Authorization: "Bearer " + token
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    return groupId;
  }

  const summary = JSON.parse(response.getContentText() || "{}");
  return summary.groupName || groupId;
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

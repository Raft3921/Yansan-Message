const MESSAGE_SHEET_NAME = "messages";
const GROUP_SHEET_NAME = "groups";
const DELETED_GROUP_SHEET_NAME = "deletedGroups";
const IMAGE_FOLDER_PROPERTY = "IMAGE_FOLDER_ID";
const SHEET_HEADERS = ["time", "sender", "userId", "groupId", "groupName", "text", "source", "messageType", "imageUrl", "stickerPackageId", "stickerId", "stickerResourceType"];
const LEGACY_SHEET_HEADERS = ["timestamp", "senderName", "userId", "text", "sourceType", "groupId"];
const GROUP_HEADERS = ["groupId", "groupName", "lastSeenAt"];
const DELETED_GROUP_HEADERS = ["groupId", "deletedAt"];

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
  const payload = getPayload(action, e.parameter, groupId);
  const body = callback + "(" + JSON.stringify(payload) + ");";

  return ContentService
    .createTextOutput(body)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function getPayload(action, parameter, groupId) {
  if (action === "groups") {
    return { groups: getGroups() };
  }

  if (action === "sendMessage") {
    return handlePagesMessage({
      text: parameter.text,
      sender: parameter.sender,
      groupId: parameter.groupId || groupId
    });
  }

  if (action === "imageData") {
    return getImageDataPayload(String(parameter.imageUrl || ""));
  }

  return { messages: getMessages(Math.min(Number(parameter.limit || 50), 50), groupId) };
}

function handleLineWebhook(events) {
  events.forEach(function (event) {
    if (event.source && event.source.groupId) {
      upsertGroup(event.source.groupId);
    }

    if (!event.message || ["text", "image", "sticker"].indexOf(event.message.type) === -1) {
      return;
    }

    const userId = event.source && event.source.userId ? event.source.userId : "";
    const sourceType = event.source && event.source.type ? event.source.type : "";
    const groupId = event.source && event.source.groupId ? event.source.groupId : "";
    const groupName = groupId ? getGroupName(groupId) : "";
    const senderName = getDisplayName(event.source, userId);
    const timestamp = new Date(event.timestamp || Date.now()).toISOString();
    const messageType = event.message.type;
    const imageUrl = messageType === "image" ? saveLineImage(event.message) : "";
    const stickerPackageId = messageType === "sticker" ? String(event.message.packageId || "") : "";
    const stickerId = messageType === "sticker" ? String(event.message.stickerId || "") : "";
    const stickerResourceType = messageType === "sticker" ? String(event.message.stickerResourceType || "") : "";

    appendMessage({
      time: timestamp,
      sender: senderName,
      userId: userId,
      groupId: groupId,
      groupName: groupName,
      text: getMessageText(event.message),
      source: sourceType,
      messageType: messageType,
      imageUrl: imageUrl,
      stickerPackageId: stickerPackageId,
      stickerId: stickerId,
      stickerResourceType: stickerResourceType
    });
  });
}

function handlePagesMessage(data) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("CHANNEL_ACCESS_TOKEN");
  const groupId = String(data.groupId || props.getProperty("GROUP_ID") || "");
  const groupName = getStoredGroupName(groupId);
  const text = String(data.text || data.body || "");

  if (!groupId || !text) {
    return {
      ok: false,
      error: "missing_fields"
    };
  }

  appendMessage({
    time: new Date().toISOString(),
    sender: String(data.senderName || data.sender || "やんさん"),
    userId: "github-pages",
    groupId: groupId,
    groupName: groupName,
    text: text,
    source: "githubPages",
    messageType: "text",
    imageUrl: "",
    stickerPackageId: "",
    stickerId: "",
    stickerResourceType: ""
  });

  const response = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
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

  return {
    ok: response.getResponseCode() >= 200 && response.getResponseCode() < 300,
    lineStatus: response.getResponseCode()
  };
}

function appendMessage(message) {
  const sheet = getSheet();
  sheet.appendRow([
    message.time,
    message.sender,
    message.userId,
    message.groupId,
    message.groupName,
    message.text,
    message.source,
    message.messageType || "text",
    message.imageUrl || "",
    message.stickerPackageId || "",
    message.stickerId || "",
    message.stickerResourceType || ""
  ]);
}

function getMessages(limit, groupId) {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  const values = sheet.getRange(2, 1, lastRow - 1, SHEET_HEADERS.length).getValues();
  const matched = [];

  for (let i = values.length - 1; i >= 0 && matched.length < limit; i -= 1) {
    const row = values[i];
    if (String(row[3] || "") !== groupId) {
      continue;
    }
    matched.unshift({
      time: row[0],
      sender: row[1],
      userId: row[2],
      groupId: row[3],
      groupName: row[4],
      text: row[5],
      source: row[6],
      messageType: row[7] || "text",
      imageUrl: row[8] || "",
      stickerPackageId: row[9] || "",
      stickerId: row[10] || "",
      stickerResourceType: row[11] || ""
    });
  }

  return matched;
}

function getMessageText(message) {
  if (message.type === "text") {
    return message.text || "";
  }
  if (message.type === "image") {
    return "写真が届きました";
  }
  if (message.type === "sticker") {
    return "スタンプが届きました";
  }
  return "";
}

function saveLineImage(message) {
  const provider = message.contentProvider || {};
  if (provider.type === "external" && provider.originalContentUrl) {
    return provider.originalContentUrl;
  }

  const token = PropertiesService.getScriptProperties().getProperty("CHANNEL_ACCESS_TOKEN");
  const response = UrlFetchApp.fetch(
    "https://api-data.line.me/v2/bot/message/" + encodeURIComponent(message.id) + "/content",
    {
      method: "get",
      headers: {
        Authorization: "Bearer " + token
      },
      muteHttpExceptions: true
    }
  );

  if (response.getResponseCode() !== 200) {
    return "";
  }

  const folder = getImageFolder();
  const contentType = getHeaderValue(response.getHeaders(), "Content-Type") || "image/jpeg";
  const extension = contentType.indexOf("png") !== -1 ? ".png" : ".jpg";
  const file = folder.createFile(response.getBlob().setName("line-image-" + message.id + extension));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return getDriveImageUrl(file);
}

function getDriveImageUrl(file) {
  const resourceKey = file.getResourceKey();
  return "https://drive.google.com/thumbnail?sz=w1200&id=" + encodeURIComponent(file.getId())
    + (resourceKey ? "&resourcekey=" + encodeURIComponent(resourceKey) : "");
}

function getImageDataPayload(imageUrl) {
  const dataUrl = getImageDataUrl(imageUrl);
  return {
    ok: Boolean(dataUrl),
    imageUrl: imageUrl,
    imageDataUrl: dataUrl
  };
}

function getImageDataUrl(imageUrl) {
  try {
    const driveFileId = getDriveFileIdFromUrl(imageUrl);
    if (driveFileId) {
      const file = DriveApp.getFileById(driveFileId);
      const blob = file.getBlob();
      return blobToDataUrl(blob);
    }

    if (/^https?:\/\//i.test(imageUrl)) {
      const response = UrlFetchApp.fetch(imageUrl, { muteHttpExceptions: true });
      if (response.getResponseCode() >= 200 && response.getResponseCode() < 300) {
        return blobToDataUrl(response.getBlob());
      }
    }
  } catch (error) {
    return "";
  }

  return "";
}

function blobToDataUrl(blob) {
  const contentType = blob.getContentType() || "image/jpeg";
  return "data:" + contentType + ";base64," + Utilities.base64Encode(blob.getBytes());
}

function getDriveFileIdFromUrl(url) {
  const idMatch = String(url || "").match(/[?&]id=([^&]+)/);
  if (idMatch) {
    return decodeURIComponent(idMatch[1]);
  }

  const pathMatch = String(url || "").match(/\/file\/d\/([^/]+)/);
  return pathMatch ? decodeURIComponent(pathMatch[1]) : "";
}

function getHeaderValue(headers, name) {
  const lowerName = name.toLowerCase();
  const keys = Object.keys(headers || {});
  for (let i = 0; i < keys.length; i += 1) {
    if (String(keys[i]).toLowerCase() === lowerName) {
      return headers[keys[i]];
    }
  }
  return "";
}

function getImageFolder() {
  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty(IMAGE_FOLDER_PROPERTY);
  if (folderId) {
    return DriveApp.getFolderById(folderId);
  }

  const folder = DriveApp.createFolder("やんさんチャット画像");
  props.setProperty(IMAGE_FOLDER_PROPERTY, folder.getId());
  return folder;
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
    migrateMessageSheet(sheet);
  }

  return sheet;
}

function migrateMessageSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const current = sheet.getRange(1, 1, 1, Math.max(SHEET_HEADERS.length, LEGACY_SHEET_HEADERS.length)).getValues()[0];
  const isNew = SHEET_HEADERS.every(function (header, index) {
    return current[index] === header;
  });

  if (isNew) {
    return;
  }

  const isLegacy = LEGACY_SHEET_HEADERS.every(function (header, index) {
    return current[index] === header;
  });

  if (!isLegacy) {
    ensureHeaders(sheet, SHEET_HEADERS);
    return;
  }

  const rowCount = lastRow - 1;
  const legacyValues = rowCount > 0
    ? sheet.getRange(2, 1, rowCount, LEGACY_SHEET_HEADERS.length).getValues()
    : [];
  const migrated = legacyValues.map(function (row) {
    const groupId = row[5] || "";
    return [
      row[0],
      row[1],
      row[2],
      groupId,
      getStoredGroupName(groupId),
      row[3],
      row[4],
      "text",
      "",
      "",
      "",
      ""
    ];
  });

  sheet.getRange(1, 1, 1, SHEET_HEADERS.length).setValues([SHEET_HEADERS]);
  if (migrated.length) {
    sheet.getRange(2, 1, migrated.length, SHEET_HEADERS.length).setValues(migrated);
  }
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
  if (isDeletedGroup(groupId)) {
    return;
  }

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
    .filter(function (group) {
      return group.groupId && !isDeletedGroup(group.groupId);
    })
    .sort(function (a, b) {
      return new Date(b.lastSeenAt).getTime() - new Date(a.lastSeenAt).getTime();
    });
}

function getStoredGroupName(groupId) {
  if (!groupId) {
    return "";
  }

  const sheet = getGroupSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return groupId;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, GROUP_HEADERS.length).getValues();
  for (let i = 0; i < values.length; i += 1) {
    if (values[i][0] === groupId) {
      return values[i][1] || groupId;
    }
  }

  return groupId;
}

function deleteGroup(groupId) {
  if (!groupId) {
    return;
  }

  addDeletedGroup(groupId);

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

function getDeletedGroupSheet() {
  const sheetId = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  const spreadsheet = SpreadsheetApp.openById(sheetId);
  const sheet = spreadsheet.getSheetByName(DELETED_GROUP_SHEET_NAME) || spreadsheet.insertSheet(DELETED_GROUP_SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(DELETED_GROUP_HEADERS);
  } else {
    ensureHeaders(sheet, DELETED_GROUP_HEADERS);
  }

  return sheet;
}

function addDeletedGroup(groupId) {
  const sheet = getDeletedGroupSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    const values = sheet.getRange(2, 1, lastRow - 1, DELETED_GROUP_HEADERS.length).getValues();
    for (let i = 0; i < values.length; i += 1) {
      if (values[i][0] === groupId) {
        sheet.getRange(i + 2, 2).setValue(new Date().toISOString());
        return;
      }
    }
  }

  sheet.appendRow([groupId, new Date().toISOString()]);
}

function isDeletedGroup(groupId) {
  if (!groupId) {
    return false;
  }

  const sheet = getDeletedGroupSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return false;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return values.some(function (row) {
    return row[0] === groupId;
  });
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

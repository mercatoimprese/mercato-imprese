const SHEET_NAME = "Aziende";
const DRIVE_FOLDER_NAME = "Mercato Imprese - Allegati";

function doGet() {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  const headers = values.shift() || [];

  const aziende = values
    .filter((row) => row.some((cell) => cell !== ""))
    .map((row) => {
      const item = {};
      headers.forEach((header, index) => {
        item[header] = row[index];
      });
      return item;
    });

  return json({ aziende });
}

function doPost(event) {
  const body = parseBody(event);
  const sheet = getSheet();
  const id = body.id || Utilities.getUuid();
  const allegato = salvaAllegato(id, body);

  sheet.appendRow([
    id,
    body.nome || "",
    body.prezzo || "",
    body.mimeType || "",
    body.fileName || "",
    allegato.url || body.file || "",
    new Date().toISOString(),
  ]);

  return json({ ok: true, id, file: allegato.url || body.file || "" });
}

function parseBody(event) {
  if (!event || !event.postData || !event.postData.contents) {
    return {};
  }

  return JSON.parse(event.postData.contents);
}

function getSheet() {
  const properties = PropertiesService.getScriptProperties();
  let spreadsheetId = properties.getProperty("SPREADSHEET_ID");
  let spreadsheet;

  if (spreadsheetId) {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } else {
    spreadsheet = SpreadsheetApp.create("Mercato Imprese");
    properties.setProperty("SPREADSHEET_ID", spreadsheet.getId());
  }

  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "id",
      "nome",
      "prezzo",
      "mimeType",
      "fileName",
      "file",
      "dataCreazione",
    ]);
  }

  return sheet;
}

function salvaAllegato(id, body) {
  if (!body.file || !body.mimeType || body.file.indexOf("base64,") === -1) {
    return {};
  }

  const base64 = body.file.split("base64,")[1];
  const bytes = Utilities.base64Decode(base64);
  const extension = extensionDaMimeType(body.mimeType);
  const fileName = body.fileName || `${id}${extension}`;
  const blob = Utilities.newBlob(bytes, body.mimeType, fileName);
  const folder = getFolder();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    url: `https://drive.google.com/uc?export=view&id=${file.getId()}`,
  };
}

function getFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function extensionDaMimeType(mimeType) {
  const map = {
    "image/jpeg": ".jpg",
    "image/png": ".png",
    "image/webp": ".webp",
    "image/gif": ".gif",
    "application/pdf": ".pdf",
  };
  return map[mimeType] || "";
}

function json(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

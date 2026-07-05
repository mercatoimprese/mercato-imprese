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
      const files = preparaFilesPerSito(item);
      item.files = files;
      if (files.length) {
        item.file = files[0].file;
        item.mimeType = files[0].mimeType;
        item.fileName = files[0].fileName;
      } else {
        item.file = "";
        item.mimeType = "";
        item.fileName = "";
      }
      return item;
    });

  return json({ apiVersion: 2, aziende });
}

function doPost(event) {
  const body = parseBody(event);
  const sheet = getSheet();

  if (body.action === "delete") {
    return json(eliminaAnnuncio(sheet, body));
  }

  const id = body.id || Utilities.getUuid();
  const allegati = salvaAllegati(id, body);
  const primoAllegato = allegati[0] || salvaAllegato(id, body);

  sheet.appendRow([
    id,
    body.nome || "",
    body.prezzo || "",
    primoAllegato.mimeType || body.mimeType || "",
    primoAllegato.fileName || body.fileName || "",
    primoAllegato.file || body.file || "",
    new Date().toISOString(),
    JSON.stringify(allegati),
  ]);

  return json({
    ok: true,
    id,
    file: primoAllegato.file || body.file || "",
    files: allegati,
  });
}

function eliminaAnnuncio(sheet, body) {
  const values = sheet.getDataRange().getValues();
  const headers = values[0] || [];
  const idIndex = headers.indexOf("id");
  const nomeIndex = headers.indexOf("nome");
  const prezzoIndex = headers.indexOf("prezzo");
  const fileIndex = headers.indexOf("file");
  let eliminati = 0;

  for (let rowIndex = values.length - 1; rowIndex >= 1; rowIndex--) {
    const row = values[rowIndex];
    const id = idIndex >= 0 ? String(row[idIndex] || "") : "";
    const key = [
      nomeIndex >= 0 ? String(row[nomeIndex] || "").trim().toLowerCase() : "",
      prezzoIndex >= 0 ? String(row[prezzoIndex] || "").trim() : "",
      fileIndex >= 0 ? String(row[fileIndex] || "").trim() : "",
    ].join("|");

    if ((body.id && id === String(body.id)) || (body.key && key === body.key)) {
      sheet.deleteRow(rowIndex + 1);
      eliminati++;
    }
  }

  return {
    ok: eliminati > 0,
    deleted: eliminati,
  };
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
      "files",
    ]);
  } else {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.indexOf("files") === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue("files");
    }
  }

  return sheet;
}

function salvaAllegati(id, body) {
  const inputFiles = Array.isArray(body.files) ? body.files : [];
  return inputFiles
    .map((item, index) => salvaAllegato(`${id}-${index + 1}`, item))
    .filter((item) => item.file);
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
    file: `https://drive.google.com/uc?export=view&id=${file.getId()}`,
    mimeType: body.mimeType,
    fileName,
  };
}

function preparaFilesPerSito(item) {
  if (item.files) {
    try {
      const files = JSON.parse(item.files);
      if (Array.isArray(files)) {
        return files.map(preparaFilePerSito).filter((file) => file.file);
      }
    } catch (err) {
      console.warn("Lista file non leggibile", err);
    }
  }

  const file = preparaFilePerSito(item);
  return file ? [file] : [];
}

function preparaFilePerSito(item) {
  if (!item.file || !item.mimeType || !String(item.mimeType).startsWith("image/")) {
    return item.file
      ? {
          file: item.file,
          mimeType: item.mimeType || "",
          fileName: item.fileName || "",
        }
      : null;
  }

  const fileId = estraiDriveId(item.file);
  if (!fileId) {
    return {
      file: item.file,
      mimeType: item.mimeType,
      fileName: item.fileName || "",
    };
  }

  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const bytes = blob.getBytes();
    if (!bytes.length) {
      return null;
    }

    return {
      file: `data:${item.mimeType};base64,${Utilities.base64Encode(bytes)}`,
      mimeType: item.mimeType,
      fileName: item.fileName || file.getName(),
    };
  } catch (err) {
    console.warn("File Drive non leggibile", err);
    return null;
  }
}

function estraiDriveId(url) {
  const patterns = [
    /drive\.google\.com\/file\/d\/([^/]+)/,
    /drive\.google\.com\/uc\?[^#]*id=([^&#]+)/,
    /drive\.google\.com\/open\?[^#]*id=([^&#]+)/,
    /drive\.google\.com\/thumbnail\?[^#]*id=([^&#]+)/,
  ];

  for (const pattern of patterns) {
    const match = String(url).match(pattern);
    if (match) {
      return decodeURIComponent(match[1]);
    }
  }

  return null;
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

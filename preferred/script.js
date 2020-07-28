// Workflow
//* Copy Entrata export over "Import Move-Ins" sheet
//* Click Process
//* Export will be located against HC asset database
//* Export will be cleaned and replaced/

const baseURL = "https://manage.happyco.com";
const assetServiceURL = baseURL + "/twirp/backend.v1.Asset.Service";
const backendServiceURL = baseURL + "/twirp/backend.v1.Inspection.Service";
const businessServiceURL = baseURL + "/twirp/backend.v1.Business.Service";
const folderServiceURL = baseURL + "/twirp/backend.v1.Folder.Service";
const inspectionServiceURL =
  baseURL + "/twirp/inspection.v1.Inspection.Service";
const templateServiceURL = baseURL + "/twirp/backend.v1.Template.Service";

const fetchAPI = (url, token, data) => {
  const options = {
    method: "POST",
    headers: { authorization: "Bearer " + token },
    contentType: "application/json",
    payload: JSON.stringify(data),
    muteHttpExceptions: true,
  };
  const resp = UrlFetchApp.fetch(url, options);
  return {
    statusCode: resp.getResponseCode(),
    data: JSON.parse(resp.getContentText()),
  };
};

// Clean pasted (imported) data of duplicate/previously emailed rows incase there are any import errors.
const cleanImportSheet = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const incomingDataSheet = ss.getSheetByName("Import Move-ins");
  const incomingData = incomingDataSheet.getDataRange().getValues();

  const logSheet = ss.getSheetByName("Import Log");
  const logData = logSheet.getDataRange().getValues();

  const previouslyInspected = logData.map((row) => row[0]);
  const incomingEmails = incomingData.map((row) => row[2]);
  const cleaned = incomingData.filter((row) => {
    let previouslyImported = previouslyInspected.includes(row[0]);
    let duplicate =
      incomingEmails.filter((email) => email === row[4]).length > 1;
    return !duplicate && !previouslyImported;
  });
  incomingDataSheet.clearContents();
  incomingDataSheet
    .getRange(1, 1, cleaned.length, cleaned[0].length)
    .setValues(cleaned);
  return true;
};

const process = () => {
  cleanImportSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const incomingDataSheet = ss.getSheetByName("Import Move-ins");

  const assetDataSheet = ss.getSheetByName("HappyCo Assets");
  const assetData = assetDataSheet.getDataRange().getValues();

  let range = incomingDataSheet.setActiveRange(
    incomingDataSheet.getRange("A2")
  );

  const deduceBuildingColumn = 8;
  const deduceUnitColumn = 9;
  const deduceBedColumn = 10;
  const locationIdColumn = 11;
  const errorColumn = 12;

  while (
    (range = incomingDataSheet.setActiveRange(
      incomingDataSheet.getRange(range.getRow() + 1, 1, 1, errorColumn)
    ))
  ) {
    var values = range.getValues()[0];
    var entrataUnit = values[0].toString();

    if (!entrataUnit) break;

    var buildingNumber = "";
    var unitNumber = "";
    var bedSpace = "";
    var occurrences = (entrataUnit.match(/-/g) || []).length;

    if (occurrences === 1) {
      var unitArr = entrataUnit.split("-");
      unitNumber = unitArr[0];
      bedSpace = unitArr[1];
    } else if (occurrences === 2) {
      var unitArr = entrataUnit.split("-");
      buildingNumber = unitArr[0];
      unitNumber = unitArr[1];
      bedSpace = unitArr[2];
    } else {
      unitNumber = entrataUnit;
    }

    // Doesn't have to be writted to the sheet.
    // This is mostly for some visual cues
    range.getCell(1, deduceBuildingColumn).setValue(buildingNumber);
    range.getCell(1, deduceUnitColumn).setValue(unitNumber);
    range.getCell(1, deduceBedColumn).setValue(bedSpace);

    const locationId = assetData.filter((asset) => asset[1] === unitNumber)[0];

    if (!!locationId) {
      range.getCell(1, locationIdColumn).setValue(locationId);
    } else {
      range
        .getCell(1, errorColumn)
        .setValue(
          "It seem like we were unable to find that unit in the HappyCo export."
        );
    }
  }
  copyData();
};

const copyData = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const primarySheet = ss.getSheetByName("Primary");
  const incomingDataSheet = ss.getSheetByName("Import Move-ins");
  const configurationDataSheet = ss.getSheetByName("Configuration");

  const templateId = configurationDataSheet.getRange("B3").getValue();

  let data = [];

  let range = incomingDataSheet.setActiveRange(
    incomingDataSheet.getRange("A2")
  );
  while (
    (range = incomingDataSheet.setActiveRange(
      incomingDataSheet.getRange(range.getRow() + 1, 1, 1, 12)
    ))
  ) {
    const row = range.getValues()[0];

    if (!row[0].toString()) break;
    if (row[10] && row[10] !== "FALSE") {
      data.push([
        "",
        templateId,
        row[10],
        row[0],
        row[2],
        `${row[4]} ${row[5]}`,
        row[1],
      ]);
    }
  }

  if (data.length > 0)
    primarySheet.getRange(4, 1, data.length, data[0].length).setValues(data);
};

const deliver = () => {
  const statusColumn = 7;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Primary");
  const importLogDataSheet = ss.getSheetByName("Import Log");
  const customMessage = sheet.getRange("B1").getValue();

  if (
    sheet.setActiveRange(sheet.getRange("A3")).getValue() != "Inspection ID"
  ) {
    throw new Error("Data cells appear to have moved");
  }

  const configurationDataSheet = ss.getSheetByName("Configuration");
  const token = configurationDataSheet.getRange("B5").getValue();

  let range = sheet.setActiveRange(sheet.getRange("A3"));

  while (
    (range = sheet.setActiveRange(
      sheet.getRange(range.getRow() + 1, 1, 1, statusColumn)
    ))
  ) {
    let row = range.getValues()[0];
    let inspectionId = row[0];
    let templateId = row[1];
    let locationId = row[2];
    let locationName = row[3];
    let email = row[4];
    let name = row[5];
    let status = row[6];

    // Fail if we don't have InspectionID AND template or location
    if (!templateId || !locationId) {
      return;
    }

    // Haven't previously created an Inspection, so lets do that.
    if (!inspectionId) {
      var resp = fetchAPI(`${inspectionServiceURL}/Create`, token, {
        locationId,
        templateId,
        scheduleDate: new Date().toISOString(),
      });
      if (resp.statusCode == 200) {
        // Inspection was created, add to #Send payload
        inspectionId = resp.data.id;
        range.getCell(1, 1).setValue(resp.data.id);
      } else {
        // Skip if inspection create fails
        range.getCell(1, statusColumn).setValue("✘ " + resp.data.msg);
        continue;
      }
    }

    // Skip this row if there's no email
    if (!email) continue;

    // Check if Inspection has already been sent
    if (!!status) {
      if (status.startsWith("✔")) {
        console.log(`Skipping sent inspection ${inspectionId}`);
        continue;
      } else {
        range.getCell(1, statusColumn).setValue("");
      }
    }

    var resp = fetchAPI(`${inspectionServiceURL}/Send`, token, {
      inspectionId,
      email,
      name,
      customMessage,
    });

    if (resp.statusCode == 200) {
      const log = [
        locationName,
        name,
        email,
        locationId,
        templateId,
        new Date().toISOString(),
      ];
      importLogDataSheet.appendRow(log);
      range.getCell(1, statusColumn).setValue("✔ Sent " + new Date());
    } else {
      range.getCell(1, statusColumn).setValue("✘ " + resp.data.msg);
    }
  }
};

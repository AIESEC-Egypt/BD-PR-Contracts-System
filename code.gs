function createContractBD() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "BD Contracts System"
  );
  const sheetData = sheet
    .getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
    .getValues();
  const referenceSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reference");
  const referenceSheetData = referenceSheet
    .getRange(1, 1, referenceSheet.getLastRow(), referenceSheet.getLastColumn())
    .getValues();
  const rowIndex = sheet.getLastRow();
  const contractType = sheet.getRange(rowIndex, 2).getValue();
  const partnerNameCol = sheet
    .createTextFinder("Name of Partner or Organization: ")
    .matchEntireCell(true)
    .findAll()
    .map((x) => x.getColumn());
  if (partnerNameCol.length > 1) {
    if (contractType == "Memorandum of Agreement ( PR )") {
      var partnerName = sheet.getRange(rowIndex, partnerNameCol[1]).getValue();
      var folder = DriveApp.getFolderById("1HhVevhIhsuBkKJw9a_TEY2G8cjDvXvL9");
      var cc_emails = mcvpB2C;
    } else if (contractType == "Memorandum of Understanding") {
      var partnerName = sheet.getRange(rowIndex, partnerNameCol[0]).getValue();
      var folder = DriveApp.getFolderById("1hfD9oc1D1HYEWBDByOptTR3SL_jVTwnc");
      var cc_emails = mcvpBD;
    } else if (contractType == "Memorandum of Agreement %5 ( PR )") {
      var partnerName = sheet.getRange(rowIndex, partnerNameCol[1]).getValue();
      var folder = DriveApp.getFolderById("1xwWSGISuJfruzLTkVllrgCB4yaBoIYL9");
      var cc_emails = mcvpB2C;
    }
  }
  const contractIDIndex = referenceSheet
    .createTextFinder(contractType)
    .matchEntireCell(true)
    .findAll()
    .map((x) => x.getRow())[0];
  Logger.log(contractIDIndex);
  const contractID = referenceSheet.getRange(contractIDIndex, 2).getValue();
  Logger.log(contractID);

  const template = DriveApp.getFileById(contractID);
  const name = `AIESEC in Egypt - ${contractType} - ${partnerName}`;
  const newFile = template.makeCopy(name, folder);
  console.log(newFile.getUrl());
  const doc = DocumentApp.openById(newFile.getId());
  const docBody = doc.getBody();
  var sendCol = sheet
    .createTextFinder("Email Sent?")
    .matchEntireCell(true)
    .findAll()
    .map((x) => x.getColumn());
  if (sheet.getRange(rowIndex, sendCol).getValue() == true) return;
  Logger.log("I am here");
  sheet
    .getRange(
      rowIndex,
      sheet
        .createTextFinder("Link of The contract")
        .matchEntireCell(true)
        .findAll()
        .map((x) => x.getColumn())
    )
    .setValue(newFile.getUrl());
  const pairs = referenceSheet
    .createTextFinder(`${contractType}`)
    .matchEntireCell(true)
    .findAll()
    .map((x) => [x.getRow(), x.getColumn()]);
  Logger.log(pairs);
  for (let i = 1; i < pairs.length; i++) {
    if (
      referenceSheetData[pairs[i][0] - 1][pairs[i][1]] !=
      "Email of AIESEC Representative:"
    ) {
      Logger.log("Loop");
      Logger.log(referenceSheetData[pairs[i][0] - 1][pairs[i][1]]);
      var colIndex = sheet
        .createTextFinder(referenceSheetData[pairs[i][0] - 1][pairs[i][1]])
        .matchEntireCell(true)
        .findAll()
        .map((x) => x.getColumn());
      Logger.log("colIndex");
      Logger.log(colIndex[0]);
      Logger.log(colIndex[1]);
      if (colIndex.length > 1) {
        if (contractType == "Memorandum of Agreement ( PR )") {
          colIndex = colIndex[1];
          var id = 00;
        } else if (contractType == "Memorandum of Understanding") {
          colIndex = colIndex[0];
          var id = 11;
        } else if (contractType == "Memorandum of Agreement %5 ( PR )") {
          colIndex = colIndex[1];
          var id = 22;
        }
      }
      var replaced = referenceSheetData[pairs[i][0] - 1][pairs[i][1] + 1];
      var value = sheetData[rowIndex - 1][colIndex - 1];
      if (
        referenceSheetData[pairs[i][0] - 1][pairs[i][1]] ==
        "AIESEC in Egypt – Local Committee Branch"
      ) {
        var lc = sheetData[rowIndex - 1][colIndex - 1];
      }
      if (
        referenceSheetData[pairs[i][0] - 1][pairs[i][1]] == "Reference Code"
      ) {
        var indices = referenceSheet
          .createTextFinder(`${lc}`)
          .matchEntireCell(true)
          .findAll()
          .map((x) => [x.getRow(), x.getColumn() + 1]);
        var lcCode = lcMap[`${lc}`];
        var date = Utilities.formatDate(new Date(), "GMT+3", dateFormat);
        var value = lcCode + id + date + Math.floor(Math.random() * 100000 + 1);
        sheet.getRange(rowIndex, colIndex).setValue(value);
        Logger.log("ID");
        Logger.log(value);
      }

      if (replaced.includes("Date")) {
        var value = Utilities.formatDate(value, "GMT+3", "dd/MM/yyyy");
      }
      docBody.replaceText(replaced, value);
    } else {
      var email =
        sheetData[rowIndex - 1][
          parseInt(referenceSheetData[pairs[i][0] - 1][pairs[i][1] + 1]) - 1
        ];
    }
  }
  doc.saveAndClose();
  // if(contractType == "Memorandum of Agreement ( PR )" || contractType == "Memorandum of Agreement %5 ( PR )"){
  //   sendPRContract(sheet,email,mcvpB2C,rowIndex,doc,sendCol)
  // }
}

// function sendPRContract(sheet,email,mcvpB2C,rowIndex,doc,sendCol){
//   var linkCol = sheet.createTextFinder("Link of The contract").matchEntireCell(true).findAll().map(x => x.getColumn())
//   var doc = DocumentApp.openByUrl(sheet.getRange(rowIndex,linkCol).getValue())
//   var docName = doc.getName()
//   var docID = DocumentApp.openByUrl(sheet.getRange(rowIndex,linkCol).getValue()).getId()
//   var file = DriveApp.getFileById(docID)
//   var docblob = file.getAs("application/pdf")
//   docblob.setName(doc.getName() + ".pdf");
//   var file = DriveApp.createFile(docblob);
//   var fileId = file.getId()
//   var lcName = sheet.getRange(rowIndex,16).getValue()
//   var email = sheet.getRange(rowIndex,18).getValue()
//   moveFileId(fileId,prFolder)
//   MailApp.sendEmail
//       ({
//         to:`${email}`,
//         subject:`${docName}`,
//         cc:`${mcvpB2C}`,
//         body:`Greeting from AIESEC in Egypt.\n\nYou can find a copy of the contract that should be signed in the next few days with the promising partner.\n\nThank you.\n\nPlease, dont' reply. This is an automated mail.`,
//         attachments:[doc.getAs(MimeType.PDF)]
//       })
//         sheet.getRange(rowIndex,sendCol).setValue(true).setBackground("green").setFontColor("white")
// }

function onEdit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "BD Contracts System"
  );
  var sendCol = sheet
    .createTextFinder("Email Sent?")
    .matchEntireCell(true)
    .findAll()
    .map((x) => x.getColumn());
  var linkCol = sheet
    .createTextFinder("Link of The contract")
    .matchEntireCell(true)
    .findAll()
    .map((x) => x.getColumn());
  var row = e.range.getRow();
  var column = e.range.getColumn();
  if (column == 24) {
    if (
      e.range.getValue() == "Proceed" &&
      sheet.getRange(row, 2).getValue() == "Memorandum of Understanding" &&
      sheet.getRange(row, 21).getValue() == ""
    ) {
      console.log("done");
      var doc = DocumentApp.openByUrl(sheet.getRange(row, linkCol).getValue());
      var docName = doc.getName();
      var docID = DocumentApp.openByUrl(
        sheet.getRange(row, linkCol).getValue()
      ).getId();
      var file = DriveApp.getFileById(docID);
      var docblob = file.getAs("application/pdf");
      docblob.setName(doc.getName() + ".pdf");
      var file = DriveApp.createFile(docblob);
      var fileId = file.getId();
      var lcName = sheet.getRange(row, 5).getValue();
      var email = sheet.getRange(row, 7).getValue();
      moveFileId(fileId, bdFolders[`${lcName}`]);
      MailApp.sendEmail({
        to: `${email}`,
        subject: `${docName}`,
        cc: "t.yahia@aiesec.org.eg",
        body: `Greeting from AIESEC in Egypt.\n\nYou can find a copy of the contract that should be signed in the next few days with the promising partner.\n\nThank you.\n\nPlease, dont' reply. This is an automated mail.`,
        attachments: [doc.getAs(MimeType.PDF)],
      });
      sheet
        .getRange(row, sendCol)
        .setValue(true)
        .setBackground("green")
        .setFontColor("white");
    } else if (
      e.range.getValue() == "Proceed" &&
      (sheet.getRange(row, 2).getValue() == "Memorandum of Agreement ( PR )" ||
        sheet.getRange(row, 2).getValue() ==
          "Memorandum of Agreement %5 ( PR )") &&
      sheet.getRange(row, 21).getValue() == ""
    ) {
      var linkCol = sheet
        .createTextFinder("Link of The contract")
        .matchEntireCell(true)
        .findAll()
        .map((x) => x.getColumn());
      var doc = DocumentApp.openByUrl(sheet.getRange(row, linkCol).getValue());
      var docName = doc.getName();
      var docID = DocumentApp.openByUrl(
        sheet.getRange(row, linkCol).getValue()
      ).getId();
      var file = DriveApp.getFileById(docID);
      var docblob = file.getAs("application/pdf");
      docblob.setName(doc.getName() + ".pdf");
      var file = DriveApp.createFile(docblob);
      var fileId = file.getId();
      var lcName = sheet.getRange(row, 16).getValue();
      var email = sheet.getRange(row, 18).getValue();
      moveFileId(fileId, prFolder);
      MailApp.sendEmail({
        to: `${email}`,
        subject: `${docName}`,
        cc: `${mcvpB2C}`,
        body: `Greeting from AIESEC in Egypt.\n\nYou can find a copy of the contract that should be signed in the next few days with the promising partner.\n\nThank you.\n\nPlease, dont' reply. This is an automated mail.`,
        attachments: [doc.getAs(MimeType.PDF)],
      });
      sheet
        .getRange(row, sendCol)
        .setValue(true)
        .setBackground("green")
        .setFontColor("white");
    }
  }
}

function moveFileId(fileId, toFolderId) {
  var file = DriveApp.getFileById(fileId);
  var source_folder = DriveApp.getFileById(fileId).getParents().next();
  var folder = DriveApp.getFolderById(toFolderId);
  folder.addFile(file);
  source_folder.removeFile(file);
}

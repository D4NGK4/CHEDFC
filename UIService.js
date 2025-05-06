

function getImageUrl() {
  var fileId = "1GcJ8_KrVSXIPxOBZb27e1Oa8f0WIyNSl";
  var file = DriveApp.getFileById(fileId);
  var blob = file.getBlob();
  var contentType = blob.getContentType();
  var base64Data = Utilities.base64Encode(blob.getBytes());
  return "data:" + contentType + ";base64," + base64Data;
}

function getFormList() {
  try {
    const mappedData = mapTemplateData();
    const result = [{
      number: 1,
      formName: mappedData["File Name"],
      templateId: mappedData["File ID"],
      url: `https://docs.google.com/document/d/${mappedData["File ID"]}/edit`,
      qr: `https://quickchart.io/qr?size=350x350&text=${encodeURIComponent(`https://docs.google.com/document/d/${mappedData["File ID"]}/edit`)}`,
      Fields: mappedData["Fields"]
    }];

    Logger.log("Processed Data: " + JSON.stringify(result));
    return result;
  } catch (e) {
    Logger.log("Error: " + e.toString());
    return [];
  }
}

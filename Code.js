function doGet() {
  const html = HtmlService.createTemplateFromFile('index');
  const evaluated = html.evaluate();
  evaluated.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  evaluated.setTitle('CHEDXFC');
  evaluated.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return evaluated;
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getUserProfile() {
  const response = UrlFetchApp.fetch('https://people.googleapis.com/v1/people/me?personFields=names,photos', {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    }
  });
  return JSON.parse(response.getContentText());
}

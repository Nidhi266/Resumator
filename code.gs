// Function to connect to the Google Sheet
function SheetConn(SheetName) {
  /*
  This function connects to the Google Sheet.
  It is needed in order to access the sheet and read and write data.
  
  @parameters: sheetname
  @return: connection string
  */
  
  // Google Sheet ID  
  var strFileID = "1Dx7FB7rX8HMtnN8Vs_jQuJDO5xw2pBc0HL6gy0uoC4Q";
  
  var ss = SpreadsheetApp.openById(strFileID);
  var sheet = ss.getSheetByName(SheetName);
  
  return sheet;  
}

// Function to get the current user's email
function GetUserEmail() {
  return Session.getActiveUser().getEmail();
}

// Function called when the Google Site page is loaded
function doGet(e) {
  /*
  This function gets called when the Google Site page is loaded.
  */

  // This will be the first page the user sees
  var mainscreen = "index";  
  
  var SiteName = "Resumator  - Google App Script"; 
  
  return HtmlService.createHtmlOutputFromFile(mainscreen)
    .setTitle(SiteName)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// Function to add a new entry to the Google Sheet
function AddNewCreate(FormData) {
  try {
    var Sheet = SheetConn("form");
    // Form object
    var d1 = FormData;
    
    var values =  [
      d1['user_name'],
      d1['user_phone'],
      d1['user_email'],
      d1['user_LinkedIN'],
      d1['user_location'],
      d1['user_Summary'],
      d1['user_skills'],
      d1['user_Education'],
      d1['user_Degree'],
      d1['Education_Location'],
      d1['User_Graduation_year'],
      d1['user_Company_Name'],
      d1['user_Role'],
      d1['Company_Start_year'],
      d1['Company_End_year'],
      d1['Company_Responsiblities'],
      d1['Certification_Name'],
      d1['CertificationLink'],
      d1['user_dttm'],
    ];
    var Add =  Sheet.appendRow(values);

    var data = {
      status:"success",
      msg: "Successfully added to the sheet",
    };  
    return data;
  } catch (error) {    
    // If there's an error, show the error message
    var data = {
      status:"Failed",
      msg: error.toString(),
    };  
    return data;  
  } 
}

// Function to create new Google Docs based on a template
function createNewGoogleDocs() {
  // This value should be the ID of your document template
  const googleDocTemplate = DriveApp.getFileById('1CEvalp36Pc8IEIgkdjHEeHyLAUqM_7VXqFYEqX0iZW8');
  
  // This value should be the ID of the folder where you want your completed documents stored
  const destinationFolder = DriveApp.getFolderById('1kSDd5sP2TE14fFyZWhek2ffZjTj6Ln0g');
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('form');
  const rows = sheet.getDataRange().getValues();
  
  // Start processing each spreadsheet row
  rows.forEach((row, index) => {
    if (index === 0 || row[19]) return;
    const copy = googleDocTemplate.makeCopy(`${row[1]}, ${row[0]} Resume 1` , destinationFolder)
    const doc = DocumentApp.openById(copy.getId())
    const body = doc.getBody();
    
    // Replace placeholders in the document with values from the spreadsheet row
    body.replaceText("{{Upload Image}}", "");
    body.replaceText('{{Name}}', row[0]);
    body.replaceText('{{Phone}}', row[1]);
    body.replaceText('{{Your email}}', row[2]);
    body.replaceText('{{LinkedIN}}', row[3]);
    body.replaceText('{{Your location}}', row[4]);
    body.replaceText('{{Summary}}', row[5]);
    body.replaceText('{{Skills}}', row[6]);
    body.replaceText('{{Education1}}', row[7]);
    body.replaceText('{{Degree1}}', row[8]);
    body.replaceText('{{Location1}}', row[9]);
    body.replaceText('{{Graduation_year1}}', row[10]);
    body.replaceText('{{Company_Name 1}}', row[11]);
    body.replaceText('{{Role 1}}', row[12]);
    body.replaceText('{{Start_year 1}}', row[13]);
    body.replaceText('{{End_year 1}}', row[14]);
    body.replaceText('{{Responsibilities 1}}', row[15]);
    body.replaceText('{{Certification_Name 1}}', row[16]);
    body.replaceText('{{Certification_Link 1}}', row[17]);
    
    doc.saveAndClose();
    const url = doc.getUrl();
    sheet.getRange(index + 1, 20).setValue(url);
  });
}

// Function to send emails based on status
function onSendemailsOpen(event) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('form');
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  values.forEach((row, i) => {
    const recipient = row[2];
    const status = row[21];

    if (status === "Completed") {
      const subject = 'Resume updated';
      const body = "Please click on the link to access the resume: " + row[19];
      GmailApp.sendEmail(recipient, subject, body);
    }
  });
}

// Function to interact with ChatGPT
function sendToChatGPT(input) {
  var apiKey = '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';
  var apiUrl = 'https://api.openai.com/v1/engines/text-davinci-003/completions';

  var requests = [];

  for (var i = 1; i <= 2; i++) {
    var maxTokens = 50; // Adjust this value based on your requirements

    var payload = {
      prompt: input,
      max_tokens: maxTokens
    };

    var options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + apiKey
      },
      payload: JSON.stringify(payload)
    };

    requests.push(options);
  }

  var responses = requests.map(function (options) {
    return UrlFetchApp.fetch(apiUrl, options);
  });

  var chatGptResponses = responses.map(function (response) {
    var responseData = JSON.parse(response.getContentText());
    return responseData.choices[0].text;
  });

  var combinedResponse = chatGptResponses.join(' ');

  return combinedResponse;
}

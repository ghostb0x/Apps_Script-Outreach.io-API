// Using the Apps Script OAuth2 Library to connect to Outreach.io's API

function getOutreachService() {
  return (
    OAuth2.createService('outreach')
      .setAuthorizationBaseUrl(
        'https://api.outreach.io/oauth/authorize'
      )
      .setTokenUrl('https://api.outreach.io/oauth/token')

      // Set the client ID and secret, from the Google Developers Console.
      .setClientId('CLIENT_ID')
      .setClientSecret('CLIENT_SECRET')

      // Set the name of the callback function in the script referenced
      // above that should be invoked to complete the OAuth flow.
      .setCallbackFunction('authCallback')

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties())

      // Set the scopes to request (space-separated for Google services).
      .setScope('mailings.read sequenceSteps.read mailboxes.read')
  );
}

/**
 * Logs the redirect URI to register.
 */
function logRedirectUri() {
  var service = getOutreachService();
  Logger.log(service.getRedirectUri());
}

function showAuthUrl() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var login_sheet = ss.getSheetByName('Login');
  var outreachService = getOutreachService();
  Logger.log(outreachService.hasAccess());
  if (!outreachService.hasAccess()) {
    var authorizationUrl = outreachService.getAuthorizationUrl();
    login_sheet.getRange(10, 2).setValue(authorizationUrl);
  }
}

function authCallback(request) {
  var outreachService = getOutreachService();
  var isAuthorized = outreachService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput(
      'Success! You can close this tab.'
    );
  } else {
    return HtmlService.createHtmlOutput(
      'Denied. You can close this tab'
    );
  }
}

function makeRequest() {
  var outreachService = getOutreachService();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var login_sheet = ss.getSheetByName('Login');
  var start = Number(login_sheet.getRange(16, 2).getValues());
  var end = Number(login_sheet.getRange(16, 3).getValues());

  //Loop over range of mailing ids
  for (var i = start; i <= end; i++) {
    //write each mailing id to Batch sheet row

    var batch_sheet = ss.getSheetByName('Batch');
    //Logger.log(batch_sheet);
    var lastRow = batch_sheet.getLastRow();
    //Logger.log(lastRow);

    batch_sheet.getRange(lastRow + 1, 1).setValue(i);

    try {
      // Using Outreach.io api to grab emails sent from Sales Users with specific properties
      // e.g. delivery date/time, email address, and step type
      var url =
        'https://api.outreach.io/api/v2/mailings/' +
        i +
        '?include=sequenceStep&fields[mailing]=deliveredAt,mailboxAddress&fields[sequenceStep]=stepType';

      var response = UrlFetchApp.fetch(url, {
        headers: {
          Authorization: 'Bearer ' + outreachService.getAccessToken(),
          'Content-Type': 'application/vnd.api+json',
        },
      });

      //parse JSON response
      var json = response.getContentText();

      var obj = JSON.parse(json);

      //define response contents
      var data = obj['data'];
      var included = obj['included'];

      if (included.length > 0) {
        var included = included[0];
        Logger.log(included);
        var delivery_date = data['attributes'][
          'deliveredAt'
        ].substring(0, 10);
        var mailbox_address = data['attributes']['mailboxAddress'];
        var step_type = included['attributes']['stepType'];

        // add data to Google sheet
        batch_sheet.getRange(lastRow + 1, 2).setValue(delivery_date);
        batch_sheet
          .getRange(lastRow + 1, 3)
          .setValue(mailbox_address);
        batch_sheet.getRange(lastRow + 1, 4).setValue(step_type);
      }

      batch_sheet.getRange(lastRow + 1, 5).setValue('done');
    } catch (err) {
      batch_sheet.getRange(lastRow + 1, 5).setValue('error');
    }
  }
}

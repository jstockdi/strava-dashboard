var scriptProperties = PropertiesService.getScriptProperties();

var spreadsheetId = scriptProperties.getProperty("SPREADSHEET_ID");
var sheetName = scriptProperties.getProperty("SHEET_NAME");
var clientId = scriptProperties.getProperty("CLIENT_ID");
var clientSecret = scriptProperties.getProperty("CLIENT_SECRET");

HEADER=['Name','Distance','Moving Time','Average Speed', 'Start Date'];

function logProps(){
  Logger.log("spreadsheetId: " + spreadsheetId);  
  Logger.log("sheetName: " + sheetName);  
  Logger.log("clientId: " + clientId);  
  Logger.log("clientSecret: " + clientSecret);
}

function run(){

  var stravaResults = getStravaResults();
  
  writeFullResults(stravaResults);
  
  Logger.log('Completed...');
};


function writeFullResults(stravaResults){
  
  var fullListArray = stravaResults.map(function(activity){
    return [
      activity.name,
      activity.miles,
      activity.movingTime,
      activity.averageSpeed,
      activity.startDateLocal
    ];
  });
  
  var sheet = getOrCreateSheet(sheetName);
  Logger.log('Found Sheet: ' + sheet.getName());
  sheet.clear();
    
  
  [HEADER].concat(fullListArray).forEach(function(row){
    sheet.appendRow(row);
  });
  
};


function getStravaResults() {
  
  var unixTime = 0;
  var responseCount = 30;  //default count to force first call
  var lastUnixTime = -1;
  
  var results = [];
    
  while(responseCount > 0){
      
    lastUnixTime = unixTime;
    var activities = getActivities(unixTime);
    responseCount = activities.length;
    
    if(responseCount > 0){
      var lastActivityDate = activities[responseCount - 1].startDate;
      unixTime = lastActivityDate.getTime() / 1000;
      
      results = results.concat(activities);
      
      Logger.log('unixTime: ' + unixTime + ' lastUnixTime: ' + lastUnixTime + ' responseCount: ' + responseCount);
      
    }
  }
  
  return results;
};


function getActivities(start_date){
  
  var service = getService();
  var stravaResults = [];
  
  if (!service.hasAccess()) {
    Logger.log('NO ACCESS');
    
  }else{
    var url = 'https://www.strava.com/api/v3/athlete/activities?per_page=125&after=' + start_date;
    Logger.log('url: ' + url);
    
    var response = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: 'Bearer ' + service.getAccessToken()
      }
    });
    
    var results = JSON.parse(response.getContentText());

    if (results.length == 0) {
      Logger.log('No new data');
    }else{      
      stravaResults = results.map(function(result){
        var name = result['name']; 
        var miles = result['distance'] * 0.000621371;
        var averageSpeed = result['average_speed'] * 0.000621371 * 60 * 60;
        var movingTime = result['moving_time'] / 60 / 60;
        var startDateLocal = new Date(Date.parse(result['start_date_local']));
        var startDate = new Date(Date.parse(result['start_date']));
        
        return {
          'name' : name,
          'miles' : miles,
          'movingTime' : movingTime,
          'averageSpeed' : averageSpeed,
          'startDateLocal' : startDateLocal,
          'startDate' : startDate
        };
      });
    }
  }
  
  return stravaResults;
  
};

function getOrCreateSheet(sheetName) {
  
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  Logger.log('Spreadsheet: ' + spreadsheet);
  
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    Logger.log('Sheet "%s" does not exists, adding new one.', sheetName);
    sheet = spreadsheet.insertSheet(sheetName)
  } 
  
  return sheet;
};




/**
 * Configures the service.
 */
function getService() {
  var service =  OAuth2.createService('Strava')
      // Set the endpoint URLs.
      .setAuthorizationBaseUrl('https://www.strava.com/oauth/authorize')
      .setTokenUrl('https://www.strava.com/oauth/token')

      // Set the client ID and secret.
      .setClientId(clientId)
      .setClientSecret(clientSecret)

      // Set the name of the callback function that should be invoked to complete
      // the OAuth flow.
      .setCallbackFunction('authCallback')

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties())
      //Include private activities when retrieving activities.
      .setScope ('view_private');
  
   if (service.hasAccess()) {
    var url = 'https://www.strava.com/api/v3/athlete';
    
    Logger.log('Using access token: ' + service.getAccessToken());
    var response = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: 'Bearer ' + service.getAccessToken()
      }
    });
    var result = JSON.parse(response.getContentText());
    
  } else {
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log('Load the authorizationUrl: %s',
        authorizationUrl);
  }

  return service;
}



/**
 * Handles the OAuth callback.
 */
function authCallback(request) {
  var service = getService();
  var authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('Success!!!');
  } else {
    return HtmlService.createHtmlOutput('Denied');
  }
}



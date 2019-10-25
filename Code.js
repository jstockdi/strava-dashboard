var scriptProperties = PropertiesService.getScriptProperties();

var spreadsheetId = scriptProperties.getProperty("SPREADSHEET_ID");
var sheetName = scriptProperties.getProperty("SHEET_NAME");
var clientId = scriptProperties.getProperty("CLIENT_ID");
var clientSecret = scriptProperties.getProperty("CLIENT_SECRET");

HEADER=['Name','Distance','Moving Time','Average Speed', 'Start Date'];


function doGet(e){
  
  var code = e.parameter.code;
  Logger.log("AccessToken: " + code);
  
  var access_token = getAPIAccessToken(code);
  Logger.log('access_token: ' + access_token);
  
  var stravaResults = getStravaResults(0, access_token);
  
  writeFullResults(stravaResults);
  
  Logger.log('Completed...');
  
  
  return HtmlService.createHtmlOutput('Success');
  
}

// curl -X POST 'https://www.strava.com/oauth/token'
//  -d 'client_id='
//  -d 'client_secret='
//  -d 'grant_type=authorization_code'
//   -d 'code=<>'
function getAPIAccessToken(code){

  var response = UrlFetchApp.fetch('https://www.strava.com/oauth/token', {
    'method' : 'post',
    'payload' : {
      'client_id': clientId,
      'client_secret': clientSecret,
      'grant_type' : 'authorization_code',
      'code' : code
    }
   });
  
  var contentText = response.getContentText();
  Logger.log('contentText: ' + contentText);
  
  var jsonResponse = JSON.parse(contentText);
  
  return jsonResponse.access_token;
  
}

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
  
  sheet.appendRow(HEADER)

  var range = sheet.getRange(2, 1, fullListArray.length, 5);
  range.setValues(fullListArray);
  
};


function getStravaResults(startTime, access_token) {
  
  var unixTime = startTime;
  var responseCount = 30;  //default count to force first call
  var lastUnixTime = -1;
  
  var results = [];
    
  while(responseCount > 0){
      
    lastUnixTime = unixTime;
    var activities = getActivities(unixTime, access_token);
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


function getActivities(start_date, access_token){
  

    var stravaResults = [];

    var url = 'https://www.strava.com/api/v3/athlete/activities?per_page=125&scope=activity:read_permission&after=' + start_date;
    Logger.log('url: ' + url);
    
    var response = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: 'Bearer ' + access_token
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
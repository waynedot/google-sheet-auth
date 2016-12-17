/*
https://github.com/waynedot/google-sheet-auth
*/

var fs = require('fs');
var readline = require('readline');
var google = require('googleapis');
var googleAuth = require('google-auth-library');

// If modifying these scopes, delete your previously saved credentials
// at ~/.credentials/sheets.googleapis.com-nodejs-quickstart.json
// Robert: 我不知道這下面三行在說啥
var SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/urlshortener'];
var TOKEN_DIR = (process.env.HOME || process.env.HOMEPATH || process.env.USERPROFILE) + '/.credentials/';
var TOKEN_PATH = TOKEN_DIR + 'sheets.googleapis.com-nodejs-quickstart.json';

// user values
var numberRow = 3;
var spreadsheetId = '161VIxFLW8wZrG-Ow0xqaejGQEbvbc558NWhlT1PqO6w';

// Load client secrets from a local file.
fs.readFile('client_secret.json', function processClientSecrets(err, content) {
  if (err) {
    console.log('Error loading client secret file: ' + err);
    return;
  }
  // Authorize a client with the loaded credentials, then call the
  // Google Sheets API.
  authorize(JSON.parse(content), getUrl);
});

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 *
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
  var clientSecret = credentials.installed.client_secret;
  var clientId = credentials.installed.client_id;
  var redirectUrl = credentials.installed.redirect_uris[0];
  var auth = new googleAuth();
  var oauth2Client = new auth.OAuth2(clientId, clientSecret, redirectUrl);

  // Check if we have previously stored a token.
  fs.readFile(TOKEN_PATH, function(err, token) {
    if (err) {
      getNewToken(oauth2Client, callback);
    } else {
      oauth2Client.credentials = JSON.parse(token);
      callback(oauth2Client);
    }
  });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 *
 * @param {google.auth.OAuth2} oauth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback to call with the authorized
 *     client.
 */
function getNewToken(oauth2Client, callback) {
  var authUrl = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES
  });
  console.log('Authorize this app by visiting this url: ', authUrl);
  var rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });
  rl.question('Enter the code from that page here: ', function(code) {
    rl.close();
    oauth2Client.getToken(code, function(err, token) {
      if (err) {
        console.log('Error while trying to retrieve access token', err);
        return;
      }
      oauth2Client.credentials = token;
      storeToken(token);
      callback(oauth2Client);
    });
  });
}

/**
 * Store token to disk be used in later program executions.
 *
 * @param {Object} token The token to store to disk.
 */
function storeToken(token) {
  try {
    fs.mkdirSync(TOKEN_DIR);
  } catch (err) {
    if (err.code != 'EEXIST') {
      throw err;
    }
  }
  fs.writeFile(TOKEN_PATH, JSON.stringify(token));
  console.log('Token stored to ' + TOKEN_PATH);
}

/**
 * using the auth to get the value from google sheet
 *
 * @param {google.auth.OAuth2} auth The auth to get the values from the google sheet
 */
function getUrl(auth) {
  var longURL = new Array();
  var sheets = google.sheets('v4');
  sheets.spreadsheets.values.get({
    auth: auth,
    spreadsheetId: spreadsheetId,
    range: 'A1:A'+numberRow,
  }, function(err, response) {
    if (err) {
      console.log('The API returned an error: ' + err);
      return;
    }
    var rows = response.values;
    if (rows.length == 0) {
      console.log('No data found.');
    } else {
      console.log('Values:');
      for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        // print the value of the row first
        console.log('%s', row[0]);
        longURL.push(row[0]);
      }

      // shorten the Url and update to the google sheet
      var promiseArray = longURL.map(shortenUrl);
      Promise.all(promiseArray)
      .then(shortUrl => {
        console.log("All the urls were shortened");
        // update to the google sheet
        updateUrl(shortUrl, auth)
      })
      .catch(reason => {
        console.log(reason);
      })
    }
  });

  /**
   * using Promise to guarantees all the async task done, and update to google sheet
   *
   * @ {String Array} url
   */
  function shortenUrl(url) {
    return new Promise(
      function (resolve, reject) {
        var urlshortener = google.urlshortener('v1');

        var resource = {
          longUrl: url
        };

        // shorten url
        urlshortener.url.insert({
          resource,
          auth: auth
        }, function (err, response) {
          if (!err) {
            console.log('Short url is', response.id);
            resolve(response.id);
          } else {
            console.log('Shotening url returned an error:', err);
            reject(err);
          }
        });
    });
  }
}

/**
 * using Promise to guarantees all the async task done, and update to google sheet
 *
 * @param {String Array} shortURL
 * @param {google.auth.OAuth2} auth
 */
function updateUrl(shortURL, auth) {
  var sheets = google.sheets('v4');
  sheets.spreadsheets.values.update({
    auth: auth,
    spreadsheetId: spreadsheetId,
    range: 'B1:B'+numberRow,
    resource:{
      range: 'B1:B'+numberRow,
      majorDimension: 'COLUMNS',
      values: [
        shortURL
      ],
    },
    valueInputOption: 'RAW',
  }, function(err, response) {
    if (err) {
      console.log('Updating te sheet returned an error: ' + err);
      return;
    } else {
    	console.log('Updateing to sheet success');
    }
  });
}

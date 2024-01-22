/**
 * Temple Rodgers - 30/08/2023
 * on startup create the search input box on the right of the sheet
 * SearchFileForm, addMenu and onOpen came from 
 * https://codewithcurt.com/how-to-search-google-drive-on-google-sheets-using-apps-script/
 */
function SearchForFileForm()
{
  var form = HtmlService.createHtmlOutputFromFile('searchInput').setTitle('Search Files');
  SpreadsheetApp.getUi().showSidebar(form);
}

function addMenu()
{
  var menu = SpreadsheetApp.getUi().createMenu('File Search');
  menu.addItem('Search Shared Drives for Files', 'SearchForFileForm');
  menu.addToUi(); 
}

function onOpen(e)
{
  addMenu(); 
}

/**
 * list as many shared drives as you want (maxResults)
 * useDomainAdminAccess = false so you get only shared drives you can access
 * Implement shared drive support is here: https://developers.google.com/drive/api/guides/enable-shareddrives
 */
/**
 * driveLister based on the following:  
 * https://stackoverflow.com/questions/70289187/how-to-get-a-pagetoken-for-the-files-list-api-endpoint-of-the-google-drives-api#:~:text=function%20myFunction()%20%7B
 * 
 */
  //
  // set global variables because 
  // they're used in different functions
  var lastRow = 2;
  var foundRecords = 'false';
  //
  //
 function driveLister(searchString) {
  // as we're looping, we need to know which row we're on at the end of each loop
  // clear the spreadsheet outside the loop and
  // pass the last row to the loop array
  // create a new tab named as per the searchString
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet(searchString);
  var lastRow = sheet.getLastRow()+1;
  // add the headings row to the active sheet and set the last row counter
  //const heads = [['Name', 'ID', 'Link', 'Created Date', 'Modified Data', 'Mime Type', 'Size', 'Shared Drive ID', 'Shared Drive Name', 'Description Text']];
  const heads = [['Name', 'ID', 'Link', 'Created Date', 'Modified Data', 'Mime Type', 'Size', 'Shared Drive Name', 'Description Text']];
  sheet.getRange(1, 1, 1, 9).setValues([...heads]);

  // run the row counter in the subroutine
  lastRow = 2;

    //
    // now run the query and get the first page of results
    //
    let driveList = [];
    let pageToken = null;
    do {
      const obj = Drive.Drives.list({
        "useDomainAccess": false,
        "orderBy": "title",
        "maxResults": 100,
        "supportsAllDrives": true,
        "fields": "nextPageToken,items",
        "pageToken": pageToken
      });
      if (obj.items.length > 0) driveList = [...driveList, ...obj.items];
      pageToken = obj.nextPageToken;
    } while(pageToken);
  driveList.forEach(drive => generateDriveFiles(drive,sheet,searchString));
  //
    if(sheet.getLastRow() > 1)
    {
      return "<span style=\"font-weight: bold\" >Found Records</span>";
    }
    else
    {
      return "<span style=\"font-weight: bold\" >No Records Found</span>";
    }
  }
/**
 * 
 */
function generateDriveFiles(drive,sheet,searchString) {
  // get all files on this drive
  let filesList = [];
    //    exclude folders
    const filesQuery = "trashed = false AND fullText contains \'" + searchString + "\' AND mimeType != 'application/vnd.google-apps.folder'";
    filesList = driveCall_(filesQuery,drive.id);

  // constructing the 2d array for google sheets
  const res = filesList.map(f => {
    return [f.name, f.id, f.webViewLink, new Date(f.createdTime), new Date(f.modifiedTime), f.mimeType, f.quotaBytesUsed,f.driveName, f.description]
  });

  // writing the results to the report
  if (res.length != 0) { 
      foundRecords = 'true'; // to get success message, but it's not working yet
      sheet.getRange(lastRow, 1, res.length, res[0].length).setValues([...res]);
      lastRow = sheet.getLastRow()+1;
      return;
    }
}
/**
 * Make Drive API v3 files.lists calls
 * @param {String} optional query term
 * @return {Object} files resource object array
 * 
 * Method: files.list described in the link below
 * https://developers.google.com/drive/api/reference/rest/v3/files/list?apix_params=%7B%22q%22%3A%22name%20contains%20%27exam%27%22%2C%22fields%22%3A%22files(name)%22%7D
 * Try Google Workspace APIs here
 * https://developers.google.com/workspace/explore?filter=
 */
function driveCall_(filesQuery,drive) {
  // options
  const options = {
    muteHttpExceptions: true,
    method: "GET",
    headers: { authorization: "Bearer " + ScriptApp.getOAuthToken() }
  };

  // variables
  let pageToken = null;
  let filesList = [];

  // loop for drive api calls
  do {
    const params = {
/**
 * I DON'T FULLY UNDERSTAND HOW options and params are working together from here on
 *    "pageSize": 1000,
 *    REST:Resource: files can be found here: https://developers.google.com/drive/api/reference/rest/v2/files
 */
      "fields": "files(id,name,createdTime,modifiedTime,size,parents,webViewLink,mimeType,quotaBytesUsed,driveId,description),nextPageToken",
      'corpora': "drive",
      'supportsAllDrives': true,
      'includeItemsFromAllDrives': true,
      'driveId': drive
    }

    // additional parameters
    if (pageToken) params.pageToken = pageToken;
    if (filesQuery) params.q = filesQuery;

    // construct the call querystring
    const queryString = Object.keys(params).map(function (p) {
      return [encodeURIComponent(p), encodeURIComponent(params[p])].join("=");
    }).join("&");
    const url = "https://www.googleapis.com/drive/v3/files?" + queryString;
    const response = JSON.parse(UrlFetchApp.fetch(url, options).getContentText());
  if ("error" in response) {
    // Handle the error case
    const errorMessage = response.error.message;
    Logger.log("Error: " + errorMessage + "-> " + drive);
  } else {
      if (response.files.length > 0) {
        filesList = [...filesList, ...response.files];
        
        // Extract unique drive IDs from the response
        const uniqueDriveIds = new Set(response.files.map(file => file.driveId));
        
        // Fetch shared drive information for each unique drive ID
        uniqueDriveIds.forEach(driveId => {
          const driveInfoUrl = `https://www.googleapis.com/drive/v3/drives/${driveId}`;
          const driveInfoResponse = JSON.parse(UrlFetchApp.fetch(driveInfoUrl, options).getContentText());

          // Check for errors in drive info response
          if ("error" in driveInfoResponse) {
            Logger.log("Error retrieving drive info: " + driveInfoResponse.error.message);
          } else {
            // Update filesList with shared drive name
            const driveName = driveInfoResponse.name;
            filesList.forEach(file => {
              if (file.driveId === driveId) {
                file.driveName = driveName;
              }
            });
          }
        });
      }
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
  return filesList;
}

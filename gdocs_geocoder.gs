/*****************************************************************************\
* Batch Google Geocoding Script and Data Exporter *
* Author: Dale Kunce *
* http://normalhabit.com/ *
* Thanks to Alan Christopher Thomas for the google geocoding script
* Thanks to the MapBox Team for the export to GeoJSON script
\*****************************************************************************/

function onOpen() {
    // Add the Geocode menu
    SpreadsheetApp.getActiveSpreadsheet().addMenu("Geocoder", [{
        name: "Geocode addresses",
        functionName: 'geocode'
    },{
      name: 'Export GeoJSON',
      functionName: 'gjDialog'
  }]);
}

function geocode() {
    // Get the current spreadsheet, sheet, range and selected addresses
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = SpreadsheetApp.getActiveSheet();
    var range = SpreadsheetApp.getActiveRange();
    var addresses = range.getValues();

    // Determine the first row and column to geocode
    var row = range.getRow();
    var column = range.getColumn();

    // Set default destination columns
    var destination = new Array();
    destination[0] = column + 1;
    destination[1] = column + 2;

    // Prompt for latitude and longitude columns
    var response = Browser.inputBox("Coordinate Columns",
        "Please specify which columns should contain the latitude " +
        "and longitude values [ie. 'C,D', 'A,F', etc]. Leave blank to " +
        "insert new columns.",
        Browser.Buttons.OK_CANCEL);
    if (response == 'cancel') return;
    if (response == '')
        sheet.insertColumnsAfter(column, 2);
    else {
        var coord_columns = response.split(',');
        destination[0] = sheet.getRange(coord_columns[0] + '1').getColumn();
        destination[1] = sheet.getRange(coord_columns[1] + '1').getColumn();
    }

    // Initialize the geocoder and set loading status
    var geocoder = Maps.newGeocoder();
    var count = range.getHeight();
    spreadsheet.toast(count + " addresses are currently being geocoded. " +
                      "Please wait.", "Loading...", -1);
  
    // Iterate through addresses and geocode
    for (i in addresses) {
        var location = geocoder.geocode(
            addresses[i]).results[0].geometry.location;
        sheet.getRange(row, destination[0]).setValue(location.lat);
        sheet.getRange(row++, destination[1]).setValue(location.lng);
        Utilities.sleep(200);
    }

    // Remove loading status
    spreadsheet.toast("Geocoding is now complete.", "Finished", -1);
}

// UI to set up GeoJSON export
function gjDialog() {
  var headersRaw = getHeaders(sheet, activeRange, 1);

  // Create a new UI
  var app = UiApp.createApplication()
    .setTitle('Export GeoJSON')
    .setStyleAttribute('width', '460')
    .setStyleAttribute('padding', '20');

  // Create a grid to hold the form
  var grid = app.createGrid(4, 2);

  // Add form elements to the grid
  grid.setWidget(0, 0, app.createLabel('Unique ID:'));
  grid.setWidget(0, 1, app.createListBox().setName('idBox').setId('idBox'));
  grid.setWidget(1, 0, app.createLabel('Longitude:'));
  grid.setWidget(1, 1, app.createListBox().setName('lonBox').setId('lonBox'));
  grid.setWidget(2, 0, app.createLabel('Latitude:'));
  grid.setWidget(2, 1, app.createListBox().setName('latBox').setId('latBox'));

  // Set the list boxes to the header values
  for (var i = 0; i < headersRaw.length; i++) {
    app.getElementById('idBox').addItem(headersRaw[i]);
    app.getElementById('lonBox').addItem(headersRaw[i]);
    app.getElementById('latBox').addItem(headersRaw[i]);
  }

  // Create a vertical panel...
  var panel = app.createVerticalPanel().setId('settingsPanel');

  panel.add(app.createLabel(
    'To format your spreadsheet as GeoJSON file, select the following columns:'
  ).setStyleAttribute('margin-bottom', '20'));

  // ...and add the grid to the panel
  panel.add(grid);

  // Create a button and click handler; pass in the grid object as a
  // callback element and the handler as a click handler
  // Identify the function b as the server click handler
  var button = app.createButton('Export')
      .setStyleAttribute('margin-top', '10')
      .setId('export');
  var handler = app.createServerClickHandler('exportGJ');
  handler.addCallbackElement(grid);
  button.addClickHandler(handler);

  // Add the button to the panel and the panel to the application,
  // then display the application app in the spreadsheet doc
  grid.setWidget(3, 1, button);
  app.add(panel);
  ss.show(app);
}

// Handle submits by updating the settings object, calling the
// export function, updates the UI
function exportGJ(e) {
  settings = {
    id: e.parameter.idBox,
    lon: e.parameter.lonBox,
    lat: e.parameter.latBox
  };
  
  // Update ui to show status
  updateUi();
  
  // Create GeoJSON file and pass back it's filepath
  var file = createGJFile();
  
  // Update ui to deliver file
  displayFile(file);
}

function updateUi() {
  // Create a new UI instance
  var app = UiApp.createApplication()
    .setTitle('Export GeoJSON')
    .setStyleAttribute('width', '460')
    .setStyleAttribute('padding', '20');

  // Add a status message to the UI
  app.add(app.createLabel(
    'Exporting your file...')
    .setStyleAttribute('margin-bottom', '10')
    .setId('exportingLabel'));

  // Show the new UI
  ss.show(app);
}

function displayFile(file) {
  // Create a new UI instance
  var app = UiApp.createApplication()
    .setTitle('Export GeoJSON')
    .setStyleAttribute('width', '460')
    .setStyleAttribute('padding', '20');
  
  // Notify the user that the file is done and in the Google Docs list
  app.add(
    app.createLabel('The GeoJSON file has been saved in your Google Docs List.')
    .setStyleAttribute('margin-bottom', '10')
  );

  // And provide a link to it
  app.add(
    app.createAnchor('Download GeoJSON File', file.getUrl())
    .setStyleAttribute('font-size', '150%')
  );
  
  // Show the new UI
  ss.show(app);
}

// Get headers within a sheet and range
function getHeaders(sheet, range, columnHeadersRowIndex) {
    var numColumns = range.getEndColumn() - range.getColumn() + 1;
    var headersRange = sheet.getRange(columnHeadersRowIndex,
        range.getColumn(), 1, numColumns);
    return headersRange.getValues()[0];
}

// Create the GeoJSON file and returns its filepath
function createGJFile() {
  return DocsList.createFile(
    (cleanCamel(ss.getName()) || 'unsaved') + '-' + Date.now() + '.geojson',
    Utilities.jsonStringify({
      type: 'FeatureCollection',
      features: getRowsData(sheet, activeRange, 1)
    })
  );
}
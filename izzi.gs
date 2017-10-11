//[X][Y]
//CORE VERSION OF IZZI
//BURRITO VERSION - Working towards release
//0. Escape Illegal characters
//1. Starting an empty list with two values in breaks the object if it's mid-line
//2. Change the wrapper
//3. Multiple sheets
//4. Automatic drawers for the XML function
//5. Work out how to get the side control in
//6. Turn on/off xml encoding
//7. Multiple files

//izzi code -------------------------------------------------------

//Entry point
function doGet() {

  //grab the current spreadsheet, find the sheets, prep the empty starting value
  var sheets = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheets();

  var sheetValue = "";

  //Loop through the sheets
  for (var i = 0; i < sheets.length; i++){
    var sheet = sheets[i];
    var sheetName = sheet.getSheetName();

    //If there's not an exclamation in the sheet name: TODO replace with user input
    if (sheetName.indexOf('!') < 0)
    {
      //Grab the sheet's values as a range
      var range = sheet.getSheetValues(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());

      //Pass into the convert Sheet function
      sheetValue += convertSheet(range, sheetName);

      //Wrap this resulting information with a Collection title
      sheetValue = wrapElement(sheetName +"Collection", sheetValue);
    }
  }

  //Add the data header for the browser to read
  var xmlHeader = "<?xml version='1.0' encoding='UTF-8'?>";

  //Return the value to the content service
  return ContentService.createTextOutput(xmlHeader + sheetValue)
  .setMimeType(ContentService.MimeType.XML)
  ;
}

//The function that converts a single sheet into xml
function convertSheet(range, sheetName){

  //The total value to be delivered at the end
  var totalValue = "";

  //Find where our 'tag' is - Where the top row & the item title column meets. (TODO: Make this optional user input)
  var topRow = 0;
  var titleCol = 0;
  while(range[topRow][titleCol] != "XML"){

    if (topRow < 60){
      topRow++;
    }else{
      topRow = 0;
      titleCol++;
    }
  }

  //Count down the rows from the tag until the next row is empty. This provides the final row with data.
  var finalRow = topRow;
  while (hasData(range[finalRow+1][titleCol])){
    finalRow++;
  }

  //Count across the cols from the tag until the next col row is empty. This provides the final col with data.
  var finalCol = titleCol;
  while (hasData(range[topRow][finalCol+1])){
    finalCol++;
  }

  //Iterate through every row under  the top row
  for (var row = topRow+1; row <= finalRow; row++){

    //This is the data we'll be returning at the end of this for loop.
    var dataRow = "";

    //This lock allows us to skip an empty row, and it should start off turned off, as we don't know what's in this row yet.
    var skipEmptyRows = false;

    //Now, for every column to the right of the title column
    for (var column = titleCol+1; column <= finalCol; column++){

      //constants for the wrapping and adding of cell data for ease of reading/debugging
      var cell = range[row][column];
      var baseParent = range[topRow][column];
      var parentRow = topRow-1;

      //If the cell has some data in it
      if (hasData(cell))
      {
        //Open up it's parents
        dataRow += openParents(range, topRow, column, parentRow, "");

        //Wrap the cell inside the base parent
        dataRow += wrapElement(baseParent, cell);

        //Now close it's parents naturally
        dataRow += closeParents(range, topRow, column, parentRow, "", false);

        //As we've found some data in this row, we'll want the lock to be switched off.
        skipEmptyRows = false;
      }
      else
      {
        //If we aren't skipping empty rows
        if (skipEmptyRows == false)
        {
          //If this is the final cell in the row, we'll just want to close it's parents forcefully
          if (isFinalCell(range, finalCol, column, row))
          {
            dataRow += closeParents(range, topRow, column, parentRow, "", true);
          }
          else
          {
            //If it's not the final cell, we'll need to close up any loose parents before we go onto the next cell.
            dataRow += closeLooseParents(range, topRow, column, row, parentRow, "");
          }

          //Now we'll want to start skipping empty
          skipEmptyRows = true;
        }
      }
    }

    //wrap the sum of row inside of what the row started with. Include attributes.
    totalValue += wrapRowAndAttributes(range, titleCol, topRow, row, dataRow);

  }

  totalValue = wrapElement(sheetName, totalValue);
  return totalValue;
}

//A recursive function to determine all the parent tags we will need to open for this particular row.
function openParents(range, topRow, column, row, cell){

  //Grab the parent we're currently examining
  var parent = range[row][column];

  //If the parent we're examining isn't empty...
  if (hasData(parent)){

    //And the cell to the left isn't the same (indicating a new parent)
    //or the cell is on the topRow (indicating it's a bottom-level parent)
    if ((range[row][column-1] != parent) || (row == topRow)){

      //Append the cell with a a new parent tag
      cell = openElement(parent) + cell;
    }
  }

  //If we're not at the top of the spreadsheet, move up a row and call this function again
  if (row > 0){
    row = row-1;
    return openParents(range,topRow,column, row, cell);
  }
  //If we are, it's time to stop - return the cell
  else
  {
    return cell;
  }
}

//A function to close all parents that might still be open
function closeLooseParents(range, topRow, column, row, parentRow, cell){

  //If this cell isn't empty, and the cell to the right is empty
 if (!(hasData(range[row][column])) && (hasData(range[row][column+1])))
 {
    //Time to close up the parents.
    return cell += closeParents(range,topRow,column, parentRow, false);
  }
  else
  {
    //Move on to the cell to right and call this function again.
    column = column + 1;
    return closeLooseParents(range, topRow, column, row, parentRow, cell);
  }

}

//This will close up any parents for the cell. It'll need the column of the current cell, a row to search up (starting with the topRow) and the cell contents.
//isUnnatural closes are for premature closing at the end of a row, while natural are for all other situations
function closeParents(range, topRow, column, row, cell, isUnnatural){

  //Grab the parent we're currently examining
  var parent = range[row][column];

  //First, if the parent we're examining isn't empty...
  if (hasData(parent)){

    //if the force is on...
    if (isUnnatural)
    {
      //... And the cell to the left isn't the same, close that parent tag up
      if (range[row][column-1] == parent){
        cell = cell + closeElement(parent);
      }
    }
    else
    {
      //... And the cell to the right isn't the same, or is on the tag Row (so is a repeated value) then close that parent tag up
      if ((range[row][column+1] != parent) || (row == topRow)){
        cell = cell + closeElement(parent);
      }
    }
  }

  //If we're not at the top of the spreadsheet, then move up a row and call this function again
  if (row > 0){
    row = row-1;
    return closeParents(range, topRow, column, row, cell, isUnnatural);

  //If we are, great, then just return the cell and end this madness
  }else{
    return cell;
  }
}

//A function that determines if a cell is the final cell
function isFinalCell(range, finalCol, column, row){

  var isFinal = true;

  //Search along the row given. If you find any with data in it, it can't be final
  for(var currentCol = column; currentCol <= finalCol; currentCol++){
    if (hasData(range[row][currentCol]))
    {
     isFinal = false;
    }
  }

  //If you've got to the finalCol without finding a single value, this is indeed the final cell.
  return isFinal;
}


//For finding attribtues of opens. These sit to the left of the topRow (ID's can be handled with @)
function wrapRowAndAttributes(range, titleCol, topRow, row, dataRow){
  var data = "";

  //Search to the left of the titleCol until you reach the end of the spreadsheet
  for (var attribCol = titleCol-1; attribCol >= 0; attribCol--){

    //when you find something with data
    if (range[row][attribCol] != ""){

      //add it to the data var with the attribute format
      data += applyRowWithAttribute(range[topRow][attribCol],range[row][attribCol]);
      }
    }

  //when you're at the end of the spread, return what you've put together
  return openElement(range[row][titleCol] + data) + dataRow + closeElement(range[row][titleCol]);

}

//Wrap data in XML
function wrapElement (outer, inner){
  open = openElement(outer);
  close = closeElement(outer);

  value = open + inner + close;
  return value;
}

//Apply an attribute (id='1')
function applyRowWithAttribute(attr, value){
  return " " + attr + " = '" + value +"'";
}

//Open xml tag
function openElement(inner){

  //If there's an @ in the text, sort that as an ID
  if (inner.indexOf('@') >= 0){
    var splitInner = inner.split('@',2);
    inner = splitInner[0];
    inner += applyRowWithAttribute("ID", splitInner[1]);
  }

  return "<" + inner + ">";
}

//Close xml tag
function closeElement(inner){

  //If there's an @ in the text, remove it
  if (inner.indexOf('@') >= 0){
    inner = inner.split('@',2);
    inner = inner[0];
  }

  return "</" + inner + ">" + "\n";
}

//Convert a column & row into location co-ords
function at(column, row){
 return convertColumn(column) + (row+1);
}

//Recursively convert a column into it's original format
function convertColumn(i) {
  return (i >= 26 ? convertColumn((i / 26 >> 0) - 1) : '') +
    'ABCDEFGHIJLKMNOPQRSTUVWXYZ'[i % 26 >> 0];
}

//Check if a cell has data
function hasData(cell){
  if (cell == null) return 0;
  return cell.length != 0;
}

//Google Sheets specific code -------------------------------------------------------

//Add the Export Sidebar to the UI
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Custom Menu')
      .addItem('Open Export Sidebar', 'showSidebar')
      .addToUi();
}

//Sidebar function
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Export with izzi')
      .setWidth(300);

  SpreadsheetApp.getUi()
      .showSidebar(html);
}

//This connects to the sidebar's Export button
function export(){

  var value = beautifyXml(escapeXml(doGet().getContent()));

  // Display a modal dialog box with custom HtmlService content.
  var htmlOutput = HtmlService

     .createHtmlOutput(value)
     .setWidth(500)
     .setHeight(500);

  SpreadsheetApp.getUi()
    .showModalDialog(htmlOutput, 'izzi Export');
}

//Escape xml - A StackOverflow solution!
function escapeXml(xml) {
    return xml.replace(/[<>&'"]/g, function (c) {
        switch (c) {
            case '<': return '&lt;';
            case '>': return '&gt;';
            case '&': return '&amp;';
            case '\'': return '&apos;';
            case '"': return '&quot;';
        }
    });
}

//Beautify
function beautifyXml(xml) {
    return xml.replace(/(&lt;)\/([a-zA-Z]+)(&gt;)/g, function (c) {
        return c + "<br/>"

    });
}

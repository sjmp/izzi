//HELLS VERSION - Basic release

//izzi code -------------------------------------------------------

//Entry point
function doGet() 
{
  return convertSheets(true, "");
}

function convertSheets(doGet, name)
{ 
  //Grab the current spreadsheet, find the sheets, prep the empty starting value
  var sheets = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheets();

  var sheetValue = "";
  
  //Loop through each sheet
  for (var i = 0; i < sheets.length; i++)
  {
    var sheet = sheets[i];
    var sheetName = sheet.getSheetName();

    //If there's not an exclamation in the sheet name
    if ((((sheetName.indexOf('!') < 0) && doGet))||((sheetName == name) && !(doGet)))
    {
      //Grab the sheet's values as a range
      var range = sheet.getSheetValues(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());

      //Pass into the convert Sheet function
      sheetValue += convertSheet(range, sheetName);

      //Wrap this resulting information with a Collection title
      sheetValue = wrapElement(sheetName +"Collection", sheetValue);
    }
  }

  //Add an xmlDataHeader
  var xmlHeader = "<?xml version='1.0' encoding='UTF-8'?>";

  //Return the value to the content service
  return ContentService.createTextOutput(xmlHeader + sheetValue)
     .setMimeType(ContentService.MimeType.XML);
}

//The function that converts a range of data into xml
function convertSheet(range, sheetName)
{
  //The total value to be delivered at the end
  var totalValue = "";

  //Find where our 'tag' is - Where the top row & the item title column meets.
  var topRow = 0;
  var titleCol = 0;
  while(range[topRow][titleCol] != "XML")
  {
    if (topRow < 998){
      topRow++;
    }else{
      topRow = 0;
      titleCol++;
    }
  }

  //Count down the rows from the tag until the next row is empty. This provides the final row with data.
  var finalRow = topRow;
  while (hasData(range[finalRow+1][titleCol]))
  {
    finalRow++;
  }

  //Count across the cols from the tag until the next col row is empty. This provides the final col with data.
  var finalCol = titleCol;
  while (hasData(range[topRow][finalCol+1]))
  {
    finalCol++;
  }

  //Iterate through every row under the top row
  for (var row = topRow+1; row <= finalRow; row++)
  {
    //This is the data we'll be returning at the end of this for loop.
    var dataRow = "";

    //This lock allows us to skip empty cells, and it should start off turned off, as we don't know what's in this row yet.
    var skipEmptyCells = false;

    //Now, for every column to the right of the title column
    for (var column = titleCol+1; column <= finalCol; column++){

      //The cell we are currently examining
      var cell = range[row][column];
      
      //The lowest level parent
      var baseParent = range[topRow][column];
      
      //The parent row above that
      var parentRow = topRow-1;

      //If the cell has some data in it
      if (hasData(cell))
      {
        //Open up it's parents
        dataRow += openParents(range, topRow, column, parentRow, "");

        //Wrap the cell inside the base parent
        dataRow += wrapElement(baseParent, cell);

        //Now close it's parents
        dataRow += closeParents(range, topRow, column, parentRow, "", false);

        //As we've found some data, we'll want the lock to be switched off, so we can close up the parents if we later find an empty cell
        skipEmptyCells = false;
      }
      else
      {
        //If we aren't skipping empty cells
        if (skipEmptyCells == false)
        {
          //If this is the final cell in the row, we'll need to force close it's parents
          if (isFinalCell(range, finalCol, column, row))
          {
            dataRow += closeParents(range, topRow, column, parentRow, "", true);
          }
          else
          {
            //If it's not the final cell, we'll need to close up any loose parents naturally before we go onto the next cell.
            dataRow += closeLooseParents(range, topRow, column, row, parentRow, "");
          }

          //Now we'll want to start skipping empty cells
          skipEmptyCells = true;
        }
      }
    }

    //wrap the sum of row inside of what the row started with. Include attributes.
    totalValue += wrapRowAndAttributes(range, titleCol, topRow, row, dataRow);
  }

  return wrapElement(sheetName, totalValue);
}

//A recursive function to determine all the parent tags we will need to open for this particular cell.
function openParents(range, topRow, column, row, cell)
{
  //Grab the parent we're currently examining
  var parent = range[row][column];

  //If the parent we're examining isn't empty...
  if (hasData(parent))
  {
    //And the cell to the left isn't the same (indicating a new parent)
    //or the cell is on the topRow (indicating it's a bottom-level parent)
    if ((range[row][column-1] != parent) || (row == topRow))
    {
      //Append the cell with a a new parent tag
      cell = openElement(parent) + cell;
    }
  }

  //If we're not at the top of the spreadsheet, move up a row and call this function again
  if (row > 0)
  {
    row = row-1;
    return openParents(range, topRow, column, row, cell);
  }
  
  //If we are at the top of the spreadsheet, it's time to stop and return the cell
  return cell;
}

//A function to close all parents that might still be open. 
//Used when a cell of data is empty but further cells on the row are not. 
function closeLooseParents(range, topRow, column, row, parentRow, cell)
{
  //If this cell is empty, and the cell to the right isn't empty
 if (!(hasData(range[row][column])) && (hasData(range[row][column+1])))
 {
    //Time to close up this cell's parents.
    return cell += closeParents(range, topRow, column, parentRow, true);
  }
  else
  {
    //Move on to the cell to right and call this function again.
    column = column + 1;
    return closeLooseParents(range, topRow, column, row, parentRow, cell);
  }
}

//This will close up any parents for the cell. 
//forceClose closes any parent if the one to the left isn't the same, non-forceCloses check the parents to the right.
//forceCloses are used to finish off data rows.
function closeParents(range, topRow, column, parentRow, cell, forceClose)
{
  //Grab the parent we're currently examining
  var parent = range[parentRow][column];

  //First, if the parent we're examining isn't empty...
  if (hasData(parent))
  {
    //And you're 'force closing' the parent
    if (forceClose)
    {
      //Compare the parent to the left - If it's the same, close the parent.
      if (range[parentRow][column-1] == parent){
        cell = cell + closeElement(parent);
      }
    }
    else
    {
      //If the parent to to the right isn't the same, or is on the topRow (so must be a repeated value), close the parent
      if ((range[parentRow][column+1] != parent) || (parentRow == topRow)){
        cell = cell + closeElement(parent);
      }
    }
  }

  //If we're not at the top of the spreadsheet, then move up a row and call this function again
  if (parentRow > 0)
  {
    parentRow = parentRow-1;
    return closeParents(range, topRow, column, parentRow, cell, forceClose);
  }
  
  //If we are, great, then just return the cell
  return cell;
}

//A function that determines if a cell is the final cell
function isFinalCell(range, finalCol, column, row)
{
  var isFinal = true;

  //Search along the row given. If you find any with data in it, it can't be final
  for(var currentCol = column; currentCol <= finalCol; currentCol++)
  {
    if (hasData(range[row][currentCol]))
    {
     isFinal = false;
    }
  }

  //If you've got to the finalCol without finding a single value, this is indeed the final cell.
  return isFinal;
}


//For finding attribtues of opens. These sit to the left of the topRow (ID's can be handled with @)
function wrapRowAndAttributes(range, titleCol, topRow, row, dataRow)
{
  var data = "";

  //Search to the left of the titleCol until you reach the end of the spreadsheet
  for (var attribCol = titleCol-1; attribCol >= 0; attribCol--)
  {
    //when you find something with data
    if (range[row][attribCol] != "")
    {
      //add it to the data var with the attribute format
      data += applyRowWithAttribute(range[topRow][attribCol],range[row][attribCol]);
      }
    }

  //when you're at the end of the spread, return what you've put together
  return openElement(range[row][titleCol] + data) + dataRow + closeElement(range[row][titleCol]);
}

//Wrap data in XML
function wrapElement (outer, inner)
{
  open = openElement(outer);
  close = closeElement(outer);
  return open + inner + close;
}

//Apply an attribute (id='1')
function applyRowWithAttribute(attr, value)
{
  return " " + attr + " = '" + value +"'";
}

//Open xml tag
function openElement(inner)
{
  //If there's an @ in the text, sort that as an ID
  if (inner.indexOf('@') >= 0)
  {
    var splitInner = inner.split('@',2);
    inner = splitInner[0];
    inner += applyRowWithAttribute("ID", splitInner[1]);
  }

  return "<" + inner + ">";
}

//Close xml tag
function closeElement(inner)
{
  //If there's an @ in the text, remove it
  if (inner.indexOf('@') >= 0){
    inner = inner.split('@',2);
    inner = inner[0];
  }

  return "</" + inner + ">" + "\n";
}

//Convert a column & row into location co-ords
function at(column, row)
{
 return convertColumn(column) + (row+1);
}

//Recursively convert a column into it's original format
function convertColumn(i) 
{
  return (i >= 26 ? convertColumn((i / 26 >> 0) - 1) : '') +
    'ABCDEFGHIJLKMNOPQRSTUVWXYZ'[i % 26 >> 0];
}

//Check if a cell has data
function hasData(cell)
{
  if (cell == null) return 0;
  return cell.length != 0;
}

//Google Sheets specific code -------------------------------------------------------

//Add the Export Sidebar to the UI
function onOpen() 
{
  SpreadsheetApp.getUi()
      .createMenu('Custom Menu')
      .addItem('Open Export Sidebar', 'showSidebar')
      .addToUi();
}

//Sidebar function
function showSidebar() 
{
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Export with izzi')
      .setWidth(300);

  SpreadsheetApp.getUi()
      .showSidebar(html);
}

//This connects to the sidebar's Export button
function export(sheetName)
{ 
  Logger.log(sheetName);
  
  var value = beautifyXml(escapeXml(convertSheets(false, sheetName).getContent()));

  // Display a modal dialog box with custom HtmlService content.
  var htmlOutput = HtmlService

     .createHtmlOutput(value)
     .setWidth(500)
     .setHeight(500);

  SpreadsheetApp.getUi()
    .showModalDialog(htmlOutput, 'izzi Export');
}

//Escape xml - A StackOverflow solution!
function escapeXml(xml) 
{
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
function beautifyXml(xml) 
{
    return xml.replace(/(&lt;)\/([a-zA-Z]+)(&gt;)/g, function (c) {
        return c + "<br/>"
    });
}

//Return the sheet names
function getNames(){
  var sheets = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheets();
  
  var sheetNames = new Array();
  
  for(var i=0; i < sheets.length; i++)
  {
    sheetNames.push(sheets[i].getName());
  }
  
  return sheetNames;

}

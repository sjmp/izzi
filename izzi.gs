//[X][Y]
//CORE VERSION OF IZZI
//PASTA VERSION - Now allowing for empty spaces. Some breakages, but much fewer!
//1. Starting an empty list with two values in breaks the object if it's mid-line
//2. Change the wrapper
//3. Multiple sheets
//4. Automatic drawers for the XML function
//5. Work out how to get the side control in
//6. Turn on/off xml encoding

function doGet() {

  //grab the Test spread by URL, get all sheets, and loop through adding to a totalvalue. This will need manual control eventually
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1dK0mIKqR__Y4oHJOzxij8La97PvKvOwY-8utJV0aGqY/edit#gid=0");
  var sheets = ss.getSheets();
  var sheetValue = "";

  for (var i = 0; i < sheets.length; i++){
    var sheetName = sheets[i].getSheetName();

    //If there's not an exclamation in the sheet name
    if (sheetName.indexOf('!') < 0)
    {
      var range = sheets[i].getRange('A:BI').getValues();
      sheetValue += convertSheet(range, sheets[i].getSheetName());
    }

  }

  sheetValue = wrapDataWithin("ScenarioCollection",sheetValue);
  var xmlStyle = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>";

  //This returns the actual value. Can't send multiple files - Is pity.
  return ContentService.createTextOutput(xmlStyle + sheetValue)
  .setMimeType(ContentService.MimeType.XML)
  ;
}

function convertSheet(range, sheetName){

   //The total value to be delivered at the end
  var totalValue = "";

  //find the XML tag and determine where the values begin
  var tagRow = 0;
  var tagCol = 0;
  while(range[tagRow][tagCol] != "XML"){
    //This maxes out at 49
    if (tagRow < 60){
      tagRow++;
    }else{
      tagRow = 0;
      tagCol++;
    }
  }

  //Count down the XML row the next row is empty. Provides last row of data.
  var finalRow = tagRow;
  while (range[finalRow+1][tagCol] != ""){
    finalRow++;
  }

  //Count down the XML col the next row is empty. Provides last col of data.
  var finalCol = tagCol;
  while (range[tagRow][finalCol+1] != ""){
    finalCol++;
  }

  //The XML report - for testing purposes
//  totalValue += wrapDataWithin("izziReport", "r" + finalRow + "c" + finalCol);

  //The main loop of the code. This runs through each cell within the boundaries of the XMLTag
  for (var row = tagRow+1; row <= finalRow; row++){

    //This is the data we'll be returning at the end.
    var dataRow = "";

    //The lock should be turned off for each new row.
    var skipEmptyRows = false;

    for (var column = tagCol+1; column <= finalCol; column++){

      //constants for the wrapping and adding of cell data for ease of reading/debugging
      var cell = range[row][column];
      var element = range[tagRow][column];
      var parentRow = tagRow-1;

      //If the cell has some value in it
      if (cell.length != 0)
      {
        //Determine parent opening tags
        dataRow += openParents(column,parentRow,"");

        //Then wrap the cell inside the element
        dataRow += wrapDataWithin(element,cell);

        //Determine parent closing tags
        dataRow += closeParents(column,parentRow,"",false);

        //No longer skip empty rows
        skipEmptyRows = false;
      }
      else
      {
        //If the lock is off...
        if (skipEmptyRows == false)
        {
          //If this is the final cell in the row, force close the parents
          if (isFinalCell(column, row, finalCol))
          {
            dataRow += closeParents(column,parentRow,"",true);
          }
          else
          {
            //Determine closing parents of that cell naturally
            dataRow += closeLooseParents(column, row, parentRow,"");
          }

          //Now turn back on the lock
          skipEmptyRows = true;
        }
      }
    }

    //wrap the sum of row inside of what the row started with. Include attributes.
    totalValue += wrapRowAndAttributes(row,dataRow);

  }

  totalValue = wrapDataWithin(sheetName, totalValue);

  return totalValue;

  //A function that determines if a cell isn't the final cell
  function isFinalCell(column, row, finalCol){

    var isFinal = true;

    //Search along this row from this point, if you find any with data in it, it can't be final
    for(var currentCol = column; currentCol <= finalCol; currentCol++){
      if (range[row][currentCol].length != 0)
      {
       isFinal = false;
      }
    }

    //If you've got to the finalCol without finding a single value, this is indeed the final cell.
    return isFinal;

  }

  //A function that seeks out the next empty cell with a neighbour to close it up
  function closeLooseParents(column, row, parentRow, cell){

    //If the next cell to the right is empty but has a cell to the next of it...
    if (range[row][column].length == 0 && range[row][column+1].length != 0){

      var testingvalue = ""
      //+ convertColumn(column) + (row+1) + " "
      ;

      //...Wrap up it's parents
      return cell += closeParents(column, parentRow, testingvalue, false);
    }
    else
    {
      //If not, move on to the next cell
      column = column + 1;
      return closeLooseParents(column, row, parentRow, cell);
    }

  }

  //A function to determine all the opening parent tags we will need for this particular row.
  function openParents(column,row,cell){

    //Grab the parent we're currently examining
    var parent = range[row][column];

    //First, if the parent we're examining isn't empty...
    if (parent != ""){

      //... And the cell to the left isn't the same, or is on the tag Row (so is a repeated value) open up a new parent tag
      if ((range[row][column-1] != parent)|| (row == tagRow)){
        cell = wrapOpen(parent) + cell;
      }
    }

    //If we're not at the top of the spreadsheet, then move up a row and call this function again
    if (row > 0){
      row = row-1;
      return openParents(column, row, cell);

    //If we are, great, then just return the cell and end this madness
    }else{
      return cell;
    }
  }


  //This will add every parent needed into the cell. It'll need the column of the current cell, a row to search up (starting with the tagRow) and the cell contents.
  //Unnatural closes are for premature closing at the end of a row, while natural are for all other situations
  function closeParents(column, row, cell, unnatural){

    //Grab the parent we're currently examining
    var parent = range[row][column];

    //First, if the parent we're examining isn't empty...
    if (parent != ""){

      //if the force is on...
      if (unnatural)
      {
        //... And the cell to the left isn't the same, close that parent tag up
        if (range[row][column-1] == parent){
          cell = cell + wrapClose(parent);
        }
      }
      else
      {
        //... And the cell to the right isn't the same, or is on the tag Row (so is a repeated value) then close that parent tag up
        if ((range[row][column+1] != parent) || (row == tagRow)){
          cell = cell + wrapClose(parent);
        }
      }
    }

    //If we're not at the top of the spreadsheet, then move up a row and call this function again
    if (row > 0){
      row = row-1;
      return closeParents(column, row, cell, unnatural);

    //If we are, great, then just return the cell and end this madness
    }else{
      return cell;
    }
  }

  //For finding attribtues of opens. These sit to the left of the tagRow (ID's can be handled with @)
  function wrapRowAndAttributes(row,dataRow){
    var data = "";

    //Search to the left of the tagCol until you reach the end of the spreadsheet
    for (var attribCol = tagCol-1; attribCol >= 0; attribCol--){

      //when you find something with data
      if (range[row][attribCol] != ""){

        //add it to the data var with the attribute format
        data += formatAttrRow(range[tagRow][attribCol],range[row][attribCol]);
      }
    }

    //when you're at the end of the spread, return what you've put together
    return wrapOpen(range[row][tagCol] + data) + dataRow + wrapClose(range[row][tagCol]);

  }
}

//basic wrapping function
function wrapDataWithin (outer, inner){
  open = wrapOpen(outer);
  close = wrapClose(outer);

  value = open + inner + close;
  return value;
}

//format attribute row
function formatAttrRow(attr, value){
  return " " + attr + " = '" + value +"'";
}

//basic open xml - this can take ID's as @ symbols
function wrapOpen(inner){

  //If there's an @ in the text, sort that as an ID
  if (inner.indexOf('@') >= 0){
    var splitInner = inner.split('@',2);
    inner = splitInner[0];
    inner += formatAttrRow("ID", splitInner[1]);
  }

  var open = "<";
  var close = ">";
  return open + inner + close;
}

//basic close xml
function wrapClose(inner){

  //If there's an @ in the text, remove it
  if (inner.indexOf('@') >= 0){
    inner = inner.split('@',2);
    inner = inner[0];
  }

  var open = "</";
  var close = ">";
  return open + inner + close + "\n";
}

//Convert a column & row into location co-ords
function at(column, row){
 return convertColumn(column) + (row+1);
}

function convertColumn(i) {
  return (i >= 26 ? convertColumn((i / 26 >> 0) - 1) : '') +
    'ABCDEFGHIJLKMNOPQRSTUVWXYZ'[i % 26 >> 0];
}

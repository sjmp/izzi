//[X][Y]
//CORE VERSION OF IZZI
//BURRITO VERSION - Working towards release
//1. Starting an empty list with two values in breaks the object if it's mid-line
//2. Change the wrapper
//3. Multiple sheets
//4. Automatic drawers for the XML function
//5. Work out how to get the side control in
//6. Turn on/off xml encoding
//7. Multiple files

function doGet() {
  
  //grab the current spreadsheet, find the sheets, prep the empty starting value
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheetValue = "";
  
  //Loop through the sheets
  for (var i = 0; i < sheets.length; i++){
    var sheet = sheets[i];
    var sheetName = sheet.getSheetName();

    //If there's not an exclamation in the sheet name
    if (sheetName.indexOf('!') < 0)
    {
      //Grab the sheet's values as a range
      var range = sheet.getSheetValues(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
      
      //Pass into the convert Sheet function
      sheetValue += convertSheet(range, sheetName);
      
      //Wrap this resulting information with a Collection title
      sheetValue = wrapDataWithin(sheetName +"Collection",sheetValue);
    }
  }
 
  //Add the data header for the browser to read
  var xmlHeader = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>";
  
  //Return the value to the content service
  return ContentService.createTextOutput(xmlHeader + sheetValue)
  .setMimeType(ContentService.MimeType.XML)
  ;
}

//The function that converts a single sheet into xml
function convertSheet(range, sheetName){
  
  //The total value to be delivered at the end
  var totalValue = "";
  
  //Find where our 'tag' is - Where the top row & the item title column meets. (TODO: Make this user input)
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
      var element = range[topRow][column];
      var parentRow = topRow-1;
      
      //If the cell has some value in it
      if (hasData(cell))
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
      if ((range[row][column-1] != parent)|| (row == topRow)){
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

  
  //This will add every parent needed into the cell. It'll need the column of the current cell, a row to search up (starting with the topRow) and the cell contents.
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
        if ((range[row][column+1] != parent) || (row == topRow)){
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
  
  //For finding attribtues of opens. These sit to the left of the topRow (ID's can be handled with @)
  function wrapRowAndAttributes(row,dataRow){
    var data = "";
    
    //Search to the left of the titleCol until you reach the end of the spreadsheet
    for (var attribCol = titleCol-1; attribCol >= 0; attribCol--){
      
      //when you find something with data
      if (range[row][attribCol] != ""){
        
        //add it to the data var with the attribute format
        data += formatAttrRow(range[topRow][attribCol],range[row][attribCol]);
      }
    }
   
    //when you're at the end of the spread, return what you've put together
    return wrapOpen(range[row][titleCol] + data) + dataRow + wrapClose(range[row][titleCol]);
    
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

//Convert a column into it's original format
function convertColumn(i) {
  return (i >= 26 ? convertColumn((i / 26 >> 0) - 1) : '') +
    'ABCDEFGHIJLKMNOPQRSTUVWXYZ'[i % 26 >> 0];
}

function hasData(cell){
  return cell.length != 0;
}


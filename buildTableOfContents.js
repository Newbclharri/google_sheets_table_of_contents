/**
 * Builds the table of contents for sheets in the active spreadsheet
 * @param {string} sheetName: default value is "Table of Contents"
 * @param {integer} numRowsHeader: default value is 1
 * @param {boolean} sort: default value is false
 */
function buildTableOfContents(options = {}) {
  let {sheetName, numRowsHeader, sort} = options;
  sheetName = sheetName || "Table of Contents";
  numRowsHeader = numRowsHeader || 1;
  sort = sort || false;
  ////VARIABLES////
  const ss = SpreadsheetApp.getActive();
  //create default table of contents sheet if one does not exit
  if(!ss.getSheetByName(sheetName)) ss.insertSheet(sheetName);

  //get the table of contents sheet
  const sheetTableOfCont = ss.getSheetByName(sheetName);

  //list of all sheets in the active Spreadsheet
  const sheets = ss.getSheets();

  //Spreadsheet url for creating hypertext
  const ssUrl = ss.getUrl();

  //remove the table of contents sheet from list of sheets to add to table of contents
  const filteredSheets = sheets.filter(sheet => sheet.getName() !== sheetTableOfCont.getName());

  const lastRow = sheetTableOfCont.getLastRow();

  //variables for navigating table of contents rows and columns
  let rowStart = numRowsHeader + 1, columnStart = 1, numRows = 1;

  try{
    //create the table of contents sheet header
    range = sheetTableOfCont.getRange(1,1)
      .setValue("Table of Contents")
      .setFontWeight("bold")
    sheetTableOfCont.setFrozenRows(1);
    
    //clear all data below header
    range = sheetTableOfCont.getRange(rowStart,columnStart,lastRow)
    range.clear();   

  }catch(err){
    console.log(ss.getName())
    console.log("buildTableOfContents: ", err)
  }finally{
    filteredSheets.forEach((sheet, index) => {
      const sheetId = sheet.getSheetId();
      const sheetName = sheet.getName();
      const linkString = ssUrl + "#gid=" + sheetId;
      const linkStyle = SpreadsheetApp.newTextStyle()
        .setUnderline(false)
        .setBold(true)
        .build();
      const link = SpreadsheetApp.newRichTextValue()
        .setText(sheetName)
        .setLinkUrl(linkString)
        .setTextStyle(linkStyle)
        .build();

      //insert a link to each sheet into each subsequent row        
      range = sheetTableOfCont.getRange(rowStart + index, columnStart, numRows);
      range.setRichTextValue(link);      
    });

    //optionally sort table of contents alphabetically
    if(sort) sheetTableOfCont.sort(1);
  }
}

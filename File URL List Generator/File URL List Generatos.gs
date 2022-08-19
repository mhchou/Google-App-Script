//the first time it will ask for write access permission
//allow in order to create new spreadsheet

function Allurl() {
  //Get all the files inside the named folder
  var foldername = "Pictures";
  var folders = DriveApp.getFoldersByName(foldername);
  var contents = folders.next().getFiles();
/*
  //Create new spreadsheet
  var folderlisting = "Picture URLs" //Name of the spreadsheet being created
  var writefolder = DriveApp.getFolderById("<Id>"); //Write to this folder //Energy Assets Pictures
  var ss = SpreadsheetApp.create(folderlisting);
  var temp = DriveApp.getFileById(ss.getId());
  writefolder.addFile(temp);
  DriveApp.getRootFolder().removeFile(temp);
  var sheet = ss.getActiveSheet();
  
  sheet.appendRow(["Number", "Link", "LastModifiedDate"]);
*/  
  
  //Or append to an extisting sheet and append to it
  var sheet = SpreadsheetApp.openById("<Id>").getActiveSheet();

  var file;
  var name;
  var newname;
  var link;
  var lastupdate;
  var updateddate;

  while (contents.hasNext()) {
    file = contents.next();
    name = file.getName();
    newname = name.substr(0, name.lastIndexOf(".")); //remove .extension
    link = file.getUrl();
    lastupdate = file.getLastUpdated();
    updateddate = Utilities.formatDate(lastupdate, "GMT-7", "MM/dd/yyyy"); //formate date
    sheet.appendRow([newname, link, updateddate]); //0 will not stay if # start with 0
  }

  //sort by number ascending
  var getrange = sheet.getRange("A:A");
  var LastRow = getrange.getLastRow();
  var range = sheet.getRange("A2:C" + LastRow);
  range.sort(1);
}

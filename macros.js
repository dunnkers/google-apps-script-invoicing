/** @OnlyCurrentDoc */

// let user choose folder manually
// @see https://developers.google.com/apps-script/guides/dialogs
var INVOICE_FOLDER = '1S7yfG1OfsjldHx69Z6yAJ2R8kJVY41KK';

function onOpen() {
    SpreadsheetApp
        .getUi()
        .createMenu('Invoicing')
        .addItem('Create invoice', 'createInvoice')
        .addToUi();
} // onOpen()

// https://gist.github.com/ixhd/3660885
function exportSheet(sheetName) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();

    // hide sheets
    for (var i = 0; i < sheets.length; i++) {
        if (sheets[i].getSheetName() != sheetName) {
            sheets[i].hideSheet();
        }
    }

    // save sheet as a PDF file
    var file = DriveApp.createFile(ss.getBlob());

    // show sheets
    for (var i = 0; i < sheets.length; i++) {
        sheets[i].showSheet();
    }

    // re-select previously active sheet
    SpreadsheetApp.setActiveSheet(sheet);

    return file;
}

function sendInvoiceByEmail() {
    var name = Browser.inputBox('ID Check', 'Enter your name', Browser.Buttons.OK_CANCEL);
}

function createInvoice() {
    var file = exportSheet('Factuur');
    var today = Utilities.formatDate(new Date(), "GMT+1", "yyyy-MM-dd");
    file.setName(today);
    
    // move file by making a copy
    var folder = DriveApp.getFolderById(INVOICE_FOLDER);
    var name = today + ' Factuur ';
    file.makeCopy(name, folder);
    Browser.msgBox('Successfully saved ' + file.getName() + ' to folder `' + 
        folder.getName() + '`');
    file.setTrashed(true);
}

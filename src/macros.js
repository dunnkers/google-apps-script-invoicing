/** @OnlyCurrentDoc */

// let user choose folder manually
// @see https://developers.google.com/apps-script/guides/dialogs
var INVOICE_FOLDER = '1S7yfG1OfsjldHx69Z6yAJ2R8kJVY41KK';
var INVOICE_SHEET = 'Factuur';
var INVOICE_CREATION_SHEET = 'Factureren';

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
    const ok = sheet => sheet.getSheetName() !== sheetName;

    // params
    var invoice = SpreadsheetApp.getActiveSheet();
    var namedRanges = invoice.getNamedRanges();
    var customer = namedRanges.forEach(function (namedRange) {
        if (namedRange.getName() === 'FactuurKlant') {
            Browser.msgBox(namedRange.getRange().getValue());
        }
    });

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

function getParams() {
    var sheet = SpreadsheetApp.get();

}

function createInvoice() {
    var file = exportSheet(INVOICE_SHEET);
    var today = Utilities.formatDate(new Date(), "GMT+1", "yyyy-MM-dd");
    
    // move file by making a copy
    var folder = DriveApp.getFolderById(INVOICE_FOLDER);
    var name = today + ' Factuur ';
    var newFile = file.makeCopy(name, folder);
    Browser.msgBox('Successfully saved `' + newFile.getName() + 
        '` to folder `' + folder.getName() + '`');
    file.setTrashed(true);
}

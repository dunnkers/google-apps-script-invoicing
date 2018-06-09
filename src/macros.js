/** @OnlyCurrentDoc */

// let user choose folder manually
// @see https://developers.google.com/apps-script/guides/dialogs
const INVOICE_FOLDER = '1S7yfG1OfsjldHx69Z6yAJ2R8kJVY41KK';
const INVOICE_SHEET = 'Factuur';
const INVOICE_CREATION_SHEET = 'Factureren';

const PARAMS = [
    'FactuurKlant',
    'FactuurNetto',
    'FactuurBTW',
    'FactuurTotaal',
    'FactuurNummer',
    'FactuurPeriode'
];

function onOpen() {
    SpreadsheetApp
        .getUi()
        .createMenu('Invoicing')
        .addItem('Create invoice', 'createInvoice')
        .addToUi();
} // onOpen()

// https://gist.github.com/ixhd/3660885
function exportSheet(sheetName) {
    const sheet = SpreadsheetApp.getActiveSheet();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();

    // hide sheets
    const filterSheets = sheet => sheet.getSheetName() !== sheetName;
    sheets.filter(filterSheets).forEach(sheet => sheet.hideSheet());

    // params
    const invoice = SpreadsheetApp.getActiveSheet();

    // save sheet as a PDF file
    const file = DriveApp.createFile(ss.getBlob());

    // show sheets
    sheets.forEach(sheet => sheet.showSheet());

    // re-select previously active sheet
    SpreadsheetApp.setActiveSheet(sheet);

    return file;
}

function sendInvoiceByEmail() {
    const name = Browser.inputBox('ID Check', 'Enter your name', Browser.Buttons.OK_CANCEL);
}

function getParams(sheetName, toMap) {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    const params = {};
    const namedRanges = sheet.getNamedRanges();
    namedRanges
        .filter(namedRange => toMap.includes(namedRange.getName()))
        .forEach(mapIt => {
            params[mapIt.getName()] = mapIt.getRange().getValue();
        });

    return params;
}

function getFirstFreeRow() {
    cs.getRange('A3:E3')
}

function getQuarter(date) {
    var month = date.getMonth() + 1;
    return (Math.ceil(month / 3));
}

function createInvoice() {
    const params = getParams(INVOICE_SHEET, PARAMS);
    const file = exportSheet(INVOICE_SHEET);
    const today = Utilities.formatDate(new Date(), "GMT+1", "yyyy-MM-dd");

    const { FactuurKlant, FactuurBTW, FactuurNetto, FactuurTotaal,
        FactuurNummer, FactuurPeriode } = params;
    
    // insert new invoice data in creation sheet
    const cs = SpreadsheetApp.getActive().getSheetByName(INVOICE_CREATION_SHEET);
    let i = 3, range;
    while (true) {
        const empty = value => !value || value === '';
        range = cs.getRange(`A${i}:F${i}`);
        const values = range.getValues()[0];
        if (values.every(empty)) {
            break;
        }
        i ++;
    }
    range.setValues([[FactuurNummer, FactuurPeriode, FactuurKlant, FactuurNetto, FactuurBTW, FactuurTotaal] ]);


    // update invoice number
    const is = SpreadsheetApp.getActive().getSheetByName(INVOICE_SHEET);
    is.getNamedRanges().forEach(namedRange => {
        if (namedRange.getName() === 'FactuurNummer') {
            namedRange.getRange().setValue(FactuurNummer + 1);
        }
    });
    
    // move file by making a copy
    const folder = DriveApp.getFolderById(INVOICE_FOLDER);
    const year = FactuurPeriode.getFullYear();
    const quarter = getQuarter(FactuurPeriode);
    const month = FactuurPeriode.toLocaleDateString('nl', {
        month: 'long'
    });
    const name = `${today} Factuur ${FactuurKlant} ${year}-Q${quarter}-${month}`;
    file.setName(name);
    const newFile = file.makeCopy(name, folder);
    Browser.msgBox(`Successfully saved \`${newFile.getName()}\` to folder \` 
                    ${folder.getName()}\``);
    file.setTrashed(true);
}

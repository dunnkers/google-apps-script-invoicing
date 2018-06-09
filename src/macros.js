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

function onOpen() { // eslint-disable-line
    SpreadsheetApp
        .getUi()
        .createMenu('Invoicing')
        .addItem('Create invoice', 'createInvoice')
        .addToUi();
} // onOpen()

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

function getQuarter(date) {
    var month = date.getMonth() + 1;
    return (Math.ceil(month / 3));
}

function createInvoice() { // eslint-disable-line
    const params = getParams(INVOICE_SHEET, PARAMS);
    // const file = exportSheet(INVOICE_FOLDER, INVOICE_SHEET);
    const today = Utilities.formatDate(new Date(), 'GMT+1', 'yyyy-MM-dd');
    
    // UPDATE IN RANGE DATA
    const { FactuurKlant, FactuurBTW, FactuurNetto, FactuurTotaal,
        FactuurNummer, FactuurPeriode } = params;

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
    const month = Utilities.formatDate(FactuurPeriode, 'Europe/Amsterdam', 'MMM');
    const name = `${today} Factuur ${FactuurKlant} ${year}-Q${quarter}-${month}`;
    // file.setName(name);


    // https://docs.google.com/spreadsheets/d/ID/export?format=pdf&gid=SHEET_ID
    // https://stackoverflow.com/questions/48163613/google-script-blob-conversion
    
    const paramss = {
        method: 'get',
        headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
    };
    const id = SpreadsheetApp.getActive().getId();
    const sheetid = SpreadsheetApp.getActive().getSheetByName(INVOICE_SHEET).getSheetId();
    const url = `https://docs.google.com/spreadsheets/d/${id}/export` +
                `?format=pdf&gid=${sheetid}`;
    const blob = UrlFetchApp.fetch(url, paramss).getBlob().setName(name);
    const dir = DriveApp.getFolderById(INVOICE_FOLDER);
    const file = dir.createFile(blob);

 

    // insert new invoice data in creation sheet
    const cs = SpreadsheetApp.getActive().getSheetByName(INVOICE_CREATION_SHEET);
    let i = 3, range;
    while (true) { // eslint-disable-line
        const empty = value => !value || value === '';
        range = cs.getRange(`A${i}:G${i}`);
        const values = range.getValues()[0];
        if (values.every(empty)) {
            break;
        }
        i++;
    }
    const link = `=HYPERLINK("${file.getUrl()}"; "${name}")`;
    range.setValues([[FactuurNummer, FactuurPeriode, FactuurKlant, FactuurNetto, FactuurBTW, FactuurTotaal, link]]);


    SpreadsheetApp.getActiveSpreadsheet().toast( 
        `Successfully saved \`${file.getName()}\` to folder \` 
                    ${folder.getName()}\``,
        'Invoice created!');
}

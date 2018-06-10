import { getParams, saveSheetAsPDF, getInvoiceName } from './util';
import {
    INVOICE_CREATION_SHEET, INVOICE_FOLDER, INVOICE_SHEET, PARAMS
} from './settings';

global.onOpen = () => {
    SpreadsheetApp
        .getUi()
        .createMenu('Invoicing')
        .addItem('Create invoice', 'createInvoice')
        .addToUi();
};

global.createInvoice = () => {
    const {
        FactuurNummer,
        FactuurPeriode,
        FactuurKlant,
        FactuurBTW,
        FactuurNetto,
        FactuurTotaal
    } = getParams(INVOICE_SHEET, PARAMS);


    // update invoice number
    const is = SpreadsheetApp.getActive().getSheetByName(INVOICE_SHEET);
    is.getNamedRanges().forEach(namedRange => {
        if (namedRange.getName() === 'FactuurNummer') {
            namedRange.getRange().setValue(FactuurNummer + 1);
        }
    });

    // construct invoice filename
    const name = getInvoiceName(FactuurPeriode, FactuurKlant);
    
    // Save invoice sheet as PDF
    const folder = DriveApp.getFolderById(INVOICE_FOLDER);
    const file = saveSheetAsPDF(INVOICE_SHEET, INVOICE_FOLDER, name);
    SpreadsheetApp.getActiveSpreadsheet().toast(
        `Successfully saved \`${file.getName()}\` to folder \` 
                    ${folder.getName()}\``,
        'Invoice created!');


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


};
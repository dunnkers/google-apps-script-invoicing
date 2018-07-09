import { saveSheetAsPDF, getInvoiceName, insertRow, mapNamedRanges } from './util';
import {
    INVOICE_MAPPING,
    INVOICE_FOLDER,
    EXPORTED_INVOICES_RANGE
} from './settings';
import { InvoiceEmail } from './invoice_email';
import { EmailPreview } from './email_preview';

global.onOpen = () => {
    SpreadsheetApp
        .getUi()
        .createMenu('Invoicing')
        .addItem('Create invoice', 'createInvoice')
        .addItem('Make draft email', 'createInvoiceEmail')
        .addToUi();
};

global.createInvoice = () => {
    // spreadsheet
    const spreadsheet = SpreadsheetApp.getActive();

    // fetch invoice data
    const invoiceData = mapNamedRanges(INVOICE_MAPPING);
    const {
        invoiceNumber: number,
        invoicePeriod: period,
        invoiceCustomer: customer
    } = invoiceData;

    // construct invoice filename
    const name = getInvoiceName(period.getValue(), customer.getValue());

    // Save invoice sheet as PDF
    const invoiceSheet = number.getSheet().getName();
    const folder = DriveApp.getFolderById(INVOICE_FOLDER);
    const file = saveSheetAsPDF(invoiceSheet, INVOICE_FOLDER, name);
    spreadsheet.toast(`Successfully saved \`${file.getName()}\` ` +
                      `to folder \`${folder.getName()}\``, 'Invoice created!');

    // insert new invoice data in creation sheet
    const hyperLink = `=HYPERLINK("${file.getUrl()}"; "${name}")`;
    const dataArray = Object.keys(invoiceData).map(key => {
        return invoiceData[key].getValue();
    });
    const rowData = [
        ...dataArray,
        hyperLink
    ];
    insertRow(EXPORTED_INVOICES_RANGE, rowData);

    // update invoice number
    number.setValue(number.getValue() + 1);
};

global.createInvoiceEmail = () => {
    // var file = DriveApp.getFileById('1rzcAtcGz781CT0mxK_FZm6XhKAa7qlBV');
    // GmailApp.createDraft('jeroenoverschie@gmail.com', 'Factuur juli', 'Zie bijlage', {
    //     attachments: [file.getAs(MimeType.PDF)],
    //     htmlBody: InvoiceEmail
    // });

    
    let preview = EmailPreview(InvoiceEmail);

    // Display a modal dialog box with custom HtmlService content.
    var htmlOutput = HtmlService
        .createHtmlOutput(preview)
        .setWidth(500)
        .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Email invoice');
};

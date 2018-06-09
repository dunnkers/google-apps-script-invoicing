import Util from './util';

export const createInvoice = () => {
    SpreadsheetApp.getActiveSpreadsheet().toast('gozer', Util.getDuder());
};
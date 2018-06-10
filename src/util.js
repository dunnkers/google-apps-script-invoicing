import { SHEET_EXPORT_URL } from './settings';

export const getParams = (sheetName, toMap) => {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    const namedRanges = sheet.getNamedRanges();

    const params = {};
    namedRanges
        .filter(namedRange => toMap.includes(namedRange.getName()))
        .forEach(mapIt => {
            params[mapIt.getName()] = mapIt.getRange().getValue();
        });

    return params;
};

export const getQuarter = date => {
    var month = date.getMonth() + 1;
    return (Math.ceil(month / 3));
};

export const saveSheetAsPDF = (sheetName, driveFolderId, name) => {
    // Construct export URL
    const spreadsheet = SpreadsheetApp.getActive();
    const id = spreadsheet.getId();
    const sheetid = spreadsheet.getSheetByName(sheetName).getSheetId();
    const url = `${SHEET_EXPORT_URL}/${id}/export?format=pdf&gid=${sheetid}`;

    // Use UrlFetchApp to trigger PDF download
    // https://stackoverflow.com/questions/48163613/google-script-blob-conversion
    const params = {
        method: 'get',
        headers: {
            Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
        }
    };
    const blob = UrlFetchApp.fetch(url, params).getBlob();
    
    // Finally save PDF to folder
    const dir = DriveApp.getFolderById(driveFolderId);
    blob.setName(name);
    const file = dir.createFile(blob);
    return file;
};

export const getInvoiceName = (date, customer) => {
    const year = date.getFullYear();
    const quarter = getQuarter(date);
    const month = Utilities.formatDate(date, 'Europe/Amsterdam', 'MMM');
    const today = Utilities.formatDate(new Date(), 'GMT+1', 'yyyy-MM-dd');
    const name = `${today} Factuur ${customer} ${year}-Q${quarter}-${month}`;
    return name;
};
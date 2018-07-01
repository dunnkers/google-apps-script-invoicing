const SHEET_EXPORT_URL = 'https://docs.google.com/spreadsheets/d';

export const mapNamedRanges = mapping => {
    const spreadsheet = SpreadsheetApp.getActive();
    const result = {};

    Object.keys(mapping).forEach(key => {
        let namedRange = mapping[key];
        result[key] = spreadsheet.getRangeByName(namedRange);
    });

    return result;
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
    blob.setName(name.endsWith('.pdf') ? name : `${name}.pdf`); // ensure ext.
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

export const getFirstEmptyRow = (range) => {
    const offset = range.getRow();
    const matrix = range.getValues();
    const empty = value => !value || value === '';

    for (let i = 0; i < matrix.length; i ++) {
        if (matrix[i].every(empty)) {
            return offset + i;
        }
    }

    return null;
};

export const insertRow = (rangeName, rowData) => {
    const spreadsheet = SpreadsheetApp.getActive();
    const range = spreadsheet.getRangeByName(rangeName);
    
    if (range === null) {
        return false;
    }

    const emptyRow = getFirstEmptyRow(range);
    if (emptyRow === null) {
        return false;
    }

    const colIndex = range.getColumn();
    const width = range.getWidth();
    const rangeSheet = range.getSheet();
    const insertionRange = rangeSheet.getRange(emptyRow, colIndex, 1, width);
    insertionRange.setValues([ rowData ]);

    return true;
};
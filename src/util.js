export const getParams = (sheetName, toMap) => {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    const params = {};
    const namedRanges = sheet.getNamedRanges();
    namedRanges
        .filter(namedRange => toMap.includes(namedRange.getName()))
        .forEach(mapIt => {
            params[mapIt.getName()] = mapIt.getRange().getValue();
        });

    return params;
};

export const getQuarter = (date) => {
    var month = date.getMonth() + 1;
    return (Math.ceil(month / 3));
};

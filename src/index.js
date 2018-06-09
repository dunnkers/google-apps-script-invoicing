import Util from './util';

export const dudertje = function maat() {
    console.log('lolmaat');
};

export default function gozerwijfie() {
    SpreadsheetApp.getActiveSpreadsheet().toast('gozer', Util.getDuder());
}

global.hello = function () {
    Logger.log('hello');
};
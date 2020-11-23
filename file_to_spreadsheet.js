const mainSheet = 'Sheet1'
const driveFolder = 'updated entries 112220'
const protectedSheets = ['Number Occurrence',]

function checkHistory () {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetArray = ss.getSheets();

    let logArray = sheetArray.map((sheet) => sheet.getName());

    return logArray;
};

function checkFiles () {
    const dApp = DriveApp;
    const folderIterate = dApp.getFoldersByName(driveFolder);
    const folder = folderIterate.next();
    
    const filesIterate = folder.getFiles();
    let fileArray = [];
    
    while (filesIterate.hasNext()) {
        let file = filesIterate.next();
        fileArray.push(file.getName());
    };
    return fileArray;
};

function createTab(name) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.insertSheet(name);

    importCSV(name);
    const codeWord = getWord(name);
    ss.getRange('K2:K').setValue(codeWord);
};

function getWord(name) {
    const splitFile = name.split('_')
    const splitName = splitFile[2].split('-');
    return splitName[0];
};

function importCSV(name) {
    let file = DriveApp.getFilesByName(name).next();
    let csvData = Utilities.parseCsv(file.getBlob().getDataAsString());

    sheet = SpreadsheetApp.getActiveSheet();
    sheet.getRange(1,1,csvData.length, csvData[0].length).setValues(csvData);
};

function updateFormula() {
    const filesList = checkHistory();

    let ranges = [];
    for (i=0;i<filesList.length;i++) {
      if (filesList[i] != mainSheet && !protectedSheets.includes(filesList[i])) {
            ranges.push(`'${filesList[i]}'!A2:K`)
        };
    };
    
    const data = ranges.join(';');
    const formula = `=QUERY({${data}},"select * where Col8 <> ''")`

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
    const cell = sheet.getRange('A2');

    cell.setValue(formula);
};

function updateTabs() {
    driveFiles = checkFiles();
    historyFiles = checkHistory();

    for (i=0;i<driveFiles.length;i++) {
        if(historyFiles.includes(driveFiles[i])) {
            continue;
        } else {
            createTab(driveFiles[i]);
        };
    };
    updateFormula();
};
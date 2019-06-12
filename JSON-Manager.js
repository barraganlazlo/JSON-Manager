var documentProperties = PropertiesService.getDocumentProperties();


function onInstall(e) {
    onOpen(e);
}

function onOpen(e) {
    SpreadsheetApp.getUi()
        .createAddonMenu()
        .addItem('Open', 'openSideBar')
        .addToUi();
}

function convertToJSON() {
    var columnArray = documentProperties.getProperty("columnArray");
    var sheet = SpreadsheetApp.getActiveSheet();
    var nbFrozen = sheet.getFrozenRows();
    var data = sheet.getDataRange().getValues();
    var nbColumns = data[nbFrozen - 1].length;
    var json = '[';
    var names = new Array(nbColumns);
    Logger.log(JSON.stringify(data[1][0]));
    var ts = "";
    for (var i = 0; i < nbFrozen; i++) {
        var lastname = null;
        for (var j = 0; j < nbColumns; j++) {
            var name = data[i][j];
            if (name != null && name != "") {
                if (names[j]) {
                    names[j] += JSON.stringify(name);
                } else {
                    names[j] = JSON.stringify(name);
                }
                if (i < nbFrozen - 1) {
                    names[j] += " : { ";
                }
            } else if (i == nbFrozen - 1) {
                if (lastname != null && lastname != "") {
                    names[j] = "}";
                }
            }
            lastname = name;
        }
    }
    Logger.log(names);
    for (var i = nbFrozen; i < data.length; i++) {
        if (i > nbFrozen) {
            json += ',';
        }
        json += '{';
        for (var j = 0; j < nbColumns; j++) {
            var value = data[i][j];
            if (value) {

            }
            if (names[j] != null && names[j] != "") {
                if (names[j] == '}') {
                    json += names[j]
                } else {
                    if (j > 0) {
                        json += ",";
                    }
                    json += names[j] + ' : ' + JSON.stringify(value);
                }
            }
        }
        json += '}';
    }
    json += ']';
    Logger.log(json);
    return json;
}

function getColumnArray() {
    return documentProperties.getProperty("columnArray");
}

function setColumnArray(array) {
    return documentProperties.setProperty({ columnArray: array });
}

function addColumnArray(columnId) {
    var columnArray = documentProperties.getProperty("columnArray");
    if (columnArray == null) {
        columnArray = new Array();
    }
    if (!columnArray.includes(columnId)) {
        columnArray.push(columnId);
    }
    setColumnArray(columnArray);

}

function removeColumnArray(columnId) {
    var columnArray = documentProperties.getProperty("columnArray");
    if (!columnArray) {
        columnArray = new Array();
    }
    columnArray = columnArray.filter(el => el != columnId);
    setColumnArray(columnArray);
}

function download() {

}

function openSideBar() {
    var html = HtmlService.createTemplateFromFile('sidebar').evaluate().setTitle("JSON Manager");
    SpreadsheetApp.getUi().showSidebar(html);
}

function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent();
}
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
    var columnArray = getColumnArray();
    var sheet = SpreadsheetApp.getActiveSheet();
    var nbFrozen = sheet.getFrozenRows();
    var data = sheet.getDataRange().getValues();
    var nbColumns = data[nbFrozen - 1].length;
    var json = '[';
    var names = new Array(nbColumns);
    console.log(JSON.stringify(data[1][0]));
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
    console.log(names);
    for (var i = nbFrozen; i < data.length; i++) {
        if (i > nbFrozen) {
            json += ',';
        }
        json += '{';
        for (var j = 0; j < nbColumns; j++) {
            var value = data[i][j];
            if (names[j] != null && names[j] != "") {
                if (names[j] == '}') {
                    json += names[j]
                } else {
                    if (j > 0) {
                        json += ",";
                    }
                    //arrays
                    if (columnArray.indexOf(numberToLetter(j)) > -1) { //=includes
                        if (value && value != "") {
                            value = value.replace(/\s/g, '');
                            value = value.split(",");
                            console.log(value);
                        } else {
                            velue = null;
                        }
                    }
                    json += names[j] + ' : ' + JSON.stringify(value);
                }
            }
        }
        json += '}';
    }
    json += ']';
    console.log(json);
    return json;
}

function getColumnArray() {
    return JSON.parse(PropertiesService.getDocumentProperties().getProperty("columnArray"));
}

function setColumnArray(array) {
    return PropertiesService.getDocumentProperties().setProperties({ "columnArray": JSON.stringify(array) });
}

function addColumnArray(columnId) {
    var columnArray = getColumnArray();
    console.log(columnArray);
    if (columnArray == null || !columnArray || columnArray == "") {
        columnArray = [];
    }
    if (!(columnArray.indexOf(columnId) > -1)) { //=includes(el)
        console.log(JSON.stringify(columnArray));
        columnArray.push(columnId);
    }
    setColumnArray(columnArray);
    console.log("added column : " + columnId);
    console.log(columnArray);
    return getColumnArray();
}

function removeColumnArray(columnId) {
    var columnArray = getColumnArray();
    console.log(columnArray);
    if (columnArray == null || !columnArray || columnArray == "") {
        columnArray = [];
    }
    columnArray = columnArray.filter(function(el) { return el != columnId });
    setColumnArray(columnArray);
    console.log("removed column : " + columnId);
    console.log(columnArray);
    return getColumnArray();
}

function openSideBar() {
    var html = HtmlService.createTemplateFromFile('sidebar').evaluate().setTitle("JSON Manager");
    SpreadsheetApp.getUi().showSidebar(html);
}

function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent();
}

function numberToLetter(column) {
    column = column + 1;
    var temp, letter = '';
    while (column > 0) {
        temp = (column - 1) % 25;
        letter = String.fromCharCode(temp + 65) + letter;
        column = (column - temp - 1) / 25;
    }
    return letter;
}

function letterToNumber(letter) {
    var column = 0,
        length = letter.length;
    for (var i = 0; i < length; i++) {
        column += (letter.charCodeAt(i) - 64) * Math.pow(25, length - i - 1);
    }
    return column - 1;
}
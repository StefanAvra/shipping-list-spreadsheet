function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Versand')
        .addItem('Retoureschein', 'showReturnSlip')
        .addItem('Format wiederherstellen', 'reformat')
        .addSeparator()
        .addItem('Abschließen', 'finalize')
        .addToUi();
}


function onEdit(e) {
    // remove hint if new tracking id added
    var row = e.range.getRow();
    var col = e.range.getColumn();
    if (row == 4 && col == 2) {
        var shippingList = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
        shippingList.getRange("B4").setBackground("white")
    }
}

function showReturnSlip() {
    // will open a return slip that is supposed to be added to the package for each item
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
    var helper = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1];
    var template = HtmlService.createTemplateFromFile('template.html');
    template.customer = ''
    template.orderID = ''
    template.ticketID = ''
    template.model = ''
    template.reason = ''
    template.user = ''
    template.store = ''
    template.other = ''
    template.date = Utilities.formatDate(new Date(), "GMT+1", "dd.MM.yyyy")
    var activeRow = activeSheet.getActiveCell().getRow();
    if (activeRow >= 6 && activeRow <= helper.getRange("G2").getValue()) { // size of the list should be dynamic :/
        currentOrder = activeSheet.getRange("A" + activeRow + ":F" + activeRow).getValues()
        template.customer = currentOrder[0][0]
        template.orderID = currentOrder[0][1]
        template.ticketID = currentOrder[0][2]
        template.model = currentOrder[0][4]
        template.reason = currentOrder[0][3]
        template.user = currentOrder[0][5]
        template.store = helper.getRange("D1").getValue()
    } else {
        template.other = 'Zeile außerhalb der Liste markiert'
    }
    var html = HtmlService.createHtmlOutput(template.evaluate().getContent())
        .setWidth(800)
        .setHeight(640);
    SpreadsheetApp.getUi()
        .showModalDialog(html, 'Reklaformular');
}

function finalize() {
    var today = new Date()
    // set sender name + date
    var shippingList = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var helper = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1];
    var helperRange = helper.getDataRange()
    var userMail = Session.getEffectiveUser().getEmail()
    var name = helperRange.createTextFinder(userMail).matchEntireCell(true).findNext().offset(0, 1).getValue()
    shippingList.getRange("F3").setValue(name) // unnecessary since user prints document before setting these values
    shippingList.getRange("D3").setValue("=TODAY()")


    // save copy
    var trackID = shippingList.getRange("B4").getValue();
    var lastRow = shippingList.getDataRange().getLastRow();
    var amountRows = lastRow - 5;
    var valueRange = shippingList.getRange("A6:F" + lastRow);
    var sentList = SpreadsheetApp.openById(helper.getRange("F2").getValue()).getSheets()[0];
    var sentListLastRow = sentList.getLastRow();
    sentList.insertRowAfter(sentListLastRow);
    var sentListRange = sentList.getRange("C" + (sentListLastRow + 1) + ":H" + (sentListLastRow + amountRows))

    var values = valueRange.getValues();

    sentListRange.setValues(values)
    for (var i = amountRows; i > 0; i--) {
        sentList.getRange("A" + (sentListLastRow + i)).setValue(today);
        sentList.getRange("B" + (sentListLastRow + i)).setValue('=HYPERLINK("https://www.dhl.de/de/privatkunden/pakete-empfangen/verfolgen.html?piececode=' + trackID.toString() + '";"' + trackID + '")');
    }

    // clear form
    valueRange.setValue('')
    shippingList.getRange("F3").setValue("")
    shippingList.getRange("B3").setValue(today)
    shippingList.getRange("B4").setBackground("orange")
    reformat()

}


function reformat() {
    var shippingList = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var templateRange = SpreadsheetApp.getActiveSpreadsheet().getSheets()[2].getRange("A1:F26");
    shippingList.clearFormats()
    templateRange.copyFormatToRange(shippingList, 1, 6, 1, 26)
}
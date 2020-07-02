function onOpen() {
    SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
        .createMenu('Custom Menu')
        .addItem('Show dialog', 'showDialog')
        .addToUi();
  }
  
  function showDialog() {
    var currentOrder = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getCurrentCell().getValue();
    var template = HtmlService.createTemplateFromFile('template.html');
    template.customer = ''
    template.orderID = ''
    template.ticketID = ''
    template.model = ''
    template.reason = ''
    template.user = ''
    template.store = ''
    template.date = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy")
    var html = HtmlService.createHtmlOutput(template.evaluate().getContent())
        .setWidth(800)
        .setHeight(640);
    SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
        .showModalDialog(html, 'Reklaformular');
  }
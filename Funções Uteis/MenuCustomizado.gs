function bomDia() {
  var ui = SpreadsheetApp.getUi();
  ui.alert("Bom dia!");
}

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Ações")
    .addItem("Bom dia", "bomDia")
    .addToUi();
} 

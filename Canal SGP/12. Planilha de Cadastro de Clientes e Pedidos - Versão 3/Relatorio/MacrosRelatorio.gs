function FormRelatorio() {

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaProduto = planilha.getSheetByName("Produtos");
  var guiaCliente = planilha.getSheetByName("Clientes");
  var guiaVendedor = planilha.getSheetByName("Vendedores");
  var guiaEstado = planilha.getSheetByName("Estados");

  var ultimaLinha = guiaProduto.getLastRow();

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var dadosLinhas = guiaProduto.getRange(2, 1, ultimaLinha, 1).getValues();

  var listaUnica = [...new Set(dadosLinhas.flat())];

  var listaLinhas = [];

  for(var i = 0; i < listaUnica.length; i++){
    listaLinhas.push([listaUnica[i]]);
  }

  var list1 = listaLinhas.sort();

  var ultimaLinha = guiaCliente.getLastRow();

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var list2 = guiaCliente.getRange(2, 2, ultimaLinha, 1).getValues();

  list2.sort();

  var ultimaLinha = guiaVendedor.getLastRow();

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var list3 = guiaVendedor.getRange(2, 1, ultimaLinha, 1).getValues();

  list3.sort();

  var ultimaLinha = guiaEstado.getLastRow();

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var list4 = guiaEstado.getRange(2, 1, ultimaLinha, 1).getValues();

  list4.sort();
  
  var Form = HtmlService.createTemplateFromFile("Relatorio");

  Form.list1 = list1.map(function(r){
    return r[0];
  })

  Form.list2 = list2.map(function(r){
    return r[0];
  })

  Form.list3 = list3.map(function(r){
    return r[0];
  })
  
  Form.list4 = list4.map(function(r){
    return r[0];
  })

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

  MostrarForm.setTitle("RELATÓRIO PRODUTOS").setHeight(510).setWidth(1200);

  SpreadsheetApp.getUi().showModalDialog(MostrarForm, "RELATÓRIO PRODUTOS");
  
}

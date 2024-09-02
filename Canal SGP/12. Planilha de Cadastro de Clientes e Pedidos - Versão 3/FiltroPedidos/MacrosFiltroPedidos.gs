function FormFiltroPedidos() {

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

  dadosLinhas.length = 0;
  listaUnica.length = 0;

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

  var list3 = guiaVendedor.getRange(2,1,ultimaLinha,1).getValues();

  list3.sort();

  var ultimaLinha = guiaEstado.getLastRow();

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var list4 = guiaEstado.getRange(2,1,ultimaLinha,1).getValues();

  list4.sort();

  var Form = HtmlService.createTemplateFromFile("FiltroPedidos");

  Form.list1 = list1.map(function(r){
    return r[0];
  });

  Form.list2 = list2.map(function(r){
    return r[0];
  });

  Form.list3 = list3.map(function(r){
    return r[0];
  });

  Form.list4 = list4.map(function(r){ 
    return r[0];
  });

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

  MostrarForm.setTitle("FILTRO PEDIDOS").setHeight(600).setWidth(1100);

  SpreadsheetApp.getUi().showModalDialog(MostrarForm, "FILTRO PEDIDOS");
  
}


function buscaDados(){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaCidades = planilha.getSheetByName("Estados/Cidades");
  var guiaProdutos = planilha.getSheetByName("Produtos");
  var guiaPedido = planilha.getSheetByName("Pedidos");

  var ultimaLinha = guiaCidades.getLastRow();
  var dadosCidades = guiaCidades.getRange(2, 1, ultimaLinha, 2).getValues();

  var ultimaLinha = guiaProdutos.getLastRow();
  var dadosProdutos = guiaProdutos.getRange(2, 1, ultimaLinha, 3).getValues();

  var ultimaLinha = guiaPedido.getLastRow();
  var dadosPedidos = guiaPedido.getRange(2, 1, ultimaLinha, 16).getValues();

  for(var i = 0; i < dadosPedidos.length; i++){

    var Data = Utilities.formatDate(new Date(dadosPedidos[i][2]), planilha.getSpreadsheetTimeZone(),"dd/MM/yyyy");

    dadosPedidos[i][2] = Data;

  }

  var arrays = {
    dadosCidades: dadosCidades,
    dadosProdutos: dadosProdutos,
    dadosPedidos: dadosPedidos,
  }

  return arrays;

}

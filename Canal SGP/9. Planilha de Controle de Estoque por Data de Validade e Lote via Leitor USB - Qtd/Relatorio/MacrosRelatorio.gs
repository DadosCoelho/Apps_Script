function FormRelatorio() {

  var ultimaLinha = guiaProduto.getLastRow();

  if (ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var dadosProdutos = guiaProduto.getRange(2,2, ultimaLinha,1).getValues();

  var listaUnica = [...new Set(dadosProdutos.flat())];

  var listaProdutos = [];
    
  for (var i = 0; i < listaUnica.length; i++) {
    listaProdutos.push([listaUnica[i]]);
  }

  var list = listaProdutos.sort();

  dadosProdutos.length = 0;
  listaUnica.length = 0;
  
  var Form = HtmlService.createTemplateFromFile("FormRelatorio");

  Form.list = list.map(function(r){ 
    return r[0];
  }); 

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

  MostrarForm.setTitle("RELATÓRIO").setHeight(600).setWidth(1100);

  SpreadsheetApp.getUi().showModalDialog(MostrarForm,"RELATÓRIO");
  
}

function buscaDados(){  

  var ultimaLinha = guiaProduto.getLastRow() - 1;

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var dadosProdutos = guiaProduto.getRange(2, 1, ultimaLinha, 2).getValues();

  var ultimaLinha = guiaEntrada.getLastRow() - 1;

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var dadosEntrada = guiaEntrada.getRange(2, 1, ultimaLinha, 7).getValues();

  for(var i = 0; i < dadosEntrada.length; i++){

    var DataValidade = Utilities.formatDate(new Date(dadosEntrada[i][5]), planilha.getSpreadsheetTimeZone(),"dd/MM/yyyy");

    dadosEntrada[i][5] = DataValidade;

    var Data = Utilities.formatDate(new Date(dadosEntrada[i][6]), planilha.getSpreadsheetTimeZone(),"dd/MM/yyyy");

    dadosEntrada[i][6] = Data;

  }

  var ultimaLinha = guiaSaida.getLastRow() - 1;

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }
  
  var dadosSaida = guiaSaida.getRange(2, 1, ultimaLinha, 7).getValues();  

  for(var i = 0; i < dadosSaida.length; i++){

    var DataValidade = Utilities.formatDate(new Date(dadosSaida[i][5]), planilha.getSpreadsheetTimeZone(),"dd/MM/yyyy");

    dadosSaida[i][5] = DataValidade;

    var Data = Utilities.formatDate(new Date(dadosSaida[i][6]), planilha.getSpreadsheetTimeZone(),"dd/MM/yyyy");

    dadosSaida[i][6] = Data;

  }

  var arrays = {    
    dadosProdutos: dadosProdutos,
    dadosEntrada: dadosEntrada,
    dadosSaida: dadosSaida,
  }

  return arrays;

}

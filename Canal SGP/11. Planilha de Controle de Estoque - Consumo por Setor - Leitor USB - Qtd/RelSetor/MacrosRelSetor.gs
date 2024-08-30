function FormRelSetor() {

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
  
  var ultimaLinha = guiaSetor.getLastRow();

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var dadosSetor = guiaSetor.getRange(2, 1, ultimaLinha, 1).getValues();

  var listaUnica = [...new Set(dadosSetor.flat())];

  var listaSetores = [];

  for(var i = 0; i < listaUnica.length; i++){
    listaSetores.push([listaUnica[i]]);
  }

  dadosSetor.length = 0;
  listaUnica.length = 0;

  var list2 = listaSetores.sort();

  var Form = HtmlService.createTemplateFromFile("FormRelSetor");

  Form.list = list.map(function(r){ 
    return r[0];
  }); 

  Form.list2 = list2.map(function(r){ 
    return r[0];
  }); 

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

  MostrarForm.setTitle("RELATÓRIO CONSUMO SETORES").setHeight(600).setWidth(1100);

  SpreadsheetApp.getUi().showModalDialog(MostrarForm,"RELATÓRIO CONSUMO SETORES");
  
}

function buscaDadosRelSetores(){  

  var ultimaLinha = guiaProduto.getLastRow() - 1;

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var dadosProdutos = guiaProduto.getRange(2, 1, ultimaLinha, 2).getValues();

  var ultimaLinha = guiaSetor.getLastRow() - 1;

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var dadosSetor = guiaSetor.getRange(2, 1, ultimaLinha, 1).getValues();

  var ultimaLinha = guiaSaida.getLastRow() - 1;

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }
  
  var dadosSaida = guiaSaida.getRange(2, 1, ultimaLinha, 6).getValues();  

  for(var i = 0; i < dadosSaida.length; i++){   

    var Data = Utilities.formatDate(new Date(dadosSaida[i][5]), planilha.getSpreadsheetTimeZone(),"dd/MM/yyyy");

    dadosSaida[i][5] = Data;

  }

  var arrays = {    
    dadosProdutos: dadosProdutos,
    dadosSetor:dadosSetor,   
    dadosSaida: dadosSaida,
  }

  return arrays;

}

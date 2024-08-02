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

function buscaDadosRel(){  

  var ultimaLinha = guiaProduto.getLastRow() - 1;

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var dadosProdutos = guiaProduto.getRange(2, 1, ultimaLinha, 4).getValues();

  var ultimaLinha = guiaContagem.getLastRow() - 1;

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var dadosContagem = guiaContagem.getRange(2, 1, ultimaLinha, 6).getValues();

  for(var i = 0; i < dadosContagem.length; i++){

    var Data = Utilities.formatDate(new Date(dadosContagem[i][5]), planilha.getSpreadsheetTimeZone(),"dd/MM/yyyy");

    dadosContagem[i][5] = Data;

  }  

  var arrays = {    
    dadosProdutos: dadosProdutos,
    dadosContagem: dadosContagem,    
  }

  return arrays;

}

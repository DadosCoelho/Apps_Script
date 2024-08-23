function FormSaida() {

  var ultimaLinha = guiaLocalizacao.getLastRow();

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var dadosLocalizacao = guiaLocalizacao.getRange(2, 1, ultimaLinha, 1).getValues();

  var listaUnica = [...new Set(dadosLocalizacao.flat())];

  var listaLocalizacao = [];

  for(var i = 0; i < listaUnica.length; i++){
    listaLocalizacao.push([listaUnica[i]]);
  }

  dadosLocalizacao.length = 0;
  listaUnica.length = 0;

  var list = listaLocalizacao.sort();
  
  var Form = HtmlService.createTemplateFromFile("FormSaida");

  Form.list = list.map(function(r){
    return r[0];
  });  

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

  MostrarForm.setTitle("SAÍDA").setHeight(600).setWidth(1100);

  SpreadsheetApp.getUi().showModalDialog(MostrarForm,"SAÍDA");
  
}

function dadosBaixaSaida(){
  
  var ultimaLinha = guiaProduto.getLastRow() - 1;

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var dadosProdutos = guiaProduto.getRange(2, 1, ultimaLinha, 2).getValues();

  var ultimaLinha = guiaEntrada.getLastRow() - 1;

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }
  
  var dadosEntrada = guiaEntrada.getRange(2, 1, ultimaLinha, 5).getValues();  

  var ultimaLinha = guiaSaida.getLastRow() - 1;

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var dadosSaida = guiaSaida.getRange(2, 1, ultimaLinha, 5).getValues();

  var arrays = {
    dadosProdutos: dadosProdutos,    
    dadosEntrada: dadosEntrada,
    dadosSaida: dadosSaida,
  }

  return arrays;

}


function SalvarSaida(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

    var linha = guiaSaida.getLastRow() + 1;

    var maiorId = Math.max.apply(null, guiaSaida.getRange("A2:A").getValues());
    var novoId = maiorId + 1;

    for(var i = 0; i < Dados.length; i++){

      guiaSaida.getRange(linha,1).setValue(novoId);
      guiaSaida.getRange(linha,2).setValue(Dados[i][0]);
      guiaSaida.getRange(linha,3).setValue(Dados[i][1]);
      guiaSaida.getRange(linha,4).setValue(Dados[i][2]);
      guiaSaida.getRange(linha,5).setValue(Dados[i][3]);
      guiaSaida.getRange(linha,6).setValue(Dados[i][4]);
      guiaSaida.getRange(linha,7).setValue(Dados[i][5]);

      linha = linha + 1;
      novoId = novoId + 1;

    }

    return "REGISTRADO COM SUCESSO!";

  }

}

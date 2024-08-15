function FormEntrada() {
  
  var Form = HtmlService.createTemplateFromFile("FormEntrada");  

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

  MostrarForm.setTitle("ENTRADA").setHeight(600).setWidth(1100);

  SpreadsheetApp.getUi().showModalDialog(MostrarForm,"ENTRADA");
  
}


function buscaProdutos(){

  var ultimaLinha = guiaProduto.getLastRow() - 1;

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var dadosProdutos = guiaProduto.getRange(2, 1, ultimaLinha, 2).getValues();
  
  var arrays = {
    dadosProdutos: dadosProdutos,   
  }
  
  return arrays;

}


function SalvarRegistros(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

    var linha = guiaEntrada.getLastRow() + 1;

    var maiorId = Math.max.apply(null, guiaEntrada.getRange("A2:A").getValues());
    var novoId = maiorId + 1;

    for(var i = 0; i < Dados.length; i++){

      guiaEntrada.getRange(linha,1).setValue(novoId);
      guiaEntrada.getRange(linha,2).setValue(Dados[i][0]);
      guiaEntrada.getRange(linha,3).setValue(Dados[i][1]);
      guiaEntrada.getRange(linha,4).setValue(Dados[i][2]);
      guiaEntrada.getRange(linha,5).setValue(Dados[i][3]);
      guiaEntrada.getRange(linha,6).setValue(Dados[i][4]);
      guiaEntrada.getRange(linha,7).setValue(Dados[i][5]);

      linha = linha + 1;
      novoId = novoId + 1;

    }

    return "REGISTRADO COM SUCESSO!";

  }

}

function FormInventario() {
  
  var Form = HtmlService.createTemplateFromFile("FormInventario");  

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

  MostrarForm.setTitle("INVENTÁRIO").setHeight(600).setWidth(1100);

  SpreadsheetApp.getUi().showModalDialog(MostrarForm,"INVENTÁRIO");
  
}


function buscaProdutos(){

  var ultimaLinha = guiaProduto.getLastRow() - 1;

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var dadosProdutos = guiaProduto.getRange(2, 1, ultimaLinha, 4).getValues();
  
  var arrays = {
    dadosProdutos: dadosProdutos,   
  }
  
  return arrays;

}


function SalvarRegistros(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

    var linha = guiaContagem.getLastRow() + 1;

    var maiorId = Math.max.apply(null, guiaContagem.getRange("A2:A").getValues());
    var novoId = maiorId + 1;

    for(var i = 0; i < Dados.length; i++){

      guiaContagem.getRange(linha,1).setValue(novoId);
      guiaContagem.getRange(linha,2).setValue(Dados[i][0]);
      guiaContagem.getRange(linha,3).setValue(Dados[i][1]);
      guiaContagem.getRange(linha,4).setValue(Dados[i][2]);
      guiaContagem.getRange(linha,5).setValue(Dados[i][3]);
      guiaContagem.getRange(linha,6).setValue(Dados[i][4]);
      guiaContagem.getRange(linha,7).setValue(Dados[i][5]);

      linha = linha + 1;
      novoId = novoId + 1;

    }

    return "REGISTRADO COM SUCESSO!";

  }

}

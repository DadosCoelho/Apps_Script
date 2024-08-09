function FormSaida() {
  
    var Form = HtmlService.createTemplateFromFile("FormSaida");  
  
    var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
  
    MostrarForm.setTitle("SAÍDA").setHeight(600).setWidth(1100);
  
    SpreadsheetApp.getUi().showModalDialog(MostrarForm,"SAÍDA");
    
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
        guiaSaida.getRange(linha,8).setValue(Dados[i][6]);
  
        linha = linha + 1;
        novoId = novoId + 1;
  
      }
  
      return "REGISTRADO COM SUCESSO!";
  
    }
  
  }
  
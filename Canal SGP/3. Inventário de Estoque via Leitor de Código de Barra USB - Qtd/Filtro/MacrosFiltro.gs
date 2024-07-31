function FormFiltro() {

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
  
  var Form = HtmlService.createTemplateFromFile("FormFiltro");

  Form.list = list.map(function(r){ 
    return r[0];
  }); 

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

  MostrarForm.setTitle("FILTRO").setHeight(600).setWidth(1100);

  SpreadsheetApp.getUi().showModalDialog(MostrarForm,"FILTRO");
  
}

function dadosFiltro(){
  
  var ultimaLinha = guiaContagem.getLastRow() - 1;

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var dadosContagem = guiaContagem.getRange(2, 1, ultimaLinha, 6).getValues();

  for(var i = 0; i < dadosContagem.length; i++){

    var Data = Utilities.formatDate(new Date(dadosContagem[i][4]), planilha.getSpreadsheetTimeZone(),"dd/MM/yyyy");

    dadosContagem[i][4] = Data;

    if(dadosContagem[i][5] != ""){
      var Hora = dadosContagem[i][5].getHours().toString().padStart(2, '0');
      var Minutos = dadosContagem[i][5].getMinutes().toString().padStart(2, '0');
      var Segundos = dadosContagem[i][5].getSeconds().toString().padStart(2, '0');
      var Hora =  Hora + ":" + Minutos + ":" + Segundos;
      dadosContagem[i][5] = Hora;
    }

  }

  var arrays = {    
    dadosContagem: dadosContagem,
  }

  return arrays;

}

function ExcluirLinha(Id){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){    

    var ultimaLinha = guiaContagem.getLastRow(); 
    var dadosContagem = guiaContagem.getRange(2, 1,ultimaLinha,1).getValues();   

    for(var i = 0; i < dadosContagem.length; i++){

      if(dadosContagem[i][0] == Id){
        var linha = i + 2;
        guiaContagem.deleteRow(linha);
        dadosContagem.length = 0;        
        return "EXCLUÍDO COM SUCESSO!";
      }

    }

    dadosContagem.length = 0;     
    return "ID NÃO ENCONTRADO!";

  }

}

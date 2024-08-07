function FormFiltroEntrada() {

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
  
  var Form = HtmlService.createTemplateFromFile("FiltroEntrada");

  Form.list = list.map(function(r){ 
    return r[0];
  }); 

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

  MostrarForm.setTitle("FILTRO ENTRADA").setHeight(600).setWidth(1100);

  SpreadsheetApp.getUi().showModalDialog(MostrarForm,"FILTRO ENTRADA");
  
}

function dadosFiltroEntrada(){
  
  var ultimaLinha = guiaEntrada.getLastRow() - 1;

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }
  
  var dadosEntrada = guiaEntrada.getRange(2, 1, ultimaLinha, 6).getValues();

  for(var i = 0; i < dadosEntrada.length; i++){

    var Data = Utilities.formatDate(new Date(dadosEntrada[i][4]), planilha.getSpreadsheetTimeZone(),"dd/MM/yyyy");

    dadosEntrada[i][4] = Data;

    if(dadosEntrada[i][5] != ""){
      var Hora = dadosEntrada[i][5].getHours().toString().padStart(2, '0');
      var Minutos = dadosEntrada[i][5].getMinutes().toString().padStart(2, '0');
      var Segundos = dadosEntrada[i][5].getSeconds().toString().padStart(2, '0');
      var Hora =  Hora + ":" + Minutos + ":" + Segundos;
      dadosEntrada[i][5] = Hora;
    }

  }

  var arrays = {    
    dadosEntrada: dadosEntrada,
  }

  return arrays;

}

function ExcluirEntrada(Id){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){    

    var ultimaLinha = guiaEntrada.getLastRow();
    var dadosEntrada = guiaEntrada.getRange(2, 1,ultimaLinha,1).getValues();   

    for(var i = 0; i < dadosEntrada.length; i++){

      if(dadosEntrada[i][0] == Id){
        var linha = i + 2;
        guiaEntrada.deleteRow(linha);
        dadosEntrada.length = 0;        
        return "EXCLUÍDO COM SUCESSO!";
      }

    }

    dadosEntrada.length = 0;     
    return "ID NÃO ENCONTRADO!";

  }

}


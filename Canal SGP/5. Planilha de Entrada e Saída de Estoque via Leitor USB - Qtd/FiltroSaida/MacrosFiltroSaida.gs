function FormFiltroSaida() {

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
  
  var Form = HtmlService.createTemplateFromFile("FiltroSaida");

  Form.list = list.map(function(r){ 
    return r[0];
  }); 

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

  MostrarForm.setTitle("FILTRO SAÍDA").setHeight(600).setWidth(1100);

  SpreadsheetApp.getUi().showModalDialog(MostrarForm,"FILTRO SAÍDA");
  
}

function dadosFiltroSaida(){
  
  var ultimaLinha = guiaSaida.getLastRow() - 1;

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }
  
  var dadosSaida = guiaSaida.getRange(2, 1, ultimaLinha, 6).getValues();  

  for(var i = 0; i < dadosSaida.length; i++){    

    var Data = Utilities.formatDate(new Date(dadosSaida[i][4]), planilha.getSpreadsheetTimeZone(),"dd/MM/yyyy");    

    if(dadosSaida[i][5] != ""){
      var Hora = dadosSaida[i][5].getHours().toString().padStart(2, '0');
      var Minutos = dadosSaida[i][5].getMinutes().toString().padStart(2, '0');
      var Segundos = dadosSaida[i][5].getSeconds().toString().padStart(2, '0');
      var Hora =  Hora + ":" + Minutos + ":" + Segundos;
      dadosSaida[i][5] = Hora;
    }

    dadosSaida[i][4] = Data;    

  }

  var arrays = {    
    dadosSaida: dadosSaida,
  }  

  return arrays;

}

function ExcluirSaida(Id){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){   

    var ultimaLinha = guiaSaida.getLastRow();
    var dadosSaida = guiaSaida.getRange(2, 1,ultimaLinha,1).getValues();   

    for(var i = 0; i < dadosSaida.length; i++){

      if(dadosSaida[i][0] == Id){
        var linha = i + 2;
        guiaSaida.deleteRow(linha);
        dadosSaida.length = 0;        
        return "EXCLUÍDO COM SUCESSO!";
      }

    }

    dadosSaida.length = 0;     
    return "ID NÃO ENCONTRADO!";

  }

}


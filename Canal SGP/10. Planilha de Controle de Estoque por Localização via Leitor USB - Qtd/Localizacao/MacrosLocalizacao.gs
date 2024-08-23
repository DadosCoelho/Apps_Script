function FormLocalizacao() {  

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
  
  var Form = HtmlService.createTemplateFromFile("FormLocalizacao");

  Form.list = list.map(function(r){
    return r[0];
  });

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

  MostrarForm.setTitle("CADASTRO DE LOCALIZAÇÕES").setHeight(200).setWidth(385);

  SpreadsheetApp.getUi().showModalDialog(MostrarForm,"CADASTRO DE LOCALIZAÇÕES");

}

function SalvarLocalizacao(Localizacao){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){    

    var ultimaLinha = guiaLocalizacao.getLastRow();

    var dadosLocalizacao = guiaLocalizacao.getRange(2,1,ultimaLinha,1).getValues();

    for(var i = 0; i < dadosLocalizacao.length; i++){
      if(dadosLocalizacao[i][0] == Localizacao){
        dadosLocalizacao.length = 0;
        return "LOCALIZAÇÃO JÁ CADASTRADA!";
      }
    }

    var linha = ultimaLinha + 1;

    guiaLocalizacao.getRange(linha,1).setValue(Localizacao);
    guiaLocalizacao.getRange("A2:A").sort([{column: 1, ascending: true}]);

    dadosLocalizacao.length = 0;
    return "REGISTRADO COM SUCESSO!";

  }

}


function EditarLocalizacao(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){
      
    var Lista = Dados.Lista;    
    var Localizacao = Dados.Localizacao;

    var ultimaLinha = guiaEntrada.getLastRow();

    var dadosEntrada = guiaEntrada.getRange(2, 5, ultimaLinha, 1).getValues();

    var ver = dadosEntrada.filter(function(value, i, arr){
      return Localizacao == arr[i][0];
    });

    if(ver.length > 0){
      dadosEntrada.length = 0;
      return "NÃO PODE SER EDITADO. JÁ TEM LANÇAMENTO DE ENTRADA!";
    }

    var ultimaLinha = guiaLocalizacao.getLastRow();

    var dadosLocalizacao = guiaLocalizacao.getRange(2, 1, ultimaLinha, 1).getValues();

    for(var i = 0; i < dadosLocalizacao.length; i++){
      if(dadosLocalizacao[i][0] == Lista){
        var linha = i + 2;        
        guiaLocalizacao.getRange(linha,1).setValue(Localizacao);
        guiaLocalizacao.getRange("A2:A").sort([{column: 1, ascending: true}]);       
        dadosLocalizacao.length = 0;
        dadosEntrada.length = 0; 
        return "EDITADO COM SUCESSO!";
      }
    }
    
    dadosEntrada.length = 0;
    dadosLocalizacao.length = 0;   
    return "PRODUTO NÃO ENCONTRADO!";

  }

}


function ExcluirLocalizacao(Localizacao){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){ 

    var ultimaLinha = guiaEntrada.getLastRow();

    var dadosEntrada = guiaEntrada.getRange(2, 5, ultimaLinha, 1).getValues();

    var ver = dadosEntrada.filter(function(value, i, arr){
      return Localizacao == arr[i][0];
    });

    if(ver.length > 0){
      dadosEntrada.length = 0;     
      return "NÃO PODE SER EXCLUÍDO. JÁ TEM LANÇAMENTO DE ENTRADA!";
    }

    var ultimaLinha = guiaLocalizacao.getLastRow();

    var dadosLocalizacao = guiaLocalizacao.getRange(2,1,ultimaLinha,1).getValues();

    for(var i = 0; i < dadosLocalizacao.length; i++){
      if(dadosLocalizacao[i][0] == Localizacao){
        var linha = i + 2;
        guiaLocalizacao.deleteRow(linha);
        dadosEntrada.length = 0;
        dadosLocalizacao.length = 0;
        ver.length = 0;
        return "EXCLUÍDO COM SUCESSO!";
      }
    }

    dadosEntrada.length = 0;
    dadosLocalizacao.length = 0;
    ver.length = 0;

    return "PRODUTO NÃO ENCONTRADO!";

  }

}


function AtualizarLocalizacoes(){  

  var ultimaLinha = guiaLocalizacao.getLastRow();

  if(ultimaLinha == 0){
    var ultimaLinha = 1;
  }

  var dadosLocalizacao = guiaLocalizacao.getRange(2, 1, ultimaLinha, 1).getValues();

  var listaUnica = [...new Set(dadosLocalizacao.flat())];

  var listaLocalizacao = [];

  for(var i = 0; i < listaUnica.length; i++){
    listaLocalizacao.push([listaUnica[i]]);
  }

  dadosLocalizacao.length = 0;
  listaUnica.length = 0;

  var lista = listaLocalizacao.sort();

  return lista;

}

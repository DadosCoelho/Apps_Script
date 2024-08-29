function FormSetor() {  

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

  var list = listaSetores.sort();
  
  var Form = HtmlService.createTemplateFromFile("FormSetor");

  Form.list = list.map(function(r){
    return r[0];
  });

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

  MostrarForm.setTitle("CADASTRO DE SETORES").setHeight(200).setWidth(385);

  SpreadsheetApp.getUi().showModalDialog(MostrarForm,"CADASTRO DE SETORES");

}

function SalvarSetor(Setor){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){    

    var ultimaLinha = guiaSetor.getLastRow();

    var dadosSetor = guiaSetor.getRange(2,1,ultimaLinha,1).getValues();

    for(var i = 0; i < dadosSetor.length; i++){
      if(dadosSetor[i][0] == Setor){
        dadosSetor.length = 0;
        return "SETOR JÁ CADASTRADO!";
      }
    }

    var linha = ultimaLinha + 1;

    guiaSetor.getRange(linha,1).setValue(Setor);
    guiaSetor.getRange("A2:A").sort([{column: 1, ascending: true}]);

    dadosSetor.length = 0;
    return "REGISTRADO COM SUCESSO!";

  }

}


function EditarSetor(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){
      
    var Lista = Dados.Lista;    
    var Setor = Dados.Setor;

    var ultimaLinha = guiaSaida.getLastRow();

    var dadosSaida = guiaSaida.getRange(2, 5, ultimaLinha, 1).getValues();

    var ver = dadosSaida.filter(function(value, i, arr){
      return Setor == arr[i][0];
    });

    if(ver.length > 0){
      dadosSaida.length = 0;     
      return "NÃO PODE SER EDITADO. JÁ TEM LANÇAMENTO!";
    }

    var ultimaLinha = guiaSetor.getLastRow();

    var dadosSetor = guiaSetor.getRange(2, 1, ultimaLinha, 1).getValues();

    for(var i = 0; i < dadosSetor.length; i++){
      if(dadosSetor[i][0] == Lista){
        var linha = i + 2;        
        guiaSetor.getRange(linha,1).setValue(Setor);
        guiaSetor.getRange("A2:A").sort([{column: 1, ascending: true}]);       
        dadosSetor.length = 0;
        dadosSaida.length = 0;
        return "EDITADO COM SUCESSO!";
      }
    }
    
    dadosSetor.length = 0;
    dadosSaida.length = 0;
      
    return "PRODUTO NÃO ENCONTRADO!";

  }

}


function ExcluirSetor(Setor){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){ 

    var ultimaLinha = guiaSaida.getLastRow();

    var dadosSaida = guiaSaida.getRange(2, 5, ultimaLinha, 1).getValues();

    var ver = dadosSaida.filter(function(value, i, arr){
      return Setor == arr[i][0];
    });

    if(ver.length > 0){
      dadosSaida.length = 0;     
      return "NÃO PODE SER EXCLUÍDO. JÁ TEM LANÇAMENTO!";
    }

    var ultimaLinha = guiaSetor.getLastRow();

    var dadosSetor = guiaSetor.getRange(2,1,ultimaLinha,1).getValues();

    for(var i = 0; i < dadosSetor.length; i++){
      if(dadosSetor[i][0] == Setor){
        var linha = i + 2;
        guiaSetor.deleteRow(linha);
        dadosSaida.length = 0;
        dadosSetor.length = 0;
        ver.length = 0;
        return "EXCLUÍDO COM SUCESSO!";
      }
    }

    dadosSaida.length = 0;
    dadosSetor.length = 0;
    ver.length = 0;

    return "PRODUTO NÃO ENCONTRADO!";

  }

}


function AtualizarSetores(){  

  var ultimaLinha = guiaSetor.getLastRow();

  if(ultimaLinha == 0){
    var ultimaLinha = 1;
  }

  var dadosSetor = guiaSetor.getRange(2, 1, ultimaLinha, 1).getValues();

  var listaUnica = [...new Set(dadosSetor.flat())];

  var listaSetores = [];

  for(var i = 0; i < listaUnica.length; i++){
    listaSetores.push([listaUnica[i]]);
  }

  dadosSetor.length = 0;
  listaUnica.length = 0;

  var lista = listaSetores.sort();

  return lista;

}

function FormEntrada(Id) {

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaProduto = planilha.getSheetByName("Produtos");

  var ultimaLinha = guiaProduto.getLastRow();

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var dadosLinha = guiaProduto.getRange(2, 1, ultimaLinha, 1).getValues();

  var listaUnica = [...new Set(dadosLinha.flat())];

  var listaLinhas = [];

  for(var i = 0; i < listaUnica.length; i++){
    listaLinhas.push([listaUnica[i]]);
  }

  dadosLinha.length = 0;
  listaUnica.length = 0;

  var list = listaLinhas.sort();
   
  var Form = HtmlService.createTemplateFromFile("FormEntrada");

  Form.list = list.map(function(r){
    return r[0];
  });

  Form.Id = Id;

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

  MostrarForm.setTitle("ENTRADA ESTOQUE").setHeight(335).setWidth(650);

  SpreadsheetApp.getUi().showModalDialog(MostrarForm, "ENTRADA ESTOQUE");
  
}

function buscaProdutos(){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaProduto = planilha.getSheetByName("Produtos");

  var ultimaLinha = guiaProduto.getLastRow();

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var dadosProdutos = guiaProduto.getRange(2, 1,ultimaLinha,3).getValues();

  return dadosProdutos;
  
}

function SalvarEntrada(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaEntrada = planilha.getSheetByName("Entradas");

    var maiorId = Math.max.apply(null, guiaEntrada.getRange("A2:A").getValues());
    var novoId = maiorId + 1;

    var dataQuebrada = Dados.Data.split("/");

    var Ano = dataQuebrada[0];
    var Mes = dataQuebrada[1];
    var Dia = dataQuebrada[2];

    var Data = Dia + "/" + Mes + "/" + Ano;

    var linha = guiaEntrada.getLastRow() + 1;

    guiaEntrada.getRange(linha, 1).setValue(novoId);
    guiaEntrada.getRange(linha, 2).setValue(Data);

    var Data = new Date(Dados.Data);
    var m = Data.getMonth();

    var meses = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];

    var Mes = meses[m];

    guiaEntrada.getRange(linha,3).setValue(Mes);
    guiaEntrada.getRange(linha,4).setValue(Ano);
    guiaEntrada.getRange(linha,5).setValue(Dados.Linha);
    guiaEntrada.getRange(linha,6).setValue(Dados.Marca);
    guiaEntrada.getRange(linha,7).setValue(Dados.Produto);
    guiaEntrada.getRange(linha,8).setValue(Dados.Cod);
    guiaEntrada.getRange(linha,9).setValue(Dados.Nf);
    guiaEntrada.getRange(linha,10).setValue(Dados.Valor);
    guiaEntrada.getRange(linha,11).setValue(Dados.Qtd);
    guiaEntrada.getRange(linha,12).setValue(Dados.Pu);
    guiaEntrada.getRange(linha,13).setValue(Dados.Obs);

    return "SALVO COM SUCESSO!"; 

  }

}


function PesquisarEntrada(id){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaEntrada = planilha.getSheetByName("Entradas");

  var ultimaLinha = guiaEntrada.getLastRow();

  var dados = guiaEntrada.getRange(2,1,ultimaLinha,13).getValues();

  for(var i = 0; i < dados.length; i++){

    if(dados[i][0] == id){

      var data = new Date(dados[i][1]);

      var Dia = data.getDate();
      var Mes = data.getMonth() + 1;
      var Ano = data.getFullYear();

      var Data = Ano + "-" + Mes + "-" + Dia;

      var Linha = dados[i][4];
      var Marca = dados[i][5];
      var Produto = dados[i][6];
      var Cod = dados[i][7];
      var Nf = dados[i][8];

      var V = dados[i][9].toLocaleString({style: 'decimal', decimal: 'pt-BR'});
      var Valor = V.replace(/\./g,"");

      var Q = dados[i][10].toLocaleString({style: 'decimal', decimal: 'pt-BR'});
      var Qtd = Q.replace(/\./g,"");

      var P = dados[i][11].toLocaleString({style: 'decimal', decimal: 'pt-BR'});
      var Pu = P.replace(/\./g,"");

      var Obs = dados[i][12];

      dados.length = 0;

      return ([Data, Linha, Marca, Produto, Cod, Nf, Valor, Qtd, Pu, Obs]);

    }

  }
  
  dados.length = 0; 

  return "NÃO ENCONTRADO!";

}

function EditarEntrada(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaEntrada = planilha.getSheetByName("Entradas");

    var ultimaLinha = guiaEntrada.getLastRow();

    var dados = guiaEntrada.getRange(2, 1, ultimaLinha, 1).getValues();

    for(var i = 0; i < dados.length; i++){

      if(dados[i][0] == Dados.Id){

        var linha = i + 2;

        var dataQuebrada = Dados.Data.split("/");

        var Ano = dataQuebrada[0];
        var Mes = dataQuebrada[1];
        var Dia = dataQuebrada[2];

        var Data = Dia + "/" + Mes + "/" + Ano;

        guiaEntrada.getRange(linha,2).setValue(Data);

        var Data = new Date(Dados.Data);
        var m = Data.getMonth();

        var meses = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];

        var Mes = meses[m];

        guiaEntrada.getRange(linha,3).setValue(Mes);
        guiaEntrada.getRange(linha,4).setValue(Ano);
        guiaEntrada.getRange(linha,5).setValue(Dados.Linha);
        guiaEntrada.getRange(linha,6).setValue(Dados.Marca);
        guiaEntrada.getRange(linha,7).setValue(Dados.Produto);
        guiaEntrada.getRange(linha,8).setValue(Dados.Cod);
        guiaEntrada.getRange(linha,9).setValue(Dados.Nf);
        guiaEntrada.getRange(linha,10).setValue(Dados.Valor);
        guiaEntrada.getRange(linha,11).setValue(Dados.Qtd);
        guiaEntrada.getRange(linha,12).setValue(Dados.Pu);
        guiaEntrada.getRange(linha,13).setValue(Dados.Obs);

        dados.length = 0;
        return "EDITADO COM SUCESSO!";

      }

    }   
    
    dados.length = 0;
    return "ID NÃO ENCONTRADO!";

  }

}

function ExcluirEntrada(id){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaEntrada = planilha.getSheetByName("Entradas");

    var ultimaLinha = guiaEntrada.getLastRow();

    var dados = guiaEntrada.getRange(2, 1, ultimaLinha, 1).getValues();

    for(var i = 0; i < dados.length; i++){

      if(dados[i][0] == id){

        var linha = i + 2;
        guiaEntrada.deleteRow(linha);

        dados.length = 0;
        return "EXCLUÍDO COM SUCESSO!";

      }

    }

    dados.length = 0;
    return "NÃO ENCONTRADO!";

  }  

}
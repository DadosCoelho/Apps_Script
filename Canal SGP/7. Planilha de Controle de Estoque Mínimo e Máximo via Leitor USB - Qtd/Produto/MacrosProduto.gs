function onOpen(){
  SpreadsheetApp.getUi()
 .createMenu('Formulários') 
 .addItem('Produtos', 'FormProduto') 
 .addItem('Entrada', 'FormEntrada')
 .addItem('Saída', 'FormSaida')
 .addItem('Filtro Entrada', 'FormFiltroEntrada') 
 .addItem('Filtro Saída', 'FormFiltroSaida')  
 .addItem('Relatório', 'FormRelatorio')   
 .addToUi();
}

var planilha = SpreadsheetApp.getActiveSpreadsheet();    
var guiaEntrada = planilha.getSheetByName("Entrada");
var guiaProduto = planilha.getSheetByName("Produtos");
var guiaSaida = planilha.getSheetByName("Saida");

function FormProduto() {
  
  var Form = HtmlService.createTemplateFromFile("FormProduto");

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

  MostrarForm.setTitle("CADASTRO DE PRODUTOS").setHeight(245).setWidth(385);

  SpreadsheetApp.getUi().showModalDialog(MostrarForm,"CADASTRO DE PRODUTOS");

}

function Chamar(Arquivo){
  return HtmlService.createHtmlOutputFromFile(Arquivo).getContent();
}

function SalvarProduto(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

    var Cod = Dados.Cod;   
    var Produto = Dados.Produto; 
    var Minimo = Dados.Minimo;
    var Maximo = Dados.Maximo;

    var ultimaLinha = guiaProduto.getLastRow();

    var dadosProdutos = guiaProduto.getRange(2,1,ultimaLinha,3).getValues();

    for(var i = 0; i < dadosProdutos.length; i++){

      if(dadosProdutos[i][0] == Cod){
        dadosProdutos.length = 0;
        return "PRODUTO JÁ CADASTRADO!";
      }

    }

    var linha = ultimaLinha + 1;

    guiaProduto.getRange(linha,1).setValue(Cod);
    guiaProduto.getRange(linha,2).setValue(Produto);
    guiaProduto.getRange(linha,3).setValue(Minimo);
    guiaProduto.getRange(linha,4).setValue(Maximo);
    
    guiaProduto.getRange("A2:D").sort([{column: 2, ascending: true}]);

    dadosProdutos.length = 0;

    ProdutoEntrada(Cod, Produto);
    ProdutoSaida(Cod, Produto);

    return "REGISTRADO COM SUCESSO!";

  }

}


function PesquisarProduto(Cod){  

  var ultimaLinha = guiaProduto.getLastRow();

  var dadosProdutos = guiaProduto.getRange(2, 1, ultimaLinha, 4).getValues();

  for(var i = 0; i < dadosProdutos.length; i++){

    if(dadosProdutos[i][0] == Cod){        
      
      var Produto = dadosProdutos[i][1];      

      var EM = dadosProdutos[i][2].toLocaleString({style: 'decimal', decimal: 'pt-BR'}); 
      var Minimo = EM.replace(/\./g,"");

      var MX = dadosProdutos[i][3].toLocaleString({style: 'decimal', decimal: 'pt-BR'}); 
      var Maximo = MX.replace(/\./g,"");
      
      dadosProdutos.length = 0;       

      return ([Produto, Minimo, Maximo]);

    }

  }

  dadosProdutos.length = 0;     
  return "PRODUTO NÃO ENCONTRADO!";

}


function EditarProduto(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){
      
    var Cod = Dados.Cod;    
    var Produto = Dados.Produto;
    var Minimo = Dados.Minimo;
    var Maximo = Dados.Maximo;

    var ultimaLinha = guiaProduto.getLastRow();

    var dadosProdutos = guiaProduto.getRange(2, 1, ultimaLinha, 2).getValues();

    for(var i = 0; i < dadosProdutos.length; i++){

      if(dadosProdutos[i][0] == Cod){

        var linha = i + 2;
        
        guiaProduto.getRange(linha,2).setValue(Produto);
        guiaProduto.getRange(linha,3).setValue(Minimo);
        guiaProduto.getRange(linha,4).setValue(Maximo);
       
        dadosProdutos.length = 0;       

        ProdutoEntrada(Cod, Produto);
        ProdutoSaida(Cod, Produto);

        return "EDITADO COM SUCESSO!";

      }

    }
    
    dadosProdutos.length = 0; 
   
    return "PRODUTO NÃO ENCONTRADO!";

  }

}

function ExcluirProduto(Cod){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){ 

    var ultimaLinha = guiaEntrada.getLastRow();

    var dadosEntrada = guiaEntrada.getRange(2, 1, ultimaLinha, 6).getValues();

    var ver = dadosEntrada.filter(function(value, i, arr){
      return Cod == arr[i][1];
    });

    if(ver.length > 0){
      dadosEntrada.length = 0;     
      return "NÃO PODE SER EXCLUÍDO. JÁ TEM LANÇAMENTO DE ENTRADA!";
    }

    var ultimaLinha = guiaProduto.getLastRow();

    var dadosProdutos = guiaProduto.getRange(2,1,ultimaLinha,2).getValues();

    for(var i = 0; i < dadosProdutos.length; i++){

      if(dadosProdutos[i][0] == Cod){

        var linha = i + 2;
        guiaProduto.deleteRow(linha);

        dadosEntrada.length = 0;
        dadosProdutos.length = 0;
        ver.length = 0;
        return "EXCLUÍDO COM SUCESSO!";

      }

    }

    dadosEntrada.length = 0;
    dadosProdutos.length = 0;
    ver.length = 0;

    return "PRODUTO NÃO ENCONTRADO!";

  }

}


function ProdutoEntrada(Cod, Produto){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){    

    var ultimaLinha = guiaEntrada.getLastRow();

    var dadosEntrada = guiaEntrada.getRange(2, 2, ultimaLinha, 1).getValues();

    for(var i = 0; i < dadosEntrada.length; i++){

      if(dadosEntrada[i][0] == Cod){

        var linha = i + 2;
        guiaEntrada.getRange(linha,3).setValue(Produto);

      }

    }

    dadosEntrada.length = 0;

  }

}

function ProdutoSaida(Cod, Produto){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){    

    var ultimaLinha = guiaSaida.getLastRow();

    var dadosSaida = guiaSaida.getRange(2, 2, ultimaLinha, 1).getValues();

    for(var i = 0; i < dadosSaida.length; i++){

      if(dadosSaida[i][0] == Cod){

        var linha = i + 2;
        guiaSaida.getRange(linha,3).setValue(Produto);

      }

    }

    dadosSaida.length = 0;

  }

}


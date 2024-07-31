function onOpen(){
 SpreadsheetApp.getUi()
 .createMenu('Formulários') 
 .addItem('Produtos', 'FormProduto') 
 .addItem('Inventário', 'FormInventario') 
 .addItem('Relatório', 'FormRelatorio')
 .addItem('Filtro', 'FormFiltro')  
 .addToUi();
}

var planilha = SpreadsheetApp.getActiveSpreadsheet();
var guiaProduto = planilha.getSheetByName("Produtos"); 
var guiaContagem = planilha.getSheetByName("Contagem");

function FormProduto() {
  
  var Form = HtmlService.createTemplateFromFile("FormProduto");

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

  MostrarForm.setTitle("CADASTRO DE PRODUTOS").setHeight(190).setWidth(385);

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
    
    guiaProduto.getRange("A2:B").sort([{column: 2, ascending: true}]);

    dadosProdutos.length = 0;

    ProdutoContagem(Cod, Produto);

    return "REGISTRADO COM SUCESSO!";

  }

}

function ProdutoContagem(Cod, Produto){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){ 

    var ultimaLinha = guiaContagem.getLastRow();

    var dadosContagem = guiaContagem.getRange(2, 2, ultimaLinha, 1).getValues();

    for(var i = 0; i < dadosContagem.length; i++){
      if(dadosContagem[i][0] == Cod){
        var linha = i + 2;
        guiaContagem.getRange(linha,3).setValue(Produto);
      }
    }

    dadosContagem.length = 0;

  }

}

function PesquisarProduto(Cod){  

  var ultimaLinha = guiaProduto.getLastRow();

  var dadosProdutos = guiaProduto.getRange(2, 1, ultimaLinha, 2).getValues();

  for(var i = 0; i < dadosProdutos.length; i++){
    if(dadosProdutos[i][0] == Cod){ 
      var Produto = dadosProdutos[i][1];      
      dadosProdutos.length = 0; 
      return (Produto);
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

    var ultimaLinha = guiaProduto.getLastRow();

    var dadosProdutos = guiaProduto.getRange(2, 1, ultimaLinha, 1).getValues();

    for(var i = 0; i < dadosProdutos.length; i++){
      if(dadosProdutos[i][0] == Cod){
        var linha = i + 2;        
        guiaProduto.getRange(linha,2).setValue(Produto);    
        dadosProdutos.length = 0; 
        ProdutoContagem(Cod, Produto);
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

    var ultimaLinha = guiaContagem.getLastRow();
    var dadosContagem = guiaContagem.getRange(2, 2, ultimaLinha, 1).getValues();

    var ver = dadosContagem.filter(function(value, i, arr){
      return Cod == arr[i][0];
    });

    if(ver.length > 0){
      dadosContagem.length = 0;     
      return "NÃO PODE SER EXCLUÍDO. JÁ TEM LANÇAMENTO DE CONTAGEM!";
    }

    var ultimaLinha = guiaProduto.getLastRow();
    var dadosProdutos = guiaProduto.getRange(2, 1, ultimaLinha,1).getValues();

    for(var i = 0; i < dadosProdutos.length; i++){
      if(dadosProdutos[i][0] == Cod){
        var linha = i + 2;
        guiaProduto.deleteRow(linha);
        dadosContagem.length = 0;
        dadosProdutos.length = 0;
        ver.length = 0;
        return "EXCLUÍDO COM SUCESSO!";
      }
    }

    dadosContagem.length = 0;
    dadosProdutos.length = 0;
    ver.length = 0;
    return "PRODUTO NÃO ENCONTRADO!";

  }

}




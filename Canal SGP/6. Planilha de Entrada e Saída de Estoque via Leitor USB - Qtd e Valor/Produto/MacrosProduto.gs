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

  MostrarForm.setTitle("CADASTRO DE PRODUTOS").setHeight(245).setWidth(382);

  SpreadsheetApp.getUi().showModalDialog(MostrarForm,"CADASTRO DE PRODUTOS");

}

function Chamar(Arquivo){
  return HtmlService.createHtmlOutputFromFile(Arquivo).getContent();
}


function PesquisarProduto(Cod){  

  var ultimaLinha = guiaProduto.getLastRow();

  var dadosProdutos = guiaProduto.getRange(2, 1, ultimaLinha, 4).getValues();

  for(var i = 0; i < dadosProdutos.length; i++){

    if(dadosProdutos[i][0] == Cod){        
      
      var Produto = dadosProdutos[i][1];     

      var C = dadosProdutos[i][2].toLocaleString({style: 'decimal', decimal: 'pt-BR'}); 
      var Compra = C.replace(/\./g,"");

      var V = dadosProdutos[i][3].toLocaleString({style: 'decimal', decimal: 'pt-BR'}); 
      var Venda = V.replace(/\./g,"");
      
      dadosProdutos.length = 0;       

      return ([Produto, Compra, Venda]);

    }

  }

  dadosProdutos.length = 0;     
  return "PRODUTO NÃO ENCONTRADO!";

}


function SalvarProduto(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

    var Cod = Dados.Cod;   
    var Produto = Dados.Produto; 
    var Compra = Dados.Compra;
    var Venda = Dados.Venda;

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
    guiaProduto.getRange(linha,3).setValue(Compra);
    guiaProduto.getRange(linha,4).setValue(Venda);
    
    guiaProduto.getRange("A2:D").sort([{column: 2, ascending: true}]);

    dadosProdutos.length = 0;

    ProdutoEntrada(Cod, Produto, Compra, Venda);
    ProdutoSaida(Cod, Produto, Compra, Venda);

    return "REGISTRADO COM SUCESSO!";

  }

}


function EditarProduto(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){
      
    var Cod = Dados.Cod;    
    var Produto = Dados.Produto;
    var Compra = Dados.Compra;
    var Venda = Dados.Venda;

    var ultimaLinha = guiaProduto.getLastRow();

    var dadosProdutos = guiaProduto.getRange(2, 1, ultimaLinha, 2).getValues();

    for(var i = 0; i < dadosProdutos.length; i++){

      if(dadosProdutos[i][0] == Cod){

        var linha = i + 2;
        
        guiaProduto.getRange(linha,2).setValue(Produto);
        guiaProduto.getRange(linha,3).setValue(Compra);
        guiaProduto.getRange(linha,4).setValue(Venda);
       
        dadosProdutos.length = 0;

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


function ProdutoEntrada(Cod, Produto, Compra){  

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){    

    var ultimaLinha = guiaEntrada.getLastRow();

    var dadosEntrada = guiaEntrada.getRange(2, 2, ultimaLinha, 3).getValues();

    for(var i = 0; i < dadosEntrada.length; i++){

      if(dadosEntrada[i][0] == Cod){

        var linha = i + 2;        

        var Qtd = dadosEntrada[i][2];

        if(Qtd == "" || Compra == ""){
          dadosEntrada.length = 0;
          return false;          
        } 

        var ConvertCompra = parseFloat(Compra.replace(/\,/g,'.'));
      
        var V = parseFloat(Qtd * ConvertCompra).toString();        
        
        var Valor = V.replace(/\./g,',');        

        guiaEntrada.getRange(linha,3).setValue(Produto);
        guiaEntrada.getRange(linha,5).setValue(Valor);

      }

    }

    dadosEntrada.length = 0;

  }

}

function ProdutoSaida(Cod, Produto, Compra, Venda){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){    

    var ultimaLinha = guiaSaida.getLastRow();

    var dadosSaida = guiaSaida.getRange(2, 2, ultimaLinha, 3).getValues();

    for(var i = 0; i < dadosSaida.length; i++){

      if(dadosSaida[i][0] == Cod){

        var linha = i + 2;        

        var Qtd = dadosSaida[i][2];

        var ConvertCompra = parseFloat(Compra.replace(/\,/g,'.'));
        var ConvertVenda = parseFloat(Venda.replace(/\,/g,'.'));
      
        var C = parseFloat(Qtd * ConvertCompra).toString();
        var Custo = C.replace(/\./g,',');

        var V = parseFloat(Qtd * ConvertVenda).toString(); 
        var Venda = V.replace(/\./g,',');

        guiaSaida.getRange(linha,3).setValue(Produto);
        guiaSaida.getRange(linha,5).setValue(Custo);
        guiaSaida.getRange(linha,6).setValue(Venda); 

      }

    }

    dadosSaida.length = 0;

  }

}


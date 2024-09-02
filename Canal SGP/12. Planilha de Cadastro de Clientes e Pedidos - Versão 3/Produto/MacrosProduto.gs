function FormProdutos() {

var planilha = SpreadsheetApp.getActiveSpreadsheet();
var guiaProduto = planilha.getSheetByName("Produtos");
var ultimaLinha = guiaProduto.getLastRow();

var dadosLinhas = guiaProduto.getRange(2,1,ultimaLinha,1).getValues();

var b = {};

for(var i = 0; i < dadosLinhas.length; i++){
  b[dadosLinhas[i][0]] = dadosLinhas[i][0];
}

var listaUnica = [];

for(var key in b){
  listaUnica.push([key]);
}

dadosLinhas.length = 0;

var list = listaUnica;

list.sort();

var Form = HtmlService.createTemplateFromFile("FormProduto");

Form.list = list.map(function(r){return r[0];});

var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

MostrarForm.setTitle("Cadastro de Produtos").setHeight(380).setWidth(510);

SpreadsheetApp.getUi().showModalDialog(MostrarForm,"Cadastro de Produtos");

}


function ListaProdutos(Linha){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaProduto = planilha.getSheetByName("Produtos");

  var ultimaLinha = guiaProduto.getLastRow();

  var dadosProdutos = guiaProduto.getRange(2,1,ultimaLinha,2).getValues();

  var produtos = [];

  for(var i = 0; i < dadosProdutos.length; i++){
    if(dadosProdutos[i][0] == Linha){
      produtos.push([dadosProdutos[i][1]]);
    }
  }

  dadosProdutos.length = 0;
  return produtos.sort();

}


function SalvarProduto(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

    var Linha = Dados.Linha;
    var Produto = Dados.Produto;
    var Preco = Dados.Preco;

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaProduto = planilha.getSheetByName("Produtos");

    var ultimaLinha = guiaProduto.getLastRow();

    var dadosProdutos = guiaProduto.getRange(2,1,ultimaLinha,2).getValues();

    for(var i = 0; i<dadosProdutos.length; i++){

      if(dadosProdutos[i][0] == Linha && dadosProdutos[i][1] == Produto){
          return "PRODUTO JÁ CADASTRADO!";
      }

    }

    var linha = guiaProduto.getLastRow() + 1;

    guiaProduto.getRange(linha,1).setValue(Linha);
    guiaProduto.getRange(linha,2).setValue(Produto);
    guiaProduto.getRange(linha,3).setValue(Preco);

    guiaProduto.getRange("A2:C").sort([{column: 1, ascending: true},{column: 2, ascending: true}]);

     return "REGISTRADO COM SUCESSO!";

  }

}

function AtualizarListaLinhas(){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaProduto = planilha.getSheetByName("Produtos");

  var ultimaLinha = guiaProduto.getLastRow();

  var dadosLinhas = guiaProduto.getRange(2,1,ultimaLinha,1).getValues();

  var b = {};

  for(var i = 0; i < dadosLinhas.length; i++){
    b[dadosLinhas[i][0]] = dadosLinhas[i][0];
  }

  var listaUnica = [];

  for(var key in b){
    listaUnica.push([key]);
  }

  dadosLinhas.length = 0;

  return listaUnica.sort();

}

function PesquisarProduto(Dados){

    var Linha = Dados.Linha;
    var Produto = Dados.Produto;

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaProduto = planilha.getSheetByName("Produtos");

    var ultimaLinha = guiaProduto.getLastRow();

    var dadosProdutos = guiaProduto.getRange(2,1,ultimaLinha,3).getValues();

    for(var i = 0; i <dadosProdutos.length; i++){

      if(dadosProdutos[i][0] == Linha && dadosProdutos[i][1] == Produto){

        var Preco = dadosProdutos[i][2].toLocaleString({style: 'decimal',decimal: 'pt-BR'});
        var Preco = Preco.replace(/\./g,"");
        
        dadosProdutos.length = 0;

        return ([Preco]);

      }

    }
   
    dadosProdutos.length = 0;    
    return "NÃO ENCONTRADO!";

}

function EditarProduto(Dados){

const user = LockService.getScriptLock();
user.tryLock(10000);

if(user.hasLock()){

  var LinhaLista = Dados.LinhaLista;
  var ProdutoLista = Dados.ProdutoLista;

  var Linha = Dados.Linha;
  var Produto = Dados.Produto;
  var Preco = Dados.Preco;

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaProduto = planilha.getSheetByName("Produtos");
  var guiaPedido = planilha.getSheetByName("Pedidos");

  var ultimaLinha = guiaProduto.getLastRow();

  var dadosProdutos = guiaProduto.getRange(2,1,ultimaLinha,2).getValues();

  var ultimaLinha = guiaPedido.getLastRow();

  var dadosPedidos = guiaPedido.getRange(2,6,ultimaLinha,2).getValues();

  var ver = dadosPedidos.filter(function(value,i,arr){
    return LinhaLista == arr[i][0] && ProdutoLista == arr[i][1];
  })

  for(var i = 0; i <dadosProdutos.length; i++){

    if(dadosProdutos[i][0] == LinhaLista && dadosProdutos[i][1] == ProdutoLista ){

        var linha = i + 2;

        if(ver.length < 1){

            guiaProduto.getRange(linha,1).setValue(Linha);
            guiaProduto.getRange(linha,2).setValue(Produto);
            guiaProduto.getRange(linha,3).setValue(Preco);

            dadosProdutos.length = 0;
            dadosPedidos.length = 0;

            return "EDITADO COM SUCESSO!";

        }

        if(ver.length > 0){

          guiaProduto.getRange(linha,3).setValue(Preco);

          dadosProdutos.length = 0;
          dadosPedidos.length = 0;
          ver.length = 0;

          return "EDITADO APENAS O PREÇO. PRODUTO JÁ POSSUI LANÇAMENTO DE PEDIDO.";

        }

    }

  }

dadosProdutos.length = 0;
dadosPedidos.length = 0;
ver.length = 0;

return "PRODUTO NÃO ENCONTRADO!";

}

}


function ExcluirProduto(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

    var Linha = Dados.LinhaLista;
    var Produto = Dados.ProdutoLista;

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaProduto = planilha.getSheetByName("Produtos");
    var guiaPedido = planilha.getSheetByName("Pedidos");

    var ultimaLinha = guiaProduto.getLastRow();

    var dadosProdutos = guiaProduto.getRange(2,1,ultimaLinha,2).getValues();

    var ultimaLinha = guiaPedido.getLastRow();

    var dadosPedidos = guiaPedido.getRange(2,6,ultimaLinha,2).getValues();

    var ver = dadosPedidos.filter(function(value,i,arr){
      return Linha == arr[i][0] && Produto == arr[i][1];
    });

    if(ver.length > 0){
      dadosProdutos.length = 0;
      dadosPedidos.length = 0;
      ver.length = 0;
      return "PRODUTO NÃO PODE SER EXCLUÍDO. PORQUE JÁ TEM LANÇAMENTO DE PEDIDO!";
    }

    for(var i = 0; i < dadosProdutos.length; i++){

      if(dadosProdutos[i][0] == Linha && dadosProdutos[i][1] == Produto){

        var linha = i + 2;
        guiaProduto.deleteRow(linha);

        dadosProdutos.length = 0;
        dadosPedidos.length = 0;
        ver.length = 0;

        return "EXCLUÍDO COM SUCESSO!";

      }

    }

    dadosProdutos.length = 0;
    dadosPedidos.length = 0;
    ver.length = 0;

    return "PRODUTO NÃO ENCONTRADO!";

  }

}

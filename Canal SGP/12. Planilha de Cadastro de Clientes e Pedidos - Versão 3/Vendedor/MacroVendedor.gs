function onOpen(){

  SpreadsheetApp.getUi()
  .createMenu('Formulários')
  .addItem('Cadastro de Vendedores', 'FormVendedor')
  .addItem('Cadastro de Produtos', 'FormProdutos')
  .addItem('Cadastro de Clientes', 'FormCliente')
  .addItem('Cadastro de Pedidos', 'FormPedido')
  .addItem('Filtro Clientes', 'FormFiltroClientes')
  .addItem('Filtro Pedidos', 'FormFiltroPedidos')
  .addItem('Relatório', 'FormRelatorio')
  .addItem('Gráficos', 'FormGrafico')
  .addToUi();

}

function FormVendedor() {

var planilha = SpreadsheetApp.getActiveSpreadsheet();
var guiaVendedor = planilha.getSheetByName("Vendedores");

var ultimaLinha = guiaVendedor.getLastRow() - 1;

if(ultimaLinha == 0){
  ultimaLinha = 1;
}

var list  = guiaVendedor.getRange(2,1,ultimaLinha,1).getValues();

list.sort();

var Form = HtmlService.createTemplateFromFile("FormVendedor");

Form.list = list.map(function(r){return r[0];});

var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

MostrarForm.setTitle("Cadastro de Vendedores").setHeight(300).setWidth(510);

SpreadsheetApp.getUi().showModalDialog(MostrarForm,"Cadastro de Vendedores");

}


function Chamar(Arquivo){
  return HtmlService.createHtmlOutputFromFile(Arquivo).getContent();
}


function SalvarVendedor(Dados){

  const user = LockService.getScriptLock();

  user.tryLock(10000);

  if(user.hasLock()){

    var Nome = Dados.Nome;
    var Telefone = Dados.Telefone;

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaVendedor = planilha.getSheetByName("Vendedores");

    var ultimaLinha = guiaVendedor.getLastRow();

    var dadosVendedor = guiaVendedor.getRange(2,1,ultimaLinha,1).getValues();

    for(var linha = 0; linha<dadosVendedor.length; linha ++){

      if(dadosVendedor[linha][0] == Nome){
        dadosVendedor.length = 0;
        return "VENDEDOR JÁ CADASTRADO!";
      }

    }

    var linha = guiaVendedor.getLastRow();
    var linha = linha + 1;

    guiaVendedor.getRange(linha,1).setValue(Nome);
    guiaVendedor.getRange(linha,2).setValue(Telefone);

    dadosVendedor.length = 0;
    return "REGISTRADO COM SUCESSO!";

  }  

}


function AtualizarListaVendedores(){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaVendedor = planilha.getSheetByName("Vendedores");

  var ultimaLinha = guiaVendedor.getLastRow() - 1;

  if(ultimaLinha == 0){
    var ultimaLinha = 1;
  }

  var list = guiaVendedor.getRange(2, 1, ultimaLinha,1).getValues();

  return list.sort();

}


function PesquisarVendedor(NomeVendedor){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaVendedor = planilha.getSheetByName("Vendedores");

  var ultimaLinha = guiaVendedor.getLastRow();

  var dadosVendedor = guiaVendedor.getRange(2,1,ultimaLinha,2).getValues();

  for(var linha = 0; linha<dadosVendedor.length; linha++) {

    if(dadosVendedor[linha][0] == NomeVendedor){

      var Nome = dadosVendedor[linha][0];
      var Telefone = dadosVendedor[linha][1];

      dadosVendedor.length = 0;
      return ([Nome,Telefone]);

    }

  }

  dadosVendedor.length = 0;
  return "NÃO ENCONTRADO!";

}


function EditarVendedor(Dados){

  const user = LockService.getScriptLock();

  user.tryLock(10000);

  if(user.hasLock()){

    var Vendedor = Dados.NomeLista;
    var Nome = Dados.Nome;
    var Telefone = Dados.Telefone;

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaVendedor = planilha.getSheetByName("Vendedores");
    var guiaPedido = planilha.getSheetByName("Pedidos");

    var ultimaLinha = guiaVendedor.getLastRow();

    var dadosVendedor = guiaVendedor.getRange(2,1,ultimaLinha,1).getValues();

    var ultimaLinha = guiaPedido.getLastRow();

    var dadosPedidos = guiaPedido.getRange(2,12,ultimaLinha,1).getValues();

    var ver = dadosPedidos.filter(function(value,i,arr){
        return Vendedor == arr[i][0];
    });

    for(var i = 0; i<dadosVendedor.length; i++){

        if(dadosVendedor[i][0] == Vendedor){

            if(ver.length < 1){

                var linha = i + 2;

                guiaVendedor.getRange(linha,1).setValue(Nome);
                guiaVendedor.getRange(linha,2).setValue(Telefone);

                dadosPedidos.length = 0;
                dadosVendedor.length = 0;
                ver.length = 0;
                return "EDITADO COM SUCESSO!";

            }

            if(ver.length > 0){

              var linha = i + 2;
              guiaVendedor.getRange(linha,2).setValue(Telefone);

              dadosPedidos.length = 0;
              dadosVendedor.length = 0;
              ver.length = 0;
              return "VENDEDOR JÁ POSSUI LANÇAMENTOS DE PEDIDO. EDITADO APENAS TELEFONE!";

            }

        }

    }

  dadosPedidos.length = 0;
  dadosVendedor.length = 0;
  ver.length = 0;
  return "VENDEDOR NÃO ENCONTRADO!"; 

  }  

}

function ExcluirVendedor(Vendedor){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaVendedor = planilha.getSheetByName("Vendedores");
    var guiaPedido = planilha.getSheetByName("Pedidos");

    var ultimaLinha = guiaVendedor.getLastRow();
    var dadosVendedor = guiaVendedor.getRange(2,1,ultimaLinha,1).getValues();

    var ultimaLinha = guiaPedido.getLastRow();
    var dadosPedidos = guiaPedido.getRange(2,12,ultimaLinha,1).getValues();

    var ver = dadosPedidos.filter(function(value,i,arr){
      return Vendedor == arr[i][0];
    });

    if(ver.length > 0){
        dadosVendedor.length = 0;
        dadosPedidos.length = 0;
        ver.length = 0;
        return "VENDEDOR JÁ POSSUI LANÇAMENTO DE PEDIDOS. NÃO PODE SER EXCLUÍDO!";
    }

    for(var i = 0; i < dadosVendedor.length; i++){

      if(dadosVendedor[i][0] == Vendedor){
        var linha = i + 2;
        guiaVendedor.deleteRow(linha);
        dadosVendedor.length = 0;
        dadosPedidos.length = 0;
        ver.length = 0;
        return "EXCLUÍDO COM SUCESSO!";
      }

    }

    dadosVendedor.length = 0;
    dadosPedidos.length = 0;
    ver.length = 0;
    return "VENDEDOR NÃO ENCONTRADO!";

  }

}


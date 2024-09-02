function FormPedido(Id) {

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaCliente = planilha.getSheetByName("Clientes");
    var guiaEstado = planilha.getSheetByName("Estados");
    var guiaProduto = planilha.getSheetByName("Produtos");
    var guiaVendedor = planilha.getSheetByName("Vendedores");
  
    var ultimaLinha = guiaCliente.getLastRow();
  
    if(ultimaLinha == 0){
      var ultimaLinha = 1;
    }
  
    var list = guiaCliente.getRange(2, 2, ultimaLinha, 1).getValues();
  
    list.sort();
  
    var ultimaLinha = guiaEstado.getLastRow() - 1;
  
    if(ultimaLinha == 0){
      var ultimaLinha = 1;
    }
  
    var list2 = guiaEstado.getRange(2, 1, ultimaLinha, 1).getValues();
  
    list2.sort();
  
    var ultimaLinha = guiaProduto.getLastRow() - 1;
  
    if(ultimaLinha == 0){
      var ultimaLinha = 1;
    }
  
    var dadosLinhas = guiaProduto.getRange(2,1, ultimaLinha, 1).getValues();
  
    var listaUnica = [...new Set(dadosLinhas.flat())];
  
    var listaLinhas = [];
  
    for(var i = 0; i < listaUnica.length; i++){
      listaLinhas.push([listaUnica[i]]);
    }  
  
    var list3 = listaLinhas.sort();
  
    var ultimaLinha = guiaVendedor.getLastRow() - 1;
  
    if(ultimaLinha == 0){
      var ultimaLinha = 1;
    }
  
    var list4 = guiaVendedor.getRange(2, 1,ultimaLinha,1).getValues();
  
    list4.sort();
  
    var Form = HtmlService.createTemplateFromFile("FormPedido");
  
    Form.list = list.map(function(r){
      return r[0];
    });
  
    Form.list2 = list2.map(function(r){
      return r[0];
    });
  
    Form.list3 = list3.map(function(r){
      return r[0];
    });
  
    Form.list4 = list4.map(function(r){
      return r[0];
    });
  
    Form.Id = Id;
  
    var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
  
    MostrarForm.setTitle("Cadastro de Pedidos").setHeight(650).setWidth(800);
  
    SpreadsheetApp.getUi().showModalDialog(MostrarForm, "Cadastro de Pedidos");
    
  }
  
  function buscaListas(){
  
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaCidades = planilha.getSheetByName("Estados/Cidades");
  
    var ultimaLinha = guiaCidades.getLastRow();
  
    var dadosCidades = guiaCidades.getRange(2, 1, ultimaLinha, 2).getValues();
  
    var guiaClientes = planilha.getSheetByName("Clientes");
  
    var ultimaLinha = guiaClientes.getLastRow();
  
    var dadosClientes = guiaClientes.getRange(2, 2, ultimaLinha, 7).getValues();
  
    var guiaProdutos = planilha.getSheetByName("Produtos");
  
    var ultimaLinha = guiaProdutos.getLastRow();
  
    var dadosProdutos = guiaProdutos.getRange(2, 1, ultimaLinha, 3).getValues();
  
    var arrays = {
      dadosCidades: dadosCidades,
      dadosClientes: dadosClientes,
      dadosProdutos: dadosProdutos,
    }
    
     return arrays;
  
  }
  
  function buscaPedidoId(){
  
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaPedido = planilha.getSheetByName("Pedidos");
  
    var novoPedido = Math.max.apply(null, guiaPedido.getRange("B2:B").getValues());
    var novoPedido = novoPedido + 1;
  
    var novoId = Math.max.apply(null, guiaPedido.getRange("A2:A").getValues());
    var novoId = novoId + 1;
  
    var dados = {
      novoPedido: novoPedido,
      novoId: novoId,
    }
  
    return dados;
  
  }
  
  function SalvarPedido(Dados){
  
    const user = LockService.getScriptLock();
    user.tryLock(10000);
  
    if(user.hasLock()){
  
      var planilha = SpreadsheetApp.getActiveSpreadsheet();
      var guiaPedido = planilha.getSheetByName("Pedidos");
  
      var linha = guiaPedido.getLastRow() + 1;
  
      var dataQuebrada = Dados.Data.split("/");
  
      var Ano = dataQuebrada[0];
      var Mes = dataQuebrada[1];
      var Dia = dataQuebrada[2];
  
      var Data = Dia + "/" + Mes + "/" + Ano;
  
      guiaPedido.getRange(linha,3).setValue(Data);
  
      var data = new Date(Dados.Data);
      var m = data.getMonth();
  
      var meses = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];
  
      var Mes = meses[m];
  
      guiaPedido.getRange(linha,1).setValue(Dados.Id);
      guiaPedido.getRange(linha,2).setValue(Dados.Pedido);
      guiaPedido.getRange(linha,4).setValue(Mes);
      guiaPedido.getRange(linha,5).setValue(Ano);
      guiaPedido.getRange(linha,6).setValue(Dados.Linha);
      guiaPedido.getRange(linha,7).setValue(Dados.Produto);
      guiaPedido.getRange(linha,8).setValue(Dados.Qtd);
      guiaPedido.getRange(linha,9).setValue(Dados.Preco);
      guiaPedido.getRange(linha,10).setValue(Dados.Total);
      guiaPedido.getRange(linha,11).setValue(Dados.Cliente);
      guiaPedido.getRange(linha,12).setValue(Dados.Vendedor);
      guiaPedido.getRange(linha,13).setValue(Dados.Estado);
      guiaPedido.getRange(linha,14).setValue(Dados.Cidade);
      guiaPedido.getRange(linha,15).setValue(Dados.Status);
      guiaPedido.getRange(linha,16).setValue(Dados.Obs);
  
      return "REGISTRADO COM SUCESSO!";
  
    }
  
  }
  
  function PesquisarPedido(Id){
  
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaPedido = planilha.getSheetByName("Pedidos");
  
    var ultimaLinha = guiaPedido.getLastRow();
  
    var dados = guiaPedido.getRange(2, 1, ultimaLinha, 16).getValues();
  
    for(var i = 0; i < dados.length; i++){
  
      if(dados[i][0] == Id){
  
        var Pedido = dados[i][1];
  
        var Data = Utilities.formatDate(new Date(dados[i][2]), planilha.getSpreadsheetTimeZone(), "yyyy-MM-dd");
  
        var Linha = dados[i][5];
        var Produto = dados[i][6];
        var Qtd = dados[i][7];
        var Preco = dados[i][8];
  
        var T = dados[i][9].toLocaleString({style:'decimal', decimal: 'pt-BR'});
        var Total = T.replace(/\./g,"");
  
        var Cliente = dados[i][10];
        var Vendedor = dados[i][11];
        var Estado = dados[i][12];
        var Cidade = dados[i][13];
        var Status = dados[i][14];
        var Obs = dados[i][15];
  
        dados.length = 0;
  
        return ([Pedido,Data,Linha,Produto,Qtd,Preco,Total,Cliente,Vendedor,Estado,Cidade,Status,Obs]);
  
      }
  
    }
    
    dados.length = 0;
    return "ID NÃO ENCONTRADO!";
  
  }
  
  function EditarPedido(Dados){
  
    const user = LockService.getScriptLock();
    user.tryLock(10000);
  
    if(user.hasLock()){
  
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaPedido = planilha.getSheetByName("Pedidos");
  
    var ultimaLinha = guiaPedido.getLastRow();
  
    var dados = guiaPedido.getRange(2, 1, ultimaLinha, 1).getValues();
  
    for(var i = 0; i < dados.length; i++){
  
      if(dados[i][0] == Dados.Id){
  
        var linha = i + 2;
  
        var dataQuebrada = Dados.Data.split("/");
  
        var Ano = dataQuebrada[0];
        var Mes = dataQuebrada[1];
        var Dia = dataQuebrada[2];
  
        var Data = Dia + "/" + Mes + "/" + Ano;
  
        guiaPedido.getRange(linha,3).setValue(Data);
  
        var data = new Date(Dados.Data);
        var m = data.getMonth();
  
        var meses = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];
  
        var Mes = meses[m];
       
        guiaPedido.getRange(linha,2).setValue(Dados.Pedido);
        guiaPedido.getRange(linha,4).setValue(Mes);
        guiaPedido.getRange(linha,5).setValue(Ano);
        guiaPedido.getRange(linha,6).setValue(Dados.Linha);
        guiaPedido.getRange(linha,7).setValue(Dados.Produto);
        guiaPedido.getRange(linha,8).setValue(Dados.Qtd);
        guiaPedido.getRange(linha,9).setValue(Dados.Preco);
        guiaPedido.getRange(linha,10).setValue(Dados.Total);
        guiaPedido.getRange(linha,11).setValue(Dados.Cliente);
        guiaPedido.getRange(linha,12).setValue(Dados.Vendedor);
        guiaPedido.getRange(linha,13).setValue(Dados.Estado);
        guiaPedido.getRange(linha,14).setValue(Dados.Cidade);
        guiaPedido.getRange(linha,15).setValue(Dados.Status);
        guiaPedido.getRange(linha,16).setValue(Dados.Obs);
  
        dados.length = 0;
        return "EDITADO COM SUCESSO!";
  
      }
  
     }
  
     dados.length = 0;
     return "ID NÃO ENCONTRADO!";
  
    }
  
  }
  
  function ExcluirPedido(Id){
  
    const user = LockService.getScriptLock();
    user.tryLock(10000);
  
    if(user.hasLock()){
  
      var planilha = SpreadsheetApp.getActiveSpreadsheet();
      var guiaPedido = planilha.getSheetByName("Pedidos");
  
      var ultimaLinha = guiaPedido.getLastRow();
  
      var dados = guiaPedido.getRange(2, 1, ultimaLinha, 1).getValues();
  
      for(var i = 0; i < dados.length; i++){
  
        if(dados[i][0] == Id){
  
          var linha = i + 2;
          guiaPedido.deleteRow(linha);
  
          dados.length = 0;
          return "EXCLUÍDO COM SUCESSO!";
  
        }
  
      }
  
      dados.length = 0;
      return "ID NÃO ENCONTRADO!";
  
    }  
  
  }
  
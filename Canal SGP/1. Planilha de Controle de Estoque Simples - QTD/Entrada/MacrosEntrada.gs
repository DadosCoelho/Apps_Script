function Linhamarca(linha){
  
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guia = planilha.getSheetByName("Linhas/Marca");
  var localPesquisa = guia.getRange(1, 2, 1, guia.getLastColumn()).getValues()[0];
    
  var resultado = localPesquisa.Pesquisa(linha);
  
  if (resultado !=-1){  
  
  var coluna = resultado + 2;
   
           var qtdlinha = 1; 
           while(guia.getRange(qtdlinha,coluna).isBlank() == false) {
              qtdlinha = qtdlinha + 1;
            };
         
         var qtdlinha = qtdlinha - 2;
         var dados = guia.getRange(2, coluna, qtdlinha).getValues();
         
         dados.sort();

        return dados;
  
  
  }else{
  
  return "LINHA NÃO ENCONTRADA"
  
  }

}


Array.prototype.Pesquisa = function(Procura){

  if (Procura == "") return false;
  
  for (var Linha= 0; Linha<this.length; Linha ++ )

  if (this[Linha]==Procura) return Linha;
  
  return -1

}


function listaProduto(dados){
  
  var Linha = dados.Linha;
  var Marca = dados.Marca;

  var planilha =SpreadsheetApp.getActiveSpreadsheet();
  var guiaProduto = planilha.getSheetByName("Produtos");

  var dadosPlan = guiaProduto.getRange(2, 1, guiaProduto.getLastRow() - 1, 3).getValues();

  var produtos = [];

  for(var linha = 0; linha<dadosPlan.length; linha++){         
     
     if(dadosPlan[linha][0] == Linha && dadosPlan[linha][1] == Marca){ 
        
        var p = dadosPlan[linha][2];
        produtos.push(p);
         
     }    

   }

  var b = {};

  for (var i = 0; i < produtos.length; i++) {
      b[produtos[i]] = produtos[i];
  }
  
  var listaProdutos = [];
    
  for (var key in b) {
      listaProdutos.push([key]);
  }

  dadosPlan.length = 0;
  produtos.length = 0;

  return listaProdutos.sort();

}


function SalvarEntrada(Dados){

const user = LockService.getScriptLock();
user.tryLock(10000);

if(user.hasLock()){

var planilha = SpreadsheetApp.getActiveSpreadsheet();
var guiaEntrada = planilha.getSheetByName("Entradas");

var novoid = Math.max.apply(null, guiaEntrada.getRange("A2:A").getValues()); 
var novoid = novoid + 1

var linha = guiaEntrada.getLastRow() + 1;

var dataQuebrada = Dados.Data.split("/");
var Ano = dataQuebrada[0];
var Mes = dataQuebrada[1];
var Dia = dataQuebrada[2];
var Data = Dia + "/" + Mes + "/" + Ano;      

var Data = new Date(Dados.Data);
var m = Data.getMonth();

var meses = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];
var Mes = meses[m];

guiaEntrada.getRange(linha, 1).setValue(novoid);
guiaEntrada.getRange(linha, 2).setValue(Data);
guiaEntrada.getRange(linha, 3).setValue(Mes);
guiaEntrada.getRange(linha, 4).setValue(Ano);
guiaEntrada.getRange(linha, 5).setValue(Dados.Linha);
guiaEntrada.getRange(linha, 6).setValue(Dados.Marca);
guiaEntrada.getRange(linha, 7).setValue(Dados.Produto);
guiaEntrada.getRange(linha, 8).setValue(Dados.Cod);
guiaEntrada.getRange(linha, 9).setValue(Dados.Qtd);
guiaEntrada.getRange(linha, 10).setValue(Dados.Nf);
guiaEntrada.getRange(linha, 11).setValue(Dados.Obs);

return "SALVO COM SUCESSO!";

}

}


function PesquisarEntrada(id){

var planilha = SpreadsheetApp.getActiveSpreadsheet()
var guiaEntrada = planilha.getSheetByName("Entradas");

var dados = guiaEntrada.getRange(2, 1, guiaEntrada.getLastRow(),11).getValues();

for(var linha = 0; linha<dados.length; linha++){         
     
     if(dados[linha][0] == id){                   
        
        var V = dados[linha][8].toLocaleString({style: 'decimal', decimal: 'pt-BR'});
        var V = V.replace(".","");

        var valor = V;
        
        var data = new Date(dados[linha][1]);
            
        var d = data.getDate();
        var m = data.getMonth() + 1;
        var a = data.getFullYear();
        
        var datacarregar = a + "-" + m + "-" + d
        
        var Carregar={}; 

        Carregar.Id = dados[linha][0];
        Carregar.Data = datacarregar;
        Carregar.Linha = dados[linha][4];      
        Carregar.Marca = dados[linha][5];         
        Carregar.Produto = dados[linha][6];
        Carregar.Cod = dados[linha][7];
        Carregar.Qtd = valor;  
        Carregar.Nf = dados[linha][9];
        Carregar.Obs = dados[linha][10];
        
        dados.length = 0;

        return ([Carregar.Id,Carregar.Data, Carregar.Linha, Carregar.Marca, Carregar.Produto, Carregar.Cod, Carregar.Qtd, Carregar.Nf,Carregar.Obs]);
         
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

var dados = guiaEntrada.getRange(2, 1, guiaEntrada.getLastRow()).getValues();

for(var linha = 0; linha<dados.length; linha++){
          
     if(dados[linha][0] == Dados.Id){   
     
      var linha = linha + 2;      
      
      var dataQuebrada = Dados.Data.split("/");
      var Ano = dataQuebrada[0];
      var Mes = dataQuebrada[1];
      var Dia = dataQuebrada[2];
      var Data = Dia + "/" + Mes + "/" + Ano;      

      var Data = new Date(Dados.Data);
      var m = Data.getMonth();

      var meses = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];
      var Mes = meses[m] ; 
      
      guiaEntrada.getRange(linha, 2).setValue(Data);
      guiaEntrada.getRange(linha, 3).setValue(Mes);
      guiaEntrada.getRange(linha, 4).setValue(Ano);
      guiaEntrada.getRange(linha, 5).setValue(Dados.Linha);
      guiaEntrada.getRange(linha, 6).setValue(Dados.Marca);
      guiaEntrada.getRange(linha, 7).setValue(Dados.Produto);
      guiaEntrada.getRange(linha, 8).setValue(Dados.Cod);
      guiaEntrada.getRange(linha, 9).setValue(Dados.Qtd);
      guiaEntrada.getRange(linha, 10).setValue(Dados.Nf);    
      guiaEntrada.getRange(linha, 11).setValue(Dados.Obs);
      
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

    var dados = guiaEntrada.getRange(2, 1, guiaEntrada.getLastRow()).getValues();

    for(var linha = 0; linha<dados.length; linha++){
              
        if(dados[linha][0] == id){           
          
              var linha = linha + 2;
              
              guiaEntrada.deleteRow(linha);           

              dados.length = 0;

              return "EXCLUÍDO COM SUCESSO!";  
                      
        }

      }

    dados.length = 0;
    
    return "NÃO ENCONTRADO!";

  }

}




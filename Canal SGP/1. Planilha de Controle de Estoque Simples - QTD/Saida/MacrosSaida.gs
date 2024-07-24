function FormSaida(Id){
   
var planilha = SpreadsheetApp.getActiveSpreadsheet();
var guialinha = planilha.getSheetByName("Linhas/Marca");

var linha = 1;

while(guialinha.getRange(linha,1).isBlank() == false){                         
    linha = linha + 1;
};

if (linha < 3){
  linha = 3;
}

var list = guialinha.getRange(2, 1,linha -2,1).getValues();

list.sort();

var Form = HtmlService.createTemplateFromFile("FormSaida");

Form.list = list.map(function(r){ return r[0];});
Form.Id = Id;

var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

MostrarForm.setTitle("SAÍDA ESTOQUE").setHeight(335).setWidth(650);
  
SpreadsheetApp.getUi().showModalDialog(MostrarForm, "SAÍDA ESTOQUE");
  
}


function SalvarSaida(Dados){

const user = LockService.getScriptLock();
user.tryLock(10000);

if(user.hasLock()){

var planilha = SpreadsheetApp.getActiveSpreadsheet();
var guiaSaida = planilha.getSheetByName("Saidas");

var novoid = Math.max.apply(null, guiaSaida.getRange("A2:A").getValues()); 
var novoid = novoid + 1;

var linha = guiaSaida.getLastRow() + 1;

var dataQuebrada = Dados.Data.split("/");
var Ano = dataQuebrada[0];
var Mes = dataQuebrada[1];
var Dia = dataQuebrada[2];
var Data = Dia + "/" + Mes + "/" + Ano;      

var Data = new Date(Dados.Data);
var m = Data.getMonth();

var meses = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];
var Mes = meses[m] ; 

guiaSaida.getRange(linha, 1).setValue(novoid);
guiaSaida.getRange(linha, 2).setValue(Data);
guiaSaida.getRange(linha, 3).setValue(Mes);
guiaSaida.getRange(linha, 4).setValue(Ano);
guiaSaida.getRange(linha, 5).setValue(Dados.Linha);
guiaSaida.getRange(linha, 6).setValue(Dados.Marca);
guiaSaida.getRange(linha, 7).setValue(Dados.Produto);
guiaSaida.getRange(linha, 8).setValue(Dados.Cod);
guiaSaida.getRange(linha, 9).setValue(Dados.Qtd);
guiaSaida.getRange(linha, 10).setValue(Dados.Obs);

return "SALVO COM SUCESSO!";

}

}

function PesquisarSaida(id){

var planilha = SpreadsheetApp.getActiveSpreadsheet()
var guiaSaida = planilha.getSheetByName("Saidas");

var dados = guiaSaida.getRange(2, 1, guiaSaida.getLastRow(),10).getValues();

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
        Carregar.Obs = dados[linha][9];
        
        dados.length = 0;

        return ([Carregar.Id,Carregar.Data, Carregar.Linha, Carregar.Marca, Carregar.Produto, Carregar.Cod, Carregar.Qtd, Carregar.Obs])     
         
         
     }

}

dados.length = 0;
return "NÃO ENCONTRADO!";

}


function EditarSaida(Dados){

const user = LockService.getScriptLock();
user.tryLock(10000);

if(user.hasLock()){

var planilha = SpreadsheetApp.getActiveSpreadsheet();
var guiaSaida = planilha.getSheetByName("Saidas");

var dados = guiaSaida.getRange(2, 1, guiaSaida.getLastRow()).getValues();

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

      guiaSaida.getRange(linha, 2).setValue(Data);
      guiaSaida.getRange(linha, 3).setValue(Mes);
      guiaSaida.getRange(linha, 4).setValue(Ano);
      guiaSaida.getRange(linha, 5).setValue(Dados.Linha);
      guiaSaida.getRange(linha, 6).setValue(Dados.Marca);
      guiaSaida.getRange(linha, 7).setValue(Dados.Produto);
      guiaSaida.getRange(linha, 8).setValue(Dados.Cod);
      guiaSaida.getRange(linha, 9).setValue(Dados.Qtd);    
      guiaSaida.getRange(linha, 10).setValue(Dados.Obs);
      
      dados.length = 0;

      return "EDITADO COM SUCESSO!";

    }
}

dados.length = 0;
return "ID NÃO ENCONTRADO!";

}

}


function ExcluirSaida(id){

const user = LockService.getScriptLock();
user.tryLock(10000);

if(user.hasLock()){

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaSaida = planilha.getSheetByName("Saidas");

    var dados = guiaSaida.getRange(2, 1, guiaSaida.getLastRow()).getValues();

    for(var linha = 0; linha<dados.length; linha++){
              
        if(dados[linha][0] == id){           
          
              var linha = linha + 2;
              
              guiaSaida.deleteRow(linha);           

              dados.length = 0;
              return "EXCLUÍDO COM SUCESSO!";
            
        }

    }

    dados.length = 0;
    return "NÃO ENCONTRADO!";

  }

}


function Saldo(Dados){

var planilha = SpreadsheetApp.getActiveSpreadsheet();
var guiaentrada = planilha.getSheetByName("Entradas");

var linha = Dados.linha;
var marca = Dados.marca;
var produto = Dados.produto;  
 
  
var dadosentrada = guiaentrada.getRange(2, 5, guiaentrada.getLastRow(),5).getValues();
var totalqtdentrada = "0";

for(var i = 0; i<dadosentrada.length; i++){
          
     if(dadosentrada[i][0] == linha && dadosentrada[i][1] == marca && dadosentrada[i][2] == produto){
     
       var qtd = parseFloat(dadosentrada[i][4]);                  
       totalqtdentrada = parseFloat(totalqtdentrada) + parseFloat(qtd);
          
    }
}

dadosentrada.length = 0;
i = 0;
      
var guiasaida = planilha.getSheetByName("Saidas");
var dadosaida = guiasaida.getRange(2, 5, guiasaida.getLastRow(),5).getValues();
var totalqtsaida = "0";

for(var i = 0; i<dadosaida.length; i++){
          
     if(dadosaida[i][0] == linha && dadosaida[i][1] == marca && dadosaida[i][2] == produto){
     
       var qtd = parseFloat(dadosaida[i][4]);                  
       totalqtsaida = parseFloat(totalqtsaida) + parseFloat(qtd);
          
    }
}

dadosaida.length = 0;
i = 0;
      
var saldo = parseFloat(totalqtdentrada) - parseFloat(totalqtsaida); 

return  saldo;


}
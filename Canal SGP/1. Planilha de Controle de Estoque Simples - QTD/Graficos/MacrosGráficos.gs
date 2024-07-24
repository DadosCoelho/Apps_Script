function FormGraficos(){
   
var planilha = SpreadsheetApp.getActiveSpreadsheet();
var guialinha = planilha.getSheetByName("Linhas/Marca");

var linha = 1;

while(guialinha.getRange(linha,1).isBlank() == false) {                         
    linha = linha + 1;
};
 
var list = guialinha.getRange(2, 1,linha -1,1).getValues();

list.sort();

var Form = HtmlService.createTemplateFromFile("FormGraficos");

Form.list = list.map(function(r){ return r[0];});
  
var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

MostrarForm.setTitle("GRÁFICO DE ESTOQUE PRODUTO").setHeight(510).setWidth(1200);
  
SpreadsheetApp.getUi().showModalDialog(MostrarForm, "GRÁFICO DE ESTOQUE PRODUTO");

  
}

function Estoque(Dadosfiltro) {    
  
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  
  var guiaEntrada = planilha.getSheetByName("Entradas");
  var guiaSaida = planilha.getSheetByName("Saidas");  
  
  var dadosEntrada = guiaEntrada.getRange(2,1, guiaEntrada.getLastRow(),11).getValues();
  var dadosSaida = guiaSaida.getRange(2,1, guiaSaida.getLastRow(),10).getValues();  

  var Linha = Dadosfiltro.Linha;
  var Marca = Dadosfiltro.Marca;
  var Produto = Dadosfiltro.Produto;
  var Ano = Dadosfiltro.Ano;
  var Mes = Dadosfiltro.Mes;

  var dadosFiltro = dadosEntrada.filter(function(value, i, arr){

  return (Mes ? Mes == arr[i][2] : true) &&  (Ano ? Ano == arr[i][3] : true) &&  (Linha ? Linha == arr[i][4] : true) &&  (Marca ? Marca == arr[i][5] : true) && (Produto ? Produto == arr[i][6] : true);

});

  var QtdEntrada = 0;
  
  for(var i = 0; i<dadosFiltro.length; i++ ){
      if(dadosFiltro[i][8] != ""){
          QtdEntrada = parseFloat(QtdEntrada) + parseFloat(dadosFiltro[i][8]);  
      }
  }

dadosFiltro.length = 0;

var dadosFiltro = dadosSaida.filter(function(value, i, arr){

  return (Mes ? Mes == arr[i][2] : true) &&  (Ano ? Ano == arr[i][3] : true) &&  (Linha ? Linha == arr[i][4] : true) &&  (Marca ? Marca == arr[i][5] : true) && (Produto ? Produto == arr[i][6] : true);

});

  var QtdSaida = 0;
  
  for(var i = 0; i<dadosFiltro.length; i++ ){
      if(dadosFiltro[i][8] != ""){
          QtdSaida = parseFloat(QtdSaida) + parseFloat(dadosFiltro[i][8]);  
      }
  }

  dadosFiltro.length = 0;
  dadosSaida.length = 0;
  dadosEntrada.length = 0;
        
  var Saldo = parseFloat(QtdEntrada) - parseFloat(QtdSaida); 
  
  var dados = new Array(["ESTOQUE", "QTD."],["Entrada", QtdEntrada], ["Saida", QtdSaida], ["Saldo",Saldo]);
  
  return dados;
 
}



function EstoqueMes(Dadosfiltro) {

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaEntrada = planilha.getSheetByName("Entradas");
  var guiaSaida = planilha.getSheetByName("Saidas"); 
  
  var dadosEntrada = guiaEntrada.getRange(2,1, guiaEntrada.getLastRow(),11).getValues();
  var dadosSaida = guiaSaida.getRange(2,1, guiaSaida.getLastRow(),10).getValues();  

  var Linha = Dadosfiltro.Linha;
  var Marca = Dadosfiltro.Marca;
  var Produto = Dadosfiltro.Produto;
  var Ano = Dadosfiltro.Ano;  
  var Mes = Dadosfiltro.Mes;

  var listameses = new Array(["JANEIRO"], ["FEVEREIRO"], ["MARÇO"], ["ABRIL"], ["MAIO"],
  ["JUNHO"], ["JULHO"], ["AGOSTO"], ["SETEMBRO"], ["OUTUBRO"], ["NOVEMBRO"], ["DEZEMBRO"]);

  var dadosFiltro = dadosEntrada.filter(function(value, i, arr){

  return (Mes ? Mes == arr[i][2] : true) && (Ano ? Ano == arr[i][3] : true) &&  (Linha ? Linha == arr[i][4] : true) &&  (Marca ? Marca == arr[i][5] : true) && (Produto ? Produto == arr[i][6] : true);

});

var QtdEntrada = 0;
    
for(var linha = 0; linha< listameses.length; linha++){
    
      var MesLista = listameses[linha][0]; 
         
      for(var i = 0; i<dadosFiltro.length; i++ ){
          if(dadosFiltro[i][8] != "" && dadosFiltro[i][2] == MesLista){
              QtdEntrada = parseFloat(QtdEntrada) + parseFloat(dadosFiltro[i][8]);  
          }
      }

if(listameses[linha][0] == "JANEIRO"){
    var Jane = QtdEntrada;
};

if(listameses[linha][0] == "FEVEREIRO"){
  var Feve = QtdEntrada;
};

if(listameses[linha][0] == "MARÇO"){
  var Mare = QtdEntrada;
};

if(listameses[linha][0] == "ABRIL"){
  var Abre = QtdEntrada;
};

if(listameses[linha][0] == "MAIO"){
  var Maie = QtdEntrada;
};

if(listameses[linha][0] == "JUNHO"){
  var June = QtdEntrada;
};

if(listameses[linha][0] == "JULHO"){
  var Jule = QtdEntrada;
};

if(listameses[linha][0] == "AGOSTO"){
  var Agoe = QtdEntrada;
};

if(listameses[linha][0] == "SETEMBRO"){
  var Sete = QtdEntrada;
};

if(listameses[linha][0] == "OUTUBRO"){
  var Oute = QtdEntrada;
};

if(listameses[linha][0] == "NOVEMBRO"){
  var Nove = QtdEntrada;
};

if(listameses[linha][0] == "DEZEMBRO"){
  var Deze = QtdEntrada;
};

QtdEntrada = 0;

};
 
dadosEntrada.length = 0;
dadosFiltro.length = 0;

var dadosFiltro = dadosSaida.filter(function(value, i, arr){

  return (Mes ? Mes == arr[i][2] : true) && (Ano ? Ano == arr[i][3] : true) &&  (Linha ? Linha == arr[i][4] : true) &&  (Marca ? Marca == arr[i][5] : true) && (Produto ? Produto == arr[i][6] : true);

});

var QtdSaida = 0;

 for(var linha = 0; linha< listameses.length; linha++){
    
    var MesLista = listameses[linha][0]; 
        
    for(var i = 0; i<dadosFiltro.length; i++ ){
        if(dadosFiltro[i][8] != "" && dadosFiltro[i][2] == MesLista){
            QtdSaida = parseFloat(QtdSaida) + parseFloat(dadosFiltro[i][8]);  
        }
    }

    if(listameses[linha][0] == "JANEIRO"){
      var Jans = QtdSaida;           
    };

    if(listameses[linha][0] == "FEVEREIRO"){
      var Fevs = QtdSaida;
    };
    
    if(listameses[linha][0] == "MARÇO"){
      var Mars = QtdSaida;
    };

    if(listameses[linha][0] == "ABRIL"){
      var Abrs = QtdSaida;
    };   
  
    if(listameses[linha][0] == "MAIO"){
      var Mais = QtdSaida;                
    };
    
    if(listameses[linha][0] == "JUNHO"){
      var Juns = QtdSaida;
    };

  if(listameses[linha][0] == "JULHO"){
    var Juls = QtdSaida;          
  };

  if(listameses[linha][0] == "AGOSTO"){
    var Agos = QtdSaida;
  };

  if(listameses[linha][0] == "SETEMBRO"){
    var Sets = QtdSaida;
  };

  if(listameses[linha][0] == "OUTUBRO"){
    var Outs = QtdSaida;
  };

  if(listameses[linha][0] == "NOVEMBRO"){
    var Novs = QtdSaida;
  };

  if(listameses[linha][0] == "DEZEMBRO"){
    var Dezs = QtdSaida;
  };

  QtdSaida = 0;    

}      
   
dadosSaida.length = 0;
dadosFiltro.length = 0;
listameses.length = 0;    

var dados = new Array(["MESES", "ENTRADA", "SAÍDA", "SALDO"],["JAN.",Jane, Jans,parseFloat(Jane) - parseFloat(Jans)],
  ["FEV.",Feve, Fevs,parseFloat(Feve) - parseFloat(Fevs)],["MAR.",Mare, Mars,parseFloat(Mare) - parseFloat(Mars)],
  ["ABR.",Abre, Abrs,parseFloat(Abre) - parseFloat(Abrs)],["MAI.",Maie, Mais,parseFloat(Maie) - parseFloat(Mais)], 
  ["JUN.",June, Juns,parseFloat(June) - parseFloat(Juns)],["JUL.",Jule, Juls,parseFloat(Jule) - parseFloat(Juls)],
  ["AGO.",Agoe, Agos,parseFloat(Agoe) - parseFloat(Agos)],["SET.",Sete, Sets,parseFloat(Sete) - parseFloat(Sets)],
  ["OUT.",Oute, Outs,parseFloat(Oute) - parseFloat(Outs)],["NOV.",Nove, Novs,parseFloat(Nove) - parseFloat(Novs)],
  ["DEZ.",Deze, Dezs,parseFloat(Deze) - parseFloat(Dezs)]); 

  return dados;
  
}



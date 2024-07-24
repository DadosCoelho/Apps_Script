function FormRelatorios(){

var planilha = SpreadsheetApp.getActiveSpreadsheet(); 
var guiaProduto = planilha.getSheetByName("Produtos");

var dados = guiaProduto.getRange(2,1, guiaProduto.getLastRow(),1).getValues();

var b = {};

for (var i = 0; i < dados.length; i++) {
    b[dados[i][0]] = dados[i][0];
}

var criterio1 = [];
  
for (var key in b) {
    criterio1.push([key]);
}

dados.length = 0;

var list1 = criterio1;

list1.sort();

var Form = HtmlService.createTemplateFromFile("Relatorio");

Form.list1 = list1.map(function(r){ return r[0];});

var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

MostrarForm.setTitle("RELATÓRIO DE ESTOQUE").setHeight(510).setWidth(1100);
  
SpreadsheetApp.getUi().showModalDialog(MostrarForm, "RELATÓRIO DE ESTOQUE");

}


function Rel2(dadosRelatorio){

var planilha = SpreadsheetApp.getActiveSpreadsheet();

var guiaEntrada = planilha.getSheetByName("Entradas");
var guiaSaida = planilha.getSheetByName("Saidas"); 
var guiaProduto = planilha.getSheetByName("Produtos"); 
var guiarel = planilha.getSheetByName("Relatório"); 
    
var listaprodutos = guiaProduto.getRange(2, 1, guiaProduto.getLastRow()- 1, 6).getValues();

var dadosEntrada = guiaEntrada.getRange(2,1, guiaEntrada.getLastRow(),11).getValues();

var Ano = dadosRelatorio.Ano;
var Mes = dadosRelatorio.Mes;
var Linha = dadosRelatorio.Linha;
var Marca = dadosRelatorio.Marca;
var Produto = dadosRelatorio.Produto;

var dadosFiltro = dadosEntrada.filter(function(value, i, arr){

return (Mes ? Mes == arr[i][2] : true) &&  (Ano ? Ano == arr[i][3] : true) &&  (Linha ? Linha == arr[i][4] : true) &&  (Marca ? Marca == arr[i][5] : true) && (Produto ? Produto == arr[i][6] : true);

});
  
  var dadosrel = [] ;
  var dadostabela = [] ;
  
      for(var l = 0; l< listaprodutos.length; l++){
    
        var lproduto = listaprodutos[l][0];
        var lmarca = listaprodutos[l][1];
        var nomeproduto = listaprodutos[l][2];
 
        var TE = parseFloat("0");

        for(var i = 0; i<  dadosFiltro.length; i++ ){         
            
               if(dadosFiltro[i][4] == lproduto && dadosFiltro[i][5] == lmarca && dadosFiltro[i][6] == nomeproduto){

                if(dadosFiltro[i][8] != ""){                                 
                  TE = parseFloat(TE) + parseFloat(dadosFiltro[i][8]);
                }

              }        
 
        }

      if (TE > 0){        
        listaprodutos[l][3] = TE;
        dadosrel.push(listaprodutos[l]);         
       }
     
    }
    
  
  if (dadosrel.length == "0"){
    return "NÃO EXISTEM DADOS PARA ESTE FILTRO";  
  }  

  listaprodutos.length = 0;
  dadosEntrada.length = 0;
  dadosFiltro.length = 0;

  var dadoSaida = guiaSaida.getRange(2,1, guiaSaida.getLastRow(),10).getValues();  

  var dadosFiltro = dadoSaida.filter(function(value, i, arr){

      return (Mes ? Mes == arr[i][2] : true) &&  (Ano ? Ano == arr[i][3] : true) &&  (Linha ? Linha == arr[i][4] : true) &&  (Marca ? Marca == arr[i][5] : true) && (Produto ? Produto == arr[i][6] : true);

  });

  for(var l = 0; l< dadosrel.length; l++){        
 
        var lproduto = dadosrel[l][0];
        var lmarca = dadosrel[l][1];
        var nomeproduto = dadosrel[l][2];
        
        var TS = parseFloat("0");

        for(var i = 0; i<  dadosFiltro.length; i++ ){
             
              if(dadosFiltro[i][4] == lproduto && dadosFiltro[i][5] == lmarca && dadosFiltro[i][6] == nomeproduto){

                if(dadosFiltro[i][8] != ""){                                 
                  TS = parseFloat(TS) + parseFloat(dadosFiltro[i][8]);              
                }

              }          
          
        }
 
        dadosrel[l][4] = TS;
        dadosrel[l][5] = parseFloat(dadosrel[l][3]) - parseFloat(TS);
         
        dadostabela.push(dadosrel[l]);        
     
    }
   
   guiarel.getRange("A2:F").clear();
   guiarel.getRange(2, 1, dadostabela.length, dadostabela[0].length).setValues(dadostabela);
  
   guiarel.getRange("A2:F").sort([{column: 1, ascending: true}, {column: 4, ascending: false}]);

  var dados = guiarel.getRange(2,1, guiarel.getLastRow() - 1, 6).getValues();

   listaprodutos.length = 0;
   dadoSaida.length = 0;
   dadosrel.length = 0;
   dadostabela.length = 0;
  
  return dados;

}






var PRINT_OPTIONS = {
  'size': 7,               // Tamanho do papel. 0 = carta, 1 = tablóide, 2 = Ofício, 3 = declaração, 4 = executivo, 5 = fólio, 6 = A3, 7 = A4, 8 = A5, 9 = B4, 10 = B
  'fzr': false,            // repetir cabeçalhos de linha
  'portrait': false,        // falso = paisagem
  'fitw': true,            // ajustar a janela ou tamanho real
  'gridlines': false,      // mostrar linhas de grade
  'printtitle': false,
  'sheetnames': false,
  'pagenum': 'UNDEFINED',  // CENTRO = mostrar números de página / UNDEFINED = não mostrar
  'attachment': false
}

var PDF_OPTS = objectToQueryString(PRINT_OPTIONS);

function Pdf() {

  SpreadsheetApp.flush();
  
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guia = planilha.getSheetByName("Relatório");
  var guiamenu = planilha.getSheetByName("Menu");
  
  guia.showSheet().activate();
    
  var linha = guia.getLastRow();
  var area = "A1:F" + linha;
  var range = guia.getRange(area).activate();
    
  var gid = guia.getSheetId();
  
  var printRange = objectToQueryString({
    'c1': range.getColumn() - 1,
    'r1': range.getRow() - 1,
    'c2': range.getColumn() + range.getWidth() - 1,
    'r2': range.getRow() + range.getHeight() - 1
  });
  
  var url = planilha.getUrl().replace(/edit$/, '') + 'export?format=pdf' + PDF_OPTS + printRange + "&gid=" + gid;

  var htmlTemplate = HtmlService.createTemplateFromFile('Pdf');
  htmlTemplate.url = url;
  SpreadsheetApp.getUi().showModalDialog(htmlTemplate.evaluate().setHeight(10).setWidth(100), 'Gerando PDF');
   
  guiamenu.activate();
     
  
}



function objectToQueryString(obj) {
  return Object.keys(obj).map(function(key) {
    return Utilities.formatString('&%s=%s', key, obj[key]);
  }).join('');
  
  
}


function relPlan(){
  
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  
  var guiaentrada = planilha.getSheetByName("Entradas");
  var guiasaida = planilha.getSheetByName("Saidas"); 
  var guiaproduto = planilha.getSheetByName("Produtos"); 
  var guiarel = planilha.getSheetByName("Relatório"); 
    
  var listaprodutos = guiaproduto.getRange(2, 1, guiaproduto.getLastRow()- 1, 6).getValues();
  var ultimalinha = guiaentrada.getLastRow();
  var range = "A2:I" + ultimalinha;
    
  var area = guiaentrada.getRange(range);
  var dadosentrada = area.getValues();
  
  var dadosrel = [] ;
  var dadostabela = [] ;
  
      for(var l = 0; l< listaprodutos.length; l++){
    
        var lproduto = listaprodutos[l][0];
        var lmarca = listaprodutos[l][1];
         var nomeproduto = listaprodutos[l][2];
 
        var totalqtdentrada = parseFloat("0");

        for(var i = 0; i<  dadosentrada.length; i++ ){         
                  
            
               if(dadosentrada[i][4] == lproduto && dadosentrada[i][5] == lmarca && dadosentrada[i][6] == nomeproduto){

                     if(dadosentrada[i][8] != ""){
                     var qtd = parseFloat(dadosentrada[i][8]);             
                     totalqtdentrada = parseFloat(totalqtdentrada) + parseFloat(qtd);              
                    }

              } 
                 
        }
        
        listaprodutos[l][3] = totalqtdentrada;
        dadosrel.push(listaprodutos[l]);         
     
    }
    
  
  if (dadosrel.length == "0"){
  return "Não existem dados para este filtro";
  
  }  

   listaprodutos.length = 0;
   dadosentrada.length = 0;
   
  var ultimalinha = guiasaida.getLastRow();
  var range = "A2:I" + ultimalinha;
    
  var area = guiasaida.getRange(range);
  var dadossaida = area.getValues();  

       for(var l = 0; l< dadosrel.length; l++){        
 
        var lproduto = dadosrel[l][0];
        var lmarca = dadosrel[l][1];
        var nomeproduto = dadosrel[l][2];
        
        var totalqtdsaida = parseFloat("0");

        for(var i = 0; i<  dadossaida.length; i++ ){  
             
               if(dadossaida[i][4] == lproduto && dadossaida[i][5] == lmarca && dadossaida[i][6] == nomeproduto){

                 if(dadossaida[i][8] != ""){
                     var qtd = parseFloat(dadossaida[i][8]);             
                      totalqtdsaida = parseFloat(totalqtdsaida) + parseFloat(qtd);              
                   }

               }
          
        }

 
        dadosrel[l][4] = totalqtdsaida;
        dadosrel[l][5] = parseFloat(dadosrel[l][3]) - parseFloat(totalqtdsaida);
         
        dadostabela.push(dadosrel[l]);        
     
    }
   
   guiarel.getRange("A2:F").clear();
   guiarel.getRange(2, 1, dadostabela.length, dadostabela[0].length).setValues(dadostabela);  
  
   guiarel.getRange("A2:F").sort([{column: 1, ascending: true}, {column: 4, ascending: false}]);  

   listaprodutos.length = 0;
   dadossaida.length = 0;
   dadosrel.length = 0;
   dadostabela.length = 0;   

}

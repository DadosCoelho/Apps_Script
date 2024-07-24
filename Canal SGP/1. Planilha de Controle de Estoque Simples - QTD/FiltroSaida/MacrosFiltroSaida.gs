function FormFiltroSaida(Data,Linha,Marca,Produto,Cod){

var planilha = SpreadsheetApp.getActiveSpreadsheet();
var guiadados = planilha.getSheetByName("Saidas"); 
var dados = guiadados.getRange(2,5, guiadados.getLastRow(),3).getValues();

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

var Form = HtmlService.createTemplateFromFile("FormFiltroSaida");

Form.list1 = list1.map(function(r){ return r[0];});
Form.Data = Data;
Form.Linha = Linha;
Form.Marca = Marca;
Form.Produto = Produto;
Form.Cod = Cod;

var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
  
MostrarForm.setTitle("FORMULÁRIO").setHeight(600).setWidth(1100);
  
SpreadsheetApp.getUi().showModalDialog(MostrarForm, "FILTRO SAÍDA ESTOQUE"); 
  
}


function FiltroSaida(dadosRelatorio) {
  
var planilha = SpreadsheetApp.getActiveSpreadsheet();
var guiadados = planilha.getSheetByName("Saidas"); 
var dados = guiadados.getRange(1,1, guiadados.getLastRow(),10).getValues(); 

var Data1 = Utilities.formatDate(new Date(dadosRelatorio.data1), planilha.getSpreadsheetTimeZone(),"dd/MM/yyyy");

var Data2 = Utilities.formatDate(new Date(dadosRelatorio.data2), planilha.getSpreadsheetTimeZone(),"dd/MM/yyyy");

var DataInicial = converteData(Data1);
var DataFinal = converteData(Data2);

var Linha = dadosRelatorio.Linha;
var Marca = dadosRelatorio.Marca;
var Produto = dadosRelatorio.Produto;
var Cod = dadosRelatorio.Cod;

var dadosfiltro = dados.filter(function(value, i, arr){

var Data = Utilities.formatDate(new Date(arr[i][1]), planilha.getSpreadsheetTimeZone(),"dd/MM/yyyy");

return converteData(Data) >= DataInicial && converteData(Data) <= DataFinal &&  (Linha ? Linha == arr[i][4] : true) &&  (Marca ? Marca == arr[i][5] : true) &&  (Produto ? Produto == arr[i][6] : true) &&  (Cod ? Cod == arr[i][7] : true);

});

if(dadosfiltro.length == "0"){
  return "NÃO EXISTEM DADOS PARA ESTE FILTRO";
}

dados.length = 0;

 for(var i = 0; i < dadosfiltro.length; i++){          
        
  var data = Utilities.formatDate(new Date(dadosfiltro[i][1]), planilha.getSpreadsheetTimeZone(),"dd/MM/yyyy");              

  dadosfiltro[i][1] = data;

}

return dadosfiltro;

}

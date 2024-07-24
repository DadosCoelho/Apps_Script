function onOpen(){
 SpreadsheetApp.getUi()
 .createMenu('Formulários')
 .addItem('Entrada', 'FormEntrada')
 .addItem('Saída', 'FormSaida')
 .addItem('Filtro Entrada', 'FormFiltroEntrada') 
 .addItem('Filtro Saída', 'FormFiltroSaida')
 .addItem('Gráficos', 'FormGraficos')
 .addItem('Relatórios', 'FormRelatorios')
 .addItem('Nova Linha', 'FormLinha')
 .addItem('Nova Marca', 'FormMarca')
 .addItem('Novo Produto', 'FormProduto')
 .addToUi();
}

function FormEntrada(Id){
   
var planilha = SpreadsheetApp.getActiveSpreadsheet();
var guialinha = planilha.getSheetByName("Linhas/Marca");

var linha = 1; 

 while(guialinha.getRange(linha,1).isBlank() == false) {                         
     linha = linha + 1;
 };
 
if (linha < 3){
  linha = 3;
}

var list = guialinha.getRange(2, 1,linha -2,1).getValues();

list.sort();

var Form = HtmlService.createTemplateFromFile("FormEntrada");

Form.list = list.map(function(r){ return r[0];});
Form.Id = Id;

var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

MostrarForm.setTitle("ENTRADA ESTOQUE").setHeight(335).setWidth(650);
  
SpreadsheetApp.getUi().showModalDialog(MostrarForm, "ENTRADA ESTOQUE");

  
}


function Chamar(Arquivo){
  return HtmlService.createHtmlOutputFromFile(Arquivo).getContent();
}



function Marca(linha){

  
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guialinha = planilha.getSheetByName("Linhas/Marca");
  var localPesquisa = guialinha.getRange(1, 2, 1, guialinha.getLastColumn()).getValues()[0];
  
  var resultado = localPesquisa.Pesquisa(linha);
      
  if (resultado !=-1){
  
   var coluna = resultado + 2;
   
           var qtdlinha = 1; 
           while(guialinha.getRange(qtdlinha,coluna).isBlank() == false) {                         
           qtdlinha = qtdlinha + 1;
            };
         
         var qtdlinha = qtdlinha - 1;
         
         var dados = guialinha.getRange(2, coluna, qtdlinha).getValues();         
         
         dados.sort();

        return dados;
 
   
  }else{
  
  return "LINHA NÃO ENCONTRADA"
  
  }

}


function Produtos(marca){
  
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiamarca = planilha.getSheetByName("Marca/Produto");
  var localPesquisa = guiamarca.getRange(1, 2, 1, guiamarca.getLastColumn()).getValues()[0];
  
  var resultado = localPesquisa.Pesquisa(marca);
  
  if (resultado !=-1){
  
         var coluna = resultado + 2;
   
           var qtdlinha = 1; 
           while(guiamarca.getRange(qtdlinha,coluna).isBlank() == false) {                         
           qtdlinha = qtdlinha + 1;
            };
         
         var qtdlinha = qtdlinha - 1;
         
         var dados = guiamarca.getRange(2, coluna, qtdlinha).getValues();         
        
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




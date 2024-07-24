function FormLinha(){
   
var Form = HtmlService.createTemplateFromFile("FormLinha");

var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

MostrarForm.setTitle("CADASTRO DE NOVAS LINHAS").setHeight(190).setWidth(405);
  
SpreadsheetApp.getUi().showModalDialog(MostrarForm, "CADASTRO DE NOVAS LINHAS");
  
}

function SalvarLinha(Linha){

const user = LockService.getScriptLock();
user.tryLock(10000);

if(user.hasLock()){

var planilha = SpreadsheetApp.getActiveSpreadsheet();
var guialinha = planilha.getSheetByName("Linhas/Marca");

var novalinha = Linha;
var novalinha  = novalinha.toUpperCase(); 
  
      var linha = 1;
      
      while(guialinha.getRange(linha,1).isBlank() == false) {

          linha = linha +1;
          var nomelinha =  guialinha.getRange(linha,1).getValue();
          
            if(novalinha == nomelinha){             
              return "LINHA JÁ CADASTRADA!";             
            };
            
        }
      
       guialinha.getRange(linha, 1).setValue(novalinha);         

          var coluna = 1; 
          while(guialinha.getRange(1,coluna).isBlank() == false){
              coluna = coluna +1;
           }
                 
         guialinha.getRange(1, coluna).setValue(novalinha);
         
         return "REGISTRADO COM SUCESSO!";
    }

}

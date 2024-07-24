function FormMarca(){
   
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

var Form = HtmlService.createTemplateFromFile("FormMarca");

Form.list = list.map(function(r){ return r[0];});
  
var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

MostrarForm.setTitle("NOVA MARCA").setHeight(200).setWidth(400);
  
SpreadsheetApp.getUi().showModalDialog(MostrarForm, "NOVA MARCA");

  
}


function SalvarMarca(Dados){

const user = LockService.getScriptLock();
user.tryLock(10000);

if(user.hasLock()){

var planilha = SpreadsheetApp.getActiveSpreadsheet();
var guiamarca = planilha.getSheetByName("Marca/Produto");
var guialinha = planilha.getSheetByName("Linhas/Marca");

var novamarca = Dados.Marca;
var linhamarca = Dados.Linha;
  
        var linha = 1; 
        while(guiamarca.getRange(linha,1).isBlank() == false) {

              linha = linha +1;
            var nomemarca =  guiamarca.getRange(linha,1).getValue();
            
             if(novamarca == nomemarca){
                
               var coluna = 1; 
               while(guialinha.getRange(1,coluna).isBlank() == false) {

              coluna = coluna + 1;
              
              var nomelinha =  guialinha.getRange(1,coluna).getValue();
            
               if(nomelinha == linhamarca){             
 
                 var linha = 1;

                 while(guialinha.getRange(linha,coluna).isBlank() == false){
                         linha = linha + 1; 
                         
                         var marca = guialinha.getRange(linha,coluna).getValue();
                         
                         if(marca == novamarca){
                            return "MARCA JÁ CADASTRADA!";
                         }
                  };

                 guialinha.getRange(linha, coluna).setValue(novamarca); 
                 return "REGISTRADO COM SUCESSO!";
                        
                 };

               }
             
            };
              
         }         
              
          guiamarca.getRange(linha, 1).setValue(novamarca);         

          var coluna = 1;

          while(guiamarca.getRange(1,coluna).isBlank() == false){
            coluna = coluna + 1;
          }
           
          guiamarca.getRange(1, coluna).setValue(novamarca);
         
            var coluna = 1;

            while(guialinha.getRange(1,coluna).isBlank() == false){

              coluna = coluna + 1;
              
              var nomelinha =  guialinha.getRange(1,coluna).getValue();
            
             if(nomelinha == linhamarca){             
 
                 var linha = 1; 
                 while(guialinha.getRange(linha,coluna).isBlank() == false) {
                    linha = linha + 1;                        
                  };
                  
                 guialinha.getRange(linha, coluna).setValue(novamarca);
                 return "REGISTRADO COM SUCESSO!";
                        
             };

       }
   }     
}



function FormProduto(){
   
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

var Form = HtmlService.createTemplateFromFile("FormProduto");

Form.list = list.map(function(r){ return r[0];});
  
var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

MostrarForm.setTitle("NOVO PRODUTO").setHeight(200).setWidth(400);
  
SpreadsheetApp.getUi().showModalDialog(MostrarForm, "NOVO PRODUTO");

  
}

function SalvarProduto(Dados){

const user = LockService.getScriptLock();
user.tryLock(10000);

if(user.hasLock()){

var planilha = SpreadsheetApp.getActiveSpreadsheet();
var guiamarcaproduto = planilha.getSheetByName("Marca/Produto");

var novoproduto = Dados.Produto;
var marca = Dados.Marca;
var nomelinha = Dados.Linha;
          
            var coluna = 1;

            while(guiamarcaproduto.getRange(1,coluna).isBlank() == false){

              coluna = coluna + 1;
              
              var nomemarca =  guiamarcaproduto.getRange(1,coluna).getValue();
            
             if(marca == nomemarca){             
 
                 var linha = 1;

                 while(guiamarcaproduto.getRange(linha,coluna).isBlank() == false){
                      linha = linha + 1;
                      
                      var nomeproduto = guiamarcaproduto.getRange(linha,coluna).getValue(); 
                      if(nomeproduto == novoproduto){
                          return "PRODUTO JÁ CADASTRADO!";
                      }
                  };

                 guiamarcaproduto.getRange(linha, coluna).setValue(novoproduto); 
                 
                 var guiaprodutos =  planilha.getSheetByName("Produtos");  
                 
                 var linha = 1;

                 while(guiaprodutos.getRange(linha,1).isBlank() == false){
                    linha = linha + 1;       
                  };

                 guiaprodutos.getRange(linha, 1).setValue(nomelinha); 
                 guiaprodutos.getRange(linha, 2).setValue(marca); 
                 guiaprodutos.getRange(linha, 3).setValue(novoproduto);

                 return "REGISTRADO COM SUCESSO!";     
             };

       }

       return "MARCA NÃO ENCONTRADA!"; 

    }     
}

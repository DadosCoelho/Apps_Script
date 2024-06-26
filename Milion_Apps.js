function importDataF() {  
    var spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1CMzEb1fJLUR2L6knJ-J9xuRTbANC-qWwsgX_H0KN_Ok/edit?resourcekey#gid=625349933").getSheetByName('Página1');
  
  
    var formUrl = "https://docs.google.com/forms/d/1YjUp2_CD_1IrtqnBoQWB8Ru2SbUwdoLK1J77Ob5jum4/edit";
    var form = FormApp.openByUrl(formUrl);
  
    //ContaTop' Essa é a pergunta 1
    var Pergunta1 = spreadsheet.getRange('A1').getValue();
    var range1 = spreadsheet.getRange("A2:A1000");
    var choices1 = range1.getValues().map(function(row) {return row[0]});
    //DadosCoelho' Remover opções duplicadas.
    var uniqueChoices1 = [];
    choices1.forEach(function(choice) {
      if (uniqueChoices1.indexOf(choice) === -1) {
        uniqueChoices1.push(choice);
      }
    });
    var existingQuestion1 = form.getItems(FormApp.ItemType.LIST).filter(function(item) {
      return item.getTitle() === Pergunta1;
    });
    if (existingQuestion1.length > 0) {
      //DadosCoelho' Se a pergunta já existir, atualize-a.
      existingQuestion1[0].asListItem().setChoiceValues(uniqueChoices1);
    } else {
      //DadosCoelho' Caso contrário, adicione uma nova pergunta.
      var question1 = form.addListItem();
      question1.setTitle(Pergunta1);
      question1.setChoiceValues(uniqueChoices1);
    }
  
  
    //ContaTop' Essa é a pergunta 2
    var Pergunta2 = spreadsheet.getRange('B1').getValue();
    var range2 = spreadsheet.getRange("B2:B1000");
    var choices2 = range2.getValues().map(function(row) {return row[0]});
    //DadosCoelho' Remover opções duplicadas.
    var uniqueChoices2 = [];
    choices2.forEach(function(choice) {
      if (uniqueChoices2.indexOf(choice) === -1) {
        uniqueChoices2.push(choice);
      }
    });
    var existingQuestion2 = form.getItems(FormApp.ItemType.LIST).filter(function(item) {
      return item.getTitle() === Pergunta2;
    });
    if (existingQuestion2.length > 0) {
      //DadosCoelho' Se a pergunta já existir, atualize-a.
      existingQuestion2[0].asListItem().setChoiceValues(uniqueChoices2);
    } else {
      //DadosCoelho' Caso contrário, adicione uma nova pergunta.
      var question2 = form.addListItem();
      question2.setTitle(Pergunta2);
    }
  

    //ContaTop' Essa é a pergunta 3
    var Pergunta3 = spreadsheet.getRange('C1').getValue();
    var range3 = spreadsheet.getRange("C2:C1000");
    var choices3 = range3.getValues().map(function(row) {return row[0]});
    //ContaTop' Remover opções duplicadas
    var uniqueChoices3 = [];
    choices3.forEach(function(choice) {
      if (uniqueChoices3.indexOf(choice) === -1) {
        uniqueChoices3.push(choice);
      }
    });
    var existingQuestion3 = form.getItems(FormApp.ItemType.LIST).filter(function(item) {
      return item.getTitle() === Pergunta3;
    });
    if (existingQuestion3.length > 0) {
      //ContaTop' Se a pergunta já existir, atualize-a
      existingQuestion3[0].asListItem().setChoiceValues(uniqueChoices3);
    } else {
      //ContaTop' Caso contrário, adicione uma nova pergunta
      var question3 = form.addListItem();
      question3.setTitle(Pergunta3);
    }

  
        //DadosCoelho' Esse grupo de comandos é para excluir perguntas diferentes das perguntas que quero que apareça. 
    var items = form.getItems();
    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var title = item.getTitle();
      if (title !== Pergunta1 && title !== Pergunta2 && title !== Pergunta3) {
        //DadosCoelho' Se o título do item for diferente do título fornecido, remova-o do formulário.
        form.deleteItem(item);
      }
    }
  
    
  }
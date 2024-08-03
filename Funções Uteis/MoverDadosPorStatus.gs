//Para essa função é nescessario ter duas paginas "Tarefas Ativas" e "Tarefas Concluídas" com as mesmas colunas "ID", "Descrição" e "Status", a Coluna "Status" tem que se de lista suspensa com uma das opções "Concluída". 


function onEdit(e) {
  const celulaModificar = e.range;
  const pagina = e.range.getSheet();

  if (pagina.getName() === "Tarefas Ativas" //Checa se foi na página de Tarefas Ativas
  && celulaModificar.getColumn() === 3      //Checa dse foi na coluna Status
  && e.value === "Concluída"                //Checa de o valor do Status é Concluído
  ) {
      var paginaConcluidas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tarefas Concluídas");
      var numeroLinha = celulaModificar.getRow();
      var linhaModificada = pagina.getRange(numeroLinha, 1, 1, 3);
      var linhaDestino = paginaConcluidas.getRange(paginaConcluidas.getLastRow() + 1, 1, 1, 3);

      //Copia pra página Tarefas Concluídas
      linhaModificada.copyTo(linhaDestino, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

      //Remove a tarefa da página "Tarefas Ativas"
      pagina.deleteRow(numeroLinha);
    }  
}

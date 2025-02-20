//Consolidação dos Dados
function consolidarDados() {
  // Abre a planilha ativa
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Obtém todas as abas da planilha
  const sheets = spreadsheet.getSheets();
  
  // Define as abas que devem ser ignoradas
  const abasIgnoradas = ["ConsPsico", "Capa", "Lotes", "NÃO FATURADAS", "Apoio"];
  
  // Valida se aba ConsPsico existe
  if (!spreadsheet.getSheetByName("ConsPsico")) {
    // Cria aba ConsPsico
    spreadsheet.insertSheet("ConsPsico");
  } else {
    // Exibe uma caixa de diálogo para confirmar a limpeza dos dados
    const ui = SpreadsheetApp.getUi();
    const resposta = ui.alert(
      "Atenção!",
      "Os dados da aba ConsPsico serão limpos. Deseja continuar?",
      ui.ButtonSet.YES_NO
    );
    // Verifica a resposta do usuário
    if (resposta !== ui.Button.YES) {
      ui.alert("Operação cancelada pelo usuário.");
      return; // Interrompe a execução do script
    }
    // Limpa os dados anteriores na aba ConsPsico (exceto o cabeçalho)
    spreadsheet.getSheetByName("ConsPsico").getRange("A5:Z").clearContent();
  }

  const abaConsPsico = spreadsheet.getSheetByName("ConsPsico");

  // Define o cabeçalho na aba ConsPsico (com a nova coluna "PLANO")
  const cabecalho = ["PROFISSIONAL", "PLANO", "PACIENTE", "DATAS ATEN.","VALOR LÍQUIDO", "VALOR GLOSA", "MOTIVO GLOSA"];
  abaConsPsico.getRange(4, 1, 1, cabecalho.length).setValues([cabecalho]);
  
  // Inicializa um array para armazenar os dados consolidados
  let dadosConsolidados = [];
  
  // Percorre todas as abas
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    
    // Ignora as abas especificadas
    if (abasIgnoradas.includes(sheetName)) {
      return;
    }
    
    // Obtém os dados da aba atual (a partir da linha 8, considerando o cabeçalho na linha 7)
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    // Filtra os dados, ignorando o cabeçalho
    const dados = values.slice(7); // Ignora as primeiras 7 linhas (cabeçalho)
    
    // Adiciona os dados ao array consolidado, incluindo o nome da aba como "PLANO"
    dados.forEach(row => {
      if (row.join("") !== "") { // Ignora linhas vazias
        // Mantém apenas as colunas desejadas (ignorando índices 6, 9 e 10)
        const linhaFiltrada = [
          row[0], // PROFISSIONAL
          sheetName, // PLANO
          row[1], // PACIENTE
          row[2], // DATAS ATEN.
          row[6], // VALOR LÍQUIDO
          row[7], // VALOR GLOSA
          row[9]  // MOTIVO GLOSA
        ];
        dadosConsolidados.push(linhaFiltrada);
      }
    });
  });
  
  // Escreve os dados consolidados na aba ConsPsico
  if (dadosConsolidados.length > 0) {
    abaConsPsico.getRange(5, 1, dadosConsolidados.length, dadosConsolidados[0].length).setValues(dadosConsolidados);
  }
  
  // Exibe uma mensagem de conclusão ao usuário
  const ui = SpreadsheetApp.getUi();
  ui.alert("Consolidação concluída com sucesso!");
}
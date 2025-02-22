/* Acompanhamento Despesas Cartão
  Importa todos os cvs da pasta para uma planilha consolidada;
  Trata algumas linhas que devem ser ignoradas;
  Resgistra quais csv já foram lidos;
  Lookerstudio consome os dado da planilha para o acompanhamento dos gastos; 
*/

function importarCSVparaSheets() {
  var folderId = "[ID_Pasta]"; // ID da pasta no Google Drive
  var nomePlanilha = "[Planilha]"; // Nome da planilha no Google Sheets
  var pasta = DriveApp.getFolderById(folderId);
  var arquivos = pasta.getFiles();
  
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abaDados = planilha.getSheetByName(nomePlanilha);
  var abaLog = planilha.getSheetByName("Log");
  
  if (!abaDados) {
    abaDados = planilha.insertSheet(nomePlanilha);
    abaDados.appendRow(["Ano", "Mês", "Data", "Lançamento", "Categoria", "Tipo", "Valor"]);
  }
  
  if (!abaLog) {
    abaLog = planilha.insertSheet("Log");
    abaLog.appendRow(["Arquivo"]);
  }
  
  var arquivosImportados = abaLog.getDataRange().getValues().flat();

  while (arquivos.hasNext()) {
    var arquivo = arquivos.next();
    var nomeArquivo = arquivo.getName();
    var mimeType = arquivo.getMimeType();
    if (mimeType !== MimeType.CSV && mimeType !== "text/comma-separated-values") {
      Logger.log("Ignorando arquivo não CSV: " + nomeArquivo);
      continue;
    }
    
    if (arquivosImportados.includes(nomeArquivo)) {
      Logger.log("Arquivo já importado: " + nomeArquivo);
      continue;
    }
    
    var regex = /fatura-\w+-(\d{4})-(\d{2})\.csv/;
    var match = nomeArquivo.match(regex);
    
    if (!match) {
      Logger.log("Nome do arquivo inválido: " + nomeArquivo);
      continue;
    }
    
    var ano = match[1];
    var mes = match[2];
    var mesAno = Utilities.formatDate(new Date(ano, mes - 1, 1), "GMT-3", "MM/yyyy");
    
    var conteudo = arquivo.getBlob().getDataAsString();
    var linhas = conteudo.split("\n").slice(1); 
    
    for (var i = 0; i < linhas.length; i++) {
      var colunas = linhas[i].split(`","`);
      if (colunas.length < 5) continue;
      
      // Remove aspas extras
      var data = colunas[0].replace(/"/g, "").trim();
      var lancamento = colunas[1].replace(/"/g, "").trim();
      var categoria = colunas[2].replace(/"/g, "").trim();
      var tipo = colunas[3].replace(/"/g, "").trim();
      var valor = colunas[4].replace(/"/g, "").trim();

      //Uber - GIFT CARD
      if (lancamento == "GIFT CARD" && categoria == "OUTROS"){
        categoria = "TRANSPORTE"
      }
      
      //.replace(/\./g, "").replace(",", ".");
      valor = valor.replace("R$ ", "").replace(/\./g, "").replace(",", ".");
      valor = parseFloat(valor);

      // Valida se é o pagamento mês anteior
      if (lancamento != "PAGTO DEBITO AUTOMATICO") {
        abaDados.appendRow([mesAno, data, lancamento, categoria, tipo, valor]);
      }
    }
    
    var agora = Utilities.formatDate(new Date(), "GMT-3", "HH:mm dd/MM/yyyy");
    abaLog.appendRow([nomeArquivo, agora]); 
    Logger.log("Importado: " + nomeArquivo);
  }
}

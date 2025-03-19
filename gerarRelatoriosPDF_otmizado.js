function gerarRelatoriosPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // Record start time
  const startTime = new Date();

  // Get all data at once and destructure headers and rows
  const [headers, ...rows] = sheet.getDataRange().getValues();

  // Create a Map for faster lookups instead of an object
  const profissionaisMap = new Map();

  // Pre-calculate planos totals while grouping by profissional
  const planosTotals = new Map();

  // Single pass through the data
  rows.forEach(row => {
    const [profissional, plano, paciente, data, valor, valorGlosa, motivoGlosa] = row;
    
    // Initialize arrays and maps if they don't exist
    if (!profissionaisMap.has(profissional)) {
      profissionaisMap.set(profissional, []);
      planosTotals.set(profissional, new Map());
    }

    // Add row to profissional's data
    profissionaisMap.get(profissional).push(row);

    // Update plano totals
    if (!isNaN(valor) && valor !== '') {
      const valorGlosaNumber = isNaN(valorGlosa) ? 0 : valorGlosa;
      const valorLiquido = valor - valorGlosaNumber;
      const planosMap = planosTotals.get(profissional);
      planosMap.set(plano, (planosMap.get(plano) || 0) + valorLiquido);
    }
  });
  // Open pop to get name of the file
  const fileName = Browser.inputBox('Digite o nome do arquivo', 'Relatorio', Browser.Buttons.OK_CANCEL);
  // Get the name of the file
  const name = fileName === 'cancel' ? 'Relatorio' : fileName;

  // Create folder with formatted date
  const now = new Date();
  const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd_HH:mm");
  const currentFolder = getCurrentFolder();
  const folder = currentFolder.createFolder('Rel_'+ name + '_' + formattedDate);

  // Get image once instead of in the loop
  const imageUrl = getImageBase64('logoup.jpg');
    
  /*const sheetName = ss.getName();
  const prefix = sheetName.substring(0, 5);*/

  // Prepare HTML style once
  const styleHTML = `
    <style>
      table { border-collapse: collapse; }
      th, td { padding: 3px; background-color: #e1e1e1; font-size:12px; }
      body { font-family: Arial, sans-serif; }
    </style>
  `;

  // Process each profissional
  profissionaisMap.forEach((dados, prof) => {
    const planosMap = planosTotals.get(prof);
    const totalGeral = Array.from(planosMap.values()).reduce((a, b) => a + b, 0);

    // Build planos summary table
    const planosHTML = Array.from(planosMap.entries())
      .map(([plano, valor]) => `
        <tr>
          <td>${plano}</td>
          <td align="right">${valor.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}</td>
        </tr>`
      ).join('');

    // Build details table
    const detailsHTML = dados.map(d => `
      <tr>
        <td>${d[1]}</td>
        <td>${d[2]}</td>
        <td>${d[3] instanceof Date ? Utilities.formatDate(d[3], Session.getScriptTimeZone(), 'dd/MM/yyyy') : d[3]}</td>
        <td align="right">${d[4].toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}</td>
        <td align="right">${d[5].toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}</td>
        <td>${d[6]}</td>
      </tr>`
    ).join('');

    const html = `
      <html>
        <head>${styleHTML}</head>
        <body>
          <h2 align="center">Relatório de ${prof}</h2>
          <h3>Resumo por Plano</h3>
          <table border="1">
            <tr><th>Plano</th><th>Valor Líquido</th></tr>
            ${planosHTML}
            <tr>
              <td><strong>Total</strong></td>
              <td align="right"><strong>${totalGeral.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}</strong></td>
            </tr>
          </table>
          <br><br>
          <h3>Detalhamento</h3>
          <table border="1" width="100%">
            <tr>
              <th>Plano</th>
              <th>Paciente</th>
              <th>Datas Aten.</th>
              <th>Valor Líquido</th>
              <th>Valor Glosa</th>
              <th>Motivo Glosa</th>
            </tr>
            ${detailsHTML}
          </table>
          <img src="${imageUrl}" style="position: absolute; top: 50; right: 0; z-index: -1; opacity: 0.5;" />
          <p align="right" style="font-size: 10px;">Emitido em ${formattedDate}</p>
        </body>
      </html>
    `;

    const blob = HtmlService.createHtmlOutput(html)
      .getAs('application/pdf')
      .setName(`Relatorio_${prof}_${name}.pdf`);
    
    folder.createFile(blob);
  });

  // Record end time and show results
  const endTime = new Date();
  const timeDiff = (endTime - startTime) / 1000;

  //SpreadsheetApp.getUi().alert('Relatórios gerados na pasta: ' + folder.getUrl());
  SpreadsheetApp.getUi().alert("Function executed in: " + timeDiff + " seconds");
}

// Pega diretório atual
function getCurrentFolder() {
  var file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var folder = file.getParents().next();
  return folder;
}

function getImageBase64(fileName) {
  var folder = getCurrentFolder(); // Função que retorna a pasta atual
  var files = folder.getFilesByName(fileName);
  if (files.hasNext()) {
    var file = files.next();
    var blob = file.getBlob();
    var base64 = Utilities.base64Encode(blob.getBytes());
    var contentType = blob.getContentType();
    return 'data:' + contentType + ';base64,' + base64;
  }
  return null;
}

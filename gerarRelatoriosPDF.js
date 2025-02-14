// Google Apps Script para gerar relatórios em PDF por profissional
function gerarRelatoriosPDF() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();

  var headers = data[0];
  var rows = data.slice(1);

  var profissionais = {};

  // Agrupar dados por profissional
  rows.forEach(function(row) {
    var profissional = row[0];
    if (!profissionais[profissional]) profissionais[profissional] = [];
    profissionais[profissional].push(row);
  });

  var now = new Date();
  var formattedDate = now.getFullYear() + '-' 
                    + String(now.getMonth() + 1).padStart(2, '0') + '-' 
                    + String(now.getDate()).padStart(2, '0') + '_' 
                    + String(now.getHours()).padStart(2, '0') + ':' 
                    + String(now.getMinutes()).padStart(2, '0');

  //Cria diretório
  var currentFolder = getCurrentFolder();
  var folder = currentFolder.createFolder('Relatorios_Profissionais_' + formattedDate );
  //Antigo
  //var folder = DriveApp.createFolder('Relatorios_Profissionais_' + new Date().toISOString());

  for (var prof in profissionais) {
    var dados = profissionais[prof];

    // Calcular somatório por plano
    var planos = {};
    var totalGeral = 0;
    dados.forEach(function(d) {
      var plano = d[1];
      //var valor = parseFloat(d[4].toString().replace('R$', '').replace('.', '').replace(',', '.'));
      var valor = d[4];
      totalGeral += valor;
      if (!planos[plano]) planos[plano] = 0;
      planos[plano] += valor;
    });

    //Carrega a imagem e base64 para funcionar
    var imageUrl = getImageBase64('logoup.jpg');

    var html = '<html><head><style>';
    html += 'table { border-collapse: collapse; }';
    html += 'th, td { padding: 3px; background-color: #e1e1e1; }';
    html += 'body { font-family: Arial, sans-serif; background-image: url("' + imageUrl + '"); background-repeat: no-repeat; background-position: top right; background-size: 200px 200px; }';
    html += '</style></head><body>';
    html += '<h2 align="center">Relatório de ' + prof + '</h2>';
    html += '<h3>Resumo por Plano</h3>';
    html += '<table border="1"><tr><th>Plano</th><th>Valor Líquido</th></tr>';
    for (var p in planos) {
      html += '<tr><td>' + p + '</td><td align="right">' + planos[p].toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }) + '</td></tr>';
    }
    html += '<tr><td><strong>Total</strong></td><td align="right"><strong>' + totalGeral.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }) + '</strong></td></tr>';
    html += '</table>';

    html += '<h3>Detalhamento</h3>';
    html += '<table border="1"><tr><th>Plano</th><th>Paciente</th><th>Datas Aten.</th><th>Valor Líquido</th><th>Valor Glosa</th><th>Motivo Glosa</th></tr>';
    dados.forEach(function(d) {
      html += '<tr>';
      html += '<td>' + d[1] + '</td>';
      html += '<td>' + d[2] + '</td>';
      html += '<td>' + d[3] + '</td>';
      html += '<td align="right">' + d[4].toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }) + '</td>';
      html += '<td align="right">' + d[5].toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }) + '</td>';
      html += '<td>' + d[6] + '</td>';
      html += '</tr>';
    });
    html += '</table>';
    html += '<p align="right" style="font-size: 10px;">Emitio em ' + formattedDate + '</p>';
    html += '<img src="' + imageUrl + '" style="position: absolute; top: 0; right: 0; z-index: -1;" />';
    html += '</body></html>';

    //Para pegar as 5 primeiras letras da planilha
    var sheetName = ss.getName();
    var prefix = sheetName.substring(0, 5);

    var blob = HtmlService.createHtmlOutput(html).getAs('application/pdf').setName('Relatorio_' + prof + '_' + prefix +'.pdf');
    //Antigo
    //var blob = HtmlService.createHtmlOutput(html).getAs('application/pdf').setName('Relatorio_' + prof + '.pdf');
    folder.createFile(blob);

    //Gera em html
    //var blob = Utilities.newBlob(html, 'text/html', 'Relatorio_' + prof + '_' + prefix + '.html');
    //folder.createFile(blob);

  }
  SpreadsheetApp.getUi().alert('Relatórios gerados na pasta: ' + folder.getUrl());
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

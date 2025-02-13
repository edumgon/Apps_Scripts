//Script antigo, com poucas coisas dinamicas, alguma criadas partindo da gravação de macros.
//Muito para ser otimazado
function CriaAbasPlanos() {
  var apoio = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Apoio");
  var nrow = apoio.getRange("A1:A200").getValues().flat().indexOf("Planos") + 3 ;
  var lrow = apoio.getLastRow() - nrow +1;
  var listaPlanos = apoio.getRange(nrow,1,lrow,1).getValues();
//  var listaPlanos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Apoio").getRange("A47:A61").getValues();
//  Logger.log(apoio.getLastRow());
  var nnrow = nrow + lrow -1;
  for(var plano in listaPlanos){
    var names = listaPlanos[plano];
    if(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(names[0])){
      Browser.msgBox('Remova as abas dos planos primeiro!  \\n Plano '+names[0]+' ainda existente!');
      return 0;
    }
  }
  var capa = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Capa");
  //Deletar colunas
  var ntotal = capa.getRange("B2").getValues().flat().indexOf("Total") + lrow;
  var ultcol = capa.getLastColumn() -3;
  capa.deleteColumns(4,ultcol);
  for(var plano in listaPlanos.reverse()){
    var names = listaPlanos[plano];
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    //salva os valores
    var nps = apoio.getRange(nnrow, 2).getValue();
    var npc = apoio.getRange(nnrow, 3).getValue();
    var npd = apoio.getRange(nnrow, 4).getValue();
    var npa = apoio.getRange(nnrow, 5).getValue();
    var vps = apoio.getRange(nnrow, 6).getValue();
    var vpc = apoio.getRange(nnrow, 7).getValue();
    var vpd = apoio.getRange(nnrow, 8).getValue();
    var vpa = apoio.getRange(nnrow, 9).getValue();
    var vss = apoio.getRange(nnrow, 10).getValue();
    var vsc = apoio.getRange(nnrow, 11).getValue();
    var vsd = apoio.getRange(nnrow, 12).getValue();
    var vsa = apoio.getRange(nnrow, 13).getValue();
    spreadsheet.getSheetByName("_Modelo").activate();
    spreadsheet.duplicateActiveSheet();
    spreadsheet.getActiveSheet().setName(names[0]);
    //copia os valores
    spreadsheet.getActiveSheet().getRange("A2").setValue('1 > '+nps);
    spreadsheet.getActiveSheet().getRange("A3").setValue('2 > '+npc);
    spreadsheet.getActiveSheet().getRange("A4").setValue('3 > '+npd);
    spreadsheet.getActiveSheet().getRange("A5").setValue('4 > '+npa);
    spreadsheet.getActiveSheet().getRange("B2").setValue(vps);
    spreadsheet.getActiveSheet().getRange("B3").setValue(vpc);
    spreadsheet.getActiveSheet().getRange("B4").setValue(vpd);
    spreadsheet.getActiveSheet().getRange("B5").setValue(vpa);
    spreadsheet.getActiveSheet().getRange("C2").setValue(vss);
    spreadsheet.getActiveSheet().getRange("C3").setValue(vsc);
    spreadsheet.getActiveSheet().getRange("C4").setValue(vsd);
    spreadsheet.getActiveSheet().getRange("C5").setValue(vsa);
      //adciona coluna
      capa.insertColumnsAfter(3, 1);
      //copia
      capa.getRange("C:C").copyTo(capa.getRange("D:D"));
      capa.getRange(2,4).setValue(names[0]);
    plano++
    nnrow--
  }
  var lastcolum = columnToLetter(lrow+3)
  capa.activate();
  capa.getRange('B3').activate();
  capa.getCurrentCell().setFormula('=if(SUM(D3:'+lastcolum+'3)=0;"";SUM(D3:'+lastcolum+'3))');
  var clrow = spreadsheet.getLastRow()-4;
  //arrumar b42, ajustar para encontrar última linha
  capa.getActiveRange().autoFill(spreadsheet.getRange('B3:B'+clrow+''), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  clrow = clrow+3;
  capa.getRange(clrow,2).setFormula('=if(SUM(D'+clrow+':'+lastcolum+clrow+')=0;"";SUM(D'+clrow+':'+lastcolum +clrow+'))');
  clrow = clrow+1;
  capa.getRange(clrow,2).setFormula('=if(SUM(D'+clrow+':'+lastcolum+clrow+')=0;"";SUM(D'+clrow+':'+lastcolum +clrow+'))');

  //Esconde modelo
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("_Modelo").activate();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("_Modelo").hideSheet();
 };

function myFunction() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A7:D202').activate();
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).sort({column: 1, ascending: true});
};

function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}


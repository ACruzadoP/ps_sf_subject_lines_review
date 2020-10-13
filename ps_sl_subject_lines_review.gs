function ps_sl_subject_lines_review() {
  
  var important_sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Income');
  var important_sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Working sheet');
  var important_sheet3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Outcome');
  
  var dataRange_important_sheet1 = important_sheet1.getDataRange();
  var values_important_sheet1 = dataRange_important_sheet1.getValues();
  
  var dataRange_important_sheet2 = important_sheet2.getDataRange();
  var values_important_sheet2 = dataRange_important_sheet2.getValues();
  
  var dataRange_important_sheet3 = important_sheet3.getDataRange();
  var values_important_sheet3 = dataRange_important_sheet3.getValues();
  
  var ArrayofTicketNumbersperWrongSL = '';
  
  for (var i = 1; i <= values_important_sheet2.length; i++){
    if (values_important_sheet2[i-1][1] == 'Fail'){
      important_sheet3.getRange(important_sheet3.getDataRange().getValues().filter(String).length+1,1).setValue(values_important_sheet2[i-1][0].toString());
      for (var j = 1; j <= values_important_sheet1.length; j++){
        if (values_important_sheet1[j-1][1] == values_important_sheet2[i-1][0].toString()){
          if (ArrayofTicketNumbersperWrongSL == ''){
            ArrayofTicketNumbersperWrongSL = values_important_sheet1[j-1][0].toString();
          }
          else{
            ArrayofTicketNumbersperWrongSL = ArrayofTicketNumbersperWrongSL + ', ' + values_important_sheet1[j-1][0].toString();
          }
        }
      }
      important_sheet3.getRange(important_sheet3.getDataRange().getValues().filter(String).length,2).setValue(ArrayofTicketNumbersperWrongSL);
      var cell = important_sheet3.getRange(important_sheet3.getDataRange().getValues().filter(String).length,2);
      cell.setNumberFormat("@");
      ArrayofTicketNumbersperWrongSL = '';
    }
  }
}
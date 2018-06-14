function Schedule_Win_Loss_Recorder_Builder_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schedule - Win / Loss Recording");

  primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule - Builder');
  primary_sheet.activate()
  var col_range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schedule - Builder").getLastRow();
  var score_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Score Options");

  primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule - Win / Loss Recording');
  primary_sheet.activate()

  var sheet_last_row = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schedule - Win / Loss Recording").getLastRow();

  var num_rows = col_range - sheet_last_row

  if (num_rows > 0) {
    addrow_(num_rows, sheet_last_row)
  }

  for (i = 2; i < col_range; i++) {

    var var_blank_cell = ss.getRange(i, 1).getValue()
    var winner_cell = ss.getRange(i, 4, 1, 1);
    var winner_range = ss.getRange(i, 1, 1, 2);
    var score_cell = ss.getRange(i, 6, 1, 1);
    var score_range = score_sheet.getRange(2, 1, 2, 1);

    if (var_blank_cell != "") {

      var winner_rule = SpreadsheetApp.newDataValidation().requireValueInRange(winner_range).build();
      winner_cell.setDataValidation(winner_rule);

      var score_rule = SpreadsheetApp.newDataValidation().requireValueInRange(score_range).build();
      score_cell.setDataValidation(score_rule);

    } else {
      winner_cell.clearDataValidations();
      score_cell.clearDataValidations();
    }

  }

}



function Schedule_Win_Loss_Formulas_() {

  // gets the laste row from the schedule builder sheet
  primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule - Builder');
  primary_sheet.activate()
  // stores last row
  var col_range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schedule - Builder").getLastRow();


  //  switches to the schedule sheet
  primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule - Win / Loss Recording');
  primary_sheet.activate()
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schedule - Win / Loss Recording");

  //  loops through the sheet and populates the formulas all the way down
  for (i = 2; i < col_range; i++) {
    var col_A_formula = sheet.getRange(i, 1, 1, 1).getFormulasR1C1();
    var col_B_formula = sheet.getRange(i, 2, 1, 1).getFormulasR1C1();
    var col_C_formula = sheet.getRange(i, 3, 1, 1).getFormulasR1C1();
    var col_E_formulas = sheet.getRange(i, 5, 1, 1).getFormulasR1C1();

    var col_A_new_range = sheet.getRange(i + 1, 1, 1, 1);
    var col_B_new_range = sheet.getRange(i + 1, 2, 1, 1);
    var col_C_new_range = sheet.getRange(i + 1, 3, 1, 1);
    var col_E_new_range = sheet.getRange(i + 1, 5, 1, 1);

    col_A_new_range.setFormulasR1C1(col_A_formula);
    col_B_new_range.setFormulasR1C1(col_B_formula);
    col_C_new_range.setFormulasR1C1(col_C_formula);
    col_E_new_range.setFormulasR1C1(col_E_formulas);
  }

}


function Schedule_Desktop_Formulas_() {

  primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule - Win / Loss Recording');
  primary_sheet.activate()

  // stores last row
  var col_range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schedule - Win / Loss Recording").getLastRow();

  primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule - Live On Site - Desktop');
  primary_sheet.activate()

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schedule - Live On Site - Desktop");


  for (i = 2; i < col_range - 1; i++) {

    var col_A_formula = sheet.getRange(i, 1, 1, 1).getFormulasR1C1();
    var col_B_formula = sheet.getRange(i, 2, 1, 1).getFormulasR1C1();
    var col_C_formula = sheet.getRange(i, 3, 1, 1).getFormulasR1C1();


    var col_A_new_range = sheet.getRange(i + 1, 1, 1, 1);
    var col_B_new_range = sheet.getRange(i + 1, 2, 1, 1);
    var col_C_new_range = sheet.getRange(i + 1, 3, 1, 1);

    col_A_new_range.setFormulasR1C1(col_A_formula);
    col_B_new_range.setFormulasR1C1(col_B_formula);
    col_C_new_range.setFormulasR1C1(col_C_formula);

  }

}


function Schedule_Builder_() {

  primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule - Raw Numbers');
  primary_sheet.activate()

  // stores last row
  var col_range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schedule - Raw Numbers").getLastRow();

  primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule - Builder');
  primary_sheet.activate()

  dropdown_range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule - Builder');

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schedule - Builder");

  var dropdown_range = SpreadsheetApp.getActive().getRange('\'Team Name & Numbers\'!$A$2:$A$11');

  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(dropdown_range).build();


  for (i = 2; i < col_range - 1; i++) {

    var col_A_formula = sheet.getRange(i, 1, 1, 1).getFormulasR1C1();
    var col_B_formula = sheet.getRange(i, 2, 1, 1).getFormulasR1C1();
    var col_C_formula = sheet.getRange(i, 3, 1, 1).getFormulasR1C1();
    var col_D_formula = sheet.getRange(i, 4, 1, 1).getFormulasR1C1();
    var col_E_range = sheet.getRange(i, 5, 1, 1);
    var col_F_range = sheet.getRange(i, 6, 1, 1);
    var var_blank_cell = sheet.getRange(i, 1).getValue()


    //    Populates the data validation dropdown range
    if (var_blank_cell != "") {

      col_E_range.setDataValidation(rule);
      col_F_range.setDataValidation(rule);

    } else {
      col_E_range.clearDataValidations();
      col_F_range.clearDataValidations();
    }

    var col_A_new_range = sheet.getRange(i + 1, 1, 1, 1);
    var col_B_new_range = sheet.getRange(i + 1, 2, 1, 1);
    var col_C_new_range = sheet.getRange(i + 1, 3, 1, 1);
    var col_D_new_range = sheet.getRange(i + 1, 4, 1, 1);

    col_A_new_range.setFormulasR1C1(col_A_formula);
    col_B_new_range.setFormulasR1C1(col_B_formula);
    col_C_new_range.setFormulasR1C1(col_C_formula);
    col_D_new_range.setFormulasR1C1(col_D_formula);

  }

}


function unmerge_() {

  var Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule - Win / Loss Recording');
  Sheet.activate()
  var Range = Sheet.getDataRange().activate();
  Range.breakApart()

  var Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule - Builder');
  Sheet.activate()
  var Range = Sheet.getDataRange().activate();
  Range.breakApart()
}


function Schedule_Mobile_Formulas_() {

  primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule - Win / Loss Recording');
  primary_sheet.activate()

  // stores last row
  var col_range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schedule - Win / Loss Recording").getLastRow();

  primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule - Live On Site - Mobile');
  primary_sheet.activate()

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schedule - Live On Site - Mobile");


  for (i = 2; i < col_range - 1; i++) {

    var col_A_formula = sheet.getRange(i, 1, 1, 1).getFormulasR1C1();
    var col_B_formula = sheet.getRange(i, 2, 1, 1).getFormulasR1C1();

    var col_A_new_range = sheet.getRange(i + 1, 1, 1, 1);
    var col_B_new_range = sheet.getRange(i + 1, 2, 1, 1);

    col_A_new_range.setFormulasR1C1(col_A_formula);
    col_B_new_range.setFormulasR1C1(col_B_formula);

  }

}



function league_setup() {

  unmerge_()
  Schedule_Builder_()
  Schedule_Win_Loss_Recorder_Builder_()
  Schedule_Win_Loss_Formulas_()
  Schedule_Desktop_Formulas_()
  Schedule_Mobile_Formulas_()

  //  completion alert
  SpreadsheetApp.getUi().alert("Done with League Setup!")

}

function addrow_(num_rows, row_index) {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.insertRows(row_index - 1, num_rows)
}





// desktop forumla = =IF('WIN/LOSS Finalizer'!F2 <> "",'WIN/LOSS Finalizer'!F2, IF('Schedule - Win / Loss Recording'!A2 <> "",IF( 'Schedule - Win / Loss Recording'!D2 = "",'Schedule - Win / Loss Recording'!A2 & " vs. " & 'Schedule - Win / Loss Recording'!B2, 'Schedule - Win / Loss Recording'!D2 & " vs. " & 'Schedule - Win / Loss Recording'!E2),IF('Schedule - Win / Loss Recording'!C2 <> "", 'Schedule - Win / Loss Recording'!C2,"")))

// mobile formula = =IF('WIN/LOSS Finalizer'!E2 <> "",'WIN/LOSS Finalizer'!E2,IF( 'Schedule - Win / Loss Recording'!A2 <> "",IF('Schedule - Win / Loss Recording'!D2 = "", 'Schedule - Win / Loss Recording'!C2 & ": " & 'Schedule - Win / Loss Recording'!A2 & " vs. " & 'Schedule - Win / Loss Recording'!B2, 'Schedule - Win / Loss Recording'!C2 & ": " & 'Schedule - Win / Loss Recording'!D2 & " vs. " & 'Schedule - Win / Loss Recording'!E2),IF('Schedule - Win / Loss Recording'!C2 <> "", 'Schedule - Win / Loss Recording'!C2,"")))
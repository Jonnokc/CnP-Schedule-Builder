function Winners(){
 var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schedule - Win / Loss Recording");
var col_range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schedule - Builder").getLastRow();
 var score_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Score Options");

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

    } else  {
      winner_cell.setDataValidation(null);
      score_cell.setDataValidation(null);
    }
  }
}



function schedule_formulas() {

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
     var col_A_formula = sheet.getRange(i , 1, 1, 1).getFormulasR1C1();
     var col_B_formula = sheet.getRange(i , 2, 1, 1).getFormulasR1C1();
     var col_C_formula = sheet.getRange(i , 3, 1, 1).getFormulasR1C1();
     var col_E_formulas = sheet.getRange(i , 5, 1, 1).getFormulasR1C1();
     var col_A_new_range = sheet.getRange(i + 1, 1, 1, 1);
     var col_B_new_range = sheet.getRange(i + 1, 2, 1, 1);
     var col_C_new_range = sheet.getRange(i + 1, 3, 1, 1);
     var col_E_new_range = sheet.getRange(i + 1, 5, 1, 1);
     col_A_new_range.setFormulasR1C1(col_A_formula);
     col_B_new_range.setFormulasR1C1(col_B_formula);
     col_C_new_range.setFormulasR1C1(col_C_formula);
     col_E_new_range.setFormulasR1C1(col_E_formulas);
   }
//  completion alert
  SpreadsheetApp.getUi().alert("Done with Formulas!")
}


function Schedule_Desktop_Formulas() {

 primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule - Win / Loss Recording');
 primary_sheet.activate()

 // stores last row
 var col_range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schedule - Win / Loss Recording").getLastRow();

 primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule - Live On Site - Desktop');
 primary_sheet.activate()

 var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schedule - Live On Site - Desktop");


   for (i = 2; i < col_range - 1; i++) {

     var col_A_formula = sheet.getRange(i , 1, 1, 1).getFormulasR1C1();
     var col_B_formula = sheet.getRange(i , 2, 1, 1).getFormulasR1C1();
     var col_C_formula = sheet.getRange(i , 3, 1, 1).getFormulasR1C1();

     var col_A_new_range = sheet.getRange(i + 1, 1, 1, 1);
     var col_B_new_range = sheet.getRange(i + 1, 2, 1, 1);
     var col_C_new_range = sheet.getRange(i + 1, 3, 1, 1);

     col_A_new_range.setFormulasR1C1(col_A_formula);
     col_B_new_range.setFormulasR1C1(col_B_formula);
     col_C_new_range.setFormulasR1C1(col_C_formula);

   }
  //  completion alert
  SpreadsheetApp.getUi().alert("Done with Formulas!")
}


function Schedule_Builder_formulas() {

 primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule - Raw Numbers');
 primary_sheet.activate()

 // stores last row
 var col_range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schedule - Raw Numbers").getLastRow();

 primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule - Builder');
 primary_sheet.activate()

 var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schedule - Builder");


   for (i = 2; i < col_range - 1; i++) {

     var col_A_formula = sheet.getRange(i , 1, 1, 1).getFormulasR1C1();
     var col_B_formula = sheet.getRange(i , 2, 1, 1).getFormulasR1C1();

     var col_A_new_range = sheet.getRange(i + 1, 1, 1, 1);
     var col_B_new_range = sheet.getRange(i + 1, 2, 1, 1);

     col_A_new_range.setFormulasR1C1(col_A_formula);
     col_B_new_range.setFormulasR1C1(col_B_formula);

   }
  //  completion alert
  SpreadsheetApp.getUi().alert("Done with Formulas!")
}


function unmerge() {
var Sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
var Range = Sheet.getDataRange().activate();
  Range.breakApart()
}


function Schedule_Mobile_Formulas() {

 primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule - Win / Loss Recording');
 primary_sheet.activate()

 // stores last row
 var col_range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schedule - Win / Loss Recording").getLastRow();

 primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule - Live On Site - Mobile');
 primary_sheet.activate()

 var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schedule - Live On Site - Mobile");


   for (i = 2; i < col_range - 1; i++) {

     var col_A_formula = sheet.getRange(i , 1, 1, 1).getFormulasR1C1();
     var col_B_formula = sheet.getRange(i , 2, 1, 1).getFormulasR1C1();

     var col_A_new_range = sheet.getRange(i + 1, 1, 1, 1);
     var col_B_new_range = sheet.getRange(i + 1, 2, 1, 1);

     col_A_new_range.setFormulasR1C1(col_A_formula);
     col_B_new_range.setFormulasR1C1(col_B_formula);

   }
  //  completion alert
  SpreadsheetApp.getUi().alert("Done with Formulas!")
}

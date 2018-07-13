function Schedule_Win_Loss_Recorder_Builder_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "Schedule - Win / Loss Recording"
  );

  ss.activate();
  var col_range = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Schedule - Win / Loss Recording")
    .getLastRow();

  var score_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "Score Options"
  );

  for (i = 2; i < col_range; i++) {
    var var_blank_cell = ss.getRange(i, 1).getValue();
    var winner_cell = ss.getRange(i, 4, 1, 1);
    var winner_range = ss.getRange(i, 1, 1, 2);
    var score_cell = ss.getRange(i, 6, 1, 1);
    var score_range = score_sheet.getRange(2, 1, 2, 1);

    if (var_blank_cell != "") {
      var winner_rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(winner_range)
        .build();
      winner_cell.setDataValidation(winner_rule);

      var score_rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(score_range)
        .build();
      score_cell.setDataValidation(score_rule);
    } else {
      winner_cell.clearDataValidations();
      score_cell.clearDataValidations();
    }
  }
}

function Schedule_Win_Loss_Recording_() {
  // gets the laste row from the schedule builder sheet
  primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "Schedule - Win / Loss Recording"
  );
  primary_sheet.activate();
  // stores last row
  var col_range = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Schedule - Win / Loss Recording")
    .getMaxRows();

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "Schedule - Win / Loss Recording"
  );

  col_A_formula = sheet
    .getRange(2, 1, 1, 1)
    .setFormula(
      '=IF(D2 = "",IF(\'Schedule - Builder\'!A2 <> "",\'Schedule - Builder\'!A2,""),D2)'
    );
  col_B_formula = sheet
    .getRange(2, 2, 1, 1)
    .setFormula(
      "=IF(A2 = 'Schedule - Builder'!A2,'Schedule - Builder'!B2,'Schedule - Builder'!A2)"
    );
  col_C_formula = sheet
    .getRange(2, 3, 1, 1)
    .setFormula(
      "=IF('Schedule - Builder'!C2 <> \"\",'Schedule - Builder'!C2,\"\")"
    );
  col_E_formula = sheet
    .getRange(2, 5, 1, 1)
    .setFormula('=If(D2 = A2,B2,IF(AND(D2 = B2, D2 <> ""),A2,""))');

  //  loops through the sheet and populates the formulas all the way down
  for (i = 2; i < col_range; i++) {
    var col_A_formula = sheet.getRange(i, 1, 1, 1).getFormulasR1C1();
    var col_B_formula = sheet.getRange(i, 2, 1, 1).getFormulasR1C1();
    var col_C_formula = sheet.getRange(i, 3, 1, 1).getFormulasR1C1();
    var col_E_formula = sheet.getRange(i, 5, 1, 1).getFormulasR1C1();

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

function sheet_size_() {
  var all_sheets = [
    "Schedule - Builder",
    "Schedule - Win / Loss Recording",
    "Schedule - Live On Site - Mobile",
    "Schedule - Live On Site - Desktop",
    "WIN/LOSS Finalizer"
  ];

  var arrayLength = all_sheets.length;

  for (var i = 0; i < arrayLength; i++) {
    primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      "Schedule - Raw Numbers"
    );
    primary_sheet.activate();

    // stores last row
    var schedule_row_count = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("Schedule - Raw Numbers")
      .getLastRow();

    primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      all_sheets[i]
    );
    primary_sheet.activate();

    var data_sheet_rows = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(all_sheets[i])
      .getMaxRows();

    var num_rows = schedule_row_count - data_sheet_rows;

    if (num_rows > 0) {
      addrow_(num_rows, data_sheet_rows);
    } else if (num_rows < 0) {
      deleterow_(schedule_row_count, num_rows);
    }
  }
}

function Schedule_Desktop_Formulas_() {
  primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "Schedule - Raw Numbers"
  );
  primary_sheet.activate();

  // stores last row
  var schedule_row_count = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Schedule - Raw Numbers")
    .getLastRow();

  primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "Schedule - Live On Site - Desktop"
  );
  primary_sheet.activate();

  var desktop_total_rows = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Schedule - Live On Site - Desktop")
    .getMaxRows();

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "Schedule - Live On Site - Desktop"
  );

  col_A_formula = sheet
    .getRange(2, 1, 1, 1)
    .setFormula(
      "=IF('Schedule - Builder'!B2 <> \"\", 'Schedule - Builder'!C2, \"\")"
    );
  col_B_formula = sheet
    .getRange(2, 2, 1, 1)
    .setFormula(
      "= IF('WIN/LOSS Finalizer'!F2 <> \"\", 'WIN/LOSS Finalizer'!F2, IF('Schedule - Win / Loss Recording'!A2 <> \"\", IF('Schedule - Win / Loss Recording'!D2 = \"\", \
      'Schedule - Win / Loss Recording'!A2 & \" vs. \" & 'Schedule - Win / Loss Recording'!B2, 'Schedule - Win / Loss Recording'!D2 & \" vs. \" \
      & 'Schedule - Win / Loss Recording'!E2), IF('Schedule - Win / Loss Recording'!C2 <> \"\", 'Schedule - Win / Loss Recording'!C2, \"\")))"
    );
  col_C_formula = sheet
    .getRange(2, 3, 1, 1)
    .setFormula(
      "=IF('Schedule - Builder'!A2 <> \"\", 'Schedule - Win / Loss Recording'!F2, \"\")"
    );

  for (i = 2; i < desktop_total_rows - 1; i++) {
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
  primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "Schedule - Builder"
  );
  primary_sheet.activate();

  // stores last row
  var col_range = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Schedule - Builder")
    .getMaxRows();

  dropdown_range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "Schedule - Builder"
  );

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "Schedule - Builder"
  );

  var dropdown_range = SpreadsheetApp.getActive().getRange(
    "'Team Name & Numbers'!$A$2:$A$11"
  );

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(dropdown_range)
    .build();

  var col_A_formula = sheet
    .getRange(2, 1, 1, 1)
    .setFormula(
      "=IF(E2 = \"\",IFERROR(INDEX('Team Name & Numbers'!$A:$A,MATCH('Schedule - Raw Numbers'!A2,'Team Name & Numbers'!$B:$B,0)),\"\"),E2)"
    );
  var col_B_formula = sheet
    .getRange(2, 2, 1, 1)
    .setFormula(
      "=IF(F2 = \"\",IFERROR(INDEX('Team Name & Numbers'!$A:$A,MATCH('Schedule - Raw Numbers'!B2,'Team Name & Numbers'!$B:$B,0)),\"\"),F2)"
    );
  var col_C_formula = sheet
    .getRange(2, 3, 1, 1)
    .setFormula("='Schedule - Raw Numbers'!C2");
  var col_D_formula = sheet
    .getRange(2, 4, 1, 1)
    .setFormula("='Schedule - Raw Numbers'!D2");
  var col_G_formula = sheet
    .getRange(2, 7, 1, 1)
    .setFormula(
      '=IFERROR(IF(INDEX(\'Schedule - Checker\'!E:E,MATCH(A2,\'Schedule - Checker\'!A:A,0)) \
    <> 2,"Problem with " & A2 & " In Week " & D2,IF(INDEX(\'Schedule - Checker\'!E:E,MATCH(B2,\'Schedule - Checker\'!A:A,0)) \
    <> 2, "Problem with " & B2 & " In Week " & D2,"")))'
    );

  for (i = 2; i < col_range - 1; i++) {
    var col_A_formula = sheet.getRange(i, 1, 1, 1).getFormulasR1C1();
    var col_B_formula = sheet.getRange(i, 2, 1, 1).getFormulasR1C1();
    var col_C_formula = sheet.getRange(i, 3, 1, 1).getFormulasR1C1();
    var col_D_formula = sheet.getRange(i, 4, 1, 1).getFormulasR1C1();
    var col_E_range = sheet.getRange(i, 5, 1, 1);
    var col_F_range = sheet.getRange(i, 6, 1, 1);
    var col_G_formula = sheet.getRange(i, 7, 1, 1).getFormulasR1C1();
    var var_blank_cell = sheet.getRange(i, 1).getValue();

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
    var col_G_new_range = sheet.getRange(i + 1, 7, 1, 1);

    col_A_new_range.setFormulasR1C1(col_A_formula);
    col_B_new_range.setFormulasR1C1(col_B_formula);
    col_C_new_range.setFormulasR1C1(col_C_formula);
    col_D_new_range.setFormulasR1C1(col_D_formula);
    col_G_new_range.setFormulasR1C1(col_G_formula);
  }
}

function unmerge_() {
  var all_sheets = [
    "Schedule - Builder",
    "Schedule - Win / Loss Recording",
    "Schedule - Live On Site - Mobile",
    "Schedule - Live On Site - Desktop",
    "WIN/LOSS Finalizer"
  ];

  var arrayLength = all_sheets.length;

  for (var i = 0; i < arrayLength; i++) {
    primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      all_sheets[i]
    );
    primary_sheet.activate();
    var Range = Sheet.getDataRange().activate();
    Range.breakApart();
  }
}

function Schedule_Mobile_Formulas_() {
  // stores last row
  primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "Schedule - Live On Site - Mobile"
  );
  primary_sheet.activate();

  var col_range = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Schedule - Live On Site - Mobile")
    .getMaxRows();

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "Schedule - Live On Site - Mobile"
  );

  col_A_formula = sheet
    .getRange(2, 1, 1, 1)
    .setFormula(
      "=IF('WIN/LOSS Finalizer'!E2 <> \"\",'WIN/LOSS Finalizer'!E2,IF( 'Schedule - Win / Loss Recording'!A2 <> \"\", \
  IF('Schedule - Win / Loss Recording'!D2 = \"\", 'Schedule - Win / Loss Recording'!C2 & \": \" & 'Schedule - Win / Loss Recording'!A2 & \" vs. \" \
  & 'Schedule - Win / Loss Recording'!B2, 'Schedule - Win / Loss Recording'!C2 & \": \" & 'Schedule - Win / Loss Recording'!D2 & \" vs. \" \
  & 'Schedule - Win / Loss Recording'!E2) \
  ,IF('Schedule - Win / Loss Recording'!C2 <> \"\",'Schedule - Win / Loss Recording'!C2,\"\")))"
    );

  col_B_formula = sheet
    .getRange(2, 2, 1, 1)
    .setFormula(
      "=IF('Schedule - Builder'!A2 <> \"\",'Schedule - Win / Loss Recording'!F2,\"\")"
    );

  for (i = 2; i < col_range - 1; i++) {
    var col_A_formula = sheet.getRange(i, 1, 1, 1).getFormulasR1C1();
    var col_B_formula = sheet.getRange(i, 2, 1, 1).getFormulasR1C1();

    var col_A_new_range = sheet.getRange(i + 1, 1, 1, 1);
    var col_B_new_range = sheet.getRange(i + 1, 2, 1, 1);

    col_A_new_range.setFormulasR1C1(col_A_formula);
    col_B_new_range.setFormulasR1C1(col_B_formula);
  }
}

function Win_Loss_Finalizer_() {
  // stores last row
  primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "WIN/LOSS Finalizer"
  );
  primary_sheet.activate();

  var col_range = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("WIN/LOSS Finalizer")
    .getMaxRows();

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "WIN/LOSS Finalizer"
  );

  col_A_formula = sheet
    .getRange(2, 1, 1, 1)
    .setFormula(
      "=IF('Schedule - Builder'!B2 <> \"\",'Schedule - Builder'!C2,\"\")"
    );
  col_B_formula = sheet
    .getRange(2, 2, 1, 1)
    .setFormula("='Schedule - Win / Loss Recording'!D2");
  col_C_formula = sheet
    .getRange(2, 3, 1, 1)
    .setFormula("='Schedule - Win / Loss Recording'!E2");
  col_D_formula = sheet
    .getRange(2, 4, 1, 1)
    .setFormula("='Schedule - Win / Loss Recording'!F2");
  col_E_formula = sheet
    .getRange(2, 5, 1, 1)
    .setFormula('=IF(B2 <> "", A2 & ": " & B2 & " Defeated " & C2,"")');
  col_F_formula = sheet
    .getRange(2, 6, 1, 1)
    .setFormula('=IF(B2 <> "", B2 & " Defeated " & C2,"")');

  for (i = 2; i < col_range - 1; i++) {
    var col_A_formula = sheet.getRange(i, 1, 1, 1).getFormulasR1C1();
    var col_B_formula = sheet.getRange(i, 2, 1, 1).getFormulasR1C1();
    var col_C_formula = sheet.getRange(i, 3, 1, 1).getFormulasR1C1();
    var col_D_formula = sheet.getRange(i, 4, 1, 1).getFormulasR1C1();
    var col_E_formula = sheet.getRange(i, 5, 1, 1).getFormulasR1C1();
    var col_F_formula = sheet.getRange(i, 6, 1, 1).getFormulasR1C1();

    var col_A_new_range = sheet.getRange(i + 1, 1, 1, 1);
    var col_B_new_range = sheet.getRange(i + 1, 2, 1, 1);
    var col_C_new_range = sheet.getRange(i + 1, 3, 1, 1);
    var col_D_new_range = sheet.getRange(i + 1, 4, 1, 1);
    var col_E_new_range = sheet.getRange(i + 1, 5, 1, 1);
    var col_F_new_range = sheet.getRange(i + 1, 6, 1, 1);

    col_A_new_range.setFormulasR1C1(col_A_formula);
    col_B_new_range.setFormulasR1C1(col_B_formula);
    col_C_new_range.setFormulasR1C1(col_C_formula);
    col_D_new_range.setFormulasR1C1(col_D_formula);
    col_E_new_range.setFormulasR1C1(col_E_formula);
    col_F_new_range.setFormulasR1C1(col_F_formula);
  }
}

function addrow_(num_rows, row_index) {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.insertRows(row_index, num_rows);
}

function deleterow_(row_index, num_rows) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = num_rows * -1;
  sheet.deleteRows(row_index, rows);
}

function team_rename() {
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert("Do you want to rename a team?", ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      "Team Name & Numbers"
    );
    ss.activate();

    var to_replace = ui
      .prompt("What is the name of the team you want to change?")
      .getResponseText();

    if (to_replace != "") {
      var name_check = validate_team_name_(to_replace);
    } else {
      SpreadsheetApp.getUi().alert("No team name given. Exiting.");
    }
  } else {
    ui.alert("Okiedokie. Won't update any team names!");
    return;
  }

  if (name_check === true) {
    var new_name = ui
      .prompt("Okay that name checks out. What is the new team name?")
      .getResponseText();

    var confirm_change = ui.alert(
      "Confirm you want to replace " + to_replace + " with " + new_name,
      ui.ButtonSet.YES_NO
    );

    if (response == ui.Button.YES) {
      // function which performs the team name find and replace
      replace_team_name_(to_replace, new_name);
    } else {
      ui.alert("Okiedokie. Won't update any team names!");
      return;
    }
  } else if (name_check === false) {
    ui.alert(
      "That team name does not match any existing team. I need the team name EXACTLY as it currently is. Try again."
    );
  }
}

// function to validate that the team name given is infact an existing team name.
function validate_team_name_(team_name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "Team Name & Numbers"
  );

  ss.activate();
  var col_range = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Team Name & Numbers")
    .getLastRow();

  // checks if the team name to replace is within the list of existing teams.
  for (i = 2; i < col_range; i++) {
    var existing_name = ss.getRange(i, 1).getValue();

    // return True if team name is within existing list
    if (team_name == existing_name) {
      return true;
    }
  }
  // return false if team not found
  return false;
}

// performs the actual find and replace
function replace_team_name_(team_name, new_name) {
  var all_sheets = [
    "Team Name & Numbers",
    "Schedule - Builder",
    "Schedule - Win / Loss Recording"
  ];

  var arrayLength = all_sheets.length;

  for (var k = 0; k < arrayLength; k++) {
    primary_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      all_sheets[k]
    );
    primary_sheet.activate();
    // stores last row
    var row_count = SpreadsheetApp.getActiveSpreadsheet().getLastRow();
    var r = SpreadsheetApp.getActiveSheet().getDataRange();
    var cls = r.getNumColumns();
    var i, j, a, find, repl;
    find = team_name;
    repl = new_name;
    for (i = 1; i <= row_count; i++) {
      for (j = 1; j <= cls; j++) {
        a = r.getCell(i, j).getValue();
        if (r.getCell(i, j).getFormula()) {
          continue;
        }
        try {
          a = a.replace(find, repl);
          r.getCell(i, j).setValue(a);
        } catch (err) {
          continue;
        }
      }
    }
  }

  SpreadsheetApp.getUi().alert("Done! Team Name changed.");
}

function league_setup() {
  sheet_size_();
  unmerge_();
  Schedule_Builder_();
  Schedule_Win_Loss_Recorder_Builder_();
  Schedule_Win_Loss_Formulas_();
  Schedule_Desktop_Formulas_();
  Schedule_Mobile_Formulas_();
  Win_Loss_Finalizer_();

  //  completion alert
  SpreadsheetApp.getUi().alert("Done with League Setup!");
}

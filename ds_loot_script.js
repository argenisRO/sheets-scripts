import 'google-apps-script';

function setCell(sheet, cell, value) {
    sheet.getRange(cell).setValue(value);
}

function dupe_to_day(ss, day, offset) {
    var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
    var s_master = ss.getSheetByName("Master Loot");
    var s_new = s_master.copyTo(ss);

    s_new.setName(day + " Loot");
    s_new.getRange("E1").setValue(day);

    var s_console = ss.getSheetByName("Console");
    var c_date = s_console.getRange("C3");
    var date = new Date(c_date.getValue() + offset);

    setCell(s_new, "C2", Utilities.formatDate(new Date(date.getTime() + offset * MILLIS_PER_DAY), "PDT", "MM-dd-yyyy"));

    s_new.getRange("C2").setNumberFormats(c_date.getNumberFormats());

    var protections = s_master.getProtections(SpreadsheetApp.ProtectionType.RANGE);

    for (var i = 0; i < protections.length; ++i) {
        var p = protections[i];
        var range = p.getRange().getA1Notation();
        var p_new = s_new.getRange(range).protect();

        p_new.setDescription(p.getDescription());
        p_new.setWarningOnly(p.isWarningOnly());

        if (!p_new.isWarningOnly()) {
            p_new.removeEditors(p_new.getEditors());
            p_new.addEditors(p.getEditors());
        }
    }

    ss.setActiveSheet(s_new);
    ss.moveActiveSheet(3 + offset);
    ss.setActiveSheet(s_console);
}

function sort_member() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var armada = ss.getSheetByName("Armada Information");
  var data = armada.getRange("B6:L105");
  var status = armada.getRange(5,6).getFormula();
  
  data.sort({column: 2, ascending: true});
  armada.getRange("F5:F105").setFormula(status);
}

function duplicate_week() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    var s_console = ss.getSheetByName("Console");
    var r_output = ss.getRange("F8");

    for (var i = 0; i < days.length; ++i) {
        var s_check = ss.getSheetByName(days[i] + " Loot");

        if (s_check != null) {
            var s_copy = s_check.copyTo(ss);
            var date = s_copy.getRange("C2").getDisplayValue();
            r_output.setValue("Backing up '" + days[i] + " Loot' to '" + date + " - " + days[i] + " Loot'...");
            s_copy.setName(date + " - " + days[i] + " Loot");
            ss.deleteSheet(s_check);
        }
    }

    r_output.setValue("Unprotecting master...");
    unprotect_master();
    r_output.setValue("Protecting master...");
    protect_master();

    for (var i = 0; i < days.length; ++i) {
        r_output.setValue("Copying master to " + days[i] + "...");
        dupe_to_day(ss, days[i], i);
    }

    r_output.setValue("Refreshing hunting totals...");
    update_hunting_totals();
    r_output.setValue("Refreshing loot verification...");
    update_loot_verification();

    r_output.setValue("Flushing changes...");
    SpreadsheetApp.flush();

    r_output.setValue("Done! Happy sea monster hunting!");
}

function delete_week() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    var s_console = ss.getSheetByName("Console");
    var r_output = ss.getRange("F8");

    for (var i = 0; i < days.length; ++i) {
        r_output.setValue("Deleting " + days[i] + " sheet...");
        var s = ss.getSheetByName(days[i] + " Loot");

        if (s != null) {
            ss.deleteSheet(s);
        }
    }

    r_output.setValue("Done deleting! Make sure to add them again!");
}

function update_hunting_totals() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var totals = ss.getSheetByName("Hunting Totals");
    var range = totals.getRange("H6:N105");

    range.setFormulas(range.getFormulas());
}

function update_loot_verification() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var totals = ss.getSheetByName("Loot Verification");
    var c_dates = ["E1", "J1", "O1", "T1", "Y1", "AD1", "AI1"];

    {
        var range = totals.getRange("C4:D103");
        range.setFormulas(range.getFormulas());
    } 
    {
        var range = totals.getRange("H4:I103");
        range.setFormulas(range.getFormulas());
    } 
    {
        var range = totals.getRange("M4:N103");
        range.setFormulas(range.getFormulas());
    } 
    {
        var range = totals.getRange("R4:S103");
        range.setFormulas(range.getFormulas());
    } 
    {
        var range = totals.getRange("W4:X103");
        range.setFormulas(range.getFormulas());
    } 
    {
        var range = totals.getRange("AB4:AC103");
        range.setFormulas(range.getFormulas());
    } 
    {
        var range = totals.getRange("AG4:AH103");
        range.setFormulas(range.getFormulas());
    }
    {
        for (var i = 0; i < c_dates.length; ++i) {
            var range = totals.getRange(c_dates[i]);
            range.setFormulas(range.getFormulas());
        }
    }
}

function protect_master() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var s_console = ss.getSheetByName("Console");
    var protections = s_console.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    var p_copy;

    for (var i = 0; i < protections.length; ++i) {
        var p = protections[i];
        
        if (p.getDescription() == "Perms") {
            p_copy = p;
            break;
        }
    }

    var columns = ["A:A", "C:D", "G:K"];
    var s_master = ss.getSheetByName("Master Loot");

    for (var i = 0; i < columns.length; ++i) {
        var range = s_master.getRange(columns[i]);
        var p = range.protect();

        p.removeEditors(p.getEditors());
        p.addEditors(p_copy.getEditors());
    }
}

function unprotect_master() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var s_master = ss.getSheetByName("Master Loot");
    var protections = s_master.getProtections(SpreadsheetApp.ProtectionType.RANGE);

    for (var i = 0; i < protections.length; ++i) {
        protections[i].remove();
    }
}

function SHEETNAME() {
    var ss = SpreadsheetApp.getActive();
    var name = ss.getSheetName();

    modifyCell("B1", name);
}
import 'google-apps-script';

function update_status(sheet, message) {
    sheet.getRange("H8").setValue(message);
}

function create_signup_day() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var console_ss = ss.getSheetByName("Console");

    update_status(console_ss, "Creating Sheet...");
    dup_master();
    update_status(console_ss, "Finished!");
    SpreadsheetApp.flush();
}

function dup_master() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var console_ss = ss.getSheetByName("Console");
    var nodewar_master = ss.getSheetByName("Node War (Master)");

    var date = console_ss.getRange("C8");
    var new_nodewar = nodewar_master.copyTo(ss);
    var type = (date.getValue().getDay() === 6.0) ? "Siege War" : "Node War";
    var formatted_date = Utilities.formatDate(date.getValue(), "PDT", "EEE MM/dd/yy");

    new_nodewar.setName(type + " (" + formatted_date + ")");
    new_nodewar.getRange("G1").setValue(formatted_date);
    new_nodewar.getRange("A1").setValue(console_ss.getRange("C11").getValue() + " Attendance");

    ss.setActiveSheet(new_nodewar);
    ss.moveActiveSheet(1);
    ss.setActiveSheet(console_ss);
}

function sort_members() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var gs_ss = ss.getSheetByName("Gear Score");
    var data = gs_ss.getRange("B27:Q126");
    var statust = gs_ss.getRange(10, 11).getFormula();
    var status = gs_ss.getRange(26, 11).getFormula();

    data.sort({ column: 2, ascending: true });
    gs_ss.getRange("K10:K20").setFormula(statust);
    gs_ss.getRange("K26:K126").setFormula(status);
}

function add_parties() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var nodewar_sheet = ss.getActiveSheet();
    var console_ss = ss.getSheetByName("Console");
    var members = {}
    var location_e = {
        Harass: { sheet: ["T57:T61", "W57:W61"], console: ["B29:B38"], sub: ["B40:B44"] },
        Flex: { sheet: ["Q57:Q61"], console: ["B49:B53"], sub: ["B55:B59"] },
        Cannons: { sheet: ["Q15:Q19"], console: ["J29:J33"], sub: ["J35:J39"] },
        "Cannon Killer": { sheet: ["T15:T19"], console: ["N29:N33"], sub: ["N35:N39"] },
        "Backline Peel": { sheet: ["W15:W19"], console: ["J49:J53"], sub: ["J55:J59"] },
        Structures: { sheet: ["W8:W12"], console: ["F49:F53"], sub: ["F55:F59"] },
        Defense: { sheet: ["Q8:Q12", "T8:T12"], console: ["F29:F38"], sub: ["F40:F44"] },
    };

    var location_a = {
        Attack: {
            sheet: ["Q22:Q26", "T22:T26", "W22:W26", "Q29:Q33", "T29:T33", "W29:W33", "Q36:Q40", "T36:T40", "W36:W40", "Q43:Q47", "T43:T47", "W43:W47", "Q50:Q54", "T50:T54", "W50:W54"],
            sheet_pa: ["Q64:Q68", "T64:T68", "W64:W68", "Q71:Q75", "T71:T75", "W71:W75"]
        }
    };

    for (var i = 8; i <= 71; i += 7) { nodewar_sheet.getRange(i, 16, 5, 8).clearContent(); }

    nodewar_sheet.getRange(8, 2, 100, 4).getValues().filter(function (value) { return value[2] != "-" }).map(function (player) {
        members[player[0]] = { Family_Name: player[0], Character_Name: player[1], Class: player[2], Role: player[3] };
    });

    if (console_ss.getRange("L43").getValue() === "On") {
        var fill_roles = Object.keys(location_e);
        for (var ex_sub = 0; ex_sub < fill_roles.length; ex_sub++) {
            var current_console = flatten(console_ss.getRange(location_e[fill_roles[ex_sub]].console).getValues());
            var pref_role_mem = Object.keys(members).filter(function (name) {
                return members[name].Role === fill_roles[ex_sub] && current_console.indexOf(name) === -1;
            });
            while (pref_role_mem.length < 5) { pref_role_mem.push(""); } if (pref_role_mem.length > 5) { pref_role_mem = pref_role_mem.slice(0, 5); }
            console_ss.getRange(location_e[fill_roles[ex_sub]].sub).setValues(pref_role_mem.map(function (name) { return [name]; }));
        }
    }

    for (var value in location_e) {
        var formed_party = []
        var console_names = flatten(console_ss.getRange(value.console).getValues());
        var matched_names = find_match(console_names, Object.keys(members)).map(function (name) { delete members[name]; return name; });

        if (matched_names.length < 5) {
            var sub_names = flatten(console_ss.getRange(value.sub).getValues());
            var matched_subs = find_match(sub_names, Object.keys(members));

            matched_subs.slice(0, (5 - matched_names.length)).map(function (name) { delete members[name]; matched_names.push(name); return name; });
            while (matched_names.length < 5) { matched_names.push(""); }
            formed_party.push(matched_names);
        } else if (matched_names.length > 5) {
            while (matched_names.length) {
                var sliced_party = matched_names.splice(0, 5);

                while (sliced_party.length < 5) { sliced_party.push(""); }
                formed_party.push(sliced_party);
            }
        } else {
            formed_party.push(matched_names);
        }

        formed_party.map(function (party, index) {
            nodewar_sheet.getRange(value.sheet[index]).setValues(party.map(function (name) { return [name]; }));
        });
    }

    var p_length = Math.ceil((Object.keys(members).length) / 5);
    var max_parties = new Array(p_length);
    var protected_area = [[], [], [], [], [], []];
    for (var i = 0; i < max_parties.length; i++) { max_parties[i] = []; }

    var pa_rot = Object.keys(members).filter(function (name) {
        return members[name].Class === "Wizard" || members[name].Class === "Witch" || members[name].Class === "Valkyrie";
    });

    pa_rot.map(function (name, index) {
        protected_area[index % Math.ceil((pa_rot.length) / 5)].push(name);
    });

    Object.keys(members).filter(function (name) {
        return members[name].Class === "Wizard" || members[name].Class === "Witch";
    }).map(function (name, index) {
        delete members[name];
        max_parties[index % p_length].push(name);
    });

    Object.keys(members).filter(function (name) {
        return members[name].Class === "Warrior" || members[name].Class === "Valkyrie";
    }).map(function (name, index) {
        delete members[name];
        max_parties[index % p_length].push(name);
    });

    var dps = Object.keys(members).filter(function (name) {
        return members[name].Class !== "Warrior" && members[name].Class !== "Valkyrie" && members[name].Class !== "Witch" && members[name].Class !== "Wizard";
    });

    if (p_length <= location_a.Attack.sheet.length) {
        for (var j = 0; j < max_parties.length; j++) {
            if (max_parties[j].length < 5) {
                dps.splice(0, (5 - max_parties[j].length)).map(function (name) { max_parties[j].push(name); });
                while (max_parties[j].length < 5) {
                    max_parties[j].push("");
                }
            }
            nodewar_sheet.getRange(location_a.Attack.sheet[j]).setValues(max_parties[j].map(function (name) { return [name]; }));
        }
    } else {
        Browser.msgBox("No space for the attack group. Please consider spreading members across parties in the console sheet.");
    }

    if (console_ss.getRange("L41").getValue() === "On") {
        for (var i = 0; i < protected_area.length; i++) {
            while (protected_area[i].length < 5) {
                protected_area[i].push("");
            }
            if (protected_area[i].length > 5) {
                protected_area[i] = protected_area[i].slice(0, 5);
            }
            nodewar_sheet.getRange(location_a.Attack.sheet_pa[i]).setValues(protected_area[i].map(function (name) { return [name]; }));
        }
    }
    SpreadsheetApp.flush();
}

function flatten(arrays) {
    return [].concat.apply([], arrays);
}

function find_match(one, two) {
    var shared = [];
    var x = y = 0;
    var firstx = one.concat().sort();
    var secondy = two.concat().sort();

    while (x < one.length && y < two.length) {
        if (firstx[x] === secondy[y]) {
            shared.push(firstx[x]); x++; y++;
        } else if (firstx[x] < secondy[y]) { x++; } else { y++; }
    } return shared;
}

function lastcol() {
    var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    Logger.log(ss.getLastColumn());
    return ss.getLastColumn();
}

function record_attendance() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var console_ss = ss.getSheetByName("Console");

    var date = console_ss.getRange("C8");
    var formatted_date = Utilities.formatDate(date.getValue(), "PDT", "EEE MM/dd/yy");

    var type = (date.getValue().getDay() === 6.0) ? "Siege War" : "Node War";
    var current_nodewar = ss.getSheetByName(type + " (" + formatted_date + ")");

    var attendance_sheet = ss.getSheetByName("Attendance/Payouts");
    var attendance_history = ss.getSheetByName("Attendance History");
    var last_column = attendance_history.getLastColumn();

    if (attendance_history.getRange(1, 1, 1, last_column).getDisplayValues()[0].indexOf(date.getDisplayValue()) !== -1) {
        Browser.msgBox("Date is already recorded. Make sure to set the correct war date before running the script.");
        update_status(console_ss, "Error");
    } else if (current_nodewar == null) {
        Browser.msgBox("Date does not exist. Please provide the correct war date in the console sheet and avoid changing the sheet names of the wars.");
        update_status(console_ss, "Error");
    } else {
        update_status(console_ss, "Working...");

        var previous_attendance_war = attendance_history.getRange(1, last_column, 109);
        attendance_history.insertColumnAfter(last_column);
        previous_attendance_war.copyTo(attendance_history.getRange(1, last_column + 1, 109));
        current_nodewar.getRange(59, 11, 100).copyTo((attendance_history.getRange(4, last_column + 1, 100)), { contentsOnly: true });
        current_nodewar.getRange(47, 8, 5).copyTo((attendance_history.getRange(105, last_column + 1, 5)), { contentsOnly: true });
        attendance_history.getRange(1, last_column + 1).setValue(console_ss.getRange("C8").getDisplayValue());

        var previous_attendance_payout = attendance_sheet.getRange(1, attendance_sheet.getLastColumn() - 3, 103);
        attendance_sheet.insertColumnAfter(attendance_sheet.getLastColumn() - 3);
        previous_attendance_payout.copyTo(attendance_sheet.getRange(1, attendance_sheet.getLastColumn() - 3, 103));
        attendance_sheet.getRange(1, attendance_sheet.getLastColumn() - 3).setValue(console_ss.getRange("C8").getDisplayValue());

        var attendance_percent = attendance_sheet.getRange(4, attendance_sheet.getLastColumn() - 1, 103);
        var attendance_status = attendance_sheet.getRange(4, attendance_sheet.getLastColumn(), 103).getFormulas();
        attendance_sheet.insertColumnAfter(attendance_sheet.getLastColumn() - 1);
        attendance_percent.copyTo(attendance_sheet.getRange(4, attendance_sheet.getLastColumn() - 1, 103));
        attendance_sheet.getRange(1, attendance_sheet.getLastColumn() - 2).copyTo(attendance_sheet.getRange(1, attendance_sheet.getLastColumn() - 1));
        attendance_sheet.deleteColumn(attendance_sheet.getLastColumn() - 2);
        attendance_sheet.getRange(4, attendance_sheet.getLastColumn(), 103).setFormulas(attendance_status);

        update_status(console_ss, "Successfully added '" + current_nodewar.getSheetName() + "' to the attendance.");

        ss.deleteSheet(current_nodewar);
        SpreadsheetApp.flush();
    }
}
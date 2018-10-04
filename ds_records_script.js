import 'google-apps-script';

function update_members() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var signup = ss.getSheetByName("Sign Up");
    var current_members = ss.getSheetByName("Current Members");
    var stored_info = {}

    signup.getRange(11, 2, 30, 22).getValues().filter(function (value) {
        return value[0] != ""
    }).map(function (member) {
        stored_info[member[0]] = {
            Family_Name:    member[0],
            Character_Name: member[1],
            Class:          member[3],
            Status:         member[4],
            AP:             member[5],
            AP_Awake:       member[6],
            DP:             member[7],
            Fame:           member[8],
            Level:          member[10],
            Axe:            member[11],
            Gear:           member[12],
            Discord:        member[13],
            Wars:           member[14],
            Sailing:        member[15],
            Boat:           member[16],
            BoatP:          member[17],
            BoatC:          member[18],
            BoatPl:         member[19],
            BoatS:          member[20],
            Date:           member[21],
        }
    });

    var current_membs = flatten(current_members.getRange("B6:B105").getValues());
    var last_row = current_membs.filter(function (value) {
        return value != "";
    }).length + 6;
    var column_positions = [2, 3, 12, 13, 14, 15, 16, 17, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30];

    for (var memb in stored_info) {
        if (memb.Status == "New Member" && current_membs.indexOf(memb.Family_Name) === -1) {
            if (last_row <= 105) {
                // This structure is used to maintain Google Sheets rules.
                current_members.getRange(last_row, 2, 1, 29).setValues(
                    [
                        [
                            [memb.Family_Name],
                            [memb.Character_Name],
                            [],
                            [],
                            [],
                            [],
                            [],
                            [],
                            [],
                            [],
                            [memb.Class],
                            ["Unassigned"],
                            [memb.AP],
                            [memb.AP_Awake],
                            [memb.DP],
                            [memb.Fame],
                            [],
                            [memb.Level],
                            [memb.Axe],
                            [memb.Gear],
                            [memb.Discord],
                            [memb.Wars],
                            [memb.Sailing],
                            [memb.Boat],
                            [memb.BoatP],
                            [memb.BoatC],
                            [memb.BoatPl],
                            [memb.BoatS],
                            [memb.Date]
                        ]
                    ]
                );
            } else {
                Browser.msgBox("Failed to finish because the members list is full.");
                return;
            }
        } else if (memb.Status == "Update Info" && current_membs.indexOf(memb.Family_Name) != -1) {
            var member_row = current_membs.lastIndexOf(memb.Family_Name) + 6;
            var count = 0;

            for (var [key, val] in memb) {
                if (key == "Status" || val == "" || val == "-") {} else {
                    current_members.getRange(member_row, column_positions[count]).setValue(val);
                }
                count++;
            }
        }
        memb.Status == "New Member" ? last_row++ : {};
    }
    sort_members()
    signup.getRangeList(["B11:J40", "L11:V40"]).clearContent();

}

function sort_members() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var current_members = ss.getSheetByName("Current Members");
    var renown = current_members.getRange(5, 18).getFormula();

    current_members.getRange("B6:AD105").sort({
        column: 2,
        ascending: true
    });
    current_members.getRange(5, 18, 101).setFormula(renown)
}

function flatten(array) {
    return [].concat.apply([], array);
}
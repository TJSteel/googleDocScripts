/* Coded by Jay Steel, any issues e-mail me at - jaysteel78@gmail.com

ASCII - http://patorjk.com/software/taag/#p=display&c=c&f=Big&t=Enter%20Text

*/

var administratorsList = ["jaysteel78@gmail.com", // TJSteel
    "b.williams0913@gmail.com", //Big Ben
    "youshi1@att.net", //Ted / Darkzer
    "driftinhero@gmail.com",
    "thomdebruijn92@gmail.com", // Koenigsegg R
    "zak.scholes@btinternet.com"], // Scholesy
    // List of people who are given editor access to the spreadsheet
    editorsList = ["ianb7116@gmail.com", //AMS Turismo
        "b.williams0913@gmail.com", //Ax4x Big Ben
        "Luke.edwards89@gmail.com", //Ax4x Cowboy
        "mattc4348@gmail.com", //Ax4x Matt
        "flyinmikeyj@gmail.com", //Ax4x Mikey J
        "johnihle@gmail.com", //Ax4x Se7en
        "driftinhero@gmail.com", //DDM DriftiN
        "Mrdriv3gaming@gmail.com", //DDM Driv3
        "hastingsgaming.yt@gmail.com", //DDM Jam
        "ritchard1995@gmail.com", //DDM MeaD 212
        "hemides@gmail.com", //DOR Corvid
        "ruttemartijn@gmail.com", //DOR Martxn
        "gatonoob@gmail.com", //DRR gato
        "cpiotrowski311@yahoo.com", //HCR ChicaiN
        "jan_szczecinski@yahoo.com", //Jan Filip S
        "thomdebruijn92@gmail.com", //Koenigsegg R
        "larbart01@gmail.com", //larrybtheg
        "haydenboydis@gmail.com", //LMP Bubz
        "xcallumx98@gmail.com", //LSEM Yuuko
        "cameronhillman@gmail.com", //M0STW4NT3D 46
        "Brooksy0907@live.com", //MRT Brooksy
        "keywalsheuk@hotmail.com", //NothelleN340
        "rwb_1118@yahoo.com", //SFM Cowmaster
        "youshi1@att.net", //SFM Darkzer
        "jordan.russell.000@gmail.com", //SFM SnowStorm
        "thethrashmetalguru@gmail.com", //Springaahhh
        "steviep847@yahoo.com", //SSR M3yh3m724
        "andrewjwisdom@gmail.com", //SVR Boxer
        "mick_c_cars@hotmail.co.uk", //SVR Coyote
        "josh_nowaczyk@yahoo.com", //SVR Incinerate
        "alexmacdonald25@yahoo.com", //SVR Spector
        "jaysteel78@gmail.com", //TJSteel
        "danielsouthby16@hotmail.co.uk", //TPR South
        "derek.herchko@gmail.com"], //WhoIsJohn117
    driversLicenceNumberColumn = 1, //A
    driversUserNameColumn = 2, //B

    entryListHeaders = 4,
    entryListCarNumberColumn = 5, //E
    entryListLicenceNumberColumnNum1 = 6, //F
    entryListLicenceNumberColumnNum2 = 8, //H
    entryListLicenceNumberColumnNum3 = 10, //J
    entryListLicenceNumberColumnNum4 = 12, //L
    entryListLicenceNumberColumnNum5 = 14, //N
    entryListLicenceNumberColumnNum6 = 16, //P    
    sheetNameEntryList = "Entry List",
    sheetNameSpotterGuide = "Spotter Guide";

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    // Or DocumentApp or FormApp.
    ui.createMenu("TJ's Scripts")
        .addItem("Add Driver by Licence Number", "AddDriverByLicenceNumber")
        .addItem("Add Driver by Gamertag", "AddDriverByGamertag")
        .addItem("Add Drivers to Qualifying Sheet", "updateQualifyingDrivers")
        .addItem("Add Team Numbers to Spotter Guide", "addTeamToSpotterGuide") //this line adds an option to the menu bar which calls a function called addTeamToSpotterGuide
        .addItem("Remove Team Numbers from Spotter Guide", "removeTeamFromSpotterGuide")
        .addItem("Clone First Table Entry", "cloneSpottersLayout")
        //.addItem("Sort Sheet", "SortSheet")
        .addToUi();
}

function addAdmin() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.addEditors(editorsList);
    var sheet = ss.getSheets();
    for (var i = 0; i < sheet.length; i++) {
        sheet[i].protect().addEditors(administratorsList);
    }
}

function finaliseDoc() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    deleteTriggers();
    removeProtection(ss);
    removeFormulas(ss);
    removeEditors(ss);
}

function removeProtection(ss) {
    var sheet = ss.getSheets();
    for (var i = 0; i < sheet.length; i++) {
        sheet[i].protect().remove();
    }
}
function removeFormulas(ss) {
    var sheet = ss.getSheets();
    for (var i = 0; i < sheet.length; i++) {
        var range = sheet[i].getRange(1, 1, sheet[i].getMaxRows(), sheet[i].getMaxColumns());
        range.setValues(range.getValues());
    }
}
function removeEditors(ss) {
    var editors = ss.getEditors();
    for (var i = 0, len = editors.length; i < len; i++) {
        ss.removeEditor(editors[i])
    }
}
// Delete all existing Script Triggers
function deleteTriggers() {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0, length = triggers.length; i < length; i++) {
        ScriptApp.deleteTrigger(triggers[i]);
    }
}

function addTeamToSpotterGuide() //this is the script which will scan all teams and add them to the guide
{
    //firstly we'll start with a list of variables to tell the script what your data looks like
    var sheetCurrentRow = 1, //set to the first car numbers row
        sheetCurrentColumn = 8, //set to first car numbers column, this is H but numbers are easier to work with in scripts
        tableColumnCount = 2, //how many teams displayed on each row basically
        tableWidth = 13, //how many columns wide per team
        tableHeight = 10, //how many rows per team
        ss = SpreadsheetApp.getActiveSpreadsheet(), //loads the spreadsheet into code

        //we also want to get all the car numbers into memory from the entry list so lets do that
        sheetEntryList = ss.getSheetByName(sheetNameEntryList),
        sheetEntryListLength = sheetEntryList.getMaxRows(),
        rangeEntryListCarNumbers = sheetEntryList.getRange(1, entryListCarNumberColumn, sheetEntryListLength), //(row, column, numRows)
        dataCarNumbers = rangeEntryListCarNumbers.getValues(),

        //we also need to know where the spotters guide is and which column / row we are currently editing
        sheetSpotterGuide = ss.getSheetByName(sheetNameSpotterGuide),
        tableCurrentColumn = 1; //for checking how far accross the table we are before dropping to next row

    removeTeamFromSpotterGuide(); //start by clearing all numbers from the sheet before re-adding them

    //sort the data into numerical order
    dataCarNumbers.sort(function (a, b) { return a - b; });

    //now we have data into the variables, we will now loop through the car numbers and populate the doc
    for (var currentCar = 0; currentCar < dataCarNumbers.length; currentCar++)
    /*
    the for loop above creates a variable called currentCar and sets this to 0, this is because the array of car numbers starts at 0
    it then checks if this is less than the length of the dataCarNumbers array and if it is it will execute the following code
    finally, once executed it will increment the currentCar variable by 1, currentCar++ is the same as currentCar = currentCar + 1
    */ {
        if (isNaN(dataCarNumbers[currentCar]) == false && dataCarNumbers[currentCar] > 0)
        /*check that the car in the array we are looping is a number (isNaN = is Not a Number, if text it's true, if number it's false)
        and also check that the number is > 0*/ {
            //now we've found a new number to add, we need now to add it to the currently selected cell
            sheetSpotterGuide.getRange(sheetCurrentRow, sheetCurrentColumn).setValue(dataCarNumbers[currentCar]);

            //now this is set, we need to move to the next cell to edit
            if (tableCurrentColumn < tableColumnCount)
            //if we haven't just edited the last column, move right
            {
                tableCurrentColumn++;
                sheetCurrentColumn += (tableWidth + 1); //an extra 1 for the border
            }
            else
            //else if we have just edited the last column, move down and left
            {
                tableCurrentColumn = 1;
                sheetCurrentColumn -= ((tableWidth + 1) * (tableColumnCount - 1));
                sheetCurrentRow += (tableHeight + 1);
            }
        }
    }
}

function removeTeamFromSpotterGuide() {
    //firstly we'll start with a list of variables to tell the script what your data looks like
    var sheetCurrentRow = 1, //set to the first car numbers row
        sheetCurrentColumn = 8, //set to first car numbers column, this is H but numbers are easier to work with in scripts
        tableColumnCount = 2, //how many teams displayed on each row basically
        tableWidth = 13, //how many columns wide per team
        tableHeight = 10, //how many rows per team
        ss = SpreadsheetApp.getActiveSpreadsheet(), //loads the spreadsheet into code

        //we also need to know where the spotters guide is and which column / row we are currently editing
        sheetSpotterGuide = ss.getSheetByName(sheetNameSpotterGuide),
        tableCurrentColumn = 1, //for checking how far accross the table we are before dropping to next row
        sheetSpotterGuideRows = sheetSpotterGuide.getMaxRows();

    while (sheetCurrentRow < sheetSpotterGuideRows) //run a loop for as long as we haven't ran off the bottom of the doc
    {
        //remove first number
        sheetSpotterGuide.getRange(sheetCurrentRow, sheetCurrentColumn).setValue("");

        //now this is set, we need to move to the next cell to edit
        if (tableCurrentColumn < tableColumnCount)
        //if we haven't just edited the last column, move right
        {
            tableCurrentColumn++;
            sheetCurrentColumn += (tableWidth + 1); //an extra 1 for the border
        }
        else
        //else if we have just edited the last column, move down and left
        {
            tableCurrentColumn = 1;
            sheetCurrentColumn -= ((tableWidth + 1) * (tableColumnCount - 1));
            sheetCurrentRow += (tableHeight + 1);
        }
    }
}

function cloneSpottersLayout() {
    //firstly we'll start with a list of variables to tell the script what your data looks like
    var sheetCurrentRow = 1, //set to the first car numbers row
        sheetCurrentColumn = 8, //set to first car numbers column, this is H but numbers are easier to work with in scripts
        tableColumnCount = 2, //how many teams displayed on each row basically
        tableWidth = 13, //how many columns wide per team
        tableHeight = 10, //how many rows per team
        ss = SpreadsheetApp.getActiveSpreadsheet(), //loads the spreadsheet into code

        //we also need to know where the spotters guide is and which column / row we are currently editing
        sheetSpotterGuide = ss.getSheetByName(sheetNameSpotterGuide),
        tableCurrentColumn = 1, //for checking how far accross the table we are before dropping to next row
        sheetSpotterGuideRows = sheetSpotterGuide.getMaxRows(),
        rangeTableEntryMaster = sheetSpotterGuide.getRange(sheetCurrentRow, sheetCurrentColumn, tableHeight, tableWidth);

    while (sheetCurrentRow < sheetSpotterGuideRows) //run a loop for as long as we haven't ran off the bottom of the doc
    {
        //copy range from table master to current selection
        rangeTableEntryMaster.copyTo(sheetSpotterGuide.getRange(sheetCurrentRow, sheetCurrentColumn, tableHeight, tableWidth));

        //now this is set, we need to move to the next cell to edit
        if (tableCurrentColumn < tableColumnCount)
        //if we haven't just edited the last column, move right
        {
            tableCurrentColumn++;
            sheetCurrentColumn += (tableWidth + 1); //an extra 1 for the border
        }
        else
        //else if we have just edited the last column, move down and left
        {
            tableCurrentColumn = 1;
            sheetCurrentColumn -= ((tableWidth + 1) * (tableColumnCount - 1));
            sheetCurrentRow += (tableHeight + 1);
        }
    }
    addTeamToSpotterGuide(); //re-add correct numbers
}


function updateQualifyingDrivers() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var QualiSheet = ss.getSheetByName("Qualifying");
    var entryList = ss.getSheetByName("Entry List");
    var entryListLength = entryList.getMaxRows() - entryListHeaders;
    var QualiDrivers = QualiSheet.getRange("A2:A" + QualiSheet.getMaxRows()).getValues();
    var EnteredDrivers1 = entryList.getRange(entryListHeaders + 1, entryListLicenceNumberColumnNum1, entryListLength).getValues();
    var EnteredDrivers2 = entryList.getRange(entryListHeaders + 1, entryListLicenceNumberColumnNum2, entryListLength).getValues();
    var EnteredDrivers3 = entryList.getRange(entryListHeaders + 1, entryListLicenceNumberColumnNum3, entryListLength).getValues();
    var EnteredDrivers4 = entryList.getRange(entryListHeaders + 1, entryListLicenceNumberColumnNum4, entryListLength).getValues();
    var EnteredDrivers5 = entryList.getRange(entryListHeaders + 1, entryListLicenceNumberColumnNum5, entryListLength).getValues();
    var EnteredDrivers6 = entryList.getRange(entryListHeaders + 1, entryListLicenceNumberColumnNum6, entryListLength).getValues();
    var EnteredDrivers = new Array(0);
    var driverFound = false;
    //EnteredDrivers[0] = 0;
    for (var i = 0; i <= EnteredDrivers1.length; i++) {
        if (EnteredDrivers1[i] > 0) {
            EnteredDrivers.push(EnteredDrivers1[i]);
        }
        if (EnteredDrivers2[i] > 0) {
            EnteredDrivers.push(EnteredDrivers2[i]);
        }
        if (EnteredDrivers3[i] > 0) {
            EnteredDrivers.push(EnteredDrivers3[i]);
        }
        if (EnteredDrivers4[i] > 0) {
            EnteredDrivers.push(EnteredDrivers4[i]);
        }
        if (EnteredDrivers5[i] > 0) {
            EnteredDrivers.push(EnteredDrivers5[i]);
        }
        if (EnteredDrivers6[i] > 0) {
            EnteredDrivers.push(EnteredDrivers6[i]);
        }
    }
    for (var i = 0; i < EnteredDrivers.length; i++) {
        driverFound = false;
        for (var j = 0; j < QualiDrivers.length; j++) {
            if (parseInt(EnteredDrivers[i]) == parseInt(QualiDrivers[j])) {
                driverFound = true;
                break;
            }
        }
        if (driverFound == false) {
            for (var k = 0; k < QualiDrivers.length; k++) {
                if (QualiDrivers[k] == 0) {
                    QualiDrivers[k] = EnteredDrivers[i];
                    break;
                }
            }
        }
    }
    QualiSheet.getRange("A2:A" + QualiSheet.getMaxRows()).setValues(QualiDrivers);
}


function AddDriverByLicenceNumber() {
    // Firstly check which sheet is being edited
    // If editing Entry List do the following
    // Only run if it's on data and not the headers
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    if (sheet.getName() == "Entry List" && sheet.getActiveCell().getRow() >= 5 && sheet.getActiveCell().getColumn() >= 5 && sheet.getActiveCell().getColumn() <= 19) {
        var enteredLicenceNumber = Browser.inputBox('Add New Driver', 'Enter licence number to be added', Browser.Buttons.OK_CANCEL);
        if (enteredLicenceNumber !== "" && enteredLicenceNumber !== "cancel") {
            AddDriver(enteredLicenceNumber, "Not Searching For This");
        }
    } else {
        Browser.msgBox("The cursor is in the wrong place, you must be on the Entry list, and have a row selected which is a valid team, and a column which is a valid driver number (inside E5:S1000)")
    }
}

function AddDriverByGamertag() {
    // Firstly check which sheet is being edited
    // If editing Entry List do the following
    // Only run if it's on data and not the headers
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    if (sheet.getName() == "Entry List" && sheet.getActiveCell().getRow() >= 5 && sheet.getActiveCell().getColumn() >= 5 && sheet.getActiveCell().getColumn() <= 19) {
        var enteredGamertag = Browser.inputBox('Add New Driver', 'Enter gamertag to be added (not case sensitive)', Browser.Buttons.OK_CANCEL);
        if (enteredGamertag !== "" && enteredGamertag !== "cancel") {
            AddDriver("Not Searching For This", enteredGamertag);
        }
    } else {
        Browser.msgBox("The cursor is in the wrong place, you must be on the Entry list, and have a row selected which is a valid team, and a column which is a valid driver number (inside E5:S1000)")
    }
}

function AddDriver(EnteredLicenceNumber, EnteredUserName) {
    // Make sure they have a specific driver selected
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        selectedDriver = 0,
        activeCell = ss.getActiveSheet().getActiveCell(),
        selectedRow = activeCell.getRow(),
        selectedColumn = activeCell.getColumn(),
        sheetDrivers = ss.getSheetByName("Drivers"),
        sheetEntryList = ss.getSheetByName("Entry List");

    switch (selectedColumn) {
        case entryListLicenceNumberColumnNum1:
        case entryListLicenceNumberColumnNum1 + 1:
            selectedDriver = 1;
            break;
        case entryListLicenceNumberColumnNum2:
        case entryListLicenceNumberColumnNum2 + 1:
            selectedDriver = 2;
            break;
        case entryListLicenceNumberColumnNum3:
        case entryListLicenceNumberColumnNum3 + 1:
            selectedDriver = 3;
            break;
        case entryListLicenceNumberColumnNum4:
        case entryListLicenceNumberColumnNum4 + 1:
            selectedDriver = 4;
            break;
        case entryListLicenceNumberColumnNum5:
        case entryListLicenceNumberColumnNum5 + 1:
            selectedDriver = 5;
            break;
        case entryListLicenceNumberColumnNum6:
        case entryListLicenceNumberColumnNum6 + 1:
            selectedDriver = 6;
            break;
    }

    if (selectedDriver !== 0) {
        var entryListLastRow = sheetEntryList.getMaxRows();
        var driversLastRow = sheetDrivers.getMaxRows();
        var driverFound = 0;

        // Import full list of licence numbers to memory to check (quicker than getting from sheet one at a time)
        var dataDriversLicenceNumber = sheetDrivers.getRange(3, driversLicenceNumberColumn, driversLastRow, 1).getValues();
        var dataDriversUserName = sheetDrivers.getRange(3, driversUserNameColumn, driversLastRow, 1).getValues();

        // Loop through all imported data and search for the driver
        for (var i = 0; i <= (driversLastRow - 3); i++) {
            // If driver list is empty then stop searching
            if (dataDriversLicenceNumber[i][0] == "" || dataDriversUserName[i][0] == "") {
                break;
            }
            // If driver found, add them to the doc
            else if (dataDriversLicenceNumber[i][0] == EnteredLicenceNumber || dataDriversUserName[i][0].toLowerCase() == EnteredUserName.toLowerCase()) {
                driverFound = i + 3;
                break;
            }
        }

        // If the driver can be found, add it to the doc, else return not found
        if (driverFound > 0) {
            var dataEntryListLicenceNumber = sheetEntryList.getRange(entryListHeaders + 1, driversLicenceNumberColumn, entryListLastRow - entryListHeaders + 1, 10).getValues();
            for (var r = 0; r <= (entryListLastRow - 5); r++) {
                for (var c = 0; c <= 10; c += 2) {
                    if (dataEntryListLicenceNumber[r][c] == dataDriversLicenceNumber[driverFound - 3][0]) {
                        Browser.msgBox("This driver is already listed " + driverFound);
                        break;
                    }
                }
            }
            // PUT CODE HERE TO ADD DRIVER TO THE DOC DEPENDING WHICH DRIVER THEY HAVE SELECTED
            switch (selectedDriver) {
                case 1:
                    sheetEntryList.getRange(selectedRow, entryListLicenceNumberColumnNum1).setValue(dataDriversLicenceNumber[driverFound - 3][0]);
                    Browser.msgBox("Driver 1 added successfully");
                    break;
                case 2:
                    sheetEntryList.getRange(selectedRow, entryListLicenceNumberColumnNum2).setValue(dataDriversLicenceNumber[driverFound - 3][0]);
                    Browser.msgBox("Driver 2 added successfully");
                    break;
                case 3:
                    sheetEntryList.getRange(selectedRow, entryListLicenceNumberColumnNum3).setValue(dataDriversLicenceNumber[driverFound - 3][0]);
                    Browser.msgBox("Driver 3 added successfully");
                    break;
                case 4:
                    sheetEntryList.getRange(selectedRow, entryListLicenceNumberColumnNum4).setValue(dataDriversLicenceNumber[driverFound - 3][0]);
                    Browser.msgBox("Driver 4 added successfully");
                    break;
                case 5:
                    sheetEntryList.getRange(selectedRow, entryListLicenceNumberColumnNum5).setValue(dataDriversLicenceNumber[driverFound - 3][0]);
                    Browser.msgBox("Driver 5 added successfully");
                    break;
                case 6:
                    sheetEntryList.getRange(selectedRow, entryListLicenceNumberColumnNum6).setValue(dataDriversLicenceNumber[driverFound - 3][0]);
                    Browser.msgBox("Driver 6 added successfully");
                    break;
            }
        }
        else {
            Browser.msgBox("Driver Not Found");
        }
    }
}

/*


function SortSheet(){
var LastRow = SpreadsheetApp.getActiveSheet().getMaxRows();
SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Drivers").getRange("A4:" + LastColumn + LastRow).sort([{column: CarNumberColumnNum, ascending: true}, {column: LicenceNumberColumnNum, ascending: true}]);
}

*/

//59 65 79 2c 20 61 20 66 65 6c 6c 6f 77 20 6e 65 72 64 2c 20 49 20 77 72 6f 74 65 20 74 68 69 73 20 6f 6e 20 32 36 2f 30 35 2f 32 30 31 36 2c 20 49 27 6d 20 67 75 65 73 73 69 6e 67 20 69 74 27 73 20 6e 6f 77 20 74 68 65 20 79 65 61 72 20 32 30 32 30 2c 20 70 6c 65 61 73 65 20 65 2d 6d 61 69 6c 20 6d 65 20 74 6f 20 74 65 6c 6c 20 6d 65 20 68 6f 77 20 77 72 6f 6e 67 20 49 20 61 6d 2c 20 4a 61 79 2e
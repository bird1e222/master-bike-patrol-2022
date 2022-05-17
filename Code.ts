// function renameSheet(
//   let currentName = "Sommaire avec ajustement";
//   let newName = `Sommaire avec
//     ajustement`
// ) {
//   // Get a reference to the sheet using its existing name
//   // and then rename it using the setName() method.
//   SpreadsheetApp.getActive()
//     .getActiveSheet()
//     .setName(newName);
// }

var gasObjects = {
  activeSpreadsheet: function() {
    // eslint-disable-next-line no-undef
    return SpreadsheetApp.getActiveSpreadsheet();
  },
  activeSheet: function() {
    // eslint-disable-next-line no-undef
    return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  },
  developper: function() {
    return ["dev.patrouilleur.ds@gmail.com"];
  },
  activeUser: function() {
    // eslint-disable-next-line no-undef
    return Session.getActiveUser().getEmail();
  },
  effectiveUser: function() {
    // eslint-disable-next-line no-undef
    return Session.getActiveUser().getEmail();
  },
  targetFolder: function() {
    // eslint-disable-next-line no-undef
    return DriveApp.getFolderById("1S5livaq1_Dn_81ivt8_eV7M8vUkqJ3Kv");
  },
  templateFile: function() {
    // eslint-disable-next-line no-undef
    return DriveApp.getFileById("1lKEiCFs2bkIHiseXn4IR5Ob5AUkQ4JbsX2dkYhJDpkQ");
  },
  ui: function() {
    // eslint-disable-next-line no-undef
    return SpreadsheetApp.getUi();
  }
};

const COLUMN_HEADERS_OFFSET = 2;
const lastColumn = gasObjects.activeSheet().getLastColumn();
const frozenRows = gasObjects.activeSheet().getFrozenRows();
const columnHeaderRange = SpreadsheetApp.getActiveSheet().getRange(
  frozenRows,
  2,
  1,
  lastColumn - 1
);

console.log(`columnHeaderRange = ${columnHeaderRange}`);
console.log(
  `columnHeaderRange.getA1Notation() = ${columnHeaderRange.getA1Notation()}`
);

const columnHeaderRangeValues = columnHeaderRange.getValues();

/**
 * ColumnHeader object contains selected column's number from Timesheet
 * This is useful to establish which data is in which column and therfore
 * not to have to hard code column number, should columns be added, moved or deleted.
 *
 * @class ColumnHeaders
 */
class ColumnHeaders {
  _fullNameIndex: string;
  _shortNameIndex: string;
  _firstNameIndex: string;
  _lastNameIndex: string;
  _teamLeaderIndex: string;
  _personalEmailIndex: string;
  _patrollEmailIndex: string;
  _passwordIndex: string;
  _phoneHomeIndex: string;
  _phoneMobileIndex: string;
  _postalAddressIndex: string;
  _versionDeployedIndex: string;
  _documentIDIndex: string;
  _documentLinkIndex: string;
  _deployedIndex: string;
  _documentNameIndex: string;
  _columnHeaderValues: any;
  /**
   *Creates an instance of ColumnHeaders.
   * @memberof ColumnHeaders
   */
  constructor() {
    this._columnHeaderValues = columnHeaderRangeValues;

    this._fullNameIndex = this._columnHeaderValues[0].indexOf(`Nom
complet`);
    console.log(`this._fullNameIndex = ${this._fullNameIndex}`);

    //     this._fullName =
    //       this._columnHeaderValues[0].indexOf(`Nom
    // complet`);

    this._shortNameIndex = this._columnHeaderValues[0].indexOf(`Nom
abrégé`);

    this._firstNameIndex = this._columnHeaderValues[0].indexOf(`Prénom`);

    this._lastNameIndex = this._columnHeaderValues[0].indexOf(`Nom`);

    this._teamLeaderIndex = this._columnHeaderValues[0].indexOf(`Chef
d'équipe`);

    this._personalEmailIndex = this._columnHeaderValues[0].indexOf(`Courriel
personnel`);

    this._patrollEmailIndex = this._columnHeaderValues[0].indexOf(`Courriel
patrouilleur`);

    this._passwordIndex = this._columnHeaderValues[0].indexOf(`Mot de passe`);

    this._phoneHomeIndex = this._columnHeaderValues[0].indexOf(`Téléphone
résidence`);

    this._phoneMobileIndex = this._columnHeaderValues[0].indexOf(`Cellulaire`);

    this._postalAddressIndex = this._columnHeaderValues[0].indexOf(`Adresse`);

    this._versionDeployedIndex = this._columnHeaderValues[0].indexOf(`Version
Déployée`);
    console.log(`this._columnHeaderValues[0] = ${this._columnHeaderValues[0]}`);

    this._documentIDIndex = this._columnHeaderValues[0].indexOf(`Document ID`);

    this._documentLinkIndex = this._columnHeaderValues[0].indexOf(`Lien vers le
document`);

    this._deployedIndex = this._columnHeaderValues[0].indexOf(`Déployé`);

    this._documentNameIndex = this._columnHeaderValues[0].indexOf(
      `Nom du document`
    );
  }

  get fullNameIndex() {
    return this._fullNameIndex;
  }

  get shortNameIndex() {
    return this._shortNameIndex;
  }

  get firstNameIndex() {
    return this._firstNameIndex;
  }

  get lastNameIndex() {
    return this._lastNameIndex;
  }

  get teamLeaderIndex() {
    return this._teamLeaderIndex;
  }

  get personalEmailIndex() {
    return this._personalEmailIndex;
  }

  get patrollEmailIndex() {
    return this._patrollEmailIndex;
  }

  get passwordIndex() {
    return this._passwordIndex;
  }

  get phoneHomeIndex() {
    return this._phoneHomeIndex;
  }

  get phoneMobileIndex() {
    return this._phoneMobileIndex;
  }

  get postalAddressIndex() {
    return this._postalAddressIndex;
  }

  get versionDeployedIndex() {
    console.log(`this._versionDeployedIndex = ${this._versionDeployedIndex}`);
    return this._versionDeployedIndex;
  }

  get documentIDIndex() {
    return this._documentIDIndex;
  }

  get documentLinkIndex() {
    return this._documentLinkIndex;
  }

  get deployedIndex() {
    return this._deployedIndex;
  }

  get documentNameIndex() {
    return this._documentNameIndex;
  }

  get columnHeaderValues() {
    console.log(`this._columnHeaderValues) = ${this._columnHeaderValues}`);
    return this._columnHeaderValues;
  }
}

const COL_HEADERS = new ColumnHeaders();

/* function addEditor() {
  let userEmail = "michel.sabourin.patrouilleur.ds@Gmail.com";
  let fileId = "16aEkBuu0gFqzDHx1SpyynBeuI3gSB4XlXNaixA4PrcE";
  let file = DriveApp.getFileById(fileId);

  file.addEditor(userEmail);
}
 */
const ACCENTED =
  "ÀÁÂÃÄÅàáâãäåÒÓÔÕÕÖØòóôõöøÈÉÊËèéêëðÇçÐÌÍÎÏìíîïÙÚÛÜùúûüÑñŠšŸÿýŽž-";
const REGULAR =
  "AAAAAAaaaaaaOOOOOOOooooooEEEEeeeeeCcDIIIIiiiiUUUUuuuuNnSsYyyZz.";
const REGEXP = new RegExp("[" + ACCENTED + "]", "g");

function replaceDiacritics(str) {
  function replace(match) {
    var p = ACCENTED.indexOf(match);
    return REGULAR[p];
  }
  return str.replace(REGEXP, replace);
}

/**
 * Create custom menu
 *
 */
/* exported onOpen() */
// eslint-disable-next-line no-unused-vars
function onOpen() {
  let ui = gasObjects.ui();
  ui.createMenu("Patrouille à vélo")
    .addItem("Créer feuille de temps pour patrouilleur", "menu.createTimeSheet")
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu("Advanced")
        .addItem("Rename current sheet", "renameCurrentSheet")
    )
    .addSeparator()
    .addSubMenu(ui.createMenu("Sub-menu").addItem("Second item", "menu.item2"))
    .addToUi();
}

function getLinkTo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  let ssID = ss.getId();
  var sheet = ss.getSheetByName("Heures par patrouilleur");
  let sheetID = sheet.getSheetId();

  let linkToRange = sheet.getRange("AD3:AD45");
  let linkToValues = linkToRange.getValues();
  let hyperlinkRange = sheet.getRange("AC3:AC45");
  let hyperlinkValues = hyperlinkRange.getValues();

  for (let row in hyperlinkValues) {
    for (let col in hyperlinkValues[row]) {
    }
    hyperlinkValues[
      row
    ][0] = `https://docs.google.com/spreadsheets/d/${ssID}/edit#gid=${sheetID}&range=${sheet
      .getRange(Number(row) + 3, 30)
      .getA1Notation()}`;
  }

  hyperlinkRange.setValues(hyperlinkValues);

  /*  sheet
    .getRange("AB3:AB45")
    .setValue(
      `https://docs.google.com/spreadsheets/d/${ssID}/edit#gid=${sheetID}&range=${sheet
        .getRange("AC3")
        .getA1Notation()}`
    ); */
}

function createHyperLinkWithFormula() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sourcesheet = "Sheet1";
  let source = ss.getSheetByName(sourcesheet);
  let target = ss.getSheetByName("Heures par patrouilleur");
  let cell = `&AC3`;
  let formula =
    '=HYPERLINK("https://docs.google.com/spreadsheets/d/' +
    "1F-7QzOJwxXOUs7iBmuedBTejvesP2VjyPm0uTIy4z5M" +
    "/edit#gid=1900061308&range=AC3" +
    '","' +
    "Michel S." +
    '")';
  let link = `https://docs.google.com/spreadsheets/d/1F-7QzOJwxXOUs7iBmuedBTejvesP2VjyPm0uTIy4z5M/edit#gid=1900061308&range="`;
  let text = "Michel";
  let value = `=HYPERLINK("${link}${cell}, "${text}")`;
  let sheet = SpreadsheetApp.getActiveSheet();
  let range = sheet.getRange("K3");
  range.setFormula(formula);
  //range.setValue(value);
}

{
  //let value = SpreadsheetApp.newRichTextValue().setText(text).setLinkUrl(linkUrl)      //setLinkUrl(link).build();
  //SpreadsheetApp.getActiveSheet().getRange('A1').setRichTextValue(value);
}

function renameCurrentSheet() {
  let sheet = SpreadsheetApp.getActiveSheet();
  sheet.setName(` Heures et Interventions
             par Parcours`);
}

// eslint-disable-next-line no-unused-vars
var menu = {
  /**
   * Create a new spreadsheet for this member.
   * Register this member in the master spreadsheet.
   */
  createTimeSheet: function() {
    let activeSheet = gasObjects.activeSheet();
    let currentCell = activeSheet.getCurrentCell();
    let httpAddressPartOne = "https://docs.google.com/spreadsheets/d/";
    let httpAddressPartThree = "/edit#gid=2083169682";
    let newDate = new Date();
    let options = {
      weekday: "long",
      year: "numeric",
      month: "short",
      day: "2-digit",
      hour: "numeric",
      minute: "numeric"
    };
    /* @ts-ignore */
    const dateTimeFormat = new Intl.DateTimeFormat("fr", options);
    const [
      { value: weekday },
      ,
      { value: day },
      ,
      { value: month },
      ,
      { value: year },
      ,
      { value: hour },
      ,
      { value: minute }
    ] = dateTimeFormat.formatToParts(newDate);

    console.log(`currentCell.getRow() = ${currentCell.getRow()}`);
    console.log(
      `COL_HEADERS.versionDeployedIndex = ${COL_HEADERS.versionDeployedIndex}`
    );
    console.log(
      `COL_HEADERS.documentNameIndex = ${COL_HEADERS.documentNameIndex}`
    );
    console.log(
      `Number(COL_HEADERS.versionDeployedIndex) + 1 = ${Number(
        COL_HEADERS.versionDeployedIndex
      ) + 1}`
    );
    console.log(
      `Number(COL_HEADERS.documentNameIndex) - Number(COL_HEADERS.versionDeployedIndex) + 1 = ${Number(
        COL_HEADERS.documentNameIndex
      ) -
        Number(COL_HEADERS.versionDeployedIndex) +
        1}`
    );

    // Retrieve current values of selected patrol (current row)
    let currentRange = activeSheet.getRange(
      currentCell.getRow(),
      Number(COL_HEADERS.versionDeployedIndex) + 2,
      1,
      Number(COL_HEADERS.documentNameIndex) -
        Number(COL_HEADERS.versionDeployedIndex) +
        1
    );
    let currValues = currentRange.getValues();

    console.log(
      `Number(COL_HEADERS.patrollEmailIndex + 1) = ${Number(
        COL_HEADERS.patrollEmailIndex + 1
      )}`
    );
    // Retrieve current values of selected patrol (current row)
    let currentEmailCol = Number(COL_HEADERS.patrollEmailIndex) + 2;
    console.log(`currentEmailCol) = ${currentEmailCol}`);
    let currentEmailRange = activeSheet.getRange(
      currentCell.getRow(),
      currentEmailCol,
      1,
      1
    );
    let currEmailValue = currentEmailRange.getValue();
    console.log(
      `currentEmailRange.getA1Notation() = ${currentEmailRange.getA1Notation()}`
    );
    console.log(`currEmailValue = ${currEmailValue}`);

    console.log(`currValues[0] = ${currValues[0]}`);
    console.log(`currValues[0].length = ${currValues[0].length}`);
    console.log(`currValues[0][0].length = ${currValues[0][0].length}`);
    console.log(
      `currentRange.getA1Notation() = ${currentRange.getA1Notation()}`
    );

    console.log(`currValues.length = ${currValues.length}`);
    console.log(`currValues[0].length = ${currValues[0].length}`);

    // Create array for new content to update the master file
    //let newValues = [[""]];

    // Copy content of current Master File's row into the new array
    console.log(`currValues[0] = ${currValues[0]}`);
    let newValues = currValues;
    console.log(`newValues[0].length = ${newValues[0].length}`);
    console.log(`newValues[0] = ${newValues[0]}`);

    // Get and open template file
    let templateFile = gasObjects.templateFile();
    let templateSS = SpreadsheetApp.open(templateFile);

    // Initialize values to be written in Master File's row or in patrol's new spreadsheet
    //console.log(`currValues[0][COL_HEADERS.fullNameIndex] = ${currValues[0][COL_HEADERS.fullNameIndex]}`);
    let memberNameRange = activeSheet.getRange(
      currentCell.getRow(),
      Number(COL_HEADERS.fullNameIndex) + 2,
      1,
      1
    );

    console.log(
      `memberNameRange.getA1Notation() = ${memberNameRange.getA1Notation()}`
    );
    console.log(`memberNameRange.getValue() = ${memberNameRange.getValue()}`);

    let memberName = memberNameRange.getValue();
    console.log(`memberName = ${memberName}`);
    let docMiddleName = "- Feuille de temps 2022 -";
    // let docVersion = currValues[0][COL_HEADERS.versionDeployedIndex];
    let templateFileName = templateSS.getName();
    let templateDocVer = templateFileName.substr(
      templateFileName.length - 4,
      4
    );
    let documentTitle = `${memberName} ${docMiddleName} ${templateDocVer}`;

    // Make a copy of the template to create the new patrol's file in the target folder
    let newFile = templateFile.makeCopy(
      documentTitle,
      gasObjects.targetFolder()
    );

    console.log(` Before Drive.Permissions.insert Bruce ${memberName}`);

    // Share new file with patrol and team leaders, without notification
    Drive.Permissions.insert(
      {
        emailAddress: "bruce.porter.patrouilleur.ds@gmail.com",
        role: "writer",
        type: "user",
        value: "bruce.porter.patrouilleur.ds@gmail.com"
      },
      newFile.getId(),
      {
        sendNotificationEmails: "false"
      }
    );

    console.log(` After Drive.Permissions.insert Bruce ${memberName}`);

    Drive.Permissions.insert(
      {
        role: "writer",
        type: "user",
        value: "michel.gaudreau.patrouilleur.ds@gmail.com"
      },
      newFile.getId(),
      {
        sendNotificationEmails: "false"
      }
    );

    Drive.Permissions.insert(
      {
        role: "writer",
        type: "user",
        value: "sylvain.roy.patrouilleur.ds@gmail.com"
      },
      newFile.getId(),
      {
        sendNotificationEmails: "false"
      }
    );

    console.log(
      ` COL_HEADERS.patrollEmailIndex =  ${COL_HEADERS.patrollEmailIndex}`
    );
    console.log(
      ` currValues[0][COL_HEADERS.patrollEmailIndex] =  ${
        currValues[0][COL_HEADERS.patrollEmailIndex]
      }`
    );
    //let memberEmail = `${currValues[0][COL_HEADERS.patrollEmailIndex]}`
    let courrielPatrouilleur = `${
      currValues[0][COL_HEADERS.patrollEmailIndex]
    }`;
    console.log(` courrielPatrouilleur =  ${courrielPatrouilleur}`);

    console.log(` currValues[0] =  ${currValues[0]}`);
    console.log(` currEmailValue =  ${currEmailValue}`);

    try {
      Drive.Permissions.insert(
        {
          role: "writer",
          type: "user",
          value: currEmailValue
        },
        newFile.getId(),
        {
          sendNotificationEmails: "false"
        }
      );
    } catch (error) {}

    // Share new file with patrol and team leaders
    //let editorEmail = ["bruce.porter.patrouilleur.ds@gmail.com"]
    /*     let editorEmail = [
      `${currValues[0][COL_HEADERS.patrollEmailIndex]}`,
      "bruce.porter.patrouilleur.ds@gmail.com",
      "michel.gaudreau.patrouilleur.ds@gmail.com",
      "sylvain.roy.patrouilleur.ds@gmail.com",
      "dev.patrouilleur.ds@gmail.com"
    ];
 */
    //    console.log(`editorEmail = ${editorEmail}`);

    /*    try {
      newFile.addEditors(editorEmail);
    } catch (error) {
      console.log(`error for newFile.addEditors(editorEmail) = ${error}`);
    }
 */
    // Continue to feed the array with new content
    newValues[0][
      Number(COL_HEADERS.versionDeployedIndex) - 13
    ] = templateDocVer;
    newValues[0][Number(COL_HEADERS.documentIDIndex) - 13] = newFile.getId();
    newValues[0][
      Number(COL_HEADERS.documentLinkIndex) - 13
    ] = `${httpAddressPartOne}${newFile.getId()}${httpAddressPartThree}`;
    newValues[0][
      Number(COL_HEADERS.deployedIndex) - 13
    ] = `${weekday}, ${day} ${month} ${year}  ${hour}:${minute}`;
    newValues[0][
      Number(COL_HEADERS.documentNameIndex) - 13
    ] = newFile.getName();

    console.log(`newValues.length = ${newValues.length}`);

    for (let m = 0; m < newValues.length; m++) {
      for (let n = 0; n < newValues[m].length; n++) {
        console.log(`newValues[m][n] = ${newValues[m][n]}`);
      }
    }
    console.log(`newValues = ${newValues}`);

    // Update the master file with new content
    console.log(`newValues = ${newValues}`);
    currentRange.setValues(newValues);

    // Update the new patrol's file with personal information
    let newFileSS = SpreadsheetApp.open(newFile);

    // The code below logs the name of the first named range.
    /*     let namedRanges = newFileSS.getNamedRanges();
    Logger.log(namedRanges.length);
    for (var i = 0; i < namedRanges.length; i++) {
//      Logger.log(namedRanges[i].getName());
      console.log(`i = ${i}  /  namedRanges[i].getName() = ${namedRanges[i].getName()}`);
    }
 */

    console.log(
      `newFileSS.getNamedRanges().values = ${newFileSS.getNamedRanges().values}`
    );
    console.log(`newFileSS.getName() = ${newFileSS.getName()}`);
    let postalAddressRange = newFileSS.getRangeByName("_patrolAddress");
    console.log(
      `newFileSS.getRangeByName("_patrolAddress") = ${newFileSS.getRangeByName(
        "_patrolAddress"
      )}`
    );
    console.log(`postalAddressRange = ${postalAddressRange}`);
    let fullNameRange = newFileSS.getRangeByName("_patrolFullName");
    let shortNameRange = newFileSS.getRangeByName("_patrolShortName");
    // let geoCodeRange = newFileSS.getRangeByName("_userAddressGeoCode");
    // let distanceRange = newFileSS.getRangeByName("_pathMilestoneDistanceAPI");

    console.log(
      `currValues[0][COL_HEADERS.postalAddressIndex] = ${
        currValues[0][COL_HEADERS.postalAddressIndex]
      }`
    );
    console.log(`before postalAddressRange.setValue`);
    postalAddressRange.setValue(
      `${currValues[0][COL_HEADERS.postalAddressIndex]}`
    );
    console.log(`after postalAddressRange.setValue`);
    console.log(`before currValues[0][COL_HEADERS.fullNameIndex]`);
    console.log(
      `currValues[0][COL_HEADERS.fullNameIndex] = ${
        currValues[0][COL_HEADERS.fullNameIndex]
      }`
    );
    fullNameRange.setValue(`${currValues[0][COL_HEADERS.fullNameIndex]}`);
    console.log(`after currValues[0][COL_HEADERS.fullNameIndex]`);
    shortNameRange.setValue(`${currValues[0][COL_HEADERS.shortNameIndex]}`);

    // Copy geocode formula and paste as value
    // let geocodeFormula = geoCodeRange.getValue();
    // if (geocodeFormula != "#ERROR!") {
    //   geoCodeRange.setValue(`${geocodeFormula}`);

    //   // Copy computed distance formulas and paste them as values
    //   let distanceFormula = distanceRange.getValue();
    //   distanceRange.setValue(`${distanceFormula}`);
    // }
  },

  item2: function() {
    gasObjects.ui().alert("You clicked: Second item");
  }
};

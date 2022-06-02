/**
 * @file Contains javascript code attached to Google sheets
 * "_Master - Feuilles de temps 2022".
 * The code is rather Typescript code by is converted to javascript when we
 * use "clasp" terminal command to upload it to the file.
 *
 * Link to the spreadsheet
 * REF https://docs.google.com/spreadsheets/d/1F-7QzOJwxXOUs7iBmuedBTejvesP2VjyPm0uTIy4z5M/edit?usp=sharing
 *
 * Link to the code repository (Github)
 * REF https://github.com/bird1e222/master-bike-patrol-2022.git
 *
 * @license MIT
 */

/**
 * Examples of available comments when using extensions "Better comments" and
 * "Todo Tree"
 */
/**
 * BUG This is a bug
 * WARN This is a warning
 * FIXME
 * REF https://primer.style/octicons/alert-24
 * TODO
 * REVIEW
 * ? Question
 * ! Important
 * * Highlight comment
 */

/**
 * Links to Visual Studio Code Market Place extensions
 * REF https://marketplace.visualstudio.com/items?itemName=aaron-bond.better-comments
 * REF https://marketplace.visualstudio.com/items?itemName=Gruntfuggly.todo-tree
 * Link to page describing of to use Better comments and Todo tree together.
 * REF https://tomfreudenberg.medium.com/vscode-extensions-better-comments-todo-tree-d5a9c46eaea5
 */

// eslint-disable-next-line @typescript-eslint/no-unused-vars
import { Console } from 'console';

/**
 * @description Constant that keeps together all Google Apps Script (gas) objects
 *  required to interact with the spreadsheet.
 * @global
 * @constant
 */
const gasObjects = {
  activeSpreadsheet: function () {
    // eslint-disable-next-line no-undef
    return SpreadsheetApp.getActiveSpreadsheet();
  },
  activeSheet: function () {
    // eslint-disable-next-line no-undef
    return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  },
  developper: function () {
    return ['dev.patrouilleur.ds@gmail.com'];
  },
  activeUser: function () {
    // eslint-disable-next-line no-undef
    return Session.getActiveUser().getEmail();
  },
  effectiveUser: function () {
    // eslint-disable-next-line no-undef
    return Session.getActiveUser().getEmail();
  },
  /** @description Returns identification (ID) of folder where template files are located */
  // WARN This ID is hard coded but will have to be changed when we start the next season
  // FIXME Find a way to change this ID when creating next season's template files
  targetFolder: function () {
    // eslint-disable-next-line no-undef
    return DriveApp.getFolderById('1S5livaq1_Dn_81ivt8_eV7M8vUkqJ3Kv');
  },
  /** @description Returns identification (ID) of template file
  // WARN This ID is hard coded but will have to be changed when we start next season
  // FIXME Find a way to change this ID when creating next season's template files
   */
  templateFile: function () {
    // eslint-disable-next-line no-undef
    return DriveApp.getFileById('1lKEiCFs2bkIHiseXn4IR5Ob5AUkQ4JbsX2dkYhJDpkQ');
  },
  ui: function () {
    // eslint-disable-next-line no-undef
    return SpreadsheetApp.getUi();
  }
};

/**
 * @constant
 * @global
 * @desc The next global constants are attributes of Timesheet spreadsheet
 */
const lastColumn = gasObjects.activeSheet().getLastColumn();
const frozenRows = gasObjects.activeSheet().getFrozenRows();
const columnHeaderRange = SpreadsheetApp.getActiveSheet().getRange(
  frozenRows,
  2,
  1,
  lastColumn - 1
);
const columnHeaderRangeValues = columnHeaderRange.getValues();

/**
 * @class
 * @classdesc The ColumnHeader class contains selected columns index of
 * Timesheet spreadsheet. This is essential to know which data is in which
 * column and, therfore, not have to hard code column number, should these
 * columns be added, moved or deleted.
 * @example
 * const COL_HEADERS = new ColumnHeaders();
 * COL_HEADERS.patrolEmailIndex // return index of column containing patrol's
 * email.
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
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  _columnHeaderValues: any;

  constructor() {
    this._columnHeaderValues = columnHeaderRangeValues;
    // WARN Do not change column header labels in spreadsheet. Otherwise the following properties won't match anymore
    // TODO Add this warning in spreadsheet as well
    this._fullNameIndex = this._columnHeaderValues[0].indexOf(`Nom
complet`);
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
    this._documentIDIndex = this._columnHeaderValues[0].indexOf(`Document ID`);
    this._documentLinkIndex = this._columnHeaderValues[0].indexOf(`Lien vers le
document`);
    this._deployedIndex = this._columnHeaderValues[0].indexOf(`Déployé`);
    this._documentNameIndex =
      this._columnHeaderValues[0].indexOf(`Nom du document`);
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
  get patrolEmailIndex() {
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
    return this._columnHeaderValues;
  }
}

/**
 * @description Creates an array containing the index and label of every
 * selected columns in the spreadsheet
 * @instance
 * @global
 */
const COL_HEADERS = new ColumnHeaders();

/**
 * @function
 * @description This function is called within Timesheet spreadsheet and
 * creates a custom menu. It is triggered every time the spreadsheet is opened.
 */
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function onOpen() {
  const ui = gasObjects.ui();
  ui.createMenu('Patrouille à vélo')
    .addItem('Créer feuille de temps pour patrouilleur', 'menu.createTimeSheet')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu('Advanced')
        .addItem('Rename current sheet', 'renameCurrentSheet')
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('Sub-menu').addItem('Second item', 'menu.item2'))
    .addToUi();
}

/**
 * @function
 * @description This function generates a [n x 1] array of hyperlinks which target
 *  cells in column "Patrouilleur", in summary table. (where n = number of patrols)
 *  These hyperlinks are used in detail table to jump to summary table.
 *  It is manually called in Timesheet spreadsheet.
 */
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function getLinkTo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssID = ss.getId();
  const sheet = ss.getSheetByName('Heures par patrouilleur');
  const sheetID = sheet.getSheetId();
  const frozenRows = sheet.getFrozenRows();

  /**
   * @description "hyperlinkToPatrolName" is a named range in spreadsheet's tab
   *  "Heures par patrouilleur". The range contains hyperlinks which, when
   *  clicked, allows user to go from summary table to detailed table, for a
   *  selected patrol.
   */
  const hyperlinkToSummaryRange = sheet.getRange('hyperlinkToSummaryTable');
  const hyperlinkToSummaryValues = hyperlinkToSummaryRange.getValues();
  const hyperlinkPart1 = 'https://docs.google.com/spreadsheets/d/';

  /**
   * @description "hyperlinkTargetPatrol" is a named range in spreadsheet.
   * It contains links used to go from summary table to detailed table, for
   * currently selected patrol, in tab "Heures par patrouilleur"
   */
  const hyperlinkToDetailRange = sheet.getRange('hyperlinkToDetailTable');
  // const hyperlinkTargetRow = hyperlinkTargetRange.getRow();
  const patrolNameColumnNumber = hyperlinkToDetailRange.getColumn();

  for (const row in hyperlinkToSummaryValues) {
    hyperlinkToSummaryValues[
      row
    ][0] = `${hyperlinkPart1}${ssID}/edit#gid=${sheetID}&range=${sheet
      .getRange(Number(row) + frozenRows, patrolNameColumnNumber)
      .getA1Notation()}`;
  }

  hyperlinkToSummaryRange.setValues(hyperlinkToSummaryValues);
}

// TODO This variable is way too long. Find a way to extract some functions.
// eslint-disable-next-line @typescript-eslint/no-unused-vars
const menu = {
  /**
   * Create a new spreadsheet for this member.
   * Register this member in the master spreadsheet.
   */
  createTimeSheet: function () {
    const activeSheet = gasObjects.activeSheet();
    const currentCell = activeSheet.getCurrentCell();
    const httpAddressPartOne = 'https://docs.google.com/spreadsheets/d/';
    const httpAddressPartThree = '/edit#gid=2083169682';
    const newDate = new Date();
    const options = {
      weekday: 'long',
      year: 'numeric',
      month: 'short',
      day: '2-digit',
      hour: 'numeric',
      minute: 'numeric'
    };
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    /* @ts-ignore */
    const dateTimeFormat = new Intl.DateTimeFormat('fr', options);
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

    // Retrieve current values of selected patrol (current row)
    const currentRange = activeSheet.getRange(
      currentCell.getRow(),
      Number(COL_HEADERS.versionDeployedIndex) + 2,
      1,
      Number(COL_HEADERS.documentNameIndex) -
        Number(COL_HEADERS.versionDeployedIndex) +
        1
    );
    const currValues = currentRange.getValues();

    // Retrieve current values of selected patrol (current row)
    const currentEmailCol = Number(COL_HEADERS.patrolEmailIndex) + 2;
    const currentEmailRange = activeSheet.getRange(
      currentCell.getRow(),
      currentEmailCol,
      1,
      1
    );
    const currEmailValue = currentEmailRange.getValue();

    // Create array for new content to update the master file
    //let newValues = [[""]];

    // Copy content of current Master File's row into the new array
    const newValues = currValues;

    // Get and open template file
    const templateFile = gasObjects.templateFile();
    const templateSS = SpreadsheetApp.open(templateFile);

    // Initialize values to be written in Master File's row or in patrol's new spreadsheet
    //console.log(`currValues[0][COL_HEADERS.fullNameIndex] = ${currValues[0][COL_HEADERS.fullNameIndex]}`);
    const memberNameRange = activeSheet.getRange(
      currentCell.getRow(),
      Number(COL_HEADERS.fullNameIndex) + 2,
      1,
      1
    );

    const memberName = memberNameRange.getValue();
    const docMiddleName = '- Feuille de temps 2022 -';
    // let docVersion = currValues[0][COL_HEADERS.versionDeployedIndex];
    const templateFileName = templateSS.getName();
    const templateDocVer = templateFileName.substr(
      templateFileName.length - 4,
      4
    );
    const documentTitle = `${memberName} ${docMiddleName} ${templateDocVer}`;

    // Make a copy of the template to create new patrol's file in the target folder
    const newFile = templateFile.makeCopy(
      documentTitle,
      gasObjects.targetFolder()
    );

    // Share new file with patrol and team leaders, without notification
    Drive.Permissions.insert(
      {
        emailAddress: 'bruce.porter.patrouilleur.ds@gmail.com',
        role: 'writer',
        type: 'user',
        value: 'bruce.porter.patrouilleur.ds@gmail.com'
      },
      newFile.getId(),
      {
        sendNotificationEmails: 'false'
      }
    );

    Drive.Permissions.insert(
      {
        role: 'writer',
        type: 'user',
        value: 'michel.gaudreau.patrouilleur.ds@gmail.com'
      },
      newFile.getId(),
      {
        sendNotificationEmails: 'false'
      }
    );

    Drive.Permissions.insert(
      {
        role: 'writer',
        type: 'user',
        value: 'sylvain.roy.patrouilleur.ds@gmail.com'
      },
      newFile.getId(),
      {
        sendNotificationEmails: 'false'
      }
    );

    try {
      Drive.Permissions.insert(
        {
          role: 'writer',
          type: 'user',
          value: currEmailValue
        },
        newFile.getId(),
        {
          sendNotificationEmails: 'false'
        }
      );
      // eslint-disable-next-line no-empty
    } catch (error) {}

    // Continue to feed the array with new content
    newValues[0][Number(COL_HEADERS.versionDeployedIndex) - 13] =
      templateDocVer;
    newValues[0][Number(COL_HEADERS.documentIDIndex) - 13] = newFile.getId();
    newValues[0][
      Number(COL_HEADERS.documentLinkIndex) - 13
    ] = `${httpAddressPartOne}${newFile.getId()}${httpAddressPartThree}`;
    newValues[0][
      Number(COL_HEADERS.deployedIndex) - 13
    ] = `${weekday}, ${day} ${month} ${year}  ${hour}:${minute}`;
    newValues[0][Number(COL_HEADERS.documentNameIndex) - 13] =
      newFile.getName();

    // Update the master file with new content
    currentRange.setValues(newValues);

    // Update the new patrol's file with personal information
    const newFileSS = SpreadsheetApp.open(newFile);

    // The code below logs the name of the first named range.
    /*     let namedRanges = newFileSS.getNamedRanges();
    Logger.log(namedRanges.length);
    for (let i = 0; i < namedRanges.length; i++) {
//      Logger.log(namedRanges[i].getName());
      console.log(`i = ${i}  /  namedRanges[i].getName() = ${namedRanges[i].getName()}`);
    }
 */

    // const postalAddressRange = newFileSS.getRangeByName('_patrolAddress');
    const fullNameRange = newFileSS.getRangeByName('_patrolFullName');
    const shortNameRange = newFileSS.getRangeByName('_patrolShortName');
    // let geoCodeRange = newFileSS.getRangeByName("_userAddressGeoCode");
    // let distanceRange = newFileSS.getRangeByName("_pathMilestoneDistanceAPI");

    fullNameRange.setValue(`${currValues[0][COL_HEADERS.fullNameIndex]}`);
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

  item2: function () {
    gasObjects.ui().alert('You clicked: Second item');
  }
};

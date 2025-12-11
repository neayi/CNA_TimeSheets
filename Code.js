function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Feuilles de temps')
      .addItem('Générer les feuilles de temps', 'createTimeSheets')
      .addToUi();
}

function getValue(row, columnName, headers) {
  let colIndex = headers.indexOf(columnName);
  if (colIndex == -1)
    return null;
    
  return row[colIndex];
}

function getDateValue(row, columnName, headers) {
  let colIndex = headers.indexOf(columnName);
  if (colIndex == -1)
    return null;

  let dateValue = row[colIndex];

  if (dateValue instanceof Date == false)
    throw new Error("Cell " + columnName + " is not a date");

  return dateValue;
}

function createTimeSheets()
{
  // Find the project in AirTable
  let params = this.getParams();

  let year = String(params['Année']);
  let supervisor = String(params['Nom du superviseur']);

  let dateSignature = params['Date à indiquer dans la feuille de temps'];

  if (year.length == 0 || dateSignature == '') {
    SpreadsheetApp.getUi().alert("Veuillez vérifier les paramètres dans l'onglet Accueil !");
    return;
  }

  dateSignature = this.dateToString(dateSignature);
  const options = {
        month: 'long',
        year: 'numeric'
      };
      
  const exportFolder = getOrCreateSubfolder(params['Project acronym']);

  // Get the data from the sheet "Import temps déclarés"
  let dataRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Import temps déclarés").getDataRange();
  let dataValues = dataRange.getValues();

  let headers = dataValues.shift();

  let persons = new Map();
  let firstYear = null;
  let lastYear = null;

  // First, get all the people and all the years for this project
  dataValues.forEach(row => {
    let project = this.getValue(row, "Projet", headers);

    if (project != params['Project acronym'])
      return;

    let personName = this.getValue(row, "Collaborateur", headers);
    let person = {
      times: new Map(),
      name: personName
    };

    if (persons.has(personName)) {
      person = persons.get(personName);
    }

    let declaredDate = this.getDateValue(row, "Mois", headers);

    if (firstYear === null || firstYear > declaredDate.getFullYear())
      firstYear = declaredDate.getFullYear();

    if (lastYear === null || lastYear < declaredDate.getFullYear())
      lastYear = declaredDate.getFullYear();

    let declaredTime = this.getValue(row, "Temps (jours)", headers);
    let workPackageShortname = this.getValue(row, "Work package", headers);
    workPackageShortname = workPackageShortname.replace(' - ' + params['Project acronym'], "");

    let monthObject = {
      declaredTime: 0,
      workPackages: new Set()
    };

    if (person.times.has(getKeyForDate(declaredDate))) {
      monthObject = person.times.get(getKeyForDate(declaredDate));
    }

    monthObject.declaredTime += declaredTime;
    monthObject.workPackages.add(workPackageShortname);

    person.times.set(getKeyForDate(declaredDate), monthObject);

    persons.set(personName, person);    
  });

  for (let year = firstYear; year <= lastYear; year++) {

    persons.forEach((person, personName) => {

      let values = [];
      let signatures = [];
      let supervisorSignatures = [];

      let hasTimeInYear = false;

      for (let m = 1; m <= 12; m++) {
        let dateKey = year + "-" + m;

        if (!person.times.has(dateKey)) {
          values.push([0, ""]);
        }
        else {
          hasTimeInYear = true;
          const declaredTime = person.times.get(dateKey);
          let numberOfDays = Math.round(declaredTime.declaredTime * 10) / 10; // un chiffre après la virgule
          values.push([numberOfDays, Array.from(declaredTime.workPackages).join("\n")]);
        }

        signatures.push(["Date: " + dateSignature + "\n\nSignature:", '=vlookup($C$7; Accueil!$A$17:$C$36; 3; false)']);
        supervisorSignatures.push(["Date: " + dateSignature + "\nName: " + supervisor + "\n\nSignature:", '=Image("' + '' + '")']);
      }

      if (!hasTimeInYear) {
        let ss = SpreadsheetApp.getActiveSpreadsheet();

        let existingSheet = ss.getSheetByName(person + " " + year);
        if (existingSheet)
          ss.deleteSheet(existingSheet);
        return;
      }
      
      let sheet = this.createTimeSheet(person.name, year);

      sheet.getRange(11, 2, 12, 2).setValues(values);
      sheet.getRange(11, 5, 12, 2).setValues(signatures);
      sheet.getRange(11, 7, 12, 3).setValues(supervisorSignatures);

      sheet.getRange(1, 2, 1, 3).setValues([[params['Project number'], params['Project acronym'], params['Call identifier']]]);

      sheet.getRange(5, 3).setValue(params['Project acronym']);
      sheet.getRange(6, 3).setValue(params['Participant Name']);
      sheet.getRange(7, 3).setValue(person.name);

      // column H
      sheet.getRange(5, 8).setValue(params['Project number']);
      
      // column I
      sheet.getRange(3, 9).setValue(year);

      // flush
      SpreadsheetApp.flush();
      
      // Export to PDF
      exportToPDF(sheet, exportFolder);
    });
  }
}

function getKeyForDate(aDate) {
  return aDate.getFullYear() + "-" + (aDate.getMonth()+1);
}

function createTimeSheet(person, year) 
{
  let ss = SpreadsheetApp.getActiveSpreadsheet();

  let existingSheet = ss.getSheetByName(person + " " + year);
  if (existingSheet)
    ss.deleteSheet(existingSheet);

  return ss.getSheetByName("Template").copyTo(ss).setName(person + " " + year);
}

function getParams()
{
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Accueil");

    let range = sheet.getDataRange();
    let values = range.getValues(); 
    let params = {};

    values.forEach(row => {
      params[row[0]] = row[1];
    });

    return params;
}

function dateToString(aDate)
{
  return aDate.toLocaleDateString("fr-FR", { day: '2-digit' })+ "/" +
         aDate.toLocaleDateString("fr-FR", { month: 'numeric' })+ "/" +
         aDate.toLocaleDateString("fr-FR", { year: 'numeric' });
}

function getOrCreateSubfolder(projectname) {
  const parentFolder = DriveApp.getFolderById('1MIyHmDXXSFMRaMPjUjYBJjsD-vaRoemp');

  // Vérifie si le sous-dossier existe déjà
  const subfolders = parentFolder.getFoldersByName(projectname);

  let subfolder;
  if (subfolders.hasNext()) {
    // Il existe déjà
    subfolder = subfolders.next();
    Logger.log("Sous-dossier existant : " + subfolder.getUrl());
  } else {
    // Sinon, on le crée
    subfolder = parentFolder.createFolder(projectname);
    Logger.log("Sous-dossier créé : " + subfolder.getUrl());
  }

  return subfolder;
}

function exportToPDF(sheet, folder) {
  const sheetId = sheet.getSheetId();
  const spreadsheetId = sheet.getParent().getId();

  const filename = sheet.getName() + '.pdf';
  // Delete any existing file with the same name
  const existingFiles = folder.getFilesByName(filename);
  while (existingFiles.hasNext()) {
    const file = existingFiles.next();
    file.setTrashed(true);
  }

  // Construire l’URL d’export (paramètres détaillés ci-dessous)
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?` +
    `format=pdf&` +
    `portrait=true&` +                 // orientation portrait
    `size=A4&` +                       // taille de page
    `fitw=true&` +                     // ajuster à la largeur
    `sheetnames=false&` +              // ne pas afficher le nom de la feuille
    `printtitle=false&` +              // ne pas répéter le titre
    `pagenumbers=false&` +             // pas de numéros de page
    `gridlines=false&` +               // pas de quadrillage
    `fzr=false&` +                     // ne pas figer les lignes
    `gid=${sheetId}`;                  // identifiant de l’onglet à exporter

  // Récupération du contenu PDF
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token },
  });

  // Sauvegarde dans Google Drive
  const blob = response.getBlob().setName(filename);
  const file = folder.createFile(blob);

  Logger.log('PDF exporté : ' + file.getUrl());
}

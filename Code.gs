const templateHeaders = ['TEST NUMBER', 'DATE',	'TIME',	'TESTER',	'APP NO',	'TEST MODE',	'EARTH CURRENT',	'EARTH',	'IEC',	'INS',	'LEAD CONTINUITY',	'USER',	'SITE',	'TEXT']
const lock = LockService.getDocumentLock();

function getLock(timeout) {
  try {
    lock.waitLock(timeout);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Could not obtain lock on document. Check no other scripts are running and try again.');
  };
};

function mainCard(status) {
  return CardService.newCardBuilder().addSection(CardService.newCardSection().setHeader("Create template").setCollapsible(false).addWidget(
    CardService.newTextParagraph().setText("Set up a template for your PAT results with the correct headers and conditional formatting. \n\nIf you are adding to an existing template you can skip this step.\n\n")
  ).addWidget(
    CardService.newTextButton().setText('New Template').setOnClickAction(CardService.newAction().setFunctionName('makeTemplate')).setTextButtonStyle(CardService.TextButtonStyle.FILLED).setBackgroundColor("#C6263E")
  ).addWidget(
    CardService.newTextParagraph().setText("&nbsp;")
  )).addSection(CardService.newCardSection().setHeader("Import results").setCollapsible(false).addWidget(
    CardService.newTextParagraph().setText("Create an ASCII export from your Seaward Apollo machine and paste the contents below to import to this sheet.\n\n")
  ).addWidget(
    CardService.newTextInput().setFieldName('textInput').setTitle('Seaward ASCII Output').setMultiline(true)
  ).addWidget(
    CardService.newTextParagraph().setText("&nbsp;")
  ).addWidget(
    CardService.newTextButton().setText('Add Results').setOnClickAction(CardService.newAction().setFunctionName('importData')).setTextButtonStyle(CardService.TextButtonStyle.FILLED).setBackgroundColor("#C6263E")
  ).addWidget(
    CardService.newTextParagraph().setText("&nbsp;\n"+status)
  ));
};

function onHomepage(e) {
  return mainCard('').build();
}

function makeTemplate() {
  getLock(5000);
  if (SpreadsheetApp.getActiveSheet().getLastRow() == 0) {
    SpreadsheetApp.getActiveSheet().setFrozenRows(1);
    SpreadsheetApp.getActiveSheet().appendRow(templateHeaders);
    let rules = SpreadsheetApp.getActiveSheet().getConditionalFormatRules();
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEndsWith('P').setBackground("#B7E1CD").setRanges([SpreadsheetApp.getActiveSheet().getRange("H:K")]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEndsWith('F').setBackground("#F4C7C3").setRanges([SpreadsheetApp.getActiveSheet().getRange("H:K")]).build());
    SpreadsheetApp.getActiveSheet().setConditionalFormatRules(rules);
    SpreadsheetApp.flush();

  } else {
      SpreadsheetApp.getUi().alert('This sheet already contains data.\n Create or switch to an empty sheet before creating your template');
  };
  lock.releaseLock;
};

function importData(e) {
  getLock(5000);
  const sheetHeaders = SpreadsheetApp.getActiveSheet().getRange(1, 1, 1, SpreadsheetApp.getActiveSheet().getMaxColumns()).getValues().flat();
  const source = e.formInput.textInput;
  if (source == null || source.trim == "") {
    SpreadsheetApp.getUi().alert('"ASCII Output" cannot be empty.\n Paste the data from your Seaward Apollo machine and try again.');
  } else if (sheetHeaders.some(r=> templateHeaders.includes(r)) == false) {
    SpreadsheetApp.getUi().alert('No matching headers were found.\n Create a template and try again.');
  } else {
    let entryArray = [];
    const sheetHeadersLargest = sheetHeaders.toSorted(function(a, b){return b.length - a.length});
    const sourceEntry = source.split('\n\n');
    for (thisSourceEntry = 0; thisSourceEntry < sourceEntry.length; thisSourceEntry++) {
      if (sourceEntry[thisSourceEntry].trim().length !== 0) {
        const sourceLines = sourceEntry[thisSourceEntry].split('\n');
        let entryOutput = [];
        let entryJson = {};
        for (thisSourceLine = 0; thisSourceLine < sourceLines.length; thisSourceLine++) {
          if (sourceLines[thisSourceLine].trim().length !== 0) {
            for (thisSheetHeader=0; thisSheetHeader < sheetHeaders.length; thisSheetHeader++) {
              if (sourceLines[thisSourceLine].trim().startsWith(sheetHeadersLargest[thisSheetHeader])) {
                const key = sourceLines[thisSourceLine].trim().substring(0, sheetHeadersLargest[thisSheetHeader].length);
                const value = sourceLines[thisSourceLine].trim().substring(sheetHeadersLargest[thisSheetHeader].length+1).trim();
                if (key in entryJson) {
                  entryJson[key] = value != "" ? entryJson[key]+", "+value : entryJson[key];
                } else {
                  entryJson[key] = value;
                };
                sourceLines[thisSourceLine] = "";
                break;
              };
            };
          };
        };
        for (thisSheetHeader=0; thisSheetHeader < sheetHeaders.length; thisSheetHeader++) {
          if (sheetHeaders[thisSheetHeader] != "" && sheetHeaders[thisSheetHeader] in entryJson) {
            entryOutput.push(entryJson[sheetHeaders[thisSheetHeader]]);
          } else {
            entryOutput.push('');
          };
        };
        if (entryOutput.join('') != "") {
          entryArray.push(entryOutput);
        }
      };
    }
    if (entryArray.length == 0) {
        SpreadsheetApp.getUi().alert('No valid results were found.\n Check your ASCII output and try again');
    } else {
      SpreadsheetApp.getActiveSheet().getRange(SpreadsheetApp.getActiveSheet().getLastRow()+1,1,entryArray.length,sheetHeaders.length).setValues(entryArray);
      SpreadsheetApp.flush();
      return CardService.newNavigation().updateCard(mainCard('Successfully imported '+entryArray.length+' results.').build());
    }
  }
  lock.releaseLock;
}
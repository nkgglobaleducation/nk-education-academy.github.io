const DATABASE_ID =
'1i-yJx4H3PUKhaADvxDPJIVV3Z4gmPtjdGbuejQwWpCw';



function saveCurrentPortfolio1(csvText) {

  const sheet =
    SpreadsheetApp
      .openById(DATABASE_ID)
      .getSheetByName(
        'CURRENT_PORTFOLIO_1'
      );

  sheet.clearContents();

  sheet
    .getRange(1,1,1,9)
    .setValues([[
      'Instrument',
      'Qty',
      'Avg Cost',
      'LTP',
      'Invested',
      'Current Value',
      'P&L',
      'Net Chg %',
      'Day Chg %'
    ]]);

  const data =
    Utilities.parseCsv(csvText);

  const rows =
    data.slice(1);

  if(rows.length){

    sheet
      .getRange(
        2,
        1,
        rows.length,
        rows[0].length
      )
      .setValues(rows);

  }

}



function saveCurrentPortfolio2(csvText) {

  const sheet =
    SpreadsheetApp
      .openById(DATABASE_ID)
      .getSheetByName(
        'CURRENT_PORTFOLIO_2'
      );

  sheet.clearContents();

  sheet
    .getRange(1,1,1,9)
    .setValues([[
      'Instrument',
      'Qty',
      'Avg Cost',
      'LTP',
      'Invested',
      'Current Value',
      'P&L',
      'Net Chg %',
      'Day Chg %'
    ]]);

  const data =
    Utilities.parseCsv(csvText);

  const rows =
    data.slice(1);

  if(rows.length){

    sheet
      .getRange(
        2,
        1,
        rows.length,
        rows[0].length
      )
      .setValues(rows);

  }

}



function saveCurrentScreener(
  screenerText
) {

  const sheet =
    SpreadsheetApp
      .openById(DATABASE_ID)
      .getSheetByName(
        'CURRENT_SCREENER'
      );

  sheet.clearContents();

  sheet
    .getRange(1,1)
    .setValue('Stock');

  const rows =
    screenerText
      .split(/\r?\n/)
      .filter(x => x.trim())
      .map(
        x => [
          x.trim().toUpperCase()
        ]
      );

  if(rows.length){

    sheet
      .getRange(
        2,
        1,
        rows.length,
        1
      )
      .setValues(rows);

  }

}



function saveScreenerSnapshot(
  screenerText
) {

  const sheet =
    SpreadsheetApp
      .openById(DATABASE_ID)
      .getSheetByName(
        'SCREENER_HISTORY'
      );

  if(sheet.getLastRow() === 0){

    sheet
      .getRange(1,1,1,2)
      .setValues([[
        'Date',
        'Stock'
      ]]);

  }

  const today =
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );

  const rows =
    screenerText
      .split(/\r?\n/)
      .filter(x => x.trim())
      .map(
        x => [
          today,
          x.trim().toUpperCase()
        ]
      );

  if(rows.length){

    sheet
      .getRange(
        sheet.getLastRow()+1,
        1,
        rows.length,
        2
      )
      .setValues(rows);

  }

}



function loadPortfolio1() {

  return SpreadsheetApp
    .openById(DATABASE_ID)
    .getSheetByName(
      'CURRENT_PORTFOLIO_1'
    )
    .getDataRange()
    .getValues();

}



function loadPortfolio2() {

  return SpreadsheetApp
    .openById(DATABASE_ID)
    .getSheetByName(
      'CURRENT_PORTFOLIO_2'
    )
    .getDataRange()
    .getValues();

}



function loadCurrentScreener() {

  const data =
    SpreadsheetApp
      .openById(DATABASE_ID)
      .getSheetByName(
        'CURRENT_SCREENER'
      )
      .getDataRange()
      .getValues();

  return data
    .slice(1)
    .flat()
    .filter(x => x);

}



function loadAIResearch() {

  const data =
    SpreadsheetApp
      .openById(DATABASE_ID)
      .getSheetByName(
        'AI_RESEARCH'
      )
      .getDataRange()
      .getValues();

  return data.slice(1);

}



function clearAIResearch() {

  const sheet =
    SpreadsheetApp
      .openById(DATABASE_ID)
      .getSheetByName(
        'AI_RESEARCH'
      );

  sheet.clearContents();

  sheet
    .getRange(1,1,1,4)
    .setValues([[
      'Stock',
      'Suggestion',
      'Alternate',
      'Rationale'
    ]]);

}



function testDatabase() {

  Logger.log(
    SpreadsheetApp
      .openById(DATABASE_ID)
      .getName()
  );

}

function updateLivePrices() {

  const liveSheet =
    SpreadsheetApp
      .openById(DATABASE_ID)
      .getSheetByName(
        'LIVE_PRICES'
      );

  const screener =
    loadCurrentScreener();

  liveSheet.clearContents();

  liveSheet
    .getRange(1,1,1,4)
    .setValues([[
      'Stock',
      'LTP',
      'Day Change %',
      'Last Updated'
    ]]);

  const rows =
    screener.map(
      stock => [
        stock,
        '=GOOGLEFINANCE("NSE:' +
        stock +
        '","price")',

        '=GOOGLEFINANCE("NSE:' +
        stock +
        '","changepct")',

        new Date()
      ]
    );

  if(rows.length){

    liveSheet
      .getRange(
        2,
        1,
        rows.length,
        4
      )
      .setValues(rows);

  }

}
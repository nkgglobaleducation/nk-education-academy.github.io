function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('AI Portfolio Manager');
}

function include(filename) {
  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();
}

function processPortfolio(
  portfolio1Csv,
  portfolio2Csv,
  screenerText
) {

  saveCurrentPortfolio1(
    portfolio1Csv
  );

  saveCurrentPortfolio2(
    portfolio2Csv
  );

  saveCurrentScreener(
    screenerText
  );

  saveScreenerSnapshot(
    screenerText
  );

  try {

    updateAIResearch();

  } catch (e) {

    Logger.log(
      e.toString()
    );

  }

  return buildDashboard();

}

function analyzeExistingDatabase() {

  return buildDashboard();

}

function buildDashboard() {

  const p1 =
    loadPortfolio1();

  const p2 =
    loadPortfolio2();

  const screener =
    new Set(
      loadCurrentScreener()
        .map(
          x =>
            String(x)
              .trim()
              .toUpperCase()
        )
    );

  const aiMap =
    loadAIResearchMap();

  let html = '';

  html += buildSummaryCards(
    p1,
    p2,
    screener,
    aiMap
  );

  html += `
  <h2>Portfolio Dashboard</h2>

  <table
  border="1"
  cellpadding="6"
  cellspacing="0"
  style="
  border-collapse:collapse;
  width:100%;
  font-size:13px;
  ">

  <tr>
    <th>Instrument</th>
    <th>Portfolio</th>
    <th>Qty</th>
    <th>Avg Cost</th>
    <th>LTP</th>
    <th>Invested</th>
    <th>Current Value</th>
    <th>P&L</th>
    <th>Net Chg %</th>
    <th>Day Chg %</th>
    <th>Screener Match</th>
    <th>AI Suggestion</th>
    <th>AI Alternate</th>
    <th>AI Rationale</th>
    <th>Historical Persistence</th>
  </tr>
  `;

  html += buildRows(
    p1,
    'P1',
    screener,
    aiMap
  );

  html += buildRows(
    p2,
    'P2',
    screener,
    aiMap
  );

  html += `
  </table>
  `;

  return html;

}

function buildSummaryCards(
  p1,
  p2,
  screener,
  aiMap
) {

  let invested = 0;
  let current = 0;
  let pnl = 0;

  let screenerMatches = 0;
  let aiBuys = 0;

  const allRows =
    p1.slice(1)
      .concat(
        p2.slice(1)
      );

  allRows.forEach(row => {

    invested +=
      Number(row[4]) || 0;

    current +=
      Number(row[5]) || 0;

    pnl +=
      Number(row[6]) || 0;

    const symbol =
      String(row[0] || '')
        .trim()
        .toUpperCase();

    if (
      screener.has(symbol)
    ) {
      screenerMatches++;
    }

    const ai =
      aiMap[symbol];

    if (
      ai &&
      ai.suggestion === 'BUY'
    ) {
      aiBuys++;
    }

  });

  const returnPct =
    invested
      ? (
          pnl /
          invested *
          100
        ).toFixed(2)
      : 0;

  return `

  <div style="
    display:flex;
    gap:15px;
    flex-wrap:wrap;
    margin-bottom:20px;
  ">

    <div style="padding:12px;border:1px solid #ccc;">
      <b>Total Invested</b><br>
      ₹ ${invested.toFixed(0)}
    </div>

    <div style="padding:12px;border:1px solid #ccc;">
      <b>Current Value</b><br>
      ₹ ${current.toFixed(0)}
    </div>

    <div style="padding:12px;border:1px solid #ccc;">
      <b>Total P&L</b><br>
      ₹ ${pnl.toFixed(0)}
    </div>

    <div style="padding:12px;border:1px solid #ccc;">
      <b>Return %</b><br>
      ${returnPct}%
    </div>

    <div style="padding:12px;border:1px solid #ccc;">
      <b>Screener Matches</b><br>
      ${screenerMatches}
    </div>

    <div style="padding:12px;border:1px solid #ccc;">
      <b>AI BUY Count</b><br>
      ${aiBuys}
    </div>

  </div>

  `;

}

function buildRows(
  data,
  portfolio,
  screener,
  aiMap
) {

  let html = '';

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const row =
      data[i];

    const symbol =
      String(
        row[0] || ''
      )
      .trim()
      .toUpperCase();

    const ai =
      aiMap[symbol] || {};

    const persistence =
      getHistoricalPersistence(
        symbol
      );

    const pnl =
      Number(row[6]) || 0;

    const pnlColor =
      pnl >= 0
        ? 'green'
        : 'red';

    html += `

    <tr>

      <td>${symbol}</td>

      <td>${portfolio}</td>

      <td>${row[1]}</td>
      <td>${row[2]}</td>
      <td>${row[3]}</td>

      <td>${row[4]}</td>
      <td>${row[5]}</td>

      <td style="
        color:${pnlColor};
        font-weight:bold;
      ">
      ${row[6]}
      </td>

      <td>${row[7]}</td>
      <td>${row[8]}</td>

      <td>
      ${
        screener.has(symbol)
        ? 'YES'
        : 'NO'
      }
      </td>

      <td>
      ${
        ai.suggestion || ''
      }
      </td>

      <td>
      ${
        ai.alternate || ''
      }
      </td>

      <td>
      ${
        ai.rationale || ''
      }
      </td>

      <td>
      ${persistence}
      </td>

    </tr>

    `;

  }

  return html;

}

function getHistoricalPersistence(
  symbol
) {

  const data =
    SpreadsheetApp
      .openById(
        DATABASE_ID
      )
      .getSheetByName(
        'SCREENER_HISTORY'
      )
      .getDataRange()
      .getValues();

  let count = 0;

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const stock =
      String(
        data[i][1] || ''
      )
      .trim()
      .toUpperCase();

    if (
      stock === symbol
    ) {
      count++;
    }

  }

  return count;

}
function updateAIResearch() {

  const stocks =
    loadCurrentScreener();

  if (!stocks.length) {

    Logger.log(
      'No screener stocks found'
    );

    return;
  }

  const apiKey =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        'GEMINI_API_KEY'
      );

  if (!apiKey) {

    throw new Error(
      'GEMINI_API_KEY not found'
    );

  }

  const prompt =
`
You are a professional Indian equity portfolio manager.

Return ONLY valid JSON.

No markdown.
No code fences.
No explanations.

For every stock provide:

1. suggestion
(BUY / HOLD / REDUCE)

2. alternate

3. rationale

Keep rationale under 20 words.

Stocks:

${stocks.join('\n')}

Return format:

[
 {
  "stock":"DIXON",
  "suggestion":"BUY",
  "alternate":"SYRMA",
  "rationale":"EMS leader benefiting from outsourcing"
 }
]
`;

  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key='
    + apiKey;

  const payload = {
    contents: [
      {
        parts: [
          {
            text: prompt
          }
        ]
      }
    ]
  };

  const response =
    UrlFetchApp.fetch(
      url,
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );

  const result =
    JSON.parse(
      response.getContentText()
    );

  if (
    !result.candidates ||
    !result.candidates.length
  ) {

    Logger.log(
      response.getContentText()
    );

    throw new Error(
      'Gemini returned no candidates'
    );

  }

  let text =
    result
      .candidates[0]
      .content
      .parts[0]
      .text;

  text =
    text
      .replace(/```json/g,'')
      .replace(/```/g,'')
      .trim();

  const aiData =
    JSON.parse(text);

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

  const rows =
    aiData.map(
      x => [
        x.stock || '',
        x.suggestion || '',
        x.alternate || '',
        x.rationale || ''
      ]
    );

  if (rows.length) {

    sheet
      .getRange(
        2,
        1,
        rows.length,
        4
      )
      .setValues(rows);

  }

  Logger.log(
    rows.length +
    ' AI records written'
  );

}



function loadAIResearchMap() {

  const data =
    loadAIResearch();

  const map = {};

  data.forEach(row => {

    const symbol =
      String(
        row[0] || ''
      )
      .trim()
      .toUpperCase();

    if (!symbol) return;

    map[symbol] = {

      suggestion:
        row[1] || '',

      alternate:
        row[2] || '',

      rationale:
        row[3] || ''

    };

  });

  return map;

}
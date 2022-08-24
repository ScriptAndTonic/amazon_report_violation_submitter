require('dotenv').config();
const { GoogleSpreadsheet } = require('google-spreadsheet');
const creds = require('./config/googleApiKey.json');

const run = async () => {
  const gdoc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID);
  await gdoc.useServiceAccountAuth(creds);

  await gdoc.loadInfo();
  console.log(`Loaded ${gdoc.title}`);

  const sheet = gdoc.sheetsByTitle['PENTRU ANDREI'];
  console.log(`Loaded ${sheet.title}`);

  const allRows = await sheet.getRows();
  console.log(`Loaded ${allRows.length} rows`);

  const rowsToProcess = allRows.filter((r) => r.Status === 'Process');
  console.log(`${rowsToProcess.length} rows to process`);

  for (let index = 0; index < rowsToProcess.length; index++) {
    const row = rowsToProcess[index];
    try {
      const title = row.Title;
      console.log(`---> Processing row: ${title};`);
      const date = row.Date;
      const author = row.Author;
      const url = row.URL;
      const itemID = row.Variation;
      var violation = '';

      for (let index = 1; index <= 7; index++) {
        const violationCell = row['Violation ' + index];
        if (violationCell === null || violationCell.trim() === '') {
          const lastExistingViolationIndex = index - 1;
          violation = row['Violation ' + lastExistingViolationIndex];
          console.log(`Selected Violation ${lastExistingViolationIndex}`);
          break;
        }
      }

      const body = `ASIN or ISBN of the product: ${itemID}
    Title of the review: ${title}
    Name of the reviewer: ${author}
    Date of the review as it appears on our website: ${date}
    Direct link to the review or post (click the 'Comments' link after the review, and copy or paste the URL that displays in your web browser): ${url}
    Required action: ${violation}`;
      console.log('\x1b[36m%s\x1b[0m', body);
      row.Status = 'Completed';
    } catch (error) {
      row.Status = 'Failed';
    } finally {
      await row.save();
    }
  }
};

run();

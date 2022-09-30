require('dotenv').config();
const fs = require('fs');
const cookiesFilePath = 'cookies.json';
const puppeteer = require('puppeteer-extra');
puppeteer.use(require('puppeteer-extra-plugin-repl')());
const { GoogleSpreadsheet } = require('google-spreadsheet');
const creds = require('./config/googleApiKey.json');
var readline = require('readline');
var rl = readline.createInterface({ input: process.stdin, output: process.stdout, terminal: false });
const nodemailer = require('nodemailer');
const transporter = nodemailer.createTransport({
  host: 'smtp.gmail.com',
  port: 465,
  auth: {
    user: process.env.GMAIL_USERNAME,
    pass: process.env.GMAIL_PASS,
  },
});

const run = async () => {
  const gdoc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID);
  await gdoc.useServiceAccountAuth(creds);

  await gdoc.loadInfo();
  console.log(`Loaded ${gdoc.title} gdoc`);

  const sheet = gdoc.sheetsByTitle[process.env.SPREADSHEET_SHEET_NAME];
  console.log(`Loaded ${sheet.title} sheet`);

  const allRows = await sheet.getRows();
  console.log(`Loaded ${allRows.length} rows`);

  const rowsToProcess = allRows.filter((r) => r.Status === 'Process');
  if (rowsToProcess.length > 0) {
    console.log(`${rowsToProcess.length} rows to process`);
  } else {
    console.log('Nothing to process, exiting...');
    process.exit(0);
  }

  const browserPage = await launchBrowserSession();
  console.log('Browser launched');

  for (let index = 0; index < rowsToProcess.length; index++) {
    const row = rowsToProcess[index];
    try {
      let lastExistingViolationIndex = 0;
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
          lastExistingViolationIndex = index - 1;
          violation = row['Violation ' + lastExistingViolationIndex];
          console.log(`Selected Violation ${lastExistingViolationIndex}`);
          break;
        }
      }

      const body = `ASIN or ISBN of the product: ${itemID}\nTitle of the review: ${title}\nName of the reviewer: ${author}\nDate of the review as it appears on our website: ${date}\nDirect link to the review or post (click the 'Comments' link after the review, and copy or paste the URL that displays in your web browser): ${url}\nRequired action: ${violation}`;

      const caseId = await submitAmazonViolation(browserPage, body);
      await reportViolation(browserPage, url);
      await sendViolationEmails(body);
      row[`Violation ${lastExistingViolationIndex}`] = body;
      row[`Case ID ${lastExistingViolationIndex}`] = caseId;
      row.Status = 'Completed';
    } catch (error) {
      console.error(error);
      row.Status = 'Failed';
    } finally {
      await row.save();
    }
    console.log(`Throttling for ${process.env.THROTTLE_SECONDS} seconds`);
    await new Promise((r) => setTimeout(r, process.env.THROTTLE_SECONDS * 1000));
  }
  console.log('Process finished, exiting...');
  process.exit();
};

const sendViolationEmails = async (body) => {
  await transporter.sendMail({
    from: process.env.GMAIL_USERNAME,
    to: process.env.SEND_VIOLATION_EMAIL_TO,
    subject: 'Review Violation',
    text: body,
  });
  await transporter.sendMail({
    from: process.env.GMAIL_USERNAME,
    to: process.env.SEND_VIOLATION_EMAIL_TO2,
    subject: 'Review Violation',
    text: body,
  });
};

const reportViolation = async (page, url) => {
  await page.goto(url);
  const popupPromise = new Promise((x) => page.once('popup', x));
  await page.click('.report-abuse-link');
  const popup = await popupPromise;
  await popup.click('.a-button-primary');
  await popup.close();
  console.log('Violation reported');
};

const launchBrowserSession = async () => {
  const browser = await puppeteer.launch({ headless: false });
  const page = await browser.newPage();
  browser.set;
  await loadCookies(page);
  await page.setViewport({ width: 1280, height: 1200 });
  return page;
};

const submitAmazonViolation = async (page, violationText) => {
  await page.goto(process.env.AMAZON_SELLER_REPORT_URL);
  await new Promise((r) => setTimeout(r, 2000));
  console.log('Waiting to pass login...');
  await page.waitForSelector('div.hh-title', { visible: true, timeout: 60000 });
  console.log('Passed login');
  await page.goto(process.env.AMAZON_SELLER_REPORT_IFRAME_URL);
  await new Promise((r) => setTimeout(r, 2000));
  await page.waitForSelector('[text="Short description"]');
  await repeatTab(page, 2);
  await clearInput(page, '#root > div > form > div > div.hill-primary-input-container > div > kat-textarea');
  await page.keyboard.type(violationText);

  await repeatTab(page, 8);

  await page.keyboard.press('Enter');
  console.log('Violation Report Submitted');
  await new Promise((r) => setTimeout(r, 10000));
  await page.goto(process.env.AMAZON_SELLER_CASE_LOG_URL);
  const caseRowId = await page.evaluate(() => document.querySelector('tr[id^="case_row"]').id);
  const caseId = caseRowId.replace('case_row_', '');
  console.log(`Case ID: ${caseId}`);

  return caseId;
};

const clearInput = async (page, selector) => {
  const value = await page.$eval(selector, (el) => el.value || el.innerText || '');
  await page.focus(selector);
  for (let i = 0; i < value.length; i++) {
    await page.keyboard.press('Backspace');
  }
};

const saveCookies = (cookiesObject) => {
  fs.writeFile(cookiesFilePath, JSON.stringify(cookiesObject), (err) => {
    if (err) {
      console.log('The file could not be written.', err);
    }
    console.log('Session has been successfully saved');
  });
};

const loadCookies = async (page) => {
  const previousSession = fs.existsSync(cookiesFilePath);
  if (previousSession) {
    const cookiesString = fs.readFileSync(cookiesFilePath);
    const parsedCookies = JSON.parse(cookiesString);
    if (parsedCookies.length !== 0) {
      for (let cookie of parsedCookies) {
        await page.setCookie(cookie);
      }
      console.log('Session has been loaded in the browser');
    }
  }
};

const repeatTab = async (page, count) => {
  for (let index = 0; index < count; index++) {
    await page.keyboard.press('Tab');
    await new Promise((r) => setTimeout(r, 200));
  }
};

run();

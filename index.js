const path = require('path');
require('dotenv').config({ path: path.resolve(__dirname, '.env', '.env') });
const { chromium } = require('playwright');
const fs = require('fs');
const xlsx = require('xlsx');
const { google } = require('googleapis');

// Constants
const LOGIN_URL = 'https://admin.aipartner.com:9010';
const REQUEST_PAGE_URL = 'https://admin.aipartner.com:9010/safeCare/management/request';
const MANAGEMENT_PAGE_URL = 'https://admin.aipartner.com:9010/safeCare/management';

const DOWNLOAD_DIR = path.join(__dirname, 'downloads');
const REQUEST_FILE_PATH = path.join(DOWNLOAD_DIR, 'request.xlsx');
const MANAGEMENT_FILE_PATH = path.join(DOWNLOAD_DIR, 'management.xlsx');

const SPREADSHEET_ID = '1lUxjLHDGpLjIG52dlwnKnyXYO0rGLlAZxivOCGgHVJs';
const CREDENTIALS_PATH = path.join(__dirname, 'credentials.json');

async function downloadExcelFiles() {
  console.log('Starting Playwright download process...');
  
  if (!fs.existsSync(DOWNLOAD_DIR)) {
    fs.mkdirSync(DOWNLOAD_DIR, { recursive: true });
  }

  // Clear existing downloads if they exist safely
  try { if (fs.existsSync(REQUEST_FILE_PATH)) fs.unlinkSync(REQUEST_FILE_PATH); } catch (e) { console.warn('Could not cleanly unlink old request file:', e.message); }
  try { if (fs.existsSync(MANAGEMENT_FILE_PATH)) fs.unlinkSync(MANAGEMENT_FILE_PATH); } catch (e) { console.warn('Could not cleanly unlink old management file:', e.message); }

  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ acceptDownloads: true });
  const page = await context.newPage();

  try {
    // 1. Login
    console.log(`Navigating to ${LOGIN_URL}...`);
    await page.goto(LOGIN_URL, { waitUntil: 'networkidle' });

    console.log('Filling in credentials...');
    const adminId = process.env.ADMIN_ID;
    const adminPwd = process.env.ADMIN_PWD;

    if (!adminId || !adminPwd) {
      throw new Error('ADMIN_ID or ADMIN_PWD not found in environment variables.');
    }

    await page.fill('input[name="id"]', adminId);
    await page.fill('input[name="pwd"]', adminPwd);

    console.log('Clicking login button and waiting for navigation...');
    await Promise.all([
      page.waitForNavigation({ waitUntil: 'networkidle' }).catch(() => console.log('Navigation wait resolved dynamically...')),
      page.click('button[type="submit"], .login-btn, input[type="submit"], button:has-text("로그인")')
    ]);
    
    // Explicit 3 second pause checking local URL context
    await page.waitForTimeout(3000);
    console.log('After login URL:', page.url());

    // 4. Checking session stability logic protecting Request redirects
    if (page.url().includes('login')) {
      console.log('URL still indicates login state. Waiting an additional 5 seconds...');
      await page.waitForTimeout(5000);
    }

    console.log('Login successful.');

    // Intercept Network Response (Fallback for JS-driven downloads)
    async function interceptAndDownload(targetPageUrl, routePattern, savePath) {
      console.log(`\nNavigating to target page: ${targetPageUrl}`);
      await page.goto(targetPageUrl, { waitUntil: 'networkidle' });

      // Create a promise to resolve exactly when the network intercepts the XHR matching excelDownload
      const downloadPromise = new Promise((resolve) => {
        page.route(routePattern, async (route) => {
          console.log(`Intercepted download request: ${route.request().url()}`);
          try {
            const response = await route.fetch();
            const buffer = await response.body();
            fs.writeFileSync(savePath, buffer);
            console.log(`File successfully saved to disk at: ${savePath}`);
            await route.fulfill({ response }); // Fulfill it so the UI doesn't hang
          } catch (e) {
            console.log(`Error processing intercepted route: ${e.message}`);
          } finally {
            await page.unroute(routePattern); // Unbind listener securely
            resolve();
          }
        }, { times: 1 });
      });

      console.log(`\n--- DEBUG INFO ---`);
      console.log(`Current URL: ${page.url()}`);
      
      await page.waitForLoadState('networkidle');
      await page.screenshot({ path: './debug-screenshot.png' });
      
      const buttons = await page.$$eval('button', 
        btns => btns.map(b => ({id: b.id, text: b.textContent.trim()}))
      );
      console.log('Buttons found:', JSON.stringify(buttons));
      console.log(`------------------\n`);

      console.log(`Clicking active #excelBtn...`);
      await page.click('#excelBtn');
      
      console.log(`Waiting for JS XHR network response buffer execution...`);
      await Promise.race([
        downloadPromise,
        new Promise((_, reject) => setTimeout(() => reject(new Error("Timeout waiting for intercepted excelDownload route. The button click did not trigger a network fetch.")), 60000))
      ]);
    }

    // 2. Download Request Excel
    await interceptAndDownload(REQUEST_PAGE_URL, '**/excelDownload**', REQUEST_FILE_PATH);

    // 3. Download Management Excel
    await interceptAndDownload(MANAGEMENT_PAGE_URL, '**/excelDownload**', MANAGEMENT_FILE_PATH);

  } catch (error) {
    console.error('An error occurred during Playwright automation:\n', error);
    throw error;
  } finally {
    await browser.close();
  }
  
  // Verify files exist
  console.log('Verifying downloaded files...');
  if (!fs.existsSync(REQUEST_FILE_PATH)) throw new Error(`File not found: ${REQUEST_FILE_PATH}`);
  if (!fs.existsSync(MANAGEMENT_FILE_PATH)) throw new Error(`File not found: ${MANAGEMENT_FILE_PATH}`);
  console.log('Download and verification complete.\n');
}

async function uploadToGoogleSheets() {
  console.log('Starting Google Sheets upload process...');
  
  if (!fs.existsSync(CREDENTIALS_PATH)) {
    throw new Error(`Credentials file not found at ${CREDENTIALS_PATH}`);
  }

  const auth = new google.auth.GoogleAuth({
    keyFile: CREDENTIALS_PATH,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  const sheets = google.sheets({ version: 'v4', auth });

  const processFileAndUpload = async (filePath, sheetName) => {
    console.log(`Reading Excel file: ${filePath}`);
    const workbook = xlsx.readFile(filePath);
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    
    // Convert to 2D array: array of arrays avoiding sparse rows
    const data = xlsx.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

    if (!data || data.length === 0) {
      console.log(`No data found in ${filePath}. Skipping upload.`);
      return;
    }

    console.log(`Clearing existing data in sheet: ${sheetName}`);
    try {
      await sheets.spreadsheets.values.clear({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${sheetName}'`,
      });
    } catch(err) {
      console.warn(`Could not clear sheet '${sheetName}'. It might not exist or another error occurred: ${err.message}`);
    }

    console.log(`Uploading ${data.length} rows to sheet: ${sheetName}`);
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${sheetName}'!A1`,
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: data,
      },
    });
    console.log(`Upload successful for sheet: ${sheetName}`);
  };

  try {
    await processFileAndUpload(REQUEST_FILE_PATH, '신청관리');
    await processFileAndUpload(MANAGEMENT_FILE_PATH, '관리리스트');
  } catch (error) {
    console.error('An error occurred during Google Sheets upload:\n', error);
    throw error;
  }
}

async function main() {
  try {
    await downloadExcelFiles();
    await uploadToGoogleSheets();
    console.log('Automation script completed successfully.');
  } catch (error) {
    console.error('Automation failed due to an error:\n', error);
    process.exit(1);
  }
}

main();

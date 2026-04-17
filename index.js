const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');
const { google } = require('googleapis');

const DOWNLOAD_DIR = path.join(__dirname, 'downloads');
const REQUEST_FILE_PATH = path.join(DOWNLOAD_DIR, 'request.xlsx');
const MANAGEMENT_FILE_PATH = path.join(DOWNLOAD_DIR, 'management.xlsx');

const SPREADSHEET_ID = '1lUxjLHDGpLjIG52dlwnKnyXYO0rGLlAZxivOCGgHVJs';
const CREDENTIALS_PATH = path.join(__dirname, 'credentials.json');

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
    let data = xlsx.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

    // Exclude Column B(1), D(3), E(4), F(5)
    const excludedIndices = new Set([1, 3, 4, 5]);

    if (sheetName === '관리리스트' && data.length > 0) {
      const headers = data[0];
      const additionalCols = [
        '계약일', '잔금일', '임대차만료일', '결제일시', '환불일시',
        '결제상태', '청약번호', '증권번호', '발급완료일', '버전'
      ];
      additionalCols.forEach(colName => {
        const idx = headers.indexOf(colName);
        if (idx !== -1) {
          excludedIndices.add(idx);
        }
      });
    }

    data = data.map(row => row.filter((_, index) => !excludedIndices.has(index)));

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
    } catch (err) {
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
    await uploadToGoogleSheets();
    console.log('Upload script completed successfully.');
  } catch (error) {
    console.error('Upload failed due to an error:\n', error);
    process.exit(1);
  }
}

main();

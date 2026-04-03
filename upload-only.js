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
    let data = xlsx.utils.sheet_to_json(worksheet, { 
      header: 1, 
      defval: '',
      range: 1  // Start from row 2 (0-indexed) since row 1 is empty
    });

    if (data.length > 0) {
      const headers = data[0];
      console.log(`\n[${sheetName}] Original Excel headers:`, headers);

      let columnsToKeep = [];
      
      if (sheetName === '신청관리') {
        columnsToKeep = [
          '중개업소명', '주택유형', '거래유형', '소재지', '매매/보증 금액',
          '신청일시', '진행상태', '결제구분', '결제금액', '테스트회원'
        ];
      } else if (sheetName === '관리리스트') {
        columnsToKeep = [
          '중개업소명', '주택유형', '거래유형', '소재지', '매매금액',
          '신청일시', '진행상태', '결제구분', '결제금액', '테스트회원'
        ];
      }

      const keepIndices = [];
      const keptHeaderNames = [];
      
      columnsToKeep.forEach(colName => {
        const idx = headers.findIndex(h => typeof h === 'string' && h.trim() === colName);
        if (idx !== -1) {
          keepIndices.push(idx);
          keptHeaderNames.push(headers[idx]);
        } else {
          console.warn(`[${sheetName}] WARNING: Setup expected column '${colName}' but it was not found in Excel!`);
        }
      });

      console.log(`[${sheetName}] Kept columns:`, keptHeaderNames);

      const locationKeptIdx = keptHeaderNames.indexOf('소재지');

      data = data.map((row, rowIndex) => keepIndices.map((idx, keptIdx) => {
        let val = row[idx] !== undefined ? row[idx] : '';
        
        // Truncate 소재지 column (address) to first 3 words, skipping the header row
        if (rowIndex > 0 && keptIdx === locationKeptIdx && typeof val === 'string' && val.trim() !== '') {
          const words = val.trim().split(' ');
          val = words.slice(0, 3).join(' ');
        }
        
        return val;
      }));
    }

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
    await uploadToGoogleSheets();
    console.log('Upload script completed successfully.');
  } catch (error) {
    console.error('Upload failed due to an error:\n', error);
    process.exit(1);
  }
}

main();

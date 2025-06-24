function importRealmAffiliateData() {
  const reportsFolderId = '1jh4C6VJGjzIJzitTOIqlL1ltR6YBofIr';
  const sheetId = '1IrhXeGfltFxdPCOyRTBchr-33pG6WLv8-1jypiXBI2E';
  const sheetName = 'Data My affiliates';
  const targetAffiliates = ['Adv_014', 'adv_041', 'Adv_022'];

  // Get current month/year folder name
  const today = new Date();
  const year = today.getFullYear().toString();
  const monthNumber = today.getMonth() + 1;
  const monthName = today.toLocaleString('en-US', { month: 'long' });
  const monthFolderName = `${monthNumber}-${monthName}${year}`; // e.g. 6-June2025

  // Traverse Reports > Realm > 2025 > 6-June2025
  let folder = DriveApp.getFolderById(reportsFolderId);
  for (const name of ['Realm', year, monthFolderName]) {
    const subfolders = folder.getFoldersByName(name);
    if (!subfolders.hasNext()) {
      Logger.log(`❌ Subfolder "${name}" not found inside "${folder.getName()}"`);
      return;
    }
    folder = subfolders.next();
  }

  // Diagnostic: list all files in the folder
const debugFiles = folder.getFiles();
Logger.log('📁 Realm Files found in folder:');
while (debugFiles.hasNext()) {
  const file = debugFiles.next();
  Logger.log('   ➤ ' + file.getName());
}


  // Get latest file from that month folder
  const files = folder.getFiles();
  let latestFile = null;
  let latestDate = new Date(0);
  while (files.hasNext()) {
    const file = files.next();
    if (file.getDateCreated() > latestDate) {
      latestDate = file.getDateCreated();
      latestFile = file;
    }
  }

  if (!latestFile) {
    Logger.log('❌ No file found in the current month Realm folder..');
    return;
  }

  const content = latestFile.getBlob().getDataAsString();
  const rows = Utilities.parseCsv(content);
  if (rows.length < 2) {
    Logger.log('❌ Realm CSV is empty or malformed.');
    return;
  }

  const data = rows.slice(1); // skip header
  const filtered = data.filter(row => targetAffiliates.includes(row[1]));

  //Today date
  const todayFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

  // Rearranged output
 const transformed = filtered.map(row => {
    const newRow = new Array(10).fill('');
    newRow[3] = row[2]; // C → D
    newRow[2] = row[3]; // D → C
    newRow[4] = row[5]; // F → E
    newRow[5] = row[6]; // G → F
    newRow[6] = row[7]; // H → G
    newRow[7] = row[8]; // I → H
    newRow[8] = row[9]; // J → I
    newRow[9] = row[1]; // B → J
    newRow[1] = todayFormatted;
    return newRow;
  });

  if (transformed.length === 0) {
    Logger.log('⚠️ No matching affiliates in Realm.');
    return;
  }

  // Open destination sheet
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`❌ Sheet "${sheetName}" not found.`);
    return;
  }

  // Find first empty row in column J
  const colJ = sheet.getRange('J:J').getValues();
  let firstEmptyRow = colJ.findIndex(r => !r[0]);
  if (firstEmptyRow === -1) firstEmptyRow = sheet.getLastRow() + 1;
  else firstEmptyRow += 1;

  // Paste into the sheet
  sheet.getRange(firstEmptyRow, 1, transformed.length, transformed[0].length).setValues(transformed);
  Logger.log(`✅ Realm: Inserted ${transformed.length} rows from "${latestFile.getName()}" at row ${firstEmptyRow}`);
}

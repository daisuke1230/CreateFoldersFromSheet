function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('建立資料夾')
    .addItem('執行建立', 'createFoldersFromSheet')
    .addToUi();
}

function createFoldersFromSheet() {
  const PARENT_FOLDER_ID = '要建立的資料夾位置 ID'; // 

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const projectName = sheet.getRange('A2').getValue();

  if (!projectName) {
    SpreadsheetApp.getUi().alert("請先輸入專案名稱於 A2 儲存格");
    return;
  }

  // 建立主資料夾
  let parentFolder;
  try {
    parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
  } catch (e) {
    SpreadsheetApp.getUi().alert("無法存取指定的上層資料夾 ID，請確認是否正確並具有存取權限");
    return;
  }

  const rootFolder = parentFolder.createFolder(projectName);

  // 讀取資料（含標題列、勾選列與子子資料夾列）
  const lastColumn = sheet.getLastColumn();
  const data = sheet.getRange(1, 1, 10, lastColumn).getValues(); // 第 1~10 列全欄
  const headerRow = data[0];       // 第 1 列：資料夾名稱
  const selectionRow = data[1];    // 第 2 列：是否要建立（"V"）
  const subfolderRows = data.slice(2, 10); // 第 3~10 列：子子資料夾

  // 建立子資料夾及其子項目
  let count = 0;
  for (let col = 1; col < lastColumn; col++) { // 從第 2 欄（B欄）開始
    const isSelected = selectionRow[col];
    const folderTitle = headerRow[col];

    if (isSelected === "V" && folderTitle) {
      const number = count.toString().padStart(2, '0');
      const childFolderName = `${number}_${folderTitle}`;
      const childFolder = rootFolder.createFolder(childFolderName);
      count++;

      // 建立對應欄位下的子子資料夾（第 3~10 列）
      for (let row = 0; row < subfolderRows.length; row++) {
        const subfolderName = subfolderRows[row][col];
        if (subfolderName && subfolderName.toString().trim() !== "") {
          childFolder.createFolder(subfolderName.toString().trim());
        }
      }
    }
  }

  // 回填主資料夾連結至 B13
  sheet.getRange("B13").setFormula(`=HYPERLINK("${rootFolder.getUrl()}", "${rootFolder.getName()}")`);
  SpreadsheetApp.getUi().alert("資料夾已建立！");
}

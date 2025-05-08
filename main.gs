/**
 * =======================
 *  CONFIG 區（依需求修改）
 * =======================
 */
const CONFIG = {
  PARENT_FOLDER_ID: '1NX695wfrxWGAUUBU3UYOj7-N6InVYvrW', // ← 父層資料夾 ID
  SHEET_ROWS: { HEADER: 1, SELECT: 2, SUB_START: 3, SUB_END: 10 }, // 列區段
  LINK_CELL: 'B13', // 主資料夾連結回填位置
  CARD_LINK_CELL: 'B14',                 // ⬅ Trello 卡片回填位置
  SELECT_FLAG: 'V',  // 判斷勾選用字元
};

/**
 * onOpen：為試算表加入自訂選單
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('建立資料夾')
    .addItem('執行建立', 'createFoldersFromSheet')
    .addToUi();
}

/**
 * 主入口：建立資料夾結構並回填連結
 */
function createFoldersFromSheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const projectName = sheet.getRange('A2').getValue().toString().trim();

  if (!projectName) {
    SpreadsheetApp.getUi().alert('請先在 A2 輸入專案名稱');
    return;
  }

  const parentFolder = safeGetFolderById(CONFIG.PARENT_FOLDER_ID);
  if (!parentFolder) return;

  // 建立主資料夾
  const rootFolder = parentFolder.createFolder(projectName);

  // 讀取表格資料
  const { headers, selections, subRows } = readSheetData(sheet);

  // 建立子資料夾與第二層子資料夾
  buildFolderTree(rootFolder, headers, selections, subRows);

  // 回填主資料夾連結
  sheet.getRange(CONFIG.LINK_CELL)
       .setFormula(`=HYPERLINK("${rootFolder.getUrl()}", "${rootFolder.getName()}")`);

  SpreadsheetApp.getUi().alert('資料夾已建立完成！');

 // ➊ 建 Trello 卡
  createTrelloCard_(projectName, rootFolder.getUrl());

  // 既有：回填 Drive 連結
  sheet.getRange(CONFIG.LINK_CELL)
       .setFormula(`=HYPERLINK("${rootFolder.getUrl()}", "${rootFolder.getName()}")`);
  SpreadsheetApp.getUi().alert('資料夾已建立完成並同步開卡！');

}

/* ---------- 工具函式區 ---------- */

/**
 * 由 ID 取得資料夾；失敗時顯示警告並回傳 null
 * @param {string} id - Google Drive 資料夾 ID
 * @return {Folder|null}
 */
function safeGetFolderById(id) {
  try {
    return DriveApp.getFolderById(id);
  } catch (err) {
    SpreadsheetApp.getUi().alert('無法存取父層資料夾，請確認 ID 是否正確以及帳號權限');
    return null;
  }
}

/**
 * 讀取標題列、選擇列及子子資料夾列
 * @param {Sheet} sheet
 * @return {{headers:string[], selections:string[], subRows:string[][]}}
 */
function readSheetData(sheet) {
  const lastCol = sheet.getLastColumn();
  const { HEADER, SELECT, SUB_START, SUB_END } = CONFIG.SHEET_ROWS;

  const headers    = sheet.getRange(HEADER, 2, 1, lastCol - 1).getValues()[0];
  const selections = sheet.getRange(SELECT, 2, 1, lastCol - 1).getValues()[0];
  const subRows    = sheet.getRange(SUB_START, 2, SUB_END - SUB_START + 1, lastCol - 1).getValues();

  return { headers, selections, subRows };
}

/**
 * 根據試算表資料建立第一層與第二層資料夾
 * @param {Folder} rootFolder - 主資料夾
 * @param {string[]} headers - 第一列的資料夾名稱
 * @param {string[]} selections - 第二列的勾選旗標
 * @param {string[][]} subRows - 第三列到第十列的子資料夾名稱
 */
function buildFolderTree(rootFolder, headers, selections, subRows) {
  let index = 0; // 編號計數器（00、01、02…）

  headers.forEach((title, colIdx) => {
    if (selections[colIdx] !== CONFIG.SELECT_FLAG || !title) return;

    const prefix   = index.toString().padStart(2, '0');
    const topName  = `${prefix}_${title}`;
    const topFolder = rootFolder.createFolder(topName);
    index++;

    // 建立第二層子資料夾
    subRows.forEach(row => {
      const subName = (row[colIdx] || '').toString().trim();
      if (subName) topFolder.createFolder(subName);
    });
  });
}

/*****************************************
 * Trello 開卡函式
 *****************************************/
/**
 * 在 Trello 指定 List 建卡，並回填卡片連結
 * @param {string} projectName - 卡片標題
 * @param {string} folderUrl   - Google Drive 資料夾 URL
 */
function createTrelloCard_(projectName, folderUrl) {
  const props = PropertiesService.getScriptProperties();
  const key     = props.getProperty('TRELLO_KEY');
  const token   = props.getProperty('TRELLO_TOKEN');
  const listId  = props.getProperty('TRELLO_LIST_ID');
  const baseUrl = props.getProperty('TRELLO_BASE_URL') || 'https://api.trello.com/1';

  if (!key || !token || !listId) {
    SpreadsheetApp.getUi().alert('Trello 憑證未設定完整，請檢查 Script Properties');
    return;
  }

  const payload = {
    name: projectName,
    desc: `Google Drive 資料夾：\n${folderUrl}`,
    idList: listId,
    key,
    token
  };

  const options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload,
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(`${baseUrl}/cards`, options);
    const card = JSON.parse(response.getContentText());

    // 回填 Trello 卡片 URL
    SpreadsheetApp.getActiveSheet()
      .getRange(CONFIG.CARD_LINK_CELL)
      .setFormula(`=HYPERLINK("${card.url}", "Trello Card")`);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Trello 建卡失敗，請檢查日誌');
    Logger.log(e);
  }
}

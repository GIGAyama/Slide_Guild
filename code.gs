/**
 * Slide Guild - 冒険の書
 * GIGA Standard v2 Compliant
 * Ver 2.1 (Thumbnail Robust Mode)
 * * 【先生へ】
 * 以下の CONFIG オブジェクトの中身を、以前メモしたIDに書き換えてください。
 */
const CONFIG = {
  // ▼▼▼ ここから書き換えエリア ▼▼▼
  MASTER_SS_ID: '1LU6pAxEHlYDI40pIBNa4DQWt8xvSd94BBtbL9Mfpy1c',
  STORAGE_FOLDER_ID: '1ixAqyqy7H_QwjVqgCfrjVEoLgQLyH8Zh'
  // ▲▲▲ ここまで書き換えエリア ▲▲▲
};

// ==========================================
// ⚙️ 定数定義
// ==========================================
const APP_NAME = "Slide Guild";

// ==========================================
// 🚀 初期化 & UI表示
// ==========================================
function onOpen() {
  SlidesApp.getUi()
    .createMenu('💎 スライドギルド')
    .addItem('▶ アプリを起動 (きどう)', 'showSidebar')
    .addSeparator()
    .addItem('🔧 管理者セットアップ (先生用)', 'setupAdmin')
    .addToUi();
}

function showSidebar() {
  // IDが空の場合のチェック
  if (!CONFIG.MASTER_SS_ID || !CONFIG.STORAGE_FOLDER_ID) {
    const ui = SlidesApp.getUi();
    ui.alert('⚠️ 設定エラー', 'スプレッドシートIDまたはフォルダIDが設定されていません。\nコード.gsのCONFIGを書き換えてください。', ui.ButtonSet.OK);
    return;
  }

  const html = HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(APP_NAME)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SlidesApp.getUi().showSidebar(html);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// 🔧 管理者用セットアップ機能
// ==========================================
function setupAdmin() {
  const ui = SlidesApp.getUi();
  const response = ui.alert(
    '管理者セットアップ',
    '新しいデータベースと保存フォルダを作成しますか？\n(先生が最初に1回だけ行います)',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  try {
    const folderName = `SlideGuild_Data_${new Date().getFullYear()}`;
    const folder = DriveApp.createFolder(folderName);
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const ss = SpreadsheetApp.create(`SlideGuild_DB_${new Date().getFullYear()}`);
    const file = DriveApp.getFileById(ss.getId());
    file.moveTo(folder);

    const sheet = ss.getSheets()[0];
    sheet.setName('submissions');
    // ヘッダー行
    const headers = ['timestamp', 'userId', 'questId', 'slideId', 'slideUrl', 'title', 'likes', 'deletedAt'];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#fff2cc').setFontWeight('bold');

    // コピー用コード生成
    const newConfigCode = `const CONFIG = {
  // ▼▼▼ ここから書き換えエリア ▼▼▼
  MASTER_SS_ID: '${ss.getId()}',
  STORAGE_FOLDER_ID: '${folder.getId()}'
  // ▲▲▲ ここまで書き換えエリア ▲▲▲
};`;

    const htmlOutput = HtmlService.createHtmlOutput(`
      <p style="font-family:sans-serif">セットアップ完了！以下のコードをコピーして、<b>コード.gsの先頭に上書き</b>してください。</p>
      <textarea style="width:100%; height:100px; font-family:monospace; border:2px solid #f1c40f; padding:5px;">${newConfigCode}</textarea>
      <button onclick="google.script.host.close()" style="margin-top:10px; padding:5px 15px;">閉じる</button>
    `).setWidth(400).setHeight(300);
    
    ui.showModalDialog(htmlOutput, '✅ 設定完了');

  } catch (e) {
    ui.alert(`エラー: ${e.toString()}`);
  }
}

// ==========================================
// 📤 提出機能 (Submit)
// ==========================================
function submitSlide(questId, questTitle) {
  if (!CONFIG.MASTER_SS_ID) throw new Error('管理者設定が未完了です');

  try {
    const presentation = SlidesApp.getActivePresentation();
    const slideId = presentation.getId();
    const userEmail = Session.getActiveUser().getEmail();
    
    // コピー作成
    const sourceFile = DriveApp.getFileById(slideId);
    const targetFolder = DriveApp.getFolderById(CONFIG.STORAGE_FOLDER_ID);
    
    const now = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
    const newFileName = `${questTitle}_${userEmail}_${timestamp}`;
    
    const newFile = sourceFile.makeCopy(newFileName, targetFolder);
    // 確実に公開設定にする
    newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const newSlideId = newFile.getId();
    const previewUrl = `https://docs.google.com/presentation/d/${newSlideId}/preview`;

    // DB記録
    const ss = SpreadsheetApp.openById(CONFIG.MASTER_SS_ID);
    const sheet = ss.getSheetByName('submissions');
    
    sheet.appendRow([
      timestamp,
      userEmail,
      questId,
      newSlideId,
      previewUrl,
      presentation.getName(),
      0, 
      "" 
    ]);

    return { success: true };

  } catch (e) {
    throw new Error(`提出失敗: ${e.toString()}`);
  }
}

// ==========================================
// 🖼️ ギャラリー取得
// ==========================================
function getGalleryData() {
  if (!CONFIG.MASTER_SS_ID) return [];

  try {
    const ss = SpreadsheetApp.openById(CONFIG.MASTER_SS_ID);
    const sheet = ss.getSheetByName('submissions');
    const data = sheet.getDataRange().getValues();
    data.shift(); // ヘッダー除去
    
    // 最新20件
    const recentData = data.filter(row => row[7] === "").reverse().slice(0, 20);

    return recentData.map((row) => {
      let thumbBase64 = null;
      try {
        const file = DriveApp.getFileById(row[3]); // slideId
        const blob = file.getThumbnail();
        if (blob) {
          thumbBase64 = Utilities.base64Encode(blob.getBytes());
        }
      } catch (e) {
        // 画像取得エラー時はnullのままにする（クライアント側でダミー画像を表示）
        console.warn('Thumb error for slide ' + row[3]);
      }

      return {
        rowIndex: findRowIndex(sheet, row[3]),
        timestamp: row[0],
        questId: row[2],
        slideId: row[3],
        title: row[5],
        likes: row[6],
        thumbnail: thumbBase64
      };
    });
  } catch (e) {
    console.error(e);
    return [];
  }
}

function findRowIndex(sheet, slideId) {
  const ids = sheet.getRange("D:D").getValues().flat();
  const index = ids.indexOf(slideId);
  return index !== -1 ? index + 1 : -1;
}

// ==========================================
// ❤️ いいね機能
// ==========================================
function addLike(rowIndex) {
  const ss = SpreadsheetApp.openById(CONFIG.MASTER_SS_ID);
  const sheet = ss.getSheetByName('submissions');
  const cell = sheet.getRange(rowIndex, 7); 
  const current = cell.getValue();
  cell.setValue(current + 1);
}

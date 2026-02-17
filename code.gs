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
  STORAGE_FOLDER_ID: '1ixAqyqy7H_QwjVqgCfrjVEoLgQLyH8Zh',
  TEACHER_PASSWORD: 'admin' // 先生用パスワード (変更してください)
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
    
    // DB作成
    const ss = SpreadsheetApp.create(`SlideGuild_DB_${new Date().getFullYear()}`);
    const file = DriveApp.getFileById(ss.getId());
    file.moveTo(folder);

    // 1. Submissions Sheet
    const sheet = ss.getSheets()[0];
    sheet.setName('submissions');
    // ヘッダー定義更新: Gamification columns added
    // validated: 'approvals', 'reviewedBy' (JSON), 'status' ('pending'|'approved'|'rejected')
    const headers = [
      'timestamp', 'userId', 'questId', 'slideId', 'slideUrl', 'title', 'likes', 'deletedAt', 'thumbnailFileId',
      'approvals', 'reviewedBy', 'status' 
    ];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#fff2cc').setFontWeight('bold');

    // 2. Quests Sheet
    const questSheet = ss.insertSheet('quests');
    const questHeaders = ['id', 'title', 'description', 'level', 'tags', 'demoSlideId', 'isActive'];
    questSheet.appendRow(questHeaders);
    questSheet.setFrozenRows(1);
    questSheet.getRange(1, 1, 1, questHeaders.length).setBackground('#d9ead3').setFontWeight('bold');

    // 3. Users Sheet (New for Gamification)
    const usersSheet = ss.insertSheet('users');
    // xp: 経験値, level: レベル, clearedQuests: クリア済みクエストID(JSON), lastReviewDate: 最終評価日, dailyReviewCount: 本日の評価回数
    const userHeaders = ['userId', 'xp', 'level', 'clearedQuests', 'lastReviewDate', 'dailyReviewCount'];
    usersSheet.appendRow(userHeaders);
    usersSheet.setFrozenRows(1);
    usersSheet.getRange(1, 1, 1, userHeaders.length).setBackground('#c9daf8').setFontWeight('bold');

    // コピー用コード生成
    const newConfigCode = `const CONFIG = {
  // ▼▼▼ ここから書き換えエリア ▼▼▼
  MASTER_SS_ID: '${ss.getId()}',
  STORAGE_FOLDER_ID: '${folder.getId()}'
  // ▲▲▲ ここまで書き換えエリア ▲▲▲
};`;

    const htmlOutput = HtmlService.createHtmlOutput(`
      <p style="font-family:sans-serif">【v2.5アップデート】<br>セットアップ完了！以下のコードをコピーして、<b>コード.gsの先頭に上書き</b>してください。<br><small>※古いスプレッドシートのデータは移行されません。必要手動でコピーしてください。</small></p>
      <textarea style="width:100%; height:100px; font-family:monospace; border:2px solid #f1c40f; padding:5px;">${newConfigCode}</textarea>
      <button onclick="google.script.host.close()" style="margin-top:10px; padding:5px 15px;">閉じる</button>
    `).setWidth(400).setHeight(300);
    
    ui.showModalDialog(htmlOutput, '✅ 設定完了');

  } catch (e) {
    ui.alert(`エラー: ${e.toString()}`);
  }
}

// ------------------------------------------
// 📜 Quest Data Management
// ------------------------------------------

// 管理者用: JSONテキストを受け取ってクエストを一括登録
function saveQuestData(jsonString, password) {
  if (password !== getTeacherPassword()) {
     throw new Error('パスワードが違います');
  }
  if (!CONFIG.MASTER_SS_ID) throw new Error('管理者設定が未完了です');
  
  try {
    const quests = JSON.parse(jsonString);
    if (!Array.isArray(quests)) throw new Error('JSONは配列形式である必要があります');

    const ss = SpreadsheetApp.openById(CONFIG.MASTER_SS_ID);
    let sheet = ss.getSheetByName('quests');
    if (!sheet) {
      sheet = ss.insertSheet('quests');
      sheet.appendRow(['id', 'title', 'description', 'level', 'tags', 'demoSlideId', 'isActive']);
    }

    // 既存データをクリア（ヘッダー以外）
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    }

    const rows = quests.map(q => [
      q.id || Utilities.getUuid(),
      q.title,
      q.description,
      q.level,
      Array.isArray(q.tags) ? q.tags.join(',') : q.tags,
      q.demoSlideId || '',
      true // isActive default
    ]);

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    }
    
    return { success: true, count: rows.length };

  } catch (e) {
    throw new Error(`インポート失敗: ${e.toString()}`);
  }
}

// ユーザー用: クエスト一覧取得
// ユーザー用: クエスト一覧取得
function getQuestData() {
  // 設定がない場合は空配列を返す（エラーにしない）
  if (!CONFIG.MASTER_SS_ID) {
    console.warn("MASTER_SS_ID is not set.");
    return [];
  }

  try {
    const ss = SpreadsheetApp.openById(CONFIG.MASTER_SS_ID);
    let sheet = ss.getSheetByName('quests');
    
    // シートがない場合は自動復旧
    if (!sheet) {
      console.warn("Quests sheet not found. Recovering...");
      sheet = initQuestsSheet(ss);
    }

    const data = sheet.getDataRange().getValues();
    // ヘッダーのみの場合は空
    if (data.length <= 1) return [];

    data.shift(); // ヘッダー除去
    
    // isActiveなものだけ返す
    const activeQuests = data.filter(row => {
      const isActive = row[6];
      // 厳密な判定: true (boolean) または "true" (string, case-insensitive) または 空文字(デフォルト有効とする場合)
      // ここでは「FALSE」や「false」と明記されていなければ有効とみなすロジックに変更
      if (typeof isActive === 'string') {
        return isActive.toLowerCase() !== 'false';
      }
      return isActive !== false; 
    });

    console.log(`Fetched ${activeQuests.length} active quests.`);

    return activeQuests.map(row => ({
      id: row[0],
      title: row[1],
      description: row[2],
      level: Number(row[3]),
      tags: row[4] ? row[4].toString().split(',') : [],
      demoSlideId: row[5]
    }));
  } catch(e) {
    console.warn("Quest Fetch Error", e);
    // 失敗時は空配列
    return [];
  }
}

// 🛠️ クエストシートの初期化・復旧
function initQuestsSheet(ss) {
  let sheet = ss.getSheetByName('quests');
  if (!sheet) {
    sheet = ss.insertSheet('quests');
  }
  
  // ヘッダー再設定
  // 既存データがあるかもしれないので、1行目が空の場合のみヘッダー追加
  if (sheet.getLastRow() === 0) {
    const questHeaders = ['id', 'title', 'description', 'level', 'tags', 'demoSlideId', 'isActive'];
    sheet.appendRow(questHeaders);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, questHeaders.length).setBackground('#d9ead3').setFontWeight('bold');
    
    // デフォルトデータの投入
    const defaultQuests = [
      [Utilities.getUuid(), '画像の召喚', '「挿入」メニューから好きな画像を入れよう', 1, 'image', '', true],
      [Utilities.getUuid(), '魔法の文字', 'ワードアートを使って、名前を派手に書こう', 1, 'text', '', true]
    ];
    sheet.getRange(2, 1, defaultQuests.length, defaultQuests[0].length).setValues(defaultQuests);
    console.log("Recovered quests sheet with default data.");
  }
  
  return sheet;
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
    
    // Check clearance status
    const profile = getUserProfile(); // From gamification.gs
    if (profile && profile.clearedQuests.includes(questId)) {
      throw new Error('このクエストは既にクリア済みです！');
    }

    // コピー作成
    const sourceFile = DriveApp.getFileById(slideId);
    const targetFolder = DriveApp.getFolderById(CONFIG.STORAGE_FOLDER_ID);
    
    const now = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
    
    // 1. スライド本体のコピー
    const newFileName = `${questTitle}_${userEmail}_${timestamp}`;
    const newFile = sourceFile.makeCopy(newFileName, targetFolder);
    newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const newSlideId = newFile.getId();
    
    // 埋め込み用URL
    const embedUrl = `https://docs.google.com/presentation/d/${newSlideId}/embed?start=false&loop=false&delayms=3000`;

    // 2. サムネイル画像(PNG)の生成と保存
    // スライドの1ページ目を取得
    const slides = presentation.getSlides();
    if (slides.length === 0) throw new Error('スライドが空です');
    const firstPageId = slides[0].getObjectId();
    
    // サムネイル生成用URL (export/png)
    // 注意: GASから自身のトークンでフェッチする
    const exportUrl = `https://docs.google.com/presentation/d/${slideId}/export/png?id=${slideId}&pageid=${firstPageId}`;
    const options = {
      headers: {
        Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(exportUrl, options);
    if (response.getResponseCode() !== 200) {
      throw new Error('サムネイル生成に失敗しました: ' + response.getContentText());
    }
    
    const blob = response.getBlob().setName(`${newFileName}.png`);
    const thumbFile = targetFolder.createFile(blob);
    thumbFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const thumbFileId = thumbFile.getId();

    // 3. DB記録
    const ss = SpreadsheetApp.openById(CONFIG.MASTER_SS_ID);
    const sheet = ss.getSheetByName('submissions');
    
    sheet.appendRow([
      timestamp,
      userEmail,
      questId,
      newSlideId,
      embedUrl,
      questTitle, // Use Quest Title instead of presentation.getName()
      0, 
      "",
      thumbFileId // 新規カラム
    ]);

    return { success: true };

  } catch (e) {
    console.error(e);
    throw new Error(`提出失敗: ${e.toString()}`);
  }
}

// ==========================================
// 🖼️ ギャラリー取得
// ==========================================
function getGalleryData(filterType) {
  if (!CONFIG.MASTER_SS_ID) {
    console.warn("MASTER_SS_ID not set");
    return [];
  }

  try {
    const ss = SpreadsheetApp.openById(CONFIG.MASTER_SS_ID);
    const sheet = ss.getSheetByName('submissions');
    if (!sheet) {
      console.warn("Submissions sheet not found");
      return [];
    }

    // データ全取得
    const data = sheet.getDataRange().getValues();
    
    // データがない（ヘッダーのみ含む）場合
    if (data.length <= 1) {
      console.log("No data in submissions sheet");
      return [];
    }
    
    // ヘッダー除去
    data.shift(); 
    
    // Current User Email
    const currentUserEmail = Session.getActiveUser().getEmail();

    // インメモリでオブジェクト化（行番号を保持するため）
    // 行番号は 2行目から始まるので index + 2
    const allRows = data.map((row, index) => {
      let thumbUrl = 'https://dummyimage.com/640x360/cccccc/ffffff&text=No+Image';
      const thumbId = row[8]; // I列 (thumbnailFileId)
      if (thumbId) {
        // Google Drive image direct link (New format)
        thumbUrl = `https://lh3.googleusercontent.com/d/${thumbId}`;
      }
      
      const submitterAsync = row[1]; // B列 userId
      const reviewedByJson = row[10] || "[]";
      let reviewedBy = [];
      try { reviewedBy = JSON.parse(reviewedByJson); } catch (e) {}

      return {
        rowIndex: index + 2, // シート上の行番号
        timestamp: row[0],
        userId: submitterAsync,
        questId: row[2],
        slideId: row[3],
        embedUrl: row[4],
        title: row[5],
        likes: row[6],
        deletedAt: row[7],
        thumbnailUrl: thumbUrl,
        approvals: row[9] || 0, // J列
        reviewedBy: row[10] || "[]", // K列 (JSON string)
        status: row[11] || "pending", // L列
        isMine: (submitterAsync === currentUserEmail),
        hasReviewed: reviewedBy.includes(currentUserEmail)
      };
    });

    // フィルタリング（削除されていないもの）
    let activeRows = allRows.filter(item => {
      const d = item.deletedAt;
      // 緩い判定で 0 や "0" も許可
      return !d || d == 0 || d === ""; 
    });

    // Apply Custom Filters
    if (filterType === 'mine') {
      activeRows = activeRows.filter(item => item.isMine);
    } else if (filterType === 'unreviewed') {
      // Unreviewed means: Not approved yet AND I haven't reviewed it yet AND it's not mine
      activeRows = activeRows.filter(item => 
        item.status !== 'approved' && 
        !item.hasReviewed && 
        !item.isMine
      );
    }

    console.log(`Initial Rows: ${allRows.length} -> Active: ${activeRows.length} (Filter: ${filterType})`);

    // 最新順にして20件取得 (mineの場合はもっと多くてもいいかも？一旦20)
    const recentItems = activeRows.reverse().slice(0, 20);
    
    const jsonResponse = JSON.stringify(recentItems);
    return jsonResponse;

  } catch (e) {
    console.error("getGalleryData Fatal Error:", e);
    return "[]"; 
  }
}
// findRowIndex function removed as it is no longer needed

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

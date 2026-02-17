
// ==========================================
// 🦸 ユーザー・ゲーミフィケーション管理
// ==========================================

/**
 * ユーザープロファイル取得 (XP, Level, クリア済みクエスト)
 */
function getUserProfile() {
  if (!CONFIG.MASTER_SS_ID) return null;
  const email = Session.getActiveUser().getEmail();
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.MASTER_SS_ID);
    let sheet = ss.getSheetByName('users');
    if (!sheet) {
      // シートがない場合は一時的にデフォルト値を返す（エラーにしない）
      return { userId: email, xp: 0, level: 1, clearedQuests: [], dailyReviewCount: 0 };
    }

    const data = sheet.getDataRange().getValues();
    // ヘッダーのみ
    if (data.length <= 1) {
      return { userId: email, xp: 0, level: 1, clearedQuests: [], dailyReviewCount: 0 };
    }

    // ユーザー検索
    // userId (A列), xp (B列), level (C列), clearedQuests (D列), lastReviewDate (E列), dailyReviewCount (F列)
    const userRow = data.find(r => r[0] === email);

    if (userRow) {
      // 今日の日付確認 (dailyReviewCountリセット用)
      const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
      // userRow[4] might be Date object or string
      let lastReviewDate = "";
      if (userRow[4] instanceof Date) {
        lastReviewDate = Utilities.formatDate(userRow[4], Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else if (typeof userRow[4] === 'string') {
        lastReviewDate = userRow[4].split('T')[0]; 
      }
      
      let currentDailyCount = Number(userRow[5]) || 0;
      if (lastReviewDate !== todayStr) {
        currentDailyCount = 0; // 日付が変わったらリセット
      }
      
      let cleared = [];
      try {
        cleared = JSON.parse(userRow[3] || "[]");
      } catch (e) {
         // Silently fail parse
      }

      return {
        userId: email,
        xp: Number(userRow[1]) || 0,
        level: Number(userRow[2]) || 1,
        clearedQuests: cleared,
        dailyReviewCount: currentDailyCount
      };
    } else {
      // 未登録ユーザー
      return { userId: email, xp: 0, level: 1, clearedQuests: [], dailyReviewCount: 0 };
    }

  } catch (e) {
    console.warn("getUserProfile Error", e);
    return null;
  }
}

/**
 * 相互評価（レビュー）の実施
 * @param {number} rowIndex シート上の行番号 (2〜)
 * @param {boolean} isApproved 合格の場合はtrue, やりなおしの場合はfalse
 */
function reviewSubmission(rowIndex, isApproved) {
  if (!CONFIG.MASTER_SS_ID) throw new Error("設定エラー");
  
  // ロックを取得（同時書き込み防止）
  const lock = LockService.getScriptLock();
  try {
      lock.waitLock(10000); // 10秒待機
  } catch (e) {
      throw new Error("サーバーが混み合っています。もう一度お試しください。");
  }

  try {
    const email = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.openById(CONFIG.MASTER_SS_ID);
    const subSheet = ss.getSheetByName('submissions');
    const userSheet = ss.getSheetByName('users');
    
    // 1. 投稿データの取得
    // 行番号から直接取得
    // B列(userId) も必要なので取得範囲を広げるか、個別に取る
    // userId is col 2 (B). J is 10.
    // Efficiency: get B and J-L. (2 and 10,11,12).
    // Or just get the whole row? Row is simpler.
    
    // rowIndex is 1-based.
    const rowValues = subSheet.getRange(rowIndex, 1, 1, 12).getValues()[0];
    const submitterEmail = rowValues[1]; // B
    
    // 自己評価チェック
    if (submitterEmail === email) {
        return { success: false, message: "自分の作品は評価できません！" };
    }

    let approvals = Number(rowValues[9]) || 0; // J (index 9)
    let reviewedByJson = rowValues[10] || "[]"; // K (index 10)
    let status = rowValues[11] || "pending"; // L (index 11)
    
    let reviewedBy = [];
    try {
      reviewedBy = JSON.parse(reviewedByJson);
    } catch (e) {}

    // 二重投票チェック
    if (reviewedBy.includes(email)) {
      return { success: false, message: "すでに評価済みです！" };
    }

    // 2. ユーザーの評価回数チェック & XP付与
    const rewardResult = updateReviewerStats(email, userSheet); 
    const xpGained = rewardResult.xpGained;

    // 3. 投稿データの更新
    reviewedBy.push(email);
    
    // 合格評価ならカウントアップ、不合格ならカウント維持（ただしレビュー済みにはなる）
    if (isApproved) {
      approvals++;
    }
    
    let isCleared = false;
    
    // 合格判定 (5人以上)
    if (approvals >= 5 && status !== 'approved') {
      status = 'approved';
      isCleared = true;
      
      // 投稿者にボーナスXP付与
      // 投稿者のEmailはB列(2列目)にある
      const submitterEmail = subSheet.getRange(rowIndex, 2).getValue();
      const questId = subSheet.getRange(rowIndex, 3).getValue();
      
      grantClearBonus(submitterEmail, questId, userSheet); 
    }

    // 書き込み (J, K, L -> index 10, 11, 12 in 1-based sheet coords)
    // subSheet.getRange(rowIndex, 10, 1, 3).setValues([[approvals, JSON.stringify(reviewedBy), status]]);
    subSheet.getRange(rowIndex, 10).setValue(approvals);
    subSheet.getRange(rowIndex, 11).setValue(JSON.stringify(reviewedBy));
    subSheet.getRange(rowIndex, 12).setValue(status);

    let msg = "";
    if (xpGained > 0) {
        msg = `評価完了！ +${xpGained} XP`;
    } else {
        msg = "評価完了！ (本日のXP上限です)";
    }

    return { 
      success: true, 
      xpGained: xpGained, 
      isCleared: isCleared, 
      message: msg
    };

  } catch(e) {
    console.error(e);
    throw new Error("評価処理に失敗しました: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// 内部ヘルパー: レビュアーのXP更新と回数制限チェック
function updateReviewerStats(email, sheet) {
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  // ヘッダー除外して検索
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      rowIndex = i + 1;
      break;
    }
  }

  // 今日の日付 (YYYY-MM-DD)
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  let xp = 0;
  let level = 1;
  let count = 0;
  let lastDateStr = "";
  let clearedQuestsJson = "[]";

  if (rowIndex > 1) {
    xp = Number(data[rowIndex - 1][1]) || 0;
    level = Number(data[rowIndex - 1][2]) || 1;
    clearedQuestsJson = data[rowIndex - 1][3] || "[]";
    
    const d = data[rowIndex - 1][4];
    if (d instanceof Date) {
        lastDateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (typeof d === 'string') {
        lastDateStr = d.split('T')[0];
    }
    
    count = Number(data[rowIndex - 1][5]) || 0;
  }

  // 日付チェック
  if (lastDateStr !== todayStr) {
    count = 0; // リセット
    lastDateStr = todayStr;
  }

  let xpGained = 0;
  // 1日5回までXP付与
  if (count < 5) {
    xpGained = 10;
    xp += xpGained;
    count++;
  }

  // レベル計算 (簡易: XP / 100)
  const newLevel = Math.floor(xp / 100) + 1;

  if (rowIndex > 1) {
    // 更新
    // B(2):XP, C(3):Level, E(5):LastDate, F(6):Count
    // 範囲指定して一括更新
    sheet.getRange(rowIndex, 2).setValue(xp);
    sheet.getRange(rowIndex, 3).setValue(newLevel);
    sheet.getRange(rowIndex, 5).setValue(lastDateStr);
    sheet.getRange(rowIndex, 6).setValue(count);
  } else {
    // 新規登録
    sheet.appendRow([email, xp, newLevel, clearedQuestsJson, lastDateStr, count]);
  }

  return { xpGained: xpGained };
}

// 内部ヘルパー: クリアボーナス付与
function grantClearBonus(email, questId, sheet) {
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      rowIndex = i + 1;
      break;
    }
  }

  // 今日の日付
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  if (rowIndex === -1) {
    // ユーザー未登録なら作る (万が一)
    sheet.appendRow([email, 100, 2, JSON.stringify([questId]), todayStr, 0]);
    return 100;
  }

  let xp = Number(data[rowIndex - 1][1]) || 0;
  let level = Number(data[rowIndex - 1][2]) || 1;
  let cleared = [];
  try {
    cleared = JSON.parse(data[rowIndex - 1][3] || "[]");
  } catch(e) {}

  if (!cleared.includes(questId)) {
    cleared.push(questId);
    xp += 100; // クリアボーナス
    const newLevel = Math.floor(xp / 100) + 1;
    
    // 更新
    sheet.getRange(rowIndex, 2).setValue(xp);
    sheet.getRange(rowIndex, 3).setValue(newLevel);
    sheet.getRange(rowIndex, 4).setValue(JSON.stringify(cleared));
    return 100;
  }
  
  return 0;
}

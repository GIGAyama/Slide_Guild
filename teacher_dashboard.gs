
// ==========================================
// 👩‍🏫 先生用ダッシュボード (Dashboard)
// ==========================================

// 🔐 パスワード管理
function getTeacherPassword() {
  // 1. Script Property (Dynamic)
  const prop = PropertiesService.getScriptProperties().getProperty('TEACHER_PASSWORD');
  if (prop) return prop;
  
  // 2. Config (Fallback)
  return CONFIG.TEACHER_PASSWORD || 'admin';
}

function changeTeacherPassword(currentPassword, newPassword) {
  if (currentPassword !== getTeacherPassword()) {
    throw new Error('現在のパスワードが違います');
  }
  if (!newPassword || newPassword.length < 4) {
    throw new Error('新しいパスワードは4文字以上にしてください');
  }
  
  PropertiesService.getScriptProperties().setProperty('TEACHER_PASSWORD', newPassword);
  return { success: true };
}

function getTeacherDashboardData(password) {
  try {
    const currentPassword = getTeacherPassword();
    console.log("Input Password: '" + password + "' (Type: " + typeof password + ")");
    console.log("Current Password: '" + currentPassword + "' (Type: " + typeof currentPassword + ")");

    if (String(password) !== String(currentPassword)) {
      console.warn("Password Mismatch");
      throw new Error('パスワードが違います');
    }
    if (!CONFIG.MASTER_SS_ID) {
      console.error("MASTER_SS_ID is missing");
      return null;
    }
    
    console.log("Opening SS: " + CONFIG.MASTER_SS_ID);
    const ss = SpreadsheetApp.openById(CONFIG.MASTER_SS_ID);
    
    // 1. Students Lists
    const userSheet = ss.getSheetByName('users');
    let students = [];
    if (userSheet) {
      const data = userSheet.getDataRange().getValues();
      if (data.length > 1) {
        data.shift(); // remove header
        students = data.map(row => ({
          email: row[0],
          xp: Number(row[1]) || 0,
          level: Number(row[2]) || 1,
          clearedCount: (function() { try { return JSON.parse(row[3] || "[]").length; } catch(e) { return 0; } })(),
          lastActive: row[4] ? Utilities.formatDate(new Date(row[4]), Session.getScriptTimeZone(), "yyyy/MM/dd") : "-"
        }));
      }
    } else {
        console.warn("Users sheet not found");
    }

    // 2. Recent Submissions (for monitoring)
    const subSheet = ss.getSheetByName('submissions');
    let recentSubs = [];
    if (subSheet) {
      const data = subSheet.getDataRange().getValues();
      if (data.length > 1) {
        data.shift();
        // Get last 10
        const last10 = data.reverse().slice(0, 10);
        recentSubs = last10.map(row => ({
          timestamp: row[0] ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm") : "",
          userId: row[1],
          title: row[5],
          status: row[11]
        }));
      }
    } else {
        console.warn("Submissions sheet not found");
    }

    const result = {
      students: students.sort((a, b) => b.xp - a.xp), // Sort by XP desc
      recentSubmissions: recentSubs
    };
    
    console.log("Dashboard Data Prepared. Students: " + students.length);
    // Return JSON string to avoid serialization issues
    return JSON.stringify(result);
    
  } catch (e) {
    console.error("getTeacherDashboardData Fatal Error: " + e.toString());
    // Return error as JSON too? Or throw? Throwing is better for FailureHandler.
    throw e;
  }
}

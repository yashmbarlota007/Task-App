// =====================================================================
// YASH MARKETING - TASK OPERATIONS & STAFF SCORING BACKEND (Code.gs)
// Live GitHub Sync: Automatically pulls Index.html live from GitHub
// Features:
// 1. Live Link between GAS and GitHub (Zero manual deployment needed)
// 2. 7-Period Score Engine (Daily, Weekly, Monthly, Last Month, Curr Qtr, Last Qtr, Yearly)
// 3. Multiple Reference Photos Upload to Google Drive
// 4. Dynamic Daily / Weekly / Monthly Owner Assigned Tasks & Tracking
// 5. Role Routine Tasks & Store Manager Dispatch Operations
// =====================================================================

function doGet(e) {
  var empId = (e && e.parameter && e.parameter.empId) ? e.parameter.empId : "";
  
  // LIVE LINK WITH GITHUB: Directly fetch the latest Index.html from GitHub main branch
  try {
    var githubUrl = "https://raw.githubusercontent.com/yashmbarlota007/Task-App/main/Index.html?t=" + new Date().getTime();
    var response = UrlFetchApp.fetch(githubUrl, {
      headers: { "Cache-Control": "no-cache, no-store, max-age=0" },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      var html = response.getContentText();
      html = html.replace('<?= empId ?>', empId);
      return HtmlService.createHtmlOutput(html)
        .setTitle("Yash Marketing - Task Operations & Staff Portal")
        .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  } catch (err) {
    Logger.log("Live GitHub Fetch fallback: " + err.message);
  }
  
  // Local fallback if GitHub is unreachable
  var template = HtmlService.createTemplateFromFile("Index");
  template.empId = empId;
  return template.evaluate()
    .setTitle("Yash Marketing - Task Operations & Staff Portal")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =====================================================================
// 1. AUTHENTICATION & LOGIN
// =====================================================================
function verifyLogin(id, pass) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Staff_List") || ss.getSheetByName("Users") || ss.getSheetByName("Staff");
    if (!sheet) {
      return { success: false, message: "Staff_List sheet nahi mili!" };
    }
    
    var data = sheet.getDataRange().getValues();
    var cleanId = String(id).trim().toLowerCase();
    var cleanPass = String(pass).trim();
    
    for (var i = 1; i < data.length; i++) {
      var rowEmpId = String(data[i][0]).trim().toLowerCase();
      var rowPass = String(data[i][1]).trim();
      var rowName = String(data[i][2]).trim();
      var rowRole = String(data[i][3]).trim();
      var rowActive = data[i][4];
      
      // SSO Bypass or Password match
      if (rowEmpId === cleanId) {
        if (cleanPass === "SSO_BYPASS" || rowPass === cleanPass) {
          if (rowActive !== false && String(rowActive).toLowerCase() !== "false") {
            var metrics = getStaffScoreMetrics(rowName);
            return {
              success: true,
              name: rowName,
              role: rowRole,
              totalPoints: metrics.daily || 0,
              scores: metrics
            };
          } else {
            return { success: false, message: "Aapka account inactive hai. Kripya admin se sampark karein." };
          }
        } else {
          return { success: false, message: "Galat Password! Kripya dobara koshish karein." };
        }
      }
    }
    return { success: false, message: "Employee ID nahi mili!" };
  } catch (err) {
    return { success: false, message: "Login Error: " + err.message };
  }
}

// =====================================================================
// 2. ADMIN: STAFF LIST FOR DROPDOWNS
// =====================================================================
function getStaffListForAdmin() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Staff_List") || ss.getSheetByName("Users") || ss.getSheetByName("Staff");
    if (!sheet) return ["All Staff"];
    
    var data = sheet.getDataRange().getValues();
    var list = ["All Staff"];
    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][2]).trim();
      var role = String(data[i][3] || "").trim();
      var active = data[i][4];
      if (name && active !== false && String(active).toLowerCase() !== "false") {
        if (!role.toLowerCase().includes("owner") && !role.toLowerCase().includes("admin")) {
          list.push("👤 " + name);
        }
      }
    }
    return list;
  } catch (err) {
    return ["All Staff"];
  }
}

// =====================================================================
// 3. ADMIN: ASSIGN TASK (MULTIPLE PHOTOS & FREQUENCY SUPPORT)
// =====================================================================
function assignAdhocTask(staff, typeVal, desc, deadline, photosData, fileType, fileName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Adhoc_Tasks");
    if (!sheet) {
      sheet = ss.insertSheet("Adhoc_Tasks");
      sheet.appendRow(["Assign Date", "Assigned To", "Frequency", "Task Description", "Deadline", "Reference Photo", "Status"]);
    }
    
    // Normalize frequency name: "Daily Task" -> "Daily", "Weekly Task" -> "Weekly", etc.
    var cleanFreq = String(typeVal || "Extra (One Time)").replace(/\s*Task/i, '').trim();
    if (!cleanFreq) cleanFreq = "Extra (One Time)";
    
    var cleanStaff = String(staff).replace("👤 ", "").trim();
    
    // Folder in Drive for task reference photos
    var photoLinks = [];
    var folderName = "Yash Marketing - Reference Photos";
    var folders = DriveApp.getFoldersByName(folderName);
    var targetFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    
    // Check if multiple photos received as Array
    if (Array.isArray(photosData) && photosData.length > 0) {
      for (var i = 0; i < photosData.length; i++) {
        var p = photosData[i];
        if (p && p.base64Data) {
          var pType = p.fileType || "image/jpeg";
          var pName = p.fileName || ("photo_" + (i + 1) + ".jpg");
          var blob = Utilities.newBlob(Utilities.base64Decode(p.base64Data), pType, pName);
          var file = targetFolder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          photoLinks.push(file.getUrl());
        }
      }
    } else if (photosData && typeof photosData === "string" && photosData.length > 50) {
      // Single photo backwards compatibility
      var sType = fileType || "image/jpeg";
      var sName = fileName || "ref_photo.jpg";
      var blob = Utilities.newBlob(Utilities.base64Decode(photosData), sType, sName);
      var file = targetFolder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      photoLinks.push(file.getUrl());
    }
    
    var refPhotoCell = photoLinks.length > 0 ? photoLinks.join(" , ") : "No Photo";
    var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+5:30", "yyyy-MM-dd HH:mm");
    
    sheet.appendRow([todayStr, cleanStaff, cleanFreq, desc, deadline, refPhotoCell, "Pending"]);
    return "Task successfully assign ho gaya!";
  } catch (err) {
    return "Error: " + err.message;
  }
}

// =====================================================================
// 4. STAFF: GET ASSIGNED TASKS (NORMALIZED FOR DAILY/WEEKLY/MONTHLY)
// =====================================================================
function getAdhocTasksForStaff(staffName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Adhoc_Tasks");
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var list = [];
    var cleanStaff = String(staffName || "").trim().toLowerCase();
    
    for (var i = 1; i < data.length; i++) {
      var assignedTo = String(data[i][1]).trim();
      var assignedToClean = assignedTo.toLowerCase();
      var status = String(data[i][6]).trim();
      
      if (assignedToClean === cleanStaff || assignedToClean === "all staff") {
        var rawFreq = String(data[i][2] || "Extra (One Time)").replace(/\s*Task/i, '').trim();
        list.push({
          index: i + 1,
          assignDate: data[i][0],
          assignedTo: assignedTo,
          type: rawFreq, // "Daily", "Weekly", "Monthly", or "Extra (One Time)"
          task: data[i][3],
          deadline: data[i][4],
          refPhoto: data[i][5],
          status: status
        });
      }
    }
    return list;
  } catch (err) {
    return [];
  }
}

// =====================================================================
// 5. STAFF: GET ROUTINE ROLE TASKS
// =====================================================================
function getTasksForRole(roleName, staffName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Role_Tasks") || ss.getSheetByName("Tasks");
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var roleClean = String(roleName || "").trim().toLowerCase();
    var tasks = [];
    
    // Check submissions done today by this staff
    var subSheet = ss.getSheetByName("Task_Submissions");
    var todayDoneMap = {};
    if (subSheet) {
      var subData = subSheet.getDataRange().getValues();
      var now = new Date();
      var startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
      
      for (var s = 1; s < subData.length; s++) {
        var sStaff = String(subData[s][1]).trim().toLowerCase();
        if (sStaff === String(staffName || "").trim().toLowerCase()) {
          var sDate = new Date(subData[s][0]).getTime();
          if (sDate >= startOfToday) {
            var sKey = String(subData[s][4]).trim() + "___" + String(subData[s][5] || "").trim();
            todayDoneMap[sKey] = String(subData[s][7]);
          }
        }
      }
    }
    
    for (var i = 1; i < data.length; i++) {
      var targetRole = String(data[i][0]).trim().toLowerCase();
      if (targetRole === roleClean || targetRole === "all") {
        var type = String(data[i][1]).trim();
        var mainTask = String(data[i][2]).trim();
        var subTask = String(data[i][3] || "").trim();
        var deadline = String(data[i][4] || "N/A").trim();
        var sop = String(data[i][5] || "").trim();
        var sopVideo = String(data[i][6] || "").trim();
        
        var key = mainTask + "___" + subTask;
        var status = todayDoneMap[key] || "Pending";
        
        tasks.push({
          role: data[i][0],
          type: type, // "Daily", "Weekly", "Monthly"
          task: mainTask,
          subTask: subTask,
          deadline: deadline,
          sop: sop,
          sopVideo: sopVideo,
          status: status
        });
      }
    }
    return tasks;
  } catch (err) {
    return [];
  }
}

// =====================================================================
// 6. STAFF & OWNER: 7-PERIOD SCORE ENGINE
// Daily, Weekly, Monthly, Last Month, Curr Qtr, Last Qtr, Yearly
// =====================================================================
function getStaffScoreMetrics(staffName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Task_Submissions");
    var emptyScores = { daily: 0, weekly: 0, monthly: 0, lastMonth: 0, currentQuarter: 0, lastQuarter: 0, yearly: 0 };
    if (!sheet) return emptyScores;
    
    var data = sheet.getDataRange().getValues();
    var now = new Date();
    
    // 1. Daily (Today from 00:00:00)
    var startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0, 0);
    
    // 2. Weekly (Monday of current week)
    var dayOfWeek = now.getDay() === 0 ? 7 : now.getDay();
    var startOfWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() - dayOfWeek + 1, 0, 0, 0, 0);
    
    // 3. Monthly (1st of current calendar month)
    var startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0);
    
    // 4. Last Month (1st to last day of previous month)
    var startOfLastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1, 0, 0, 0, 0);
    var endOfLastMonth = new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59, 999);
    
    // 5. Current Quarter (Q1: Jan-Mar, Q2: Apr-Jun, Q3: Jul-Sep, Q4: Oct-Dec)
    var currentQtrNum = Math.floor(now.getMonth() / 3);
    var startOfCurrentQtr = new Date(now.getFullYear(), currentQtrNum * 3, 1, 0, 0, 0, 0);
    
    // 6. Last Quarter
    var lastQtrYear = currentQtrNum === 0 ? now.getFullYear() - 1 : now.getFullYear();
    var lastQtrMonthIndex = currentQtrNum === 0 ? 9 : (currentQtrNum - 1) * 3;
    var startOfLastQtr = new Date(lastQtrYear, lastQtrMonthIndex, 1, 0, 0, 0, 0);
    var endOfLastQtr = new Date(now.getFullYear(), currentQtrNum * 3, 0, 23, 59, 59, 999);
    
    // 7. Yearly (1st Jan of current year)
    var startOfYear = new Date(now.getFullYear(), 0, 1, 0, 0, 0, 0);
    
    var scores = { daily: 0, weekly: 0, monthly: 0, lastMonth: 0, currentQuarter: 0, lastQuarter: 0, yearly: 0 };
    var cleanStaff = String(staffName || "").trim().toLowerCase();
    
    for (var i = 1; i < data.length; i++) {
      var rowName = String(data[i][1] || "").trim().toLowerCase();
      if (cleanStaff && rowName !== cleanStaff) continue;
      
      var rawDate = data[i][0];
      if (!rawDate) continue;
      var subDate = new Date(rawDate);
      if (isNaN(subDate.getTime())) continue;
      
      var status = String(data[i][7] || "");
      if (status.indexOf("Skipped") !== -1) continue; // Skipped tasks = 0 pts
      
      var pts = parseInt(data[i][10]) || 10; // Default 10 points per completed task
      
      if (subDate >= startOfToday) scores.daily += pts;
      if (subDate >= startOfWeek) scores.weekly += pts;
      if (subDate >= startOfMonth) scores.monthly += pts;
      if (subDate >= startOfLastMonth && subDate <= endOfLastMonth) scores.lastMonth += pts;
      if (subDate >= startOfCurrentQtr) scores.currentQuarter += pts;
      if (subDate >= startOfLastQtr && subDate <= endOfLastQtr) scores.lastQuarter += pts;
      if (subDate >= startOfYear) scores.yearly += pts;
    }
    
    return scores;
  } catch (err) {
    return { daily: 0, weekly: 0, monthly: 0, lastMonth: 0, currentQuarter: 0, lastQuarter: 0, yearly: 0 };
  }
}

// =====================================================================
// 7. OWNER: OPTIMIZED LEADERBOARD FOR ALL STAFF (SINGLE PASS)
// =====================================================================
function getStaffLeaderboard(period) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var staffSheet = ss.getSheetByName("Staff_List") || ss.getSheetByName("Users") || ss.getSheetByName("Staff");
    var staffList = [];
    var pKey = period || "daily";
    if (!staffSheet) return [];
    
    var sData = staffSheet.getDataRange().getValues();
    var staffInfoMap = {};
    for (var s = 1; s < sData.length; s++) {
      var sName = String(sData[s][2] || "").trim();
      var sRole = String(sData[s][3] || "Staff").trim();
      var sActive = sData[s][4];
      if (sName && sActive !== false && String(sActive).toLowerCase() !== "false") {
        if (!sRole.toLowerCase().includes("owner") && !sRole.toLowerCase().includes("admin")) {
          var cleanKey = sName.toLowerCase();
          staffInfoMap[cleanKey] = {
            name: sName,
            role: sRole,
            metrics: { daily: 0, weekly: 0, monthly: 0, lastMonth: 0, currentQuarter: 0, lastQuarter: 0, yearly: 0 }
          };
        }
      }
    }
    
    var subSheet = ss.getSheetByName("Task_Submissions");
    if (subSheet) {
      var data = subSheet.getDataRange().getValues();
      var now = new Date();
      var startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0, 0);
      var dayOfWeek = now.getDay() === 0 ? 7 : now.getDay();
      var startOfWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() - dayOfWeek + 1, 0, 0, 0, 0);
      var startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0);
      var startOfLastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1, 0, 0, 0, 0);
      var endOfLastMonth = new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59, 999);
      var currentQtrNum = Math.floor(now.getMonth() / 3);
      var startOfCurrentQtr = new Date(now.getFullYear(), currentQtrNum * 3, 1, 0, 0, 0, 0);
      var lastQtrYear = currentQtrNum === 0 ? now.getFullYear() - 1 : now.getFullYear();
      var lastQtrMonthIndex = currentQtrNum === 0 ? 9 : (currentQtrNum - 1) * 3;
      var startOfLastQtr = new Date(lastQtrYear, lastQtrMonthIndex, 1, 0, 0, 0, 0);
      var endOfLastQtr = new Date(now.getFullYear(), currentQtrNum * 3, 0, 23, 59, 59, 999);
      var startOfYear = new Date(now.getFullYear(), 0, 1, 0, 0, 0, 0);
      
      for (var i = 1; i < data.length; i++) {
        var rowName = String(data[i][1] || "").trim().toLowerCase();
        if (!staffInfoMap[rowName]) continue;
        
        var rawDate = data[i][0];
        if (!rawDate) continue;
        var subDate = new Date(rawDate);
        if (isNaN(subDate.getTime())) continue;
        
        var status = String(data[i][7] || "");
        if (status.indexOf("Skipped") !== -1) continue;
        
        var pts = parseInt(data[i][10]) || 10;
        var m = staffInfoMap[rowName].metrics;
        if (subDate >= startOfToday) m.daily += pts;
        if (subDate >= startOfWeek) m.weekly += pts;
        if (subDate >= startOfMonth) m.monthly += pts;
        if (subDate >= startOfLastMonth && subDate <= endOfLastMonth) m.lastMonth += pts;
        if (subDate >= startOfCurrentQtr) m.currentQuarter += pts;
        if (subDate >= startOfLastQtr && subDate <= endOfLastQtr) m.lastQuarter += pts;
        if (subDate >= startOfYear) m.yearly += pts;
      }
    }
    
    for (var k in staffInfoMap) {
      var item = staffInfoMap[k];
      staffList.push({
        name: item.name,
        role: item.role,
        score: item.metrics[pKey] || 0,
        metrics: item.metrics
      });
    }
    
    staffList.sort(function(a, b) { return b.score - a.score; });
    return staffList;
  } catch (err) {
    return [];
  }
}

// =====================================================================
// 8. TASK SUBMISSION WITH PROOF & POINTS
// =====================================================================
function submitTaskWithProof(name, role, type, task, subTask, deadline, base64Data, mimeType, fileName, isAdhoc, adhocIndex, actionType) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var subSheet = ss.getSheetByName("Task_Submissions");
    if (!subSheet) {
      subSheet = ss.insertSheet("Task_Submissions");
      subSheet.appendRow(["Timestamp", "Staff Name", "Role", "Type", "Task", "SubTask", "Deadline", "Status", "Proof Link", "Action", "Points Earned"]);
    }
    
    var proofUrl = "No Proof";
    if (base64Data && mimeType && fileName) {
      var folderName = "Yash Marketing - Task Proofs";
      var folders = DriveApp.getFoldersByName(folderName);
      var targetFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
      var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
      var file = targetFolder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      proofUrl = file.getUrl();
    }
    
    var isSkip = actionType === "Skip";
    var pointsEarned = isSkip ? 0 : 10;
    var finalStatus = isSkip ? "Skipped" : (proofUrl !== "No Proof" ? "Uploaded (Done)" : "Completed");
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+5:30", "yyyy-MM-dd HH:mm:ss");
    
    subSheet.appendRow([nowStr, name, role, type, task, subTask || "", deadline || "", finalStatus, proofUrl, actionType || "Done", pointsEarned]);
    
    // If adhoc task, update status in Adhoc_Tasks sheet
    if (isAdhoc && adhocIndex) {
      var adSheet = ss.getSheetByName("Adhoc_Tasks");
      if (adSheet) {
        var row = parseInt(adhocIndex);
        if (row >= 2 && row <= adSheet.getLastRow()) {
          adSheet.getRange(row, 7).setValue(finalStatus);
        }
      }
    }
    
    return { status: finalStatus, pointsEarned: pointsEarned, proofUrl: proofUrl };
  } catch (err) {
    return { status: "Error: " + err.message, pointsEarned: 0 };
  }
}

// =====================================================================
// 9. ADMIN: STAFF PROGRESS REPORT
// =====================================================================
function getStaffProgressReport(cleanName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var staffSheet = ss.getSheetByName("Staff_List") || ss.getSheetByName("Users") || ss.getSheetByName("Staff");
    var role = "Staff";
    if (staffSheet) {
      var sData = staffSheet.getDataRange().getValues();
      for (var s = 1; s < sData.length; s++) {
        if (String(sData[s][2]).trim().toLowerCase() === String(cleanName).trim().toLowerCase()) {
          role = String(sData[s][3] || "Staff").trim();
          break;
        }
      }
    }
    
    var routine = getTasksForRole(role, cleanName);
    var adhoc = getAdhocTasksForStaff(cleanName);
    
    return { role: role, routine: routine, adhoc: adhoc };
  } catch (err) {
    return { role: "Staff", routine: [], adhoc: [] };
  }
}

// =====================================================================
// 10. ADMIN: ALL ADHOC HISTORY
// =====================================================================
function getAllAdhocHistory() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Adhoc_Tasks");
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var list = [];
    for (var i = data.length - 1; i >= 1; i--) {
      list.push({
        date: data[i][0],
        assignedTo: data[i][1],
        type: data[i][2],
        task: data[i][3],
        deadline: data[i][4],
        refPhoto: data[i][5],
        status: data[i][6]
      });
    }
    return list;
  } catch (err) {
    return [];
  }
}

// =====================================================================
// 11. STORE MANAGER: LIVE ORDERS DISPATCH
// =====================================================================
function getLiveOrdersForManager() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Orders");
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var orders = [];
    for (var i = 1; i < data.length; i++) {
      var status = String(data[i][6] || "");
      if (!status.toLowerCase().includes("delivered") && !status.toLowerCase().includes("completed")) {
        orders.push({
          orderId: data[i][0],
          shopName: data[i][1],
          orderTime: data[i][3],
          items: data[i][4],
          status: status || "Pending Dispatch",
          rowIndex: i + 1
        });
      }
    }
    return orders;
  } catch (err) {
    return [];
  }
}
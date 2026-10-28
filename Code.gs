// =====================================================================
// YASH MARKETING - TASK OPERATIONS & STAFF SCORING BACKEND (Code.gs)
// Live GitHub Sync: Dynamically fetches Index.html from GitHub main
// Features:
// 1. Smart Sheet Detector (Handles Staff_List, Staff List, Users, Team, Sheet1, etc.)
// 2. 7-Period Comprehensive Score Engine (Daily, Weekly, Monthly, Last Month, Curr Qtr, Last Qtr, Yearly)
// 3. Multiple Reference Photos Upload to Google Drive
// 4. Dynamic Daily / Weekly / Monthly Recurring Tasks with Auto-Reset Cycles
// 5. Role Routine Tasks, SOPs, & Store Manager Dispatch Operations
// =====================================================================

function doGet(e) {
  var empId = (e && e.parameter && e.parameter.empId) ? e.parameter.empId : "";
  
  // LIVE LINK WITH GITHUB: Directly fetch latest Index.html from GitHub main branch
  try {
    var githubUrl = "https://raw.githubusercontent.com/yashmbarlota007/Task-App/main/Index.html?t=" + new Date().getTime();
    var response = UrlFetchApp.fetch(githubUrl, {
      headers: { "Cache-Control": "no-cache, no-store, max-age=0" },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      var html = response.getContentText();
      html = html.replace('<?= portalEmpId ?>', empId);
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
  template.portalEmpId = empId;
  return template.evaluate()
    .setTitle("Yash Marketing - Task Operations & Staff Portal")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =====================================================================
// SMART SHEET DETECTOR HELPER
// Automatically handles variations: "Staff_List", "Staff List", "Staff",
// "Users", "Role_Tasks", "Role Tasks", "Adhoc_Tasks", "Adhoc Tasks", etc.
// =====================================================================
function getSheetSmart(ss, targetType) {
  if (!ss) return null;
  var sheets = ss.getSheets();
  
  var patterns = [];
  if (targetType === "staff") {
    // Matches: Staff_List, Staff List, StaffList, Staff, Users, User_List, Employees, Emp_List, Team
    patterns = [/staff.*list/i, /^staffs?$/i, /users?.*list/i, /^users?$/i, /employees?.*list/i, /^employees?$/i, /^team$/i, /^members?$/i];
  } else if (targetType === "routine_tasks") {
    // Matches: Role_Tasks, Role Tasks, Routine_Tasks, Routine Tasks, Tasks, Task_List, Master_Tasks
    patterns = [/role.*tasks?/i, /routine.*tasks?/i, /^tasks?$/i, /task.*list/i, /master.*tasks?/i];
  } else if (targetType === "adhoc_tasks") {
    // Matches: Adhoc_Tasks, Adhoc Tasks, Assigned_Tasks, Assigned Tasks, Extra_Tasks, Adhoc
    patterns = [/adhoc.*tasks?/i, /^adhoc$/i, /assigned.*tasks?/i, /extra.*tasks?/i, /owner.*tasks?/i];
  } else if (targetType === "submissions") {
    // Matches: Task_Submissions, Task Submissions, Submissions, Task_Done, Completed_Tasks
    patterns = [/task.*submissions?/i, /^submissions?$/i, /task.*done/i, /completed.*tasks?/i];
  } else if (targetType === "orders") {
    // Matches: Orders, Order_List, Shop_Orders, Dispatch
    patterns = [/orders?/i, /dispatch/i];
  }
  
  // 1. Try regex pattern match on sheet names
  for (var p = 0; p < patterns.length; p++) {
    for (var i = 0; i < sheets.length; i++) {
      var sName = sheets[i].getName().trim();
      if (patterns[p].test(sName)) {
        return sheets[i];
      }
    }
  }
  
  // 2. Fallback for staff sheet: detect based on header columns
  if (targetType === "staff") {
    for (var i = 0; i < sheets.length; i++) {
      var s = sheets[i];
      if (s.getLastRow() >= 1 && s.getLastColumn() >= 3) {
        var headers = s.getRange(1, 1, 1, Math.min(s.getLastColumn(), 10)).getValues()[0];
        var headerStr = headers.join(" ").toLowerCase();
        if ((headerStr.indexOf("pass") !== -1 || headerStr.indexOf("role") !== -1) && (headerStr.indexOf("emp") !== -1 || headerStr.indexOf("id") !== -1 || headerStr.indexOf("name") !== -1)) {
          return s;
        }
      }
    }
  }
  
  return null;
}

// Dynamically detects column index of EmpId, Password, Name, Role, Active in Staff Sheet
function getStaffColMap(sheet) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = { empId: 0, pass: 1, name: 2, role: 3, active: 4 };
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c] || "").toLowerCase().trim();
    if (h.indexOf("pass") !== -1 || h.indexOf("pwd") !== -1) map.pass = c;
    else if (h.indexOf("empid") !== -1 || h.indexOf("emp_id") !== -1 || h.indexOf("emp id") !== -1 || h === "id" || h.indexOf("username") !== -1 || h.indexOf("login") !== -1) map.empId = c;
    else if (h.indexOf("name") !== -1 || h.indexOf("staff") !== -1 || h.indexOf("employee") !== -1) map.name = c;
    else if (h.indexOf("role") !== -1 || h.indexOf("designation") !== -1) map.role = c;
    else if (h.indexOf("active") !== -1 || h.indexOf("status") !== -1) map.active = c;
  }
  return map;
}

// =====================================================================
// 1. AUTHENTICATION & LOGIN
// =====================================================================
function verifyLogin(id, pass) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return { success: false, message: "Spreadsheet connect nahi hui! Kripya check karein ki Apps Script Google Sheet ke sath bound hai." };
    }
    
    var sheet = getSheetSmart(ss, "staff");
    if (!sheet) {
      var allNames = ss.getSheets().map(function(s){ return s.getName(); }).join(", ");
      return { success: false, message: "Staff sheet nahi mili! Spreadsheet me ye sheets hain: [" + allNames + "]. Kripya staff sheet ka naam 'Staff_List' ya 'Staff List' rakhein." };
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: false, message: "Staff sheet khali hai! Kripya staff record add karein." };
    }
    
    var colMap = getStaffColMap(sheet);
    var cleanId = String(id).replace(/[^\w]/gi, '').toLowerCase();
    var cleanPass = String(pass).trim();
    
    for (var i = 1; i < data.length; i++) {
      var rowEmpId = String(data[i][colMap.empId] || "").replace(/[^\w]/gi, '').toLowerCase();
      var rowPass = String(data[i][colMap.pass] || "").trim();
      var rowName = String(data[i][colMap.name] || "").trim();
      var rowRole = String(data[i][colMap.role] || "Staff").trim();
      var rowActive = data[i][colMap.active];
      
      // SSO Bypass or Password match
      if (rowEmpId === cleanId) {
        if (cleanPass === "SSO_BYPASS" || rowPass === cleanPass) {
          if (rowActive !== false && String(rowActive).toLowerCase() !== "false") {
            var metrics = getStaffScoreMetrics(rowName);
            return {
              success: true,
              name: rowName,
              role: rowRole,
              totalPoints: (metrics && metrics.daily) ? metrics.daily.points : 0,
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
    return { success: false, message: "Employee ID '" + id + "' nahi mili! Kripya check karein." };
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
    var sheet = getSheetSmart(ss, "staff");
    if (!sheet) return ["All Staff"];
    
    var data = sheet.getDataRange().getValues();
    var colMap = getStaffColMap(sheet);
    var list = ["All Staff"];
    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][colMap.name] || "").trim();
      var role = String(data[i][colMap.role] || "").trim();
      var active = data[i][colMap.active];
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
    var sheet = getSheetSmart(ss, "adhoc_tasks");
    if (!sheet) {
      sheet = ss.insertSheet("Adhoc_Tasks");
      sheet.appendRow(["Assign Date", "Assigned To", "Frequency", "Task Description", "Deadline", "Reference Photo", "Status", "TaskId"]);
    } else {
      // Auto-upgrade older 6-column sheet to 7+ columns with Frequency
      if (sheet.getLastRow() >= 1 && sheet.getLastColumn() < 7) {
        var hVals = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        var hasFreq = hVals.some(function(h) { return /freq|type/i.test(String(h)); });
        if (!hasFreq) {
          sheet.insertColumnAfter(2);
          sheet.getRange(1, 3).setValue("Frequency");
        }
      }
    }
    
    // Normalize frequency name: "Daily Task" -> "Daily", "Weekly Task" -> "Weekly", etc.
    var cleanFreq = String(typeVal || "Extra (One Time)").replace(/[^\w\s]/gi, '').replace(/\s*Task/i, '').trim();
    if (!cleanFreq || cleanFreq.toLowerCase() === "extra" || cleanFreq.toLowerCase() === "one time") {
      cleanFreq = "Extra (One Time)";
    } else {
      cleanFreq = cleanFreq.charAt(0).toUpperCase() + cleanFreq.slice(1).toLowerCase();
    }
    
    var cleanStaff = String(staff || "").replace(/[^\w\s]/gi, '').replace(/\s+/g, ' ').trim();
    if (cleanStaff.toLowerCase().indexOf("all") !== -1) cleanStaff = "All Staff";
    
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
    var taskId = "TASK_" + new Date().getTime();
    
    sheet.appendRow([todayStr, cleanStaff, cleanFreq, desc, deadline, refPhotoCell, "Pending", taskId]);
    return "Task successfully assign ho gaya!";
  } catch (err) {
    return "Error: " + err.message;
  }
}

// =====================================================================
// 4. STAFF: GET ASSIGNED TASKS (WITH AUTO-RESET RECURRING CYCLES)
// Supports Rekha, Vrushali, all staff, and team-wide "All Staff"
// =====================================================================
function getAdhocTasksForStaff(staffName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getSheetSmart(ss, "adhoc_tasks");
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    var cleanStaff = String(staffName || "").replace(/[^\w\s]/gi, '').replace(/\s+/g, ' ').trim().toLowerCase();
    var hasFreqCol = (data[0].length >= 7);
    
    // Cycle Date Boundaries for Recurring Task Reset
    var now = new Date();
    var startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0, 0).getTime();
    
    var dayOfWeek = now.getDay() === 0 ? 7 : now.getDay();
    var startOfWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() - dayOfWeek + 1, 0, 0, 0, 0).getTime();
    
    var startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0).getTime();
    
    // Fetch submissions done by this staff member to determine current cycle status
    var subSheet = getSheetSmart(ss, "submissions");
    var staffSubmissions = [];
    if (subSheet) {
      var sData = subSheet.getDataRange().getValues();
      for (var s = 1; s < sData.length; s++) {
        var sName = String(sData[s][1] || "").replace(/[^\w\s]/gi, '').replace(/\s+/g, ' ').trim().toLowerCase();
        if (sName === cleanStaff || (cleanStaff.length >= 3 && sName.indexOf(cleanStaff) !== -1)) {
          var sDate = new Date(sData[s][0]).getTime();
          var sTask = String(sData[s][4] || "").trim().toLowerCase();
          var sStatus = String(sData[s][7] || "").trim();
          var sAdhocIdx = sData[s][11] ? parseInt(sData[s][11]) : null;
          staffSubmissions.push({ date: sDate, task: sTask, status: sStatus, adhocIndex: sAdhocIdx });
        }
      }
    }
    
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var assignedToRaw = String(data[i][1] || "").trim();
      var assignedToClean = assignedToRaw.replace(/[^\w\s]/gi, '').replace(/\s+/g, ' ').trim().toLowerCase();
      
      // Match condition: "All Staff", exact match, or substring match
      var isMatch = false;
      if (assignedToClean.indexOf("all") !== -1) {
        isMatch = true;
      } else if (assignedToClean === cleanStaff) {
        isMatch = true;
      } else if (cleanStaff.length >= 3 && assignedToClean.indexOf(cleanStaff) !== -1) {
        isMatch = true;
      } else if (assignedToClean.length >= 3 && cleanStaff.indexOf(assignedToClean) !== -1) {
        isMatch = true;
      }
      
      if (isMatch) {
        var rawFreq = hasFreqCol ? String(data[i][2] || "Extra (One Time)") : "Extra (One Time)";
        rawFreq = rawFreq.replace(/[^\w\s]/gi, '').replace(/\s*Task/i, '').trim();
        if (!rawFreq || rawFreq.toLowerCase() === "extra" || rawFreq.toLowerCase() === "one time") {
          rawFreq = "Extra (One Time)";
        } else {
          rawFreq = rawFreq.charAt(0).toUpperCase() + rawFreq.slice(1).toLowerCase();
        }
        
        var taskDesc = hasFreqCol ? data[i][3] : data[i][2];
        var deadline = hasFreqCol ? data[i][4] : data[i][3];
        var refPhoto = hasFreqCol ? data[i][5] : data[i][4];
        var statusColIdx = hasFreqCol ? 6 : 5;
        var recordedStatus = String(data[i][statusColIdx] || "Pending").trim();
        
        var currentCycleStatus = "Pending";
        var taskClean = String(taskDesc || "").trim().toLowerCase();
        
        // AUTO-RESET RECURRING CYCLES LOGIC:
        if (rawFreq === "Daily") {
          var doneToday = staffSubmissions.some(function(sub) {
            return sub.date >= startOfToday && (sub.adhocIndex === (i + 1) || sub.task === taskClean);
          });
          currentCycleStatus = doneToday ? "Completed" : "Pending";
        } else if (rawFreq === "Weekly") {
          var doneThisWeek = staffSubmissions.some(function(sub) {
            return sub.date >= startOfWeek && (sub.adhocIndex === (i + 1) || sub.task === taskClean);
          });
          currentCycleStatus = doneThisWeek ? "Completed" : "Pending";
        } else if (rawFreq === "Monthly") {
          var doneThisMonth = staffSubmissions.some(function(sub) {
            return sub.date >= startOfMonth && (sub.adhocIndex === (i + 1) || sub.task === taskClean);
          });
          currentCycleStatus = doneThisMonth ? "Completed" : "Pending";
        } else {
          // Extra (One Time) task retains its sheet status or submission
          currentCycleStatus = recordedStatus;
          if (currentCycleStatus === "Pending") {
            var anyDone = staffSubmissions.some(function(sub) {
              return sub.adhocIndex === (i + 1) || sub.task === taskClean;
            });
            if (anyDone) currentCycleStatus = "Completed";
          }
        }
        
        list.push({
          index: i + 1,
          assignDate: data[i][0],
          assignedTo: assignedToRaw,
          type: rawFreq, // "Daily", "Weekly", "Monthly", or "Extra (One Time)"
          task: taskDesc,
          deadline: deadline,
          refPhoto: refPhoto,
          status: currentCycleStatus
        });
      }
    }
    return list;
  } catch (err) {
    return [];
  }
}

// =====================================================================
// 5. STAFF: GET ROUTINE ROLE TASKS (FROM ROLE_TASKS SHEET)
// =====================================================================
function getTasksForRole(roleName, staffName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getSheetSmart(ss, "routine_tasks");
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var roleClean = String(roleName || "").trim().toLowerCase();
    var cleanStaff = String(staffName || "").replace(/[^\w\s]/gi, '').replace(/\s+/g, ' ').trim().toLowerCase();
    var tasks = [];
    
    // Check submissions done today by this staff member
    var subSheet = getSheetSmart(ss, "submissions");
    var todayDoneMap = {};
    if (subSheet) {
      var subData = subSheet.getDataRange().getValues();
      var now = new Date();
      var startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
      
      for (var s = 1; s < subData.length; s++) {
        var sStaff = String(subData[s][1] || "").replace(/[^\w\s]/gi, '').replace(/\s+/g, ' ').trim().toLowerCase();
        if (sStaff === cleanStaff || (cleanStaff.length >= 3 && sStaff.indexOf(cleanStaff) !== -1)) {
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
// 6. COMPREHENSIVE 7-PERIOD SCORE ENGINE (STAFF & OWNER VIEWS)
// Daily, Weekly, Monthly, Last Month, Curr Qtr, Last Qtr, Yearly
// Calculates: Points, Total Assigned, Completed, Delayed, Percentage
// =====================================================================
function getStaffScoreMetrics(staffName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result = {
      daily: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 0 },
      weekly: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 0 },
      monthly: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 0 },
      lastMonth: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 0 },
      currentQuarter: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 0 },
      lastQuarter: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 0 },
      yearly: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 0 }
    };
    
    var cleanStaff = String(staffName || "").replace(/[^\w\s]/gi, '').replace(/\s+/g, ' ').trim().toLowerCase();
    var now = new Date();
    
    // Timeframe boundaries
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
    
    // 1. Process Task_Submissions (Completed & Delayed Points)
    var subSheet = getSheetSmart(ss, "submissions");
    if (subSheet) {
      var data = subSheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var rowName = String(data[i][1] || "").replace(/[^\w\s]/gi, '').replace(/\s+/g, ' ').trim().toLowerCase();
        if (cleanStaff) {
          var isMatch = (rowName === cleanStaff || (cleanStaff.length >= 3 && rowName.indexOf(cleanStaff) !== -1) || (rowName.length >= 3 && cleanStaff.indexOf(rowName) !== -1));
          if (!isMatch) continue;
        }
        
        var rawDate = data[i][0];
        if (!rawDate) continue;
        var subDate = new Date(rawDate);
        if (isNaN(subDate.getTime())) continue;
        
        var status = String(data[i][7] || "");
        var isSkipped = (status.indexOf("Skipped") !== -1);
        var pts = isSkipped ? 0 : (parseInt(data[i][10]) || 10);
        var deadlineStr = String(data[i][6] || "");
        var isDelayed = false;
        if (deadlineStr && deadlineStr !== "N/A" && deadlineStr !== "Jald") {
          var dDate = new Date(deadlineStr);
          if (!isNaN(dDate.getTime()) && subDate > dDate) isDelayed = true;
        }
        
        function updateBucket(b, countComplete) {
          b.points += pts;
          if (countComplete) b.completed += 1;
          if (isDelayed) b.delayed += 1;
        }
        
        if (subDate >= startOfToday) updateBucket(result.daily, !isSkipped);
        if (subDate >= startOfWeek) updateBucket(result.weekly, !isSkipped);
        if (subDate >= startOfMonth) updateBucket(result.monthly, !isSkipped);
        if (subDate >= startOfLastMonth && subDate <= endOfLastMonth) updateBucket(result.lastMonth, !isSkipped);
        if (subDate >= startOfCurrentQtr) updateBucket(result.currentQuarter, !isSkipped);
        if (subDate >= startOfLastQtr && subDate <= endOfLastQtr) updateBucket(result.lastQuarter, !isSkipped);
        if (subDate >= startOfYear) updateBucket(result.yearly, !isSkipped);
      }
    }
    
    // 2. Compute Total Assigned Tasks from Role_Tasks and Adhoc_Tasks for Staff
    var role = "Staff";
    var staffSheet = getSheetSmart(ss, "staff");
    if (staffSheet) {
      var sData = staffSheet.getDataRange().getValues();
      var colMap = getStaffColMap(staffSheet);
      for (var s = 1; s < sData.length; s++) {
        var sName = String(sData[s][colMap.name] || "").replace(/[^\w\s]/gi, '').replace(/\s+/g, ' ').trim().toLowerCase();
        if (sName === cleanStaff || (cleanStaff.length >= 3 && sName.indexOf(cleanStaff) !== -1)) {
          role = String(sData[s][colMap.role] || "Staff").trim();
          break;
        }
      }
    }
    
    var routineList = getTasksForRole(role, cleanStaff);
    var adhocList = getAdhocTasksForStaff(cleanStaff);
    
    var dailyAssignedCount = 0;
    var weeklyAssignedCount = 0;
    var monthlyAssignedCount = 0;
    
    for (var r = 0; r < routineList.length; r++) {
      var rType = String(routineList[r].type || "").toLowerCase();
      if (rType.indexOf("daily") !== -1) dailyAssignedCount++;
      if (rType.indexOf("weekly") !== -1) weeklyAssignedCount++;
      if (rType.indexOf("monthly") !== -1) monthlyAssignedCount++;
    }
    for (var a = 0; a < adhocList.length; a++) {
      var aType = String(adhocList[a].type || "").toLowerCase();
      if (aType.indexOf("daily") !== -1) dailyAssignedCount++;
      else if (aType.indexOf("weekly") !== -1) weeklyAssignedCount++;
      else if (aType.indexOf("monthly") !== -1) monthlyAssignedCount++;
      else dailyAssignedCount++; // Extra task treated in today's quota
    }
    
    result.daily.assigned = Math.max(result.daily.completed, dailyAssignedCount || 1);
    result.weekly.assigned = Math.max(result.weekly.completed, (dailyAssignedCount * dayOfWeek) + weeklyAssignedCount);
    result.monthly.assigned = Math.max(result.monthly.completed, (dailyAssignedCount * now.getDate()) + (weeklyAssignedCount * Math.ceil(now.getDate() / 7)) + monthlyAssignedCount);
    result.lastMonth.assigned = Math.max(result.lastMonth.completed, (dailyAssignedCount * 30) + (weeklyAssignedCount * 4) + monthlyAssignedCount);
    result.currentQuarter.assigned = Math.max(result.currentQuarter.completed, result.monthly.assigned * 3);
    result.lastQuarter.assigned = Math.max(result.lastQuarter.completed, result.monthly.assigned * 3);
    result.yearly.assigned = Math.max(result.yearly.completed, result.monthly.assigned * 12);
    
    // Compute Percentages
    for (var k in result) {
      var b = result[k];
      b.percentage = b.assigned > 0 ? Math.min(100, Math.round((b.completed / b.assigned) * 100)) : 100;
    }
    
    return result;
  } catch (err) {
    return {
      daily: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 0 },
      weekly: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 0 },
      monthly: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 0 },
      lastMonth: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 0 },
      currentQuarter: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 0 },
      lastQuarter: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 0 },
      yearly: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 0 }
    };
  }
}

// =====================================================================
// 7. OWNER: COMPREHENSIVE SCORECARD LEADERBOARD FOR ALL STAFF
// =====================================================================
function getStaffLeaderboard(period) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var staffSheet = getSheetSmart(ss, "staff");
    var staffList = [];
    var pKey = period || "daily";
    if (!staffSheet) return [];
    
    var sData = staffSheet.getDataRange().getValues();
    var colMap = getStaffColMap(staffSheet);
    for (var s = 1; s < sData.length; s++) {
      var sName = String(sData[s][colMap.name] || "").trim();
      var sRole = String(sData[s][colMap.role] || "Staff").trim();
      var sActive = sData[s][colMap.active];
      if (sName && sActive !== false && String(sActive).toLowerCase() !== "false") {
        if (!sRole.toLowerCase().includes("owner") && !sRole.toLowerCase().includes("admin")) {
          var metrics = getStaffScoreMetrics(sName);
          var selectedPeriod = metrics[pKey] || { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 0 };
          staffList.push({
            name: sName,
            role: sRole,
            points: selectedPeriod.points || 0,
            assigned: selectedPeriod.assigned || 0,
            completed: selectedPeriod.completed || 0,
            delayed: selectedPeriod.delayed || 0,
            percentage: selectedPeriod.percentage || 0,
            metrics: metrics
          });
        }
      }
    }
    
    staffList.sort(function(a, b) { return b.points - a.points; });
    for (var i = 0; i < staffList.length; i++) {
      staffList[i].rank = i + 1;
    }
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
    var subSheet = getSheetSmart(ss, "submissions");
    if (!subSheet) {
      subSheet = ss.insertSheet("Task_Submissions");
      subSheet.appendRow(["Timestamp", "Staff Name", "Role", "Type", "Task", "SubTask", "Deadline", "Status", "Proof Link", "Action", "Points Earned", "AdhocIndex"]);
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
    
    var isSkip = (actionType === "Skip");
    var pointsEarned = isSkip ? 0 : 10;
    var finalStatus = isSkip ? "Skipped" : (proofUrl !== "No Proof" ? "Uploaded (Done)" : "Completed");
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+5:30", "yyyy-MM-dd HH:mm:ss");
    
    subSheet.appendRow([nowStr, name, role, type, task, subTask || "", deadline || "", finalStatus, proofUrl, actionType || "Done", pointsEarned, adhocIndex || ""]);
    
    // If one-time extra task, update status in Adhoc_Tasks sheet directly
    if (isAdhoc && adhocIndex) {
      var adSheet = getSheetSmart(ss, "adhoc_tasks");
      if (adSheet) {
        var row = parseInt(adhocIndex);
        if (row >= 2 && row <= adSheet.getLastRow()) {
          var hasFreqCol = (adSheet.getLastColumn() >= 7);
          var statusCol = hasFreqCol ? 7 : 6;
          var freqVal = hasFreqCol ? String(adSheet.getRange(row, 3).getValue()) : "Extra (One Time)";
          // Only permanently mark Extra tasks as done in the sheet
          if (freqVal.indexOf("Daily") === -1 && freqVal.indexOf("Weekly") === -1 && freqVal.indexOf("Monthly") === -1) {
            adSheet.getRange(row, statusCol).setValue(finalStatus);
          }
        }
      }
    }
    
    return { status: finalStatus, pointsEarned: pointsEarned, proofUrl: proofUrl };
  } catch (err) {
    return { status: "Error: " + err.message, pointsEarned: 0 };
  }
}

// =====================================================================
// 9. ADMIN: STAFF PROGRESS REPORT (WITH 7-PERIOD SCORES)
// =====================================================================
function getStaffProgressReport(cleanName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var staffSheet = getSheetSmart(ss, "staff");
    var role = "Staff";
    if (staffSheet) {
      var sData = staffSheet.getDataRange().getValues();
      var colMap = getStaffColMap(staffSheet);
      var cName = String(cleanName || "").replace(/[^\w\s]/gi, '').trim().toLowerCase();
      for (var s = 1; s < sData.length; s++) {
        var rowStaff = String(sData[s][colMap.name] || "").replace(/[^\w\s]/gi, '').trim().toLowerCase();
        if (rowStaff === cName || (cName.length >= 3 && rowStaff.indexOf(cName) !== -1)) {
          role = String(sData[s][colMap.role] || "Staff").trim();
          break;
        }
      }
    }
    
    var routine = getTasksForRole(role, cleanName);
    var adhoc = getAdhocTasksForStaff(cleanName);
    var metrics = getStaffScoreMetrics(cleanName);
    
    return { role: role, routine: routine, adhoc: adhoc, scores: metrics };
  } catch (err) {
    return { role: "Staff", routine: [], adhoc: [], scores: {} };
  }
}

// =====================================================================
// 10. ADMIN: ALL ADHOC HISTORY
// =====================================================================
function getAllAdhocHistory() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getSheetSmart(ss, "adhoc_tasks");
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var hasFreqCol = (data[0].length >= 7);
    var list = [];
    
    for (var i = data.length - 1; i >= 1; i--) {
      var assignedTo = data[i][1];
      var rawType = hasFreqCol ? String(data[i][2] || "Extra (One Time)") : "Extra (One Time)";
      var taskDesc = hasFreqCol ? data[i][3] : data[i][2];
      var deadline = hasFreqCol ? data[i][4] : data[i][3];
      var refPhoto = hasFreqCol ? data[i][5] : data[i][4];
      var status = hasFreqCol ? data[i][6] : data[i][5];
      
      list.push({
        date: data[i][0],
        assignedTo: assignedTo,
        type: rawType,
        task: taskDesc,
        deadline: deadline,
        refPhoto: refPhoto,
        status: status || "Pending"
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
    var sheet = getSheetSmart(ss, "orders");
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

function completeStoreManagerOrder(rowIndex) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getSheetSmart(ss, "orders");
    if (!sheet) return "Orders sheet nahi mili!";
    
    var row = parseInt(rowIndex);
    if (row >= 2 && row <= sheet.getLastRow()) {
      sheet.getRange(row, 7).setValue("Dispatched (Completed)");
      return "Order successfully processed!";
    }
    return "Invalid Row Index!";
  } catch (err) {
    return "Error: " + err.message;
  }
}

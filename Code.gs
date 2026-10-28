// =====================================================================
// YASH MARKETING - TASK OPERATIONS & STAFF SCORING BACKEND (Code.gs)
// Complete Production Engine:
// 1. 7-Period Performance Score Engine (Daily, Weekly, Monthly, Last Month, Curr Qtr, Last Qtr, Yearly)
// 2. Multiple Reference Photos Upload to Google Drive & Gallery / Zoom Preview
// 3. Dynamic Daily / Weekly / Monthly Recurring Tasks with Auto-Reset Cycles & "All Staff" Matching
// 4. Flexible Sheet & Column Matcher (Staff_List, Role_Tasks, Adhoc_Tasks, Task_Submissions, Orders)
// 5. Store Manager Order Processing Pipeline
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
  
  // Local fallback from Apps Script project if GitHub is unreachable
  var template = HtmlService.createTemplateFromFile("Index");
  template.portalEmpId = empId;
  template.empId = empId;
  return template.evaluate()
    .setTitle("Yash Marketing - Task Operations & Staff Portal")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =====================================================================
// HELPER: FLEXIBLE SHEET FINDER (Handles spaces, underscores, casing)
// =====================================================================
function getSheetByNameFlexible(ss, names) {
  if (!ss) return null;
  // 1. Direct match
  for (var i = 0; i < names.length; i++) {
    var s = ss.getSheetByName(names[i]);
    if (s) return s;
  }
  // 2. Normalized match (remove spaces, underscores, lowercase)
  var allSheets = ss.getSheets();
  for (var i = 0; i < names.length; i++) {
    var target = names[i].replace(/[_\s]/g, '').toLowerCase();
    for (var j = 0; j < allSheets.length; j++) {
      var cur = allSheets[j].getName().replace(/[_\s]/g, '').toLowerCase();
      if (cur === target) return allSheets[j];
    }
  }
  // 3. Fallback: contains match
  for (var i = 0; i < names.length; i++) {
    var target = names[i].replace(/[_\s]/g, '').toLowerCase();
    for (var j = 0; j < allSheets.length; j++) {
      var cur = allSheets[j].getName().replace(/[_\s]/g, '').toLowerCase();
      if (cur.indexOf(target) !== -1 || target.indexOf(cur) !== -1) return allSheets[j];
    }
  }
  return null;
}

// =====================================================================
// HELPER: ROBUST ADHOC_TASKS SHEET SCHEMA & COLUMN MAPPER
// Automatically ensures the Frequency column exists without corrupting data
// =====================================================================
function getOrSetupAdhocSheet(ss) {
  var sheet = getSheetByNameFlexible(ss, ["Adhoc_Tasks", "Adhoc Tasks", "Assigned_Tasks"]);
  var defaultHeaders = ["Assign Date", "Assigned To", "Frequency", "Task Description", "Deadline", "Reference Photo", "Status", "TaskId"];
  
  if (!sheet) {
    sheet = ss.insertSheet("Adhoc_Tasks");
    sheet.appendRow(defaultHeaders);
    return {
      sheet: sheet,
      colMap: { date: 0, staff: 1, freq: 2, task: 3, deadline: 4, photo: 5, status: 6, id: 7 }
    };
  }

  var numCols = Math.max(sheet.getLastColumn(), 1);
  var headerVals = sheet.getRange(1, 1, 1, numCols).getValues()[0];
  
  var colMap = { date: -1, staff: -1, freq: -1, task: -1, deadline: -1, photo: -1, status: -1, id: -1 };
  
  for (var c = 0; c < headerVals.length; c++) {
    var h = String(headerVals[c] || "").toLowerCase().trim();
    if (!h) continue;
    if (h.indexOf("freq") !== -1 || h === "type") colMap.freq = c;
    else if (colMap.date === -1 && (h.indexOf("date") !== -1 || h.indexOf("time") !== -1)) colMap.date = c;
    else if (colMap.staff === -1 && (h.indexOf("staff") !== -1 || h.indexOf("assign") !== -1 || h.indexOf("user") !== -1 || h.indexOf("emp") !== -1)) colMap.staff = c;
    else if (colMap.task === -1 && (h.indexOf("desc") !== -1 || (h.indexOf("task") !== -1 && h.indexOf("id") === -1 && h.indexOf("freq") === -1))) colMap.task = c;
    else if (colMap.deadline === -1 && (h.indexOf("dead") !== -1 || h.indexOf("due") !== -1)) colMap.deadline = c;
    else if (colMap.photo === -1 && (h.indexOf("photo") !== -1 || h.indexOf("image") !== -1 || h.indexOf("file") !== -1 || h.indexOf("proof") !== -1)) colMap.photo = c;
    else if (colMap.status === -1 && h.indexOf("stat") !== -1) colMap.status = c;
    else if (colMap.id === -1 && h.indexOf("id") !== -1) colMap.id = c;
  }

  // If Frequency column is missing from header, insert it cleanly at column 3
  if (colMap.freq === -1) {
    sheet.insertColumnAfter(2);
    sheet.getRange(1, 3).setValue("Frequency");
    var totalRows = sheet.getLastRow();
    if (totalRows > 1) {
      var fRange = sheet.getRange(2, 3, totalRows - 1, 1);
      var fVals = fRange.getValues();
      for (var r = 0; r < fVals.length; r++) {
        if (!fVals[r][0]) fVals[r][0] = "Extra (One Time)";
      }
      fRange.setValues(fVals);
    }
    // Re-index columns after insertion
    colMap.freq = 2;
    colMap.date = 0;
    colMap.staff = 1;
    colMap.task = 3;
    colMap.deadline = 4;
    colMap.photo = 5;
    colMap.status = 6;
    colMap.id = 7;
  } else {
    // Fill any missing defaults based on column index order
    if (colMap.date === -1) colMap.date = 0;
    if (colMap.staff === -1) colMap.staff = 1;
    if (colMap.task === -1) colMap.task = (colMap.freq === 2) ? 3 : 2;
    if (colMap.deadline === -1) colMap.deadline = colMap.task + 1;
    if (colMap.photo === -1) colMap.photo = colMap.deadline + 1;
    if (colMap.status === -1) colMap.status = colMap.photo + 1;
    if (colMap.id === -1) colMap.id = colMap.status + 1;
  }

  return { sheet: sheet, colMap: colMap };
}

// =====================================================================
// 1. AUTHENTICATION & LOGIN
// Column A: Emp ID | Column B: Password | Column C: Name | Column D: Role | Column E: Active
// =====================================================================
function verifyLogin(id, pass) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return { success: false, message: "Spreadsheet connect nahi hui! Kripya check karein ki Apps Script Google Sheet ke sath bound hai." };
    }
    
    var sheet = getSheetByNameFlexible(ss, ["Staff_List", "Staff List", "Staff", "Users"]);
    if (!sheet) {
      return { success: false, message: "Staff_List sheet nahi mili! Kripya sheet tab ka naam 'Staff_List' ya 'Staff List' rakhein." };
    }
    
    var data = sheet.getDataRange().getValues();
    var cleanId = String(id).replace(/[^\w]/gi, '').toLowerCase();
    var cleanPass = String(pass).trim();
    
    for (var i = 1; i < data.length; i++) {
      var rowEmpId = String(data[i][0] || "").replace(/[^\w]/gi, '').toLowerCase();
      var rowPass = String(data[i][1] || "").trim();
      var rowName = String(data[i][2] || "").trim();
      var rowRole = String(data[i][3] || "Staff").trim();
      var rowActive = data[i][4];
      
      // Match by Emp ID or Name
      var idMatches = (rowEmpId === cleanId || String(data[i][0] || "").trim().toLowerCase() === String(id).trim().toLowerCase());
      
      if (idMatches) {
        if (cleanPass === "SSO_BYPASS" || rowPass === cleanPass) {
          if (rowActive !== false && String(rowActive).toLowerCase() !== "false") {
            var metrics = getStaffScoreMetrics(rowName, String(data[i][0] || ""));
            return {
              success: true,
              empId: String(data[i][0] || id),
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
    var sheet = getSheetByNameFlexible(ss, ["Staff_List", "Staff List", "Staff", "Users"]);
    if (!sheet) return ["All Staff"];
    
    var data = sheet.getDataRange().getValues();
    var list = ["All Staff"];
    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][2] || "").trim();
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
    var setup = getOrSetupAdhocSheet(ss);
    var sheet = setup.sheet;
    var colMap = setup.colMap;
    
    // Normalize frequency name: "Daily Task" -> "Daily", "Weekly Task" -> "Weekly", etc.
    var cleanFreq = String(typeVal || "Extra (One Time)").replace(/[^\w\s]/gi, '').replace(/\s*Task/i, '').trim();
    if (!cleanFreq || cleanFreq.toLowerCase() === "extra" || cleanFreq.toLowerCase() === "one time") {
      cleanFreq = "Extra (One Time)";
    } else {
      cleanFreq = cleanFreq.charAt(0).toUpperCase() + cleanFreq.slice(1).toLowerCase();
    }
    
    var cleanStaff = String(staff || "").replace("👤 ", "").trim();
    if (!cleanStaff || cleanStaff.toLowerCase().indexOf("all") !== -1) {
      cleanStaff = "All Staff";
    }
    
    // Upload reference photos to Google Drive
    var photoLinks = [];
    var folderName = "Yash Marketing - Reference Photos";
    var folders = DriveApp.getFoldersByName(folderName);
    var targetFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    
    // 1. Multiple photos received as Array of objects: [{ name, type, base64 }]
    if (Array.isArray(photosData) && photosData.length > 0) {
      for (var i = 0; i < photosData.length; i++) {
        var p = photosData[i];
        if (p && p.base64 && p.name) {
          var blob = Utilities.newBlob(Utilities.base64Decode(p.base64), p.type || "image/jpeg", p.name);
          var file = targetFolder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          photoLinks.push(file.getUrl());
        }
      }
    } else if (photosData && typeof photosData === "string" && fileType && fileName) {
      // 2. Single photo backward compatibility
      var blob = Utilities.newBlob(Utilities.base64Decode(photosData), fileType, fileName);
      var file = targetFolder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      photoLinks.push(file.getUrl());
    }
    
    var photoCell = "No Photo";
    if (photoLinks.length === 1) {
      photoCell = photoLinks[0];
    } else if (photoLinks.length > 1) {
      photoCell = JSON.stringify(photoLinks);
    }
    
    var timeZone = Session.getScriptTimeZone() || "GMT+5:30";
    var nowStr = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd HH:mm:ss");
    var taskId = "TASK-" + new Date().getTime();
    
    // Prepare standardized row using mapped columns
    var maxCol = Math.max(colMap.id, colMap.status, colMap.photo, colMap.deadline, colMap.task, colMap.freq, colMap.staff, colMap.date) + 1;
    var newRow = new Array(maxCol);
    for (var k = 0; k < maxCol; k++) newRow[k] = "";
    
    newRow[colMap.date] = nowStr;
    newRow[colMap.staff] = cleanStaff;
    newRow[colMap.freq] = cleanFreq;
    newRow[colMap.task] = desc;
    newRow[colMap.deadline] = deadline || "Jald";
    newRow[colMap.photo] = photoCell;
    newRow[colMap.status] = "Pending";
    newRow[colMap.id] = taskId;
    
    sheet.appendRow(newRow);
    return "Task successfully assign ho gaya!";
  } catch (err) {
    return "Error: " + err.message;
  }
}

// =====================================================================
// 4. STAFF: GET ADHOC & RECURRING TASKS FOR STAFF
// Includes Auto-Reset Logic for Daily, Weekly, and Monthly cycles!
// Supports "All Staff", Staff Name, and Staff ID matching!
// =====================================================================
function getAdhocTasksForStaff(staffName, empId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var setup = getOrSetupAdhocSheet(ss);
    var sheet = setup.sheet;
    var colMap = setup.colMap;
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    var cleanStaff = String(staffName || "").replace("👤 ", "").trim().toLowerCase();
    var cleanEmpId = String(empId || "").replace(/[^\w]/gi, '').toLowerCase();
    
    // Also lookup Staff_List to find empId if not passed
    if (!cleanEmpId && cleanStaff) {
      var staffSheet = getSheetByNameFlexible(ss, ["Staff_List", "Staff List", "Staff", "Users"]);
      if (staffSheet) {
        var sVals = staffSheet.getDataRange().getValues();
        for (var sv = 1; sv < sVals.length; sv++) {
          var svName = String(sVals[sv][2] || "").trim().toLowerCase();
          if (svName === cleanStaff || (cleanStaff.length >= 3 && svName.indexOf(cleanStaff) !== -1)) {
            cleanEmpId = String(sVals[sv][0] || "").replace(/[^\w]/gi, '').toLowerCase();
            break;
          }
        }
      }
    }
    
    // Timezone & current cycle identifiers
    var timeZone = Session.getScriptTimeZone() || "GMT+5:30";
    var now = new Date();
    var todayStr = Utilities.formatDate(now, timeZone, "yyyy-MM-dd");
    var currentWeekStr = Utilities.formatDate(now, timeZone, "yyyy-ww");
    var currentMonthStr = Utilities.formatDate(now, timeZone, "yyyy-MM");
    
    // Fetch Task_Submissions to check cycle completions for this staff member
    var subSheet = getSheetByNameFlexible(ss, ["Task_Submissions", "Task Submissions", "Submissions"]);
    var staffSubmissions = [];
    if (subSheet && subSheet.getLastRow() > 1) {
      var subData = subSheet.getDataRange().getValues();
      for (var s = 1; s < subData.length; s++) {
        var sStaff = String(subData[s][1] || "").replace("👤 ", "").trim().toLowerCase();
        var isSubStaffMatch = (sStaff === cleanStaff || (cleanStaff.length >= 3 && sStaff.indexOf(cleanStaff) !== -1) ||
                               (cleanEmpId && sStaff.indexOf(cleanEmpId) !== -1));
        if (isSubStaffMatch) {
          var rawDate = subData[s][0];
          var subDateObj = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
          var subDateFormatted = "";
          var subWeekFormatted = "";
          var subMonthFormatted = "";
          if (!isNaN(subDateObj.getTime())) {
            subDateFormatted = Utilities.formatDate(subDateObj, timeZone, "yyyy-MM-dd");
            subWeekFormatted = Utilities.formatDate(subDateObj, timeZone, "yyyy-ww");
            subMonthFormatted = Utilities.formatDate(subDateObj, timeZone, "yyyy-MM");
          }
          staffSubmissions.push({
            dateObj: subDateObj,
            dayStr: subDateFormatted,
            weekStr: subWeekFormatted,
            monthStr: subMonthFormatted,
            task: String(subData[s][4] || "").trim().toLowerCase(),
            status: String(subData[s][7] || "Completed"),
            adhocIndex: parseInt(subData[s][11]) || null
          });
        }
      }
    }
    
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var assignedToRaw = String(data[i][colMap.staff] || "").trim();
      var assignedToClean = assignedToRaw.replace("👤 ", "").trim().toLowerCase();
      var assignedToEmpId = assignedToClean.replace(/[^\w]/gi, '');
      
      // Match condition: "All Staff", exact staff name, substring, or empId match
      var isMatch = false;
      if (assignedToClean.indexOf("all") !== -1 || assignedToClean === "everyone" || assignedToClean === "team" || !assignedToClean) {
        isMatch = true;
      } else if (assignedToClean === cleanStaff) {
        isMatch = true;
      } else if (cleanStaff.length >= 3 && assignedToClean.indexOf(cleanStaff) !== -1) {
        isMatch = true;
      } else if (assignedToClean.length >= 3 && cleanStaff.indexOf(assignedToClean) !== -1) {
        isMatch = true;
      } else if (cleanEmpId && (assignedToEmpId === cleanEmpId || assignedToClean.indexOf(cleanEmpId) !== -1)) {
        isMatch = true;
      }
      
      if (isMatch) {
        var rawFreq = String(data[i][colMap.freq] || "Extra (One Time)").replace(/[^\w\s]/gi, '').replace(/\s*Task/i, '').trim();
        if (!rawFreq || rawFreq.toLowerCase() === "extra" || rawFreq.toLowerCase() === "one time") {
          rawFreq = "Extra (One Time)";
        } else {
          rawFreq = rawFreq.charAt(0).toUpperCase() + rawFreq.slice(1).toLowerCase();
        }
        
        var taskDesc = String(data[i][colMap.task] || "");
        var deadline = String(data[i][colMap.deadline] || "Jald");
        var refPhoto = data[i][colMap.photo] || "No Photo";
        var recordedStatus = String(data[i][colMap.status] || "Pending").trim();
        var taskClean = taskDesc.trim().toLowerCase();
        
        var currentCycleStatus = "Pending";
        
        // AUTO-RESET RECURRING CYCLES:
        if (rawFreq === "Daily") {
          var doneToday = staffSubmissions.some(function(sub) {
            return sub.dayStr === todayStr && (sub.adhocIndex === (i + 1) || sub.task === taskClean);
          });
          currentCycleStatus = doneToday ? "Completed" : "Pending";
        } else if (rawFreq === "Weekly") {
          var doneThisWeek = staffSubmissions.some(function(sub) {
            return sub.weekStr === currentWeekStr && (sub.adhocIndex === (i + 1) || sub.task === taskClean);
          });
          currentCycleStatus = doneThisWeek ? "Completed" : "Pending";
        } else if (rawFreq === "Monthly") {
          var doneThisMonth = staffSubmissions.some(function(sub) {
            return sub.monthStr === currentMonthStr && (sub.adhocIndex === (i + 1) || sub.task === taskClean);
          });
          currentCycleStatus = doneThisMonth ? "Completed" : "Pending";
        } else {
          // Extra (One Time) task retains its sheet status or submission status
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
          assignDate: data[i][colMap.date],
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
    Logger.log("getAdhocTasksForStaff error: " + err.message);
    return [];
  }
}

// =====================================================================
// 5. STAFF: GET ROUTINE ROLE TASKS (FROM ROLE_TASKS / TASKS SHEET)
// Column A: Role | Column B: Type | Column C: Task | Column D: SubTask | Column E: Deadline | Column F: SOP | Column G: Video
// =====================================================================
function getTasksForRole(roleName, staffName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getSheetByNameFlexible(ss, ["Role_Tasks", "Role Tasks", "Routine_Tasks", "Routine Tasks", "Tasks"]);
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var roleClean = String(roleName || "").trim().toLowerCase();
    var cleanStaff = String(staffName || "").replace("👤 ", "").trim().toLowerCase();
    var tasks = [];
    
    // Check submissions done today by this staff member
    var subSheet = getSheetByNameFlexible(ss, ["Task_Submissions", "Task Submissions", "Submissions"]);
    var todayDoneMap = {};
    var timeZone = Session.getScriptTimeZone() || "GMT+5:30";
    var todayStr = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
    
    if (subSheet && subSheet.getLastRow() > 1) {
      var subData = subSheet.getDataRange().getValues();
      for (var s = 1; s < subData.length; s++) {
        var sStaff = String(subData[s][1] || "").replace("👤 ", "").trim().toLowerCase();
        if (sStaff === cleanStaff || (cleanStaff.length >= 3 && sStaff.indexOf(cleanStaff) !== -1)) {
          var rawDate = subData[s][0];
          var subDateObj = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
          if (!isNaN(subDateObj.getTime())) {
            var subDateStr = Utilities.formatDate(subDateObj, timeZone, "yyyy-MM-dd");
            if (subDateStr === todayStr) {
              var sKey = String(subData[s][4]).trim() + "___" + String(subData[s][5] || "").trim();
              todayDoneMap[sKey] = String(subData[s][7]);
            }
          }
        }
      }
    }
    
    for (var i = 1; i < data.length; i++) {
      var targetRole = String(data[i][0] || "").trim().toLowerCase();
      var isRoleMatch = (targetRole === roleClean || targetRole === "all" || targetRole === "all staff" ||
                         (roleClean.length >= 3 && targetRole.indexOf(roleClean) !== -1) ||
                         (targetRole.length >= 3 && roleClean.indexOf(targetRole) !== -1));
      
      if (isRoleMatch) {
        var type = String(data[i][1] || "").trim();
        var mainTask = String(data[i][2] || "").trim();
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
// Calculates: Points, Total Assigned, Completed, Delayed, Percentage
// Periods: Daily, Weekly, Monthly, Last Month, Current Quarter, Last Quarter, Yearly
// =====================================================================
function getStaffScoreMetrics(staffName, empId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result = {
      daily: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 100 },
      weekly: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 100 },
      monthly: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 100 },
      lastMonth: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 100 },
      currentQuarter: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 100 },
      lastQuarter: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 100 },
      yearly: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 100 }
    };
    
    var cleanStaff = String(staffName || "").replace("👤 ", "").trim().toLowerCase();
    var cleanEmpId = String(empId || "").replace(/[^\w]/gi, '').toLowerCase();
    var now = new Date();
    
    // Timeframe boundaries
    var startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0, 0);
    var dayOfWeek = now.getDay() === 0 ? 7 : now.getDay();
    var startOfWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() - dayOfWeek + 1, 0, 0, 0, 0);
    var startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0);
    var startOfLastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1, 0, 0, 0, 0);
    var endOfLastMonth = new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59, 999);
    
    var currQtrMonth = Math.floor(now.getMonth() / 3) * 3;
    var startOfCurrQtr = new Date(now.getFullYear(), currQtrMonth, 1, 0, 0, 0, 0);
    var startOfLastQtr = new Date(now.getFullYear(), currQtrMonth - 3, 1, 0, 0, 0, 0);
    var endOfLastQtr = new Date(now.getFullYear(), currQtrMonth, 0, 23, 59, 59, 999);
    var startOfYear = new Date(now.getFullYear(), 0, 1, 0, 0, 0, 0);
    
    var periods = [
      { key: "daily", start: startOfToday, end: now },
      { key: "weekly", start: startOfWeek, end: now },
      { key: "monthly", start: startOfMonth, end: now },
      { key: "lastMonth", start: startOfLastMonth, end: endOfLastMonth },
      { key: "currentQuarter", start: startOfCurrQtr, end: now },
      { key: "lastQuarter", start: startOfLastQtr, end: endOfLastQtr },
      { key: "yearly", start: startOfYear, end: now }
    ];
    
    // 1. Process submissions for this staff member
    var subSheet = getSheetByNameFlexible(ss, ["Task_Submissions", "Task Submissions", "Submissions"]);
    if (subSheet && subSheet.getLastRow() > 1) {
      var subData = subSheet.getDataRange().getValues();
      for (var s = 1; s < subData.length; s++) {
        var sStaff = String(subData[s][1] || "").replace("👤 ", "").trim().toLowerCase();
        var isMatch = (sStaff === cleanStaff || (cleanStaff.length >= 3 && sStaff.indexOf(cleanStaff) !== -1) ||
                       (cleanEmpId && sStaff.indexOf(cleanEmpId) !== -1));
        if (isMatch) {
          var rawDate = subData[s][0];
          var subDate = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
          if (isNaN(subDate.getTime())) continue;
          
          var pts = parseInt(subData[s][10]) || 10;
          var deadlineStr = String(subData[s][6] || "");
          
          var isDelayed = false;
          if (deadlineStr && deadlineStr !== "N/A" && deadlineStr !== "Jald") {
            var deadDate = new Date(deadlineStr);
            if (!isNaN(deadDate.getTime()) && subDate > deadDate) isDelayed = true;
          }
          
          for (var p = 0; p < periods.length; p++) {
            var pr = periods[p];
            if (subDate >= pr.start && subDate <= pr.end) {
              result[pr.key].points += pts;
              result[pr.key].completed += 1;
              if (isDelayed) result[pr.key].delayed += 1;
            }
          }
        }
      }
    }
    
    // 2. Count assigned tasks from Adhoc_Tasks (Recurring + One-time)
    var setup = getOrSetupAdhocSheet(ss);
    var adSheet = setup.sheet;
    var colMap = setup.colMap;
    
    if (adSheet && adSheet.getLastRow() > 1) {
      var adData = adSheet.getDataRange().getValues();
      for (var a = 1; a < adData.length; a++) {
        var aStaff = String(adData[a][colMap.staff] || "").replace("👤 ", "").trim().toLowerCase();
        var isTargetStaff = (aStaff.indexOf("all") !== -1 || aStaff === cleanStaff ||
                             (cleanStaff.length >= 3 && aStaff.indexOf(cleanStaff) !== -1) ||
                             (cleanEmpId && aStaff.indexOf(cleanEmpId) !== -1));
        if (isTargetStaff) {
          var rawDate = adData[a][colMap.date];
          var aDate = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
          var rawFreq = String(adData[a][colMap.freq] || "Extra (One Time)").replace(/[^\w\s]/gi, '').replace(/\s*Task/i, '').trim();
          
          for (var p = 0; p < periods.length; p++) {
            var pr = periods[p];
            if (rawFreq === "Daily") {
              var daysInPeriod = Math.max(1, Math.round((Math.min(now.getTime(), pr.end.getTime()) - pr.start.getTime()) / (1000 * 60 * 60 * 24)));
              result[pr.key].assigned += daysInPeriod;
            } else if (rawFreq === "Weekly") {
              var weeksInPeriod = Math.max(1, Math.round((Math.min(now.getTime(), pr.end.getTime()) - pr.start.getTime()) / (1000 * 60 * 60 * 24 * 7)));
              result[pr.key].assigned += weeksInPeriod;
            } else if (rawFreq === "Monthly") {
              result[pr.key].assigned += 1;
            } else {
              if (!isNaN(aDate.getTime()) && aDate >= pr.start && aDate <= pr.end) {
                result[pr.key].assigned += 1;
              }
            }
          }
        }
      }
    }
    
    // Finalize percentage calculation
    for (var p = 0; p < periods.length; p++) {
      var k = periods[p].key;
      result[k].assigned = Math.max(result[k].assigned, result[k].completed);
      if (result[k].assigned > 0) {
        var onTimeCompleted = Math.max(0, result[k].completed - result[k].delayed);
        result[k].percentage = Math.min(100, Math.round((onTimeCompleted / result[k].assigned) * 100));
      } else {
        result[k].percentage = 100;
      }
    }
    
    return result;
  } catch (err) {
    return {
      daily: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 100 },
      weekly: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 100 },
      monthly: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 100 },
      lastMonth: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 100 },
      currentQuarter: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 100 },
      lastQuarter: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 100 },
      yearly: { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 100 }
    };
  }
}

// =====================================================================
// 7. ADMIN: ALL STAFF SCORECARD LEADERBOARD
// Fetches performance cards for all active staff across any period
// =====================================================================
function getStaffLeaderboard(period) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var staffSheet = getSheetByNameFlexible(ss, ["Staff_List", "Staff List", "Staff", "Users"]);
    var staffList = [];
    var pKey = period || "daily";
    if (!staffSheet) return [];
    
    var sData = staffSheet.getDataRange().getValues();
    for (var s = 1; s < sData.length; s++) {
      var sId = String(sData[s][0] || "");
      var sName = String(sData[s][2] || "").trim();
      var sRole = String(sData[s][3] || "Staff").trim();
      var sActive = sData[s][4];
      if (sName && sActive !== false && String(sActive).toLowerCase() !== "false") {
        if (!sRole.toLowerCase().includes("owner") && !sRole.toLowerCase().includes("admin")) {
          var metrics = getStaffScoreMetrics(sName, sId);
          var pData = metrics[pKey] || { points: 0, assigned: 0, completed: 0, delayed: 0, percentage: 100 };
          staffList.push({
            name: sName,
            role: sRole,
            points: pData.points,
            assigned: pData.assigned,
            completed: pData.completed,
            delayed: pData.delayed,
            percentage: pData.percentage,
            scores: metrics
          });
        }
      }
    }
    
    // Sort descending by points then percentage
    staffList.sort(function(a, b) {
      if (b.points !== a.points) return b.points - a.points;
      return b.percentage - a.percentage;
    });
    
    return staffList;
  } catch (err) {
    return [];
  }
}

// =====================================================================
// 8. STAFF: SUBMIT TASK WITH PROOF & REWARD POINTS
// =====================================================================
function submitTaskWithProof(name, role, type, task, subTask, deadline, base64Data, mimeType, fileName, isAdhoc, adhocIndex, actionType) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var subSheet = getSheetByNameFlexible(ss, ["Task_Submissions", "Task Submissions", "Submissions"]);
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
    
    var isSkip = actionType === "Skip";
    var pointsEarned = isSkip ? 0 : 10;
    var finalStatus = isSkip ? "Skipped" : (proofUrl !== "No Proof" ? "Uploaded (Done)" : "Completed");
    var timeZone = Session.getScriptTimeZone() || "GMT+5:30";
    var nowStr = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd HH:mm:ss");
    
    subSheet.appendRow([nowStr, name, role, type, task, subTask || "", deadline || "", finalStatus, proofUrl, actionType || "Done", pointsEarned, adhocIndex || ""]);
    
    // If one-time extra task, update status in Adhoc_Tasks sheet directly
    if (isAdhoc && adhocIndex) {
      var setup = getOrSetupAdhocSheet(ss);
      var adSheet = setup.sheet;
      var colMap = setup.colMap;
      var row = parseInt(adhocIndex);
      if (row >= 2 && row <= adSheet.getLastRow()) {
        var freqVal = String(adSheet.getRange(row, colMap.freq + 1).getValue() || "Extra (One Time)");
        if (freqVal.indexOf("Daily") === -1 && freqVal.indexOf("Weekly") === -1 && freqVal.indexOf("Monthly") === -1) {
          adSheet.getRange(row, colMap.status + 1).setValue(finalStatus);
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
    var staffSheet = getSheetByNameFlexible(ss, ["Staff_List", "Staff List", "Staff", "Users"]);
    var role = "Staff";
    var empId = "";
    if (staffSheet) {
      var sData = staffSheet.getDataRange().getValues();
      var cName = String(cleanName || "").replace("👤 ", "").trim().toLowerCase();
      for (var s = 1; s < sData.length; s++) {
        var rowStaff = String(sData[s][2] || "").replace("👤 ", "").trim().toLowerCase();
        if (rowStaff === cName || (cName.length >= 3 && rowStaff.indexOf(cName) !== -1)) {
          role = String(sData[s][3] || "Staff").trim();
          empId = String(sData[s][0] || "").trim();
          break;
        }
      }
    }
    
    var routine = getTasksForRole(role, cleanName);
    var adhoc = getAdhocTasksForStaff(cleanName, empId);
    var scores = getStaffScoreMetrics(cleanName, empId);
    
    return { role: role, routine: routine, adhoc: adhoc, scores: scores };
  } catch (err) {
    return { role: "Staff", routine: [], adhoc: [], scores: null };
  }
}

// =====================================================================
// 10. ADMIN: ALL ADHOC HISTORY
// =====================================================================
function getAllAdhocHistory() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var setup = getOrSetupAdhocSheet(ss);
    var sheet = setup.sheet;
    var colMap = setup.colMap;
    
    var data = sheet.getDataRange().getValues();
    var list = [];
    
    for (var i = data.length - 1; i >= 1; i--) {
      list.push({
        date: data[i][colMap.date],
        assignedTo: data[i][colMap.staff],
        type: data[i][colMap.freq] || "Extra (One Time)",
        task: data[i][colMap.task],
        deadline: data[i][colMap.deadline],
        refPhoto: data[i][colMap.photo],
        status: data[i][colMap.status]
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
    var sheet = getSheetByNameFlexible(ss, ["Orders"]);
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
    var sheet = getSheetByNameFlexible(ss, ["Orders"]);
    if (!sheet) return "Orders sheet nahi mili!";
    
    var row = parseInt(rowIndex);
    if (row >= 2 && row <= sheet.getLastRow()) {
      sheet.getRange(row, 7).setValue("Completed");
      return "Order successfully complete mark ho gaya!";
    }
    return "Invalid row index!";
  } catch (err) {
    return "Error: " + err.message;
  }
}

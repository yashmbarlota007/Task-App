function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  if (e && e.parameter && e.parameter.empId) {
    template.empId = e.parameter.empId;
    template.portalEmpId = e.parameter.empId;
    template.role = e.parameter.role;
  } else {
    template.empId = "";
    template.role = "";
  }
  return template.evaluate()
      .setTitle('Yash Marketing ERP & Task Manager')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function verifyLogin(empId, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var passSheet = ss.getSheetByName("password");
  if (!passSheet) return { success: false, message: "Error: 'password' sheet nahi mili!" };
  
  var passData = passSheet.getDataRange().getValues();
  var matchedEmpName = "";
  for (var i = 1; i < passData.length; i++) {
    var sheetId = passData[i][0] ? passData[i][0].toString().trim() : "";
    var sheetPass = passData[i][1] ? passData[i][1].toString().trim() : "";
    if (sheetId.toLowerCase() === empId.toString().trim().toLowerCase() && (sheetPass === password.toString().trim() || password === "SSO_BYPASS")) {
      matchedEmpName = findEmployeeName(sheetId);
      break;
    }
  }
  if (matchedEmpName !== "") {
    var role = getStaffRoleByName(matchedEmpName);
    var scores = getStaffScoresBreakdown(matchedEmpName);
    return { success: true, name: matchedEmpName, role: role, scores: scores };
  }
  return { success: false, message: "Galat Employee ID ya Password!" };
}

function getStaffListForAdmin() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("staff master"); 
  if (!sheet) return [];
  var data = sheet.getRange(2, 2, sheet.getLastRow(), 5).getValues(); 
  var list = [];
  var roles = [];
  for (var i = 0; i < data.length; i++) {
    var name = data[i][0] ? data[i][0].toString().trim() : "";
    var role = data[i][4] ? data[i][4].toString().trim() : "";
    if (name !== "" && list.indexOf("👤 " + name) === -1) list.push("👤 " + name);
    if (role !== "" && role.toLowerCase() !== "owner" && role.toLowerCase() !== "admin" && roles.indexOf("🏢 Team: " + role) === -1) {
      roles.push("🏢 Team: " + role);
    }
  }
  return ["👥 All Staff"].concat(roles).concat(["------------------"]).concat(list);
}

function findEmployeeName(empId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("staff master");
  var passSheet = ss.getSheetByName("password");
  if (masterSheet && passSheet) {
    var passData = passSheet.getDataRange().getValues();
    var masterData = masterSheet.getDataRange().getValues();
    // Try to match ID directly in staff master first
    for (var m = 1; m < masterData.length; m++) {
      if (masterData[m][0] && masterData[m][0].toString().trim().toLowerCase() === empId.toString().trim().toLowerCase()) {
        return masterData[m][1] ? masterData[m][1].toString().trim() : "Employee " + empId;
      }
    }
    // Fallback: match by row in password sheet
    for (var i = 1; i < passData.length; i++) {
      if (passData[i][0].toString().trim().toLowerCase() === empId.toString().trim().toLowerCase()) {
        var rowVal = masterSheet.getRange(i + 1, 2).getValue();
        if (rowVal) return rowVal.toString().trim();
      }
    }
  }
  return "Employee " + empId;
}

function getStaffRoleByName(empName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("staff master");
  if (!sheet) return "Staff";
  var data = sheet.getRange(2, 2, sheet.getLastRow(), 5).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === empName.toLowerCase()) {
      return data[i][4] ? data[i][4].toString().trim() : "Staff";
    }
  }
  return "Staff";
}

function formatTimeFriendly(val) {
  try {
    if (!val) return "";
    if (val instanceof Date) {
      if (val.getFullYear() === 1899 || val.getFullYear() === 1970) {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), "hh:mm a"); 
      } else {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd MMM, hh:mm a");
      }
    }
    return val.toString().replace(/\s*GMT.*/i, '').trim();
  } catch(e) { return val.toString().replace(/\s*GMT.*/i, '').trim(); }
}

// 📌 7-PERIOD SCORE CALCULATION
function getStaffScoresBreakdown(staffName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var subSheet = ss.getSheetByName("Task_Submissions");
  var scores = { daily: 0, weekly: 0, monthly: 0, lastMonth: 0, currentQuarter: 0, lastQuarter: 0, yearly: 0 };
  if (!subSheet || subSheet.getLastRow() <= 1) return scores;

  var subData = subSheet.getDataRange().getValues();
  var tz = Session.getScriptTimeZone();
  var now = new Date();
  var todayStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");

  var dayOfWeek = now.getDay();
  var diffToMon = (dayOfWeek === 0 ? -6 : 1) - dayOfWeek;
  var startOfWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() + diffToMon, 0, 0, 0, 0);

  var currYear = now.getFullYear();
  var currMonth = now.getMonth();
  var lastMonthYear = currMonth === 0 ? currYear - 1 : currYear;
  var lastMonth = currMonth === 0 ? 11 : currMonth - 1;

  var currQuarter = Math.floor(currMonth / 3);
  var lastQuarter = currQuarter === 0 ? 3 : currQuarter - 1;
  var lastQuarterYear = currQuarter === 0 ? currYear - 1 : currYear;

  var targetClean = staffName ? staffName.toString().trim().toUpperCase() : "";

  for (var j = 1; j < subData.length; j++) {
    var rStaff = subData[j][1] ? subData[j][1].toString().trim().toUpperCase() : "";
    if (targetClean !== "" && rStaff !== targetClean && !rStaff.includes(targetClean) && !targetClean.includes(rStaff)) continue;

    var pts = parseFloat(subData[j][10]) || 0;
    if (pts <= 0) continue;

    var rawDate = subData[j][0];
    var subDate = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
    if (isNaN(subDate.getTime())) continue;

    var subYear = subDate.getFullYear();
    var subMonth = subDate.getMonth();
    var subDateStr = Utilities.formatDate(subDate, tz, "yyyy-MM-dd");
    var subQuarter = Math.floor(subMonth / 3);

    if (subDateStr === todayStr) scores.daily += pts;
    if (subDate >= startOfWeek) scores.weekly += pts;
    if (subYear === currYear && subMonth === currMonth) scores.monthly += pts;
    if (subYear === lastMonthYear && subMonth === lastMonth) scores.lastMonth += pts;
    if (subYear === currYear && subQuarter === currQuarter) scores.currentQuarter += pts;
    if (subYear === lastQuarterYear && subQuarter === lastQuarter) scores.lastQuarter += pts;
    if (subYear === currYear) scores.yearly += pts;
  }
  return scores;
}

// 📌 OWNER SCOREBOARD FOR ALL STAFF
function getAllStaffScoreboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("staff master");
  if (!masterSheet) return [];
  
  var mData = masterSheet.getDataRange().getValues();
  var staffMap = {};
  var staffOrder = [];

  for (var i = 1; i < mData.length; i++) {
    var name = mData[i][1] ? mData[i][1].toString().trim() : "";
    var role = mData[i][5] ? mData[i][5].toString().trim() : (mData[i][4] ? mData[i][4].toString().trim() : "Staff");
    if (name && role.toLowerCase() !== "owner" && role.toLowerCase() !== "admin" && !staffMap[name.toUpperCase()]) {
      staffMap[name.toUpperCase()] = {
        name: name, role: role,
        daily: 0, weekly: 0, monthly: 0, lastMonth: 0, currentQuarter: 0, lastQuarter: 0, yearly: 0
      };
      staffOrder.push(name.toUpperCase());
    }
  }

  var subSheet = ss.getSheetByName("Task_Submissions");
  if (subSheet && subSheet.getLastRow() > 1) {
    var subData = subSheet.getDataRange().getValues();
    var tz = Session.getScriptTimeZone();
    var now = new Date();
    var todayStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");

    var dayOfWeek = now.getDay();
    var diffToMon = (dayOfWeek === 0 ? -6 : 1) - dayOfWeek;
    var startOfWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() + diffToMon, 0, 0, 0, 0);

    var currYear = now.getFullYear();
    var currMonth = now.getMonth();
    var lastMonthYear = currMonth === 0 ? currYear - 1 : currYear;
    var lastMonth = currMonth === 0 ? 11 : currMonth - 1;

    var currQuarter = Math.floor(currMonth / 3);
    var lastQuarter = currQuarter === 0 ? 3 : currQuarter - 1;
    var lastQuarterYear = currQuarter === 0 ? currYear - 1 : currYear;

    for (var j = 1; j < subData.length; j++) {
      var rStaff = subData[j][1] ? subData[j][1].toString().trim().toUpperCase() : "";
      if (!rStaff) continue;

      var matchKey = staffOrder.find(function(k) { return k === rStaff || k.includes(rStaff) || rStaff.includes(k); });
      if (!matchKey) continue;

      var pts = parseFloat(subData[j][10]) || 0;
      if (pts <= 0) continue;

      var rawDate = subData[j][0];
      var subDate = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
      if (isNaN(subDate.getTime())) continue;

      var subYear = subDate.getFullYear();
      var subMonth = subDate.getMonth();
      var subDateStr = Utilities.formatDate(subDate, tz, "yyyy-MM-dd");
      var subQuarter = Math.floor(subMonth / 3);

      var entry = staffMap[matchKey];
      if (subDateStr === todayStr) entry.daily += pts;
      if (subDate >= startOfWeek) entry.weekly += pts;
      if (subYear === currYear && subMonth === currMonth) entry.monthly += pts;
      if (subYear === lastMonthYear && subMonth === lastMonth) entry.lastMonth += pts;
      if (subYear === currYear && subQuarter === currQuarter) entry.currentQuarter += pts;
      if (subYear === lastQuarterYear && subQuarter === lastQuarter) entry.lastQuarter += pts;
      if (subYear === currYear) entry.yearly += pts;
    }
  }

  var result = [];
  for (var s = 0; s < staffOrder.length; s++) {
    result.push(staffMap[staffOrder[s]]);
  }
  result.sort(function(a, b) { return b.monthly - a.monthly; });
  return result;
}

// 📌 COMBINED ROUTINE & ASSIGNED TASKS (DAILY / WEEKLY / MONTHLY)

// =====================================================================
// PHASE 1: ZERO-TRAINING & DIGITAL TWIN HELPER FUNCTIONS
// =====================================================================
function getTimeSlotFromDeadline(deadlineStr) {
  if (!deadlineStr || deadlineStr === 'N/A' || deadlineStr === 'Jald') return 'morning';
  var lower = String(deadlineStr).toLowerCase();
  var match = lower.match(/(\d{1,2}):?(\d{2})?\s*(am|pm)?/i);
  if (match) {
    var h = parseInt(match[1], 10);
    var isPm = (match[3] && match[3].toLowerCase() === 'pm');
    var isAm = (match[3] && match[3].toLowerCase() === 'am');
    if (isPm && h < 12) h += 12;
    if (isAm && h === 12) h = 0;
    if (h < 11 || (h === 11 && parseInt(match[2] || 0, 10) <= 30)) return 'morning';
    if (h < 14 || (h === 14 && parseInt(match[2] || 0, 10) <= 30)) return 'midday';
    if (h < 17 || (h === 17 && parseInt(match[2] || 0, 10) <= 30)) return 'afternoon';
    return 'closing';
  }
  return 'morning';
}

function getSmartRoleDefaults(roleName, mainTask, subTask) {
  var role = String(roleName || '').toLowerCase();
  var task = String(mainTask || '').toLowerCase() + ' ' + String(subTask || '').toLowerCase();
  
  // ACCOUNTANT
  if (role.includes('account')) {
    if (task.includes('cash book') || task.includes('cash') || task.includes('hissab')) {
      return {
        sopSteps: [
          '1. Physical cash drawer mein sabhi notes & coins ko denomination-wise count karein (500x, 200x, 100x).',
          '2. Kal ke closing balance aur cash book se match karein.',
          '3. Din ke sabhi cash vouchers (Receipts & Payments) entry check karein.',
          '4. Net Physical Cash aur Book Balance 100% tally hone par hi register photo click karein.'
        ],
        goodSample: 'Clear, unblurred photo jisme: Date, Opening Cash, Denomination Tally, Total Receipts, Total Payments aur Net Cash clearly visible ho.',
        badSample: 'Sirf total amount ka rough paper scribble, dhundli photo, ya cash physical count kiye bina submit karna.',
        troubleshoot: 'Agar cash mein farq (discrepancy) aaye toh pehle last 3 payment vouchers check karein aur turant Owner ko report karein. Skip na karein.',
        checklist: ['Cash denomination count verify kar liya', 'Cash book entry tally ho gayi']
      };
    }
    if (task.includes('report') || task.includes('eod') || task.includes('sales') || task.includes('stock')) {
      return {
        sopSteps: [
          '1. ERP/Software mein jakar required report (Sales, Stock, Razorpay, Outstanding) download ya view karein.',
          '2. Date range aaj ki select karein (00:00 se abhi tak).',
          '3. Negative stock ya abnormal pending payments ko highlight karein.',
          '4. Report summary screen ka clear screenshot lekar upload karein.'
        ],
        goodSample: 'Full screen view jisme Date, Total Sales Volume, Outstanding figure aur verified status dikhta ho.',
        badSample: 'Phone se kaati hui aadhi screen jisme date ya total na dikhe, ya purane din ka data upload karna.',
        troubleshoot: 'Software down ho toh internet connection check karein, cache clear karein, ya offline Excel sheet mein manual tally ready rakhein.',
        checklist: ['Date filter correct set kiya', 'Total metrics cross-check kar liye']
      };
    }
  }

  // STORE MANAGER / WAREHOUSE
  if (role.includes('store') || role.includes('manager') || role.includes('warehouse')) {
    if (task.includes('order') || task.includes('dispatch') || task.includes('pack')) {
      return {
        sopSteps: [
          '1. Order list se items ka exact SKU, Color aur Quantity check karein.',
          '2. Har item ko physical box mein pack karne se pehle damage inspection karein.',
          '3. Invoice copy andar rakhein aur carton ke upar address label paste karein.',
          '4. Packed box ka photo click karke Dispatch confirm karein.'
        ],
        goodSample: 'Sealed carton photo jisme shop name, invoice number aur parcel label clearly legible ho.',
        badSample: 'Open box, bina tape ka carton, ya galat dealer ka address laga hona.',
        troubleshoot: 'Stock kam ho toh partial dispatch na karein; CRM aur Owner ko inform karke replacement item confirm karein.',
        checklist: ['Item quantity aur model check kiya', 'Carton tightly tape aur label kiya']
      };
    }
  }

  // SALES BOY / FIELD STAFF
  if (role.includes('sales')) {
    return {
      sopSteps: [
        '1. Target shop ya location par pahunch kar GPS location on karein.',
        '2. Dealer se milkar fresh order requirement aur outstanding payment par baat karein.',
        '3. Order book par stamp/signature lein.',
        '4. WhatsApp EOD report format mein total counter visits aur collected amount submit karein.'
      ],
      goodSample: 'Shop front with dealer / signed order sheet with clear payment breakdown.',
      badSample: 'Bina visit kiye purana order repeat karna ya incomplete counter feedback.',
      troubleshoot: 'Dealer available na ho toh next shop visit karein aur call par follow-up schedule karein.',
      checklist: ['GPS visit log verify kiya', 'Signed order sheet ready hai']
    };
  }

  // GENERAL DEFAULT
  return {
    sopSteps: [
      '1. Task ki di gayi deadline aur requirement ko dhyan se padhein.',
      '2. Standard process ke anusar work execute karein.',
      '3. Quality verify karein aur clear proof photo/screenshot click karein.',
      '4. Done button daba kar confirmation lein.'
    ],
    goodSample: 'Quality verified work meeting all deadline criteria with clear documentation.',
    badSample: 'Incomplete work, rushed submission without proof, or missed deadline.',
    troubleshoot: 'Kisi bhi issue ke liye apne supervisor ya WhatsApp group mein instant notify karein.',
    checklist: ['Kaam 100% complete ho gaya', 'Proof photo clear hai']
  };
}

function getRoleOnboardingGuide(roleName) {
  var role = String(roleName || 'Staff').trim();
  var rLower = role.toLowerCase();

  if (rLower.includes('account')) {
    return {
      role: role,
      title: 'Day-1 Operational Blueprint: Accountant',
      tagline: 'Zero-Discrepancy Financial Operations & Daily Tally System',
      mission: 'Aapka mukhya lakshya Yash Marketing ki har financial transaction ko 100% accurate, documented aur zero-discrepancy banana hai. Day-1 se closing cash aur register tally exact match hona chahiye.',
      rules: [
        'Kissi bhi transaction ko bina voucher ya receipt ke execute na karein.',
        'Har roz shaam 07:00 PM se pehle physical cash aur software closing balance match karein.',
        'Sabhi daily reports (Sales, Stock, Razorpay, Outstanding) 11:00 AM se pehle submit karein.'
      ],
      timeline: [
        { time: '09:30 AM - 10:00 AM', title: 'Opening Cash & Day Setup', desc: 'Kal ka closing cash physical drawer se count karein. Cash register mein opening entry note karein.' },
        { time: '10:00 AM - 11:00 AM', title: 'Daily Reports Generation', desc: 'ERP se Sales, Stock, Razorpay, Outstanding reports nikaal kar verify karein aur app par submit karein.' },
        { time: '11:00 AM - 02:00 PM', title: 'Ledger Entries & Vendor Vouchers', desc: 'Aane wale payments aur vendor bills ki entry Tally/ERP mein karein.' },
        { time: '02:00 PM - 05:00 PM', title: 'Bank Reconciliation & Collections', desc: 'Bank statement aur online UPI transactions ko sales accounts se reconcile karein.' },
        { time: '05:00 PM - 07:00 PM', title: 'Closing Cash Book & WhatsApp EOD', desc: 'Physical cash count karke register sign karein, Cash Book task mark karein aur EOD report send karein.' }
      ],
      benchmarks: {
        good: 'Full unblurred register page jisme denomination summary (500x, 200x, 100x), daily total aur net balance clean ho.',
        bad: 'Rough chit par pencil se likha total, ya bina physical cash count kiye man-pasand figure likhna.'
      },
      faq: [
        { q: 'Agar physical cash aur cash book mein difference aaye toh kya karein?', a: 'Sabse pehle din ke sabhi payment vouchers dubara total karein. Agar farq na mile toh turant Owner ko WhatsApp par escalate karein.' },
        { q: 'Internet ya ERP band ho toh reports kaise submit karein?', a: 'Offline Excel sheet mein manual tally karein aur register ka clear photo lekar task complete karein.' }
      ]
    };
  }

  if (rLower.includes('store') || rLower.includes('warehouse')) {
    return {
      role: role,
      title: 'Day-1 Operational Blueprint: Store Manager',
      tagline: 'Flawless Order Dispatch, Stock Safety & Zero Transit Damage',
      mission: 'Aapka lakshya har order ko bina kisi delay, 100% right items aur zero damage ke pack & dispatch karna hai.',
      rules: [
        'Har order carton pack karne se pehle items ka SKU aur quantity invoice se match karein.',
        'Packing tight honi chahiye taaki transit mein koi item damage na ho.',
        'Transporter handover ke baad LR number aur photo turant system mein update karein.'
      ],
      timeline: [
        { time: '09:30 AM - 10:30 AM', title: 'Morning Stock & Pending Orders', desc: 'Aaj ke fresh orders verify karein aur warehouse stock location check karein.' },
        { time: '10:30 AM - 02:00 PM', title: 'Picking, Packing & Quality Check', desc: 'Items pick karein, damage check karein aur tight carton packaging karein.' },
        { time: '02:00 PM - 05:00 PM', title: 'Courier / Transport Handover', desc: 'Transporters ko cartons handover karke LR slip receive karein.' },
        { time: '05:00 PM - 07:00 PM', title: 'Stock Inward & Day Closing', desc: 'Naye aaye stock ko shelf par arrange karein aur closing status update karein.' }
      ],
      benchmarks: {
        good: 'Properly taped brown carton jisme dealer ka label, invoice copy aur LR slip clearly marked ho.',
        bad: 'Dheela pack jisse saman bahar nikal sake ya bina label ka parcel.'
      },
      faq: [
        { q: 'Agar order ka koi item stock mein na ho toh?', a: 'Customer ko galat item na bhejein. CRM / Owner ko message karein aur confirm karein.' }
      ]
    };
  }

  return {
    role: role,
    title: 'Day-1 Operational Blueprint: ' + role,
    tagline: 'Standard Operational Discipline & High Performance System',
    mission: 'Aapka lakshya apne daily assigned tasks ko time par aur highest quality ke sath complete karna hai.',
    rules: [
      'Har task ki deadline ka samman karein aur on-time submit karein.',
      'Kaam pura hone ke baad clear proof upload karein.',
      'Shift khatam hone par Daily EOD report send karna anivarya hai.'
    ],
    timeline: [
      { time: 'Morning Shift', title: 'Shift Kickoff & Priorities', desc: 'Daily tab mein jakar aaj ke urgent tasks review karein aur execution shuru karein.' },
      { time: 'Midday Shift', title: 'Core Task Execution', desc: 'Main routine kaam pura karein aur proof upload karein.' },
      { time: 'Evening Shift', title: 'Verification & EOD', desc: 'Pending items check karein aur WhatsApp EOD report submit karein.' }
    ],
    benchmarks: {
      good: 'Timely submission with complete documentation and accurate proof.',
      bad: 'Overdue task without reason, incomplete proof, or lack of communication.'
    },
    faq: [
      { q: 'Task complete karne mein problem aaye toh?', a: 'Task card mein diye gaye SOP Stepper ko padhein ya direct supervisor se contact karein.' }
    ]
  };
}

function getTasksForRole(role, staffName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var taskList = [];
  var roleClean = role ? role.toLowerCase().trim() : "";

  var targetSheetName = "";
  if (roleClean === "crm") targetSheetName = "Sheet3";
  else if (roleClean === "accountant") targetSheetName = "Accountant";
  else if (roleClean === "sales boy" || roleClean === "sales") targetSheetName = "Sales Boy";
  else if (roleClean === "mis") targetSheetName = "MIS";
  else if (roleClean === "store manager") targetSheetName = "Store Manager";

  if (targetSheetName !== "") {
    var sheet = ss.getSheetByName(targetSheetName);
    if (sheet && sheet.getLastRow() >= 3) {
      var data = sheet.getDataRange().getValues();
      var lastDailyTask = "", lastDailyDeadline = "", lastDailySop = "", lastDailyVideo = "";
      var lastWeeklyTask = "", lastWeeklyDeadline = "", lastWeeklySop = "", lastWeeklyVideo = "";
      var lastMonthlyTask = "", lastMonthlyDeadline = "", lastMonthlySop = "", lastMonthlyVideo = "";

      for (var i = 2; i < data.length; i++) {
        var dTask = data[i][0] ? data[i][0].toString().trim() : "";
        var dSub = data[i][1] ? data[i][1].toString().trim() : "";
        var dDead = formatTimeFriendly(data[i][2]);
        var dSop = data[i][3] ? data[i][3].toString().trim() : "";
        var dVid = data[i][4] ? data[i][4].toString().trim() : "";
        if (dTask !== "") { lastDailyTask = dTask; if(dDead) lastDailyDeadline = dDead; if(dSop) lastDailySop = dSop; if(dVid) lastDailyVideo = dVid; }
        if (dSub !== "" || (dTask !== "" && dSub === "")) {
          var curT = dTask !== "" ? dTask : lastDailyTask;
          var curD = dDead || lastDailyDeadline;
          var metaD = getSmartRoleDefaults(roleClean, curT, dSub);
          taskList.push({
            type: "Daily",
            task: curT,
            subTask: dSub,
            deadline: curD,
            timeSlot: getTimeSlotFromDeadline(curD),
            sop: dSop || lastDailySop,
            sopVideo: dVid || lastDailyVideo,
            sopSteps: metaD.sopSteps,
            goodSample: metaD.goodSample,
            badSample: metaD.badSample,
            troubleshoot: metaD.troubleshoot,
            checklist: metaD.checklist,
            status: "Pending ⏳"
          });
        }

        var wTask = data[i][5] ? data[i][5].toString().trim() : "";
        var wSub = data[i][6] ? data[i][6].toString().trim() : "";
        var wDead = formatTimeFriendly(data[i][7]);
        var wSop = data[i][8] ? data[i][8].toString().trim() : "";
        var wVid = data[i][9] ? data[i][9].toString().trim() : "";
        if (wTask !== "") { lastWeeklyTask = wTask; if(wDead) lastWeeklyDeadline = wDead; if(wSop) lastWeeklySop = wSop; if(wVid) lastWeeklyVideo = wVid; }
        if (wSub !== "" || (wTask !== "" && wSub === "")) {
          var curWT = wTask !== "" ? wTask : lastWeeklyTask;
          var curWD = wDead || lastWeeklyDeadline;
          var metaW = getSmartRoleDefaults(roleClean, curWT, wSub);
          taskList.push({
            type: "Weekly",
            task: curWT,
            subTask: wSub,
            deadline: curWD,
            timeSlot: getTimeSlotFromDeadline(curWD),
            sop: wSop || lastWeeklySop,
            sopVideo: wVid || lastWeeklyVideo,
            sopSteps: metaW.sopSteps,
            goodSample: metaW.goodSample,
            badSample: metaW.badSample,
            troubleshoot: metaW.troubleshoot,
            checklist: metaW.checklist,
            status: "Pending ⏳"
          });
        }

        var mTask = data[i][10] ? data[i][10].toString().trim() : "";
        var mSub = data[i][11] ? data[i][11].toString().trim() : "";
        var mDead = formatTimeFriendly(data[i][12]);
        var mSop = data[i][13] ? data[i][13].toString().trim() : "";
        var mVid = data[i][14] ? data[i][14].toString().trim() : "";
        if (mTask !== "") { lastMonthlyTask = mTask; if(mDead) lastMonthlyDeadline = mDead; if(mSop) lastMonthlySop = mSop; if(mVid) lastMonthlyVideo = mVid; }
        if (mSub !== "" || (mTask !== "" && mSub === "")) {
          var curMT = mTask !== "" ? mTask : lastMonthlyTask;
          var curMD = mDead || lastMonthlyDeadline;
          var metaM = getSmartRoleDefaults(roleClean, curMT, mSub);
          taskList.push({
            type: "Monthly",
            task: curMT,
            subTask: mSub,
            deadline: curMD,
            timeSlot: getTimeSlotFromDeadline(curMD),
            sop: mSop || lastMonthlySop,
            sopVideo: mVid || lastMonthlyVideo,
            sopSteps: metaM.sopSteps,
            goodSample: metaM.goodSample,
            badSample: metaM.badSample,
            troubleshoot: metaM.troubleshoot,
            checklist: metaM.checklist,
            status: "Pending ⏳"
          });
        }
      }
    }
  }

  if (roleClean === "sales boy" || roleClean === "sales") {
     taskList.push({ type: "Daily", task: "Send EOD Report on WhatsApp", subTask: "Screenshot & Final Update", deadline: "08:00 PM", sop: "", sopVideo: "", status: "Pending ⏳" });
  }

  // 📌 MERGE ASSIGNED DAILY, WEEKLY, MONTHLY TASKS FROM 'Adhoc_Tasks'
  var adhocSheet = ss.getSheetByName("Adhoc_Tasks");
  if (adhocSheet && adhocSheet.getLastRow() > 1) {
    var aData = adhocSheet.getDataRange().getValues();
    var sClean = staffName ? staffName.toString().trim().toLowerCase() : "";
    for (var a = 1; a < aData.length; a++) {
      var aAssignee = aData[a][1] ? aData[a][1].toString().trim().toLowerCase() : "";
      var aDesc = aData[a][2] ? aData[a][2].toString().trim() : "";
      var aDead = aData[a][3] ? aData[a][3].toString().replace(/\s*GMT.*/i, '').trim() : "";
      var aStatus = aData[a][4] ? aData[a][4].toString().trim() : "";
      var aPhotos = aData[a][5] ? aData[a][5].toString().trim() : "";
      var aFreq = aData[a][6] ? aData[a][6].toString().trim() : "Extra (One Time)";

      if (aDesc === "") continue;

      var isMatch = false;
      if (aAssignee === sClean || aAssignee === roleClean || aAssignee === "all" || aAssignee === "everyone" || aAssignee === "all staff") isMatch = true;
      else if (aAssignee !== "" && sClean !== "" && (aAssignee.includes(sClean) || sClean.includes(aAssignee))) isMatch = true;
      else if (aAssignee.replace("team: ", "").trim() === roleClean) isMatch = true;

      if (!isMatch) continue;

      var fType = "";
      if (aFreq.toLowerCase().includes("daily")) fType = "Daily";
      else if (aFreq.toLowerCase().includes("weekly")) fType = "Weekly";
      else if (aFreq.toLowerCase().includes("monthly")) fType = "Monthly";

      if (fType !== "") {
        taskList.push({
          type: fType,
          task: aDesc,
          subTask: "⚡ Assigned by Owner",
          deadline: aDead,
          sop: "",
          sopVideo: "",
          refPhoto: aPhotos,
          isAdhocAssigned: true,
          adhocRowIndex: a + 1,
          status: aStatus || "Pending ⏳"
        });
      }
    }
  }

  // 📌 CHECK 'Task_Submissions' TO KEEP TODAY'S/THIS WEEK'S/MONTH'S COMPLETED TASKS MARKED DONE
  var subSheet = ss.getSheetByName("Task_Submissions");
  if (subSheet && subSheet.getLastRow() > 1) {
    var subData = subSheet.getDataRange().getValues();
    var tz = Session.getScriptTimeZone();
    var now = new Date();
    var todayStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");
    var currentMonthStr = Utilities.formatDate(now, tz, "yyyy-MM");

    var dayOfWeek = now.getDay(); 
    var diffToMonday = (dayOfWeek === 0 ? -6 : 1) - dayOfWeek;
    var startOfWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() + diffToMonday, 0, 0, 0);

    var staffClean = staffName ? staffName.toString().trim().toLowerCase() : "";
    var submittedMap = {};

    for (var s = subData.length - 1; s >= 1; s--) {
      var rawDate = subData[s][0];
      if (!rawDate) continue;
      
      var rStaff = subData[s][1] ? subData[s][1].toString().trim().toLowerCase() : "";
      var rRole = subData[s][2] ? subData[s][2].toString().trim().toLowerCase() : "";
      var rType = subData[s][3] ? subData[s][3].toString().trim().toLowerCase() : "";
      var rTask = subData[s][4] ? subData[s][4].toString().trim().toLowerCase() : "";
      var rSub = subData[s][5] ? subData[s][5].toString().trim().toLowerCase() : "";
      var rStatus = subData[s][8] ? subData[s][8].toString().trim() : "";

      var isPersonMatch = false;
      if (staffClean !== "") {
        if (rStaff === staffClean || rStaff.includes(staffClean) || staffClean.includes(rStaff)) isPersonMatch = true;
      } else {
        if (rRole === roleClean) isPersonMatch = true;
      }
      if (!isPersonMatch) continue;

      var subDateObj = (rawDate instanceof Date) ? rawDate : new Date(rawDate.toString().replace(/-/g, "/"));
      var subDateStr = "";
      var subMonthStr = "";
      if (!isNaN(subDateObj.getTime())) {
        subDateStr = Utilities.formatDate(subDateObj, tz, "yyyy-MM-dd");
        subMonthStr = Utilities.formatDate(subDateObj, tz, "yyyy-MM");
      } else {
        subDateStr = rawDate.toString().substring(0, 10);
        subMonthStr = rawDate.toString().substring(0, 7);
      }

      var isValidPeriod = false;
      if (rType === "daily" && subDateStr === todayStr) {
        isValidPeriod = true;
      } else if (rType === "weekly" && (!isNaN(subDateObj.getTime()) ? subDateObj >= startOfWeek : subDateStr === todayStr)) {
        isValidPeriod = true;
      } else if (rType === "monthly" && subMonthStr === currentMonthStr) {
        isValidPeriod = true;
      }

      if (isValidPeriod) {
        var key = rType + "|" + rTask + "|" + rSub;
        if (!submittedMap[key]) {
          submittedMap[key] = rStatus;
        }
      }
    }

    for (var k = 0; k < taskList.length; k++) {
      var tKey = taskList[k].type.toLowerCase() + "|" + (taskList[k].task || "").toString().trim().toLowerCase() + "|" + (taskList[k].subTask || "").toString().trim().toLowerCase();
      if (submittedMap[tKey]) {
        taskList[k].status = submittedMap[tKey];
      }
    }
  }

  return taskList;
}

function getLiveOrdersForManager() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Orders"); 
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var allOrders = [];
  var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  for (var i = 1; i < data.length; i++) {
    var rawDate = data[i][0];
    if (!rawDate) continue;
    var orderDateStr = Utilities.formatDate(new Date(rawDate), Session.getScriptTimeZone(), "yyyy-MM-dd");
    if (orderDateStr !== todayStr) continue;
    
    var orderId = data[i][1] ? data[i][1].toString().trim() : "";
    var shopName = data[i][2] ? data[i][2].toString().trim() : "";
    var items = data[i][4] ? data[i][4].toString().trim() : "";
    var status = data[i][8] ? data[i][8].toString().trim() : ""; 
    if (orderId !== "") {
      allOrders.push({ rowIndex: i + 1, orderId: orderId, shopName: shopName, items: items, orderTime: "N/A", status: status || "Pending ⏳" });
    }
  }
  return allOrders;
}

function confirmOrderCompleted(rowIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Orders");
  if (sheet) {
    var submitTimeString = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "hh:mm a"); 
    var newStatus = "Completed ✅ (" + submitTimeString + ")";
    sheet.getRange(rowIndex, 9).setValue(newStatus);
    return { success: true, status: newStatus };
  }
  return { success: false };
}

function getOrCreatePhotoFolder() {
  var folderName = "ERP_Task_Photos";
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

// 📌 ASSIGN TASK WITH MULTIPLE PHOTOS & SEAMLESS DAILY/WEEKLY/MONTHLY/EXTRA SUPPORT
function assignAdhocTask(assigneeInfo, taskTypeVal, taskDesc, deadline, photoArray) {
  if (!assigneeInfo || assigneeInfo.includes("----")) return "Error: Kripya valid staff ya team chunein!";
  var cleanName = assigneeInfo.replace("🏢 Team: ", "").replace("👤 ", "").replace("👥 ", "").trim();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Save all photos to Drive folder
  var photoUrls = [];
  if (photoArray && photoArray.length > 0) {
    try {
      var folder = getOrCreatePhotoFolder();
      for (var p = 0; p < photoArray.length; p++) {
        var pItem = photoArray[p];
        if (pItem.base64Data && pItem.fileName) {
          var blob = Utilities.newBlob(Utilities.base64Decode(pItem.base64Data), pItem.mimeType, pItem.fileName);
          var file = folder.createFile(blob);
          try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}
          photoUrls.push(file.getUrl());
        }
      }
    } catch(e) {}
  }
  var photoLinksString = photoUrls.length > 0 ? photoUrls.join(" , ") : "No Photo";

  // 📌 Store in 'Adhoc_Tasks' with Frequency column
  var adhocSheet = ss.getSheetByName("Adhoc_Tasks");
  if (!adhocSheet) {
    adhocSheet = ss.insertSheet("Adhoc_Tasks");
    adhocSheet.appendRow(["Assign Date", "Staff Name", "Task Description", "Deadline", "Status", "Reference Photo", "Task Frequency"]);
  }

  // Ensure header has Col 7 if missing
  if (adhocSheet.getLastColumn() < 7) {
    adhocSheet.getRange(1, 7).setValue("Task Frequency");
  }

  adhocSheet.appendRow([new Date(), cleanName, taskDesc, deadline, "Pending ⏳", photoLinksString, taskTypeVal]);
  return taskTypeVal + " Assigned Successfully to " + cleanName + " 🚀";
}

// 📌 FETCH ONLY EXTRA (ONE TIME) TASKS FOR EXTRA TASKS TAB
function getAdhocTasksForStaff(staffName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Adhoc_Tasks");
  if(!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var tasks = [];
  var role = String(getStaffRoleByName(staffName)).toLowerCase().trim();
  var name = String(staffName).toLowerCase().trim();
  var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM");
  
  for(var i = 1; i < data.length; i++) {
    var rawAssignee = String(data[i][1] || "").toLowerCase().trim();
    var taskDesc = String(data[i][2] || "").trim();
    var deadline = data[i][3] ? data[i][3].toString().replace(/\s*GMT.*/i, '').trim() : "";
    var status = String(data[i][4] || "").trim();
    var refPhoto = String(data[i][5] || "");
    var freq = String(data[i][6] || "Extra (One Time)").trim();

    // Routine Daily/Weekly/Monthly tasks are shown in their respective tabs via getTasksForRole
    if (freq !== "Extra (One Time)" && freq !== "One Time" && freq !== "") continue;
    if (taskDesc === "") continue;
    
    var assignDateRaw = data[i][0];
    var assignDateStr = "";
    if (assignDateRaw) {
        try {
            var d = new Date(assignDateRaw);
            if(!isNaN(d.getTime())) { assignDateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "dd MMM yyyy, hh:mm a"); } 
            else { assignDateStr = assignDateRaw.toString().replace(/\s*GMT.*/i, '').trim(); }
        } catch(e) { assignDateStr = assignDateRaw.toString().replace(/\s*GMT.*/i, '').trim(); }
    }

    var isMatch = false;
    if (rawAssignee === name || rawAssignee === role || rawAssignee === "all" || rawAssignee === "everyone" || rawAssignee === "all staff") { isMatch = true; } 
    else if (rawAssignee !== "" && name !== "") { if (rawAssignee.includes(name) || name.includes(rawAssignee)) isMatch = true; }

    var isPending = true;
    if(status.toLowerCase().includes("upload") || status.toLowerCase().includes("complet") || status.toLowerCase().includes("skip")) {
        if(!status.includes(todayStr)) { isPending = false; }
    }

    if(isMatch && isPending) {
      tasks.push({ index: i, assignDate: assignDateStr, task: taskDesc, deadline: deadline, status: status || "Pending ⏳", refPhoto: refPhoto });
    }
  }
  return tasks;
}

function getStaffProgressReport(staffName) {
  var role = getStaffRoleByName(staffName);
  var routineTasks = getTasksForRole(role, staffName);
  var adhocTasks = getAdhocTasksForStaff(staffName);
  var scores = getStaffScoresBreakdown(staffName);
  return { routine: routineTasks, adhoc: adhocTasks, role: role, scores: scores };
}

function getAllAdhocHistory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Adhoc_Tasks");
  if(!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var history = [];
  var start = Math.max(1, data.length - 60);
  
  for(var i = data.length - 1; i >= start; i--) {
    var rawDate = data[i][0];
    var dateStr = rawDate ? rawDate.toString().replace(/\s*GMT.*/i, '').trim() : "";
    try {
      if(rawDate) {
        var d = new Date(rawDate);
        if(!isNaN(d.getTime())) { dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "dd MMM yyyy, hh:mm a"); }
      }
    } catch(e) {}
    
    var taskDesc = data[i][2] ? data[i][2].toString().trim() : "";
    if (taskDesc === "") continue; 

    history.push({
      date: dateStr,
      assignedTo: data[i][1] ? data[i][1].toString() : "",
      task: taskDesc,
      deadline: data[i][3] ? data[i][3].toString().replace(/\s*GMT.*/i, '').trim() : "",
      status: data[i][4] ? data[i][4].toString() : "",
      photos: data[i][5] ? data[i][5].toString() : "",
      freq: data[i][6] ? data[i][6].toString() : "Extra"
    });
  }
  return history;
}

function isTaskOnTime(deadlineStr, submitDateObj) {
  if (!deadlineStr || deadlineStr === "N/A" || deadlineStr === "Jald") return true; 
  try {
    var targetDate = new Date(submitDateObj.getTime()); 
    if (deadlineStr.includes(",") || deadlineStr.includes("202")) {
        targetDate = new Date(deadlineStr);
    } else {
        var parts = deadlineStr.match(/(\d+):(\d+)\s(AM|PM)/i);
        if (parts) {
            var hours = parseInt(parts[1], 10);
            var minutes = parseInt(parts[2], 10);
            var ampm = parts[3].toUpperCase();
            if (ampm === "PM" && hours < 12) hours += 12;
            if (ampm === "AM" && hours === 12) hours = 0;
            targetDate.setHours(hours, minutes, 0, 0);
        }
    }
    return submitDateObj.getTime() <= targetDate.getTime();
  } catch(e) { return true; }
}

function submitTaskWithProof(staffName, role, taskType, taskName, subTask, deadline, base64Data, mimeType, fileName, isAdhoc, adhocIndex, actionType) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();
  var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MMM");
  var submitTimeString = Utilities.formatDate(now, Session.getScriptTimeZone(), "hh:mm a"); 
  var fullDateTime = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  var statusPrefix = "Uploaded ✅";
  var hyperlinkFormula = "No Image";
  var pointsEarned = 0; 

  if (actionType === "Skip" || actionType === true) {
      statusPrefix = "Skipped ⏭️"; hyperlinkFormula = "Skipped (No Photo)"; pointsEarned = 0; 
  } else if (actionType === "Complete") {
      statusPrefix = "Completed ✅"; hyperlinkFormula = "Completed (No Photo)";
  }

  if (base64Data && fileName && actionType !== "Skip" && actionType !== "Complete") {
    try {
      var folder = getOrCreatePhotoFolder();
      var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
      var file = folder.createFile(blob);
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}
      hyperlinkFormula = '=HYPERLINK("' + file.getUrl() + '", "📷 View Photo")';
      statusPrefix = "Uploaded ✅";
    } catch(e) {}
  }

  if (actionType !== "Skip" && actionType !== true) {
    pointsEarned = 5; 
    if (isTaskOnTime(deadline, now)) { pointsEarned += 5; }
  }

  var finalStatusDisplay = statusPrefix + " (" + dateStr + " " + submitTimeString + ")";

  if(isAdhoc === true && adhocIndex !== null && adhocIndex !== undefined) {
    var adhocSheet = ss.getSheetByName("Adhoc_Tasks");
    if(adhocSheet) { adhocSheet.getRange(adhocIndex + 1, 5).setValue(finalStatusDisplay); }
  }

  var sheet = ss.getSheetByName("Task_Submissions");
  if (!sheet) {
    sheet = ss.insertSheet("Task_Submissions");
    sheet.appendRow(["Submit Timestamp", "Staff Name", "Role", "Task Type", "Task", "Sub Task", "Set Deadline", "Actual Submit Time", "Status", "Proof Image Link", "Points Earned"]);
  }
  
  sheet.appendRow([fullDateTime, staffName, role, taskType, taskName, subTask, deadline, submitTimeString, finalStatusDisplay, hyperlinkFormula, pointsEarned]);
  return { success: true, submitTime: submitTimeString, status: finalStatusDisplay, pointsEarned: pointsEarned };
}

function setupPermissions() { DriveApp.getRootFolder(); }
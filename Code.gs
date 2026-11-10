function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  var empId = (e && e.parameter && e.parameter.empId) ? e.parameter.empId : "";
  var role = (e && e.parameter && e.parameter.role) ? e.parameter.role : "";
  template.empId = empId;
  template.portalEmpId = empId;
  template.role = role;
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
    // Auto-ensure staff sheet exists in background
    try {
      var staffSh = ss.getSheetByName(matchedEmpName);
      if (!staffSh) {
        setupStaffTaskSheet(matchedEmpName, role);
      }
    } catch(e) {}
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
    if (role !== "" && role.toLowerCase() !== "owner" && role.toLowerCase() !== "ceo" && role.toLowerCase() !== "admin" && roles.indexOf("🏢 Team: " + role) === -1) {
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
      var r = data[i][4] ? data[i][4].toString().trim() : "Staff";
      if (r.toLowerCase() === "owner") return "CEO";
      return r;
    }
  }
  return "Staff";
}

function formatTimeFriendly(val) {
  try {
    if (!val) return "";
    if (val instanceof Date) {
      if (val.getFullYear() <= 1970) {
        var h = val.getHours();
        var m = val.getMinutes();
        var ampm = h >= 12 ? "PM" : "AM";
        var h12 = h % 12;
        if (h12 === 0) h12 = 12;
        var hStr = (h12 < 10 ? "0" : "") + h12;
        var mStr = (m < 10 ? "0" : "") + m;
        return hStr + ":" + mStr + " " + ampm;
      } else {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd MMM, hh:mm a");
      }
    }
    var str = val.toString().replace(/\s*GMT.*/i, "").trim();
    var timeMatch = str.match(/(\d{1,2}):(\d{2})(?::\d{2})?\s*(am|pm)?/i);
    if (timeMatch && !timeMatch[3]) {
      var rawH = parseInt(timeMatch[1], 10);
      var rawM = parseInt(timeMatch[2], 10);
      if (!isNaN(rawH) && !isNaN(rawM) && rawH < 24 && rawM < 60) {
        var a = rawH >= 12 ? "PM" : "AM";
        var hh = rawH % 12;
        if (hh === 0) hh = 12;
        var hhStr = (hh < 10 ? "0" : "") + hh;
        var mmStr = (rawM < 10 ? "0" : "") + rawM;
        var convertedTime = hhStr + ":" + mmStr + " " + a;
        return str.replace(timeMatch[0], convertedTime);
      }
    }
    return str;
  } catch(e) { return (val || "").toString().replace(/\s*GMT.*/i, "").trim(); }
}

// 📌 7-PERIOD SCORE CALCULATION
function getStaffScoresBreakdown(staffName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var subSheet = ss.getSheetByName("Task_Submissions");
  var role = getStaffRoleByName(staffName);

  var rawPts = { daily: 0, weekly: 0, monthly: 0, lastMonth: 0, currentQuarter: 0, lastQuarter: 0, yearly: 0 };
  
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

  if (subSheet && subSheet.getLastRow() > 1) {
    var subData = subSheet.getDataRange().getValues();
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

      if (subDateStr === todayStr) rawPts.daily += pts;
      if (subDate >= startOfWeek) rawPts.weekly += pts;
      if (subYear === currYear && subMonth === currMonth) rawPts.monthly += pts;
      if (subYear === lastMonthYear && subMonth === lastMonth) rawPts.lastMonth += pts;
      if (subYear === currYear && subQuarter === currQuarter) rawPts.currentQuarter += pts;
      if (subYear === lastQuarterYear && subQuarter === lastQuarter) rawPts.lastQuarter += pts;
      if (subYear === currYear) rawPts.yearly += pts;
    }
  }

  // Count routine tasks for realistic target calculation
  var routineDaily = 0;
  var routineWeekly = 0;
  var routineMonthly = 0;
  try {
    var rTasks = getTasksForRole(role, staffName);
    routineDaily = rTasks.filter(function(t) { return t.type === 'Daily'; }).length;
    routineWeekly = rTasks.filter(function(t) { return t.type === 'Weekly'; }).length;
    routineMonthly = rTasks.filter(function(t) { return t.type === 'Monthly'; }).length;
  } catch(e) {}

  if (routineDaily === 0) routineDaily = Math.max(1, Math.ceil(rawPts.daily / 10));

  var dailyMax = Math.max(rawPts.daily, routineDaily * 10, 10);
  var dailyPct = Math.min(100, Math.round((rawPts.daily / dailyMax) * 100));

  var dayOfWeekNum = (dayOfWeek === 0 ? 7 : dayOfWeek);
  var weeklyExpected = (routineDaily * dayOfWeekNum) + routineWeekly;
  var weeklyMax = Math.max(rawPts.weekly, weeklyExpected * 10, 10);
  var weeklyPct = Math.min(100, Math.round((rawPts.weekly / weeklyMax) * 100));

  var dayOfMonth = now.getDate();
  var monthlyExpected = (routineDaily * dayOfMonth) + (routineWeekly * Math.ceil(dayOfMonth / 7)) + routineMonthly;
  var monthlyMax = Math.max(rawPts.monthly, monthlyExpected * 10, 10);
  var monthlyPct = Math.min(100, Math.round((rawPts.monthly / monthlyMax) * 100));

  var lastMoExpected = (routineDaily * 30) + (routineWeekly * 4) + routineMonthly;
  var lastMoMax = Math.max(rawPts.lastMonth, lastMoExpected * 10, 10);
  var lastMoPct = Math.min(100, Math.round((rawPts.lastMonth / lastMoMax) * 100));

  var qtrMax = Math.max(rawPts.currentQuarter, monthlyExpected * 3 * 10, 10);
  var qtrPct = Math.min(100, Math.round((rawPts.currentQuarter / qtrMax) * 100));

  var lastQtrMax = Math.max(rawPts.lastQuarter, lastMoExpected * 3 * 10, 10);
  var lastQtrPct = Math.min(100, Math.round((rawPts.lastQuarter / lastQtrMax) * 100));

  var yrMax = Math.max(rawPts.yearly, lastMoExpected * 12 * 10, 10);
  var yrPct = Math.min(100, Math.round((rawPts.yearly / yrMax) * 100));

  return {
    daily: { percentage: dailyPct, earned: rawPts.daily, max: dailyMax },
    weekly: { percentage: weeklyPct, earned: rawPts.weekly, max: weeklyMax },
    monthly: { percentage: monthlyPct, earned: rawPts.monthly, max: monthlyMax },
    lastMonth: { percentage: lastMoPct, earned: rawPts.lastMonth, max: lastMoMax },
    currentQuarter: { percentage: qtrPct, earned: rawPts.currentQuarter, max: qtrMax },
    lastQuarter: { percentage: lastQtrPct, earned: rawPts.lastQuarter, max: lastQtrMax },
    yearly: { percentage: yrPct, earned: rawPts.yearly, max: yrMax }
  };
}

// 📌 CEO SCOREBOARD FOR ALL STAFF
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
    if (name && role.toLowerCase() !== "owner" && role.toLowerCase() !== "ceo" && role.toLowerCase() !== "admin" && !staffMap[name.toUpperCase()]) {
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
    var raw = staffMap[staffOrder[s]];
    
    var dailyMax = Math.max(raw.daily, 70);
    var dailyPct = Math.min(100, Math.round((raw.daily / dailyMax) * 100));

    var weeklyMax = Math.max(raw.weekly, 300);
    var weeklyPct = Math.min(100, Math.round((raw.weekly / weeklyMax) * 100));

    var monthlyMax = Math.max(raw.monthly, 1000);
    var monthlyPct = Math.min(100, Math.round((raw.monthly / monthlyMax) * 100));

    var lastMoMax = Math.max(raw.lastMonth, 2100);
    var lastMoPct = Math.min(100, Math.round((raw.lastMonth / lastMoMax) * 100));

    var qtrMax = Math.max(raw.currentQuarter, 2100);
    var qtrPct = Math.min(100, Math.round((raw.currentQuarter / qtrMax) * 100));

    var lastQtrMax = Math.max(raw.lastQuarter, 6300);
    var lastQtrPct = Math.min(100, Math.round((raw.lastQuarter / lastQtrMax) * 100));

    var yrMax = Math.max(raw.yearly, 25200);
    var yrPct = Math.min(100, Math.round((raw.yearly / yrMax) * 100));

    result.push({
      name: raw.name,
      role: raw.role,
      rawMonthly: raw.monthly,
      daily: { percentage: dailyPct, earned: raw.daily, max: dailyMax },
      weekly: { percentage: weeklyPct, earned: raw.weekly, max: weeklyMax },
      monthly: { percentage: monthlyPct, earned: raw.monthly, max: monthlyMax },
      lastMonth: { percentage: lastMoPct, earned: raw.lastMonth, max: lastMoMax },
      currentQuarter: { percentage: qtrPct, earned: raw.currentQuarter, max: qtrMax },
      lastQuarter: { percentage: lastQtrPct, earned: raw.lastQuarter, max: lastQtrMax },
      yearly: { percentage: yrPct, earned: raw.yearly, max: yrMax }
    });
  }
  result.sort(function(a, b) { return b.rawMonthly - a.rawMonthly; });
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
        troubleshoot: 'Agar cash mein farq (discrepancy) aaye toh pehle last 3 payment vouchers check karein aur turant CEO ko report karein. Skip na karein.',
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
        troubleshoot: 'Stock kam ho toh partial dispatch na karein; CRM aur CEO ko inform karke replacement item confirm karein.',
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
        { q: 'Agar physical cash aur cash book mein difference aaye toh kya karein?', a: 'Sabse pehle din ke sabhi payment vouchers dubara total karein. Agar farq na mile toh turant CEO ko WhatsApp par escalate karein.' },
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
        { q: 'Agar order ka koi item stock mein na ho toh?', a: 'Customer ko galat item na bhejein. CRM / CEO ko message karein aur confirm karein.' }
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


// =====================================================================
// 📌 GET TASKS FOR STAFF & ROLE (PERSONAL STAFF SHEET OR ROLE SHEET)
// Supports both:
// 1. Modern 14-Column Table: Frequency, Task, SubTask, Deadline, Slot, Points, SOP, GoodSample, GoodPhoto, BadSample, BadPhoto, Troubleshoot, RefDoc, Video
// 2. Legacy 15-Column Layout (Daily cols 0-4, Weekly cols 5-9, Monthly cols 10-14)
// Priority: Sheet named after Staff Name (e.g. "Komal Gurap") -> Sheet named after Role (e.g. "Accountant")
// =====================================================================
function getTasksForRole(role, staffName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var taskList = [];
  var roleClean = role ? role.toLowerCase().trim() : "";
  var staffClean = staffName ? staffName.toString().trim() : "";

  // 1. FIND TARGET SHEET: Check Staff Name tab first, then Role tab
  var sheet = null;
  var sheetNameUsed = "";

  if (staffClean !== "") {
    sheet = ss.getSheetByName(staffClean);
    if (sheet) sheetNameUsed = staffClean;
    else {
      // Case-insensitive search across all sheets
      var allSheets = ss.getSheets();
      for (var s = 0; s < allSheets.length; s++) {
        var sName = allSheets[s].getName().trim();
        if (sName.toLowerCase() === staffClean.toLowerCase()) {
          sheet = allSheets[s];
          sheetNameUsed = sName;
          break;
        }
      }
    }
  }

  // 2. Fallback to Role tab if no personal staff tab
  if (!sheet && roleClean !== "") {
    var roleTarget = "";
    if (roleClean === "crm") roleTarget = "Sheet3";
    else if (roleClean === "accountant") roleTarget = "Accountant";
    else if (roleClean === "sales boy" || roleClean === "sales") roleTarget = "Sales Boy";
    else if (roleClean === "mis") roleTarget = "MIS";
    else if (roleClean === "store manager") roleTarget = "Store Manager";
    else roleTarget = role;

    sheet = ss.getSheetByName(roleTarget);
    if (sheet) sheetNameUsed = roleTarget;
    else {
      var allSheets = ss.getSheets();
      for (var r = 0; r < allSheets.length; r++) {
        var rName = allSheets[r].getName().trim();
        if (rName.toLowerCase() === roleClean.toLowerCase() || (roleTarget && rName.toLowerCase() === roleTarget.toLowerCase())) {
          sheet = allSheets[r];
          sheetNameUsed = rName;
          break;
        }
      }
    }
  }

  // 3. PARSE SHEET DATA IF FOUND
  if (sheet && sheet.getLastRow() >= 2) {
    var data = sheet.getDataRange().getValues();
    
    // Check if it's the Modern 14-Column table or Legacy format
    var h0 = data[0].map(function(h) { return String(h || '').toLowerCase().trim(); });
    var isModernTable = h0.some(function(h) {
      return h.includes('frequency') || h === 'type' || (h.includes('task') && !h.includes('daily'));
    });

    if (!isModernTable && data.length > 2) {
      var h1 = data[1].map(function(h) { return String(h || '').toLowerCase().trim(); });
      if (h1.some(function(h) { return h.includes('frequency') || h === 'type' || h.includes('sop'); })) {
        isModernTable = true;
      }
    }

    if (isModernTable) {
      // A. MODERN 14-COLUMN FORMAT PARSING
      var colMap = {
        type: -1, task: -1, subTask: -1, deadline: -1, timeSlot: -1, points: -1,
        sop: -1, goodSample: -1, goodPhoto: -1, badSample: -1, badPhoto: -1,
        troubleshoot: -1, refPhoto: -1, video: -1
      };

      h0.forEach(function(h, idx) {
        if (h.includes('freq') || h === 'type') colMap.type = idx;
        else if (h.includes('subtask') || h.includes('sub task') || h === 'item') colMap.subTask = idx;
        else if (h.includes('task') && colMap.task === -1) colMap.task = idx;
        else if (h.includes('dead') || (h.includes('time') && !h.includes('slot'))) colMap.deadline = idx;
        else if (h.includes('slot') || h.includes('shift')) colMap.timeSlot = idx;
        else if (h.includes('point') || h.includes('score')) colMap.points = idx;
        else if (h.includes('sop') || h.includes('step') || h.includes('procedure')) colMap.sop = idx;
        else if (h.includes('good') && (h.includes('photo') || h.includes('img') || h.includes('url') || h.includes('link'))) colMap.goodPhoto = idx;
        else if (h.includes('good') || h.includes('acceptable')) colMap.goodSample = idx;
        else if (h.includes('bad') && (h.includes('photo') || h.includes('img') || h.includes('url') || h.includes('link'))) colMap.badPhoto = idx;
        else if (h.includes('bad') || h.includes('mistake') || h.includes('rejected')) colMap.badSample = idx;
        else if (h.includes('trouble') || h.includes('help') || h.includes('faq')) colMap.troubleshoot = idx;
        else if (h.includes('ref') || h.includes('doc') || h.includes('form')) colMap.refPhoto = idx;
        else if (h.includes('video') || h.includes('vid') || h.includes('youtube')) colMap.video = idx;
      });

      if (colMap.task === -1) colMap.task = 1;
      if (colMap.type === -1) colMap.type = 0;
      if (colMap.subTask === -1) colMap.subTask = 2;
      if (colMap.deadline === -1) colMap.deadline = 3;

      var lastTask = "";
      var lastType = "Daily";

      for (var r = 1; r < data.length; r++) {
        var row = data[r];
        var rawType = colMap.type !== -1 && row[colMap.type] ? String(row[colMap.type]).trim() : "";
        var rawTask = colMap.task !== -1 && row[colMap.task] ? String(row[colMap.task]).trim() : "";
        var rawSub = colMap.subTask !== -1 && row[colMap.subTask] ? String(row[colMap.subTask]).trim() : "";

        if (!rawTask && !rawSub) continue;

        if (rawType) lastType = rawType;
        var curType = rawType || lastType;
        if (curType.toLowerCase().includes('daily')) curType = 'Daily';
        else if (curType.toLowerCase().includes('weekly')) curType = 'Weekly';
        else if (curType.toLowerCase().includes('monthly')) curType = 'Monthly';
        else curType = 'Daily';

        if (rawTask) lastTask = rawTask;
        var curTask = rawTask || lastTask;

        var rawDead = colMap.deadline !== -1 && row[colMap.deadline] ? formatTimeFriendly(row[colMap.deadline]) : "";
        if (!rawDead) {
          if (curType === 'Weekly') rawDead = "Monday, 07:00 PM";
          else if (curType === 'Monthly') rawDead = "1st of Month, 07:00 PM";
          else rawDead = "07:00 PM";
        }
        var rawSlot = colMap.timeSlot !== -1 && row[colMap.timeSlot] ? String(row[colMap.timeSlot]).trim() : "";
        if (!rawSlot) rawSlot = getTimeSlotFromDeadline(rawDead);

        var rawPoints = colMap.points !== -1 && row[colMap.points] ? parseInt(row[colMap.points], 10) || 10 : 10;
        var rawSop = colMap.sop !== -1 && row[colMap.sop] ? String(row[colMap.sop]).trim() : "";
        var rawGood = colMap.goodSample !== -1 && row[colMap.goodSample] ? String(row[colMap.goodSample]).trim() : "";
        var rawGoodPhoto = colMap.goodPhoto !== -1 && row[colMap.goodPhoto] ? String(row[colMap.goodPhoto]).trim() : "";
        var rawBad = colMap.badSample !== -1 && row[colMap.badSample] ? String(row[colMap.badSample]).trim() : "";
        var rawBadPhoto = colMap.badPhoto !== -1 && row[colMap.badPhoto] ? String(row[colMap.badPhoto]).trim() : "";
        var rawTrouble = colMap.troubleshoot !== -1 && row[colMap.troubleshoot] ? String(row[colMap.troubleshoot]).trim() : "";
        var rawRef = colMap.refPhoto !== -1 && row[colMap.refPhoto] ? String(row[colMap.refPhoto]).trim() : "";
        var rawVideo = colMap.video !== -1 && row[colMap.video] ? String(row[colMap.video]).trim() : "";

        // Fallback to smart role defaults if specific cell is empty
        var smartMeta = getSmartRoleDefaults(roleClean, curTask, rawSub);

        var sopSteps = [];
        if (rawSop) {
          sopSteps = rawSop.split(/\r?\n/).map(function(s) { return s.trim(); }).filter(Boolean);
        } else {
          sopSteps = smartMeta.sopSteps;
        }

        taskList.push({
          type: curType,
          task: curTask,
          subTask: rawSub || curTask,
          deadline: rawDead,
          timeSlot: rawSlot.toLowerCase(),
          points: rawPoints,
          sop: rawSop || smartMeta.sopSteps.join('\n'),
          sopSteps: sopSteps,
          goodSample: rawGood || smartMeta.goodSample,
          goodPhoto: rawGoodPhoto,
          badSample: rawBad || smartMeta.badSample,
          badPhoto: rawBadPhoto,
          troubleshoot: rawTrouble || smartMeta.troubleshoot,
          refPhoto: rawRef,
          sopVideo: rawVideo,
          sourceSheet: sheetNameUsed,
          status: "Pending ⏳"
        });
      }
    } else {
      // B. LEGACY 15-COLUMN FORMAT PARSING
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
            sop: dSop || lastDailySop || metaD.sopSteps.join('\n'),
            sopVideo: dVid || lastDailyVideo,
            sopSteps: dSop ? dSop.split(/\r?\n/).map(function(s){return s.trim();}).filter(Boolean) : metaD.sopSteps,
            goodSample: metaD.goodSample,
            goodPhoto: "",
            badSample: metaD.badSample,
            badPhoto: "",
            troubleshoot: metaD.troubleshoot,
            checklist: metaD.checklist,
            sourceSheet: sheetNameUsed,
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
            sop: wSop || lastWeeklySop || metaW.sopSteps.join('\n'),
            sopVideo: wVid || lastWeeklyVideo,
            sopSteps: wSop ? wSop.split(/\r?\n/).map(function(s){return s.trim();}).filter(Boolean) : metaW.sopSteps,
            goodSample: metaW.goodSample,
            goodPhoto: "",
            badSample: metaW.badSample,
            badPhoto: "",
            troubleshoot: metaW.troubleshoot,
            checklist: metaW.checklist,
            sourceSheet: sheetNameUsed,
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
            sop: mSop || lastMonthlySop || metaM.sopSteps.join('\n'),
            sopVideo: mVid || lastMonthlyVideo,
            sopSteps: mSop ? mSop.split(/\r?\n/).map(function(s){return s.trim();}).filter(Boolean) : metaM.sopSteps,
            goodSample: metaM.goodSample,
            goodPhoto: "",
            badSample: metaM.badSample,
            badPhoto: "",
            troubleshoot: metaM.troubleshoot,
            checklist: metaM.checklist,
            sourceSheet: sheetNameUsed,
            status: "Pending ⏳"
          });
        }
      }
    }
  }

  // 4. Fallback if taskList is empty: auto-generate standard tasks
  if (taskList.length === 0) {
    var defs = [
      { task: "Morning Shift Setup & Opening Review", sub: "Priority Verification", dead: "10:00 AM", slot: "morning", type: "Daily" },
      { task: "Core Operations & Checklist Execution", sub: "Daily Workflow", dead: "02:00 PM", slot: "midday", type: "Daily" },
      { task: "Closing Summary & WhatsApp EOD Report", sub: "EOD Verification", dead: "07:00 PM", slot: "closing", type: "Daily" }
    ];
    for (var d = 0; d < defs.length; d++) {
      var itemDef = defs[d];
      var metaDef = getSmartRoleDefaults(roleClean, itemDef.task, itemDef.sub);
      taskList.push({
        type: itemDef.type,
        task: itemDef.task,
        subTask: itemDef.sub,
        deadline: itemDef.dead,
        timeSlot: itemDef.slot,
        sop: metaDef.sopSteps.join('\n'),
        sopSteps: metaDef.sopSteps,
        goodSample: metaDef.goodSample,
        goodPhoto: "",
        badSample: metaDef.badSample,
        badPhoto: "",
        troubleshoot: metaDef.troubleshoot,
        checklist: metaDef.checklist,
        sourceSheet: "Default",
        status: "Pending ⏳"
      });
    }
  }

  if (roleClean === "sales boy" || roleClean === "sales") {
     taskList.push({ type: "Daily", task: "Send EOD Report on WhatsApp", subTask: "Screenshot & Final Update", deadline: "08:00 PM", timeSlot: "closing", sop: "", sopSteps: ["Take screenshot of day visits", "Share in group"], sopVideo: "", goodSample: "Complete counter visit list", badSample: "Incomplete visit log", troubleshoot: "If network slow, send SMS", status: "Pending ⏳" });
  }

  // 5. MERGE ASSIGNED DAILY, WEEKLY, MONTHLY TASKS FROM 'Adhoc_Tasks'
  var adhocSheet = ss.getSheetByName("Adhoc_Tasks");
  if (adhocSheet && adhocSheet.getLastRow() > 1) {
    var aData = adhocSheet.getDataRange().getValues();
    var sClean = staffName ? staffName.toString().trim().toLowerCase() : "";
    // 🔻 DESCENDING ORDER: Merge newest assigned tasks first
    for (var a = aData.length - 1; a >= 1; a--) {
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
        var metaA = getSmartRoleDefaults(roleClean, aDesc, aFreq);
        taskList.unshift({
          type: fType,
          task: "[Assigned] " + aDesc,
          subTask: aFreq,
          deadline: aDead || "N/A",
          timeSlot: getTimeSlotFromDeadline(aDead),
          sop: metaA.sopSteps.join('\n'),
          sopVideo: "",
          sopSteps: metaA.sopSteps,
          goodSample: metaA.goodSample,
          goodPhoto: "",
          badSample: metaA.badSample,
          badPhoto: "",
          troubleshoot: metaA.troubleshoot,
          checklist: metaA.checklist,
          refPhoto: aPhotos,
          sourceSheet: "Adhoc_Tasks",
          status: aStatus || "Pending ⏳"
        });
      }
    }
  }

  return taskList;
}


// =====================================================================
// 📌 SYNC ROUTINE TASK LIST WITH LIVE SUBMISSIONS
// Matches Daily (today), Weekly (this week), Monthly (this month) tasks
// Ensures submitted tasks STAY COMPLETED upon page refresh!
// =====================================================================
function syncTaskListWithSubmissions(taskList, staffName) {
  if (!taskList || taskList.length === 0 || !staffName) return taskList;

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var subSheet = ss.getSheetByName("Task_Submissions");
    if (!subSheet || subSheet.getLastRow() <= 1) return taskList;

    var subData = subSheet.getDataRange().getValues();
    var tz = Session.getScriptTimeZone();
    var now = new Date();
    var todayStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");
    var todayDateStr = Utilities.formatDate(now, tz, "dd-MMM");

    var dayOfWeek = now.getDay();
    var diffToMon = (dayOfWeek === 0 ? -6 : 1) - dayOfWeek;
    var startOfWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() + diffToMon, 0, 0, 0, 0);

    var currYear = now.getFullYear();
    var currMonth = now.getMonth();

    var sClean = staffName.toString().trim().toLowerCase();

    var dailySubmissions = {};
    var weeklySubmissions = {};
    var monthlySubmissions = {};

    function cleanTaskKey(str) {
      return (str || "").toString().toLowerCase().replace(/[^a-z0-9]/g, "");
    }

    for (var i = subData.length - 1; i >= 1; i--) {
      var rowStaff = subData[i][1] ? subData[i][1].toString().trim().toLowerCase() : "";
      if (!rowStaff) continue;

      var isStaffMatch = (rowStaff === sClean || rowStaff.replace(/\s+/g, '') === sClean.replace(/\s+/g, '') || rowStaff.includes(sClean) || sClean.includes(rowStaff));
      if (!isStaffMatch) continue;

      var rawDate = subData[i][0];
      if (!rawDate) continue;
      var subDate = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
      if (isNaN(subDate.getTime())) {
        if (typeof rawDate === "string") {
          var dMatch = rawDate.match(/^(\d{4})-(\d{2})-(\d{2})/);
          if (dMatch) {
            subDate = new Date(parseInt(dMatch[1], 10), parseInt(dMatch[2], 10) - 1, parseInt(dMatch[3], 10));
          }
        }
      }
      if (isNaN(subDate.getTime())) continue;

      var subDateStr = Utilities.formatDate(subDate, tz, "yyyy-MM-dd");
      var status = subData[i][8] ? subData[i][8].toString().trim() : "";
      var isToday = (subDateStr === todayStr) || (status.indexOf(todayDateStr) !== -1);
      var isThisWeek = (subDate >= startOfWeek);
      var isThisMonth = (subDate.getFullYear() === currYear && subDate.getMonth() === currMonth);

      var isDoneStatus = (status.indexOf("Uploaded") !== -1 || status.indexOf("Completed") !== -1 || status.indexOf("Skipped") !== -1);
      if (!isDoneStatus) continue;

      var subTask = subData[i][4] ? subData[i][4].toString().trim() : "";
      var subSubTask = subData[i][5] ? subData[i][5].toString().trim() : "";
      var rawLink = subData[i][9] ? subData[i][9].toString().trim() : "";
      var submitTime = subData[i][7] ? subData[i][7].toString().trim() : "";

      var photoUrl = "";
      if (rawLink) {
        var m = rawLink.match(/=HYPERLINK\("([^"]+)"/i) || rawLink.match(/(https?:\/\/[^\s",)]+)/i);
        if (m) photoUrl = m[1];
      }

      var subEntry = {
        status: status,
        photoUrl: photoUrl,
        submitTime: submitTime,
        date: subDateStr
      };

      var cTask = cleanTaskKey(subTask.replace(/^\[assigned\]\s*/i, ''));
      var cSub = cleanTaskKey(subSubTask);

      var kExact = cTask + "__" + cSub;
      var kTaskOnly = cTask + "__";
      var kBothSame = cTask + "__" + cTask;
      var kSubOnly = "__" + cSub;

      function storeInMap(mapObj) {
        if (!mapObj[kExact]) mapObj[kExact] = subEntry;
        if (!cSub || cSub === cTask) {
          if (!mapObj[kTaskOnly]) mapObj[kTaskOnly] = subEntry;
          if (!mapObj[kBothSame]) mapObj[kBothSame] = subEntry;
        } else {
          if (!mapObj[kSubOnly]) mapObj[kSubOnly] = subEntry;
        }
      }

      if (isToday) storeInMap(dailySubmissions);
      if (isThisWeek) storeInMap(weeklySubmissions);
      if (isThisMonth) storeInMap(monthlySubmissions);
    }

    for (var t = 0; t < taskList.length; t++) {
      var item = taskList[t];
      var type = (item.type || "Daily").toLowerCase();

      var targetMap = dailySubmissions;
      if (type.indexOf("week") !== -1) targetMap = weeklySubmissions;
      else if (type.indexOf("month") !== -1) targetMap = monthlySubmissions;

      var itemTask = cleanTaskKey((item.task || "").replace(/^\[assigned\]\s*/i, ''));
      var itemSub = cleanTaskKey(item.subTask);

      var hasDistinctSub = itemSub && itemSub !== itemTask && itemSub !== "actionrequired";

      var mExact = itemTask + "__" + itemSub;
      var mBothSame = itemTask + "__" + itemTask;
      var mTaskOnly = itemTask + "__";
      var mSubOnly = "__" + itemSub;

      var matched = targetMap[mExact];
      if (!matched && hasDistinctSub) {
        matched = targetMap[mSubOnly];
      }
      if (!matched && !hasDistinctSub) {
        matched = targetMap[mBothSame] || targetMap[mTaskOnly];
      }

      if (matched) {
        item.status = matched.status;
        item.isDone = true;
        if (matched.photoUrl) item.proofUrl = matched.photoUrl;
        if (matched.submitTime) item.submitTime = matched.submitTime;
      }
    }
  } catch(e) {
    Logger.log("syncTaskListWithSubmissions error: " + e);
  }

  // 6. SYNCHRONIZE WITH LIVE SUBMISSIONS (TODAY FOR DAILY, THIS WEEK FOR WEEKLY, THIS MONTH FOR MONTHLY)
  syncTaskListWithSubmissions(taskList, staffName);

  return taskList;
}

// =====================================================================
// 📌 SETUP STAFF TASK SHEETS AUTO-CREATOR (CAN BE CALLED FROM APP OR ADMIN)
// Creates a clean, beautifully formatted 14-column task table for any staff
// =====================================================================
function setupStaffTaskSheet(targetName, roleName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = targetName.trim();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  // Set formatted headers
  var headers = [
    "Frequency", "Task Name", "SubTask / Item", "Deadline", "Shift Time Slot",
    "Points", "Step-by-Step SOP (Text)", "Good Sample (Text)", "Good Sample Photo URL",
    "Bad Sample / Mistake (Text)", "Bad Sample Photo URL", "Troubleshoot (Help if Stuck)",
    "Reference Form Photo URL", "Training Video URL"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
       .setBackground("#1f6feb")
       .setFontColor("#ffffff")
       .setFontWeight("bold")
       .setHorizontalAlignment("center");
  sheet.setFrozenRows(1);

  // If newly created or empty, add role baseline rows
  if (sheet.getLastRow() <= 1) {
    var rLower = (roleName || "").toLowerCase();
    var rows = [];

    if (rLower.includes("account")) {
      rows.push([
        "Daily", "Cash Book ka hissab dena", "Physical Cash Count & Denomination", "07:00 PM", "Closing", 10,
        "1. Physical cash drawer mein sabhi notes & coins ko denomination-wise count karein (500x, 200x, 100x).\n2. Kal ke closing balance aur cash book se match karein.\n3. Din ke sabhi cash vouchers (Receipts & Payments) entry check karein.\n4. Net Physical Cash aur Book Balance 100% tally hone par hi register photo click karein.",
        "Clear, unblurred photo jisme: Date, Opening Cash, Denomination Tally, Total Receipts, Total Payments aur Net Cash clearly visible ho.",
        "https://images.unsplash.com/photo-1554224155-8d04cb21cd6c?w=600&auto=format&fit=crop&q=80",
        "Sirf total amount ka rough paper scribble, dhundli photo, ya cash physical count kiye bina submit karna.",
        "https://images.unsplash.com/photo-1584433144859-1fc3ab64a957?w=600&auto=format&fit=crop&q=80",
        "Agar cash mein farq (discrepancy) aaye toh pehle last 3 payment vouchers check karein aur turant Owner ko report karein. Skip na karein.",
        "https://images.unsplash.com/photo-1450133064473-71024230f91b?w=600&auto=format&fit=crop&q=80",
        "https://www.youtube.com/watch?v=dQw4w9WgXcQ"
      ]);
      rows.push([
        "Daily", "EOD report bhejna", "Final Day Reconciliation & Summary", "07:00 PM", "Closing", 10,
        "1. Din bhar ki sabhi sale, collection aur expense entries close karein.\n2. Day-end summary report ERP se generate karein.\n3. Summary verify karke photo upload karein aur WhatsApp par send karein.",
        "Verified ERP summary photo with matching date and zero unposted vouchers.",
        "",
        "Half-screen photo ya bina verify kiye pending entries chhod kar report bhejna.",
        "",
        "System sync delay ho toh internet restart karein aur backup spreadsheet se number verify karein.",
        "",
        ""
      ]);
      rows.push([
        "Daily", "Daily Reports", "Sales Report", "11:00 AM", "Morning", 10,
        "1. ERP mein Sales Register open karein.\n2. Aaj ki total billing aur counter-wise breakup review karein.\n3. Report export ya screenshot lekar upload karein.",
        "Full sales report showing counter totals and date.",
        "",
        "Incomplete counter list or yesterday's date.",
        "",
        "ERP login issue ho toh admin se password reset karwayein.",
        "",
        ""
      ]);
      rows.push([
        "Daily", "Daily Reports", "Stock Report", "11:00 AM", "Morning", 10,
        "1. Warehouse inventory valuation screen open karein.\n2. Negative stock ya zero stock alert list check karein.\n3. Summary upload karein.",
        "Clean stock valuation summary.",
        "",
        "Missing category filters.",
        "",
        "Physical stock se mismatch ho toh Store Manager se coordinate karein.",
        "",
        ""
      ]);
      rows.push([
        "Daily", "Daily Reports", "Razorpay Report", "11:00 AM", "Morning", 10,
        "1. Razorpay merchant dashboard login karein.\n2. Captured payments vs settled amount tally karein.\n3. Settlement screenshot upload karein.",
        "Merchant dashboard showing total settled volume and UTR references.",
        "",
        "Unsettled or pending transaction ignored.",
        "",
        "Settlement on-hold ho toh Razorpay support ticket check karein.",
        "",
        ""
      ]);
      rows.push([
        "Daily", "Daily Reports", "Outstanding report", "11:00 AM", "Morning", 10,
        "1. Debtors aging report run karein (>30 days overdue).\n2. Top 10 overdue parties ko highlight karein.\n3. Recovery update note karein.",
        "Full debtors aging summary with party-wise dues.",
        "",
        "Old cached report without today's receipts.",
        "",
        "Party dispute par Sales Boy se collection follow-up mangein.",
        "",
        ""
      ]);
      rows.push([
        "Daily", "Daily Reports", "Kelock Report", "11:00 AM", "Morning", 10,
        "1. Kelock account ledger reconcile karein.\n2. Debit-Credit balance verify karein.\n3. Confirmation screenshot attach karein.",
        "Matched ledger balance statement.",
        "",
        "Unreconciled journal entry.",
        "",
        "Discrepancy ho toh invoice bills re-verify karein.",
        "",
        ""
      ]);
    } else {
      rows.push([
        "Daily", "Morning Shift Setup", "Priority Verification", "10:00 AM", "Morning", 10,
        "1. Review today's task list.\n2. Prepare tools and workspace.\n3. Confirm starting status.",
        "Complete opening checklist verified.", "", "Skipping setup or late start.", "",
        "Contact supervisor if material or access is blocked.", "", ""
      ]);
      rows.push([
        "Daily", "Shift Closing & EOD", "Day End Wrap Up", "07:00 PM", "Closing", 10,
        "1. Clean workspace.\n2. Submit pending tasks.\n3. Send WhatsApp EOD report.",
        "Accurate EOD report submitted before 07:00 PM.", "", "Leaving without sending EOD report.", "",
        "Notify group if working overtime.", "", ""
      ]);
    }

    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.autoResizeColumns(1, headers.length);
  }

  return { success: true, sheetName: sheetName, message: "Sheet '" + sheetName + "' ready for task management!" };
}



// =====================================================================
// 📦 9-STAGE LIVE ORDER PROCESSING & 1-HOUR SLA ENGINE
// =====================================================================
var ORDER_PROCESSING_API_URL = "https://script.google.com/macros/s/AKfycbyXHJaFTSkjwu6C8ze4czs2Q42NQsgzjW814j8OLA7ztMXTW-nqmh2gkpjebmdSW4YJjA/exec";

var ORDER_STAGE_NAMES = [
  "Balance Check",                // Step 1: Accountant
  "WhatsApp 1 (Confirmation)",    // Step 2: CRM
  "Processed (Stock)",            // Step 3: Store Manager
  "Invoiced",                     // Step 4: Accountant
  "Dispatched",                   // Step 5: Store Manager
  "Logistics Allocation",         // Step 6: Store Manager
  "Delivery Comms (Call/WA)",     // Step 7: CRM
  "Delivered",                    // Step 8: Store Manager / Delivery
  "Final Call (Rating)"           // Step 9: CRM
];

var ORDER_STAGE_ROLES = [
  "Accountant",     // Step 1
  "CRM",            // Step 2
  "Store Manager",  // Step 3
  "Accountant",     // Step 4
  "Store Manager",  // Step 5
  "Store Manager",  // Step 6
  "CRM",            // Step 7
  "Store Manager",  // Step 8
  "CRM"             // Step 9
];

function fetchLiveOrderPipeline() {
  try {
    var options = {
      method: "post",
      contentType: "text/plain",
      payload: JSON.stringify({ action: "getOrders", staffName: "CEO" }),
      muteHttpExceptions: true,
      followRedirects: true
    };
    var response = UrlFetchApp.fetch(ORDER_PROCESSING_API_URL, options);
    var json = JSON.parse(response.getContentText());
    if (json && json.status === "success" && json.data && Array.isArray(json.data)) {
      return json.data;
    }
  } catch(e) {
    Logger.log("Order App API fetch error: " + e.toString());
  }

  // Fallback to local Orders sheet
  return getLocalOrdersFallback();
}

function getLocalOrdersFallback() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Orders");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var list = [];
  var tz = Session.getScriptTimeZone();
  var todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

  for (var i = 1; i < data.length; i++) {
    var rawDate = data[i][0];
    if (!rawDate) continue;
    var dStr = Utilities.formatDate(new Date(rawDate), tz, "yyyy-MM-dd");
    if (dStr !== todayStr) continue;

    var orderId = data[i][1] ? data[i][1].toString().trim() : "";
    var shopName = data[i][2] ? data[i][2].toString().trim() : "";
    var items = data[i][4] ? data[i][4].toString().trim() : "";
    var status = data[i][8] ? data[i][8].toString().trim() : "";

    if (orderId !== "") {
      list.push({
        orderId: orderId,
        shopName: shopName,
        items: items,
        date: Utilities.formatDate(new Date(rawDate), tz, "yyyy-MM-dd HH:mm:ss"),
        completedStages: (status.indexOf("Completed") !== -1 ? 9 : 2),
        stageUrls: [],
        paymentMode: "PREPAID"
      });
    }
  }
  return list;
}

function parseOrderDateString(dateStr, fallbackDate) {
  if (!dateStr) return fallbackDate;
  try {
    var d = new Date(dateStr);
    if (!isNaN(d.getTime())) return d;
  } catch(e) {}
  return fallbackDate;
}

function getEnrichedOrderPipeline() {
  var rawOrders = fetchLiveOrderPipeline();
  if (!rawOrders || !Array.isArray(rawOrders)) return [];

  var now = new Date();
  var tz = Session.getScriptTimeZone();
  var enriched = [];

  for (var i = 0; i < rawOrders.length; i++) {
    var o = rawOrders[i];
    if (!o || !o.orderId) continue;

    var completed = Number(o.completedStages || 0);
    var isFullyCompleted = (completed >= 9);
    var currentStageNum = isFullyCompleted ? 9 : (completed + 1);
    var currentStageName = isFullyCompleted ? "Fully Completed" : ORDER_STAGE_NAMES[completed];
    var currentAssignedRole = isFullyCompleted ? "Completed" : ORDER_STAGE_ROLES[completed];

    // Stage entry timestamp
    var stageEnteredTime = now;
    if (completed === 0) {
      stageEnteredTime = parseOrderDateString(o.date, now);
    } else if (o.stageUrls && o.stageUrls[completed - 1]) {
      var prevCell = String(o.stageUrls[completed - 1]);
      var tsMatch = prevCell.match(/TIMESTAMP:\s*([^\n|]+)/i);
      if (tsMatch) {
        stageEnteredTime = parseOrderDateString(tsMatch[1].trim(), now);
      } else {
        stageEnteredTime = parseOrderDateString(o.date, now);
      }
    } else {
      stageEnteredTime = parseOrderDateString(o.date, now);
    }

    // 1-Hour SLA Deadline (+60 mins)
    var deadlineTime = new Date(stageEnteredTime.getTime() + 60 * 60 * 1000);
    var diffMs = deadlineTime.getTime() - now.getTime();
    var remainingMins = Math.round(diffMs / 60000);
    var isOverdue = (!isFullyCompleted && remainingMins < 0);
    var overdueMins = isOverdue ? Math.abs(remainingMins) : 0;

    // Proof photo URLs
    var stageProof = "";
    if (o.stageUrls && Array.isArray(o.stageUrls)) {
      stageProof = o.stageUrls.filter(function(u){ return u && String(u).indexOf("http") !== -1; }).join(" , ");
    }

    enriched.push({
      orderId: o.orderId,
      shopName: o.shopName || o.dealerName || "Dealer",
      phone: o.dealerPhone || o.phone || "",
      items: o.items || "Order Items",
      orderDate: Utilities.formatDate(stageEnteredTime, tz, "dd MMM, hh:mm a"),
      dateRaw: o.date,
      paymentMode: o.paymentMode || "PREPAID",
      completedStages: completed,
      currentStageNum: currentStageNum,
      currentStageName: currentStageName,
      assignedRole: currentAssignedRole,
      isFullyCompleted: isFullyCompleted,
      isOverdue: isOverdue,
      remainingMinutes: remainingMins,
      overdueMinutes: overdueMins,
      slaDeadlineText: Utilities.formatDate(deadlineTime, tz, "hh:mm a"),
      slaDeadlineEpoch: deadlineTime.getTime(),
      proofUrls: stageProof,
      stageUrls: o.stageUrls || []
    });
  }

  // Sort: Incomplete orders first, and newest orders (Today's fresh orders) right at the top!
  enriched.sort(function(a, b) {
    if (a.isFullyCompleted !== b.isFullyCompleted) return a.isFullyCompleted ? 1 : -1;
    var timeA = a.dateRaw ? new Date(a.dateRaw).getTime() : 0;
    var timeB = b.dateRaw ? new Date(b.dateRaw).getTime() : 0;
    return timeB - timeA;
  });

  return enriched;
}

function getOrdersForRole(role, staffName) {
  var allEnriched = getEnrichedOrderPipeline();
  var roleClean = role ? role.toString().trim().toLowerCase() : "";

  // 7:00 PM Dispatch Cut-off calculations
  var now = new Date();
  var cutoffDate = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 19, 0, 0, 0);
  var cutoffDiffMins = Math.round((cutoffDate.getTime() - now.getTime()) / 60000);
  var cutoffStatus = {
    cutoffTime: "07:00 PM",
    remainingMins: cutoffDiffMins,
    isPassed: cutoffDiffMins < 0
  };

  // Management, CEO, Admin, MIS, and Manager oversee all pipeline orders
  if (roleClean === "owner" || roleClean === "ceo" || roleClean === "admin" || roleClean === "mis" || roleClean === "manager") {
    return {
      cutoff: cutoffStatus,
      orders: allEnriched
    };
  }

  var filtered = allEnriched.filter(function(o) {
    if (o.isFullyCompleted) return false;
    var assignedLower = o.assignedRole.toLowerCase();

    if (roleClean === "accountant") {
      return assignedLower === "accountant";
    }
    if (roleClean === "store manager") {
      return assignedLower === "store manager";
    }
    if (roleClean === "crm" || roleClean === "smm") {
      return assignedLower === "crm" || assignedLower === "smm";
    }
    return assignedLower === roleClean;
  });

  return {
    cutoff: cutoffStatus,
    orders: filtered
  };
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
// =====================================================================
// 📌 CEO TASK ASSIGNMENT: PERMANENT GOOGLE SHEET SCHEDULE WRITER
// Writes Daily, Weekly & Monthly tasks permanently to the staff's 14-col sheet
// =====================================================================
function assignTaskPermanently(assigneeInfo, taskTypeVal, taskName, subTask, deadlineStr, timeSlot, sopText, photoArray, priority) {
  if (!assigneeInfo || assigneeInfo.includes("----")) return "Error: Kripya valid staff ya team chunein!";
  if (!taskName || taskName.trim() === "") return "Error: Kripya task ka naam bharein!";

  var cleanName = assigneeInfo.replace("🏢 Team: ", "").replace("👤 ", "").replace("👥 ", "").trim();
  var isAllStaff = assigneeInfo.includes("👥 All Staff") || cleanName.toLowerCase() === "all staff" || cleanName.toLowerCase() === "all";
  var isTeam = assigneeInfo.includes("🏢 Team: ");
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Save reference photos to Drive if any
  var photoUrls = [];
  if (photoArray && photoArray.length > 0) {
    try {
      var folder = getOrCreatePhotoFolder();
      for (var p = 0; p < photoArray.length; p++) {
        var pItem = photoArray[p];
        if (pItem.base64Data && pItem.fileName) {
          var blob = Utilities.newBlob(Utilities.base64Decode(pItem.base64Data), pItem.mimeType || "image/jpeg", pItem.fileName);
          var file = folder.createFile(blob);
          try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}
          photoUrls.push(file.getUrl());
        }
      }
    } catch(e) {}
  }
  var photoLinksString = photoUrls.length > 0 ? photoUrls.join(" , ") : "No Photo";

  var freqClean = "Daily";
  var tLower = (taskTypeVal || "").toLowerCase();
  if (tLower.includes("daily")) freqClean = "Daily";
  else if (tLower.includes("weekly")) freqClean = "Weekly";
  else if (tLower.includes("monthly")) freqClean = "Monthly";
  else if (tLower.includes("extra") || tLower.includes("one time")) freqClean = "Extra (One Time)";

  // If Extra (One Time), record in Adhoc_Tasks sheet
  if (freqClean === "Extra (One Time)") {
    var adhocSheet = ss.getSheetByName("Adhoc_Tasks");
    if (!adhocSheet) {
      adhocSheet = ss.insertSheet("Adhoc_Tasks");
      adhocSheet.appendRow(["Assign Date", "Staff Name", "Task Description", "Deadline", "Status", "Reference Photo", "Task Frequency"]);
    }
    adhocSheet.appendRow([new Date(), cleanName, taskName + (subTask ? " - " + subTask : ""), deadlineStr, "Pending ⏳", photoLinksString, "Extra (One Time)"]);
    return "⚡ Extra One-Time Task assigned to " + cleanName + " 🚀";
  }

  // For Daily, Weekly, Monthly: Build 14-Column permanent row
  var smartDefaults = getSmartRoleDefaults(isTeam ? cleanName : getStaffRoleByName(cleanName), taskName, subTask);
  var finalSop = (sopText && sopText.trim() !== "") ? sopText.trim() : smartDefaults.sopSteps.join("\n");
  var finalSlot = (timeSlot && timeSlot.trim() !== "") ? timeSlot.trim() : getTimeSlotFromDeadline(deadlineStr);

  var newRow = [
    freqClean,                                                    // 1. Frequency (Daily, Weekly, Monthly)
    taskName.trim(),                                              // 2. Task Name
    (subTask || taskName).trim(),                                 // 3. SubTask / Item
    deadlineStr.trim(),                                           // 4. Deadline
    finalSlot,                                                    // 5. Shift Time Slot
    10,                                                           // 6. Points
    finalSop,                                                     // 7. Step-by-Step SOP
    smartDefaults.goodSample || "Verified clean entry with unblurred photo / document matching exact figures.", // 8. Good Sample
    "",                                                           // 9. Good Photo URL
    smartDefaults.badSample || "Incomplete execution, delay past deadline, or blurry/empty proof.",             // 10. Bad Sample
    "",                                                           // 11. Bad Photo URL
    smartDefaults.troubleshoot || "Agar koi issue ya discrepancy aaye toh supervisor ya CEO ko report karein.", // 12. Troubleshoot
    photoLinksString,                                             // 13. Reference Form Photo URL
    ""                                                            // 14. Training Video URL
  ];

  function appendToTargetSheet(sheetName, roleForHeaders) {
    if (!sheetName) return;
    var sh = ss.getSheetByName(sheetName);
    if (!sh) {
      setupStaffTaskSheet(sheetName, roleForHeaders || sheetName);
      sh = ss.getSheetByName(sheetName);
    }
    if (sh) {
      sh.appendRow(newRow);
    }
  }

  var targetsAppended = [];

  if (isAllStaff) {
    var masterSheet = ss.getSheetByName("staff master");
    if (masterSheet) {
      var mData = masterSheet.getDataRange().getValues();
      for (var m = 1; m < mData.length; m++) {
        var sName = mData[m][1] ? mData[m][1].toString().trim() : "";
        var sRole = mData[m][5] ? mData[m][5].toString().trim() : (mData[m][4] ? mData[m][4].toString().trim() : "Staff");
        var rLow = sRole.toLowerCase();
        if (sName && rLow !== "owner" && rLow !== "ceo" && rLow !== "admin") {
          appendToTargetSheet(sName, sRole);
          if (targetsAppended.indexOf(sName) === -1) targetsAppended.push(sName);
        }
      }
    }
  } else if (isTeam) {
    appendToTargetSheet(cleanName, cleanName);
    targetsAppended.push("Team: " + cleanName);

    var masterSheet = ss.getSheetByName("staff master");
    if (masterSheet) {
      var mData = masterSheet.getDataRange().getValues();
      for (var m = 1; m < mData.length; m++) {
        var sName = mData[m][1] ? mData[m][1].toString().trim() : "";
        var sRole = mData[m][5] ? mData[m][5].toString().trim() : (mData[m][4] ? mData[m][4].toString().trim() : "Staff");
        if (sName && sRole.toLowerCase() === cleanName.toLowerCase()) {
          var existingSh = ss.getSheetByName(sName);
          if (existingSh) {
            existingSh.appendRow(newRow);
            if (targetsAppended.indexOf(sName) === -1) targetsAppended.push(sName);
          }
        }
      }
    }
  } else {
    var staffRole = getStaffRoleByName(cleanName);
    appendToTargetSheet(cleanName, staffRole);
    targetsAppended.push(cleanName);
  }

  // Audit log in Adhoc_Tasks
  try {
    var adhocSheet = ss.getSheetByName("Adhoc_Tasks");
    if (!adhocSheet) {
      adhocSheet = ss.insertSheet("Adhoc_Tasks");
      adhocSheet.appendRow(["Assign Date", "Staff Name", "Task Description", "Deadline", "Status", "Reference Photo", "Task Frequency"]);
    }
    adhocSheet.appendRow([new Date(), cleanName, taskName + (subTask ? " - " + subTask : ""), deadlineStr, "Permanent Sheet Added ✅", photoLinksString, freqClean]);
  } catch(e) {}

  return "✅ " + freqClean + " task permanently added to Google Sheet (" + targetsAppended.join(", ") + ")! 🚀";
}

function assignAdhocTask(assigneeInfo, taskTypeVal, taskDesc, deadline, photoArray) {
  return assignTaskPermanently(assigneeInfo, taskTypeVal, taskDesc, "", deadline, "", "", photoArray, "Normal");
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
  
  // 🔻 DESCENDING ORDER: Iterate backwards from newest row to oldest row
  for(var i = data.length - 1; i >= 1; i--) {
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
  try {
    var role = getStaffRoleByName(staffName) || "Staff";
    var routineTasks = getTasksForRole(role, staffName) || [];
    var adhocTasks = getAdhocTasksForStaff(staffName) || [];
    var scores = null;
    try {
      scores = getStaffScoresBreakdown(staffName);
    } catch(errScores) {
      Logger.log("Scores calculation error for " + staffName + ": " + errScores);
    }
    return {
      success: true,
      routine: routineTasks,
      adhoc: adhocTasks,
      role: role,
      scores: scores
    };
  } catch(e) {
    Logger.log("getStaffProgressReport error: " + e);
    return {
      success: false,
      error: e.toString(),
      routine: [],
      adhoc: [],
      role: "Staff",
      scores: null
    };
  }
}

function getAllAdhocHistory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Adhoc_Tasks");
  if(!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var history = [];
  var start = Math.max(1, data.length - 150);
  
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

function parseDeadlineToDate(deadlineStr, baseDate) {
  if (!deadlineStr || deadlineStr === "N/A" || deadlineStr === "Jald") return null;
  var now = baseDate || new Date();
  var lower = String(deadlineStr).toLowerCase().trim();

  var hrs = 19;
  var mins = 0;
  var timeMatch = lower.match(/(\d{1,2}):(\d{2})\s*(am|pm)?/i);
  if (timeMatch) {
    hrs = parseInt(timeMatch[1], 10);
    mins = parseInt(timeMatch[2], 10);
    var ampm = timeMatch[3] ? timeMatch[3].toUpperCase() : "";
    if (ampm === "PM" && hrs < 12) hrs += 12;
    if (ampm === "AM" && hrs === 12) hrs = 0;
  }

  // Weekly Day Check (e.g., Monday, Every Monday)
  var daysMap = { sunday: 0, monday: 1, tuesday: 2, wednesday: 3, thursday: 4, friday: 5, saturday: 6 };
  for (var day in daysMap) {
    if (lower.indexOf(day) !== -1) {
      var targetDay = daysMap[day];
      var curDay = now.getDay();
      var diff = targetDay - curDay;
      if (diff < 0) diff += 7;
      return new Date(now.getFullYear(), now.getMonth(), now.getDate() + diff, hrs, mins, 0);
    }
  }

  // Monthly Date Check (e.g., 1st, 5th of Month, 25th)
  var strWithoutTime = lower.replace(/(\d{1,2}):(\d{2})\s*(am|pm)?/i, '').replace(/(\d{1,2})\s*(am|pm)/i, '').trim();
  var dateMatch = strWithoutTime.match(/(\d{1,2})(?:st|nd|rd|th)?(?:\s*of\s*month)?/i);
  if (dateMatch && (lower.indexOf("month") !== -1 || lower.indexOf("st") !== -1 || lower.indexOf("nd") !== -1 || lower.indexOf("rd") !== -1 || lower.indexOf("th") !== -1 || !isNaN(Number(strWithoutTime.split(/[,\s]/)[0])))) {
    var targetDateNum = parseInt(dateMatch[1], 10);
    if (targetDateNum >= 1 && targetDateNum <= 31) {
      var curDateNum = now.getDate();
      var targetMonth = now.getMonth();
      var targetYear = now.getFullYear();
      if (curDateNum > targetDateNum) {
        targetMonth++;
        if (targetMonth > 11) { targetMonth = 0; targetYear++; }
      }
      return new Date(targetYear, targetMonth, targetDateNum, hrs, mins, 0);
    }
  }

  // Direct ISO / Full Date Check
  if (lower.indexOf("-") !== -1 && (lower.indexOf("202") !== -1 || lower.indexOf("t") !== -1)) {
    var d = new Date(deadlineStr);
    if (!isNaN(d.getTime())) return d;
  }

  // Fallback to today at parsed time
  return new Date(now.getFullYear(), now.getMonth(), now.getDate(), hrs, mins, 0);
}

function isTaskOnTime(deadlineStr, submitDateObj) {
  if (!deadlineStr || deadlineStr === "N/A" || deadlineStr === "Jald") return true; 
  try {
    var targetDate = parseDeadlineToDate(deadlineStr, submitDateObj);
    if (!targetDate) return true;
    return submitDateObj.getTime() <= targetDate.getTime();
  } catch(e) { return true; }
}

// =====================================================================
// 📌 CEO CONTACT & NOTIFICATION CONFIGURATION
// =====================================================================
function getCeoContactInfo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ceo = {
    name: "Yash Barlota",
    email: "",
    phone: ""
  };

  try {
    var masterSheet = ss.getSheetByName("staff master");
    if (masterSheet) {
      var data = masterSheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var name = row[1] ? row[1].toString().trim() : "";
        var role = row[5] ? row[5].toString().trim() : (row[4] ? row[4].toString().trim() : "");
        var phone = row[2] ? row[2].toString().trim() : "";
        
        var rLow = role.toLowerCase();
        var nLow = name.toLowerCase();
        if (rLow === "owner" || rLow === "ceo" || rLow === "admin" || nLow.indexOf("yash") !== -1 || nLow.indexOf("barlota") !== -1) {
          if (name) ceo.name = name;
          if (phone) ceo.phone = phone;
          
          for (var c = 0; c < row.length; c++) {
            var val = row[c] ? row[c].toString().trim() : "";
            if (val.indexOf("@") !== -1 && val.indexOf(".") !== -1) {
              ceo.email = val;
              break;
            }
          }
          break;
        }
      }
    }
  } catch(e) {
    Logger.log("getCeoContactInfo error: " + e);
  }

  // Fallback to active/effective Google Script user email
  if (!ceo.email) {
    try {
      var eff = Session.getEffectiveUser().getEmail();
      if (eff && eff.indexOf("@") !== -1) ceo.email = eff;
    } catch(e) {}
  }
  if (!ceo.email) {
    try {
      var act = Session.getActiveUser().getEmail();
      if (act && act.indexOf("@") !== -1) ceo.email = act;
    } catch(e) {}
  }

  return ceo;
}

// =====================================================================
// 📌 CEO LIVE COMPLETED TASKS FEED (TODAY'S ACTIVITY STREAM)
// =====================================================================
function getCeoCompletedTasksFeed(limit) {
  limit = limit || 50;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Task_Submissions");
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var data = sheet.getDataRange().getValues();
  var tz = Session.getScriptTimeZone();
  var now = new Date();
  var todayStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");
  var list = [];
  var fallbackList = [];

  for (var i = data.length - 1; i >= 1; i--) {
    var rawDate = data[i][0];
    if (!rawDate) continue;
    var subDate = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
    if (isNaN(subDate.getTime())) continue;

    var dateStr = Utilities.formatDate(subDate, tz, "yyyy-MM-dd");
    var isToday = (dateStr === todayStr);

    var staffName = data[i][1] ? data[i][1].toString().trim() : "Staff";
    var role = data[i][2] ? data[i][2].toString().trim() : "Staff";
    var taskType = data[i][3] ? data[i][3].toString().trim() : "Daily";
    var task = data[i][4] ? data[i][4].toString().trim() : "";
    var subTask = data[i][5] ? data[i][5].toString().trim() : "";
    var deadline = data[i][6] ? data[i][6].toString().trim() : "";
    var submitTime = data[i][7] ? data[i][7].toString().trim() : Utilities.formatDate(subDate, tz, "hh:mm a");
    var status = data[i][8] ? data[i][8].toString().trim() : "";
    var rawLink = data[i][9] ? data[i][9].toString().trim() : "";
    var points = parseFloat(data[i][10]) || 0;

    var photoUrl = "";
    if (rawLink) {
      var match = rawLink.match(/=HYPERLINK\("([^"]+)"/i) || rawLink.match(/(https?:\/\/[^\s",)]+)/i);
      if (match) photoUrl = match[1];
    }

    var item = {
      submitTimestamp: Utilities.formatDate(subDate, tz, "yyyy-MM-dd HH:mm:ss"),
      timeDisplay: (isToday ? "" : Utilities.formatDate(subDate, tz, "dd-MMM ") ) + submitTime,
      isToday: isToday,
      staffName: staffName,
      role: role,
      taskType: taskType,
      task: task,
      subTask: subTask,
      deadline: deadline,
      status: status,
      points: points,
      isOnTime: (points >= 10),
      isSkipped: (status.toLowerCase().indexOf("skip") !== -1),
      photoUrl: photoUrl,
      hasPhoto: Boolean(photoUrl)
    };

    if (isToday) {
      list.push(item);
      if (list.length >= limit) break;
    } else if (fallbackList.length < 15) {
      fallbackList.push(item);
    }
  }

  return list.length > 0 ? list : fallbackList;
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
  var fileUrl = "";

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
      fileUrl = file.getUrl();
      hyperlinkFormula = '=HYPERLINK("' + fileUrl + '", "📷 View Photo")';
      statusPrefix = "Uploaded ✅";
    } catch(e) {}
  }

  var onTime = false;
  if (actionType !== "Skip" && actionType !== true) {
    pointsEarned = 5; 
    if (isTaskOnTime(deadline, now)) { pointsEarned += 5; onTime = true; }
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
  SpreadsheetApp.flush();

  // 🔔 NOTIFY CEO (Yash Barlota) VIA BACKGROUND EMAIL
  var ceo = getCeoContactInfo();
  if (actionType !== "Skip" && actionType !== true) {
    try {
      var recipientEmails = [];
      if (ceo && ceo.email && ceo.email.indexOf("@") !== -1) recipientEmails.push(ceo.email);
      var effEmail = Session.getEffectiveUser().getEmail();
      if (effEmail && effEmail.indexOf("@") !== -1 && recipientEmails.indexOf(effEmail) === -1) {
        recipientEmails.push(effEmail);
      }

      if (recipientEmails.length > 0) {
        var onTimeBadge = onTime ? "<span style='color:#16a34a; font-weight:bold;'>✅ On-Time (+10 Pts)</span>" : "<span style='color:#d97706; font-weight:bold;'>⚠️ Completed Late (+5 Pts)</span>";
        var photoBox = fileUrl ? 
          ("<div style='margin-top:14px; padding:12px; background:#f8fafc; border-radius:8px; border:1px solid #e2e8f0; text-align:center;'>" +
           "<a href='" + fileUrl + "' target='_blank' style='display:inline-block; padding:9px 18px; background:#2563eb; color:#ffffff; text-decoration:none; border-radius:6px; font-weight:bold; font-size:13px;'>📷 View Uploaded Proof Photo</a>" +
           "</div>") : 
          ("<div style='margin-top:10px; color:#64748b; font-size:12px;'>✓ Direct Completion (Checklist verified)</div>");

        var subject = "🔔 Task Completed: " + staffName + " (" + role + ") - " + taskName;
        var htmlBody = "<div style='font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif; max-width:540px; margin:0 auto; background:#ffffff; border:1px solid #e2e8f0; border-radius:12px; overflow:hidden; box-shadow:0 4px 6px -1px rgba(0,0,0,0.05);'>" +
          "<div style='background:#0f172a; padding:16px 20px; color:#ffffff;'>" +
            "<div style='font-size:11px; text-transform:uppercase; letter-spacing:1px; color:#94a3b8; font-weight:bold;'>Yash Marketing ERP • Task Completion Alert</div>" +
            "<div style='font-size:18px; font-weight:bold; margin-top:4px;'>Task Completed By Staff ✅</div>" +
          "</div>" +
          "<div style='padding:20px;'>" +
            "<div style='margin-bottom:14px; padding-bottom:12px; border-bottom:1px solid #f1f5f9;'>" +
              "<div style='font-size:16px; font-weight:bold; color:#0f172a;'>👤 " + staffName + " <span style='font-size:12px; color:#64748b; font-weight:normal;'>(" + role + ")</span></div>" +
            "</div>" +
            "<table style='width:100%; border-collapse:collapse; font-size:13px; line-height:1.7;'>" +
              "<tr><td style='color:#64748b; padding:4px 0; width:110px;'>📋 Task:</td><td style='color:#0f172a; font-weight:bold; padding:4px 0;'>" + taskName + (subTask ? " <span style='font-weight:normal; color:#475569;'>(" + subTask + ")</span>" : "") + "</td></tr>" +
              "<tr><td style='color:#64748b; padding:4px 0;'>🏷️ Type:</td><td style='color:#0f172a; padding:4px 0;'>" + (taskType || "Daily") + "</td></tr>" +
              "<tr><td style='color:#64748b; padding:4px 0;'>⏰ Deadline:</td><td style='color:#0f172a; padding:4px 0;'>" + (deadline || "Today") + "</td></tr>" +
              "<tr><td style='color:#64748b; padding:4px 0;'>⏱️ Submit Time:</td><td style='color:#0f172a; font-weight:bold; padding:4px 0;'>" + submitTimeString + "</td></tr>" +
              "<tr><td style='color:#64748b; padding:4px 0;'>🏆 Status:</td><td style='padding:4px 0;'>" + onTimeBadge + "</td></tr>" +
            "</table>" +
            photoBox +
            "<div style='margin-top:20px; text-align:center; border-top:1px solid #f1f5f9; padding-top:14px;'>" +
              "<a href='https://yash-marketing-master-app.vercel.app/' target='_blank' style='display:inline-block; padding:10px 20px; background:#0f172a; color:#ffffff; text-decoration:none; border-radius:8px; font-weight:bold; font-size:13px;'>Open CEO Operations Radar 🚀</a>" +
            "</div>" +
          "</div>" +
        "</div>";

        MailApp.sendEmail({
          to: recipientEmails.join(","),
          subject: subject,
          htmlBody: htmlBody
        });
      }
    } catch(mailErr) {
      Logger.log("Email notification error (non-fatal): " + mailErr);
    }
  }

  return {
    success: true,
    submitTime: submitTimeString,
    status: finalStatusDisplay,
    pointsEarned: pointsEarned,
    isOnTime: onTime,
    taskName: taskName,
    staffName: staffName,
    fileUrl: fileUrl,
    ceoPhone: ceo ? ceo.phone : "",
    ceoName: ceo ? ceo.name : "Yash Barlota"
  };
}

function setupPermissions() { DriveApp.getRootFolder(); }

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("🚀 Task App Control")
      .addItem("📋 Setup / Refresh All Staff Task Sheets", "setupAllStaffTaskSheets")
      .addToUi();
  } catch(e) {}
}

function setupAllStaffTaskSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName("staff master");
  if (!master) return "staff master sheet nahi mili!";
  var data = master.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    var name = data[i][1] ? data[i][1].toString().trim() : "";
    var role = data[i][5] ? data[i][5].toString().trim() : (data[i][4] ? data[i][4].toString().trim() : "Staff");
    if (name && role.toLowerCase() !== "owner" && role.toLowerCase() !== "ceo" && role.toLowerCase() !== "admin") {
      setupStaffTaskSheet(name, role);
      count++;
    }
  }
  return count + " staff sheets setup / verified successfully!";
}


// =====================================================================
// 📌 CEO EXECUTIVE OPERATIONS RADAR (LIVE PULSE & COMPLIANCE COMMAND)
// =====================================================================
function getCeoOperationsRadar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("staff master");
  if (!masterSheet) {
    return {
      pulse: { totalStaff: 0, totalTasksDue: 0, totalCompleted: 0, totalOverdue: 0, totalPending: 0, executionRate: 0 },
      staffList: []
    };
  }

  var mData = masterSheet.getDataRange().getValues();
  var activeStaff = [];

  for (var i = 1; i < mData.length; i++) {
    var name = mData[i][1] ? mData[i][1].toString().trim() : "";
    var role = mData[i][5] ? mData[i][5].toString().trim() : (mData[i][4] ? mData[i][4].toString().trim() : "Staff");
    var phone = mData[i][2] ? mData[i][2].toString().trim() : "";
    var rLower = role.toLowerCase();
    if (name && rLower !== "owner" && rLower !== "ceo" && rLower !== "admin") {
      activeStaff.push({ name: name, role: role, phone: phone });
    }
  }

  var totalTasksDue = 0;
  var totalCompleted = 0;
  var totalOverdue = 0;
  var totalPending = 0;
  var staffResults = [];

  var now = new Date();
  var nowMinutes = now.getHours() * 60 + now.getMinutes();

  for (var s = 0; s < activeStaff.length; s++) {
    var member = activeStaff[s];
    var tasks = [];
    try {
      tasks = getTasksForRole(member.role, member.name);
    } catch(e) {
      tasks = [];
    }

    var dailyTasks = [];
    var sCompleted = 0;
    var sOverdue = 0;
    var sPending = 0;

    for (var t = 0; t < tasks.length; t++) {
      var item = tasks[t];
      if (item.type && item.type !== "Daily" && item.type !== "Weekly" && item.type !== "Monthly") continue;

      var isDone = item.status && (item.status.indexOf("Uploaded") !== -1 || item.status.indexOf("Completed") !== -1 || item.status.indexOf("Skipped") !== -1);
      var isOverdue = false;
      var overdueText = "";

      if (!isDone) {
        var match = (item.deadline || "").match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
        if (match) {
          var h = parseInt(match[1], 10);
          var m = parseInt(match[2], 10);
          var ampm = match[3].toUpperCase();
          if (ampm === "PM" && h < 12) h += 12;
          if (ampm === "AM" && h === 12) h = 0;
          var deadMins = h * 60 + m;
          if (nowMinutes > deadMins) {
            isOverdue = true;
            var diff = nowMinutes - deadMins;
            var hrs = Math.floor(diff / 60);
            var mins = diff % 60;
            overdueText = (hrs > 0 ? hrs + "h " : "") + mins + "m late";
          }
        }
      }

      if (isDone) {
        sCompleted++;
      } else {
        sPending++;
        if (isOverdue) sOverdue++;
      }

      dailyTasks.push({
        task: item.task,
        subTask: item.subTask,
        deadline: item.deadline,
        timeSlot: item.timeSlot,
        points: item.points,
        status: item.status,
        isDone: isDone,
        isOverdue: isOverdue,
        overdueText: overdueText,
        sourceSheet: item.sourceSheet
      });
    }

    totalTasksDue += dailyTasks.length;
    totalCompleted += sCompleted;
    totalOverdue += sOverdue;
    totalPending += sPending;

    var sRate = dailyTasks.length > 0 ? Math.round((sCompleted / dailyTasks.length) * 100) : 100;

    staffResults.push({
      name: member.name,
      role: member.role,
      phone: member.phone,
      total: dailyTasks.length,
      completed: sCompleted,
      overdue: sOverdue,
      pending: sPending,
      rate: sRate,
      tasks: dailyTasks
    });
  }

  var overallRate = totalTasksDue > 0 ? Math.round((totalCompleted / totalTasksDue) * 100) : 100;

  // 📦 INJECT LIVE ORDER PROCESSING RADAR
  var orderData = getOrdersForRole("CEO", "");
  var allOrders = orderData.orders || [];
  var activeOrders = allOrders.filter(function(o){ return !o.isFullyCompleted; });
  var completedOrders = allOrders.filter(function(o){ return o.isFullyCompleted; });
  var overdueOrders = activeOrders.filter(function(o){ return o.isOverdue; });

  var stageBreakdown = {
    balanceCheck: activeOrders.filter(function(o){ return o.currentStageNum === 1; }).length,
    confirmation: activeOrders.filter(function(o){ return o.currentStageNum === 2; }).length,
    stockPicking: activeOrders.filter(function(o){ return o.currentStageNum === 3; }).length,
    invoicing: activeOrders.filter(function(o){ return o.currentStageNum === 4; }).length,
    packingDispatch: activeOrders.filter(function(o){ return o.currentStageNum === 5; }).length,
    logistics: activeOrders.filter(function(o){ return o.currentStageNum === 6; }).length,
    deliveryComms: activeOrders.filter(function(o){ return o.currentStageNum === 7; }).length,
    delivered: activeOrders.filter(function(o){ return o.currentStageNum === 8; }).length,
    finalRating: activeOrders.filter(function(o){ return o.currentStageNum === 9; }).length
  };

  return {
    pulse: {
      totalStaff: activeStaff.length,
      totalTasksDue: totalTasksDue,
      totalCompleted: totalCompleted,
      totalOverdue: totalOverdue,
      totalPending: totalPending,
      executionRate: overallRate
    },
    staffList: staffResults,
    orderPipeline: {
      totalActive: activeOrders.length,
      totalCompletedToday: completedOrders.length,
      totalOverdue: overdueOrders.length,
      cutoff: orderData.cutoff,
      stageBreakdown: stageBreakdown,
      activeOrders: activeOrders.slice(0, 30)
    },
    recentCompletions: getCeoCompletedTasksFeed(50),
    ceoContact: getCeoContactInfo()
  };
}

// =====================================================================
// 📌 SAVE SHIFT HANDOVER REPORT (AUTO-CALCULATED DAY-END AUDIT LOG)
// =====================================================================
function saveShiftHandover(staffName, role, completedCount, totalCount, pendingTasks, notes) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Shift_Handovers");
  if (!sheet) {
    sheet = ss.insertSheet("Shift_Handovers");
    var headers = ["Timestamp", "Staff Name", "Role", "Completed Tasks", "Total Tasks", "Completion %", "Pending / Carried Forward", "Handover Notes"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground("#1f6feb").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  var tz = Session.getScriptTimeZone();
  var now = new Date();
  var timeStr = Utilities.formatDate(now, tz, "yyyy-MM-dd HH:mm:ss");
  var rate = totalCount > 0 ? Math.round((completedCount / totalCount) * 100) + "%" : "100%";

  sheet.appendRow([
    timeStr,
    staffName || "",
    role || "",
    completedCount || 0,
    totalCount || 0,
    rate,
    pendingTasks || "None",
    notes || ""
  ]);

  return { success: true, timestamp: timeStr };
}


// =====================================================================
// 📊 POINTS LEDGER & AUDIT BREAKDOWN (ITEMIZED TASK POINTS WITH ON-TIME BONUS)
// =====================================================================
function getStaffPointsLedger(staffName, period) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var subSheet = ss.getSheetByName("Task_Submissions");
  var role = getStaffRoleByName(staffName);
  var periodClean = (period || 'daily').toLowerCase();

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

  // 1. Get Routine tasks for this staff
  var allRoutineTasks = [];
  try {
    allRoutineTasks = getTasksForRole(role, staffName);
  } catch(e) {}

  var dailyRoutine = allRoutineTasks.filter(function(t) { return t.type === 'Daily'; });
  var weeklyRoutine = allRoutineTasks.filter(function(t) { return t.type === 'Weekly'; });
  var monthlyRoutine = allRoutineTasks.filter(function(t) { return t.type === 'Monthly'; });

  // 2. Scan submissions
  var entries = [];
  var totalEarned = 0;
  var completedCount = 0;
  var onTimeCount = 0;
  var submittedTitlesMap = {};

  if (subSheet && subSheet.getLastRow() > 1) {
    var subData = subSheet.getDataRange().getValues();

    for (var j = subData.length - 1; j >= 1; j--) {
      var rStaff = subData[j][1] ? subData[j][1].toString().trim().toUpperCase() : "";
      if (targetClean !== "" && rStaff !== targetClean && !rStaff.includes(targetClean) && !targetClean.includes(rStaff)) continue;

      var rawDate = subData[j][0];
      var subDate = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
      if (isNaN(subDate.getTime())) continue;

      var subYear = subDate.getFullYear();
      var subMonth = subDate.getMonth();
      var subDateStr = Utilities.formatDate(subDate, tz, "yyyy-MM-dd");
      var subQuarter = Math.floor(subMonth / 3);

      var isMatchPeriod = false;
      if (periodClean === 'daily' && subDateStr === todayStr) isMatchPeriod = true;
      else if (periodClean === 'weekly' && subDate >= startOfWeek) isMatchPeriod = true;
      else if (periodClean === 'monthly' && subYear === currYear && subMonth === currMonth) isMatchPeriod = true;
      else if (periodClean === 'lastmonth' && subYear === lastMonthYear && subMonth === lastMonth) isMatchPeriod = true;
      else if (periodClean === 'currentquarter' && subYear === currYear && subQuarter === currQuarter) isMatchPeriod = true;
      else if (periodClean === 'lastquarter' && subYear === lastQuarterYear && subQuarter === lastQuarter) isMatchPeriod = true;
      else if (periodClean === 'yearly' && subYear === currYear) isMatchPeriod = true;

      if (!isMatchPeriod) continue;

      var pts = parseFloat(subData[j][10]) || 0;
      var tName = subData[j][4] ? subData[j][4].toString().trim() : "";
      var subTName = subData[j][5] ? subData[j][5].toString().trim() : "";
      var fullTitle = tName + (subTName ? " - " + subTName : "");
      var deadline = subData[j][6] ? subData[j][6].toString().trim() : "";
      var submitTime = subData[j][7] ? subData[j][7].toString().trim() : Utilities.formatDate(subDate, tz, "hh:mm a");
      var status = subData[j][8] ? subData[j][8].toString().trim() : "";
      var proofLink = subData[j][9] ? subData[j][9].toString().trim() : "";

      var compPts = pts >= 5 ? 5 : 0;
      var onTimePts = pts === 10 ? 5 : 0;
      var isOnTime = (pts === 10);

      totalEarned += pts;
      if (compPts > 0) completedCount++;
      if (onTimePts > 0) onTimeCount++;

      submittedTitlesMap[fullTitle.toLowerCase()] = true;
      submittedTitlesMap[tName.toLowerCase()] = true;

      entries.push({
        task: tName,
        subTask: subTName,
        fullTitle: fullTitle,
        deadline: deadline || "N/A",
        submitTime: submitTime,
        dateFormatted: Utilities.formatDate(subDate, tz, "dd MMM, hh:mm a"),
        completionPoints: compPts,
        onTimePoints: onTimePts,
        totalPoints: pts,
        isOnTime: isOnTime,
        status: status,
        proofLink: proofLink,
        isPending: false
      });
    }
  }

  // 3. If Daily, append any scheduled daily tasks that are still pending
  if (periodClean === 'daily') {
    for (var d = 0; d < dailyRoutine.length; d++) {
      var item = dailyRoutine[d];
      var checkKey1 = (item.task + (item.subTask ? " - " + item.subTask : "")).toLowerCase();
      var checkKey2 = item.task.toLowerCase();

      if (!submittedTitlesMap[checkKey1] && !submittedTitlesMap[checkKey2]) {
        entries.push({
          task: item.task,
          subTask: item.subTask,
          fullTitle: item.task + (item.subTask ? " - " + item.subTask : ""),
          deadline: item.deadline || "07:00 PM",
          submitTime: "Pending ⏳",
          dateFormatted: "Today",
          completionPoints: 0,
          onTimePoints: 0,
          totalPoints: 0,
          isOnTime: false,
          status: "Pending ⏳",
          proofLink: "",
          isPending: true
        });
      }
    }
  }

  // 4. Calculate Max Possible Points & Percentage
  var expectedCount = 0;
  if (periodClean === 'daily') {
    expectedCount = Math.max(entries.length, dailyRoutine.length, 1);
  } else if (periodClean === 'weekly') {
    var dayOfWeekNum = (dayOfWeek === 0 ? 7 : dayOfWeek);
    expectedCount = Math.max(entries.length, (dailyRoutine.length * dayOfWeekNum) + weeklyRoutine.length, 1);
  } else if (periodClean === 'monthly') {
    var dayOfMonth = now.getDate();
    expectedCount = Math.max(entries.length, (dailyRoutine.length * dayOfMonth) + (weeklyRoutine.length * Math.ceil(dayOfMonth / 7)) + monthlyRoutine.length, 1);
  } else {
    expectedCount = Math.max(entries.length, 1);
  }

  var maxPossiblePoints = expectedCount * 10;
  var percentage = maxPossiblePoints > 0 ? Math.min(100, Math.round((totalEarned / maxPossiblePoints) * 100)) : 0;

  var labelMap = {
    daily: 'Daily (Today)',
    weekly: 'Weekly (This Week)',
    monthly: 'Monthly (This Month)',
    lastmonth: 'Last Month',
    currentquarter: 'Current Quarter',
    lastquarter: 'Last Quarter',
    yearly: 'Yearly'
  };

  return {
    success: true,
    staffName: staffName,
    role: role,
    period: periodClean,
    periodLabel: labelMap[periodClean] || 'Points Ledger',
    earnedPoints: totalEarned,
    maxPoints: maxPossiblePoints,
    percentage: percentage,
    totalTasks: expectedCount,
    completedCount: completedCount,
    onTimeCount: onTimeCount,
    entries: entries
  };
}

var SPREADSHEET_ID = "1aWuHrdbwVkasyCHgt6z55X7Ob5S0tt127UZ1dKKCiL4";

function getActiveSpreadsheet() {
  var ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch(e) {}
  if (!ss) {
    try {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch(e) {}
  }
  return ss;
}
var DRIVE_FOLDER_ID = "1uj9sY9b74_Tzm-x_Uu6I4NqGCP2harxJ"; 

function onOpen() {
  try {
    getOrCreateDealerLocationGpsSheet();
    getOrCreateLiveTrackingSheet();
    getOrCreateProxyAuditSheet();
  } catch(e) {}
}

function doGet(e) {
  if (e && e.parameter && e.parameter.getStaffRadar) {
    try {
      var trackSheet = getOrCreateLiveTrackingSheet();
      var trackData = trackSheet ? trackSheet.getDataRange().getValues() : [];
      var list = [];
      for (var t = 1; t < trackData.length; t++) {
        if (trackData[t][0]) {
          list.push({
            name: trackData[t][0],
            lat: trackData[t][1],
            lng: trackData[t][2],
            accuracy: trackData[t][3],
            speed: trackData[t][4],
            distFromHQ: trackData[t][5],
            time: trackData[t][6],
            date: trackData[t][7],
            maps: trackData[t][8],
            status: trackData[t][9]
          });
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true, staff: list })).setMimeType(ContentService.MimeType.JSON);
    } catch(eRadar) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: eRadar.message })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  if (e && e.parameter && e.parameter.diag) {
    try {
      var ss = getActiveSpreadsheet();
      var sGps = getSheetSafely("Dealer Location GPS");
      var sMaster = getSheetSafely("Dealer Master");
      var sStaff = getSheetSafely("Live Staff Tracking");
      var sAudit = getSheetSafely("Proxy Audit Log");
      return ContentService.createTextOutput(JSON.stringify({
        gpsRows: sGps ? sGps.getLastRow() : 0,
        masterRows: sMaster ? sMaster.getLastRow() : 0,
        dealersInArea: getDealersByArea(e.parameter.area || "Sawedi + Balikashram Road"),
        allAreas: getAreas()
      })).setMimeType(ContentService.MimeType.JSON);
    } catch(errD) {
      return ContentService.createTextOutput(JSON.stringify({ error: errD.message })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  if (e && e.parameter && e.parameter.initTabs) {
    try {
      var s1 = getOrCreateDealerLocationGpsSheet();
      var s2 = getOrCreateLiveTrackingSheet();
      var s3 = getOrCreateProxyAuditSheet();
      var ss = getActiveSpreadsheet();
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        created: [s1.getName(), s2.getName(), s3.getName()],
        allSheets: ss.getSheets().map(function(s) { return s.getName(); })
      })).setMimeType(ContentService.MimeType.JSON);
    } catch(errInit) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: errInit.message
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  try { getOrCreateDealerLocationGpsSheet(); getOrCreateLiveTrackingSheet(); } catch(eInit) {}

  var template = HtmlService.createTemplateFromFile('Index');
  
  if (e && e.parameter && (e.parameter.empId || e.parameter.role)) {
    template.empId = e.parameter.empId || "";
    template.role = e.parameter.role || "";
    template.name = e.parameter.name || "";
  } else {
    template.empId = "";
    template.role = "";
    template.name = "";
  }
  
    // PRE-EMBED ALL DEALERS DIRECTLY INTO HTML TEMPLATE FOR 0-SECOND CLIENT SWITCH
  try {
    template.allDealersJson = JSON.stringify(getAllDealersGroupedByArea());
  } catch(errGroup) {
    template.allDealersJson = JSON.stringify({ areas: [], grouped: {} });
  }

  return template.evaluate()
    .setTitle('Yash Marketing App')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function testDriveUpload() {
  try {
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    folder.createFile("TestFile.txt", "Agar ye file folder me ban gayi, toh Drive bilkul theek hai!");
    Logger.log("✅ SUCCESS! Folder me file ban gayi. Permissions aur Folder ID dono sahi hain.");
  } catch (e) {
    Logger.log("❌ EXACT ERROR: " + e.message);
  }
}

function getSheetSafely(targetName) {
  var ss = getActiveSpreadsheet();
  if (!ss) return null;
  var sheets = ss.getSheets();
  
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().trim().toLowerCase() === targetName.trim().toLowerCase()) {
      return sheets[i];
    }
  }
  
  return null; 
}

function verifyLogin(empId, password) {
  var sheet = getSheetSafely("Staff Master");
  
  if (!sheet) {
    return { success: false, message: "Staff Master sheet missing!" };
  }
  
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == empId && (data[i][4] == password || password === "SSO_BYPASS")) {
      return { 
        success: true, 
        name: String(data[i][1]).trim(), 
        empId: data[i][0],
        role: String(data[i][5]).trim() 
      };
    }
  }
  
  return { success: false, message: "Invalid Employee ID or Password!" };
}

// -------------------------------------------------------------
// DYNAMIC INCENTIVE ENGINE
// -------------------------------------------------------------
function getIncentiveRules() {
  var rules = { 
    visitRate: 0, 
    pcRate: 0, 
    thousandRate: 0, 
    brandRates: {} 
  };
  
  var sheet = getSheetSafely("Incentive Master");
  
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    
    if (data.length > 1) {
      var headers = data[0];
      var colVisit = -1;
      var colBrand = -1;
      var colFocusRate = -1;
      var colGenRate = -1;
      var colThousand = -1;
      
      for (var j = 0; j < headers.length; j++) {
        var h = String(headers[j]).toLowerCase().replace(/\s+/g, '');
        
        if (h.includes("pervisit")) {
          colVisit = j;
        }
        if (h.includes("focusedbrand") && !h.includes("perpc")) {
          colBrand = j;
        }
        if (h.includes("focusedbrandperpc")) {
          colFocusRate = j;
        }
        if (h.includes("generalperpc")) {
          colGenRate = j;
        }
        if (h.includes("per1000")) {
          colThousand = j;
        }
      }

      if (colVisit !== -1) {
        rules.visitRate = parseFloat(data[1][colVisit]) || 0;
      }
      
      if (colGenRate !== -1) {
        rules.pcRate = parseFloat(data[1][colGenRate]) || 0;
      }
      
      if (colThousand !== -1) {
        rules.thousandRate = parseFloat(data[1][colThousand]) || 0;
      }

      if (colBrand !== -1 && colFocusRate !== -1) {
        for (var i = 1; i < data.length; i++) {
          var bName = String(data[i][colBrand]).trim().toUpperCase();
          var bRate = parseFloat(data[i][colFocusRate]) || 0;
          
          if (bName && bName !== "") {
            rules.brandRates[bName] = bRate;
          }
        }
      }
    }
  }
  
  return rules;
}

// -------------------------------------------------------------
// EXTERNAL ORDER CONTROL TOWER & ULTIMATE ATTRIBUTION
// -------------------------------------------------------------
function formatTimeSafe(raw) {
  if (!raw) return "";
  if (raw instanceof Date) {
    return Utilities.formatDate(raw, Session.getScriptTimeZone(), "hh:mm a");
  }
  var s = String(raw).trim();
  if (s.indexOf("GMT") !== -1 || s.indexOf("1899") !== -1) {
    try {
      var d = new Date(s);
      if (!isNaN(d.getTime())) {
        return Utilities.formatDate(d, Session.getScriptTimeZone(), "hh:mm a");
      }
    } catch(e) {}
  }
  if (s.indexOf(":") !== -1 && (s.indexOf("AM") !== -1 || s.indexOf("PM") !== -1)) {
    return s;
  }
  if (s.indexOf(":") !== -1) {
    var parts = s.split(":");
    var h = parseInt(parts[0], 10);
    var m = parseInt(parts[1], 10);
    if (!isNaN(h) && !isNaN(m)) {
      var ampm = h >= 12 ? "PM" : "AM";
      var h12 = h % 12;
      if (h12 === 0) h12 = 12;
      var mStr = m < 10 ? "0" + m : "" + m;
      var hStr = h12 < 10 ? "0" + h12 : "" + h12;
      return hStr + ":" + mStr + " " + ampm;
    }
  }
  return s;
}

function extractDateSafe(raw) {
  if (!raw) {
    return null;
  }
  
  if (raw instanceof Date) {
    var d = new Date(raw);
    d.setHours(0, 0, 0, 0);
    return d;
  }
  
  var s = String(raw).trim().split(" ")[0];
  
  if (s.includes("/")) {
    var p = s.split("/");
    if (p.length === 3) {
      return new Date(p[2], p[1] - 1, p[0], 0, 0, 0, 0);
    }
  } else if (s.includes("-")) {
    var d2 = new Date(s);
    d2.setHours(0, 0, 0, 0);
    return d2;
  }
  
  return null;
}

function fetchAndMapOrders() {
  var areaToStaff = {};
  var sheetsToScan = [];
  
  var localStaffSheet = getSheetSafely("Staff Master");
  if (localStaffSheet) {
    sheetsToScan.push(localStaffSheet);
  }
  
  var cachedExtApp = null;
  try {
    var externalId = '1np1XfpYjqXcbbNwT7_Q3XzmxxGrcYKacXQV04p1ilYw';
    cachedExtApp = SpreadsheetApp.openById(externalId);
    var extStaff = cachedExtApp.getSheetByName('Staff');
    if (extStaff) {
      sheetsToScan.push(extStaff);
    }
  } catch(e) {
    Logger.log("External Staff Fetch Error: " + e.message);
  }

  for (var s = 0; s < sheetsToScan.length; s++) {
    var staffSheet = sheetsToScan[s];
    if (staffSheet) {
      var staffData = staffSheet.getDataRange().getValues();
      for (var i = 1; i < staffData.length; i++) {
        var staffName = String(staffData[i][1]).trim(); 
        var areaStr = String(staffData[i][8] || "").trim(); 
        if (staffName && areaStr) {
          var areas = areaStr.split(',');
          for (var k = 0; k < areas.length; k++) {
            var cleanArea = areas[k].trim().toUpperCase().replace(/\s+/g, ' ');
            if (cleanArea && cleanArea !== "#N/A" && cleanArea !== "UNMAPPED") {
              areaToStaff[cleanArea] = staffName;
            }
          }
        }
      }
    }
  }

  var dealerToArea = {};
  var dealerSheet = getSheetSafely("Dealer Master");
  if (dealerSheet) {
    var dData = dealerSheet.getDataRange().getValues();
    for (var i = 1; i < dData.length; i++) {
      var dName = String(dData[i][1]).trim().toUpperCase().replace(/\s+/g, ' ');
      var dArea = String(dData[i][2]).trim().toUpperCase().replace(/\s+/g, ' ');
      if (dName && dArea && dArea !== "#N/A" && dArea !== "UNMAPPED") {
        dealerToArea[dName] = dArea;
      }
    }
  }

  var orderList = [];
  
  try {
    var extSpreadsheet = cachedExtApp || SpreadsheetApp.openById('1np1XfpYjqXcbbNwT7_Q3XzmxxGrcYKacXQV04p1ilYw');
    var orderSheet = extSpreadsheet.getSheetByName('Orders') || extSpreadsheet.getSheets()[0];
    var oData = orderSheet.getDataRange().getValues();
    
    if (oData.length <= 1) {
      return orderList;
    }

    var head = oData[0];
    var dateCol = 0;
    var shopCol = -1;
    var tCol = -1;
    var itemCol = -1;
    
    for (var c = 0; c < head.length; c++) {
      var hName = String(head[c]).toLowerCase().replace(/\s+/g, '');
      if (hName === "shopname") shopCol = c;
      if (hName.includes("total") || hName.includes("amount")) tCol = c;
      if (hName.includes("item") || hName.includes("product") || hName.includes("orderitems")) itemCol = c;
    }

    for (var j = 1; j < oData.length; j++) {
       var dObj = extractDateSafe(oData[j][dateCol]);
       
       if (dObj) {
         var shopNameRaw = shopCol !== -1 ? String(oData[j][shopCol]).trim() : "";
         var shopNameUp = shopNameRaw.toUpperCase().replace(/\s+/g, ' ');
         var amt = tCol !== -1 ? parseFloat(oData[j][tCol]) || 0 : 0;
         var itemsStr = itemCol !== -1 ? String(oData[j][itemCol]) : "";

         var mappedArea = dealerToArea[shopNameUp] || "Unmapped / No Area";
         var assignedStaff = "Unassigned";
         
         // STRICT AREA ATTRIBUTION:
         // An order is credited to a staff member ONLY IF:
         // 1. The dealer has a valid mapped area (not #N/A, not empty, not Unmapped)
         // 2. That exact area is assigned to that staff member in Staff Master
         if (mappedArea && mappedArea !== "Unmapped / No Area" && mappedArea !== "#N/A" && !mappedArea.startsWith("#")) {
           if (areaToStaff[mappedArea]) {
             assignedStaff = areaToStaff[mappedArea];
           } else {
             for (var aKey in areaToStaff) {
               if (aKey.length >= 5 && (mappedArea === aKey || mappedArea.indexOf(aKey) !== -1 || aKey.indexOf(mappedArea) !== -1)) {
                 if (aKey !== "WHOLESALE" && aKey !== "DEALER") {
                   assignedStaff = areaToStaff[aKey];
                   break;
                 }
               }
             }
           }
         }

         orderList.push({
           date: dObj,
           timeMs: dObj.getTime(),
           shop: shopNameRaw.toUpperCase(),
           area: mappedArea,
           amount: amt,
           items: itemsStr,
           staff: assignedStaff
         });
       }
    }
  } catch (e) {
    Logger.log("External Order Fetch Error: " + e.message);
  }
  
  return orderList;
}

// -------------------------------------------------------------
// SLEEPING GIANTS RETENTION ENGINE
// -------------------------------------------------------------
function getCEOStatsForRange(startTime, endTime, vData, allOrders) {
  var stats = { totalVisits: 0, totalCollections: 0, totalSales: 0 };
  var staffData = {};
  var allCompanyVisits = [];
  var allCompanyOrders = [];

  function initDealerMap(staff, area, dealer) {
    if (!staffData[staff]) {
      staffData[staff] = { visits: 0, collections: 0, sales: 0, orders: 0, areas: {}, visitsList: [], ordersList: [] };
    }
    if (!staffData[staff].areas[area]) {
      staffData[staff].areas[area] = { visits: 0, sales: 0, dealers: {} };
    }
    if (!staffData[staff].areas[area].dealers[dealer]) {
      staffData[staff].areas[area].dealers[dealer] = { visits: 0, sales: 0, logs: [], orders: [] };
    }
  }

  // 1. Process visits in range
  if (vData && vData.length > 1) {
    for (var i = 1; i < vData.length; i++) {
      var dObj = extractDateSafe(vData[i][1]);
      if (dObj) {
        var t = dObj.getTime();
        if (t >= startTime && t <= endTime) {
          var staffName = String(vData[i][3]).trim();
          var dealer = String(vData[i][5]).trim().toUpperCase();
          var area = String(vData[i][4]).trim().toUpperCase() || "UNCATEGORIZED AREA";

          initDealerMap(staffName, area, dealer);
          stats.totalVisits++;
          staffData[staffName].visits++;
          staffData[staffName].areas[area].visits++;
          staffData[staffName].areas[area].dealers[dealer].visits++;

          var task = vData[i][7];
          if (task === "Payment Collection") {
            stats.totalCollections++;
            staffData[staffName].collections++;
          }

          var rowTime = formatTimeSafe(vData[i][2]);
          var dateStr = Utilities.formatDate(dObj, Session.getScriptTimeZone(), "dd/MM/yyyy");
          var locationStr = String(vData[i][11] || "");

          var visitItem = {
            id: String(vData[i][0] || ("V" + i)),
            dateStr: dateStr,
            timeStr: rowTime,
            staff: staffName,
            area: area,
            dealer: dealer,
            task: task,
            location: locationStr
          };
          allCompanyVisits.push(visitItem);
          staffData[staffName].visitsList.push(visitItem);

          var logEntry = {
            date: dateStr, 
            time: rowTime, 
            staff: staffName, 
            area: area, 
            dealer: dealer, 
            task: task,
            feedback: String(vData[i][40] || ""), 
            photo: String(vData[i][41] || "")
          };
          staffData[staffName].areas[area].dealers[dealer].logs.push(logEntry);
        }
      }
    }
  }

  // 2. Process orders in range
  if (allOrders && allOrders.length > 0) {
    for (var j = 0; j < allOrders.length; j++) {
      var ord = allOrders[j];
      if (ord.timeMs >= startTime && ord.timeMs <= endTime) {
        var creditedStaff = ord.staff;
        var creditedArea = ord.area;
        var creditedShop = ord.shop;

        if (!creditedStaff || creditedStaff === "Unassigned" || creditedStaff === "Unmapped") {
          creditedStaff = "Unmapped";
          creditedArea = "UNMAPPED LOCATIONS";
        }

        initDealerMap(creditedStaff, creditedArea, creditedShop);

        var orderDateStr = Utilities.formatDate(ord.date, Session.getScriptTimeZone(), "dd/MM/yyyy");
        var orderItem = {
          shop: creditedShop,
          area: creditedArea,
          amount: ord.amount,
          items: ord.items,
          dateStr: orderDateStr,
          timeMs: ord.timeMs,
          staff: creditedStaff
        };
        allCompanyOrders.push(orderItem);
        staffData[creditedStaff].ordersList.push(orderItem);

        staffData[creditedStaff].areas[creditedArea].dealers[creditedShop].orders.push({
          date: orderDateStr,
          amount: ord.amount,
          items: ord.items
        });
      }
    }
  }

  // 3. Compute Metrics for each staff member
  var leaderboard = [];
  for (var name in staffData) {
    if (name === "") continue;
    var sd = staffData[name];
    var staffMetricsObj = computeEntityMetrics(sd.visitsList, sd.ordersList);
    leaderboard.push({
      name: name,
      visits: sd.visits,
      collections: sd.collections,
      sales: staffMetricsObj.metrics.revenue,
      orders: staffMetricsObj.metrics.transactions,
      strike: staffMetricsObj.metrics.conversionRatio,
      metrics: staffMetricsObj.metrics,
      dealerBreakdown: staffMetricsObj.dealerBreakdown,
      visitsList: staffMetricsObj.visitsList,
      ordersList: staffMetricsObj.ordersList,
      unvisitedOrdersList: staffMetricsObj.unvisitedOrdersList,
      areas: sd.areas
    });
  }

  leaderboard.sort(function(a, b) {
    if (b.sales !== a.sales) return b.sales - a.sales;
    return b.visits - a.visits;
  });

  var companyMetricsObj = computeEntityMetrics(allCompanyVisits, allCompanyOrders);

  stats.totalSales = companyMetricsObj.metrics.revenue;
  stats.directSales = companyMetricsObj.metrics.unvisitedRevenue;
  stats.totalAllSales = companyMetricsObj.metrics.revenue + companyMetricsObj.metrics.unvisitedRevenue;

  return {
    metrics: stats,
    companyMetrics: companyMetricsObj.metrics,
    companyDealerBreakdown: companyMetricsObj.dealerBreakdown,
    companyVisitsList: companyMetricsObj.visitsList,
    companyOrdersList: companyMetricsObj.ordersList,
    companyUnvisitedOrdersList: companyMetricsObj.unvisitedOrdersList,
    leaderboard: leaderboard
  };
}

function getCEOAllBuckets() {
  var vData = [];
  var visitSheet = getSheetSafely("Daily Operation");
  if (visitSheet) {
    vData = visitSheet.getDataRange().getValues();
  }
  var allOrders = fetchAndMapOrders();

  var now = new Date();
  now.setHours(0, 0, 0, 0);
  
  var todayTime = now.getTime();
  var todayEndTime = todayTime + 86399999;
  
  var yesterdayTime = new Date(now).setDate(now.getDate() - 1);
  var yesterdayEndTime = yesterdayTime + 86399999;
  
  var dayOfWeek = now.getDay();
  var diffToMonday = now.getDate() - dayOfWeek + (dayOfWeek === 0 ? -6 : 1);
  var startOfThisWeek = new Date(now).setDate(diffToMonday);
  
  var startOfLastWeek = new Date(startOfThisWeek).setDate(new Date(startOfThisWeek).getDate() - 7);
  var endOfLastWeek = new Date(startOfThisWeek).setDate(new Date(startOfThisWeek).getDate() - 1) + 86399999;
  
  var startOfThisMonth = new Date(now.getFullYear(), now.getMonth(), 1).getTime();
  var endOfThisMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59, 999).getTime();
  
  var startOfLastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1).getTime();
  var endOfLastMonth = new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59, 999).getTime();

  var currentQuarterStartMonth = Math.floor(now.getMonth() / 3) * 3;
  var startOfThisQuarter = new Date(now.getFullYear(), currentQuarterStartMonth, 1).getTime();
  var endOfThisQuarter = new Date(now.getFullYear(), currentQuarterStartMonth + 3, 0, 23, 59, 59, 999).getTime();

  var startOfLastQuarter = new Date(now.getFullYear(), currentQuarterStartMonth - 3, 1).getTime();
  var endOfLastQuarter = new Date(now.getFullYear(), currentQuarterStartMonth, 0, 23, 59, 59, 999).getTime();

  var startOfThisYear = new Date(now.getFullYear(), 0, 1).getTime();
  var endOfThisYear = new Date(now.getFullYear(), 11, 31, 23, 59, 59, 999).getTime();

  var ranges = {
    today: [todayTime, todayEndTime],
    yesterday: [yesterdayTime, yesterdayEndTime],
    thisWeek: [startOfThisWeek, todayEndTime],
    lastWeek: [startOfLastWeek, endOfLastWeek],
    thisMonth: [startOfThisMonth, endOfThisMonth],
    lastMonth: [startOfLastMonth, endOfLastMonth],
    thisQuarter: [startOfThisQuarter, endOfThisQuarter],
    lastQuarter: [startOfLastQuarter, endOfLastQuarter],
    thisYear: [startOfThisYear, endOfThisYear]
  };

  var buckets = {};
  for (var key in ranges) {
    buckets[key] = getCEOStatsForRange(ranges[key][0], ranges[key][1], vData, allOrders);
  }

  return {
    success: true,
    allBuckets: true,
    buckets: buckets
  };
}

function getCEOStats(filter, customStartStr, customEndStr) {
  if (filter === 'custom' && customStartStr && customEndStr) {
    var startTime = new Date(customStartStr).setHours(0, 0, 0, 0);
    var endTime = new Date(customEndStr).setHours(23, 59, 59, 999);
    var vData = [];
    var visitSheet = getSheetSafely("Daily Operation");
    if (visitSheet) {
      vData = visitSheet.getDataRange().getValues();
    }
    var allOrders = fetchAndMapOrders();
    var res = getCEOStatsForRange(startTime, endTime, vData, allOrders);
    return res;
  }

  // For standard timeframe queries, return all buckets for 0-second instant client switching
  return getCEOAllBuckets();
}

function getCEOStats(filter, customStartStr, customEndStr) {
  var now = new Date();
  now.setHours(0, 0, 0, 0);
  
  var todayTime = now.getTime();
  var todayEndTime = todayTime + 86399999;
  
  var yesterdayTime = new Date(now).setDate(now.getDate() - 1);
  var yesterdayEndTime = yesterdayTime + 86399999;
  
  var dayOfWeek = now.getDay();
  var diffToMonday = now.getDate() - dayOfWeek + (dayOfWeek === 0 ? -6 : 1);
  var startOfThisWeek = new Date(now).setDate(diffToMonday);
  
  var startOfLastWeek = new Date(startOfThisWeek).setDate(new Date(startOfThisWeek).getDate() - 7);
  var endOfLastWeek = new Date(startOfThisWeek).setDate(new Date(startOfThisWeek).getDate() - 1) + 86399999;
  
  var startOfThisMonth = new Date(now.getFullYear(), now.getMonth(), 1).getTime();
  var endOfThisMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59, 999).getTime();
  
  var startOfLastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1).getTime();
  var endOfLastMonth = new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59, 999).getTime();

  var currentQuarterStartMonth = Math.floor(now.getMonth() / 3) * 3;
  var startOfThisQuarter = new Date(now.getFullYear(), currentQuarterStartMonth, 1).getTime();
  var endOfThisQuarter = new Date(now.getFullYear(), currentQuarterStartMonth + 3, 0, 23, 59, 59, 999).getTime();

  var startOfLastQuarter = new Date(now.getFullYear(), currentQuarterStartMonth - 3, 1).getTime();
  var endOfLastQuarter = new Date(now.getFullYear(), currentQuarterStartMonth, 0, 23, 59, 59, 999).getTime();

  var startOfThisYear = new Date(now.getFullYear(), 0, 1).getTime();
  var endOfThisYear = new Date(now.getFullYear(), 11, 31, 23, 59, 59, 999).getTime();

  var startTime;
  var endTime;
  
  if (filter === 'today') { 
    startTime = todayTime; endTime = todayEndTime;
  } else if (filter === 'yesterday') { 
    startTime = yesterdayTime; endTime = yesterdayEndTime;
  } else if (filter === 'thisWeek') {
    startTime = startOfThisWeek; endTime = todayEndTime;
  } else if (filter === 'lastWeek') {
    startTime = startOfLastWeek; endTime = endOfLastWeek;
  } else if (filter === 'thisMonth') {
    startTime = startOfThisMonth; endTime = endOfThisMonth;
  } else if (filter === 'lastMonth') {
    startTime = startOfLastMonth; endTime = endOfLastMonth;
  } else if (filter === 'thisQuarter') {
    startTime = startOfThisQuarter; endTime = endOfThisQuarter;
  } else if (filter === 'lastQuarter') {
    startTime = startOfLastQuarter; endTime = endOfLastQuarter;
  } else if (filter === 'thisYear') {
    startTime = startOfThisYear; endTime = endOfThisYear;
  } else if (filter === 'custom' && customStartStr && customEndStr) {
    startTime = new Date(customStartStr).setHours(0, 0, 0, 0);
    endTime = new Date(customEndStr).setHours(23, 59, 59, 999);
  } else {
    startTime = todayTime; endTime = todayEndTime;
  }

  var stats = { 
    totalVisits: 0, 
    totalCollections: 0, 
    totalSales: 0 
  };
  
  var staffData = {};

  function initDealerMap(staff, area, dealer) {
    if (!staffData[staff]) {
      staffData[staff] = { visits: 0, collections: 0, sales: 0, orders: 0, areas: {}, visitsList: [], ordersList: [] };
    }
    if (!staffData[staff].areas[area]) {
      staffData[staff].areas[area] = { visits: 0, sales: 0, dealers: {} };
    }
    if (!staffData[staff].areas[area].dealers[dealer]) {
      staffData[staff].areas[area].dealers[dealer] = { visits: 0, sales: 0, logs: [], orders: [] };
    }
  }

  var globalLogs = []; 
  var allCompanyVisits = [];
  var allCompanyOrders = [];

  var visitSheet = getSheetSafely("Daily Operation");

  if (visitSheet) {
    var data = visitSheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      var dObj = extractDateSafe(data[i][1]);

      if (dObj) {
        var t = dObj.getTime();
        var staffName = String(data[i][3]).trim();
        var dealer = String(data[i][5]).trim().toUpperCase();
        var area = String(data[i][4]).trim().toUpperCase() || "UNCATEGORIZED AREA";

        if (t >= startTime && t <= endTime) {
          initDealerMap(staffName, area, dealer);
          
          stats.totalVisits++;
          staffData[staffName].visits++;
          staffData[staffName].areas[area].visits++;
          staffData[staffName].areas[area].dealers[dealer].visits++;
          
          var task = data[i][7];
          if (task === "Payment Collection") {
            stats.totalCollections++;
            staffData[staffName].collections++;
          }
          
          var rowTimeRaw = data[i][2];
          var rowTime = formatTimeSafe(rowTimeRaw);
          var dateStr = Utilities.formatDate(dObj, Session.getScriptTimeZone(), "dd/MM/yyyy");
          var locationStr = String(data[i][11] || "");

          var visitItem = {
            id: String(data[i][0] || ("V" + i)),
            dateStr: dateStr,
            timeStr: rowTime,
            staff: staffName,
            area: area,
            dealer: dealer,
            task: task,
            location: locationStr
          };
          allCompanyVisits.push(visitItem);
          staffData[staffName].visitsList.push(visitItem);

          var logEntry = {
            date: dateStr, 
            time: rowTime, 
            staff: staffName, 
            area: area, 
            dealer: dealer, 
            task: task,
            feedback: String(data[i][40] || ""), 
            photo: String(data[i][41] || ""),
            stock: {
              bAirpods: data[i][18], bNeckband: data[i][19], bPowerbank: data[i][20], bSmartwatch: data[i][21], bHeadphones: data[i][22],
              gAirpods: data[i][23], gNeckband: data[i][24], gPowerbank: data[i][25], gSmartwatch: data[i][26], gHeadphones: data[i][27]
            }
          };
          
          staffData[staffName].areas[area].dealers[dealer].logs.push(logEntry);
          globalLogs.push(logEntry);
        }
      }
    }
  }

  var allOrders = fetchAndMapOrders();
  
  for (var o = 0; o < allOrders.length; o++) {
    var ord = allOrders[o];
    
    if (ord.timeMs >= startTime && ord.timeMs <= endTime) {
       var creditedStaff = String(ord.staff).trim();
       var creditedArea = String(ord.area || "UNMAPPED LOCATIONS").trim().toUpperCase();
       var creditedShop = ord.shop || "Unknown Dealer";
       
       if (!creditedStaff || creditedStaff === "Unassigned") {
          creditedStaff = "System / Unmapped";
          creditedArea = "UNMAPPED LOCATIONS";
       }

       initDealerMap(creditedStaff, creditedArea, creditedShop);

       stats.totalSales += ord.amount;
       staffData[creditedStaff].sales += ord.amount;
       staffData[creditedStaff].orders++;
       
       staffData[creditedStaff].areas[creditedArea].sales += ord.amount;
       staffData[creditedStaff].areas[creditedArea].dealers[creditedShop].sales += ord.amount;
       
       var orderDateStr = Utilities.formatDate(ord.date, Session.getScriptTimeZone(), "dd/MM/yyyy");
       var orderItem = {
         shop: creditedShop,
         area: creditedArea,
         amount: ord.amount,
         items: ord.items,
         dateStr: orderDateStr,
         timeMs: ord.timeMs,
         staff: creditedStaff
       };
       allCompanyOrders.push(orderItem);
       staffData[creditedStaff].ordersList.push(orderItem);

       staffData[creditedStaff].areas[creditedArea].dealers[creditedShop].orders.push({
         date: orderDateStr,
         amount: ord.amount,
         items: ord.items
       });
    }
  }

  var leaderboard = [];
  
  for (var name in staffData) {
     if (name === "") continue;
     var sd = staffData[name];
     var staffMetricsObj = computeEntityMetrics(sd.visitsList, sd.ordersList);
     
     leaderboard.push({
        name: name, 
        visits: sd.visits, 
        collections: sd.collections, 
        sales: staffMetricsObj.metrics.revenue,
        orders: staffMetricsObj.metrics.transactions,
        strike: staffMetricsObj.metrics.conversionRatio,
        metrics: staffMetricsObj.metrics,
        dealerBreakdown: staffMetricsObj.dealerBreakdown,
        visitsList: staffMetricsObj.visitsList,
        ordersList: staffMetricsObj.ordersList,
        unvisitedOrdersList: staffMetricsObj.unvisitedOrdersList,
        areas: sd.areas
     });
  }
  
  leaderboard.sort(function(a, b) {
    if (b.sales !== a.sales) return b.sales - a.sales;
    return b.visits - a.visits;
  });

  var companyMetricsObj = computeEntityMetrics(allCompanyVisits, allCompanyOrders);

  return { 
    metrics: stats, 
    companyMetrics: companyMetricsObj.metrics,
    companyDealerBreakdown: companyMetricsObj.dealerBreakdown,
    companyVisitsList: companyMetricsObj.visitsList,
    companyOrdersList: companyMetricsObj.ordersList,
    leaderboard: leaderboard, 
    globalLogs: globalLogs 
  };
}

// ----------------------------------------------------------------------
// EOD WHATSAPP GENERATOR
// ----------------------------------------------------------------------
function generateWhatsAppEOD(staffName) {
  var dash = getStaffDashboardData(staffName);
  var staffNameClean = String(staffName).trim();
  
  var todayVisits = [];
  var visitedDealersToday = [];
  var stockOuts = [];
  var deadStocks = [];
  var firstVisitStr = "--:--";
  var lastVisitStr = "--:--";
  var activeMins = 0;
  
  var now = new Date();
  now.setHours(0, 0, 0, 0);
  var todayTime = now.getTime();
  
  var visitSheet = getSheetSafely("Daily Operation");
  
  if (visitSheet) {
    var vData = visitSheet.getDataRange().getValues();
    var timesArray = [];
    
    for (var i = 1; i < vData.length; i++) {
      if (String(vData[i][3]).trim() === staffNameClean) {
        var dObj = extractDateSafe(vData[i][1]);
        
        if (dObj && dObj.getTime() === todayTime) {
          var dealer = String(vData[i][5]).trim().toUpperCase();
          visitedDealersToday.push(dealer);
          todayVisits.push(vData[i]);
          
          var rawTime = vData[i][2];
          var formattedTime = "";
          
          if (rawTime instanceof Date) {
              formattedTime = Utilities.formatDate(rawTime, Session.getScriptTimeZone(), "HH:mm");
          } else {
              formattedTime = String(rawTime || "").trim();
          }
          
          if (formattedTime) timesArray.push(formattedTime);
          
          var stockCols = [
            { idx: 18, name: "BOAT Airpods" }, 
            { idx: 19, name: "BOAT Neckband" }, 
            { idx: 20, name: "BOAT Powerbank" },
            { idx: 23, name: "GoBoult Airpods" }, 
            { idx: 24, name: "GoBoult Neckband" },
            { idx: 33, name: "Velocity Airpods" }, 
            { idx: 34, name: "Velocity Neckband" }, 
            { idx: 35, name: "Velocity Cable" }
          ];
          
          for (var s = 0; s < stockCols.length; s++) {
            var sc = stockCols[s];
            if (vData[i][sc.idx] === "0 (Nil)") {
               stockOuts.push(dealer + " (" + sc.name + ")");
            }
            if (vData[i][sc.idx] === "10+ (High)") {
               deadStocks.push(dealer + " (" + sc.name + ")");
            }
          }
        }
      }
    }
    
    if (timesArray.length > 0) {
      firstVisitStr = timesArray[0];
      lastVisitStr = timesArray[timesArray.length - 1];
      
      try {
        var fParts = firstVisitStr.split(":");
        var lParts = lastVisitStr.split(":");
        if (fParts.length >= 2 && lParts.length >= 2) {
           var fMins = parseInt(fParts[0], 10) * 60 + parseInt(fParts[1], 10);
           var lMins = parseInt(lParts[0], 10) * 60 + parseInt(lParts[1], 10);
           activeMins = lMins - fMins;
           if (activeMins < 0) activeMins = 0; 
        }
      } catch(e) {}
    }
  }

  var allOrders = fetchAndMapOrders();
  var orderDealersToday = [];
  var rankMap = {}; 

  for (var o = 0; o < allOrders.length; o++) {
    var ord = allOrders[o];
    if (ord.timeMs === todayTime) {
      var ordStaffClean = String(ord.staff).trim();
      if (!rankMap[ordStaffClean]) {
         rankMap[ordStaffClean] = 0;
      }
      rankMap[ordStaffClean] += ord.amount;
      
      if (ordStaffClean === staffNameClean && ord.shop) {
        var clnShop = ord.shop.toUpperCase().replace(/[^A-Z0-9]/g, '');
        orderDealersToday.push(clnShop);
      }
    }
  }
    
  var sortedStaff = Object.keys(rankMap).sort(function(a, b) { return rankMap[b] - rankMap[a]; });
  var myRank = sortedStaff.indexOf(staffNameClean) + 1;
  if (myRank === 0 && dash.today.sales > 0) myRank = 1; 
  
  var zeroSaleDealers = [];
  for (var v = 0; v < visitedDealersToday.length; v++) {
    var d = visitedDealersToday[v];
    var cleanD = d.replace(/[^A-Z0-9]/g, ''); 
    var found = false;
    
    for (var t = 0; t < orderDealersToday.length; t++) {
      var orderD = orderDealersToday[t];
      if (cleanD.indexOf(orderD) !== -1 || orderD.indexOf(cleanD) !== -1) {
        found = true; 
        break;
      }
    }
    if (!found) zeroSaleDealers.push(d);
  }
  
  var productiveCount = todayVisits.length - zeroSaleDealers.length;
  if (productiveCount < 0) productiveCount = 0;
  
  var strikeRate = todayVisits.length > 0 ? Math.round((productiveCount / todayVisits.length) * 100) : 0;
  var actHrs = Math.floor(activeMins / 60);
  var actM = activeMins % 60;
  var avgMinsPerStore = todayVisits.length > 0 ? Math.round(activeMins / todayVisits.length) : 0;
  var moneyLeftOnTable = zeroSaleDealers.length * dash.rules.visitRate; 
  
  var TARGET_WEEK = 150000;
  var TARGET_MONTH = 600000;
  var paceWeek = TARGET_WEEK - dash.thisWeek.sales;
  var paceMonth = TARGET_MONTH - dash.thisMonth.sales;
  
  var focusArr = [];
  for (var b in dash.today.brandPcs) {
      var onlyLetters = b.replace(/[^A-Z]/g, '');
      if (onlyLetters.length > 1) focusArr.push(onlyLetters);
  }
  
  var uniqueFocus = [];
  for (var i = 0; i < focusArr.length; i++) {
      if (uniqueFocus.indexOf(focusArr[i]) === -1) {
         uniqueFocus.push(focusArr[i]);
      }
  }
  var focusBrandsSold = uniqueFocus.length > 0 ? uniqueFocus.join(", ") : "None";

  var nl = "\n";
  var report = "*📊 EXECUTIVE EOD REPORT*" + nl;
  report += "*👤 Agent:* " + staffNameClean + nl;
  report += "*📅 Date:* " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yyyy") + nl + nl;
  
  report += "*🔥 1. CORE CONVERSION*" + nl;
  report += "• Total Visits: *" + todayVisits.length + "*" + nl;
  report += "• Productive: *" + productiveCount + "* (" + strikeRate + "% Strike Rate)" + nl;
  report += "• Total Sales: *₹ " + dash.today.sales.toLocaleString('en-IN') + "*" + nl;
  report += "• Zero-Sale Counters: *" + zeroSaleDealers.length + "*" + nl;
  if (zeroSaleDealers.length > 0 && zeroSaleDealers.length <= 5) {
     report += "  _(" + zeroSaleDealers.join(", ") + ")_" + nl;
  } else if (zeroSaleDealers.length > 5) {
     report += "  _(" + zeroSaleDealers.slice(0, 5).join(", ") + " + " + (zeroSaleDealers.length - 5) + " more)_" + nl;
  }
  report += nl;
  
  report += "*⏱️ 2. TIME & LEVERAGE*" + nl;
  report += "• Field Active: *" + actHrs + "h " + actM + "m*" + nl;
  report += "• First Punch: *" + firstVisitStr + "*" + nl;
  report += "• Last Punch: *" + lastVisitStr + "*" + nl;
  report += "• Avg Time/Store: *" + avgMinsPerStore + " mins*" + nl;
  report += "• Route Efficiency: *" + (avgMinsPerStore > 40 ? "Needs Improvement ⚠️" : "Optimal ✅") + "*" + nl;
  report += nl;
  
  report += "*📦 3. MARKET INTELLIGENCE*" + nl;
  report += "• Stock-Out Alerts (Nil): *" + stockOuts.length + "*" + nl;
  if (stockOuts.length > 0) {
     report += "  _Follow up tomorrow on these._" + nl;
  }
  report += "• Focus Brands Pitched/Sold: *" + focusBrandsSold + "*" + nl;
  report += "• Cross-Selling Index: *" + (uniqueFocus.length >= 2 ? "High 🚀" : "Low 🔻") + "*" + nl;
  report += nl;
  
  report += "*💰 4. GAMIFICATION & INCENTIVE*" + nl;
  report += "• Today's Incentive: *₹ " + dash.today.incentive + "*" + nl;
  report += "• Money Left on Table: *~₹ " + moneyLeftOnTable + "* _(Lost due to zero-sale visits)_" + nl;
  report += "• Daily Office Rank: *#" + (myRank || "-") + "* 🏆" + nl;
  report += nl;
  
  report += "*������ 5. GROWTH TRACKER*" + nl;
  report += "• Today: *₹ " + dash.today.sales.toLocaleString('en-IN') + "*" + nl;
  report += "• Yesterday: *₹ " + dash.yesterday.sales.toLocaleString('en-IN') + "*" + nl;
  report += "• This Week: *₹ " + dash.thisWeek.sales.toLocaleString('en-IN') + "* " + (paceWeek > 0 ? "(_Need ₹" + paceWeek.toLocaleString('en-IN') + " for Target_)" : "(_Target Hit! 🎯_)") + nl;
  report += "• This Month: *₹ " + dash.thisMonth.sales.toLocaleString('en-IN') + "* " + (paceMonth > 0 ? "(_Need ₹" + paceMonth.toLocaleString('en-IN') + " for Target_)" : "(_Target Hit! 🎯_)") + nl;

  return encodeURIComponent(report);
}

function getAreas() {
  var sheet = getSheetSafely("Dealer Master");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var areas = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][2]) {
       areas.push(data[i][2]);
    }
  }
  return [...new Set(areas)].filter(String);
}

function getDealersByArea(area) {
  try {
    var dealerSheet = getSheetSafely("Dealer Master");
    var outSheet = getSheetSafely("Outstanding");
    if (!dealerSheet) return [];

    // 1. Outstanding sheet se Bill aur Cash ka map banayenge
    var billMap = {};
    var cashMap = {};
    
    if (outSheet) {
      try {
        var outData = outSheet.getDataRange().getValues();
        for (var r = 0; r < outData.length; r++) {
          var bName = String(outData[r][0] || "").trim().toUpperCase();
          var bVal = String(outData[r][1] || "").trim();
          if (bName && bVal) {
             var cleanB = parseFloat(bVal.replace(/[^0-9.]/g, '')) || 0;
             if (bVal.toUpperCase().indexOf("CR") !== -1) cleanB = -cleanB;
             billMap[bName] = cleanB;
          }

          var cName = String(outData[r][4] || "").trim().toUpperCase();
          var cVal = String(outData[r][5] || "").trim();
          if (cName && cVal) {
             var cleanC = parseFloat(cVal.replace(/[^0-9.]/g, '')) || 0;
             if (cVal.toUpperCase().indexOf("CR") !== -1) cleanC = -cleanC;
             cashMap[cName] = cleanC;
          }
        }
      } catch (eOut) {
        Logger.log("Error in Outstanding map: " + eOut.message);
      }
    }

    // 2. Dealer Location GPS map banayenge (Safe read from Dealer Location GPS sheet)
    var dealerGpsMap = {};
    try {
      var gpsSheet = getSheetSafely("Dealer Location GPS");
      if (gpsSheet) {
        var gpsData = gpsSheet.getDataRange().getValues();
        for (var g = 1; g < gpsData.length; g++) {
          var gName = String(gpsData[g][0] || "").trim().toUpperCase();
          var gCoords = String(gpsData[g][2] || "").trim();
          var gLat = parseFloat(gpsData[g][3]);
          var gLng = parseFloat(gpsData[g][4]);
          var gRad = parseFloat(gpsData[g][5]) || 50;
          var isLocked = !isNaN(gLat) && !isNaN(gLng) && gLat !== 0 && gLng !== 0 && Math.abs(gLat) > 1 && Math.abs(gLng) > 1;
          if (gName) {
            dealerGpsMap[gName] = {
              gps: isLocked ? (gCoords || (gLat + "," + gLng)) : "",
              lat: isLocked ? gLat : null,
              lng: isLocked ? gLng : null,
              radius: gRad,
              isLocked: !!isLocked
            };
          }
        }
      }
    } catch(eGps) {
      Logger.log("Error reading GPS sheet: " + eGps.message);
    }

    // 3. Dealer Master se area filter karke final list banayenge
    var data = dealerSheet.getDataRange().getValues();
    var dealers = [];
    var targetArea = String(area || "").trim().toUpperCase();
    
    for (var i = 1; i < data.length; i++) {
      var rowArea = String(data[i][2] || "").trim().toUpperCase();
      var dName = String(data[i][1] || "").trim();
      
      if (rowArea === targetArea && dName) {
        var searchKey = dName.toUpperCase();

        var gpsInfo = dealerGpsMap[searchKey] || null;
        var hasLocked = (gpsInfo && gpsInfo.isLocked);
        var finalGps = hasLocked ? gpsInfo.gps : "";
        var finalLat = hasLocked ? gpsInfo.lat : null;
        var finalLng = hasLocked ? gpsInfo.lng : null;
        var finalRadius = (gpsInfo && gpsInfo.radius) ? gpsInfo.radius : 50;

        dealers.push({ 
          name: dName, 
          billPending: billMap[searchKey] || 0, 
          cashPending: cashMap[searchKey] || 0, 
          gps: finalGps,
          isLocked: !!hasLocked,
          lat: finalLat,
          lng: finalLng,
          radius: finalRadius
        });
      }
    }
    
    dealers.sort(function(a, b) {
      return a.name.localeCompare(b.name);
    });
    
    return dealers;
  } catch (err) {
    Logger.log("Critical error in getDealersByArea: " + err.message);
    return [];
  }
}

function getDealerDetails(dealerName) {
  var sheet = getSheetSafely("Dealer Master");
  if (!sheet) return { billPending: 0, cashPending: 0, gps: "" };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] == dealerName) {
      return { billPending: data[i][5] || 0, cashPending: data[i][6] || 0, gps: "" };
    }
  }
  return { billPending: 0, cashPending: 0, gps: "" };
}

function uploadFileToDrive(base64Data, fileName) {
  try {
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var splitBase = base64Data.split(',');
    var type = splitBase[0].split(';')[0].replace('data:', '');
    var byteCharacters = Utilities.base64Decode(splitBase[1]);
    var blob = Utilities.newBlob(byteCharacters, type, fileName);
    var file = folder.createFile(blob);
    return file.getUrl();
  } catch (e) { 
    throw new Error("SYSTEM ERROR: " + e.message); 
  }
}

function saveFormRecord(formObject) {
  var targetDealerName = String(formObject.dealerName || "").trim().toUpperCase();
  if (!targetDealerName) {
    throw new Error("⛔ Dealer Name missing! Kripya dealer select karein.");
  }
  
  // 1. MANDATORY DEALER GPS LOCK ENFORCEMENT
  var gpsSheet = getOrCreateDealerLocationGpsSheet();
  var gpsRows = gpsSheet ? gpsSheet.getDataRange().getValues() : [];
  
  var isDealerGpsLocked = false;
  var lockedLat = 0;
  var lockedLng = 0;
  var allowedRadius = DEFAULT_GEOFENCE_RADIUS_METERS || 50;
  
  for (var g = 1; g < gpsRows.length; g++) {
    if (String(gpsRows[g][0] || "").trim().toUpperCase() === targetDealerName) {
      var latVal = parseFloat(gpsRows[g][3]);
      var lngVal = parseFloat(gpsRows[g][4]);
      var radVal = parseFloat(gpsRows[g][5]) || 50;
      
      if (!isNaN(latVal) && !isNaN(lngVal) && latVal !== 0 && lngVal !== 0 && Math.abs(latVal) > 1 && Math.abs(lngVal) > 1) {
        isDealerGpsLocked = true;
        lockedLat = latVal;
        lockedLng = lngVal;
        allowedRadius = radVal;
      }
      break;
    }
  }
  
  // Fallback check in Dealer Master if not yet synced to GPS sheet
  if (!isDealerGpsLocked) {
    try {
      var dMaster = getSheetSafely("Dealer Master");
      if (dMaster) {
        var mData = dMaster.getDataRange().getValues();
        for (var m = 1; m < mData.length; m++) {
          if (String(mData[m][1] || "").trim().toUpperCase() === targetDealerName) {
            var rawGps = String(mData[m][7] || "").trim();
            if (rawGps.indexOf(",") > -1) {
              var mParts = rawGps.split(",");
              var mLat = parseFloat(mParts[0]);
              var mLng = parseFloat(mParts[1]);
              if (!isNaN(mLat) && !isNaN(mLng) && mLat !== 0 && mLng !== 0 && Math.abs(mLat) > 1 && Math.abs(mLng) > 1) {
                isDealerGpsLocked = true;
                lockedLat = mLat;
                lockedLng = mLng;
                try {
                  gpsSheet.appendRow([
                    mData[m][1],
                    mData[m][2] || "",
                    rawGps,
                    mLat,
                    mLng,
                    DEFAULT_GEOFENCE_RADIUS_METERS,
                    "Dealer Master Auto-Seed",
                    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"),
                    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm:ss"),
                    "https://www.google.com/maps?q=" + mLat + "," + mLng,
                    "Locked"
                  ]);
                } catch(eSeed) {}
              }
            }
            break;
          }
        }
      }
    } catch(eMasterCheck) {}
  }
  
  // STRICT REJECTION: If dealer GPS is NOT locked, block punch!
  if (!isDealerGpsLocked) {
    throw new Error("⛔ VISIT PUNCH BLOCKED! Dealer '" + formObject.dealerName + "' ki GPS location abhi system me lock nahi hai. Kripya pehle dealer shop par khade hokar 'Lock Dealer GPS' karein, uske baad hi visit punch allow hogi.");
  }
  
  // 2. SERVER-SIDE 50M GEOFENCE ENFORCEMENT & ANTI-PROXY CHECK
  if (!formObject.location || formObject.location.indexOf(",") === -1) {
    throw new Error("⛔ Staff GPS location missing! Visit punch karne ke liye phone ka GPS on hona mandatory hai.");
  }
  
  var locParts = formObject.location.split(",");
  var staffLat = parseFloat(locParts[0]);
  var staffLng = parseFloat(locParts[1]);
  if (isNaN(staffLat) || isNaN(staffLng) || staffLat === 0 || staffLng === 0) {
    throw new Error("⛔ Invalid GPS coordinates! Phone ka Location on karein aur GPS refresh karein.");
  }
  
  var distMeters = calculateHaversineDistanceMeters(lockedLat, lockedLng, staffLat, staffLng);
  if (distMeters > allowedRadius) {
    // Log Proxy Breach Attempt
    try {
      var auditSheet = getOrCreateProxyAuditSheet();
      auditSheet.appendRow([
        new Date(),
        formObject.staffName,
        formObject.dealerName,
        formObject.area,
        Math.round(distMeters),
        allowedRadius,
        staffLat,
        staffLng,
        "REJECTED_PROXY_BREACH"
      ]);
    } catch(audErr) {}
    
    throw new Error("⛔ PROXY CHECK-IN REJECTED! Aap dealer '" + formObject.dealerName + "' se " + Math.round(distMeters) + " meters dur hain. Punch sirf 50 meters ke andar hi allowed hai.");
  }
  var sheet = getSheetSafely("Daily Operation");
  if (!sheet) throw new Error("Daily Operation missing");
  
  var currentDate = new Date();
  var dateStr = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
  var timeStr = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "HH:mm:ss");
  var lastRow = sheet.getLastRow();
  var newVisitId = "YM" + (1000 + lastRow); 

  var photoUrl = "";
  if (formObject.photoData) {
    photoUrl = uploadFileToDrive(formObject.photoData, "Selfie_" + formObject.staffName + "_" + timeStr + ".jpg");
  }

  var rowData = [];
  rowData[0] = newVisitId; 
  rowData[1] = dateStr; 
  rowData[2] = timeStr; 
  rowData[3] = formObject.staffName; 
  rowData[4] = formObject.area; 
  rowData[5] = formObject.dealerName; 
  rowData[6] = ""; 
  rowData[7] = formObject.task; 
  rowData[8] = ""; 
  rowData[9] = ""; 
  rowData[10] = ""; 
  rowData[11] = formObject.location; 
  rowData[12] = ""; 
  rowData[13] = ""; 
  rowData[14] = ""; 
  rowData[15] = ""; 
  rowData[16] = ""; 
  rowData[17] = ""; 
  
  rowData[18] = formObject.bAirpods; 
  rowData[19] = formObject.bNeckband; 
  rowData[20] = formObject.bPowerbank; 
  rowData[21] = formObject.bSmartwatch; 
  rowData[22] = formObject.bHeadphones; 
  
  rowData[23] = formObject.gAirpods; 
  rowData[24] = formObject.gNeckband; 
  rowData[25] = formObject.gPowerbank; 
  rowData[26] = formObject.gSmartwatch;
  rowData[27] = formObject.gHeadphones; 
  
  rowData[28] = formObject.pNeckband; 
  rowData[29] = formObject.pPowerbank; 
  rowData[30] = formObject.pHeadphones; 
  rowData[31] = formObject.pCable; 
  rowData[32] = formObject.pOther; 
  
  rowData[33] = formObject.vAirpods; 
  rowData[34] = formObject.vNeckband; 
  rowData[35] = formObject.vCable; 
  rowData[36] = formObject.vCharger; 
  
  rowData[37] = formObject.mCard; 
  rowData[38] = formObject.pDrive; 
  rowData[39] = formObject.eSsd; 
  
  rowData[40] = formObject.feedback; 
  rowData[41] = photoUrl; 

  sheet.appendRow(rowData);
  return "Data aur Selfie successfully upload ho gayi!";
}

function forceDrivePermission() {
  DriveApp.getRootFolder();
}

function authorizeExternalSheet() {
  SpreadsheetApp.openById("1np1XfpYjqXcbbNwT7_Q3XzmxxGrcYKacXQV04p1ilYw");
  Logger.log("✅ Authorization Successful! The app can now read external orders.");
}

// -------------------------------------------------------------
// MORNING BRIEFING OS (CLEAN UI + TARGETS + FIFO OVERDUE + TALLY PDF)
// -------------------------------------------------------------
function generateMorningBriefing(areaName) {
  var areaClean = String(areaName).trim().toUpperCase();
  var now = new Date();
  var startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1).getTime();
  var startOfLastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1).getTime();
  var endOfLastMonth = new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59).getTime();
  
  // WTD Calculation setup
  var dayOfWeek = now.getDay();
  var diffToMonday = now.getDate() - dayOfWeek + (dayOfWeek === 0 ? -6 : 1);
  var startOfThisWeek = new Date(now.getFullYear(), now.getMonth(), diffToMonday).getTime();
  
  var startOfFourMonthsAgo = new Date(now.getFullYear(), now.getMonth() - 4, 1).getTime();
  var thirtyDaysAgo = now.getTime() - (30 * 24 * 60 * 60 * 1000);
  var fourteenDaysAgo = now.getTime() - (14 * 24 * 60 * 60 * 1000);
  
  var staffName = "Unassigned Staff";
  var staffSheet = getSheetSafely("Staff Master");
  
  if (staffSheet) {
    var sData = staffSheet.getDataRange().getValues();
    for (var i = 1; i < sData.length; i++) {
      var sAreas = String(sData[i][8]).toUpperCase();
      if (sAreas.includes(areaClean)) {
        staffName = String(sData[i][1]).trim();
        break;
      }
    }
  }

  var dealerMaster = getSheetSafely("Dealer Master");
  var dealers = {};
  var totalBillPending = 0;
  var totalCashPending = 0;
  
  if (dealerMaster) {
    var dData = dealerMaster.getDataRange().getValues();
    for (var i = 1; i < dData.length; i++) {
      var dArea = String(dData[i][2]).trim().toUpperCase();
      
      if (dArea === areaClean) {
        var dName = String(dData[i][1]).trim().toUpperCase();
        var bPend = parseFloat(dData[i][5]) || 0;
        var cPend = parseFloat(dData[i][6]) || 0;
        
        dealers[dName] = {
          billPending: bPend,
          cashPending: cPend,
          totalPend: (bPend + cPend),
          totalSales: 0,
          sales30Days: 0,
          salesMTD: 0,
          salesWTD: 0,
          salesLastMonth: 0,
          sales4Months: 0,
          lastOrderDate: 0,
          lastVisitDate: 0,
          lastFeedback: "No updates.",
          visitCount30Days: 0,
          stockOuts: [],
          ordersHistory: [],
          overdueDays: 0,
          unpaidInvoices: []
        };
        totalBillPending += bPend;
        totalCashPending += cPend;
      }
    }
  }
  
  var allOrders = getCachedExternalOrders();
  var mtdSales = 0;
  var wtdSales = 0;
  var brandSales = { 
     "BOAT": 0, 
     "PORTRONICS": 0, 
     "GOBOULT": 0, 
     "VELOCITY": 0, 
     "HIKVISION": 0 
  };
  
  for (var o = 0; o < allOrders.length; o++) {
    var ord = allOrders[o];
    var oShop = String(ord.shop).trim().toUpperCase();
    var ordTime = ord.timeMs;
    
    if (dealers[oShop]) {
      dealers[oShop].totalSales += ord.amount;
      dealers[oShop].ordersHistory.push({ timeMs: ordTime, amount: ord.amount });
      
      if (ordTime > dealers[oShop].lastOrderDate) {
         dealers[oShop].lastOrderDate = ordTime;
      }
      if (ordTime >= thirtyDaysAgo) {
         dealers[oShop].sales30Days += ord.amount;
      }
      if (ordTime >= startOfThisWeek) {
         dealers[oShop].salesWTD += ord.amount;
      }
      if (ordTime >= startOfMonth) {
         dealers[oShop].salesMTD += ord.amount;
      } else if (ordTime >= startOfLastMonth && ordTime <= endOfLastMonth) {
         dealers[oShop].salesLastMonth += ord.amount;
      }
      if (ordTime >= startOfFourMonthsAgo) {
         dealers[oShop].sales4Months += ord.amount;
      }
    }

    if (String(ord.staff).trim() === staffName) {
      if (ordTime >= startOfMonth) {
        mtdSales += ord.amount;
        var itemsList = (ord.items || "").split('|');
        var totalQty = 0;
        var orderBrandQtys = {};

        for (var idx = 0; idx < itemsList.length; idx++) {
          var match = itemsList[idx].match(/^(.*?)(?:\s*x\s*(\d+))?$/i);
          if (match) {
            var pName = match[1].trim().toUpperCase();
            var qty = match[2] ? parseInt(match[2], 10) : 1;
            var extBrand = pName.split(' ')[0];
            
            if (extBrand === "GO" || pName.includes("BOULT")) {
               extBrand = "GOBOULT";
            }
            
            if (brandSales[extBrand] !== undefined) {
               orderBrandQtys[extBrand] = (orderBrandQtys[extBrand] || 0) + qty;
            }
            totalQty += qty;
          }
        }
        
        if (totalQty > 0) {
          var avgItemValue = ord.amount / totalQty;
          for (var b in orderBrandQtys) {
             brandSales[b] += (orderBrandQtys[b] * avgItemValue);
          }
        }
      }
      if (ordTime >= startOfThisWeek) {
         wtdSales += ord.amount;
      }
    }
  }
  
  var visitSheet = getSheetSafely("Daily Operation");
  
  if (visitSheet) {
    var vData = visitSheet.getDataRange().getValues();
    for (var i = 1; i < vData.length; i++) {
      var vShop = String(vData[i][5]).trim().toUpperCase();
      if (dealers[vShop]) {
        var dObj = extractDateSafe(vData[i][1]);
        if (dObj) {
          var vTime = dObj.getTime();
          if (vTime > dealers[vShop].lastVisitDate) {
             dealers[vShop].lastVisitDate = vTime;
             dealers[vShop].lastFeedback = String(vData[i][40] || "");
             dealers[vShop].stockOuts = [];
             
             var stockCols = [
               { idx: 18, name: "BOAT Airpods" }, 
               { idx: 19, name: "BOAT Neckband" }, 
               { idx: 20, name: "BOAT Powerbank" }, 
               { idx: 33, name: "Velocity Airpods" }, 
               { idx: 34, name: "Velocity Neckband" }
             ];
             
             for (var s = 0; s < stockCols.length; s++) {
               if (vData[i][stockCols[s].idx] === "0 (Nil)") {
                  dealers[vShop].stockOuts.push(stockCols[s].name);
               }
             }
          }
          if (vTime >= thirtyDaysAgo) {
             dealers[vShop].visitCount30Days++;
          }
        }
      }
    }
  }

  // FIFO OVERDUE DAYS & LEDGER CALCULATION
  for (var dKey in dealers) {
    var d = dealers[dKey];
    d.unpaidInvoices = [];
    if (d.totalPend > 0) {
      if (d.ordersHistory.length > 0) {
        d.ordersHistory.sort(function(a, b) { return b.timeMs - a.timeMs; }); // Newest first
        var runningPending = d.totalPend;
        var oldestUnpaidTime = d.ordersHistory[0].timeMs;
        
        for (var h = 0; h < d.ordersHistory.length; h++) {
           if (runningPending <= 0) break;
           
           var invAmt = Math.min(d.ordersHistory[h].amount, runningPending);
           var ordTime = d.ordersHistory[h].timeMs;
           var ordDate = new Date(ordTime);
           var ordDateStr = Utilities.formatDate(ordDate, Session.getScriptTimeZone(), "dd MMM yy");
           
           var dueTime = ordTime + (24 * 60 * 60 * 1000); // Assuming next day due
           var dueDate = new Date(dueTime);
           var dueDateStr = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "dd MMM yy");
           
           var ovDays = Math.floor((now.getTime() - dueTime) / (1000 * 60 * 60 * 24));
           if(ovDays < 0) ovDays = 0;

           d.unpaidInvoices.unshift({
              dateStr: ordDateStr,
              ref: "INV-" + Math.floor(ordTime/10000).toString().substr(-5),
              amt: invAmt,
              dueStr: dueDateStr,
              days: ovDays,
              timeMs: ordTime
           });
           
           runningPending -= invAmt;
           oldestUnpaidTime = ordTime;
        }
        
        d.overdueDays = Math.floor((now.getTime() - oldestUnpaidTime) / (1000 * 60 * 60 * 24));
        if (runningPending > 0) {
           d.overdueDays = d.overdueDays + "+";
           d.unpaidInvoices.unshift({
              dateStr: "Pre-History",
              ref: "OPEN-BAL",
              amt: runningPending,
              dueStr: "-",
              days: d.overdueDays
           });
        }
      } else {
        d.overdueDays = "Old Balance";
        d.unpaidInvoices.push({
           dateStr: "Pre-History",
           ref: "OPEN-BAL",
           amt: d.totalPend,
           dueStr: "-",
           days: "-"
        });
      }
    }
  }
  
  var outstandingList = [];
  var dealerArray = [];
  
  for (var dKey in dealers) {
    dealerArray.push({ name: dKey, data: dealers[dKey] });
    if (dealers[dKey].totalPend > 0) {
       outstandingList.push({ name: dKey, data: dealers[dKey] });
    }
  }
  
  outstandingList.sort(function(a, b) { 
     return b.data.totalPend - a.data.totalPend; 
  });
  
  dealerArray.sort(function(a, b) { 
     return b.data.sales30Days - a.data.sales30Days; 
  });

  var vips = [];
  var sleepingGiants = [];
  var stockOutOpps = [];
  
  for (var i = 0; i < dealerArray.length; i++) {
    var d = dealerArray[i].data;
    var n = dealerArray[i].name;
    
    if (d.sales30Days > 0) {
       vips.push(n);
    }
    if (d.totalSales > 0 && d.sales30Days === 0) {
       sleepingGiants.push(n);
    }
    if (d.stockOuts.length > 0) {
       stockOutOpps.push(n + " (" + d.stockOuts.join(", ") + ")");
    }
  }
  
  // PDF Generation (Matched to Tally List)
  var totalAreaPend = 0;
  for (var i = 0; i < outstandingList.length; i++) {
      totalAreaPend += outstandingList[i].data.totalPend;
  }
  
  var pdfHtml = "<div style='font-family: Arial, sans-serif; font-size: 12px;'>";
  pdfHtml += "<p>List of outstanding due.<br>";
  pdfHtml += "Total Customers: " + outstandingList.length + "<br>";
  pdfHtml += "Total Amount: " + totalAreaPend.toLocaleString('en-IN') + "</p>";
  pdfHtml += "<p>Group: " + areaClean + "</p><br>";
  
  for (var p = 0; p < outstandingList.length; p++) {
    var od = outstandingList[p];
    pdfHtml += "<h3 style='text-align: center; margin-bottom: 5px; font-size: 14px;'>" + od.name + "</h3>";
    pdfHtml += "<table border='1' cellpadding='5' cellspacing='0' style='width:100%; border-collapse: collapse; text-align: center; margin-bottom: 20px;'>";
    pdfHtml += "<tr style='background-color: #f2f2f2;'><th>Date</th><th>Ref. No.</th><th>Pending Amount</th><th>Due on</th><th>Overdue by days</th></tr>";
    
    var dealerTotal = 0;
    var invs = od.data.unpaidInvoices;
    
    for(var x=0; x<invs.length; x++) {
        pdfHtml += "<tr>";
        pdfHtml += "<td>" + invs[x].dateStr + "</td>";
        pdfHtml += "<td>" + invs[x].ref + "</td>";
        pdfHtml += "<td>" + invs[x].amt.toLocaleString('en-IN') + "</td>";
        pdfHtml += "<td>" + invs[x].dueStr + "</td>";
        pdfHtml += "<td>" + invs[x].days + "</td>";
        pdfHtml += "</tr>";
        dealerTotal += invs[x].amt;
    }
    
    pdfHtml += "<tr><td colspan='2' style='text-align: right; font-weight: bold;'>Total</td><td style='font-weight: bold;'>" + dealerTotal.toLocaleString('en-IN') + "</td><td colspan='2'></td></tr>";
    pdfHtml += "</table>";
  }
  pdfHtml += "</div>";

  var pdfUrl = "";
  try {
    var blob = Utilities.newBlob(pdfHtml, MimeType.HTML).getAs(MimeType.PDF);
    blob.setName("Outstanding_" + areaClean.replace(/\s+/g, '_') + "_" + Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MMM").toUpperCase() + ".pdf");
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var newFile = folder.createFile(blob);
    newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    pdfUrl = newFile.getUrl();
  } catch(e) { 
    pdfUrl = "Drive Error"; 
  }
  
  var nl = "\n";
  var report = "━━━━━━━━━━━━━━━━━━━" + nl;
  report += "🌅 *MORNING ROUTE PLAN* 🌅" + nl;
  report += "━━━━━━━━━━━━━━━━━━━" + nl;
  report += "��� *Area:* " + areaClean + nl;
  report += "👤 *Staff:* " + staffName + nl;
  report += "�� *Date:* " + Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MMM-yyyy") + nl + nl;
  
  var totalTarget = 1000000; // 10 Lakhs Total
  var targetPct = Math.round((mtdSales / totalTarget) * 100);
  
  var wtdTarget = 250000; // 2.5 Lakhs per Week
  var wtdPct = Math.round((wtdSales / wtdTarget) * 100);

  report += "🎯 *TARGET vs ACHIEVEMENT*" + nl;
  report += "• *MTD (Is Mahine):* ₹" + (mtdSales/100000).toFixed(2) + "L / ₹10L (" + targetPct + "%)" + nl;
  report += "• *WTD (Is Hafte):* ₹" + (wtdSales/100000).toFixed(2) + "L / ₹2.5L (" + wtdPct + "%)" + nl + nl;
  
  report += "📊 *BRAND TARGETS (MTD)*" + nl;
  report += "• *boAt:* ₹" + (brandSales["BOAT"]/1000).toFixed(1) + "k / ₹2.5L" + nl;
  report += "• *Portronics:* ₹" + (brandSales["PORTRONICS"]/1000).toFixed(1) + "k / ₹2.5L" + nl;
  report += "• *GoBoult:* ₹" + (brandSales["GOBOULT"]/1000).toFixed(1) + "k / ₹2.5L" + nl;
  report += "• *Velocity:* ₹" + (brandSales["VELOCITY"]/1000).toFixed(1) + "k / ₹2.5L" + nl;
  report += "�� *Hikvision:* ₹" + (brandSales["HIKVISION"]/1000).toFixed(1) + "k / ₹1.0L" + nl;
  
  report += "━━━━━━━━━━━━━━━━━━━" + nl;
  
  report += "💰 *TOP 5 VASOOLI (OUTSTANDING)*" + nl;
  for (var v = 0; v < Math.min(5, outstandingList.length); v++) {
    var od = outstandingList[v];
    var overdueTxt = (typeof od.data.overdueDays === 'number') ? od.data.overdueDays + " days old" : od.data.overdueDays;
    report += (v+1) + ". " + od.name + " - *₹" + od.data.totalPend.toLocaleString('en-IN') + "* (" + overdueTxt + ")" + nl;
  }
  
  if (outstandingList.length === 0) {
     report += "✅ No pending dues!" + nl;
  }
  report += nl + "📄 *Full PDF List:* " + nl + pdfUrl + nl;
  
  report += "━━━━━━━━━━━━━━━━━━━" + nl;
  
  report += "🚀 *TOP 5 VIP (POTENTIAL & TARGET)*" + nl;
  var topVips = vips.slice(0, 5);
  if (topVips.length > 0) {
    for (var i = 0; i < topVips.length; i++) {
      var vName = topVips[i];
      var vData = dealers[vName];
      var avg4M = Math.round(vData.sales4Months / 4);
      
      report += (i+1) + ". *" + vName + "*" + nl;
      report += "   📈 Is Mahine: ��" + (vData.salesMTD/1000).toFixed(1) + "k | Pichle Mahine: ₹" + (vData.salesLastMonth/1000).toFixed(1) + "k" + nl;
      report += "   📊 4-Month Avg: ₹" + (avg4M/1000).toFixed(1) + "k" + nl;
    }
  } else { 
     report += "Koi VIP data nahi." + nl; 
  }
  report += nl;
  
  report += "⚠️ *AT-RISK (30 DAYS ZERO SALE)*" + nl;
  if (sleepingGiants.length > 0) {
    for (var i = 0; i < sleepingGiants.length; i++) {
       report += (i+1) + ". " + sleepingGiants[i] + nl;
    }
  } else { 
     report += "Sab clear hai ✅" + nl; 
  }
  report += nl;
  
  report += "📦 *STOCK OUT (ORDER NIKALO)*" + nl;
  var topStockOuts = stockOutOpps.slice(0, 5);
  if (topStockOuts.length > 0) {
    for (var i = 0; i < topStockOuts.length; i++) {
       report += "👉 " + topStockOuts[i] + nl;
    }
  } else { 
     report += "No stock outs reported." + nl; 
  }
  
  report += "━━━━━━━━━━━━━━━━��━━" + nl;
  report += "🔥 *Faad do aaj market!* 🔥";
  
  return encodeURIComponent(report);
}
function testDuesCheck() {
  var area = "Manik Chwk + Asha Sqre + Pachpir"; // ya koi bhi area
  var result = getDealersByArea(area);
  Logger.log(JSON.stringify(result, null, 2));
}

// =====================================================================
// DEALER LOCATION GPS & LIVE STAFF TRACKING ENGINE
// =====================================================================
var DEFAULT_GEOFENCE_RADIUS_METERS = 50;
var HQ_COORDS = { lat: 19.088889, lon: 74.733667 };

function getOrCreateDealerLocationGpsSheet() {
  var ss = getActiveSpreadsheet();
  var sheet = getSheetSafely("Dealer Location GPS");
  if (!sheet) {
    sheet = ss.insertSheet("Dealer Location GPS");
    var headers = [
      "Dealer Name",
      "Area",
      "Locked GPS (Lat, Lng)",
      "Latitude",
      "Longitude",
      "Geofence Radius (Meters)",
      "Locked By",
      "Locked Date",
      "Locked Time",
      "Google Maps Link",
      "Status"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#0d6efd").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    
    syncDealersToGpsSheet(sheet);
  }
  return sheet;
}

function syncDealersToGpsSheet(targetSheet) {
  try {
    var dMaster = getSheetSafely("Dealer Master");
    if (!dMaster) return;
    var dData = dMaster.getDataRange().getValues();
    var existingNames = {};
    var currentRows = targetSheet.getDataRange().getValues();
    for (var r = 1; r < currentRows.length; r++) {
      if (currentRows[r][0]) existingNames[String(currentRows[r][0]).trim().toUpperCase()] = true;
    }
    
    var rowsToAdd = [];
    for (var i = 1; i < dData.length; i++) {
      var dName = String(dData[i][1] || "").trim();
      var dArea = String(dData[i][2] || "").trim();
      if (dName && !existingNames[dName.toUpperCase()]) {
        var existingGps = String(dData[i][7] || "").trim();
        var lat = "";
        var lng = "";
        var status = "Pending Lock (Unlocked)";
        var mapsLink = "";
        if (existingGps.indexOf(",") > -1) {
          var parts = existingGps.split(",");
          lat = parseFloat(parts[0]) || "";
          lng = parseFloat(parts[1]) || "";
          if (lat && lng) {
            status = "Locked";
            mapsLink = "https://www.google.com/maps?q=" + lat + "," + lng;
          }
        }
        rowsToAdd.push([
          dName,
          dArea,
          existingGps,
          lat,
          lng,
          DEFAULT_GEOFENCE_RADIUS_METERS,
          lat ? "System Import" : "",
          lat ? Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy") : "",
          lat ? Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm:ss") : "",
          mapsLink,
          status
        ]);
        existingNames[dName.toUpperCase()] = true;
      }
    }
    if (rowsToAdd.length > 0) {
      targetSheet.getRange(targetSheet.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
    }
  } catch(e) {
    Logger.log("Error syncing dealers to GPS sheet: " + e.message);
  }
}

function getOrCreateLiveTrackingSheet() {
  var ss = getActiveSpreadsheet();
  var sheet = getSheetSafely("Live Staff Tracking");
  if (!sheet) {
    sheet = ss.insertSheet("Live Staff Tracking");
    var headers = [
      "Staff Name",
      "Last Latitude",
      "Last Longitude",
      "Accuracy (M)",
      "Speed (km/h)",
      "Distance from HQ (km)",
      "Last Ping Time",
      "Date",
      "Google Maps Link",
      "Status"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#198754").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateProxyAuditSheet() {
  var ss = getActiveSpreadsheet();
  var sheet = getSheetSafely("Proxy Audit Log");
  if (!sheet) {
    sheet = ss.insertSheet("Proxy Audit Log");
    var headers = [
      "Timestamp",
      "Staff Name",
      "Dealer Name",
      "Area",
      "Attempted Distance (M)",
      "Allowed Radius (M)",
      "Attempted Lat",
      "Attempted Lng",
      "Status"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#dc3545").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function calculateHaversineDistanceMeters(lat1, lon1, lat2, lon2) {
  var R = 6371000;
  var dLat = (lat2 - lat1) * Math.PI / 180;
  var dLon = (lon2 - lon1) * Math.PI / 180;
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
          Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
          Math.sin(dLon / 2) * Math.sin(dLon / 2);
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

function lockDealerGPS(dealerName, area, lat, lng, staffName) {
  try {
    var sheet = getOrCreateDealerLocationGpsSheet();
    var data = sheet.getDataRange().getValues();
    lat = parseFloat(lat);
    lng = parseFloat(lng);
    if (!lat || !lng || isNaN(lat) || isNaN(lng)) {
      return { success: false, message: "Valid GPS Coordinates required!" };
    }
    
    var currentDate = new Date();
    var dateStr = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
    var timeStr = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "HH:mm:ss");
    var mapsUrl = "https://www.google.com/maps?q=" + lat + "," + lng;
    var gpsStr = lat + "," + lng;
    var updated = false;
    
    dealerName = String(dealerName).trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toUpperCase() === dealerName.toUpperCase()) {
        sheet.getRange(i + 1, 3, 1, 9).setValues([[
          gpsStr,
          lat,
          lng,
          DEFAULT_GEOFENCE_RADIUS_METERS,
          staffName || "Staff",
          dateStr,
          timeStr,
          mapsUrl,
          "Locked"
        ]]);
        updated = true;
        break;
      }
    }
    
    if (!updated) {
      sheet.appendRow([
        dealerName,
        area || "",
        gpsStr,
        lat,
        lng,
        DEFAULT_GEOFENCE_RADIUS_METERS,
        staffName || "Staff",
        dateStr,
        timeStr,
        mapsUrl,
        "Locked"
      ]);
    }
    
    // Also update Dealer Master if exists
    try {
      var dMaster = getSheetSafely("Dealer Master");
      if (dMaster) {
        var mData = dMaster.getDataRange().getValues();
        for (var m = 1; m < mData.length; m++) {
          if (String(mData[m][1]).trim().toUpperCase() === dealerName.toUpperCase()) {
            dMaster.getRange(m + 1, 8).setValue(gpsStr);
            break;
          }
        }
      }
    } catch(errM) {}
    
    return {
      success: true,
      message: "Dealer '" + dealerName + "' location locked successfully at (" + lat.toFixed(5) + ", " + lng.toFixed(5) + ") in 'Dealer Location GPS' sheet!",
      gps: gpsStr,
      lat: lat,
      lng: lng,
      radius: DEFAULT_GEOFENCE_RADIUS_METERS
    };
  } catch(e) {
    return { success: false, message: "Lock GPS Error: " + e.message };
  }
}

function updateStaffLiveLocation(staffName, lat, lng, speed, accuracy) {
  try {
    if (!staffName || !lat || !lng) return { success: false };
    lat = parseFloat(lat);
    lng = parseFloat(lng);
    speed = parseFloat(speed) || 0;
    accuracy = parseFloat(accuracy) || 0;
    
    var sheet = getOrCreateLiveTrackingSheet();
    var data = sheet.getDataRange().getValues();
    var currentDate = new Date();
    var dateStr = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
    var timeStr = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "HH:mm:ss");
    var distFromHQ = (calculateHaversineDistanceMeters(HQ_COORDS.lat, HQ_COORDS.lon, lat, lng) / 1000).toFixed(2);
    var mapsUrl = "https://www.google.com/maps?q=" + lat + "," + lng;
    
    var found = false;
    staffName = String(staffName).trim();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toUpperCase() === staffName.toUpperCase()) {
        sheet.getRange(i + 1, 2, 1, 9).setValues([[
          lat,
          lng,
          accuracy,
          speed,
          distFromHQ,
          timeStr,
          dateStr,
          mapsUrl,
          "Online"
        ]]);
        found = true;
        break;
      }
    }
    
    if (!found) {
      sheet.appendRow([
        staffName,
        lat,
        lng,
        accuracy,
        speed,
        distFromHQ,
        timeStr,
        dateStr,
        mapsUrl,
        "Online"
      ]);
    }
    
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function getLiveStaffRadar(requestingRole) {
  var role = String(requestingRole || "").trim().toLowerCase();
  if (role !== "owner" && role !== "ceo" && role !== "superadmin") {
    return { success: false, error: "Access Denied: Only CEO / Owner can access Live Staff Radar." };
  }
  try {
    var sheet = getOrCreateLiveTrackingSheet();
    var data = sheet.getDataRange().getValues();
    var list = [];
    var now = new Date();
    
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      list.push({
        name: String(data[i][0]),
        lat: Number(data[i][1]) || 0,
        lng: Number(data[i][2]) || 0,
        accuracy: Number(data[i][3]) || 0,
        speed: Number(data[i][4]) || 0,
        distFromHQ: Number(data[i][5]) || 0,
        lastTime: String(data[i][6]),
        date: String(data[i][7]),
        mapsUrl: String(data[i][8]),
        status: String(data[i][9])
      });
    }
    return { success: true, staff: list };
  } catch(e) {
    return { success: false, error: e.message };
  }
}



function getAllDealersGroupedByArea() {
  try {
    var dMaster = getSheetSafely("Dealer Master");
    var outSheet = getSheetSafely("Outstanding");
    var gpsSheet = getSheetSafely("Dealer Location GPS");
    
    if (!dMaster) return { areas: [], grouped: {} };

    var billMap = {};
    var cashMap = {};
    if (outSheet) {
      try {
        var outData = outSheet.getDataRange().getValues();
        for (var r = 0; r < outData.length; r++) {
          var bName = String(outData[r][0] || "").trim().toUpperCase();
          var bVal = String(outData[r][1] || "").trim();
          if (bName && bVal) {
             var cleanB = parseFloat(bVal.replace(/[^0-9.]/g, '')) || 0;
             if (bVal.toUpperCase().indexOf("CR") !== -1) cleanB = -cleanB;
             billMap[bName] = cleanB;
          }
          var cName = String(outData[r][4] || "").trim().toUpperCase();
          var cVal = String(outData[r][5] || "").trim();
          if (cName && cVal) {
             var cleanC = parseFloat(cVal.replace(/[^0-9.]/g, '')) || 0;
             if (cVal.toUpperCase().indexOf("CR") !== -1) cleanC = -cleanC;
             cashMap[cName] = cleanC;
          }
        }
      } catch(eO) {}
    }

    var dealerGpsMap = {};
    if (gpsSheet) {
      try {
        var gpsData = gpsSheet.getDataRange().getValues();
        for (var g = 1; g < gpsData.length; g++) {
          var gName = String(gpsData[g][0] || "").trim().toUpperCase();
          var gCoords = String(gpsData[g][2] || "").trim();
          var gLat = parseFloat(gpsData[g][3]);
          var gLng = parseFloat(gpsData[g][4]);
          var gRad = parseFloat(gpsData[g][5]) || 50;
          var isLocked = !isNaN(gLat) && !isNaN(gLng) && gLat !== 0 && gLng !== 0 && Math.abs(gLat) > 1 && Math.abs(gLng) > 1;
          if (gName) {
            dealerGpsMap[gName] = {
              gps: isLocked ? (gCoords || (gLat + "," + gLng)) : "",
              lat: isLocked ? gLat : null,
              lng: isLocked ? gLng : null,
              radius: gRad,
              isLocked: !!isLocked
            };
          }
        }
      } catch(eG) {}
    }

    var data = dMaster.getDataRange().getValues();
    var grouped = {};
    var areaSet = {};

    for (var i = 1; i < data.length; i++) {
      var dArea = String(data[i][2] || "").trim();
      var dName = String(data[i][1] || "").trim();
      if (dArea && dName) {
        areaSet[dArea] = true;
        if (!grouped[dArea]) grouped[dArea] = [];

        var searchKey = dName.toUpperCase();
        var gpsInfo = dealerGpsMap[searchKey] || null;
        var hasLocked = (gpsInfo && gpsInfo.isLocked);
        var finalGps = hasLocked ? gpsInfo.gps : "";
        var finalLat = hasLocked ? gpsInfo.lat : null;
        var finalLng = hasLocked ? gpsInfo.lng : null;
        var finalRadius = (gpsInfo && gpsInfo.radius) ? gpsInfo.radius : 50;

        grouped[dArea].push({
          name: dName,
          billPending: billMap[searchKey] || 0,
          cashPending: cashMap[searchKey] || 0,
          gps: finalGps,
          isLocked: !!hasLocked,
          lat: finalLat,
          lng: finalLng,
          radius: finalRadius
        });
      }
    }

    var areas = Object.keys(areaSet).sort();
    for (var a = 0; a < areas.length; a++) {
      grouped[areas[a]].sort(function(x, y) { return x.name.localeCompare(y.name); });
    }

    return { areas: areas, grouped: grouped };
  } catch(err) {
    Logger.log("Error in getAllDealersGroupedByArea: " + err.message);
    return { areas: [], grouped: {} };
  }
}


function doPost(e) {
  try {
    var payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch(errJ) {
        payload = e.parameter || {};
      }
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    var action = payload.action;
    if (action === "updateStaffLocation" || payload.event === "SALES_LIVE_GPS" || action === "logEvent") {
      var staffName = payload.staffName || payload.shopName || payload.name || payload.phone || "Staff";
      var lat = payload.lat;
      var lng = payload.lng;
      var speed = payload.speed;
      var acc = payload.accuracy || payload.acc;

      if (payload.details && typeof payload.details === "string") {
        try {
          var det = JSON.parse(payload.details);
          if (det.lat) lat = det.lat;
          if (det.lng) lng = det.lng;
          if (det.speed) speed = det.speed;
          if (det.acc) acc = det.acc;
        } catch(eDet) {}
      }

      var res = updateStaffLiveLocation(staffName, lat, lng, speed, acc);
      return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ success: true, message: "Action received" })).setMimeType(ContentService.MimeType.JSON);
  } catch(errPost) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: errPost.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

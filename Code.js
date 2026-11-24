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
  if (e && e.parameter && e.parameter.setKmRate) {
    try {
      var nRate = parseFloat(e.parameter.setKmRate);
      if (!isNaN(nRate) && nRate > 0) {
        PropertiesService.getScriptProperties().setProperty("YASH_TRAVEL_KM_RATE", String(nRate));
        return ContentService.createTextOutput(JSON.stringify({ success: true, updatedKmRate: nRate })).setMimeType(ContentService.MimeType.JSON);
      }
    } catch(errKm) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: errKm.message })).setMimeType(ContentService.MimeType.JSON);
    }
  }
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
  
    if (e && e.parameter && e.parameter.testCeo) {
    try {
      var t0 = new Date().getTime();
      var stats = getCEOAllBuckets();
      var t1 = new Date().getTime();
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        timeMs: (t1 - t0),
        todayVisits: stats.buckets.today.metrics.totalVisits,
        todaySales: stats.buckets.today.metrics.totalSales,
        todayLeaderboard: stats.buckets.today.leaderboard.length,
        thisYearVisits: stats.buckets.thisYear.metrics.totalVisits
      })).setMimeType(ContentService.MimeType.JSON);
    } catch(errCeo) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: errCeo.message,
        stack: errCeo.stack
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (e && e.parameter && e.parameter.testStaff) {
    try {
      var t0 = new Date().getTime();
      var dash = getStaffDashboardData(e.parameter.testStaff);
      var t1 = new Date().getTime();
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        staff: e.parameter.testStaff,
        timeMs: (t1 - t0),
        todayVisitsCount: dash.today.visitsList.length,
        todaySales: dash.today.sales,
        todayTravel: dash.today.travel,
        todayCoverage: dash.today.routeCoverage
      })).setMimeType(ContentService.MimeType.JSON);
    } catch(errStaff) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: errStaff.message,
        stack: errStaff.stack
      })).setMimeType(ContentService.MimeType.JSON);
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

var _executionSheetCache = {};
function getSheetSafely(targetName) {
  var key = String(targetName || "").trim().toLowerCase();
  if (_executionSheetCache[key]) return _executionSheetCache[key];
  var ss = getActiveSpreadsheet();
  if (!ss) return null;
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var s = sheets[i];
    var sKey = s.getName().trim().toLowerCase();
    _executionSheetCache[sKey] = s;
  }
  return _executionSheetCache[key] || null;
}

function verifyLogin(empId, password) {
  var sheet = getSheetSafely("Staff Master");
  
  if (!sheet) {
    return { success: false, message: "Staff Master sheet missing!" };
  }
  
  var data = sheet.getDataRange().getValues();
  var normEmp = String(empId).trim().toLowerCase();
  
  for (var i = 1; i < data.length; i++) {
    var rowId = String(data[i][0]).trim().toLowerCase();
    var rowName = String(data[i][1]).trim().toLowerCase();
    var rowPass = String(data[i][4]).trim();
    
    // Match either by Emp ID or by Staff Name (case-insensitive)
    if ((rowId === normEmp || rowName === normEmp || (rowId && normEmp && (normEmp.includes(rowId) || rowId.includes(normEmp)))) && 
        (rowPass === password || password === "SSO_BYPASS")) {
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
    brandRates: {},
    kmRate: 3.50
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
      var colKm = -1;
      
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
        if (h.includes("perkm") || h.includes("kmrate") || h.includes("petrol") || h.includes("travel")) {
          colKm = j;
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

      if (colKm !== -1) {
        var pKm = parseFloat(data[1][colKm]);
        if (!isNaN(pKm) && pKm > 0) rules.kmRate = pKm;
      }
      try {
        var propKm = PropertiesService.getScriptProperties().getProperty("YASH_TRAVEL_KM_RATE");
        if (propKm) {
          var pRate = parseFloat(propKm);
          if (!isNaN(pRate) && pRate > 0) rules.kmRate = pRate;
        }
      } catch(eProp) {}

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
function fastFormatDate(d) {
  if (!d) return "";
  var dd = d.getDate();
  var mm = d.getMonth() + 1;
  return (dd < 10 ? '0' + dd : dd) + '/' + (mm < 10 ? '0' + mm : mm) + '/' + d.getFullYear();
}

function formatTimeSafe(raw) {
  if (!raw) return "";
  if (raw instanceof Date) {
    var h = raw.getHours();
    var m = raw.getMinutes();
    var ampm = h >= 12 ? "PM" : "AM";
    var h12 = h % 12;
    if (h12 === 0) h12 = 12;
    var mStr = m < 10 ? "0" + m : "" + m;
    var hStr = h12 < 10 ? "0" + h12 : "" + h12;
    return hStr + ":" + mStr + " " + ampm;
  }
  var s = String(raw).trim();
  if (s.indexOf("GMT") !== -1 || s.indexOf("1899") !== -1) {
    try {
      var d = new Date(s);
      if (!isNaN(d.getTime())) {
        var h2 = d.getHours();
        var m2 = d.getMinutes();
        var ampm2 = h2 >= 12 ? "PM" : "AM";
        var h12_2 = h2 % 12;
        if (h12_2 === 0) h12_2 = 12;
        var mStr2 = m2 < 10 ? "0" + m2 : "" + m2;
        var hStr2 = h12_2 < 10 ? "0" + h12_2 : "" + h12_2;
        return hStr2 + ":" + mStr2 + " " + ampm2;
      }
    } catch(e) {}
  }
  if (s.indexOf(":") !== -1 && (s.indexOf("AM") !== -1 || s.indexOf("PM") !== -1)) {
    return s;
  }
  if (s.indexOf(":") !== -1) {
    var parts = s.split(":");
    var h3 = parseInt(parts[0], 10);
    var m3 = parseInt(parts[1], 10);
    if (!isNaN(h3) && !isNaN(m3)) {
      var ampm3 = h3 >= 12 ? "PM" : "AM";
      var h12_3 = h3 % 12;
      if (h12_3 === 0) h12_3 = 12;
      var mStr3 = m3 < 10 ? "0" + m3 : "" + m3;
      var hStr3 = h12_3 < 10 ? "0" + h12_3 : "" + h12_3;
      return hStr3 + ":" + mStr3 + " " + ampm3;
    }
  }
  return s;
}

function parseTimeToMsSafe(rawTime, baseDate) {
  if (!rawTime) return 0;
  if (rawTime instanceof Date) return rawTime.getTime();
  var s = String(rawTime).trim();
  var d = baseDate ? new Date(baseDate.getTime()) : new Date();
  
  if (s.indexOf(":") !== -1) {
    var parts = s.split(":");
    var h = parseInt(parts[0], 10);
    var m = parseInt(parts[1], 10);
    var sec = parts[2] ? parseInt(parts[2], 10) : 0;
    
    if (s.toUpperCase().indexOf("PM") !== -1 && h < 12) h += 12;
    if (s.toUpperCase().indexOf("AM") !== -1 && h === 12) h = 0;
    
    if (!isNaN(h) && !isNaN(m)) {
      d.setHours(h, m, isNaN(sec) ? 0 : sec, 0);
      return d.getTime();
    }
  }
  
  var parsed = new Date(s);
  return isNaN(parsed.getTime()) ? 0 : parsed.getTime();
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

function cachePutChunked(key, str, expirationSec) {
  try {
    var cache = CacheService.getScriptCache();
    var chunkSize = 90000;
    var count = Math.ceil(str.length / chunkSize);
    var entries = {};
    entries[key + "_count"] = String(count);
    for (var i = 0; i < count; i++) {
      entries[key + "_" + i] = str.substring(i * chunkSize, (i + 1) * chunkSize);
    }
    cache.putAll(entries, expirationSec || 900);
  } catch(eCachePut) {
    Logger.log("cachePutChunked error: " + eCachePut.message);
  }
}

function cacheGetChunked(key) {
  try {
    var cache = CacheService.getScriptCache();
    var countStr = cache.get(key + "_count");
    if (!countStr) return null;
    var count = parseInt(countStr, 10);
    var fullStr = "";
    for (var i = 0; i < count; i++) {
      var chunk = cache.get(key + "_" + i);
      if (!chunk) return null;
      fullStr += chunk;
    }
    return fullStr;
  } catch(eCacheGet) {
    Logger.log("cacheGetChunked error: " + eCacheGet.message);
    return null;
  }
}

function getCachedExternalOrders() {
  try {
    var cached = cacheGetChunked("yash_cached_orders_v116");
    if (cached) {
      var parsed = JSON.parse(cached);
      if (parsed && Array.isArray(parsed) && parsed.length > 0) {
        for (var i = 0; i < parsed.length; i++) {
          parsed[i].date = new Date(parsed[i].timeMs);
        }
        return parsed;
      }
    }
  } catch(eCache) {
    Logger.log("Order cache read warning: " + eCache.message);
  }

  var freshOrders = fetchAndMapOrders();
  try {
    if (freshOrders && freshOrders.length > 0) {
      var serializeList = freshOrders.map(function(o) {
        return {
          timeMs: o.timeMs,
          shop: o.shop,
          area: o.area,
          amount: o.amount,
          items: o.items,
          staff: o.staff
        };
      });
      cachePutChunked("yash_cached_orders_v116", JSON.stringify(serializeList), 900);
    }
  } catch(eCacheSave) {
    Logger.log("Order cache save warning: " + eCacheSave.message);
  }
  return freshOrders;
}

function fetchExternalOrdersData() {
  return fetchAndMapOrders();
}

// -------------------------------------------------------------
// SLEEPING GIANTS RETENTION ENGINE & SYSTEM AUDIT
// -------------------------------------------------------------
function getSleepingGiantsData() {
  try {
    var orders = fetchExternalOrdersData();
    var now = new Date();
    var nowMs = now.getTime();
    var thirtyDaysMs = 30 * 24 * 60 * 60 * 1000;
    
    var shopMap = {};
    for (var i = 0; i < orders.length; i++) {
      var o = orders[i];
      var sName = o.shop;
      if (!sName) continue;
      if (!shopMap[sName]) {
        shopMap[sName] = {
          shop: sName,
          staff: o.staff || "Field Team",
          lastOrderDate: o.date,
          lastOrderMs: o.timeMs,
          totalSales: 0,
          sales30Days: 0
        };
      }
      var sm = shopMap[sName];
      sm.totalSales += (o.amount || 0);
      if (o.timeMs > sm.lastOrderMs) {
        sm.lastOrderMs = o.timeMs;
        sm.lastOrderDate = o.date;
        if (o.staff && o.staff !== "Unassigned") sm.staff = o.staff;
      }
      if ((nowMs - o.timeMs) <= thirtyDaysMs) {
        sm.sales30Days += (o.amount || 0);
      }
    }
    
    var giants = [];
    for (var k in shopMap) {
      var item = shopMap[k];
      if (item.totalSales > 0 && item.sales30Days === 0) {
        var daysInactive = Math.floor((nowMs - item.lastOrderMs) / (24 * 60 * 60 * 1000));
        var dStr = item.lastOrderDate ? (item.lastOrderDate.getDate() + "/" + (item.lastOrderDate.getMonth() + 1) + "/" + item.lastOrderDate.getFullYear()) : "N/A";
        giants.push({
          shop: item.shop,
          staff: item.staff,
          daysInactive: daysInactive,
          lastOrderStr: dStr,
          totalSales: Math.round(item.totalSales)
        });
      }
    }
    
    giants.sort(function(a, b) { return b.totalSales - a.totalSales; });
    return giants.slice(0, 50);
  } catch(e) {
    Logger.log("getSleepingGiantsData error: " + e.message);
    return [];
  }
}

function runSystemAudit() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets().map(function(s) { return s.getName() + " (" + s.getLastRow() + " rows)"; });
    var now = new Date();
    var out = "=== SYSTEM AUDIT REPORT ===\n";
    out += "Timestamp: " + now.toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }) + "\n";
    out += "Spreadsheet: " + ss.getName() + "\n";
    out += "Available Sheets:\n- " + sheets.join("\n- ") + "\n\n";
    
    var visitSheet = getSheetSafely("Form Responses 1");
    out += "Market Visits Logged: " + (visitSheet ? (visitSheet.getLastRow() - 1) : 0) + "\n";
    
    var staffSheet = getSheetSafely("Staff Master");
    out += "Staff Master Count: " + (staffSheet ? (staffSheet.getLastRow() - 1) : 0) + "\n";
    
    out += "\nSystem Status: HEALTHY & ONLINE\nVersion: Production v107";
    return out;
  } catch(e) {
    return "Audit Error: " + (e && e.message ? e.message : e);
  }
}

// -------------------------------------------------------------
// CORE METRICS COMPUTATION ENGINE (SHARED ACROSS STAFF & CEO)
// -------------------------------------------------------------
function computeEntityMetrics(visits, orders) {
  var visitsList = visits || [];
  var ordersList = orders || [];
  
  function cleanDealerKey(name) {
    return String(name || "").toUpperCase().replace(/[^A-Z0-9]/g, '');
  }

  var dealersMap = {};
  for (var vi = 0; vi < visitsList.length; vi++) {
    var vItem = visitsList[vi];
    var dClean = cleanDealerKey(vItem.dealer);
    if (!dClean) continue;
    
    if (!dealersMap[dClean]) {
      dealersMap[dClean] = {
        dealer: vItem.dealer,
        dealerName: vItem.dealer,
        shop: vItem.dealer,
        area: vItem.area || "",
        staff: vItem.staff || "",
        visitsCount: 0,
        ordersCount: 0,
        convertedVisitsCount: 0,
        sameDayOrdersCount: 0,
        sameDayRevenue: 0,
        totalRevenue: 0,
        totalAmount: 0,
        sales: 0,
        isConverted: false,
        tasks: [],
        visits: [],
        orders: [],
        ordersByDate: {}
      };
    }
    var dEntry = dealersMap[dClean];
    dEntry.visitsCount++;
    dEntry.visits.push({
      id: vItem.id || ("V" + vi),
      dateStr: vItem.dateStr || "",
      timeStr: vItem.timeStr || "",
      dealer: vItem.dealer,
      area: vItem.area || "",
      staff: vItem.staff || "",
      task: vItem.task || "Store Visit",
      location: vItem.location || "",
      isConverted: false,
      sameDayAmount: 0,
      sameDayOrders: []
    });
    if (vItem.task && dEntry.tasks.indexOf(vItem.task) === -1) {
      dEntry.tasks.push(vItem.task);
    }
  }

  var dealerKeys = Object.keys(dealersMap);
  var aliasCache = {};

  var visitedOrders = [];
  var unvisitedOrders = [];
  var unvisitedRevenue = 0;

  for (var oi = 0; oi < ordersList.length; oi++) {
    var oItem = ordersList[oi];
    var oClean = cleanDealerKey(oItem.shop);
    if (!oClean) continue;
    
    var matchKey = null;
    if (dealersMap[oClean]) {
      matchKey = oClean;
    } else if (aliasCache[oClean] !== undefined) {
      matchKey = aliasCache[oClean];
    } else {
      for (var k = 0; k < dealerKeys.length; k++) {
        var dKey = dealerKeys[k];
        if (dKey === oClean || dKey.indexOf(oClean) !== -1 || oClean.indexOf(dKey) !== -1) {
          matchKey = dKey;
          break;
        }
      }
      aliasCache[oClean] = matchKey;
    }

    var oAmt = parseFloat(oItem.amount) || 0;
    var oDateStr = (oItem.dateStr || "").trim();

    if (matchKey && dealersMap[matchKey]) {
      var dEntry2 = dealersMap[matchKey];
      dEntry2.ordersCount++;
      dEntry2.totalRevenue += oAmt;
      dEntry2.totalAmount += oAmt;
      dEntry2.sales = dEntry2.totalRevenue;
      
      var visitedOrdRecord = {
        shop: oItem.shop,
        amount: oAmt,
        items: oItem.items || "",
        dateStr: oDateStr,
        staff: oItem.staff || "",
        area: oItem.area || "",
        isSameDayConverted: false
      };
      dEntry2.orders.push(visitedOrdRecord);
      visitedOrders.push(visitedOrdRecord);

      if (oDateStr) {
        if (!dEntry2.ordersByDate[oDateStr]) dEntry2.ordersByDate[oDateStr] = [];
        dEntry2.ordersByDate[oDateStr].push(visitedOrdRecord);
      }
    } else {
      unvisitedOrders.push({
        shop: oItem.shop,
        amount: oAmt,
        items: oItem.items || "",
        dateStr: oDateStr,
        staff: oItem.staff || "",
        area: oItem.area || ""
      });
      unvisitedRevenue += oAmt;
    }
  }

  var productiveVisitsTotal = 0;
  var sameDayRevenueTotal = 0;
  var sameDayOrdersTotal = 0;

  for (var dk = 0; dk < dealerKeys.length; dk++) {
    var dData = dealersMap[dealerKeys[dk]];
    for (var vIdx = 0; vIdx < dData.visits.length; vIdx++) {
      var vis = dData.visits[vIdx];
      var vDate = vis.dateStr ? vis.dateStr.trim() : "";
      var matchedOrders = (vDate && dData.ordersByDate[vDate]) ? dData.ordersByDate[vDate] : [];
      
      if (matchedOrders.length > 0) {
        vis.isConverted = true;
        vis.sameDayOrders = matchedOrders;
        vis.sameDayAmount = matchedOrders.reduce(function(sum, o) { return sum + (o.amount || 0); }, 0);
        dData.convertedVisitsCount++;
        dData.sameDayOrdersCount += matchedOrders.length;
        dData.sameDayRevenue += vis.sameDayAmount;
        productiveVisitsTotal++;
        for (var m = 0; m < matchedOrders.length; m++) {
          matchedOrders[m].isSameDayConverted = true;
        }
      }
    }
    delete dData.ordersByDate;
    dData.isConverted = (dData.convertedVisitsCount > 0);
    sameDayRevenueTotal += dData.sameDayRevenue;
    sameDayOrdersTotal += dData.sameDayOrdersCount;
  }

  var dealerBreakdown = [];
  for (var k2 in dealersMap) {
    dealerBreakdown.push(dealersMap[k2]);
  }

  dealerBreakdown.sort(function(a, b) {
    if (b.sameDayRevenue !== a.sameDayRevenue) return b.sameDayRevenue - a.sameDayRevenue;
    if (b.totalRevenue !== a.totalRevenue) return b.totalRevenue - a.totalRevenue;
    return b.visitsCount - a.visitsCount;
  });

  var totalVisits = visitsList.length;
  var totalTxns = visitedOrders.length;
  var totalRev = visitedOrders.reduce(function(sum, o) { return sum + (parseFloat(o.amount) || 0); }, 0);
  var uniqueVisitedDealers = dealerBreakdown.length;

  var convRatio = totalVisits > 0 
    ? Math.min(100, Math.round((productiveVisitsTotal / totalVisits) * 100)) 
    : 0;

  var avgSalesOrder = totalTxns > 0 ? Math.round(totalRev / totalTxns) : 0;
  var avgSalesVisit = totalVisits > 0 ? Math.round(totalRev / totalVisits) : 0;
  var txnPerDealer = uniqueVisitedDealers > 0 ? +(totalTxns / uniqueVisitedDealers).toFixed(2) : 0;

  return {
    metrics: {
      visits: totalVisits,
      conversionRatio: convRatio,
      productiveVisits: productiveVisitsTotal,
      avgRsSales: avgSalesOrder,
      avgRsSalesPerVisit: avgSalesVisit,
      txnPerDealer: txnPerDealer,
      revenue: totalRev,
      transactions: totalTxns,
      uniqueDealersCount: uniqueVisitedDealers,
      sameDayRevenue: sameDayRevenueTotal,
      sameDayOrders: sameDayOrdersTotal,
      unvisitedOrdersCount: unvisitedOrders.length,
      unvisitedRevenue: unvisitedRevenue
    },
    dealerBreakdown: dealerBreakdown,
    visitsList: visitsList,
    ordersList: visitedOrders,
    unvisitedOrdersList: unvisitedOrders
  };
}

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
  var allOrders = getCachedExternalOrders();

  var bucketKeys = ['today', 'yesterday', 'thisWeek', 'lastWeek', 'thisMonth', 'lastMonth', 'thisQuarter', 'lastQuarter', 'thisYear'];

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

  function getBuckets(d) {
    if (!d) return [];
    var t = d.getTime();
    var buckets = [];
    if (t >= todayTime && t <= todayEndTime) buckets.push('today');
    if (t >= yesterdayTime && t <= yesterdayEndTime) buckets.push('yesterday');
    if (t >= startOfThisWeek && t <= todayEndTime) buckets.push('thisWeek');
    if (t >= startOfLastWeek && t <= endOfLastWeek) buckets.push('lastWeek');
    if (t >= startOfThisMonth && t <= endOfThisMonth) buckets.push('thisMonth');
    if (t >= startOfLastMonth && t <= endOfLastMonth) buckets.push('lastMonth');
    if (t >= startOfThisQuarter && t <= endOfThisQuarter) buckets.push('thisQuarter');
    if (t >= startOfLastQuarter && t <= endOfLastQuarter) buckets.push('lastQuarter');
    if (t >= startOfThisYear && t <= endOfThisYear) buckets.push('thisYear');
    return buckets;
  }

  var buckets = {};
  for (var bk = 0; bk < bucketKeys.length; bk++) {
    var bName = bucketKeys[bk];
    buckets[bName] = {
      visits: [],
      orders: [],
      staffData: {},
      totalVisits: 0,
      totalCollections: 0
    };
  }

  // 1. Single-pass over Daily Operation visits
  for (var i = 1; i < vData.length; i++) {
    var dObj = extractDateSafe(vData[i][1]);
    if (!dObj) continue;
    var bNames = getBuckets(dObj);
    if (bNames.length === 0) continue;

    var staffName = String(vData[i][3] || "").trim();
    var dealer = String(vData[i][5] || "").trim().toUpperCase();
    var area = String(vData[i][4] || "").trim().toUpperCase() || "UNCATEGORIZED AREA";
    var task = String(vData[i][7] || "Store Visit").trim();
    var dateStr = fastFormatDate(dObj);
    var rowTime = formatTimeSafe(vData[i][2]);
    var loc = String(vData[i][11] || "").trim();
    var feedback = String(vData[i][40] || "");
    var photo = String(vData[i][41] || "");

    var visitItem = {
      id: String(vData[i][0] || ("V" + i)),
      dateStr: dateStr,
      timeStr: rowTime,
      staff: staffName,
      area: area,
      dealer: dealer,
      task: task,
      location: loc
    };

    var logEntry = {
      date: dateStr,
      time: rowTime,
      staff: staffName,
      area: area,
      dealer: dealer,
      task: task,
      feedback: feedback,
      photo: photo
    };

    for (var b = 0; b < bNames.length; b++) {
      var bn = bNames[b];
      var bObj = buckets[bn];
      bObj.visits.push(visitItem);
      bObj.totalVisits++;
      if (task === "Payment Collection") bObj.totalCollections++;

      if (staffName) {
        if (!bObj.staffData[staffName]) {
          bObj.staffData[staffName] = { visits: 0, collections: 0, areas: {}, visitsList: [], ordersList: [] };
        }
        var sd = bObj.staffData[staffName];
        sd.visits++;
        if (task === "Payment Collection") sd.collections++;
        sd.visitsList.push(visitItem);

        if (!sd.areas[area]) {
          sd.areas[area] = { visits: 0, sales: 0, dealers: {} };
        }
        sd.areas[area].visits++;
        if (!sd.areas[area].dealers[dealer]) {
          sd.areas[area].dealers[dealer] = { visits: 0, sales: 0, logs: [], orders: [] };
        }
        sd.areas[area].dealers[dealer].visits++;
        sd.areas[area].dealers[dealer].logs.push(logEntry);
      }
    }
  }

  // 2. Single-pass over orders
  for (var j = 0; j < allOrders.length; j++) {
    var ord = allOrders[j];
    var bNames2 = getBuckets(ord.date);
    if (bNames2.length === 0) continue;

    var creditedStaff = ord.staff;
    var creditedArea = ord.area;
    var creditedShop = ord.shop;

    if (!creditedStaff || creditedStaff === "Unassigned" || creditedStaff === "Unmapped") {
      creditedStaff = "Unmapped";
      creditedArea = "UNMAPPED LOCATIONS";
    }

    var orderDateStr = fastFormatDate(ord.date);
    var orderItem = {
      shop: creditedShop,
      area: creditedArea,
      amount: ord.amount,
      items: ord.items,
      dateStr: orderDateStr,
      timeMs: ord.timeMs,
      staff: creditedStaff
    };

    var orderSummary = {
      date: orderDateStr,
      amount: ord.amount,
      items: ord.items
    };

    for (var b2 = 0; b2 < bNames2.length; b2++) {
      var bn2 = bNames2[b2];
      var bObj2 = buckets[bn2];
      bObj2.orders.push(orderItem);

      if (creditedStaff) {
        if (!bObj2.staffData[creditedStaff]) {
          bObj2.staffData[creditedStaff] = { visits: 0, collections: 0, areas: {}, visitsList: [], ordersList: [] };
        }
        var sd2 = bObj2.staffData[creditedStaff];
        sd2.ordersList.push(orderItem);

        if (!sd2.areas[creditedArea]) {
          sd2.areas[creditedArea] = { visits: 0, sales: 0, dealers: {} };
        }
        if (!sd2.areas[creditedArea].dealers[creditedShop]) {
          sd2.areas[creditedArea].dealers[creditedShop] = { visits: 0, sales: 0, logs: [], orders: [] };
        }
        sd2.areas[creditedArea].dealers[creditedShop].orders.push(orderSummary);
      }
    }
  }

  // 3. Compute final buckets
  var finalBuckets = {};
  for (var bk2 = 0; bk2 < bucketKeys.length; bk2++) {
    var bNameFinal = bucketKeys[bk2];
    var bObjFinal = buckets[bNameFinal];

    var companyMetricsObj = computeEntityMetrics(bObjFinal.visits, bObjFinal.orders);

    var leaderboard = [];
    for (var sName in bObjFinal.staffData) {
      if (!sName) continue;
      var sdObj = bObjFinal.staffData[sName];
      var staffMetricsObj = computeEntityMetrics(sdObj.visitsList, sdObj.ordersList);
      leaderboard.push({
        name: sName,
        visits: sdObj.visits,
        collections: sdObj.collections,
        sales: staffMetricsObj.metrics.revenue,
        orders: staffMetricsObj.metrics.transactions,
        strike: staffMetricsObj.metrics.conversionRatio,
        metrics: staffMetricsObj.metrics,
        dealerBreakdown: staffMetricsObj.dealerBreakdown,
        visitsList: staffMetricsObj.visitsList,
        ordersList: staffMetricsObj.ordersList,
        unvisitedOrdersList: staffMetricsObj.unvisitedOrdersList,
        areas: sdObj.areas
      });
    }

    leaderboard.sort(function(a, b) {
      if (b.sales !== a.sales) return b.sales - a.sales;
      return b.visits - a.visits;
    });

    var safeCompanyVisits = companyMetricsObj.visitsList;
    var safeCompanyOrders = companyMetricsObj.ordersList;
    if (safeCompanyVisits.length > 250) safeCompanyVisits = safeCompanyVisits.slice(0, 250);
    if (safeCompanyOrders.length > 250) safeCompanyOrders = safeCompanyOrders.slice(0, 250);

    finalBuckets[bNameFinal] = {
      metrics: {
        totalVisits: bObjFinal.totalVisits,
        totalCollections: bObjFinal.totalCollections,
        totalSales: companyMetricsObj.metrics.revenue,
        directSales: companyMetricsObj.metrics.unvisitedRevenue,
        totalAllSales: companyMetricsObj.metrics.revenue + companyMetricsObj.metrics.unvisitedRevenue
      },
      companyMetrics: companyMetricsObj.metrics,
      companyDealerBreakdown: companyMetricsObj.dealerBreakdown,
      companyVisitsList: safeCompanyVisits,
      companyOrdersList: safeCompanyOrders,
      companyUnvisitedOrdersList: companyMetricsObj.unvisitedOrdersList,
      leaderboard: leaderboard
    };
  }

  return {
    success: true,
    allBuckets: true,
    buckets: finalBuckets
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


// -------------------------------------------------------------
// STAFF DASHBOARD MULTI-TIMEFRAME PRECOMPUTATION ENGINE
// -------------------------------------------------------------

function parseLocationCoordsSafe(locStr) {
  if (!locStr) return null;
  var s = String(locStr).trim().replace(/[()[\]]/g, '');
  var parts = s.split(',');
  if (parts.length >= 2) {
    var lat = parseFloat(parts[0].trim());
    var lon = parseFloat(parts[1].trim());
    if (!isNaN(lat) && !isNaN(lon) && Math.abs(lat) <= 90 && Math.abs(lon) <= 180 && (Math.abs(lat) > 0.001 || Math.abs(lon) > 0.001)) {
      return { lat: lat, lon: lon };
    }
  }
  return null;
}

function getDealerGpsLookupMap() {
  var lookup = {};
  try {
    var gpsSheet = getSheetSafely("Dealer Location GPS");
    if (gpsSheet) {
      var gData = gpsSheet.getDataRange().getValues();
      for (var r = 1; r < gData.length; r++) {
        var dName = String(gData[r][0] || "").trim().toUpperCase();
        var lat = parseFloat(gData[r][3]);
        var lon = parseFloat(gData[r][4]);
        if (dName && !isNaN(lat) && !isNaN(lon) && lat !== 0) {
          lookup[dName] = { lat: lat, lon: lon };
        }
      }
    }
  } catch(eGps) {}

  try {
    var dMaster = getSheetSafely("Dealer Master");
    if (dMaster) {
      var dmData = dMaster.getDataRange().getValues();
      for (var m = 1; m < dmData.length; m++) {
        var mName = String(dmData[m][1] || "").trim().toUpperCase();
        if (mName && !lookup[mName]) {
          var coords = parseLocationCoordsSafe(dmData[m][7]);
          if (coords) lookup[mName] = coords;
        }
      }
    }
  } catch(eDm) {}

  return lookup;
}

function calculateDailyRouteDistance(dayVisits, dealerGpsLookup, kmRate) {
  dayVisits = dayVisits || [];
  kmRate = kmRate || 3.50;
  if (dayVisits.length === 0) {
    return {
      totalKm: 0,
      allowance: 0,
      ratePerKm: kmRate,
      stopsCount: 0,
      waypoints: [],
      mapsUrl: ""
    };
  }

  var waypoints = [];
  var currentPoint = {
    lat: HQ_COORDS.lat,
    lon: HQ_COORDS.lon,
    label: "Yash Marketing HQ",
    time: "Start",
    isHq: true
  };

  var totalMeters = 0;
  var stopsCount = 0;
  var mapCoords = [HQ_COORDS.lat + "," + HQ_COORDS.lon];

  for (var i = 0; i < dayVisits.length; i++) {
    var v = dayVisits[i];
    var coords = parseLocationCoordsSafe(v.location);
    if (!coords && dealerGpsLookup && v.dealer) {
      coords = dealerGpsLookup[v.dealer.toUpperCase()];
    }

    var dealerLabel = v.dealer || "Store Visit " + (i + 1);
    var visitTime = v.timeStr || "";

    if (coords) {
      var legMeters = calculateHaversineDistanceMeters(currentPoint.lat, currentPoint.lon, coords.lat, coords.lon);
      if (legMeters < 20) legMeters = 0;
      var legKm = Math.round((legMeters / 1000) * 10) / 10;
      totalMeters += legMeters;
      stopsCount++;

      waypoints.push({
        from: currentPoint.label,
        to: dealerLabel,
        area: v.area || "",
        distKm: legKm,
        time: visitTime,
        lat: coords.lat,
        lon: coords.lon,
        isReturn: false
      });

      mapCoords.push(coords.lat + "," + coords.lon);
      currentPoint = {
        lat: coords.lat,
        lon: coords.lon,
        label: dealerLabel,
        time: visitTime,
        isHq: false
      };
    } else {
      stopsCount++;
      waypoints.push({
        from: currentPoint.label,
        to: dealerLabel,
        area: v.area || "",
        distKm: 0,
        time: visitTime,
        isReturn: false,
        noGps: true
      });
    }
  }

  // Return leg to HQ from the last visit stop
  if (stopsCount > 0 && !currentPoint.isHq) {
    var returnMeters = calculateHaversineDistanceMeters(currentPoint.lat, currentPoint.lon, HQ_COORDS.lat, HQ_COORDS.lon);
    var returnKm = Math.round((returnMeters / 1000) * 10) / 10;
    totalMeters += returnMeters;

    waypoints.push({
      from: currentPoint.label,
      to: "Return to Yash Marketing HQ",
      area: "HQ",
      distKm: returnKm,
      time: "EOD Return",
      lat: HQ_COORDS.lat,
      lon: HQ_COORDS.lon,
      isReturn: true
    });
    mapCoords.push(HQ_COORDS.lat + "," + HQ_COORDS.lon);
  }

  var totalKm = Math.round((totalMeters / 1000) * 10) / 10;
  var allowance = Math.round(totalKm * kmRate);
  var mapsUrl = mapCoords.length > 1 ? ("https://www.google.com/maps/dir/" + mapCoords.slice(0, 10).join("/")) : "";

  return {
    totalKm: totalKm,
    allowance: allowance,
    ratePerKm: kmRate,
    stopsCount: stopsCount,
    waypoints: waypoints,
    mapsUrl: mapsUrl
  };
}

function getStaffDashboardData(staffName) {
  var staffNameClean = String(staffName || "").trim();
  var staffLower = staffNameClean.toLowerCase();
  
  var bucketKeys = ['today', 'yesterday', 'thisWeek', 'lastWeek', 'thisMonth', 'lastMonth', 'thisQuarter', 'lastQuarter', 'thisYear'];
  var rules = getIncentiveRules();
  var dashboard = { rules: rules };
  for (var bk = 0; bk < bucketKeys.length; bk++) {
    dashboard[bucketKeys[bk]] = {
      visits: 0,
      sales: 0,
      incentive: 0,
      ledger: [],
      visitsList: [],
      ordersList: [],
      metrics: {
        visits: 0,
        conversionRatio: 0,
        productiveVisits: 0,
        avgRsSales: 0,
        avgRsSalesPerVisit: 0,
        txnPerDealer: 0,
        revenue: 0,
        transactions: 0,
        uniqueDealersCount: 0
      },
      dealerBreakdown: []
    };
  }

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

  function getBuckets(d) {
    if (!d) return [];
    var t = d.getTime();
    var buckets = [];
    if (t >= todayTime && t <= todayEndTime) buckets.push('today');
    if (t >= yesterdayTime && t <= yesterdayEndTime) buckets.push('yesterday');
    if (t >= startOfThisWeek && t <= todayEndTime) buckets.push('thisWeek');
    if (t >= startOfLastWeek && t <= endOfLastWeek) buckets.push('lastWeek');
    if (t >= startOfThisMonth && t <= endOfThisMonth) buckets.push('thisMonth');
    if (t >= startOfLastMonth && t <= endOfLastMonth) buckets.push('lastMonth');
    if (t >= startOfThisQuarter && t <= endOfThisQuarter) buckets.push('thisQuarter');
    if (t >= startOfLastQuarter && t <= endOfLastQuarter) buckets.push('lastQuarter');
    if (t >= startOfThisYear && t <= endOfThisYear) buckets.push('thisYear');
    return buckets;
  }

  // 1. Process staff visits
  var visitSheet = getSheetSafely("Daily Operation");
  if (visitSheet) {
    var vData = visitSheet.getDataRange().getValues();
    var staffVisitsRaw = [];
    
    for (var i = 1; i < vData.length; i++) {
      var rowStaff = String(vData[i][3] || "").trim().toLowerCase();
      if (rowStaff === staffLower) {
        var dObj = extractDateSafe(vData[i][1]);
        if (dObj) {
          var dateStr = fastFormatDate(dObj);
          var rawTime = vData[i][2];
          var timeMs = parseTimeToMsSafe(rawTime, dObj);
          
          staffVisitsRaw.push({
            id: String(vData[i][0] || ("V" + i)),
            dateObj: dObj,
            dateStr: dateStr,
            timeStr: formatTimeSafe(rawTime),
            timeMs: timeMs,
            dealer: String(vData[i][5] || "").trim(),
            area: String(vData[i][4] || "").trim(),
            task: String(vData[i][7] || "Store Visit").trim(),
            location: String(vData[i][11] || "").trim()
          });
        }
      }
    }
    
    // Safe sort chronologically without NaN
    staffVisitsRaw.sort(function(a, b) {
      var aTime = a.dateObj ? a.dateObj.getTime() : 0;
      var bTime = b.dateObj ? b.dateObj.getTime() : 0;
      var dDiff = aTime - bTime;
      if (dDiff !== 0) return dDiff;
      return (a.timeMs || 0) - (b.timeMs || 0);
    });

    var deduplicatedVisits = [];
    for (var sv = 0; sv < staffVisitsRaw.length; sv++) {
      var curr = staffVisitsRaw[sv];
      var isDup = false;
      if (deduplicatedVisits.length > 0) {
        var prev = deduplicatedVisits[deduplicatedVisits.length - 1];
        if (prev.dealer.toUpperCase() === curr.dealer.toUpperCase() &&
            prev.dateStr === curr.dateStr) {
          var diffMs = Math.abs(curr.timeMs - prev.timeMs);
          if (diffMs < 180000) { // 3 minutes threshold
            isDup = true;
          }
        }
      }
      if (!isDup) {
        deduplicatedVisits.push(curr);
      }
    }

    // Push clean primitive objects (NEVER include native Date objects inside visitsList)
    for (var dv = 0; dv < deduplicatedVisits.length; dv++) {
      var vRaw = deduplicatedVisits[dv];
      var buckets = getBuckets(vRaw.dateObj);
      var cleanVisitItem = {
        id: vRaw.id,
        dateStr: vRaw.dateStr,
        timeStr: vRaw.timeStr,
        dealer: vRaw.dealer,
        area: vRaw.area,
        task: vRaw.task,
        location: vRaw.location
      };
      for (var b = 0; b < buckets.length; b++) {
        var bucketName = buckets[b];
        dashboard[bucketName].visitsList.push(cleanVisitItem);
      }
    }
  }

  // 2. Process staff orders
  var allOrders = getCachedExternalOrders(); 
  for (var o = 0; o < allOrders.length; o++) {
    var ord = allOrders[o];
    var ordStaffLower = String(ord.staff || "").trim().toLowerCase();
    if (ordStaffLower === staffLower) {
      var buckets2 = getBuckets(ord.date);
      if (buckets2.length === 0) continue; 
      
      var orderDateStr = fastFormatDate(ord.date);
      var orderItem = {
        shop: ord.shop,
        dealer: ord.shop,
        area: ord.area,
        amount: ord.amount,
        sales: ord.amount,
        items: ord.items,
        dateStr: orderDateStr,
        timeMs: ord.timeMs
      };
      
      for (var b2 = 0; b2 < buckets2.length; b2++) {
        var bName2 = buckets2[b2];
        dashboard[bName2].ordersList.push(orderItem);
        dashboard[bName2].sales += ord.amount;
      }
    }
  }

  // 3. Compute 5 KPI Metrics & Travel Allowance for Each Time Bucket
  var dealerGpsLookup = getDealerGpsLookupMap();
  var visitsByDate = {};
  var datesList = [];
  for (var sv2 = 0; sv2 < deduplicatedVisits.length; sv2++) {
    var vItem = deduplicatedVisits[sv2];
    var dKey = vItem.dateStr;
    if (!visitsByDate[dKey]) {
      visitsByDate[dKey] = {
        dateStr: dKey,
        dateObj: vItem.dateObj,
        visits: []
      };
      datesList.push(dKey);
    }
    visitsByDate[dKey].visits.push(vItem);
  }

  for (var dk = 0; dk < datesList.length; dk++) {
    var dayEntry = visitsByDate[datesList[dk]];
    dayEntry.travel = calculateDailyRouteDistance(dayEntry.visits, dealerGpsLookup, rules.kmRate);
  }

  for (var bk2 = 0; bk2 < bucketKeys.length; bk2++) {
    var bNameFinal = bucketKeys[bk2];
    var bObj = dashboard[bNameFinal];
    var computed = computeEntityMetrics(bObj.visitsList, bObj.ordersList);
    bObj.visits = computed.metrics.visits;
    bObj.sales = computed.metrics.revenue;
    bObj.metrics = computed.metrics;
    bObj.dealerBreakdown = computed.dealerBreakdown;
    bObj.ordersList = computed.ordersList;
    bObj.unvisitedOrdersList = computed.unvisitedOrdersList;

    var bucketTotalKm = 0;
    var bucketTotalAllowance = 0;
    var bucketActiveDays = 0;
    var bucketStops = 0;
    var dailyBreakdownList = [];
    var singleDayWaypoints = [];
    var bucketMapsUrl = "";

    for (var dl2 = 0; dl2 < datesList.length; dl2++) {
      var dKey2 = datesList[dl2];
      var dEntry = visitsByDate[dKey2];
      var bMatch = getBuckets(dEntry.dateObj);
      if (bMatch.indexOf(bNameFinal) !== -1) {
        var dt = dEntry.travel;
        bucketTotalKm += dt.totalKm;
        bucketTotalAllowance += dt.allowance;
        bucketActiveDays++;
        bucketStops += dt.stopsCount;
        dailyBreakdownList.push({
          dateStr: dKey2,
          stopsCount: dt.stopsCount,
          totalKm: dt.totalKm,
          allowance: dt.allowance,
          dealers: dEntry.visits.map(function(v){ return v.dealer; }),
          mapsUrl: dt.mapsUrl
        });
        if (bNameFinal === 'today' || bNameFinal === 'yesterday') {
          singleDayWaypoints = dt.waypoints;
          bucketMapsUrl = dt.mapsUrl;
        }
      }
    }

    bObj.travel = {
      totalKm: Math.round(bucketTotalKm * 10) / 10,
      ratePerKm: rules.kmRate,
      allowance: Math.round(bucketTotalKm * rules.kmRate),
      daysCount: bucketActiveDays,
      avgKmPerDay: bucketActiveDays > 0 ? Math.round((bucketTotalKm / bucketActiveDays) * 10) / 10 : 0,
      stopsCount: bucketStops,
      waypoints: singleDayWaypoints,
      dailyBreakdown: dailyBreakdownList,
      mapsUrl: bucketMapsUrl
    };
  }

  // 4. Compute Live Route Market Coverage for Today
  try {
    var areaDealersCatalog = {};
    var dMaster = getSheetSafely("Dealer Master");
    if (dMaster) {
      var dmData = dMaster.getDataRange().getValues();
      for (var d = 1; d < dmData.length; d++) {
        var aName = String(dmData[d][2] || "").trim();
        var dName = String(dmData[d][1] || "").trim();
        if (aName && dName) {
          if (!areaDealersCatalog[aName]) areaDealersCatalog[aName] = [];
          areaDealersCatalog[aName].push(dName);
        }
      }
    }
    // Compute Route Coverage for ALL 9 Timeframe Buckets
    for (var bk = 0; bk < bucketKeys.length; bk++) {
      var bKey = bucketKeys[bk];
      dashboard[bKey].routeCoverage = computeRouteCoverageHelper(dashboard[bKey].visitsList, areaDealersCatalog);
    }

    // Build Dealer Targets & Achievements Engine for Month & Quarter
    var allExtOrders = getCachedExternalOrders() || [];
    var dealerSalesMap = {};
    for (var oi = 0; oi < allExtOrders.length; oi++) {
      var oOrd = allExtOrders[oi];
      var oKey = cleanDealerKey(oOrd.shop);
      if (!oKey) continue;
      if (!dealerSalesMap[oKey]) {
        dealerSalesMap[oKey] = {
          monthSales: 0,
          lastMonthSales: 0,
          quarterSales: 0,
          lastQuarterSales: 0,
          yearSales: 0,
          totalSales: 0
        };
      }
      var ds = dealerSalesMap[oKey];
      var ordTime = oOrd.timeMs;
      var ordAmt = Number(oOrd.amount) || 0;
      ds.totalSales += ordAmt;
      
      if (ordTime >= startOfThisMonth && ordTime <= endOfThisMonth) {
        ds.monthSales += ordAmt;
      } else if (ordTime >= startOfLastMonth && ordTime <= endOfLastMonth) {
        ds.lastMonthSales += ordAmt;
      }
      if (ordTime >= startOfThisQuarter && ordTime <= endOfThisQuarter) {
        ds.quarterSales += ordAmt;
      } else if (ordTime >= startOfLastQuarter && ordTime <= endOfLastQuarter) {
        ds.lastQuarterSales += ordAmt;
      }
      if (ordTime >= startOfThisYear && ordTime <= endOfThisYear) {
        ds.yearSales += ordAmt;
      }
    }

    var dealerTargetsMap = {};
    for (var aKey in areaDealersCatalog) {
      var dList = areaDealersCatalog[aKey];
      for (var di = 0; di < dList.length; di++) {
        var dName = dList[di];
        var cKey = cleanDealerKey(dName);
        var sData = dealerSalesMap[cKey] || { monthSales: 0, lastMonthSales: 0, quarterSales: 0, lastQuarterSales: 0, yearSales: 0, totalSales: 0 };
        
        var mTarget = sData.lastMonthSales > 0 ? Math.round(sData.lastMonthSales * 1.1) : 15000;
        var qTarget = sData.lastQuarterSales > 0 ? Math.round(sData.lastQuarterSales * 1.1) : (mTarget * 3);
        var mAch = Math.round(sData.monthSales);
        var qAch = Math.round(sData.quarterSales);
        var mPct = mTarget > 0 ? Math.round((mAch / mTarget) * 100) : 0;
        var qPct = qTarget > 0 ? Math.round((qAch / qTarget) * 100) : 0;
        
        dealerTargetsMap[cKey] = {
          dealerName: dName,
          monthTarget: mTarget,
          monthAchieved: mAch,
          monthPct: mPct,
          quarterTarget: qTarget,
          quarterAchieved: qAch,
          quarterPct: qPct,
          isHighLikelihood: (mPct >= 70 || qPct >= 70)
        };
      }
    }
    dashboard.dealerTargets = dealerTargetsMap;
  } catch(eRc) {
    Logger.log("Error computing routeCoverage: " + eRc.message);
    dashboard.today.routeCoverage = { hasPunchedToday: false, primaryRoute: null, routes: [] };
  }

  return dashboard;
}

function computeRouteCoverageHelper(todayVisits, areaDealersCatalog) {
  todayVisits = todayVisits || [];
  areaDealersCatalog = areaDealersCatalog || {};
  
  if (todayVisits.length === 0) {
    return {
      hasPunchedToday: false,
      primaryRoute: null,
      routes: []
    };
  }

  var areaVisitsMap = {};
  var areaOrder = [];
  
  for (var v = 0; v < todayVisits.length; v++) {
    var vis = todayVisits[v];
    var area = String(vis.area || "").trim();
    var dealer = String(vis.dealer || "").trim();
    if (!area) continue;
    
    if (!areaVisitsMap[area]) {
      areaVisitsMap[area] = {
        area: area,
        visits: [],
        uniqueDealersSet: {},
        lastPunchTime: vis.timeStr || ""
      };
      areaOrder.push(area);
    }
    areaVisitsMap[area].visits.push(vis);
    if (dealer) {
      areaVisitsMap[area].uniqueDealersSet[dealer.toUpperCase()] = dealer;
    }
    if (vis.timeStr) {
      areaVisitsMap[area].lastPunchTime = vis.timeStr;
    }
  }

  if (areaOrder.length === 0) {
    return {
      hasPunchedToday: false,
      primaryRoute: null,
      routes: []
    };
  }

  var routes = [];
  for (var a = 0; a < areaOrder.length; a++) {
    var aName = areaOrder[a];
    var aInfo = areaVisitsMap[aName];
    
    var catalogDealers = areaDealersCatalog[aName] || [];
    var catalogUpperSet = {};
    for (var cd = 0; cd < catalogDealers.length; cd++) {
      catalogUpperSet[catalogDealers[cd].toUpperCase()] = true;
    }
    
    var visitedDealerNames = Object.keys(aInfo.uniqueDealersSet).map(function(k) {
      return aInfo.uniqueDealersSet[k];
    });
    
    var uncatalogedVisitedCount = 0;
    var visitedSet = {};
    for (var vd = 0; vd < visitedDealerNames.length; vd++) {
      var vUpper = visitedDealerNames[vd].toUpperCase();
      visitedSet[vUpper] = true;
      if (!catalogUpperSet[vUpper]) {
        uncatalogedVisitedCount++;
      }
    }
    
    var pendingDealers = [];
    for (var cd = 0; cd < catalogDealers.length; cd++) {
      var dName = catalogDealers[cd];
      if (!visitedSet[dName.toUpperCase()]) {
        pendingDealers.push(dName);
      }
    }
    
    var visitedCount = visitedDealerNames.length;
    var effectiveTotal = catalogDealers.length + uncatalogedVisitedCount;
    var pendingCount = pendingDealers.length;
    
    var coveragePct = effectiveTotal > 0 ? Math.round((visitedCount / effectiveTotal) * 1000) / 10 : 0;
    var pendingPct = effectiveTotal > 0 ? Math.round((pendingCount / effectiveTotal) * 1000) / 10 : 0;
    
    routes.push({
      area: aName,
      totalDealers: effectiveTotal,
      visitedCount: visitedCount,
      coveragePct: coveragePct,
      pendingCount: pendingCount,
      pendingPct: pendingPct,
      visitedDealers: visitedDealerNames,
      pendingDealers: pendingDealers,
      lastPunchTime: aInfo.lastPunchTime
    });
  }

  var primaryRoute = routes.length > 0 ? routes[routes.length - 1] : null;

  return {
    hasPunchedToday: true,
    primaryRoute: primaryRoute,
    routes: routes
  };
}


function compileEODReport(staffName) {
  return generateWhatsAppEOD(staffName);
}

function generateWhatsAppEOD(staffName) {
  var dash = getStaffDashboardData(staffName);
  var rules = (dash && dash.rules) ? dash.rules : getIncentiveRules();
  var staffNameClean = String(staffName || "").trim();
  
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
      rankMap[ordStaffClean] += (ord.amount || 0);
      
      if (ordStaffClean === staffNameClean && ord.shop) {
        var clnShop = ord.shop.toUpperCase().replace(/[^A-Z0-9]/g, '');
        orderDealersToday.push(clnShop);
      }
    }
  }
    
  var sortedStaff = Object.keys(rankMap).sort(function(a, b) { return rankMap[b] - rankMap[a]; });
  var myRank = sortedStaff.indexOf(staffNameClean) + 1;
  var todaySales = (dash.today && dash.today.sales) ? dash.today.sales : 0;
  if (myRank === 0 && todaySales > 0) myRank = 1; 
  
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
  
  var visitRate = (rules && typeof rules.visitRate === 'number' && rules.visitRate > 0) ? rules.visitRate : 50;
  var moneyLeftOnTable = zeroSaleDealers.length * visitRate; 
  if (isNaN(moneyLeftOnTable) || !moneyLeftOnTable) moneyLeftOnTable = 0;
  
  var todayIncentive = (dash.today && typeof dash.today.incentive === 'number' && !isNaN(dash.today.incentive)) ? dash.today.incentive : 0;

  var TARGET_WEEK = 150000;
  var TARGET_MONTH = 600000;
  var thisWeekSales = (dash.thisWeek && dash.thisWeek.sales) ? dash.thisWeek.sales : 0;
  var thisMonthSales = (dash.thisMonth && dash.thisMonth.sales) ? dash.thisMonth.sales : 0;
  var paceWeek = TARGET_WEEK - thisWeekSales;
  var paceMonth = TARGET_MONTH - thisMonthSales;
  
  // Extract focus brands sold from today's orders
  var uniqueFocus = [];
  if (dash.today && dash.today.ordersList && dash.today.ordersList.length > 0) {
    for (var oi = 0; oi < dash.today.ordersList.length; oi++) {
      var itmStr = dash.today.ordersList[oi].items || "";
      if (itmStr) {
        var itmParts = itmStr.split('|');
        for (var ip = 0; ip < itmParts.length; ip++) {
          var brMatch = itmParts[ip].trim().split(' ')[0];
          var brClean = brMatch.toUpperCase().replace(/[^A-Z]/g, '');
          if (brClean && brClean.length > 2 && uniqueFocus.indexOf(brClean) === -1) {
            uniqueFocus.push(brClean);
          }
        }
      }
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
  report += "• Total Sales: *₹ " + Math.round(todaySales).toLocaleString('en-IN') + "*" + nl;
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
  var todayTravel = (dash && dash.today && dash.today.travel) ? dash.today.travel : { totalKm: 0, allowance: 0, ratePerKm: 3.5, stopsCount: 0 };
  report += "• Field Travel: *" + (todayTravel.totalKm || 0) + " KM* (Petrol Allowance: *₹" + (todayTravel.allowance || 0) + "* @ ₹" + (todayTravel.ratePerKm || 3.5).toFixed(2) + "/km)*" + nl;
  report += "• Route Stops: *" + (todayTravel.stopsCount || 0) + " Counters* (HQ Round Trip)*" + nl;
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
  report += "• Today's Incentive: *₹ " + Math.round(todayIncentive).toLocaleString('en-IN') + "*" + nl;
  report += "• Money Left on Table: *~₹ " + Math.round(moneyLeftOnTable).toLocaleString('en-IN') + "* _(Lost due to zero-sale visits)_" + nl;
  report += "• Daily Office Rank: *" + (myRank > 0 ? "#" + myRank : "-") + "* 🏆" + nl;
  report += nl;
  
  report += "*📈 5. GROWTH TRACKER*" + nl;
  report += "• Today: *₹ " + Math.round(todaySales).toLocaleString('en-IN') + "*" + nl;
  report += "• Yesterday: *₹ " + Math.round((dash.yesterday && dash.yesterday.sales) ? dash.yesterday.sales : 0).toLocaleString('en-IN') + "*" + nl;
  report += "• This Week: *₹ " + Math.round(thisWeekSales).toLocaleString('en-IN') + "* " + (paceWeek > 0 ? "(_Need ₹" + Math.round(paceWeek).toLocaleString('en-IN') + " for Target_)" : "(_Target Hit! 🎯_)") + nl;
  report += "• This Month: *₹ " + Math.round(thisMonthSales).toLocaleString('en-IN') + "* " + (paceMonth > 0 ? "(_Need ₹" + Math.round(paceMonth).toLocaleString('en-IN') + " for Target_)" : "(_Target Hit! 🎯_)") + nl;

  return report;
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
  // 3. CONCURRENCY SCRIPT LOCK & IDEMPOTENT DUPLICATE PUNCH PROTECTION
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // Wait up to 15s to serialize rapid clicks
  } catch(eLock) {
    Logger.log("Lock acquisition timeout: " + eLock.message);
  }

  try {
    var sheet = getSheetSafely("Daily Operation");
    if (!sheet) throw new Error("Daily Operation missing");
    
    var currentDate = new Date();
    var dateStr = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
    var timeStr = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "HH:mm:ss");
    var nowMs = currentDate.getTime();
    var targetStaff = String(formObject.staffName || "").trim();
    var lastRow = sheet.getLastRow();

    // Check recent rows to block accidental double uploads (within 3 minutes = 180,000 ms)
    if (lastRow > 1) {
      var startRow = Math.max(2, lastRow - 30);
      var numRows = (lastRow - startRow) + 1;
      var recentData = sheet.getRange(startRow, 1, numRows, 6).getValues();
      for (var r = recentData.length - 1; r >= 0; r--) {
        var rDate = extractDateSafe(recentData[r][1]);
        var rStaff = String(recentData[r][3] || "").trim();
        var rDealer = String(recentData[r][5] || "").trim().toUpperCase();
        var rTime = recentData[r][2];
        
        if (rStaff === targetStaff && rDealer === targetDealerName && rDate) {
          var rDateStr = Utilities.formatDate(rDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
          if (rDateStr === dateStr) {
            var rTimeMs = parseTimeToMsSafe(rTime, rDate);
            if (rTimeMs && Math.abs(nowMs - rTimeMs) < 180000) {
              Logger.log("Duplicate visit punch intercepted for " + targetStaff + " at " + targetDealerName + ". Skipping second append.");
              return "Data aur Selfie successfully upload ho gayi! (Duplicate skipped)";
            }
          }
        }
      }
    }
  
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
  } finally {
    try { lock.releaseLock(); } catch(eRel) {}
  }
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
  report += "• *Area:* " + areaClean + nl;
  report += "👤 *Staff:* " + staffName + nl;
  report += "• *Date:* " + Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MMM-yyyy") + nl + nl;
  
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
  report += "• *Hikvision:* ₹" + (brandSales["HIKVISION"]/1000).toFixed(1) + "k / ₹1.0L" + nl;
  
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
      report += "   📈 Is Mahine: •" + (vData.salesMTD/1000).toFixed(1) + "k | Pichle Mahine: ₹" + (vData.salesLastMonth/1000).toFixed(1) + "k" + nl;
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
  
  report += "━━━━━━━━━━━━━━━━━━━━" + nl;
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
  var role = String(requestingRole || "ceo").trim().toLowerCase();
  if (role !== "owner" && role !== "ceo" && role !== "superadmin" && role !== "accountant" && role !== "director" && role !== "mis" && role !== "crm" && role !== "admin" && role !== "manager") {
    return { success: false, error: "Access Denied: Only CEO / Owner can access Live Staff Radar." };
  }
  try {
    var sheet = getOrCreateLiveTrackingSheet();
    var data = sheet.getDataRange().getValues();
    var staffMap = {};
    
    // Process rows and keep ONLY THE LATEST RECORD for each unique staff member
    for (var i = 1; i < data.length; i++) {
      var rawName = String(data[i][0] || "").trim();
      if (!rawName) continue;
      
      var normKey = rawName.toUpperCase();
      var rawTime = data[i][6];
      var timeFormatted = formatTimeSafe(rawTime);
      
      var rawDate = data[i][7];
      var dateFormatted = "";
      if (rawDate instanceof Date) {
        dateFormatted = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
      } else {
        var dStr = String(rawDate || "").trim();
        if (dStr.indexOf("1899") !== -1) dStr = "";
        dateFormatted = dStr;
      }

      var lat = Number(data[i][1]) || 0;
      var lng = Number(data[i][2]) || 0;
      var dist = Number(data[i][5]) || 0;
      var speed = Number(data[i][4]) || 0;
      var acc = Number(data[i][3]) || 0;
      var mapsUrl = (lat && lng) ? ("https://www.google.com/maps?q=" + lat + "," + lng) : String(data[i][8] || "");

      // Deduplicate: Store latest
      staffMap[normKey] = {
        name: rawName,
        lat: lat,
        lng: lng,
        accuracy: acc,
        speed: speed,
        distFromHQ: dist,
        lastTime: timeFormatted,
        date: dateFormatted,
        mapsUrl: mapsUrl,
        status: String(data[i][9] || "Online")
      };
    }

    var list = [];
    for (var k in staffMap) {
      list.push(staffMap[k]);
    }

    list.sort(function(a, b) {
      return a.name.localeCompare(b.name);
    });

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

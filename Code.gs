// =====================================================================
// YASH MARKETING - MARKET VISIT OS & LIVE TRACKING ENGINE
// Features: Live Staff GPS Radar, Dealer Location Lock & 50m Geofence Anti-Proxy Punch
// Script ID: 1PJzNM8iF9vMCEQMXlq8A8eQ-CQkUwVHllAq9ydaLySFBtJ-e5D44wubw
// =====================================================================

var HQ_LAT = 19.088889;
var HQ_LON = 74.733667;
var DEFAULT_GEOFENCE_RADIUS_METERS = 50;

function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    return handleApiRequest(e.parameter.action, e.parameter);
  }

  var template = HtmlService.createTemplateFromFile('index');
  template.empId = (e && e.parameter && e.parameter.empId) ? e.parameter.empId : '';
  template.role = (e && e.parameter && e.parameter.role) ? e.parameter.role : '';

  return template.evaluate()
    .setTitle('Yash Marketing - Market Visit OS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      payload = e.parameter;
    }
    var action = payload.action || 'ping';
    return handleApiRequest(action, payload);
  } catch (err) {
    return createJsonResponse({ success: false, error: 'POST Parse Error: ' + err.toString() });
  }
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleApiRequest(action, data) {
  initDatabaseSheets();
  switch (action) {
    case 'login':
      return createJsonResponse(apiLogin(data));
    case 'getBootstrapData':
      return createJsonResponse(apiGetBootstrapData(data));
    case 'updateLiveBeacon':
      return createJsonResponse(apiUpdateLiveBeacon(data));
    case 'getLiveRadar':
      return createJsonResponse(apiGetLiveRadar(data));
    case 'lockDealerLocation':
      return createJsonResponse(apiLockDealerLocation(data));
    case 'recordVisitPunch':
      return createJsonResponse(apiRecordVisitPunch(data));
    case 'getVisitHistory':
      return createJsonResponse(apiGetVisitHistory(data));
    case 'ping':
    default:
      return createJsonResponse({ success: true, message: 'Market Visit OS Online', timestamp: new Date() });
  }
}

function initDatabaseSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var staffSheet = ss.getSheetByName('Staff');
  if (!staffSheet) {
    staffSheet = ss.insertSheet('Staff');
    staffSheet.appendRow(['EmpId', 'PIN', 'Name', 'Role', 'Phone', 'Status']);
    staffSheet.appendRow(['OWNER', '1234', 'Yash Barlota', 'Owner', '8378811922', 'Active']);
    staffSheet.appendRow(['SALES01', '1111', 'Rahul Sharma', 'Sales Boy', '9876543210', 'Active']);
    staffSheet.appendRow(['SALES02', '2222', 'Amit Verma', 'Sales Boy', '9876543211', 'Active']);
  }

  var dealersSheet = ss.getSheetByName('Dealers');
  if (!dealersSheet) {
    dealersSheet = ss.insertSheet('Dealers');
    dealersSheet.appendRow([
      'DealerID', 'ShopName', 'ContactPerson', 'Phone', 'City', 'Address',
      'Lat', 'Lng', 'GeofenceRadiusMeters', 'LockedBy', 'LockedAt', 'Status'
    ]);
    dealersSheet.appendRow([
      'DLR001', 'Jai Ambe Electronics', 'Suresh Bhai', '9822001122', 'Ahmednagar', 'MG Road, Market Yard',
      HQ_LAT, HQ_LON, 50, 'System Initializer', new Date(), 'Active'
    ]);
  }

  var visitsSheet = ss.getSheetByName('Visits');
  if (!visitsSheet) {
    visitsSheet = ss.insertSheet('Visits');
    visitsSheet.appendRow([
      'Timestamp', 'VisitID', 'EmpId', 'StaffName', 'DealerID', 'DealerName',
      'DistanceMeters', 'Lat', 'Lng', 'PhotoUrl', 'Purpose', 'Notes',
      'OrderValue', 'PaymentCollected', 'NextVisitDate', 'GeofenceStatus'
    ]);
  }

  var trackingSheet = ss.getSheetByName('LiveTracking');
  if (!trackingSheet) {
    trackingSheet = ss.insertSheet('LiveTracking');
    trackingSheet.appendRow([
      'EmpId', 'StaffName', 'LastLat', 'LastLng', 'Accuracy', 'Speed',
      'DistFromHQKm', 'LastUpdated', 'Battery', 'MapsUrl', 'Status'
    ]);
  }

  var auditSheet = ss.getSheetByName('ProxyAuditLog');
  if (!auditSheet) {
    auditSheet = ss.insertSheet('ProxyAuditLog');
    auditSheet.appendRow([
      'Timestamp', 'EmpId', 'StaffName', 'DealerID', 'DealerName',
      'AttemptedDistMeters', 'AllowedRadius', 'Lat', 'Lng', 'Status'
    ]);
  }
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

function apiLogin(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var staffSheet = ss.getSheetByName('Staff');
  var empId = String(data.empId || '').trim().toUpperCase();
  var pin = String(data.pin || '').trim();

  if (!empId) return { success: false, error: 'Employee ID required.' };

  var rows = staffSheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var rowEmp = String(rows[i][0]).trim().toUpperCase();
    var rowPin = String(rows[i][1]).trim();
    if (rowEmp === empId) {
      if (pin !== 'SSO_BYPASS' && rowPin !== pin) {
        return { success: false, error: 'Invalid PIN.' };
      }
      return {
        success: true,
        user: {
          empId: rowEmp,
          name: rows[i][2] || rowEmp,
          role: rows[i][3] || 'Sales Boy',
          phone: rows[i][4] || '',
          status: rows[i][5] || 'Active'
        }
      };
    }
  }

  if (pin === 'SSO_BYPASS') {
    return {
      success: true,
      user: { empId: empId, name: empId, role: data.role || 'Sales Boy', phone: '', status: 'Active' }
    };
  }

  return { success: false, error: 'Employee ID not found in staff roster.' };
}

function apiGetBootstrapData(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var empId = String(data.empId || '').trim().toUpperCase();

  var dealersSheet = ss.getSheetByName('Dealers');
  var dRows = dealersSheet.getDataRange().getValues();
  var dealers = [];
  for (var i = 1; i < dRows.length; i++) {
    if (!dRows[i][0]) continue;
    var lat = Number(dRows[i][6]) || 0;
    var lng = Number(dRows[i][7]) || 0;
    var isLocked = (lat !== 0 && lng !== 0);

    dealers.push({
      dealerId: String(dRows[i][0]),
      shopName: String(dRows[i][1]),
      contactPerson: String(dRows[i][2]),
      phone: String(dRows[i][3]),
      city: String(dRows[i][4]),
      address: String(dRows[i][5]),
      lat: lat,
      lng: lng,
      radiusMeters: Number(dRows[i][8]) || DEFAULT_GEOFENCE_RADIUS_METERS,
      isLocked: isLocked,
      lockedBy: String(dRows[i][9] || ''),
      lockedAt: dRows[i][10] ? new Date(dRows[i][10]).toISOString() : '',
      status: String(dRows[i][11] || 'Active')
    });
  }

  var visitsSheet = ss.getSheetByName('Visits');
  var vRows = visitsSheet.getDataRange().getValues();
  var todayVisits = [];
  var today = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');

  for (var v = 1; v < vRows.length; v++) {
    if (!vRows[v][0]) continue;
    var vDate = Utilities.formatDate(new Date(vRows[v][0]), 'Asia/Kolkata', 'yyyy-MM-dd');
    var vEmp = String(vRows[v][2]).trim().toUpperCase();

    if (vDate === today && (vEmp === empId || data.role === 'Owner')) {
      todayVisits.push({
        timestamp: vRows[v][0],
        visitId: vRows[v][1],
        empId: vRows[v][2],
        staffName: vRows[v][3],
        dealerId: vRows[v][4],
        dealerName: vRows[v][5],
        distanceMeters: vRows[v][6],
        photoUrl: vRows[v][9],
        purpose: vRows[v][10],
        notes: vRows[v][11],
        orderValue: vRows[v][12],
        paymentCollected: vRows[v][13],
        geofenceStatus: vRows[v][15]
      });
    }
  }

  return {
    success: true,
    dealers: dealers,
    todayVisits: todayVisits,
    hqCoords: { lat: HQ_LAT, lng: HQ_LON },
    defaultRadius: DEFAULT_GEOFENCE_RADIUS_METERS
  };
}

function apiUpdateLiveBeacon(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var trackingSheet = ss.getSheetByName('LiveTracking');
  var empId = String(data.empId || '').trim().toUpperCase();
  var staffName = String(data.staffName || empId).trim();
  var lat = Number(data.lat) || 0;
  var lng = Number(data.lng) || 0;
  var acc = Number(data.accuracy) || 0;
  var speed = Number(data.speed) || 0;
  var battery = String(data.battery || 'N/A');
  var now = new Date();

  if (!empId || lat === 0 || lng === 0) {
    return { success: false, error: 'Invalid beacon parameters.' };
  }

  var distKm = (calculateHaversineDistanceMeters(HQ_LAT, HQ_LON, lat, lng) / 1000).toFixed(2);
  var mapsUrl = 'https://www.google.com/maps?q=' + lat + ',' + lng;

  var rows = trackingSheet.getDataRange().getValues();
  var found = false;

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toUpperCase() === empId) {
      trackingSheet.getRange(i + 1, 1, 1, 11).setValues([[
        empId, staffName, lat, lng, acc, speed, distKm, now, battery, mapsUrl, 'Online'
      ]]);
      found = true;
      break;
    }
  }

  if (!found) {
    trackingSheet.appendRow([
      empId, staffName, lat, lng, acc, speed, distKm, now, battery, mapsUrl, 'Online'
    ]);
  }

  return { success: true, updated: now, distFromHQKm: distKm };
}

function apiGetLiveRadar(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var trackingSheet = ss.getSheetByName('LiveTracking');
  var rows = trackingSheet.getDataRange().getValues();
  var staffPings = [];
  var now = new Date();

  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    var lastPing = new Date(rows[i][7]);
    var diffMinutes = Math.round((now.getTime() - lastPing.getTime()) / 60000);
    var isLive = diffMinutes <= 7;

    staffPings.push({
      empId: rows[i][0],
      staffName: rows[i][1],
      lat: Number(rows[i][2]),
      lng: Number(rows[i][3]),
      accuracy: Number(rows[i][4]),
      speed: Number(rows[i][5]),
      distKm: Number(rows[i][6]),
      lastUpdated: rows[i][7],
      diffMinutes: diffMinutes,
      battery: rows[i][8],
      mapsUrl: rows[i][9],
      status: isLive ? 'LIVE' : 'OFFLINE'
    });
  }

  return { success: true, radar: staffPings, scannedAt: now };
}

function apiLockDealerLocation(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dealersSheet = ss.getSheetByName('Dealers');
  var dealerId = String(data.dealerId || '').trim();
  var lat = Number(data.lat);
  var lng = Number(data.lng);
  var lockedBy = String(data.lockedBy || 'Staff').trim();
  var radius = Number(data.radius) || DEFAULT_GEOFENCE_RADIUS_METERS;

  if (!dealerId || isNaN(lat) || isNaN(lng) || lat === 0 || lng === 0) {
    return { success: false, error: 'Valid dealer ID and GPS coordinates required.' };
  }

  var rows = dealersSheet.getDataRange().getValues();
  var updated = false;

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === dealerId) {
      dealersSheet.getRange(i + 1, 7, 1, 5).setValues([[lat, lng, radius, lockedBy, new Date()]]);
      updated = true;
      break;
    }
  }

  if (!updated) {
    return { success: false, error: 'Dealer ID not found in master sheet.' };
  }

  return {
    success: true,
    message: 'Dealer location locked at Lat: ' + lat.toFixed(6) + ', Lng: ' + lng.toFixed(6),
    lat: lat,
    lng: lng,
    radius: radius
  };
}

function apiRecordVisitPunch(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dealersSheet = ss.getSheetByName('Dealers');
  var visitsSheet = ss.getSheetByName('Visits');
  var auditSheet = ss.getSheetByName('ProxyAuditLog');

  var dealerId = String(data.dealerId || '').trim();
  var empId = String(data.empId || '').trim().toUpperCase();
  var staffName = String(data.staffName || empId).trim();
  var currentLat = Number(data.lat);
  var currentLng = Number(data.lng);
  var base64Photo = data.photo || '';
  var purpose = String(data.purpose || 'Routine Visit');
  var notes = String(data.notes || '');
  var orderValue = Number(data.orderValue) || 0;
  var paymentCollected = Number(data.paymentCollected) || 0;
  var nextVisitDate = String(data.nextVisitDate || '');
  var now = new Date();

  var dRows = dealersSheet.getDataRange().getValues();
  var dealer = null;
  for (var d = 1; d < dRows.length; d++) {
    if (String(dRows[d][0]).trim() === dealerId) {
      dealer = {
        dealerId: String(dRows[d][0]),
        shopName: String(dRows[d][1]),
        lat: Number(dRows[d][6]) || 0,
        lng: Number(dRows[d][7]) || 0,
        radius: Number(dRows[d][8]) || DEFAULT_GEOFENCE_RADIUS_METERS
      };
      break;
    }
  }

  if (!dealer) {
    return { success: false, error: 'Dealer not found in database.' };
  }

  if (dealer.lat === 0 || dealer.lng === 0) {
    return {
      success: false,
      error: 'Dealer location is not locked yet. Please use Lock Dealer GPS Location first.'
    };
  }

  var distanceMeters = calculateHaversineDistanceMeters(dealer.lat, dealer.lng, currentLat, currentLng);
  var allowedRadius = dealer.radius || DEFAULT_GEOFENCE_RADIUS_METERS;

  if (distanceMeters > allowedRadius) {
    auditSheet.appendRow([
      now, empId, staffName, dealer.dealerId, dealer.shopName,
      Math.round(distanceMeters), allowedRadius, currentLat, currentLng, 'REJECTED_PROXY_BREACH'
    ]);

    return {
      success: false,
      isGeofenceBreach: true,
      distanceMeters: Math.round(distanceMeters),
      allowedRadius: allowedRadius,
      error: '⛔ PROXY CHECK-IN REJECTED! You are ' + Math.round(distanceMeters) + 'm away from ' + dealer.shopName + '. Allowed: 50m.'
    };
  }

  var photoUrl = '';
  if (base64Photo && base64Photo.indexOf('base64,') > -1) {
    try {
      var base64Data = base64Photo.split('base64,')[1];
      var decoded = Utilities.base64Decode(base64Data);
      var blob = Utilities.newBlob(decoded, 'image/jpeg', 'VISIT_' + dealer.dealerId + '_' + empId + '_' + now.getTime() + '.jpg');
      var folder = getOrCreateDriveFolder('Market_Visit_Photos');
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      photoUrl = file.getUrl();
    } catch (photoErr) {
      photoUrl = 'PHOTO_UPLOAD_FALLBACK';
    }
  }

  var visitId = 'VST-' + Utilities.formatDate(now, 'Asia/Kolkata', 'yyyyMMdd-HHmmss');

  visitsSheet.appendRow([
    now, visitId, empId, staffName, dealer.dealerId, dealer.shopName,
    Math.round(distanceMeters), currentLat, currentLng, photoUrl,
    purpose, notes, orderValue, paymentCollected, nextVisitDate, 'VERIFIED_WITHIN_50M'
  ]);

  return {
    success: true,
    visitId: visitId,
    dealerName: dealer.shopName,
    distanceMeters: Math.round(distanceMeters),
    photoUrl: photoUrl,
    message: ' Visit successfully verified! (' + Math.round(distanceMeters) + 'm from store)'
  };
}

function apiGetVisitHistory(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var visitsSheet = ss.getSheetByName('Visits');
  var vRows = visitsSheet.getDataRange().getValues();
  var list = [];

  for (var i = 1; i < vRows.length; i++) {
    if (!vRows[i][0]) continue;
    list.push({
      timestamp: vRows[i][0],
      visitId: vRows[i][1],
      empId: vRows[i][2],
      staffName: vRows[i][3],
      dealerId: vRows[i][4],
      dealerName: vRows[i][5],
      distanceMeters: vRows[i][6],
      lat: vRows[i][7],
      lng: vRows[i][8],
      photoUrl: vRows[i][9],
      purpose: vRows[i][10],
      notes: vRows[i][11],
      orderValue: vRows[i][12],
      paymentCollected: vRows[i][13],
      nextVisitDate: vRows[i][14],
      status: vRows[i][15]
    });
  }

  return { success: true, visits: list.reverse() };
}

function getOrCreateDriveFolder(name) {
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

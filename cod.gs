function doGet(e) {
  // ดึงพารามิเตอร์ action
  var action = e ? e.parameter.action : null;
  
  if (action === 'ping') {
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', message: 'Connected' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'get_map_list') {
    return getMapList();
  }
  if (action === 'get_map_data') {
    return getMapData(e.parameter.id);
  }
  
  // 1. ถ้ามีการระบุว่า action=get_map ให้ไปทำงานที่ระบบแผนที่
  if (action === 'get_map') {
    return getMapFiles();
  }
  if (action === 'find_smart_slot') {
    return findSmartSlot(e.parameter.date);
  }

  // 2. ถ้าไม่ได้ระบุ action หรือเป็นการโหลดหน้าแรก ให้ดึงข้อมูลคิวกลับไป
  return getReqQueue();
}

function doPost(e) {
  try {
    // อ่านข้อมูลที่ส่งมาจาก HTML (ในรูปแบบ JSON String)
    var data = JSON.parse(e.postData.contents);
    
    // ถ้าสั่งบันทึกข้อมูลคำร้องใหม่
    if (data.action === 'save_reception') {
      return saveReception(data);
    } else if (data.action === 'save_investigation') {
      return saveInvestigation(data);
    } else if (data.action === 'delete_record') {
      return deleteRecord(data);
    } else if (data.action === 'save_survey_schedule') {
      return saveSurveySchedule(data);
    } else if (data.action === 'update_survey_item') {
      return updateSurveyItem(data);
    } else if (data.action === 'login') {
      return authenticateUser(data);
    } else if (data.action === 'save_urgent_task') {
      return saveUrgentTask(data);
    }
    
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Unknown action'}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ===============================================
// ฟังก์ชัน 1: ดึงไฟล์แผนที่ (.json) จาก Google Drive
// ===============================================
function getMapFiles() {
  // *** ต้องนำ ID โฟลเดอร์ใน Drive ที่เก็บไฟล์ .json มาใส่ที่นี่ ***
  var FOLDER_ID = '1UgByuLmCTXImak4e1zJ8ngzv_bYxsf9Z'; 
  
  try {
    var folder = DriveApp.getFolderById(FOLDER_ID);
    var files = folder.getFiles();
    var maps = [];
    var loadedNames = {};
    
    // วนลูปอ่านทุกไฟล์ในโฟลเดอร์
    while (files.hasNext()) {
      var file = files.next();
      var filename = file.getName().toLowerCase();
      
      // เช็คว่าเป็นไฟล์ .json หรือ .geojson ไหม
      if (filename.endsWith('.json') || filename.endsWith('.geojson')) {
        var baseName = filename.replace('.json', '').replace('.geojson', '');
        if (loadedNames[baseName]) continue; // กันโหลดไฟล์ซ้ำถ้ามีทั้ง .json และ .geojson ในโฟลเดอร์เดียวกัน
        loadedNames[baseName] = true;
        
        try {
          var content = file.getBlob().getDataAsString('UTF-8');
          var jsonData = JSON.parse(content); // แปลงเนื้อหาเป็น JSON object
          
          maps.push({
            filename: file.getName(),
            data: jsonData
          });
        } catch(e) {
          // ข้ามไฟล์ที่โครงสร้าง JSON พัง
          Logger.log('Error parsing JSON for file: ' + filename);
        }
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({status: 'success', maps: maps}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ===============================================
// ฟังก์ชัน 1.1: ดึงรายชื่อไฟล์แผนที่ทั้งหมด (.json) 
// ===============================================
function getMapList() {
  var FOLDER_ID = '1UgByuLmCTXImak4e1zJ8ngzv_bYxsf9Z'; 
  
  try {
    var folder = DriveApp.getFolderById(FOLDER_ID);
    var files = folder.getFiles();
    var list = [];
    var loadedNames = {};
    
    // วนลูปอ่านทุกไฟล์ในโฟลเดอร์
    while (files.hasNext()) {
      var file = files.next();
      var filename = file.getName().toLowerCase();
      
      // เช็คว่าเป็นไฟล์ .json หรือ .geojson ไหม
      if (filename.endsWith('.json') || filename.endsWith('.geojson')) {
        var baseName = filename.replace('.json', '').replace('.geojson', '');
        if (loadedNames[baseName]) continue; // กันโหลดไฟล์ซ้ำ
        loadedNames[baseName] = true;
        
        list.push({
          id: file.getId(),
          filename: file.getName()
        });
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({status: 'success', list: list}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ===============================================
// ฟังก์ชัน 1.2: ดึงข้อมูลไฟล์แผนที่ 1 ไฟล์ตาม ID ที่ระบุ
// ===============================================
function getMapData(fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    var content = file.getBlob().getDataAsString('UTF-8');
    
    // ส่ง string เปล่าๆ กลับไปเลย ไม่ต้อง parse เป็น JSON ฝั่ง GAS ช่วยลดการกิน Memory มหาศาล
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success', 
      data_string: content
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: e.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ===============================================
// ฟังก์ชัน 2: ดึงคิวรับคำร้อง จาก Google Sheet
// ===============================================
function getReqQueue() {
  // *** ต้องนำ ID ของไฟล์ Google Sheet มาใส่ที่นี่ ***
  var SHEET_ID = '1kf10jeU_1ZqqNgLB6t4JGiAs4pwne6waiktI7Ydd2wQ';
  var SHEET_NAME = 'Queue'; // ชื่อ Sheet ข้างล่าง (Tab)
  
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      // ถ้าไม่มีหน้า Sheet ชื่อนี้ ให้คืนค่าว่าง
      return ContentService.createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) { // มีแต่ Headings (หรือว่าง)
       return ContentService.createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    var headers = data[0];
    var result = [];
    
    // วนลูปเฉพาะบรรทัดที่เป็นข้อมูล (ข้าม headers บรรทัดแรกไป)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var obj = {};
      
      // Map คอลัมน์เข้าเป็น key และทำเป็นตัวเล็กทั้งหมดเพื่อความสม่ำเสมอ
      for (var j = 0; j < headers.length; j++) {
        var key = headers[j] ? headers[j].toString().trim().toLowerCase() : 'col_' + j; 
        var val = row[j];
        if (val instanceof Date) {
           var tz = ss.getSpreadsheetTimeZone();
           val = Utilities.formatDate(val, tz, "yyyy-MM-dd");
        }
        if (key) {
           obj[key] = val;
        }
      }
      
      // แมปฟิลด์สำคัญให้เป็นชื่อมาตรฐาน
      var rawId = obj.id || obj.req_id;
      var rawName = obj.name || obj.req_name;
      
      if (rawId || rawName) {
          // ประกันว่ามีทั้ง key id และ req_id ให้ frontend เรียกได้
          if (!obj.id) obj.id = rawId;
          if (!obj.req_id) obj.req_id = rawId;
          if (!obj.name) obj.name = rawName;
          if (!obj.req_name) obj.req_name = rawName;
          if (!obj.address && obj.req_address) obj.address = obj.req_address;
          
          result.push(obj);
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify([])) // คืนค่า array ว่างๆ ถ้า error
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ===============================================
// ฟังก์ชัน 3: บันทึกรับคำร้องลง Google Sheet (อัปเดตข้อมูลเก่าหรือเพิ่มใหม่)
// ===============================================
function saveReception(data) {
  var SHEET_ID = '1kf10jeU_1ZqqNgLB6t4JGiAs4pwne6waiktI7Ydd2wQ'; 
  var SHEET_NAME = 'Queue'; 
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Sheet not found'})).setMimeType(ContentService.MimeType.JSON);

    var headers = sheet.getLastColumn() > 0 ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return h.toString().trim(); }) : [];
    
    // Auto-create headers
    var keys = Object.keys(data);
    var headersChanged = false;
    for (var k = 0; k < keys.length; k++) {
      var key = keys[k];
      if (key !== 'action' && key !== 'attachments' && headers.indexOf(key) === -1) {
        sheet.getRange(1, headers.length + 1).setValue(key);
        headers.push(key);
        headersChanged = true;
      }
    }

    var idColIdx = headers.indexOf('req_id') !== -1 ? headers.indexOf('req_id') : headers.indexOf('id');
    var targetId = data.req_id || data.id;

    // Auto-increment for -NEW
    if (targetId && targetId.indexOf('-NEW') !== -1 && idColIdx !== -1) {
      var prefix = targetId.split('-NEW')[0];
      var lastRow = sheet.getLastRow();
      var nextNum = 1;
      if (lastRow > 1) {
        var allIds = sheet.getRange(2, idColIdx + 1, lastRow - 1, 1).getValues();
        var maxSeq = 0;
        for (var i = 0; i < allIds.length; i++) {
          var cid = allIds[i][0].toString();
          if (cid.indexOf(prefix) === 0) {
            var seq = parseInt(cid.split('-').pop());
            if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
          }
        }
        nextNum = maxSeq + 1;
      }
      targetId = prefix + '-' + nextNum.toString().padStart(3, '0');
      data.req_id = targetId;
      data.id = targetId;
    }

    var rowData = headers.map(function(h) { return data[h] !== undefined ? data[h] : ''; });
    var isUpdated = false;
    var lastRow = sheet.getLastRow();
    if (targetId && lastRow > 1 && idColIdx !== -1) {
      var idData = sheet.getRange(2, idColIdx + 1, lastRow - 1, 1).getValues();
      for (var r = 0; r < idData.length; r++) {
        if (idData[r][0] == targetId) {
          sheet.getRange(r + 2, 1, 1, headers.length).setValues([rowData]);
          isUpdated = true;
          break;
        }
      }
    }
    if (!isUpdated) sheet.appendRow(rowData);

    lock.releaseLock();
    return ContentService.createTextOutput(JSON.stringify({status: 'success', generated_id: targetId})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    if (lock) lock.releaseLock();
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

// ===============================================
// ฟังก์ชัน 4: บันทึกข้อมูลสอบสวนสิทธิลง Google Sheet
// ===============================================
function saveInvestigation(data) {
  var SHEET_ID = '1kf10jeU_1ZqqNgLB6t4JGiAs4pwne6waiktI7Ydd2wQ'; 
  var SHEET_NAME = 'Queue'; 
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Sheet not found'})).setMimeType(ContentService.MimeType.JSON);
    
    var headers = sheet.getLastColumn() > 0 ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return h.toString().trim(); }) : [];
    
    // Auto-create headers
    var keys = Object.keys(data);
    for (var k = 0; k < keys.length; k++) {
      var key = keys[k];
      if (key !== 'action' && key !== 'case_id' && headers.indexOf(key) === -1) {
        sheet.getRange(1, headers.length + 1).setValue(key);
        headers.push(key);
      }
    }
    
    if (!data.status) data.status = 'สอบสวนสิทธิแล้ว';

    var targetId = data.case_id;
    var lastRow = sheet.getLastRow();
    if (targetId && lastRow > 1) {
      var idColIdx = headers.indexOf('req_id') !== -1 ? headers.indexOf('req_id') : headers.indexOf('id');
      if (idColIdx !== -1) {
        var idData = sheet.getRange(2, idColIdx + 1, lastRow - 1, 1).getValues();
        for (var r = 0; r < idData.length; r++) {
          if (idData[r][0] == targetId) {
             var rowRange = sheet.getRange(r + 2, 1, 1, headers.length);
             var existingRow = rowRange.getValues()[0];
             for (var i = 0; i < headers.length; i++) {
               var hName = headers[i];
               if (data[hName] !== undefined && data[hName] !== "") existingRow[i] = data[hName];
             }
             rowRange.setValues([existingRow]);
             return ContentService.createTextOutput(JSON.stringify({status: 'success'})).setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Queue ID not found: ' + targetId})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

// ===============================================
// ฟังก์ชัน 5: ยกเลิก/ลบ คำร้อง (เปลี่ยนสถานะเป็น ยกเลิก)
// ===============================================
function deleteRecord(data) {
  var SHEET_ID = '1kf10jeU_1ZqqNgLB6t4JGiAs4pwne6waiktI7Ydd2wQ'; 
  var SHEET_NAME = 'Queue'; 
  
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'ไม่พบหน้า Sheet'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var headers = [];
    if (lastCol > 0) {
      headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return h.toString().trim(); });
    }
    
    var targetId = data.id || data.case_id; 
    var isUpdated = false;
    
    if (targetId && lastRow > 1) {
      var idColIndex = headers.indexOf('req_id');
      if (idColIndex === -1) idColIndex = headers.indexOf('id');
      
      var remarkColIndex = headers.indexOf('remark');
      
      if (idColIndex !== -1) {
        var idData = sheet.getRange(2, idColIndex + 1, lastRow - 1, 1).getValues();
        
        for (var r = 0; r < idData.length; r++) {
          if (idData[r][0] == targetId) {
             // โหลดข้อมูลแถวเดิมมาก่อน
             var existingRow = sheet.getRange(r + 2, 1, 1, headers.length).getValues()[0];
             
             // เปลี่ยนสถานะ
             var statusColIndex = headers.indexOf('status');
             if (statusColIndex !== -1) {
                 existingRow[statusColIndex] = 'ยกเลิก';
             }
             
             // บันทึกหมายเหตุ
             if (remarkColIndex !== -1 && data.remark) {
                 // ถ้ามีหมายเหตุเดิม ให้เอามาต่อกัน หรือเขียนทับ
                 var oldRemark = existingRow[remarkColIndex] || '';
                 existingRow[remarkColIndex] = oldRemark ? oldRemark + ' | [ยกเลิก] ' + data.remark : '[ยกเลิก] ' + data.remark;
             } else if (statusColIndex !== -1 && data.remark) {
                 // ถ้าไม่มีคอลัมน์ remark ให้พ่วงท้าย status ไปเลย
                 existingRow[statusColIndex] = 'ยกเลิก (' + data.remark + ')';
             }
             
             // เขียนทับ
             sheet.getRange(r + 2, 1, 1, headers.length).setValues([existingRow]);
             isUpdated = true;
             break;
          }
        }
      }
    }
    
    if (!isUpdated) {
        return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'ไม่พบเลขคิว ' + targetId}))
          .setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({status: 'success'}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Helper สำหรับแมปหัวตารางแบบไม่สนใจตัวพิมพ์เล็ก/ใหญ่
function getHeaderMap(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i].toString().trim().toLowerCase();
    if (h) map[h] = i;
  }
  return map;
}

// ===============================================
// ฟังก์ชัน 6: บันทึกแผนตารางออกรังวัด (บันทึกเป็นกลุ่ม)
// ===============================================
function saveSurveySchedule(data) {
  var SHEET_ID = '1kf10jeU_1ZqqNgLB6t4JGiAs4pwne6waiktI7Ydd2wQ'; 
  var SHEET_NAME = 'Queue'; 
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Sheet not found'})).setMimeType(ContentService.MimeType.JSON);
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return ContentService.createTextOutput(JSON.stringify({status: 'success', count: 0})).setMimeType(ContentService.MimeType.JSON);

    var hMap = getHeaderMap(sheet);
    var dateIdx = hMap['survey_date'];
    var timeIdx = hMap['survey_time'];
    var idIdx = hMap['req_id'] !== undefined ? hMap['req_id'] : hMap['id'];
    var stIdx = hMap['status'];
    
    // Create missing columns if needed
    if (dateIdx === undefined) { sheet.getRange(1, sheet.getLastColumn() + 1).setValue('survey_date'); dateIdx = sheet.getLastColumn() - 1; }
    if (timeIdx === undefined) { sheet.getRange(1, sheet.getLastColumn() + 1).setValue('survey_time'); timeIdx = sheet.getLastColumn() - 1; }

    var idData = sheet.getRange(2, idIdx + 1, lastRow - 1, 1).getValues();
    var rangeToUpdate = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    var updatedCount = 0;

    for (var i = 0; i < data.schedule.length; i++) {
      var item = data.schedule[i];
      var targetDate = item.date || data.survey_date;
      for (var r = 0; r < idData.length; r++) {
        if (idData[r][0].toString().trim() == item.id.toString().trim()) {
          if (dateIdx !== undefined) rangeToUpdate[r][dateIdx] = (targetDate && targetDate !== '-') ? targetDate : '-';
          if (timeIdx !== undefined) rangeToUpdate[r][timeIdx] = item.timeSlot;
          if (stIdx !== undefined) {
            var curS = rangeToUpdate[r][stIdx];
            if (curS !== 'รังวัดแล้ว' && curS !== 'ยกเลิก') rangeToUpdate[r][stIdx] = 'รอรังวัด';
          }
          updatedCount++;
          break;
        }
      }
    }

    if (updatedCount > 0) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).setValues(rangeToUpdate);
    return ContentService.createTextOutput(JSON.stringify({status: 'success', count: updatedCount})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

// ===============================================
// ฟังก์ชัน 7: อัปเดตสถานะรายคิว หรือ เบอร์โทร (อัปเดตเดี่ยว)
// ===============================================
function updateSurveyItem(data) {
  var SHEET_ID = '1kf10jeU_1ZqqNgLB6t4JGiAs4pwne6waiktI7Ydd2wQ'; 
  var SHEET_NAME = 'Queue'; 
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({status: 'error'})).setMimeType(ContentService.MimeType.JSON);
    
    var hMap = getHeaderMap(sheet);
    var targetId = data.case_id ? data.case_id.toString().trim() : ""; 
    
    // ค้นหาสารพัด Index ของ ID ที่อาจจะมี
    var idIndices = [];
    if (hMap['id'] !== undefined) idIndices.push(hMap['id']);
    if (hMap['req_id'] !== undefined) idIndices.push(hMap['req_id']);
    
    if (targetId && idIndices.length > 0) {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var fullData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
        
        for (var r = 0; r < fullData.length; r++) {
          var row = fullData[r];
          var matchFound = false;
          
          // ตรวจสอบ ID ทุกคอลัมน์ที่เป็นไปได้
          for (var k = 0; k < idIndices.length; k++) {
            var val = row[idIndices[k]];
            if (val && val.toString().trim() == targetId) {
              matchFound = true;
              break;
            }
          }
          
          if (matchFound) {
            var rowRange = sheet.getRange(r + 2, 1, 1, sheet.getLastColumn());
            var existingRow = row; // use the row we already fetched
            
            var sIdx = hMap['status'];
            var rIdx = hMap['remark'];
            var dIdx = hMap['survey_date'];
            var tIdx = hMap['survey_time'];

            if (data.update_type === 'status') {
              if (sIdx !== undefined) existingRow[sIdx] = data.new_status;
            } else if (data.update_type === 'status_and_remark') {
              if (sIdx !== undefined) existingRow[sIdx] = data.new_status;
              // Clear date/time explicitly
              if (dIdx !== undefined) existingRow[dIdx] = '-';
              if (tIdx !== undefined) existingRow[tIdx] = '-';
              
              if (rIdx === undefined) {
                sheet.getRange(1, sheet.getLastColumn() + 1).setValue('remark');
                existingRow.push(data.remark);
                rowRange = sheet.getRange(r + 2, 1, 1, sheet.getLastColumn());
              } else {
                existingRow[rIdx] = data.remark;
              }
            } else if (data.update_type === 'call_log') {
              // ... same call_log logic ...
              var cIdx = hMap['last_called'];
              if (cIdx === undefined) {
                sheet.getRange(1, sheet.getLastColumn() + 1).setValue('last_called');
                existingRow.push(data.call_time);
                rowRange = sheet.getRange(r + 2, 1, 1, sheet.getLastColumn());
              } else {
                existingRow[cIdx] = data.call_time;
              }
            } else if (data.update_type === 'coords') {
              // ... same coords logic ...
              var coIdx = hMap['coords'];
              if (coIdx === undefined) {
                sheet.getRange(1, sheet.getLastColumn() + 1).setValue('coords');
                existingRow.push(data.new_coords);
                rowRange = sheet.getRange(r + 2, 1, 1, sheet.getLastColumn());
              } else {
                existingRow[coIdx] = data.new_coords;
              }
            } else if (data.update_type === 'reschedule') {
              if (dIdx !== undefined) existingRow[dIdx] = (data.new_date && data.new_date !== '-') ? data.new_date : '-';
              if (tIdx !== undefined) existingRow[tIdx] = data.new_time;
              if (sIdx !== undefined) {
                existingRow[sIdx] = data.new_status || (data.new_date !== '-' ? 'รอรังวัด' : 'พร้อมรังวัด');
              }
              if (data.remark) {
                 if (rIdx === undefined) {
                   sheet.getRange(1, sheet.getLastColumn() + 1).setValue('remark');
                   existingRow.push(data.remark);
                   rowRange = sheet.getRange(r + 2, 1, 1, sheet.getLastColumn());
                 } else {
                   existingRow[rIdx] = data.remark;
                 }
              }
            }
            rowRange.setValues([existingRow]);
            SpreadsheetApp.flush();
            return ContentService.createTextOutput(JSON.stringify({status: 'success'})).setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'ID not found'})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}
// ===============================================
// ฟังก์ชัน 8: ตรวจสอบการเข้าสู่ระบบ (Login)
// ===============================================
function authenticateUser(data) {
  var SHEET_ID = '1kf10jeU_1ZqqNgLB6t4JGiAs4pwne6waiktI7Ydd2wQ'; 
  var SHEET_NAME = 'Users'; // สร้าง Tab ใหม่ชื่อ Users
  
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'ไม่พบหน้า Sheet ที่ชื่อ Users (สำหรับเก็บรหัสผ่าน)'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง'})).setMimeType(ContentService.MimeType.JSON);
    
    // ดึงข้อมูล 4 คอลัมน์: username, password, role, name
    var sheetData = sheet.getRange(2, 1, lastRow - 1, 4).getValues(); 
    
    for (var i = 0; i < sheetData.length; i++) {
      if (sheetData[i][0].toString() === data.username && sheetData[i][1].toString() === data.password) {
        return ContentService.createTextOutput(JSON.stringify({
          status: 'success', 
          role: sheetData[i][2].toString(),
          name: sheetData[i][3].toString()
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง'}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ===============================================
// ฟังก์ชัน 9: ค้นหาช่วงเวลาว่าง (Smart Slot Finder)
// ===============================================
function findSmartSlot(dateString) {
  var SHEET_ID = '1kf10jeU_1ZqqNgLB6t4JGiAs4pwne6waiktI7Ydd2wQ'; 
  var SHEET_NAME = 'Queue'; 
  
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({status: 'error'})).setMimeType(ContentService.MimeType.JSON);
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var dateIdx = headers.indexOf('survey_date');
    var timeIdx = headers.indexOf('survey_time');
    
    if (dateIdx === -1 || timeIdx === -1) {
      return ContentService.createTextOutput(JSON.stringify({ status: "found", slot: "08:30" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    var existingSlots = [];
    for (var i = 1; i < data.length; i++) {
       var rowDate = data[i][dateIdx];
       if (rowDate instanceof Date) {
         var tz = ss.getSpreadsheetTimeZone();
         rowDate = Utilities.formatDate(rowDate, tz, "yyyy-MM-dd");
       }
       if (rowDate == dateString) {
         existingSlots.push(data[i][timeIdx].toString().trim());
       }
    }
    
    var standardSlots = ["08:30", "10:00", "13:00", "14:30"];
    for (var j = 0; j < standardSlots.length; j++) {
      if (existingSlots.indexOf(standardSlots[j]) === -1) {
        return ContentService.createTextOutput(JSON.stringify({ status: "found", slot: standardSlots[j] }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: "full", message: "ขออภัย ทุกช่วงเวลาในวันนี้เต็มแล้ว" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ===============================================
// ฟังก์ชัน 10: บันทึกงานด่วนจากหน้างาน (On-site Urgent Task)
// ===============================================
function saveUrgentTask(data) {
  var SHEET_ID = '1kf10jeU_1ZqqNgLB6t4JGiAs4pwne6waiktI7Ydd2wQ'; 
  var SHEET_NAME = 'Queue'; 
  
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({status: 'error'})).setMimeType(ContentService.MimeType.JSON);
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return h.toString().trim(); });
    
    var newId = "URG-D-" + Utilities.getUuid().substr(0, 4).toUpperCase();
    
    var row = headers.map(function(h) {
      if (h === 'id' || h === 'req_id') return newId;
      if (h === 'name' || h === 'req_name') return data.name;
      if (h === 'phone' || h === 'req_phone') return data.phone;
      if (h === 'coords') return data.coords;
      if (h === 'status') return data.status || "งานด่วนหน้างาน (รอเจ้าหน้าที่มารับคำร้อง)";
      if (h === 'remark') return "[เพิ่มงานด่วนจากหน้างาน]";
      if (h === 'req_date' || h === 'timestamp') return new Date();
      return "";
    });
    
    sheet.appendRow(row);
    
    return ContentService.createTextOutput(JSON.stringify({status: 'success', id: newId}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

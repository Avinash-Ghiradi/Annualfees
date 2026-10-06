// ==========================================
// STUDENT QR CODE & DRESS MANAGEMENT SYSTEM
// Google Apps Script Backend Code (Code.gs)
// Connected Google Sheet ID: 1uhK7jOTvlKGTC-7FVQxJa3g10I9nnNce2icXcE4ccVQ
// ==========================================

const SPREADSHEET_ID = "1uhK7jOTvlKGTC-7FVQxJa3g10I9nnNce2icXcE4ccVQ";
const DATASHEET = "Data";
const STUDENTSHEET = "Students";
const CURRENCY_SYMBOL = "₹"; // Indian Rupee
const QR_CODE_FOLDER_NAME = "QR Codes";
const DROPDOWN_SHEET = "DROPDOWN";
const RECEIPT_FOLDER_NAME = "Receipts";
const WHATSAPP_API_URL = "https://api.whatsapp.com/send";

/**
 * Gets the active or linked Spreadsheet reliably.
 */
function getSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== "") {
    try {
      return SpreadsheetApp.openById(SPREADSHEET_ID.trim());
    } catch (e) {
      console.error("Error opening spreadsheet by ID: " + e.toString());
    }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Gets a sheet by name with case-insensitive and fallback matching.
 */
function getSheetByNameCaseInsensitive(ss, sheetName) {
  if (!ss) return null;
  const sheet = ss.getSheetByName(sheetName);
  if (sheet) return sheet;

  const sheets = ss.getSheets();
  const target = sheetName.trim().toLowerCase();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().trim().toLowerCase() === target) {
      return sheets[i];
    }
  }
  return null;
}

/**
 * Maps header row strings to their column index.
 */
function getHeaderMap(headerRow) {
  const map = {};
  if (!headerRow || !Array.isArray(headerRow)) return map;
  headerRow.forEach((h, idx) => {
    if (h !== undefined && h !== null) {
      const key = h.toString().trim().toLowerCase();
      map[key] = idx;
    }
  });
  return map;
}

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  return template
    .evaluate()
    .setTitle('Student QR Code System')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .addMetaTag('apple-mobile-web-app-capable', 'yes')
    .addMetaTag('mobile-web-app-capable', 'yes')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function processDressSale(formObject) {
  try {
    const ss = getSpreadsheet();
    let dataSheet = getSheetByNameCaseInsensitive(ss, DATASHEET);
    
    // Ensure Data sheet exists
    if (!dataSheet) {
      dataSheet = ss.insertSheet(DATASHEET);
      dataSheet.appendRow([
        'Timestamp', 'Student Name', 'Class', 'Song/Event', 'Quantity', 'Price', 
        'Student Code', 'Phone No', 'Dress Code', 'Total Amount', 'Admission Fees', 
        'Annual Day Fees', 'Sports Fees', 'Exam Fees', 'Library Fees', 'Medical Fees', 
        'Tour Fees', 'Lab Fees', 'Penalty Fees', 'Other Fees', 'Total Fees', 'Grand Total'
      ]);
    }
    
    const quantity = parseInt(formObject.quantity) || 0;
    const price = parseFloat(formObject.price) || 0;
    const dressTotal = quantity * price;
    
    const student = lookupStudent(formObject.studentCode);
    if (!student) {
      return { 
        success: false, 
        error: `Student with code "${formObject.studentCode}" not found. Please register the student first.` 
      };
    }
    
    if (quantity > student.remainingDress) {
      return { 
        success: false, 
        error: `Insufficient stock! Available: ${student.remainingDress}, Requested: ${quantity}` 
      };
    }
    
    let selectedFees = {};
    let totalFees = 0;
    
    if (formObject.feeCategories) {
      try {
        const feeCategories = JSON.parse(formObject.feeCategories);
        const classFeesResult = getClassFees(formObject.class_name || student.className);
        
        if (classFeesResult.success) {
          const fees = classFeesResult.fees;
          feeCategories.forEach(feeName => {
            if (fees[feeName] && fees[feeName] > 0) {
              selectedFees[feeName] = fees[feeName];
              totalFees += fees[feeName];
            }
          });
        }
      } catch (e) {
        console.log('Error parsing fee categories:', e);
      }
    }
    
    const grandTotal = dressTotal + totalFees;
    
    if (quantity <= 0) {
      return { 
        success: false, 
        error: 'Quantity must be greater than 0' 
      };
    }
    
    const timestamp = new Date().toLocaleString('en-IN', { 
      timeZone: 'Asia/Kolkata',
      hour12: false 
    });
    
    const rowData = [
      timestamp,
      formObject.student_name || student.studentName,
      formObject.class_name || student.className,
      formObject.song_event || student.songEvent,
      quantity,
      price,
      formObject.studentCode,
      formObject.phone_no || student.phoneNo || '',
      formObject.dress_code || student.dressCode || '',
      dressTotal
    ];
    
    const feeHeaders = [
      'Admission Fees', 'Annual Day Fees', 'Sports Fees', 'Exam Fees', 
      'Library Fees', 'Medical Fees', 'Tour Fees', 'Lab Fees', 
      'Penalty Fees', 'Other Fees'
    ];
    
    feeHeaders.forEach(feeHeader => {
      rowData.push(selectedFees[feeHeader] || 0);
    });
    
    rowData.push(totalFees);
    rowData.push(grandTotal);
    
    dataSheet.appendRow(rowData);
    
    const updateSuccess = updateStudentStock(formObject.studentCode, quantity);
    
    if (updateSuccess) {
      const updatedStudent = lookupStudent(formObject.studentCode);
      
      return { 
        success: true, 
        message: 'Dress and fees processed successfully!',
        studentData: updatedStudent,
        saleData: {
          timestamp: timestamp,
          studentName: formObject.student_name || student.studentName,
          className: formObject.class_name || student.className,
          songEvent: formObject.song_event || student.songEvent,
          phoneNo: formObject.phone_no || student.phoneNo || '',
          dressCode: formObject.dress_code || student.dressCode || '',
          price: price,
          quantity: quantity,
          dressTotal: dressTotal,
          fees: selectedFees,
          totalFees: totalFees,
          grandTotal: grandTotal,
          studentId: formObject.studentCode
        }
      };
    } else {
      return { success: false, error: 'Failed to update student stock' };
    }
  } catch (error) {
    Logger.log('Error in processDressSale: ' + error.toString());
    return { 
      success: false, 
      error: 'Error processing sale: ' + error.message 
    };
  }
}

function getClasses() {
  try {
    const ss = getSpreadsheet();
    let sheet = getSheetByNameCaseInsensitive(ss, DROPDOWN_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(DROPDOWN_SHEET);
      sheet.getRange('A1').setValue('Class');
      sheet.getRange('B1').setValue('Song/Event Name');
      sheet.getRange('C1').setValue('Dress Code');
      return [];
    }
    
    var data = sheet.getRange('A:A').getValues();
    var classes = data.filter(String).map(function(row) {
      return row[0];
    });
    
    if (classes.length > 0 && classes[0].toString().toLowerCase() === 'class') {
      classes.shift();
    }
    
    return [...new Set(classes)].filter(Boolean);
  } catch (error) {
    console.error('Error getting classes:', error);
    return [];
  }
}

function getSongEvents() {
  try {
    const ss = getSpreadsheet();
    var sheet = getSheetByNameCaseInsensitive(ss, DROPDOWN_SHEET);
    if (!sheet) {
      return [];
    }
    
    var data = sheet.getRange('B:B').getValues();
    var songEvents = data.filter(String).map(function(row) {
      return row[0];
    });
    
    if (songEvents.length > 0 && songEvents[0].toString().toLowerCase().includes('song')) {
      songEvents.shift();
    }
    
    return [...new Set(songEvents)].filter(Boolean);
  } catch (error) {
    console.error('Error getting song events:', error);
    return [];
  }
}

function getDressCodes() {
  try {
    const ss = getSpreadsheet();
    var sheet = getSheetByNameCaseInsensitive(ss, DROPDOWN_SHEET);
    if (!sheet) {
      return [];
    }
    
    var data = sheet.getRange('C:C').getValues();
    var dressCodes = data.filter(String).map(function(row) {
      return row[0];
    });
    
    if (dressCodes.length > 0 && dressCodes[0].toString().toLowerCase().includes('dress')) {
      dressCodes.shift();
    }
    
    return [...new Set(dressCodes)].filter(Boolean);
  } catch (error) {
    console.error('Error getting dress codes:', error);
    return [];
  }
}

function saveStudent(studentData) {
  try {
    const ss = getSpreadsheet();
    let studentSheet = getSheetByNameCaseInsensitive(ss, STUDENTSHEET);
    
    if (!studentSheet) {
      studentSheet = ss.insertSheet(STUDENTSHEET);
      studentSheet.appendRow(['Student ID', 'Class', 'Student Name', 'Song/Event Name', 'Price', 'Total Dress', 'Remaining Dress', 'Phone No', 'Dress Code', 'QR URL', 'Registration Date']);
    } else {
      const headers = studentSheet.getRange(1, 1, 1, Math.max(11, studentSheet.getLastColumn())).getValues()[0];
      const requiredColumns = [
        {name: 'Student ID', position: 0},
        {name: 'Class', position: 1},
        {name: 'Student Name', position: 2},
        {name: 'Song/Event Name', position: 3},
        {name: 'Price', position: 4},
        {name: 'Total Dress', position: 5},
        {name: 'Remaining Dress', position: 6},
        {name: 'Phone No', position: 7},
        {name: 'Dress Code', position: 8},
        {name: 'QR URL', position: 9},
        {name: 'Registration Date', position: 10}
      ];
      
      requiredColumns.forEach(function(column, index) {
        if (!headers[index] || headers[index].toString().trim() === '') {
          studentSheet.getRange(1, index + 1).setValue(column.name);
        }
      });
    }
    
    const data = studentSheet.getDataRange().getValues();
    const map = getHeaderMap(data[0]);
    const idIdx = map['student id'] !== undefined ? map['student id'] : 0;
    const nameIdx = map['student name'] !== undefined ? map['student name'] : 2;

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] && data[i][idIdx].toString().toLowerCase() === studentData.studentId.toLowerCase()) {
        return {
          success: false,
          error: 'Student ID "' + studentData.studentId + '" already exists!'
        };
      }
      
      if (data[i][nameIdx] && data[i][nameIdx].toString().toLowerCase() === studentData.studentName.toLowerCase()) {
        return {
          success: false,
          error: 'Student "' + studentData.studentName + '" already exists!'
        };
      }
    }
    
    const studentId = studentData.studentId || generateStudentId();
    const initialDress = parseInt(studentData.initialDress) || 0;
    
    let qrUrl = '';
    try {
      qrUrl = generateQRCode(studentId, studentData.studentName);
    } catch (error) {
      console.error('QR generation error: ' + error.toString());
    }
    
    const registrationDate = new Date().toLocaleString('en-IN', { 
      timeZone: 'Asia/Kolkata',
      hour12: false 
    });
    
    studentSheet.appendRow([
      studentId,
      studentData.className,
      studentData.studentName,
      studentData.songEvent,
      parseFloat(studentData.price) || 0,
      initialDress,
      initialDress,
      studentData.phoneNo || '',
      studentData.dressCode || '',
      qrUrl || '',
      registrationDate
    ]);
    
    return {
      success: true,
      studentId: studentId,
      qrUrl: qrUrl || '',
      note: qrUrl ? 'QR code generated successfully' : 'QR code generation failed - will retry later'
    };
  } catch (error) {
    console.error('Error in saveStudent: ' + error.toString());
    return {
      success: false,
      error: 'Failed to save student: ' + error.message
    };
  }
}

function generateStudentId() {
  const timestamp = new Date().getTime();
  const randomNum = Math.floor(Math.random() * 1000);
  return 'STU' + timestamp.toString().slice(-8) + randomNum.toString().padStart(3, '0');
}

function generateQRCode(studentId, studentName) {
  try {
    const qrContent = studentId;
    
    const qrUrls = [
      `https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${encodeURIComponent(qrContent)}&format=png`,
      `https://chart.googleapis.com/chart?chs=300x300&cht=qr&chl=${encodeURIComponent(qrContent)}&choe=UTF-8&chld=H|2`,
      `https://quickchart.io/qr?text=${encodeURIComponent(qrContent)}&size=300&margin=1`
    ];
    
    for (let i = 0; i < qrUrls.length; i++) {
      try {
        const response = UrlFetchApp.fetch(qrUrls[i], {
          muteHttpExceptions: true,
          followRedirects: true,
          validateHttpsCertificates: false
        });
        
        if (response.getResponseCode() === 200) {
          try {
            const qrBlob = response.getBlob();
            const folder = getOrCreateQRFolder();
            const fileName = `QR_${studentId}_${studentName.replace(/[^a-zA-Z0-9]/g, '_')}.png`;
            
            const files = folder.getFilesByName(fileName);
            if (files.hasNext()) {
              files.next().setContent(qrBlob);
            } else {
              const file = folder.createFile(qrBlob);
              file.setName(fileName);
            }
            return qrUrls[i];
          } catch (driveError) {
            console.log('Could not save to Drive, using URL only:', driveError);
            return qrUrls[i];
          }
        }
      } catch (error) {
        console.log(`QR API ${i + 1} failed:`, error);
        continue;
      }
    }
    
    console.log('All QR APIs failed');
    return '';
    
  } catch (error) {
    console.error('QR Code Generation Error:', error);
    return '';
  }
}

function getOrCreateQRFolder() {
  try {
    const folders = DriveApp.getFoldersByName(QR_CODE_FOLDER_NAME);
    if (folders.hasNext()) {
      return folders.next();
    }
    
    return DriveApp.createFolder(QR_CODE_FOLDER_NAME);
  } catch (error) {
    console.error('Error accessing Drive folder:', error);
    throw error;
  }
}

function lookupStudent(studentId) {
  try {
    const ss = getSpreadsheet();
    const studentSheet = getSheetByNameCaseInsensitive(ss, STUDENTSHEET);
    
    if (!studentSheet) {
      return null;
    }
    
    const data = studentSheet.getDataRange().getValues();
    if (data.length < 2) return null;
    
    const map = getHeaderMap(data[0]);
    const idIdx = map['student id'] ?? map['student code'] ?? map['id'] ?? 0;
    const classIdx = map['class'] ?? 1;
    const nameIdx = map['student name'] ?? map['name'] ?? 2;
    const eventIdx = map['song/event name'] ?? map['song/event'] ?? map['event'] ?? 3;
    const priceIdx = map['price'] ?? 4;
    const totalDressIdx = map['total dress'] ?? map['total stock'] ?? 5;
    const remDressIdx = map['remaining dress'] ?? map['remaining stock'] ?? map['remaining'] ?? 6;
    const phoneIdx = map['phone no'] ?? map['phone'] ?? map['mobile'] ?? 7;
    const dressCodeIdx = map['dress code'] ?? 8;
    const qrUrlIdx = map['qr url'] ?? map['qr code'] ?? 9;

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] && data[i][idIdx].toString().trim().toLowerCase() === studentId.toString().trim().toLowerCase()) {
        return {
          studentId: data[i][idIdx],
          className: data[i][classIdx],
          studentName: data[i][nameIdx],
          songEvent: data[i][eventIdx],
          price: data[i][priceIdx],
          totalDress: data[i][totalDressIdx] || 0,
          remainingDress: data[i][remDressIdx] !== undefined ? data[i][remDressIdx] : (data[i][totalDressIdx] || 0),
          phoneNo: data[i][phoneIdx] || '',
          dressCode: data[i][dressCodeIdx] || '',
          qrUrl: data[i][qrUrlIdx] || ''
        };
      }
    }
    return null;
  } catch (error) {
    console.error('Error in lookupStudent: ' + error.toString());
    return null;
  }
}

function updateStudentStock(studentId, quantity) {
  try {
    const ss = getSpreadsheet();
    const studentSheet = getSheetByNameCaseInsensitive(ss, STUDENTSHEET);
    
    if (!studentSheet) {
      throw new Error('Students sheet not found');
    }
    
    const data = studentSheet.getDataRange().getValues();
    const map = getHeaderMap(data[0]);
    const idIdx = map['student id'] ?? 0;
    const remDressIdx = map['remaining dress'] ?? 6;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] && data[i][idIdx].toString().trim().toLowerCase() === studentId.toString().trim().toLowerCase()) {
        const remainingDress = parseInt(data[i][remDressIdx]) || 0;
        const newRemainingDress = remainingDress - parseInt(quantity);
        
        if (newRemainingDress < 0) {
          throw new Error(`Insufficient Dress! Available: ${remainingDress}`);
        }
        
        studentSheet.getRange(i + 1, remDressIdx + 1).setValue(newRemainingDress);
        return true;
      }
    }
    
    throw new Error('Student not found');
  } catch (error) {
    console.error('Error in updateStudentStock: ' + error.toString());
    throw error;
  }
}

function getAllStudents() {
  try {
    const ss = getSpreadsheet();
    const studentSheet = getSheetByNameCaseInsensitive(ss, STUDENTSHEET);
    
    if (!studentSheet) {
      return [];
    }
    
    const data = studentSheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const map = getHeaderMap(data[0]);
    const idIdx = map['student id'] ?? map['student code'] ?? map['id'] ?? 0;
    const classIdx = map['class'] ?? 1;
    const nameIdx = map['student name'] ?? map['name'] ?? 2;
    const eventIdx = map['song/event name'] ?? map['song/event'] ?? map['event'] ?? 3;
    const priceIdx = map['price'] ?? 4;
    const totalDressIdx = map['total dress'] ?? map['total stock'] ?? 5;
    const remDressIdx = map['remaining dress'] ?? map['remaining stock'] ?? map['remaining'] ?? 6;
    const phoneIdx = map['phone no'] ?? map['phone'] ?? map['mobile'] ?? 7;
    const dressCodeIdx = map['dress code'] ?? 8;
    const qrUrlIdx = map['qr url'] ?? map['qr code'] ?? 9;

    const students = [];
    
    for (let i = 1; i < data.length; i++) {
      if (!data[i][idIdx] && !data[i][nameIdx]) continue;
      students.push({
        id: data[i][idIdx] || '',
        class: data[i][classIdx] || '',
        name: data[i][nameIdx] || '',
        songEvent: data[i][eventIdx] || '',
        price: data[i][priceIdx] || 0,
        totalDress: data[i][totalDressIdx] || 0,
        remainingDress: data[i][remDressIdx] !== undefined ? data[i][remDressIdx] : (data[i][totalDressIdx] || 0),
        phoneNo: data[i][phoneIdx] || '',
        dressCode: data[i][dressCodeIdx] || '',
        qrUrl: data[i][qrUrlIdx] || ''
      });
    }
    
    return students;
  } catch (error) {
    console.error('Error in getAllStudents: ' + error.toString());
    return [];
  }
}

function updateStudent(studentData) {
  try {
    const ss = getSpreadsheet();
    const studentSheet = getSheetByNameCaseInsensitive(ss, STUDENTSHEET);
    
    if (!studentSheet) {
      return { success: false, error: 'Students sheet not found' };
    }
    
    const data = studentSheet.getDataRange().getValues();
    const map = getHeaderMap(data[0]);
    const idIdx = map['student id'] ?? 0;
    const classIdx = map['class'] ?? 1;
    const nameIdx = map['student name'] ?? 2;
    const eventIdx = map['song/event name'] ?? 3;
    const priceIdx = map['price'] ?? 4;
    const totalDressIdx = map['total dress'] ?? 5;
    const remDressIdx = map['remaining dress'] ?? 6;
    const phoneIdx = map['phone no'] ?? 7;
    const dressCodeIdx = map['dress code'] ?? 8;
    const qrUrlIdx = map['qr url'] ?? 9;

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] && data[i][idIdx].toString().trim().toLowerCase() === studentData.id.toString().trim().toLowerCase()) {
        studentSheet.getRange(i + 1, classIdx + 1).setValue(studentData.class);
        studentSheet.getRange(i + 1, nameIdx + 1).setValue(studentData.name);
        studentSheet.getRange(i + 1, eventIdx + 1).setValue(studentData.songEvent);
        studentSheet.getRange(i + 1, priceIdx + 1).setValue(studentData.price);
        studentSheet.getRange(i + 1, totalDressIdx + 1).setValue(studentData.totalDress);
        studentSheet.getRange(i + 1, remDressIdx + 1).setValue(studentData.remainingDress);
        studentSheet.getRange(i + 1, phoneIdx + 1).setValue(studentData.phoneNo || '');
        studentSheet.getRange(i + 1, dressCodeIdx + 1).setValue(studentData.dressCode || '');
        
        if (studentData.name !== data[i][nameIdx]) {
          try {
            const qrUrl = generateQRCode(studentData.id, studentData.name);
            if (qrUrl) {
              studentSheet.getRange(i + 1, qrUrlIdx + 1).setValue(qrUrl);
            }
          } catch (error) {
            console.error('Error updating QR code:', error);
          }
        }
        
        return { success: true };
      }
    }
    
    return { success: false, error: 'Student not found' };
  } catch (error) {
    console.error('Error in updateStudent: ' + error.toString());
    return { success: false, error: 'Error updating student: ' + error.message };
  }
}

function deleteStudent(studentId) {
  try {
    const ss = getSpreadsheet();
    const studentSheet = getSheetByNameCaseInsensitive(ss, STUDENTSHEET);
    
    if (!studentSheet) {
      return { success: false, error: 'Students sheet not found' };
    }
    
    const data = studentSheet.getDataRange().getValues();
    const map = getHeaderMap(data[0]);
    const idIdx = map['student id'] ?? 0;

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] && data[i][idIdx].toString().trim().toLowerCase() === studentId.toString().trim().toLowerCase()) {
        studentSheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    
    return { success: false, error: 'Student not found' };
  } catch (error) {
    console.error('Error in deleteStudent: ' + error.toString());
    return { success: false, error: 'Error deleting student: ' + error.message };
  }
}

function getDashboardData() {
  try {
    const ss = getSpreadsheet();
    const studentSheet = getSheetByNameCaseInsensitive(ss, STUDENTSHEET);
    const dataSheet = getSheetByNameCaseInsensitive(ss, DATASHEET);

    if (!studentSheet) {
      return { success: false, error: 'Student sheet not found in spreadsheet' };
    }

    const studentData = studentSheet.getDataRange().getValues();
    const salesData = dataSheet ? dataSheet.getDataRange().getValues() : [];

    const totalStudents = Math.max(0, studentData.length - 1);

    let totalRevenue = 0;
    if (salesData.length > 1) {
      const salesMap = getHeaderMap(salesData[0]);
      const priceIdx = salesMap['price'] ?? 5;
      const qtyIdx = salesMap['quantity'] ?? 4;
      
      for (let i = 1; i < salesData.length; i++) {
        const price = Number(salesData[i][priceIdx]) || 0;
        const quantity = Number(salesData[i][qtyIdx]) || 0;
        totalRevenue += price * quantity;
      }
    }

    let totalDressValue = 0;
    let totalDress = 0;
    let totalRemainingDress = 0;
    let lowDressStudents = 0;

    if (studentData.length > 1) {
      const stuMap = getHeaderMap(studentData[0]);
      const priceIdx = stuMap['price'] ?? 4;
      const totalIdx = stuMap['total dress'] ?? 5;
      const remIdx = stuMap['remaining dress'] ?? 6;

      for (let i = 1; i < studentData.length; i++) {
        const price = Number(studentData[i][priceIdx]) || 0;
        const Dress = Number(studentData[i][totalIdx]) || 0;
        const remaining = Number(studentData[i][remIdx]) || 0;
        totalDressValue += price * remaining;
        totalDress += Dress;
        totalRemainingDress += remaining;

        if (remaining < Dress * 0.2 || remaining === 0) {
          lowDressStudents++;
        }
      }
    }

    let classSales = {};
    let classRevenue = {};
    let studentSales = {};
    let dailySales = {};

    if (salesData.length > 1) {
      const salesMap = getHeaderMap(salesData[0]);
      const dateIdx = salesMap['timestamp'] ?? 0;
      const nameIdx = salesMap['student name'] ?? 1;
      const classIdx = salesMap['class'] ?? 2;
      const qtyIdx = salesMap['quantity'] ?? 4;
      const priceIdx = salesMap['price'] ?? 5;

      for (let i = 1; i < salesData.length; i++) {
        const className = salesData[i][classIdx] || 'Uncategorized';
        const studentName = salesData[i][nameIdx] || 'Unknown';
        const quantity = Number(salesData[i][qtyIdx]) || 0;
        const price = Number(salesData[i][priceIdx]) || 0;
        const dateStr = salesData[i][dateIdx];

        classSales[className] = (classSales[className] || 0) + quantity;
        classRevenue[className] = (classRevenue[className] || 0) + (price * quantity);

        if (!studentSales[studentName]) {
          studentSales[studentName] = { quantity: 0, revenue: 0 };
        }
        studentSales[studentName].quantity += quantity;
        studentSales[studentName].revenue += price * quantity;

        if (dateStr) {
          const date = new Date(dateStr);
          const dateKey = !isNaN(date.getTime()) ? date.toLocaleDateString() : dateStr.toString().split(' ')[0];
          if (!dailySales[dateKey]) {
            dailySales[dateKey] = { quantity: 0, revenue: 0 };
          }
          dailySales[dateKey].quantity += quantity;
          dailySales[dateKey].revenue += price * quantity;
        }
      }
    }

    let topStudents = Object.keys(studentSales).map(function(name) {
      return {
        name: name,
        quantity: studentSales[name].quantity,
        revenue: studentSales[name].revenue
      };
    }).sort(function(a, b) {
      return b.revenue - a.revenue;
    }).slice(0, 5);

    return {
      success: true,
      totalStudents: totalStudents,
      totalRevenue: totalRevenue,
      totalSales: Math.max(0, salesData.length - 1),
      totalDressValue: totalDressValue,
      totalDress: totalDress,
      totalRemainingDress: totalRemainingDress,
      lowDressStudents: lowDressStudents,
      classSales: classSales,
      classRevenue: classRevenue,
      topStudents: topStudents,
      dailySales: dailySales,
      DressPercentage: totalDress > 0 ? Math.round((totalRemainingDress / totalDress) * 100) : 0,
      currencySymbol: CURRENCY_SYMBOL
    };
  } catch (error) {
    console.error('Error in getDashboardData: ' + error.toString());
    return { success: false, error: 'Error loading dashboard data: ' + error.message };
  }
}

function getCurrencySymbol() {
  return CURRENCY_SYMBOL;
}

function recalculateDress() {
  try {
    const ss = getSpreadsheet();
    const studentSheet = getSheetByNameCaseInsensitive(ss, STUDENTSHEET);
    const dataSheet = getSheetByNameCaseInsensitive(ss, DATASHEET);
    
    if (!studentSheet || !dataSheet) {
      return { success: false, error: 'Required sheets not found' };
    }
    
    const studentData = studentSheet.getDataRange().getValues();
    const salesData = dataSheet.getDataRange().getValues();
    
    const salesMap = getHeaderMap(salesData[0]);
    const idSalesIdx = salesMap['student code'] ?? salesMap['student id'] ?? 6;
    const qtySalesIdx = salesMap['quantity'] ?? 4;

    let salesCount = {};
    
    for (let i = 1; i < salesData.length; i++) {
      const studentId = salesData[i][idSalesIdx];
      const quantity = parseInt(salesData[i][qtySalesIdx]) || 0;
      
      if (studentId) {
        const key = studentId.toString().trim().toLowerCase();
        if (!salesCount[key]) {
          salesCount[key] = 0;
        }
        salesCount[key] += quantity;
      }
    }
    
    const stuMap = getHeaderMap(studentData[0]);
    const idStuIdx = stuMap['student id'] ?? 0;
    const totalStuIdx = stuMap['total dress'] ?? 5;
    const remStuIdx = stuMap['remaining dress'] ?? 6;

    let updatedCount = 0;
    
    for (let i = 1; i < studentData.length; i++) {
      const studentId = studentData[i][idStuIdx];
      const totalDress = parseInt(studentData[i][totalStuIdx]) || 0;
      
      if (studentId) {
        const key = studentId.toString().trim().toLowerCase();
        const totalSold = salesCount[key] || 0;
        const newRemainingDress = totalDress - totalSold;
        
        studentSheet.getRange(i + 1, remStuIdx + 1).setValue(newRemainingDress);
        updatedCount++;
      }
    }
    
    return {
      success: true,
      message: `Dress recalculated for ${updatedCount} students.`,
      updated: updatedCount
    };
  } catch (error) {
    console.error('Error in recalculateDress: ' + error.toString());
    return { success: false, error: 'Error recalculating dress: ' + error.message };
  }
}

function generateMissingQRCodes() {
  try {
    const ss = getSpreadsheet();
    const studentSheet = getSheetByNameCaseInsensitive(ss, STUDENTSHEET);
    
    if (!studentSheet) {
      return {
        success: false,
        error: 'Students sheet not found',
        processed: 0,
        skipped: 0,
        errors: 0,
        totalStudents: 0
      };
    }
    
    const data = studentSheet.getDataRange().getValues();
    const map = getHeaderMap(data[0]);
    
    let qrColumnIndex = map['qr url'] ?? map['qr code'] ?? -1;
    
    if (qrColumnIndex === -1) {
      studentSheet.getRange(1, studentSheet.getLastColumn() + 1).setValue('QR URL');
      qrColumnIndex = studentSheet.getLastColumn() - 1;
    }
    
    const idIdx = map['student id'] ?? 0;
    const nameIdx = map['student name'] ?? 2;

    let processedCount = 0;
    let skippedCount = 0;
    let errorCount = 0;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const studentId = row[idIdx];
      const studentName = row[nameIdx];
      const existingQrUrl = row[qrColumnIndex];
      
      if (!studentId || !studentName) {
        skippedCount++;
        continue;
      }
      
      if (existingQrUrl && existingQrUrl.toString().trim() !== '' && !existingQrUrl.toString().includes('failed')) {
        skippedCount++;
        continue;
      }
      
      try {
        const qrUrl = generateQRCode(studentId, studentName);
        
        if (qrUrl) {
          studentSheet.getRange(i + 1, qrColumnIndex + 1).setValue(qrUrl);
          processedCount++;
        } else {
          studentSheet.getRange(i + 1, qrColumnIndex + 1).setValue('(Generation failed)');
          errorCount++;
        }
        
        Utilities.sleep(150);
        
      } catch (error) {
        console.error(`Error generating QR for ${studentId}:`, error);
        studentSheet.getRange(i + 1, qrColumnIndex + 1).setValue('(Error: ' + error.message + ')');
        errorCount++;
      }
    }
    
    return {
      success: true,
      processed: processedCount,
      skipped: skippedCount,
      errors: errorCount,
      totalStudents: data.length - 1,
      message: `QR codes generated for ${processedCount} students. ${skippedCount} already had QR codes. ${errorCount} errors.`
    };
  } catch (error) {
    console.error('Error in generateMissingQRCodes: ' + error.toString());
    return {
      success: false,
      error: 'Error generating QR codes: ' + error.message,
      processed: 0,
      skipped: 0,
      errors: 0,
      totalStudents: 0
    };
  }
}

function checkQRCodeStatus() {
  try {
    const ss = getSpreadsheet();
    const studentSheet = getSheetByNameCaseInsensitive(ss, STUDENTSHEET);
    
    if (!studentSheet) {
      return { success: false, error: 'Students sheet not found' };
    }
    
    const data = studentSheet.getDataRange().getValues();
    const map = getHeaderMap(data[0]);
    const qrColumnIndex = map['qr url'] ?? map['qr code'] ?? -1;
    
    if (qrColumnIndex === -1) {
      return {
        success: true,
        qrColumnExists: false,
        totalStudents: Math.max(0, data.length - 1),
        studentsWithQR: 0,
        studentsWithoutQR: Math.max(0, data.length - 1)
      };
    }
    
    let withQR = 0;
    let withoutQR = 0;
    
    for (let i = 1; i < data.length; i++) {
      const qrUrl = data[i][qrColumnIndex];
      if (qrUrl && qrUrl.toString().trim() !== '' && !qrUrl.toString().includes('failed')) {
        withQR++;
      } else {
        withoutQR++;
      }
    }
    
    return {
      success: true,
      qrColumnExists: true,
      totalStudents: Math.max(0, data.length - 1),
      studentsWithQR: withQR,
      studentsWithoutQR: withoutQR,
      percentageComplete: data.length > 1 ? Math.round((withQR / (data.length - 1)) * 100) : 0
    };
  } catch (error) {
    console.error('Error in checkQRCodeStatus: ' + error.toString());
    return { success: false, error: 'Error checking QR status: ' + error.message };
  }
}

function getAllClasses() {
  try {
    const ss = getSpreadsheet();
    const studentSheet = getSheetByNameCaseInsensitive(ss, STUDENTSHEET);
    
    if (!studentSheet) {
      return [];
    }
    
    const data = studentSheet.getDataRange().getValues();
    const map = getHeaderMap(data[0]);
    const classIdx = map['class'] ?? 1;

    const classes = new Set();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][classIdx]) {
        classes.add(data[i][classIdx]);
      }
    }
    
    return Array.from(classes).sort();
  } catch (error) {
    console.error('Error in getAllClasses: ' + error.toString());
    return [];
  }
}

function getStudentsForPrinting(studentIds) {
  try {
    const ss = getSpreadsheet();
    const studentSheet = getSheetByNameCaseInsensitive(ss, STUDENTSHEET);
    
    if (!studentSheet) {
      return [];
    }
    
    const data = studentSheet.getDataRange().getValues();
    const map = getHeaderMap(data[0]);
    const idIdx = map['student id'] ?? 0;
    const classIdx = map['class'] ?? 1;
    const nameIdx = map['student name'] ?? 2;
    const eventIdx = map['song/event name'] ?? 3;
    const priceIdx = map['price'] ?? 4;
    const totalDressIdx = map['total dress'] ?? 5;
    const remDressIdx = map['remaining dress'] ?? 6;
    const phoneIdx = map['phone no'] ?? 7;
    const dressCodeIdx = map['dress code'] ?? 8;
    const qrUrlIdx = map['qr url'] ?? 9;

    const students = [];
    const idSet = new Set(studentIds.map(id => id.toString().trim().toLowerCase()));
    
    for (let i = 1; i < data.length; i++) {
      const studentId = data[i][idIdx];
      
      if (studentId && idSet.has(studentId.toString().trim().toLowerCase())) {
        students.push({
          id: studentId,
          class: data[i][classIdx] || '',
          name: data[i][nameIdx] || '',
          songEvent: data[i][eventIdx] || '',
          price: data[i][priceIdx] || 0,
          totalDress: data[i][totalDressIdx] || 0,
          remainingDress: data[i][remDressIdx] || 0,
          phoneNo: data[i][phoneIdx] || '',
          dressCode: data[i][dressCodeIdx] || '',
          qrUrl: data[i][qrUrlIdx] || ''
        });
      }
    }
    
    return students;
  } catch (error) {
    console.error('Error in getStudentsForPrinting: ' + error.toString());
    return [];
  }
}

function initializeSheets() {
  try {
    const ss = getSpreadsheet();
    
    let dataSheet = getSheetByNameCaseInsensitive(ss, DATASHEET);
    if (!dataSheet) {
      dataSheet = ss.insertSheet(DATASHEET);
      dataSheet.appendRow(['Timestamp', 'Student Name', 'Class', 'Song/Event', 'Quantity', 'Price', 'Student Code', 'Phone No', 'Dress Code', 'Total Amount']);
    }
    
    let studentSheet = getSheetByNameCaseInsensitive(ss, STUDENTSHEET);
    if (!studentSheet) {
      studentSheet = ss.insertSheet(STUDENTSHEET);
      studentSheet.appendRow(['Student ID', 'Class', 'Student Name', 'Song/Event Name', 'Price', 'Total Dress', 'Remaining Dress', 'Phone No', 'Dress Code', 'QR URL', 'Registration Date']);
    }
    
    let dropdownSheet = getSheetByNameCaseInsensitive(ss, DROPDOWN_SHEET);
    if (!dropdownSheet) {
      dropdownSheet = ss.insertSheet(DROPDOWN_SHEET);
      dropdownSheet.getRange('A1').setValue('Class');
      dropdownSheet.getRange('B1').setValue('Song/Event Name');
      dropdownSheet.getRange('C1').setValue('Dress Code');
      
      dropdownSheet.getRange('A2').setValue('Class 1');
      dropdownSheet.getRange('A3').setValue('Class 2');
      dropdownSheet.getRange('A4').setValue('Class 3');
      
      dropdownSheet.getRange('B2').setValue('Song 1');
      dropdownSheet.getRange('B3').setValue('Song 2');
      dropdownSheet.getRange('B4').setValue('Event 1');
      
      dropdownSheet.getRange('C2').setValue('Traditional');
      dropdownSheet.getRange('C3').setValue('Western');
      dropdownSheet.getRange('C4').setValue('Casual');
    }
    
    return {
      success: true,
      message: 'Sheets initialized successfully'
    };
  } catch (error) {
    console.error('Error in initializeSheets: ' + error.toString());
    return { success: false, error: 'Error initializing sheets: ' + error.message };
  }
}

function testSheets() {
  try {
    const ss = getSpreadsheet();
    const sheets = ss.getSheets();
    const sheetNames = sheets.map(sheet => sheet.getName());
    
    return {
      spreadsheetId: ss.getId(),
      sheets: sheetNames,
      hasDataSheet: sheetNames.some(s => s.toLowerCase() === DATASHEET.toLowerCase()),
      hasStudentSheet: sheetNames.some(s => s.toLowerCase() === STUDENTSHEET.toLowerCase()),
      hasDropdownSheet: sheetNames.some(s => s.toLowerCase() === DROPDOWN_SHEET.toLowerCase())
    };
  } catch (error) {
    console.error('Error in testSheets: ' + error.toString());
    return { success: false, error: 'Error testing sheets: ' + error.message };
  }
}

function bulkUpdateStudents(studentUpdates) {
  try {
    const ss = getSpreadsheet();
    const studentSheet = getSheetByNameCaseInsensitive(ss, STUDENTSHEET);
    
    if (!studentSheet) {
      return { success: false, error: 'Students sheet not found' };
    }
    
    const data = studentSheet.getDataRange().getValues();
    const map = getHeaderMap(data[0]);
    
    const idCol = (map['student id'] !== undefined ? map['student id'] : 0) + 1;
    const songEventCol = (map['song/event name'] !== undefined ? map['song/event name'] : 3) + 1;
    const priceCol = (map['price'] !== undefined ? map['price'] : 4) + 1;
    const dressCodeCol = (map['dress code'] !== undefined ? map['dress code'] : 8) + 1;
    
    let updatedCount = 0;
    const errors = [];
    
    const studentMap = {};
    for (let i = 1; i < data.length; i++) {
      const studentId = data[i][idCol - 1];
      if (studentId) {
        studentMap[studentId.toString().trim().toLowerCase()] = i + 1;
      }
    }
    
    studentUpdates.forEach((update) => {
      try {
        const key = update.studentId ? update.studentId.toString().trim().toLowerCase() : '';
        const row = studentMap[key];
        if (row) {
          if (update.songEvent !== undefined && update.songEvent !== null && update.songEvent !== '') {
            studentSheet.getRange(row, songEventCol).setValue(update.songEvent);
          }
          
          if (update.price !== undefined && update.price !== null && update.price !== '') {
            studentSheet.getRange(row, priceCol).setValue(parseFloat(update.price));
          }
          
          if (update.dressCode !== undefined && update.dressCode !== null && update.dressCode !== '') {
            studentSheet.getRange(row, dressCodeCol).setValue(update.dressCode);
          }
          
          updatedCount++;
        } else {
          errors.push(`Student ${update.studentId} not found`);
        }
      } catch (error) {
        errors.push(`Error updating student ${update.studentId}: ${error.message}`);
      }
    });
    
    return {
      success: true,
      updated: updatedCount,
      total: studentUpdates.length,
      errors: errors,
      message: `Updated ${updatedCount} of ${studentUpdates.length} students. ${errors.length > 0 ? 'Errors: ' + errors.join(', ') : ''}`
    };
    
  } catch (error) {
    console.error('Error in bulkUpdateStudents: ' + error.toString());
    return { success: false, error: 'Error in bulk update: ' + error.message };
  }
}

function getDropdownDataForBulkEdit() {
  try {
    const ss = getSpreadsheet();
    const dropdownSheet = getSheetByNameCaseInsensitive(ss, DROPDOWN_SHEET);
    const studentSheet = getSheetByNameCaseInsensitive(ss, STUDENTSHEET);
    
    const result = {
      songEvents: [],
      dressCodes: [],
      existingPrices: [],
      success: true
    };
    
    if (dropdownSheet) {
      const songData = dropdownSheet.getRange('B:B').getValues();
      const songEvents = songData.filter(String).map(row => row[0]);
      if (songEvents.length > 0 && songEvents[0].toString().toLowerCase().includes('song')) {
        songEvents.shift();
      }
      result.songEvents = [...new Set(songEvents)].filter(Boolean);

      const dressData = dropdownSheet.getRange('C:C').getValues();
      const dressCodes = dressData.filter(String).map(row => row[0]);
      if (dressCodes.length > 0 && dressCodes[0].toString().toLowerCase().includes('dress')) {
        dressCodes.shift();
      }
      result.dressCodes = [...new Set(dressCodes)].filter(Boolean);
    }
    
    if (studentSheet) {
      const data = studentSheet.getDataRange().getValues();
      const map = getHeaderMap(data[0]);
      const priceIdx = map['price'] ?? 4;

      const prices = [];
      for (let i = 1; i < data.length; i++) {
        const price = parseFloat(data[i][priceIdx]);
        if (!isNaN(price) && price > 0) {
          prices.push(price);
        }
      }
      
      const priceCounts = {};
      prices.forEach(price => {
        priceCounts[price] = (priceCounts[price] || 0) + 1;
      });
      
      result.existingPrices = Object.keys(priceCounts)
        .map(price => parseFloat(price))
        .sort((a, b) => priceCounts[b] - priceCounts[a])
        .slice(0, 10);
    }
    
    return result;
    
  } catch (error) {
    console.error('Error in getDropdownDataForBulkEdit: ' + error.toString());
    return {
      success: false,
      error: error.message,
      songEvents: [],
      dressCodes: [],
      existingPrices: []
    };
  }
}

function getQRCodeForDownload(studentId) {
  try {
    const ss = getSpreadsheet();
    const studentSheet = getSheetByNameCaseInsensitive(ss, STUDENTSHEET);
    
    if (!studentSheet) {
      return { success: false, error: 'Students sheet not found' };
    }
    
    const data = studentSheet.getDataRange().getValues();
    const map = getHeaderMap(data[0]);
    const idIdx = map['student id'] ?? 0;
    const nameIdx = map['student name'] ?? 2;
    const qrUrlIdx = map['qr url'] ?? 9;

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] && data[i][idIdx].toString().trim().toLowerCase() === studentId.toString().trim().toLowerCase()) {
        const qrUrl = data[i][qrUrlIdx] || '';
        
        if (!qrUrl || qrUrl.trim() === '') {
          return { 
            success: false, 
            error: 'No QR code available for this student' 
          };
        }
        
        try {
          const response = UrlFetchApp.fetch(qrUrl, {
            muteHttpExceptions: true,
            followRedirects: true
          });
          
          if (response.getResponseCode() === 200) {
            const blob = response.getBlob();
            const base64 = Utilities.base64Encode(blob.getBytes());
            const dataUri = 'data:' + blob.getContentType() + ';base64,' + base64;
            
            const studentName = data[i][nameIdx] || 'Unknown';
            const cleanName = studentName.replace(/[^a-z0-9]/gi, '_').replace(/_{2,}/g, '_');
            const filename = `QR_${studentId}_${cleanName}.png`;
            
            return {
              success: true,
              qrUrl: qrUrl,
              downloadUrl: dataUri,
              filename: filename,
              studentName: studentName
            };
          }
        } catch (fetchError) {
          console.log('Direct fetch failed:', fetchError);
        }
        
        const studentName = data[i][nameIdx] || 'Unknown';
        return {
          success: true,
          qrUrl: qrUrl,
          filename: `QR_${studentId}.png`,
          studentName: studentName
        };
      }
    }
    
    return { success: false, error: 'Student not found' };
    
  } catch (error) {
    console.error('Error in getQRCodeForDownload:', error);
    return { success: false, error: 'Error retrieving QR code: ' + error.message };
  }
}

function downloadQRCode(studentId) {
  return getQRCodeForDownload(studentId);
}

function generatePDFReceipt(receiptData) {
  try {
    const htmlTemplate = `
      <!DOCTYPE html>
      <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; }
          .receipt-container { max-width: 400px; margin: 0 auto; background: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
          .header { text-align: center; border-bottom: 2px solid #10b981; padding-bottom: 15px; margin-bottom: 20px; }
          .school-name { font-size: 24px; font-weight: bold; color: #10b981; margin-bottom: 5px; }
          .receipt-title { font-size: 18px; color: #333; margin-bottom: 10px; }
          .receipt-info { margin-bottom: 20px; }
          .info-row { display: flex; justify-content: space-between; margin-bottom: 8px; padding-bottom: 8px; border-bottom: 1px dashed #ddd; }
          .info-label { font-weight: bold; color: #555; }
          .info-value { color: #333; }
          .total-section { background: #f8f9fa; padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #10b981; }
          .total-row { display: flex; justify-content: space-between; font-size: 18px; font-weight: bold; }
          .footer { text-align: center; margin-top: 20px; padding-top: 15px; border-top: 1px solid #ddd; color: #666; font-size: 12px; }
          .date-time { text-align: center; color: #666; margin-bottom: 15px; font-size: 14px; }
        </style>
      </head>
      <body>
        <div class="receipt-container">
          <div class="header">
            <div class="school-name">SUCCESS SCHOOL INDI</div>
            <div class="receipt-title">Dress Sale Receipt</div>
          </div>
          
          <div class="date-time">
            ${receiptData.timestamp || new Date().toLocaleString('en-IN')}
          </div>
          
          <div class="receipt-info">
            <div class="info-row"><span class="info-label">Student:</span><span class="info-value">${receiptData.studentName || ''}</span></div>
            <div class="info-row"><span class="info-label">Class:</span><span class="info-value">${receiptData.className || ''}</span></div>
            <div class="info-row"><span class="info-label">Event:</span><span class="info-value">${receiptData.songEvent || ''}</span></div>
            <div class="info-row"><span class="info-label">Student ID:</span><span class="info-value">${receiptData.studentId || ''}</span></div>
            <div class="info-row"><span class="info-label">Phone:</span><span class="info-value">${receiptData.phoneNo || 'N/A'}</span></div>
            <div class="info-row"><span class="info-label">Dress Code:</span><span class="info-value">${receiptData.dressCode || 'N/A'}</span></div>
            <div class="info-row"><span class="info-label">Transaction ID:</span><span class="info-value">${receiptData.transactionId || generateTransactionId()}</span></div>
          </div>
          
          <div class="total-section">
            <div class="info-row"><span class="info-label">Price per Dress:</span><span class="info-value">₹${parseFloat(receiptData.price || 0).toFixed(2)}</span></div>
            <div class="info-row"><span class="info-label">Quantity:</span><span class="info-value">${receiptData.quantity || 0}</span></div>
            <div class="total-row"><span class="info-label">TOTAL AMOUNT:</span><span class="info-value">₹${parseFloat(receiptData.totalAmount || 0).toFixed(2)}</span></div>
          </div>
          
          <div class="footer">
            <p>Thank you for your purchase!</p>
            <p>SUCCESS SCHOOL INDI<br>Contact: 9964882070<br>This is a computer generated receipt</p>
          </div>
        </div>
      </body>
      </html>
    `;
    
    const blob = Utilities.newBlob(htmlTemplate, 'text/html');
    const pdf = blob.getAs('application/pdf');
    
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const filename = `Receipt_${receiptData.studentId || 'STUDENT'}_${timestamp}.pdf`;
    
    const folder = getOrCreateReceiptFolder();
    const file = folder.createFile(pdf);
    file.setName(filename);
    
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const pdfUrl = file.getUrl();
    
    return {
      success: true,
      pdfUrl: pdfUrl,
      fileId: file.getId(),
      filename: filename,
      transactionId: receiptData.transactionId || generateTransactionId()
    };
    
  } catch (error) {
    console.error('Error generating PDF:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

function getOrCreateReceiptFolder() {
  try {
    const folders = DriveApp.getFoldersByName(RECEIPT_FOLDER_NAME);
    if (folders.hasNext()) {
      return folders.next();
    }
    
    return DriveApp.createFolder(RECEIPT_FOLDER_NAME);
  } catch (error) {
    console.error('Error accessing receipt folder:', error);
    throw error;
  }
}

function generateTransactionId() {
  const timestamp = new Date().getTime();
  const randomNum = Math.floor(Math.random() * 10000);
  return `TXN${timestamp.toString().slice(-10)}${randomNum.toString().padStart(4, '0')}`;
}

function getWhatsAppLink(phoneNo, receiptData) {
  try {
    const cleanPhone = phoneNo.toString().replace(/\D/g, '');
    let whatsappPhone = cleanPhone;
    
    if (!cleanPhone.startsWith('91') && cleanPhone.length === 10) {
      whatsappPhone = '91' + cleanPhone;
    }
    
    let message = `*Dress Purchase Receipt - Success School Indi*\n\n`;
    message += `*Student:* ${receiptData.studentName || ''}\n`;
    message += `*Class:* ${receiptData.className || ''}\n`;
    message += `*Event:* ${receiptData.songEvent || ''}\n`;
    message += `*Date:* ${receiptData.timestamp || new Date().toLocaleString('en-IN')}\n\n`;
    message += `*Dress Details:*\n`;
    message += `Price per Dress: ₹${parseFloat(receiptData.price || 0).toFixed(2)}\n`;
    message += `Quantity: ${receiptData.quantity || 0}\n`;
    message += `Dress Total: ₹${parseFloat(receiptData.dressTotal || 0).toFixed(2)}\n\n`;
    
    if (receiptData.fees && Object.keys(receiptData.fees).length > 0) {
      message += `*Additional Fees:*\n`;
      Object.keys(receiptData.fees).forEach(feeName => {
        const feeAmount = receiptData.fees[feeName];
        message += `${feeName}: ₹${feeAmount.toFixed(2)}\n`;
      });
      message += `Total Fees: ₹${parseFloat(receiptData.totalFees || 0).toFixed(2)}\n\n`;
    }
    
    message += `*GRAND TOTAL: ₹${parseFloat(receiptData.grandTotal || 0).toFixed(2)}*\n\n`;
    message += `Thank you for your purchase!\n`;
    message += `Success School Indi\n`;
    message += `Contact: 9964882070`;
    
    const encodedMessage = encodeURIComponent(message);
    const whatsappLink = `https://api.whatsapp.com/send?phone=${whatsappPhone}&text=${encodedMessage}`;
    
    return {
      success: true,
      whatsappLink: whatsappLink,
      message: message
    };
  } catch (error) {
    console.error('Error generating WhatsApp link:', error);
    return { success: false, error: error.toString() };
  }
}

function sendReceiptViaWhatsApp(phoneNumber, receiptData) {
  try {
    const result = getWhatsAppLink(phoneNumber, receiptData);
    
    if (!result.success) {
      throw new Error(result.error);
    }
    
    return {
      success: true,
      whatsappLink: result.whatsappLink,
      pdfUrl: result.pdfUrl,
      transactionId: result.transactionId,
      message: `WhatsApp link generated for ${phoneNumber}. Click the link to send the receipt.`
    };
    
  } catch (error) {
    console.error('Error sending receipt via WhatsApp:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

function getStudentPhoneNumber(studentId) {
  try {
    const ss = getSpreadsheet();
    const studentSheet = getSheetByNameCaseInsensitive(ss, STUDENTSHEET);
    
    if (!studentSheet) {
      return { success: false, error: 'Students sheet not found' };
    }
    
    const data = studentSheet.getDataRange().getValues();
    const map = getHeaderMap(data[0]);
    const idIdx = map['student id'] ?? 0;
    const phoneIdx = map['phone no'] ?? 7;
    const nameIdx = map['student name'] ?? 2;
    const classIdx = map['class'] ?? 1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] && data[i][idIdx].toString().trim().toLowerCase() === studentId.toString().trim().toLowerCase()) {
        const phoneNo = data[i][phoneIdx] || '';
        if (phoneNo && phoneNo.toString().trim() !== '') {
          return {
            success: true,
            phoneNumber: phoneNo.toString().trim(),
            studentName: data[i][nameIdx] || '',
            className: data[i][classIdx] || ''
          };
        } else {
          return {
            success: false,
            error: 'No phone number found for this student'
          };
        }
      }
    }
    
    return { success: false, error: 'Student not found' };
    
  } catch (error) {
    console.error('Error getting student phone number:', error);
    return { success: false, error: error.toString() };
  }
}

function processDressSaleWithReceipt(formObject) {
  try {
    const saleResult = processDressSale(formObject);
    
    if (saleResult.success) {
      const receiptData = {
        studentId: formObject.studentCode,
        studentName: formObject.student_name,
        className: formObject.class_name,
        songEvent: formObject.song_event,
        phoneNo: formObject.phone_no || '',
        dressCode: formObject.dress_code || '',
        price: parseFloat(formObject.price) || 0,
        quantity: parseInt(formObject.quantity) || 0,
        totalAmount: (parseFloat(formObject.price) || 0) * (parseInt(formObject.quantity) || 0),
        timestamp: new Date().toLocaleString('en-IN', { 
          timeZone: 'Asia/Kolkata',
          hour12: false 
        })
      };
      
      return {
        ...saleResult,
        receiptData: receiptData,
        receiptReady: true
      };
    } else {
      return saleResult;
    }
  } catch (error) {
    console.error('Error in processDressSaleWithReceipt:', error);
    return {
      success: false,
      error: 'Error processing sale with receipt: ' + error.message
    };
  }
}

function getAllFeeTypes() {
  try {
    const ss = getSpreadsheet();
    const dropdownSheet = getSheetByNameCaseInsensitive(ss, DROPDOWN_SHEET);
    
    if (!dropdownSheet) {
      return { success: false, error: 'Dropdown sheet not found' };
    }
    
    const data = dropdownSheet.getDataRange().getValues();
    const headers = data[0];
    
    const feeTypes = [];
    for (let i = 3; i < headers.length; i++) {
      if (headers[i] && headers[i].toString().includes('Fees')) {
        feeTypes.push({
          name: headers[i],
          column: i
        });
      }
    }
    
    return {
      success: true,
      feeTypes: feeTypes,
      headers: headers
    };
    
  } catch (error) {
    console.error('Error in getAllFeeTypes: ' + error.toString());
    return { success: false, error: error.message };
  }
}

function getClassFees(className) {
  try {
    const ss = getSpreadsheet();
    const dropdownSheet = getSheetByNameCaseInsensitive(ss, DROPDOWN_SHEET);
    
    if (!dropdownSheet) {
      return { success: false, error: 'Dropdown sheet not found' };
    }
    
    const data = dropdownSheet.getDataRange().getValues();
    const headers = data[0];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toLowerCase() === className.toString().trim().toLowerCase()) {
        const fees = {};
        
        for (let j = 3; j < headers.length; j++) {
          if (headers[j] && headers[j].toString().includes('Fees')) {
            const feeValue = parseFloat(data[i][j]) || 0;
            if (feeValue > 0) {
              fees[headers[j]] = feeValue;
            }
          }
        }
        
        return {
          success: true,
          className: className,
          fees: fees,
          totalFees: Object.values(fees).reduce((sum, fee) => sum + fee, 0)
        };
      }
    }
    
    return {
      success: true,
      className: className,
      fees: {},
      totalFees: 0
    };
    
  } catch (error) {
    console.error('Error in getClassFees: ' + error.toString());
    return { success: false, error: error.message };
  }
}

function saveClassFees(classFees) {
  try {
    const ss = getSpreadsheet();
    let dropdownSheet = getSheetByNameCaseInsensitive(ss, DROPDOWN_SHEET);
    
    if (!dropdownSheet) {
      dropdownSheet = ss.insertSheet(DROPDOWN_SHEET);
      dropdownSheet.getRange('A1').setValue('Class');
      
      const feeHeaders = [
        'Admission Fees', 'Annual Day Fees', 'Sports Fees', 'Exam Fees', 
        'Library Fees', 'Medical Fees', 'Tour Fees', 'Lab Fees', 
        'Penalty Fees', 'Other Fees'
      ];
      
      const headers = ['Class', 'Song/Event Name', 'Dress Code', ...feeHeaders];
      headers.forEach((header, index) => {
        dropdownSheet.getRange(1, index + 1).setValue(header);
      });
    }
    
    const data = dropdownSheet.getDataRange().getValues();
    let classRow = -1;
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toLowerCase() === classFees.className.toString().trim().toLowerCase()) {
        classRow = i + 1;
        break;
      }
    }
    
    if (classRow === -1) {
      classRow = data.length + 1;
      dropdownSheet.getRange(classRow, 1).setValue(classFees.className);
    }
    
    const headers = data[0];
    
    if (Array.isArray(classFees.fees)) {
      classFees.fees.forEach(fee => {
        const feeName = fee.name;
        const feeAmount = fee.amount;
        const columnIndex = headers.indexOf(feeName);
        
        if (columnIndex !== -1) {
          dropdownSheet.getRange(classRow, columnIndex + 1).setValue(feeAmount);
        }
      });
    } else if (typeof classFees.fees === 'object') {
      Object.keys(classFees.fees).forEach(feeName => {
        const feeAmount = classFees.fees[feeName];
        const columnIndex = headers.indexOf(feeName);
        
        if (columnIndex !== -1) {
          dropdownSheet.getRange(classRow, columnIndex + 1).setValue(feeAmount);
        }
      });
    }
    
    return {
      success: true,
      message: 'Class fees updated successfully'
    };
    
  } catch (error) {
    console.error('Error in saveClassFees: ' + error.toString());
    return { success: false, error: error.message };
  }
}

function addNewFeeType(feeName) {
  try {
    const ss = getSpreadsheet();
    const dropdownSheet = getSheetByNameCaseInsensitive(ss, DROPDOWN_SHEET);
    
    if (!dropdownSheet) {
      return { success: false, error: 'Dropdown sheet not found' };
    }
    
    const lastColumn = dropdownSheet.getLastColumn();
    dropdownSheet.getRange(1, lastColumn + 1).setValue(feeName);
    
    return {
      success: true,
      message: `Fee type "${feeName}" added successfully`
    };
    
  } catch (error) {
    console.error('Error adding fee type:', error);
    return { success: false, error: error.message };
  }
}

function getAllFeeCategories() {
  try {
    const ss = getSpreadsheet();
    const dropdownSheet = getSheetByNameCaseInsensitive(ss, DROPDOWN_SHEET);
    
    if (!dropdownSheet) {
      return { success: false, error: 'Dropdown sheet not found' };
    }
    
    const data = dropdownSheet.getDataRange().getValues();
    const headers = data[0];
    
    const feeCategories = [];
    for (let i = 3; i < headers.length; i++) {
      if (headers[i] && headers[i].toString().includes('Fees')) {
        feeCategories.push(headers[i]);
      }
    }
    
    return {
      success: true,
      feeCategories: feeCategories
    };
    
  } catch (error) {
    console.error('Error getting fee categories:', error);
    return { success: false, error: error.message };
  }
}

function initializeFeesInDropdown() {
  try {
    const ss = getSpreadsheet();
    let dropdownSheet = getSheetByNameCaseInsensitive(ss, DROPDOWN_SHEET);
    
    if (!dropdownSheet) {
      dropdownSheet = ss.insertSheet(DROPDOWN_SHEET);
    }
    
    const feeHeaders = [
      'Admission Fees', 'Annual Day Fees', 'Sports Fees', 'Exam Fees', 
      'Library Fees', 'Medical Fees', 'Tour Fees', 'Lab Fees', 
      'Penalty Fees', 'Other Fees'
    ];
    
    const lastColumn = dropdownSheet.getLastColumn();
    const headers = [];
    for (let i = 1; i <= lastColumn; i++) {
      headers.push(dropdownSheet.getRange(1, i).getValue());
    }
    
    let currentColumn = lastColumn + 1;
    feeHeaders.forEach(feeHeader => {
      if (!headers.includes(feeHeader)) {
        dropdownSheet.getRange(1, currentColumn).setValue(feeHeader);
        currentColumn++;
      }
    });
    
    return {
      success: true,
      message: 'Fee structure initialized in dropdown sheet'
    };
    
  } catch (error) {
    console.error('Error initializing fees:', error);
    return { success: false, error: error.message };
  }
}

function getUnpaidFees(studentId) {
  try {
    const ss = getSpreadsheet();
    const dataSheet = getSheetByNameCaseInsensitive(ss, DATASHEET);
    
    if (!dataSheet) {
      return { success: true, unpaidFees: [], totalUnpaid: 0 };
    }
    
    const data = dataSheet.getDataRange().getValues();
    const headers = data[0];
    
    const feeColumns = [];
    headers.forEach((header, index) => {
      if (header && header.toString().includes('Fees') && !['Total Fees', 'Grand Total'].includes(header.toString())) {
        feeColumns.push({ name: header, index: index });
      }
    });
    
    let totalUnpaid = 0;
    const unpaidFees = [];
    const map = getHeaderMap(headers);
    const idIdx = map['student code'] ?? map['student id'] ?? 6;

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] && data[i][idIdx].toString().trim().toLowerCase() === studentId.toString().trim().toLowerCase()) {
        const timestamp = data[i][0];
        
        feeColumns.forEach(feeCol => {
          const feeAmount = parseFloat(data[i][feeCol.index]) || 0;
          if (feeAmount > 0) {
            unpaidFees.push({
              transactionDate: timestamp,
              feeName: feeCol.name,
              amount: feeAmount,
              transactionRow: i + 1
            });
            totalUnpaid += feeAmount;
          }
        });
      }
    }
    
    return {
      success: true,
      unpaidFees: unpaidFees,
      totalUnpaid: totalUnpaid,
      hasUnpaidFees: totalUnpaid > 0
    };
    
  } catch (error) {
    console.error('Error getting unpaid fees:', error);
    return { success: false, error: error.message };
  }
}

function processAdditionalPayment(formObject) {
  try {
    const ss = getSpreadsheet();
    const dataSheet = getSheetByNameCaseInsensitive(ss, DATASHEET);
    
    if (!dataSheet) {
      return { success: false, error: 'Data sheet not found' };
    }
    
    const student = lookupStudent(formObject.studentCode);
    if (!student) {
      return { 
        success: false, 
        error: `Student with code "${formObject.studentCode}" not found.` 
      };
    }
    
    const timestamp = new Date().toLocaleString('en-IN', { 
      timeZone: 'Asia/Kolkata',
      hour12: false 
    });
    
    let selectedFees = {};
    let totalFees = 0;
    
    if (formObject.feeCategories) {
      try {
        const feeCategories = JSON.parse(formObject.feeCategories);
        const classFeesResult = getClassFees(formObject.class_name || student.className);
        
        if (classFeesResult.success) {
          const fees = classFeesResult.fees;
          
          feeCategories.forEach(feeName => {
            if (fees[feeName] && fees[feeName] > 0) {
              selectedFees[feeName] = fees[feeName];
              totalFees += fees[feeName];
            }
          });
        }
      } catch (e) {
        console.log('Error parsing fee categories:', e);
      }
    }
    
    if (totalFees === 0) {
      return { 
        success: false, 
        error: 'No fees selected for payment.' 
      };
    }
    
    const rowData = [
      timestamp,
      formObject.student_name || student.studentName,
      formObject.class_name || student.className,
      formObject.song_event || student.songEvent,
      0, // Quantity = 0 for additional payments
      0, // Price = 0 for additional payments
      formObject.studentCode,
      formObject.phone_no || student.phoneNo || '',
      formObject.dress_code || student.dressCode || '',
      0 // Dress Total = 0
    ];
    
    const feeHeaders = [
      'Admission Fees', 'Annual Day Fees', 'Sports Fees', 'Exam Fees', 
      'Library Fees', 'Medical Fees', 'Tour Fees', 'Lab Fees', 
      'Penalty Fees', 'Other Fees'
    ];
    
    feeHeaders.forEach(feeHeader => {
      rowData.push(selectedFees[feeHeader] || 0);
    });
    
    rowData.push(totalFees); // Total Fees
    rowData.push(totalFees); // Grand Total
    
    dataSheet.appendRow(rowData);
    
    return { 
      success: true, 
      message: 'Additional payment processed successfully!',
      studentData: student,
      paymentData: {
        timestamp: timestamp,
        studentName: formObject.student_name || student.studentName,
        className: formObject.class_name || student.className,
        studentId: formObject.studentCode,
        fees: selectedFees,
        totalFees: totalFees,
        type: 'additional_payment'
      }
    };
    
  } catch (error) {
    Logger.log('Error in processAdditionalPayment: ' + error.toString());
    return { 
      success: false, 
      error: 'Error processing additional payment: ' + error.message 
    };
  }
}

function getStudentFeeHistory(studentId) {
  try {
    const ss = getSpreadsheet();
    const dataSheet = getSheetByNameCaseInsensitive(ss, DATASHEET);
    
    if (!dataSheet) {
      return { success: true, transactions: [] };
    }
    
    const data = dataSheet.getDataRange().getValues();
    const headers = data[0];
    const map = getHeaderMap(headers);
    const idIdx = map['student code'] ?? map['student id'] ?? 6;

    const transactions = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] && data[i][idIdx].toString().trim().toLowerCase() === studentId.toString().trim().toLowerCase()) {
        const transaction = {
          timestamp: data[i][0],
          type: parseFloat(data[i][4]) > 0 ? 'dress_sale' : 'fee_payment',
          quantity: data[i][4],
          price: data[i][5],
          dressTotal: data[i][9],
          totalFees: data[i][20] || 0,
          grandTotal: data[i][21] || 0,
          rowNumber: i + 1
        };
        
        transaction.fees = {};
        headers.forEach((header, index) => {
          if (header && header.toString().includes('Fees') && !['Total Fees', 'Grand Total'].includes(header.toString())) {
            const feeAmount = parseFloat(data[i][index]) || 0;
            if (feeAmount > 0) {
              transaction.fees[header] = feeAmount;
            }
          }
        });
        
        transactions.push(transaction);
      }
    }
    
    return {
      success: true,
      transactions: transactions,
      totalTransactions: transactions.length
    };
    
  } catch (error) {
    console.error('Error getting fee history:', error);
    return { success: false, error: error.message };
  }
}

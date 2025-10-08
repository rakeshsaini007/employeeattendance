function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Attendance App")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function ensureSheets() {
  const ss = SpreadsheetApp.getActive();
  let emp = ss.getSheetByName("Employees");
  if (!emp) {
    emp = ss.insertSheet("Employees");
    emp.appendRow(["Employee ID", "Employee Name", "Category", "Adhar Number", "Mobile Number"]);
  } else {
    const headers = emp.getRange(1, 1, 1, emp.getLastColumn()).getValues()[0];
    if (!headers.includes("Employee ID")) {
      emp.insertColumnBefore(1);
      emp.getRange(1, 1).setValue("Employee ID");
      if (emp.getLastRow() > 1) {
        const data = emp.getRange(2, 2, emp.getLastRow() - 1, 3).getValues();
        for (let i = 0; i < data.length; i++) {
          const empId = _generateEmployeeId(data[i][0], data[i][2]);
          emp.getRange(i + 2, 1).setValue(empId);
        }
      }
    }
  }
  let att = ss.getSheetByName("Attendance");
  if (!att) {
    att = ss.insertSheet("Attendance");
    att.appendRow([
      "Date","Employee ID","Employee Name","Category","Adhar Number","Mobile Number",
      "Entry Time","Exit Time","Total Hours","Duty","O.T.","Month-Year"
    ]);
  } else {
    const headers = att.getRange(1, 1, 1, att.getLastColumn()).getValues()[0];
    if (!headers.includes("Month-Year")) {
      const lastCol = att.getLastColumn();
      att.getRange(1, lastCol + 1).setValue("Month-Year");
    }
    if (!headers.includes("Employee ID")) {
      att.insertColumnAfter(1);
      att.getRange(1, 2).setValue("Employee ID");
    }
  }
  let users = ss.getSheetByName("Users");
  if (!users) {
    users = ss.insertSheet("Users");
    users.appendRow(["Username", "Password"]);
    users.appendRow(["admin", "admin123"]);
  }
  return { emp, att, users };
}

/* -------------------------
   Employee functions
   ------------------------- */
function addEmployee(payload) {
  try {
    const { emp } = ensureSheets();
    const name = String(payload.name || "").trim().toUpperCase();
    const category = String(payload.category || "").trim().toUpperCase();
    const adhar = String(payload.adhar || "").trim();
    const mobile = String(payload.mobile || "").trim();

    if (!/^\d{12}$/.test(adhar)) return { success: false, error: "Aadhaar must be 12 digits" };
    if (!/^\d{10}$/.test(mobile)) return { success: false, error: "Mobile must be 10 digits" };

    if (emp.getLastRow() >= 2) {
      const existingData = emp.getRange(2, 4, emp.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < existingData.length; i++) {
        if (String(existingData[i][0]).trim() === adhar) {
          return { success: false, error: "Aadhaar number already exists. Please use a unique Aadhaar number." };
        }
      }
    }

    const employeeId = _generateEmployeeId(name, adhar);
    emp.appendRow([employeeId, name, category, adhar, mobile]);
    return { success: true, employeeId: employeeId };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getEmployees() {
  try {
    const { emp } = ensureSheets();
    if (emp.getLastRow() < 2) return [];
    const data = emp.getRange(2, 1, emp.getLastRow() - 1, 5).getValues();
    return data.map(r => ({ empId: r[0], name: r[1], category: r[2], adhar: String(r[3]), mobile: String(r[4]) }));
  } catch (e) {
    throw new Error("getEmployees error: " + e.message);
  }
}

function getEmployeeById(empId) {
  try {
    const { emp } = ensureSheets();
    if (emp.getLastRow() < 2) return { success: false, error: "No employees found" };

    const data = emp.getRange(2, 1, emp.getLastRow() - 1, 5).getValues();
    const upperEmpId = String(empId).trim().toUpperCase();

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toUpperCase() === upperEmpId) {
        return {
          success: true,
          employee: {
            empId: data[i][0],
            name: data[i][1],
            category: data[i][2],
            adhar: String(data[i][3]),
            mobile: String(data[i][4])
          }
        };
      }
    }

    return { success: false, error: "Employee ID not found" };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function updateEmployee(payload) {
  try {
    const { emp, att } = ensureSheets();
    if (emp.getLastRow() < 2) return { success: false, error: "No employees found" };

    const oldEmpId = String(payload.empId || "").trim().toUpperCase();
    const name = String(payload.name || "").trim().toUpperCase();
    const category = String(payload.category || "").trim().toUpperCase();
    const adhar = String(payload.adhar || "").trim();
    const mobile = String(payload.mobile || "").trim();

    if (!/^\d{12}$/.test(adhar)) return { success: false, error: "Aadhaar must be 12 digits" };
    if (!/^\d{10}$/.test(mobile)) return { success: false, error: "Mobile must be 10 digits" };

    const data = emp.getRange(2, 1, emp.getLastRow() - 1, 5).getValues();
    let oldName = "";
    let currentRowIndex = -1;

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toUpperCase() === oldEmpId) {
        currentRowIndex = i;
        oldName = String(data[i][1]).trim().toUpperCase();
        break;
      }
    }

    if (currentRowIndex === -1) {
      return { success: false, error: "Employee ID not found" };
    }

    for (let i = 0; i < data.length; i++) {
      if (i !== currentRowIndex && String(data[i][3]).trim() === adhar) {
        return { success: false, error: "Aadhaar number already exists for another employee. Please use a unique Aadhaar number." };
      }
    }

    const newEmpId = _generateEmployeeId(name, adhar);

    const rowNum = currentRowIndex + 2;
    emp.getRange(rowNum, 1).setValue(newEmpId);
    emp.getRange(rowNum, 2).setValue(name);
    emp.getRange(rowNum, 3).setValue(category);
    emp.getRange(rowNum, 4).setValue(adhar);
    emp.getRange(rowNum, 5).setValue(mobile);

    _updateAttendanceRecords(att, oldEmpId, newEmpId, oldName, name, adhar, mobile);

    return { success: true, message: "Employee updated successfully across all records", newEmployeeId: newEmpId };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function _updateAttendanceRecords(att, oldEmpId, newEmpId, oldName, newName, newAdhar, newMobile) {
  try {
    if (att.getLastRow() < 2) return;

    const data = att.getRange(2, 1, att.getLastRow() - 1, 12).getValues();

    for (let i = 0; i < data.length; i++) {
      const rowEmpId = String(data[i][1]).trim().toUpperCase();
      const rowEmpName = String(data[i][2]).trim().toUpperCase();

      if (rowEmpId === oldEmpId || rowEmpName === oldName) {
        const rowNum = i + 2;

        att.getRange(rowNum, 2).setValue(newEmpId);
        att.getRange(rowNum, 3).setValue(newName);
        att.getRange(rowNum, 5).setValue(newAdhar);
        att.getRange(rowNum, 6).setValue(newMobile);
      }
    }
  } catch (e) {
    throw new Error("Error updating attendance records: " + e.message);
  }
}

/* -------------------------
   Attendance functions (Modified)
   ------------------------- */
function markAttendance(payload) {
  try {
    const { att } = ensureSheets();
    const name = String(payload.name || "").trim().toUpperCase();
    const category = String(payload.category || "").trim().toUpperCase();
    const adhar = String(payload.adhar || "").trim();
    const mobile = String(payload.mobile || "").trim();

    const now = new Date();
    const roundedNow = _roundToNearest15Minutes(now);
    const today = Utilities.formatDate(roundedNow, Session.getScriptTimeZone(), "yyyy-MM-dd");

    let hours = roundedNow.getHours();
    const minutes = String(roundedNow.getMinutes()).padStart(2, '0');
    const ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    hours = hours ? hours : 12;
    const nowTime12Hr = hours + ':' + minutes + ' ' + ampm;

    const employeeId = _generateEmployeeId(name, adhar);

    const data = att.getDataRange().getValues();
    let foundRow = -1;
    for (let i = 1; i < data.length; i++) {
      const [date, empId, empName, , , , entry, exit] = data[i];
      if (_cellToDateString(date) === today && empName === name) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow === -1) {
      const monthYear = Utilities.formatDate(roundedNow, Session.getScriptTimeZone(), "MMM-yyyy");
      att.appendRow([today, employeeId, name, category, adhar, mobile, nowTime12Hr, "", "", "", "", monthYear]);
      const range = att.getRange(att.getLastRow(), 7);
      range.setNumberFormat('@STRING@');
      return { success: true, message: `Entry marked at ${nowTime12Hr}` };
    } else {
      const exitCell = att.getRange(foundRow, 8);
      if (exitCell.getValue() === "" || String(exitCell.getValue()).trim() === "") {
        exitCell.setValue(nowTime12Hr);
        exitCell.setNumberFormat('@STRING@');

        const entryTimeRaw = att.getRange(foundRow, 7).getValue();
        const entryTime24 = _convertTo24Hr(String(entryTimeRaw));
        const exitTime24 = _convertTo24Hr(nowTime12Hr);

        const entryDate = new Date(today + "T" + entryTime24);
        const exitDate = new Date(today + "T" + exitTime24);
        let hours = (exitDate - entryDate) / (1000 * 60 * 60);
        if (hours < 0) hours += 24;
        hours = Math.round(hours * 100) / 100;

        const duty = Math.floor(hours / 8);
        const ot = Math.max(0, Math.round((hours - duty * 8) * 100) / 100);

        const monthYear = Utilities.formatDate(new Date(today), Session.getScriptTimeZone(), "MMM-yyyy");
        att.getRange(foundRow, 9).setValue(hours);
        att.getRange(foundRow, 10).setValue(duty);
        att.getRange(foundRow, 11).setValue(ot);
        att.getRange(foundRow, 12).setValue(monthYear);

        return { success: true, message: `Exit marked at ${nowTime12Hr}, Total hours = ${hours}` };
      } else {
        return { success: false, message: "Attendance already complete for today." };
      }
    }
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function markManualAttendance(payload) {
  try {
    const { att } = ensureSheets();
    const name = String(payload.name || "").trim().toUpperCase();
    const category = String(payload.category || "").trim().toUpperCase();
    const adhar = String(payload.adhar || "").trim();
    const mobile = String(payload.mobile || "").trim();
    const manualDate = String(payload.date || "").trim();
    const entryTime = String(payload.entryTime || "").trim();
    const exitTime = String(payload.exitTime || "").trim();

    if (!manualDate) return { success: false, error: "Date is required" };
    if (!entryTime) return { success: false, error: "Entry time is required" };

    const employeeId = _generateEmployeeId(name, adhar);

    const data = att.getDataRange().getValues();
    let foundRow = -1;
    for (let i = 1; i < data.length; i++) {
      const [date, empId, empName] = data[i];
      if (_cellToDateString(date) === manualDate && empName === name) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow === -1) {
      const monthYear = Utilities.formatDate(new Date(manualDate), Session.getScriptTimeZone(), "MMM-yyyy");
      if (exitTime) {
        const entryTime24 = _convertTo24Hr(entryTime);
        const exitTime24 = _convertTo24Hr(exitTime);
        const entryDate = new Date(manualDate + "T" + entryTime24);
        const exitDate = new Date(manualDate + "T" + exitTime24);
        let totalHours = (exitDate - entryDate) / (1000 * 60 * 60);
        if (totalHours < 0) totalHours += 24;
        totalHours = Math.round(totalHours * 100) / 100;
        const duty = Math.floor(totalHours / 8);
        const ot = Math.max(0, Math.round((totalHours - duty * 8) * 100) / 100);

        att.appendRow([manualDate, employeeId, name, category, adhar, mobile, entryTime, exitTime, totalHours, duty, ot, monthYear]);
        const range = att.getRange(att.getLastRow(), 7, 1, 2);
        range.setNumberFormat('@STRING@');
        return { success: true, message: `Manual attendance saved with entry at ${entryTime} and exit at ${exitTime}` };
      } else {
        att.appendRow([manualDate, employeeId, name, category, adhar, mobile, entryTime, "", "", "", "", monthYear]);
        const range = att.getRange(att.getLastRow(), 7);
        range.setNumberFormat('@STRING@');
        return { success: true, message: `Manual entry marked at ${entryTime}` };
      }
    } else {
      if (exitTime) {
        const exitCell = att.getRange(foundRow, 8);
        exitCell.setValue(exitTime);
        exitCell.setNumberFormat('@STRING@');

        const entryTimeRaw = att.getRange(foundRow, 7).getValue();
        const entryTime24 = _convertTo24Hr(String(entryTimeRaw));
        const exitTime24 = _convertTo24Hr(exitTime);

        const entryDate = new Date(manualDate + "T" + entryTime24);
        const exitDate = new Date(manualDate + "T" + exitTime24);
        let totalHours = (exitDate - entryDate) / (1000 * 60 * 60);
        if (totalHours < 0) totalHours += 24;
        totalHours = Math.round(totalHours * 100) / 100;

        const duty = Math.floor(totalHours / 8);
        const ot = Math.max(0, Math.round((totalHours - duty * 8) * 100) / 100);

        const monthYear = Utilities.formatDate(new Date(manualDate), Session.getScriptTimeZone(), "MMM-yyyy");
        att.getRange(foundRow, 9).setValue(totalHours);
        att.getRange(foundRow, 10).setValue(duty);
        att.getRange(foundRow, 11).setValue(ot);
        att.getRange(foundRow, 12).setValue(monthYear);

        return { success: true, message: `Manual exit marked at ${exitTime}, Total hours = ${totalHours}` };
      } else {
        return { success: false, error: "Attendance entry already exists for this date. Please provide exit time." };
      }
    }
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getTodayAttendance(name) {
  try {
    const { att } = ensureSheets();
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const data = att.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const [date, empId, empName, , , , entry, exit] = data[i];
      if (_cellToDateString(date) === today && empName === name) {
        return {
          exists: true,
          entryTime: _formatTimeTo12Hr(entry) || "",
          exitTime: _formatTimeTo12Hr(exit) || ""
        };
      }
    }
    return { exists: false, entryTime: "", exitTime: "" };
  } catch (e) {
    throw new Error("getTodayAttendance error: " + e.message);
  }
}

function getAttendanceByDate(date) {
  try {
    const { att } = ensureSheets();
    if (att.getLastRow() < 2) return [];
    const data = att.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (_cellToDateString(data[i][0]) === date) {
        result.push(_formatRowForClient(data[i]));
      }
    }
    return result;
  } catch (e) {
    throw new Error("getAttendanceByDate error: " + e.message);
  }
}

function getAttendanceByEmployee(name, fromDate, toDate) {
  try {
    const { att } = ensureSheets();
    if (att.getLastRow() < 2) return [];
    const data = att.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const rowDate = _cellToDateString(data[i][0]);
      const empName = data[i][2];
      if (empName === name && rowDate >= fromDate && rowDate <= toDate) {
        result.push(_formatRowForClient(data[i]));
      }
    }
    return result;
  } catch (e) {
    throw new Error("getAttendanceByEmployee error: " + e.message);
  }
}

function exportAttendanceToExcel(data) {
  try {
    const ss = SpreadsheetApp.create("Attendance_Export_" + new Date().getTime());
    const sheet = ss.getActiveSheet();
    sheet.appendRow(["Date","Employee ID","Employee Name","Category","Adhar Number","Mobile Number","Entry Time","Exit Time","Total Hours","Duty","O.T.","Month-Year"]);
    data.forEach(row => sheet.appendRow(row));
    return ss.getUrl();
  } catch (e) {
    throw new Error("exportAttendanceToExcel error: " + e.message);
  }
}

function exportAttendanceToPDF(data) {
  try {
    const ss = SpreadsheetApp.create("Attendance_PDF_" + new Date().getTime());
    const sheet = ss.getActiveSheet();
    sheet.appendRow(["Date","Employee ID","Employee Name","Category","Adhar Number","Mobile Number","Entry Time","Exit Time","Total Hours","Duty","O.T.","Month-Year"]);
    data.forEach(row => sheet.appendRow(row));
    const pdfUrl = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?format=pdf";
    return pdfUrl;
  } catch (e) {
    throw new Error("exportAttendanceToPDF error: " + e.message);
  }
}

/* -------------------------
   Helpers
   ------------------------- */
function _roundToNearest15Minutes(date) {
  const minutes = date.getMinutes();
  const remainder = minutes % 15;
  const roundedMinutes = remainder < 7.5 ? minutes - remainder : minutes + (15 - remainder);

  const roundedDate = new Date(date);
  roundedDate.setMinutes(roundedMinutes);
  roundedDate.setSeconds(0);
  roundedDate.setMilliseconds(0);

  return roundedDate;
}

function _convertTo24Hr(time12) {
  try {
    const timeStr = String(time12).trim();
    const match = timeStr.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)$/i);

    if (!match) {
      if (timeStr.match(/^\d{1,2}:\d{2}(:\d{2})?$/)) {
        return timeStr;
      }
      return "00:00:00";
    }

    let hours = parseInt(match[1]);
    const minutes = match[2];
    const seconds = match[3] || "00";
    const ampm = match[4].toUpperCase();

    if (ampm === "PM" && hours !== 12) {
      hours += 12;
    } else if (ampm === "AM" && hours === 12) {
      hours = 0;
    }

    return String(hours).padStart(2, '0') + ':' + minutes + ':' + seconds;
  } catch (e) {
    return "00:00:00";
  }
}

function _cellToDateString(cell) {
  if (cell instanceof Date) {
    return Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  const s = String(cell).trim();
  const m = s.match(/^(\d{4}-\d{2}-\d{2})/);
  if (m) return m[1];
  try {
    const d = new Date(s);
    if (!isNaN(d)) return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
  } catch (e) {}
  return s;
}

function _formatTimeTo12Hr(timeVal) {
  if (!timeVal || timeVal === "") return "";

  try {
    let timeStr = String(timeVal).trim();

    if (timeStr.match(/\d{1,2}:\d{2}(:\d{2})?\s*(AM|PM)/i)) {
      const match = timeStr.match(/(\d{1,2}):(\d{2})(:\d{2})?\s*(AM|PM)/i);
      if (match) {
        return match[1] + ":" + match[2] + " " + match[4].toUpperCase();
      }
      return timeStr;
    }

    if (timeStr.match(/^\d{1,2}:\d{2}(:\d{2})?$/)) {
      const parts = timeStr.split(":");
      let hours = parseInt(parts[0]);
      const minutes = parts[1];
      const ampm = hours >= 12 ? "PM" : "AM";
      hours = hours % 12;
      hours = hours ? hours : 12;
      return hours + ":" + minutes + " " + ampm;
    }

    if (timeStr.includes("GMT") || timeStr.includes("T")) {
      const d = new Date(timeStr);
      if (!isNaN(d)) {
        let hours = d.getHours();
        const minutes = String(d.getMinutes()).padStart(2, "0");
        const ampm = hours >= 12 ? "PM" : "AM";
        hours = hours % 12;
        hours = hours ? hours : 12;
        return hours + ":" + minutes + " " + ampm;
      }
    }

    return timeStr;
  } catch (e) {
    return String(timeVal);
  }
}

function _formatRowForClient(row) {
  return row.map((c, idx) => {
    if (c instanceof Date) {
      if (idx === 0) return Utilities.formatDate(c, Session.getScriptTimeZone(), "yyyy-MM-dd");
      return String(c);
    }
    if (idx === 6 || idx === 7) {
      return _formatTimeTo12Hr(c);
    }
    if (idx === 8 || idx === 10) {
      const num = parseFloat(c);
      if (!isNaN(num)) {
        return num.toFixed(2);
      }
    }
    return String(c);
  });
}

function _generateEmployeeId(name, adhar) {
  try {
    const namePart = String(name).replace(/[^A-Z]/g, '').substring(0, 4).padEnd(4, 'X');
    const adharStr = String(adhar);
    const adharPart = adharStr.substring(adharStr.length - 4);
    return namePart + adharPart;
  } catch (e) {
    return "XXXX0000";
  }
}

function verifyPassword(username, password) {
  try {
    const { users } = ensureSheets();
    if (users.getLastRow() < 2) return { success: false, error: "No users found" };

    const data = users.getRange(2, 1, users.getLastRow() - 1, 2).getValues();

    for (let i = 0; i < data.length; i++) {
      const [storedUsername, storedPassword] = data[i];
      if (String(storedUsername).trim() === String(username).trim() &&
          String(storedPassword).trim() === String(password).trim()) {
        return { success: true };
      }
    }

    return { success: false, error: "Invalid username or password" };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

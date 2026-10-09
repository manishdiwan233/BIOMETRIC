const SPREAD_SHEET_ID = "1DxTKgRXcWnWtEmuTKuETPFUr3lnRGSqvAJS_Dpaq6fI";
const SHEET_NAME = "Class 1";
const TEACHER_SHEET_NAME = "Teachers"; // NAYA TEACHER TAB

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('DDU-GKY Attendance Portal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================
// STUDENT CODE (BILKUL UNTOUCHED)
// ==========================================
function getStudentList() {
  try {
    const ss = SpreadsheetApp.openById(SPREAD_SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow < 4) return [];

    const range = sheet.getRange(4, 1, lastRow - 3, 6); 
    const data = range.getValues();
    
    return data.map(row => ({
      batch: row[0].toString(),     
      name: row[1].toString(),      
      mobilizer: row[2].toString(), 
      kp_id: row[3].toString(),     
      care_of: row[4].toString(),   
      contact: row[5].toString()    
    })).filter(s => s.name.trim() !== "");
    
  } catch (e) {
    return "Error: " + e.toString();
  }
}

function updateRegister(attendanceData, selectedDate) {
  try {
    const ss = SpreadsheetApp.openById(SPREAD_SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const fullData = sheet.getDataRange().getValues();
    
    const targetDate = new Date(selectedDate);
    targetDate.setHours(0, 0, 0, 0);
    
    const d = targetDate.getDate();
    const m = targetDate.getMonth() + 1;
    const formatMatch = d + "/" + m; 
    
    const dateRow = fullData[1]; 
    let colIndex = -1;

    for (let j = 0; j < dateRow.length; j++) {
      let cellValue = dateRow[j];
      if (cellValue instanceof Date) {
        let tempDate = new Date(cellValue);
        tempDate.setHours(0, 0, 0, 0);
        if (tempDate.getTime() === targetDate.getTime()) {
          colIndex = j + 1;
          break;
        }
      } else {
        let strVal = cellValue.toString().trim();
        if (strVal.includes(formatMatch)) {
          colIndex = j + 1;
          break;
        }
      }
    }
    
    if (colIndex === -1) {
      return "Error: Row 2 mein Date '" + formatMatch + "' ka column nahi mila!";
    }

    const namesInSheet = fullData.map(r => r[1].toString().trim()); 
    
    attendanceData.forEach(item => {
      let rowIndex = namesInSheet.indexOf(item.name.trim()) + 1;
      if (rowIndex > 0) {
        sheet.getRange(rowIndex, colIndex).setValue(item.status);
      }
    });
    
    return "Success! Attendance register update ho gaya hai.";
  } catch (e) {
    return "Error: " + e.toString();
  }
}

function getTodayStats(selectedDate) {
  try {
    const ss = SpreadsheetApp.openById(SPREAD_SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const fullData = sheet.getDataRange().getValues();
    
    const targetDate = new Date(selectedDate);
    targetDate.setHours(0, 0, 0, 0);
    const dateStr = targetDate.getDate() + "/" + (targetDate.getMonth() + 1);
    
    const dateRow = fullData[1]; 
    let colIndex = -1;

    for (let j = 0; j < dateRow.length; j++) {
      let cell = dateRow[j];
      if (cell instanceof Date) {
        cell.setHours(0, 0, 0, 0);
        if (cell.getTime() === targetDate.getTime()) { colIndex = j; break; }
      } else if (cell.toString().includes(dateStr)) {
        colIndex = j; break;
      }
    }
    
    if (colIndex === -1) return []; 

    return fullData.slice(3).map(row => ({
      name: row[1],      
      mobilizer: row[2], 
      status: row[colIndex] 
    })).filter(r => r.name !== "");
  } catch (e) {
    return [];
  }
}

function getFullReportData() {
  try {
    const ss = SpreadsheetApp.openById(SPREAD_SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const fullData = sheet.getDataRange().getValues();
    
    const datesRow = fullData[1].slice(6); 
    
    const reportRows = fullData.slice(3).map(row => {
      let attendance = row.slice(6); 
      let pCount = attendance.filter(status => status === 'P').length;
      let sCount = attendance.filter(status => status === 'Sunday' || status === 'S').length;
      let aCount = attendance.filter(status => status === 'A').length;
      let hCount = attendance.filter(status => status === 'Holiday' || status === 'H').length;
      
      return {
        batch: row[0],
        name: row[1],
        kp_id: row[3],
        mobilizer: row[2],
        attendance: attendance, 
        totalPresent: pCount,
        totalSunday: sCount,
        totalAbsent: aCount,
        totalHoliday: hCount
      };
    }).filter(r => r.name !== "");

    return {
      dates: datesRow.map(d => (d instanceof Date) ? (d.getDate() + "/" + (d.getMonth() + 1)) : d.toString()),
      students: reportRows
    };
  } catch (e) {
    return "Error: " + e.toString();
  }
}

function markBulkLeave(selectedDate, leaveType) {
  try {
    const ss = SpreadsheetApp.openById(SPREAD_SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const fullData = sheet.getDataRange().getValues();
    
    const targetDate = new Date(selectedDate);
    targetDate.setHours(0, 0, 0, 0);
    const dateStr = targetDate.getDate() + "/" + (targetDate.getMonth() + 1);
    
    const dateRow = fullData[1];
    let colIndex = -1;

    for (let j = 0; j < dateRow.length; j++) {
      let cell = dateRow[j];
      if (cell instanceof Date) {
        cell.setHours(0, 0, 0, 0);
        if (cell.getTime() === targetDate.getTime()) { colIndex = j + 1; break; }
      } else if (cell.toString().includes(dateStr)) {
        colIndex = j + 1; break;
      }
    }
    
    if (colIndex === -1) {
      return "Error: Chuni gayi tareekh sheet register me nahi mili!";
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow >= 4) {
      const totalStudents = lastRow - 3;
      const cellValues = Array(totalStudents).fill([leaveType]);
      sheet.getRange(4, colIndex, totalStudents, 1).setValues(cellValues);
    }
    
    return "Success: Sabhi candidates ke liye " + leaveType + " mark kar diya gaya hai.";
  } catch (e) {
    return "Error: " + e.toString();
  }
}

function markBulkRandomPresent(selectedDate, targetPercentage) {
  try {
    const ss = SpreadsheetApp.openById(SPREAD_SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const fullData = sheet.getDataRange().getValues();
    
    const targetDate = new Date(selectedDate);
    targetDate.setHours(0, 0, 0, 0);
    const dateStr = targetDate.getDate() + "/" + (targetDate.getMonth() + 1);
    
    const dateRow = fullData[1];
    let colIndex = -1;

    for (let j = 0; j < dateRow.length; j++) {
      let cell = dateRow[j];
      if (cell instanceof Date) {
        cell.setHours(0, 0, 0, 0);
        if (cell.getTime() === targetDate.getTime()) { colIndex = j + 1; break; }
      } else if (cell.toString().includes(dateStr)) {
        colIndex = j + 1; break;
      }
    }
    
    if (colIndex === -1) {
      return "Error: Chuni gayi tareekh sheet register me nahi mili!";
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow >= 4) {
      const totalStudents = lastRow - 3;
      const cellValues = [];
      
      for(let i = 0; i < totalStudents; i++) {
        let rand = Math.random() * 100;
        if (rand <= targetPercentage) {
          cellValues.push(["P"]);
        } else {
          cellValues.push(["A"]);
        }
      }
      sheet.getRange(4, colIndex, totalStudents, 1).setValues(cellValues);
    }
    
    return "Success: Target " + targetPercentage + "% probability ke mutabik P/A status distribute ho gaya hai.";
  } catch (e) {
    return "Error: " + e.toString();
  }
}

function sendDefaulterMail(targetEmail) {
  try {
    if (!targetEmail || !targetEmail.includes("@")) {
      return "Error: Kripya ek sahi Email ID enter karein!";
    }
    
    const reportData = getFullReportData();
    if (typeof reportData === 'string') return reportData;
    
    const totalCourseDays = 90;
    let defaulterStudents = [];
    let batchCounts = {}; 
    
    reportData.students.forEach(s => {
      let P = s.totalPresent || 0;
      let calculatedPercentage = (P / totalCourseDays) * 100;
      if (calculatedPercentage < 70) {
        defaulterStudents.push({
          name: s.name,
          batch: s.batch,
          present: P,
          percentage: calculatedPercentage.toFixed(2)
        });
        batchCounts[s.batch] = (batchCounts[s.batch] || 0) + 1;
      }
    });
    
    if (defaulterStudents.length === 0) {
      return "Alert: Koi bhi candidate 70% se kam attendance wala nahi mila. Mail nahi bheja gaya.";
    }
    
    let batchSummaryHtml = `
      <div style="background-color: #f1f3f5; padding: 12px; margin-bottom: 20px; border-radius: 6px; border-left: 4px solid #c82333;">
        <h4 style="margin-top: 0; margin-bottom: 8px; color: #343a40; font-size: 16px;">Batch-wise Defaulter Summary:</h4>
        <ul style="margin: 0; padding-left: 20px; color: #495057; font-size: 14px;">`;
        
    for (let batch in batchCounts) {
      batchSummaryHtml += `<li style="margin-bottom: 4px;"><b>${batch}:</b> ${batchCounts[batch]} Candidates</li>`;
    }
    batchSummaryHtml += `</ul></div>`;
    
    let emailBody = `
      <div style="font-family: Arial, sans-serif; padding: 20px; color: #333;">
        <h2 style="color: #c82333;">ALFRESCO SOLUTIONS LLP - Attendance Alert</h2>
        <p>Dear Rajesh Ji,</p>
        <p style="font-size: 15px; color: #444; line-height: 1.5;">
          Please find below the detailed list of candidates whose attendance has fallen below the mandatory 70% threshold during the 90-day training tenure: <strong style="color: #d9534f;">70% se kam</strong> hai:
        </p>
        ${batchSummaryHtml}
        <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; border: 1px solid #ddd;">
          <thead style="background-color: #f8d7da; color: #721c24;">
            <tr>
              <th align="left">Student Name</th>
              <th align="left">Batch Details</th>
              <th align="center">Present Count</th>
              <th align="center">Attendance Percentage</th>
            </tr>
          </thead>
          <tbody>`;
          
    defaulterStudents.forEach(st => {
      emailBody += `
            <tr>
              <td><b>${st.name}</b></td>
              <td>${st.batch}</td>
              <td align="center">${st.present} Days</td>
              <td align="center" style="color: #d9534f; font-weight: bold;">${st.percentage}%</td>
            </tr>`;
    });
    
    emailBody += `
          </tbody>
        </table>
        <p style="margin-top: 20px; font-size: 12px; color: #666;">This is an automated system generated report sent on behalf of Portal Admin.</p>
      </div>`;
      
    MailApp.sendEmail({
      to: targetEmail,
      subject: "DDU-GKY Portal Alert: Candidates with Shortage Attendance (< 70%)",
      htmlBody: emailBody
    });
    
    return "Success: Defaulters list (Total: " + defaulterStudents.length + " candidates) ka mail '" + targetEmail + "' par kamiyabi se bhej diya gaya hai.";
  } catch (e) {
    return "Error: Mail nahi bheja ja saka. Details: " + e.toString();
  }
}

function fixMailPermission() {
  MailApp.sendEmail(Session.getActiveUser().getEmail(), "Permission Fix", "This is to fix authorization error.");
}

// ==========================================
// NEW TEACHER CODE (STUDENT CODE KE SATH)
// ==========================================

function getTeacherList() {
  try {
    const ss = SpreadsheetApp.openById(SPREAD_SHEET_ID);
    const sheet = ss.getSheetByName(TEACHER_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow < 4) return [];

    // Assuming Teacher sheet follows same 6 column info structure
    // Col 1: ID, Col 2: Name, Col 3: Department, Col 4: Designation, Col 5: Email, Col 6: Contact
    const range = sheet.getRange(4, 1, lastRow - 3, 6); 
    const data = range.getValues();
    
    return data.map(row => ({
      emp_id: row[0].toString(),     
      name: row[1].toString(),      
      department: row[2].toString(), 
      contact: row[5].toString()    
    })).filter(t => t.name.trim() !== "");
    
  } catch (e) {
    return "Error: " + e.toString();
  }
}

function updateTeacherRegister(attendanceData, selectedDate) {
  try {
    const ss = SpreadsheetApp.openById(SPREAD_SHEET_ID);
    const sheet = ss.getSheetByName(TEACHER_SHEET_NAME);
    const fullData = sheet.getDataRange().getValues();
    
    const targetDate = new Date(selectedDate);
    targetDate.setHours(0, 0, 0, 0);
    const d = targetDate.getDate();
    const m = targetDate.getMonth() + 1;
    const formatMatch = d + "/" + m; 
    
    const dateRow = fullData[1]; 
    let colIndex = -1;

    for (let j = 0; j < dateRow.length; j++) {
      let cellValue = dateRow[j];
      if (cellValue instanceof Date) {
        let tempDate = new Date(cellValue);
        tempDate.setHours(0, 0, 0, 0);
        if (tempDate.getTime() === targetDate.getTime()) { colIndex = j + 1; break; }
      } else {
        if (cellValue.toString().trim().includes(formatMatch)) { colIndex = j + 1; break; }
      }
    }
    
    if (colIndex === -1) return "Error: Row 2 mein Date '" + formatMatch + "' ka column nahi mila!";

    const namesInSheet = fullData.map(r => r[1].toString().trim()); 
    attendanceData.forEach(item => {
      let rowIndex = namesInSheet.indexOf(item.name.trim()) + 1;
      if (rowIndex > 0) sheet.getRange(rowIndex, colIndex).setValue(item.status);
    });
    
    return "Success! Teacher attendance register update ho gaya hai.";
  } catch (e) {
    return "Error: " + e.toString();
  }
}

function getTeacherTodayStats(selectedDate) {
  try {
    const ss = SpreadsheetApp.openById(SPREAD_SHEET_ID);
    const sheet = ss.getSheetByName(TEACHER_SHEET_NAME);
    const fullData = sheet.getDataRange().getValues();
    
    const targetDate = new Date(selectedDate);
    targetDate.setHours(0, 0, 0, 0);
    const dateStr = targetDate.getDate() + "/" + (targetDate.getMonth() + 1);
    const dateRow = fullData[1]; 
    let colIndex = -1;

    for (let j = 0; j < dateRow.length; j++) {
      let cell = dateRow[j];
      if (cell instanceof Date) {
        cell.setHours(0, 0, 0, 0);
        if (cell.getTime() === targetDate.getTime()) { colIndex = j; break; }
      } else if (cell.toString().includes(dateStr)) {
        colIndex = j; break;
      }
    }
    
    if (colIndex === -1) return []; 

    return fullData.slice(3).map(row => ({
      name: row[1],      
      department: row[2], 
      status: row[colIndex] 
    })).filter(r => r.name !== "");
  } catch (e) {
    return [];
  }
}

function getTeacherFullReportData() {
  try {
    const ss = SpreadsheetApp.openById(SPREAD_SHEET_ID);
    const sheet = ss.getSheetByName(TEACHER_SHEET_NAME);
    
    // Agar sheet khali hai ya galat hai
    if (!sheet) return "Error: 'Teachers' sheet system mein nahi mili!";
    
    const fullData = sheet.getDataRange().getValues();
    if (fullData.length < 4) {
      return { dates: [], teachers: [] }; // No data found yet
    }
    
    // Safe slice for Dates (Index 6 yani Column G se aage)
    let datesRow = [];
    if (fullData[1] && fullData[1].length > 6) {
      datesRow = fullData[1].slice(6);
    }
    
    const reportRows = fullData.slice(3).map(row => {
      let attendance = row.length > 6 ? row.slice(6) : []; 
      
      // UPDATE: Ab yeh "P" aur "P - 10:30" dono ko count karega
      let pCount = attendance.filter(status => status && status.toString().trim().startsWith('P')).length;
      
      let sCount = attendance.filter(status => status === 'Sunday' || status === 'S').length;
      let aCount = attendance.filter(status => status === 'A').length;
      let hCount = attendance.filter(status => status === 'Holiday' || status === 'H').length;
      let lCount = attendance.filter(status => status === 'Leave' || status === 'L').length;
      
      return {
        emp_id: row[0] ? row[0].toString() : "",
        name: row[1] ? row[1].toString() : "",
        department: row[2] ? row[2].toString() : "",
        attendance: attendance.map(a => a ? a.toString() : ""), 
        totalPresent: pCount,
        totalSunday: sCount,
        totalAbsent: aCount,
        totalHoliday: hCount,
        totalLeave: lCount
      };
    }).filter(r => r.name !== "");

    // Dates ko safely string mein convert karna
    let formattedDates = datesRow.map(d => {
      if (d instanceof Date) {
        return d.getDate() + "/" + (d.getMonth() + 1);
      }
      return d ? d.toString() : "";
    });

    return {
      dates: formattedDates,
      teachers: reportRows
    };
  } catch (e) {
    return "Error: " + e.toString();
  }
}
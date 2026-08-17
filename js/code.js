// ========================================================================
// GLOBAL CONSTANTS (ตั้งค่าพื้นฐาน)
// ========================================================================
const SPREADSHEET_ID = "1OSx5X0na7rF6f1imn_T81QVvKGh0tj6_CBmM5uCYM_4";
const BOOKINGS_SHEET_NAME = "Bookings";
const ROOMS_SHEET_NAME = "Rooms";
const SCHOOL_LOGO_URL = "https://i.postimg.cc/C5RsRHgV/S-5677128-1.png";

// รายชื่อกอง
const DEPARTMENTS = [
  "สำนักปลัด",
  "กองคลัง",
  "กองสวัสดิการสังคม",
  "กองช่าง",
  "กองการศึกษา",
  "กองยุทธศาสตร์",
  "กองสาธารณสุขฯ"
];

// Telegram
const TELEGRAM_BOT_TOKEN = "8729170135:AAGuBUwJW6dLL8kUQ49VOq3eEac5XOI-kU8";
const TELEGRAM_CHAT_ID   = "-1004487379324";

// ========================================================================
// API ENTRY POINTS
// ========================================================================
function doGet(e) {
  return ContentService.createTextOutput("API is running... (No Login Required)").setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  let response = { success: false, error: 'Unknown action' };
  
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    
    // ทำงานตามคำสั่งที่ส่งมา (ถอดระบบเช็ครหัสผ่านออกแล้ว)
    if (action === 'getRooms') {
      response = getRoomData();
    } 
    else if (action === 'getEvents') {
      response = { success: true, data: getCalendarEvents() };
    } 
    else if (action === 'submitBooking') {
      response = submitBooking(params.formData);
    } 
    else if (action === 'deleteBooking') {
      response = deleteBooking(params.bookingId);
    } 
    else if (action === 'generateReport') {
      response = generateReportAsPdf(params.reportType, params.filterParam);
    }

  } catch (error) {
    response = { success: false, error: "API Error: " + error.message };
  }

  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

// ========================================================================
// DATA RETRIEVAL
// ========================================================================
function getRoomData() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ROOMS_SHEET_NAME);
    if (!sheet) return { success: false, error: `Sheet "${ROOMS_SHEET_NAME}" not found.` };
    if (sheet.getLastRow() < 2) return { success: true, data: [] };

    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    const dataValues = dataRange.getValues();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());

    const rooms = dataValues.map(row => {
      let room = {};
      headers.forEach((header, i) => room[header] = row[i]);
      return room;
    });
    return { success: true, data: rooms };
  } catch (error) {
    return { success: false, error: "เกิดข้อผิดพลาดในการดึงข้อมูลห้อง: " + error.message };
  }
}

function getBookings() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(BOOKINGS_SHEET_NAME);
    if (!sheet) return { success: false, error: `Sheet "${BOOKINGS_SHEET_NAME}" not found.` };
    if (sheet.getLastRow() < 2) return { success: true, data: [] };

    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    const dataValues = dataRange.getValues();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());

    const bookings = dataValues.map(row => {
      let booking = {};
      headers.forEach((headerName, index) => {
        const cellValue = row[index];
        if (headerName === 'BookingDate') {
          if (cellValue instanceof Date) {
            booking[headerName] = Utilities.formatDate(new Date(cellValue), Session.getScriptTimeZone(), "yyyy-MM-dd");
          } else if (cellValue && typeof cellValue === 'string' && cellValue.match(/^\d{4}-\d{2}-\d{2}/)) {
            booking[headerName] = cellValue.substring(0, 10);
          } else if (cellValue) {
            try { booking[headerName] = Utilities.formatDate(new Date(cellValue), Session.getScriptTimeZone(), "yyyy-MM-dd"); }
            catch (e) { booking[headerName] = cellValue ? String(cellValue).trim() : null; }
          } else {
            booking[headerName] = null;
          }
        } else if (headerName === 'StartTime' || headerName === 'EndTime') {
          if (cellValue instanceof Date) {
            booking[headerName] = Utilities.formatDate(new Date(cellValue), Session.getScriptTimeZone(), "HH:mm");
          } else if (cellValue && typeof cellValue === 'string' && cellValue.match(/^\d{1,2}:\d{2}/)) {
            let parts = String(cellValue).split(':');
            booking[headerName] = `${parts[0].padStart(2, '0')}:${parts[1].padStart(2, '0')}`;
          } else if (cellValue) {
            booking[headerName] = String(cellValue).trim();
          } else {
            booking[headerName] = null;
          }
        } else {
          booking[headerName] = (cellValue !== undefined && cellValue !== null) ? String(cellValue).trim() : "";
        }
      });
      return booking;
    });
    return { success: true, data: bookings };
  } catch (error) {
    return { success: false, error: "เกิดข้อผิดพลาดในการดึงข้อมูลการจอง: " + error.message };
  }
}

function getCalendarEvents() {
  try {
    const bookingsResult = getBookings();
    if (!bookingsResult.success || !bookingsResult.data) return [];

    const calendarEvents = bookingsResult.data.map(booking => {
      const bookingDate = booking.BookingDate;
      const startTime = booking.StartTime;
      const endTime = booking.EndTime;
      if (!bookingDate || !startTime || !endTime || !startTime.includes(':') || !endTime.includes(':')) return null;

      const isoStartDateTime = `${bookingDate}T${startTime}:00`;
      const isoEndDateTime = `${bookingDate}T${endTime}:00`;
      if (isNaN(new Date(isoStartDateTime).getTime()) || isNaN(new Date(isoEndDateTime).getTime())) return null;

      let eventTitle = (booking.Purpose || "การจอง") + (booking.RoomName ? ` (${booking.RoomName})` : "") + (booking.Name ? ` - ${booking.Name}` : "");
      return {
        id: booking.BookingID || Utilities.getUuid(),
        title: eventTitle,
        start: isoStartDateTime,
        end: isoEndDateTime,
        backgroundColor: booking.EventColor || '#3788d8',
        borderColor: booking.EventColor || '#3788d8',
        extendedProps: {
          roomName: booking.RoomName || 'N/A',
          purpose: booking.Purpose || 'N/A',
          name: booking.Name || 'N/A',
          phoneNumber: booking.PhoneNumber || '',
          attendees: booking.Attendees || '',
          notes: booking.Notes || '',
          targetAudience: booking.TargetAudience || '',
          status: booking.Status || 'N/A',
          bookingID: booking.BookingID || '',
          department: booking.Department || ''
        }
      };
    }).filter(event => event !== null);

    return calendarEvents;
  } catch (error) {
    return [];
  }
}

// ========================================================================
// SUBMIT BOOKING & DELETE & TELEGRAM
// ========================================================================
function submitBooking(formData) {
  try {
    if (!formData || !formData.nameInput || !formData.PhoneInput || !formData.roomSelect ||
        !formData.dateInput || !formData.startTimeInput || !formData.endTimeInput ||
        !formData.purposeInput || !formData.attendeesInput || !formData.departmentSelect) {
      return { success: false, error: "ข้อมูลไม่ครบถ้วน กรุณากรอกข้อมูลที่จำเป็นทั้งหมด" };
    }

    if (!DEPARTMENTS.includes(formData.departmentSelect)) {
      return { success: false, error: "กองที่เลือกไม่ถูกต้อง" };
    }

    const bookingDateStr = formData.dateInput;
    const startTimeStr = formData.startTimeInput;
    const endTimeStr = formData.endTimeInput;

    const startDateTime = new Date(`${bookingDateStr}T${startTimeStr}:00`);
    const endDateTime = new Date(`${bookingDateStr}T${endTimeStr}:00`);
    if (endDateTime <= startDateTime) {
      return { success: false, error: "เวลาสิ้นสุดต้องอยู่หลังเวลาเริ่มต้น" };
    }

    const bookingsSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(BOOKINGS_SHEET_NAME);
    
    // ตรวจสอบการจองซ้ำ
    const existingBookingsData = getBookings();
    if (existingBookingsData.success && existingBookingsData.data) {
      for (const existing of existingBookingsData.data) {
        if (existing.RoomName === formData.roomSelect && existing.BookingDate === bookingDateStr) {
          if (!existing.StartTime || !existing.EndTime) continue;
          const existingStart = new Date(`${existing.BookingDate}T${existing.StartTime}:00`);
          const existingEnd = new Date(`${existing.BookingDate}T${existing.EndTime}:00`);
          if (isNaN(existingStart.getTime()) || isNaN(existingEnd.getTime())) continue;

          if (startDateTime < existingEnd && endDateTime > existingStart) {
            return { success: false, error: `ห้อง "${formData.roomSelect}" ถูกจองแล้วในช่วงเวลา ${existing.StartTime} - ${existing.EndTime}` };
          }
        }
      }
    }

    // หาสีของห้อง
    const roomsData = getRoomData();
    let eventColor = "#808080";
    if (roomsData.success && roomsData.data) {
      const roomInfo = roomsData.data.find(r => r.RoomName === formData.roomSelect);
      if (roomInfo && roomInfo.ColorHex) eventColor = String(roomInfo.ColorHex);
    }

    const newRowData = {
      Timestamp: new Date(),
      BookingID: Utilities.getUuid(),
      Name: formData.nameInput,
      Department: formData.departmentSelect,
      PhoneNumber: formData.PhoneInput,
      RoomName: formData.roomSelect,
      Purpose: formData.purposeInput,
      TargetAudience: formData.targetInput || "",
      Attendees: formData.attendeesInput,
      BookingDate: bookingDateStr,
      StartTime: startTimeStr,
      EndTime: endTimeStr,
      Notes: formData.notesInput || "",
      EventColor: eventColor,
      Status: "Confirmed"
    };

    const sheetHeaders = bookingsSheet.getRange(1, 1, 1, bookingsSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const newRowOrdered = sheetHeaders.map(header => newRowData[header] !== undefined ? newRowData[header] : "");
    bookingsSheet.appendRow(newRowOrdered);

    sendTelegramNotification(newRowData);
    return { success: true, message: "บันทึกการจองสำเร็จ!" };
  } catch (error) {
    return { success: false, error: "เกิดข้อผิดพลาดในการบันทึก: " + error.message };
  }
}

function deleteBooking(bookingId) {
  if (!bookingId) return { success: false, error: "ไม่ได้รับ ID การจองที่ต้องการลบ" };
  
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(BOOKINGS_SHEET_NAME);
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(h => String(h).trim());
    const bookingIdColumnIndex = headers.indexOf('BookingID');

    for (let i = values.length - 1; i >= 1; i--) {
      if (values[i][bookingIdColumnIndex] === bookingId) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "การจองถูกลบเรียบร้อยแล้ว" };
      }
    }
    return { success: false, error: "ไม่พบการจองด้วย ID ที่ระบุในชีต" };
  } catch (error) {
    return { success: false, error: "เกิดข้อผิดพลาดระหว่างการลบการจอง: " + error.message };
  }
}

function sendTelegramNotification(booking) {
  try {
    const message = `📢 <b>มีการจองห้องประชุมใหม่</b>
🏢 <b>กอง:</b> ${booking.Department || "-"}
👤 <b>ผู้จอง:</b> ${booking.Name || "-"}
📞 <b>เบอร์โทร:</b> ${booking.PhoneNumber || "-"}
🏠 <b>ห้อง:</b> ${booking.RoomName || "-"}
📅 <b>วันที่:</b> ${booking.BookingDate || "-"}
⏰ <b>เวลา:</b> ${booking.StartTime || "--:--"} - ${booking.EndTime || "--:--"}
🎯 <b>วัตถุประสงค์:</b> ${booking.Purpose || "-"}
👥 <b>จำนวนผู้เข้าร่วม:</b> ${booking.Attendees || "-"}
📝 <b>หมายเหตุ:</b> ${booking.Notes || "-"}`;

    const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;
    UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ chat_id: TELEGRAM_CHAT_ID, text: message, parse_mode: "HTML" }),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.error("Error sending Telegram:", e.message);
  }
}

// ========================================================================
// PDF REPORT GENERATOR
// ========================================================================
function generateReportAsPdf(reportType, filterParam) {
  try {
    const bookingsDataResult = getBookings();
    if (!bookingsDataResult.success || !bookingsDataResult.data) {
      return { success: false, error: "รายงาน: ไม่สามารถดึงข้อมูลการจองได้" };
    }

    let allBookings = bookingsDataResult.data;
    let filteredBookings = [];
    let reportPeriodText = "";

    const tableThaiHeaders = [
      'วันที่จอง', 'ช่วงเวลา', 'ชื่อห้องประชุม', 'กอง', 'ชื่อผู้จอง',
      'วัตถุประสงค์การใช้งาน', 'จำนวนผู้เข้าร่วม', 'หมายเลขโทรศัพท์', 'สถานะการจอง'
    ];

    const timeZone = Session.getScriptTimeZone();

    // กรองข้อมูลตามประเภทรายงาน
    if (reportType === 'daily') {
      filteredBookings = allBookings.filter(b => b.BookingDate === filterParam);
      try {
        reportPeriodText = "ประจำวันที่ " + Utilities.formatDate(new Date(filterParam + "T00:00:00Z"), timeZone, "d MMMM yyyy");
      } catch (e) { reportPeriodText = "ประจำวันที่ " + filterParam; }

    } else if (reportType === 'monthly') {
      filteredBookings = allBookings.filter(b => b.BookingDate && String(b.BookingDate).startsWith(filterParam));
      const [year, month] = filterParam.split('-');
      const monthDate = new Date(parseInt(year), parseInt(month) - 1, 1);
      reportPeriodText = "ประจำเดือน " + Utilities.formatDate(monthDate, timeZone, "MMMM yyyy");

    } else if (reportType === 'yearly') {
      filteredBookings = allBookings.filter(b => b.BookingDate && String(b.BookingDate).startsWith(filterParam));
      reportPeriodText = "ประจำปี พ.ศ. " + (parseInt(filterParam) + 543);

    } else if (reportType === 'department') {
      filteredBookings = allBookings.filter(b => b.Department === filterParam);
      reportPeriodText = "กอง: " + filterParam;
    } else {
      return { success: false, error: "ประเภทรายงานไม่ถูกต้อง" };
    }

    if (filteredBookings.length === 0) {
      return { success: false, error: `ไม่พบข้อมูลการจองสำหรับ ${reportPeriodText}` };
    }

    const mainReportTitleText = "รายงานจองห้องประชุมเทศบาลตำบลบึงโขงหลง";
    const docNameForFile = `รายงานการจองห้องประชุม-${reportType}-${String(filterParam).replace(/-/g, '').replace(/\s+/g, '_')}`;
    const tempDoc = DocumentApp.create(docNameForFile + " (Temp)");
    const body = tempDoc.getBody();
    const headerSection = tempDoc.getHeader() || tempDoc.addHeader();

    body.clear();
    headerSection.clear();
    body.setMarginTop(36).setMarginBottom(72).setMarginLeft(50).setMarginRight(50);

    let commonStyleAttrs = {};
    commonStyleAttrs[DocumentApp.Attribute.FONT_FAMILY] = 'Sarabun';
    headerSection.setAttributes(commonStyleAttrs);

    let bodyStyleAttrs = {};
    bodyStyleAttrs[DocumentApp.Attribute.FONT_FAMILY] = 'Sarabun';
    bodyStyleAttrs[DocumentApp.Attribute.FONT_SIZE] = 10;
    body.setAttributes(bodyStyleAttrs);

    if (SCHOOL_LOGO_URL) {
      try {
        const imageBlob = UrlFetchApp.fetch(SCHOOL_LOGO_URL).getBlob();
        const imageParaInHeader = headerSection.appendParagraph('');
        const imageInHeader = imageParaInHeader.appendInlineImage(imageBlob);
        const desiredLogoWidth = 60;
        const aspectRatio = imageInHeader.getWidth() / imageInHeader.getHeight();
        imageInHeader.setWidth(desiredLogoWidth);
        imageInHeader.setHeight(desiredLogoWidth / aspectRatio);
        imageParaInHeader.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        imageParaInHeader.setSpacingAfter(0);
      } catch (imgError) {
        headerSection.appendParagraph("[ตราโลโก้]").setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(9).setItalic(true);
      }
    }

    const mainTitleParaInHeader = headerSection.appendParagraph(mainReportTitleText);
    mainTitleParaInHeader.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    let mainTitleStyle = {};
    mainTitleStyle[DocumentApp.Attribute.FONT_SIZE] = 16;
    mainTitleStyle[DocumentApp.Attribute.BOLD] = true;
    mainTitleParaInHeader.setAttributes(mainTitleStyle);
    mainTitleParaInHeader.setSpacingBefore(6).setSpacingAfter(0);

    const reportPeriodParaInHeader = headerSection.appendParagraph(reportPeriodText);
    reportPeriodParaInHeader.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    let reportPeriodStyle = {};
    reportPeriodStyle[DocumentApp.Attribute.FONT_SIZE] = 14;
    reportPeriodParaInHeader.setAttributes(reportPeriodStyle);
    reportPeriodParaInHeader.setSpacingBefore(4);

    filteredBookings.sort((a, b) => {
      const dateTimeA = new Date(`${a.BookingDate || '1970-01-01'}T${a.StartTime || '00:00'}`);
      const dateTimeB = new Date(`${b.BookingDate || '1970-01-01'}T${b.StartTime || '00:00'}`);
      if (dateTimeA.getTime() === dateTimeB.getTime()) {
        return String(a.RoomName || "").localeCompare(String(b.RoomName || ""), 'th');
      }
      return dateTimeA < dateTimeB ? -1 : 1;
    });

    const tableData = [tableThaiHeaders];

    filteredBookings.forEach(b => {
      let phoneNumberStr = String(b.PhoneNumber || "").trim();
      if (phoneNumberStr.length === 9 && phoneNumberStr.match(/^[1-9]\d{8}$/)) {
        phoneNumberStr = '0' + phoneNumberStr;
      } else if (phoneNumberStr.length === 8 && phoneNumberStr.match(/^[1-9]\d{7}$/)) {
        phoneNumberStr = '0' + phoneNumberStr;
      }

      let formattedBookingDate = b.BookingDate;
      try {
        if (b.BookingDate) {
          const dateObj = new Date(b.BookingDate + "T00:00:00Z");
          formattedBookingDate = Utilities.formatDate(dateObj, timeZone, "d MMM yy");
        }
      } catch (dateErr) {}

      tableData.push([
        formattedBookingDate,
        `${b.StartTime || "--:--"} - ${b.EndTime || "--:--"}`,
        b.RoomName || "ไม่มีข้อมูล",
        b.Department || "-",
        b.Name || "ไม่มีข้อมูล",
        b.Purpose || "ไม่มีข้อมูล",
        String(b.Attendees || "-"),
        phoneNumberStr || "-",
        b.Status || "Confirmed"
      ]);
    });

    const table = body.appendTable(tableData);
    table.setBorderColor("#A9A9A9");

    const headerRowStyle = {};
    headerRowStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = "#E0E0E0";
    headerRowStyle[DocumentApp.Attribute.BOLD] = true;
    headerRowStyle[DocumentApp.Attribute.FONT_SIZE] = 9;
    headerRowStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Sarabun';

    const headerRow = table.getRow(0);
    for (let i = 0; i < headerRow.getNumCells(); i++) {
      const cell = headerRow.getCell(i);
      cell.setAttributes(headerRowStyle);
      cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      cell.setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(4).setPaddingRight(4);
    }

    const dataCellBaseStyle = {};
    dataCellBaseStyle[DocumentApp.Attribute.FONT_SIZE] = 8;
    dataCellBaseStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Sarabun';

    for (let r = 1; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      const rowBgColor = (r % 2 === 0) ? "#FFFFFF" : "#F5F5F5";

      for (let c = 0; c < row.getNumCells(); c++) {
        const cell = row.getCell(c);
        let currentCellStyles = JSON.parse(JSON.stringify(dataCellBaseStyle));
        currentCellStyles[DocumentApp.Attribute.BACKGROUND_COLOR] = rowBgColor;
        cell.setAttributes(currentCellStyles);
        cell.setVerticalAlignment(DocumentApp.VerticalAlignment.TOP);

        const cellParagraph = cell.getChild(0).asParagraph();
        if (c === 0 || c === 1 || c === 6 || c === 8) {
          cellParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        } else {
          cellParagraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
        }
        cell.setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(4).setPaddingRight(4);
      }
    }

    tempDoc.saveAndClose();

    const pdfBlob = tempDoc.getAs(MimeType.PDF).setName(docNameForFile + ".pdf");
    const pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());
    
    DriveApp.getFileById(tempDoc.getId()).setTrashed(true);

    return { success: true, pdfBase64: pdfBase64, fileName: pdfBlob.getName() };
  } catch (e) {
    return { success: false, error: "เกิดข้อผิดพลาดในการสร้าง PDF: " + e.toString() };
  }
}

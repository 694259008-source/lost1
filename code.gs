/*
 * NPRU DELIVERY - GOOGLE APPS SCRIPT BACKEND
 * Full Featured Version with Admin Systems & Content Moderation
 */

const CONFIG = {
  SPREADSHEET_ID: '18Wa0_uum5oNDNuIWy2ro-PPF40-8U4uZS1CqpekGkyM',
  DRIVE_FOLDER_ID: '1_d63TWt4z86dxDR1tnx9ZYWuKzrbxAzB',
  LINE_ADMIN_USER_ID: 'Udc34599d21cc1500ce01a927543ae6a4',
  SHEETS: {
    USERS: 'Users',
    REPORTS: 'Reports',
    CATEGORIES: 'Categories',
    LOCATIONS: 'Locations',
    NOTIFICATIONS: 'Notifications',
    AUDIT_LOGS: 'AuditLogs',
    CONSENTS: 'Consents'
  },
  SESSION_SECONDS: 21600
};

const SHEET_HEADERS = {
  Users: ['userId','fullName','studentId','email','phone','password','role','status','createdAt','updatedAt'],
  Reports: ['reportId','type','itemName','categoryId','description','distinguishingMark','date','time','locationId','imageUrl','reporterId','status','createdAt','updatedAt'],
  Categories: ['categoryId','categoryName','status','createdAt'],
  Locations: ['locationId','locationName','description','status','createdAt'],
  Notifications: ['notificationId','userId','title','message','type','isRead','createdAt'],
  AuditLogs: ['logId','userId','action','target','targetId','description','createdAt'],
  Consents: ['consentId','userId','fullName','email','privacyNoticeAccepted','consentAccepted','consentDate','privacyNoticeVersion','createdAt']
};

const PRIVACY_CONFIG = {
  VERSION: '1.0'
};

function doGet() {
  return HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setTitle('NPRU Lost&Found')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getSheet(sheetName) {
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('ไม่พบ Sheet: ' + sheetName);
  }
  return sheet;
}

function generateId(prefix) {
  return prefix + '-' +
    new Date().getTime().toString(36) + '-' +
    Math.random().toString(36).substring(2, 8);
}

function safeDate(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm:ss'
    );
  }
  return value;
}

function ensureSheet(sheetName, headers) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold');
  }

  return sheet;
}

function initializeDatabase() {
  try {
    Object.keys(SHEET_HEADERS).forEach(function(name) {
      ensureSheet(name, SHEET_HEADERS[name]);
    });

    ensureDefaultMasterData();

    return {
      success: true,
      message: 'Initialize Database สำเร็จ'
    };
  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}

function getAllSheetData(sheetName) {
  try {
    const sheet = getSheet(sheetName);
    const values = sheet.getDataRange().getValues();

    if (values.length <= 1) {
      return { success: true, data: [] };
    }

    const headers = values[0];
    const data = values.slice(1)
      .filter(function(row) {
        return row[0] !== '';
      })
      .map(function(row) {
        const obj = {};
        headers.forEach(function(header, i) {
          if (header) {
            obj[header] = safeDate(row[i]);
          }
        });
        return obj;
      });

    return {
      success: true,
      data: data
    };
  } catch (error) {
    return {
      success: false,
      message: error.message,
      data: []
    };
  }
}

/* =========================
   PUBLIC MASTER DATA
========================= */

function getDefaultCategories_() {
  return [
    ['CAT001', 'กระเป๋า', 'ACTIVE'],
    ['CAT002', 'กระเป๋าสตางค์', 'ACTIVE'],
    ['CAT003', 'โทรศัพท์มือถือ', 'ACTIVE'],
    ['CAT004', 'แท็บเล็ต / iPad', 'ACTIVE'],
    ['CAT005', 'โน้ตบุ๊ก / คอมพิวเตอร์', 'ACTIVE'],
    ['CAT006', 'หูฟัง / ลำโพง / อุปกรณ์เสียง', 'ACTIVE'],
    ['CAT007', 'สายชาร์จ / หัวชาร์จ / อะแดปเตอร์', 'ACTIVE'],
    ['CAT008', 'พาวเวอร์แบงก์', 'ACTIVE'],
    ['CAT009', 'บัตรนักศึกษา / บัตรประจำตัว', 'ACTIVE'],
    ['CAT010', 'เอกสาร', 'ACTIVE'],
    ['CAT011', 'หนังสือ / สมุด / ชีทเรียน', 'ACTIVE'],
    ['CAT012', 'เครื่องเขียน', 'ACTIVE'],
    ['CAT013', 'กุญแจ / รีโมต / คีย์การ์ด', 'ACTIVE'],
    ['CAT014', 'แว่นตา', 'ACTIVE'],
    ['CAT015', 'นาฬิกา', 'ACTIVE'],
    ['CAT016', 'เครื่องประดับ', 'ACTIVE'],
    ['CAT017', 'เสื้อผ้า', 'ACTIVE'],
    ['CAT018', 'รองเท้า', 'ACTIVE'],
    ['CAT019', 'ร่ม / หมวก / อุปกรณ์กันฝน', 'ACTIVE'],
    ['CAT020', 'ขวดน้ำ / แก้วน้ำ', 'ACTIVE'],
    ['CAT021', 'อุปกรณ์กีฬา', 'ACTIVE'],
    ['CAT022', 'หมวกกันน็อก / อุปกรณ์ยานพาหนะ', 'ACTIVE'],
    ['CAT023', 'ของใช้ส่วนตัว', 'ACTIVE'],
    ['CAT024', 'อุปกรณ์การเรียน / อุปกรณ์สำนักงาน', 'ACTIVE'],
    ['CAT025', 'เครื่องสำอาง / ของใช้ส่วนตัวขนาดเล็ก', 'ACTIVE'],
    ['CAT026', 'อื่น ๆ (ระบุเอง)', 'ACTIVE']
  ];
}

function getDefaultLocations_() {
  return [
    ['LOC001', 'อาคารคณะครุศาสตร์', 'อาคารคณะครุศาสตร์ / พื้นที่การเรียนการสอน', 'ACTIVE'],
    ['LOC002', 'อาคารสิริวรัญญา', 'อาคารสิริวรัญญา', 'ACTIVE'],
    ['LOC003', 'อาคารศูนย์ศึกษาและพัฒนา จังหวัดนครปฐม (CCS)', 'อาคารศูนย์ศึกษาและพัฒนา จังหวัดนครปฐม', 'ACTIVE'],
    ['LOC004', 'อาคารหอพักอาเขียน', 'อาคารหอพักอาเขียน', 'ACTIVE'],
    ['LOC005', 'อาคารเฉลิมพระเกียรติ 50 พรรษามหาวชิราลงกรณ', 'อาคารเฉลิมพระเกียรติ 50 พรรษามหาวชิราลงกรณ', 'ACTIVE'],
    ['LOC006', 'อาคาร A1', 'อาคารเรียน A1', 'ACTIVE'],
    ['LOC007', 'อาคาร A2', 'อาคารเรียน A2', 'ACTIVE'],
    ['LOC008', 'อาคาร A3', 'อาคารเรียน A3', 'ACTIVE'],
    ['LOC009', 'อาคาร A4', 'อาคารเรียน A4', 'ACTIVE'],
    ['LOC010', 'อาคาร A5 / คณะครุศาสตร์เก่า', 'อาคารคณะครุศาสตร์เดิม', 'ACTIVE'],
    ['LOC011', 'อาคาร A6', 'อาคารเรียน A6', 'ACTIVE'],
    ['LOC012', 'อาคาร A7', 'อาคารเรียน A7', 'ACTIVE'],
    ['LOC013', 'อาคารศูนย์ภาษาและศูนย์คอมพิวเตอร์ (LI)', 'ศูนย์ภาษาและศูนย์คอมพิวเตอร์; มีหอประชุมปิ่นเกลียวในอาคาร', 'ACTIVE'],
    ['LOC014', 'สำนักวิทยบริการ / อาคารบรรณราชนครินทร์', 'อาคารห้องสมุดและสำนักวิทยบริการ', 'ACTIVE'],
    ['LOC015', 'อาคารศูนย์วิทยาศาสตร์และเทคโนโลยี', 'พื้นที่ศูนย์วิทยาศาสตร์และเทคโนโลยี', 'ACTIVE'],
    ['LOC016', 'อาคารคอมพิวเตอร์ / ศูนย์ปฏิบัติการคอมพิวเตอร์', 'อาคารคอมพิวเตอร์และห้องปฏิบัติการคอมพิวเตอร์', 'ACTIVE'],
    ['LOC017', 'อาคารกิจกรรมนักศึกษา', 'อาคารกองพัฒนานักศึกษา / กิจกรรมนักศึกษา', 'ACTIVE'],
    ['LOC018', 'อาคารปฏิบัติการการจัดการอาหาร', 'อาคารปฏิบัติการด้านการจัดการอาหาร', 'ACTIVE'],
    ['LOC019', 'อาคารเทคโนโลยีการอาหาร', 'อาคารเทคโนโลยีการอาหาร', 'ACTIVE'],
    ['LOC020', 'อาคารปฏิบัติการบัญชีและโลจิสติกส์', 'อาคารปฏิบัติการด้านบัญชีและโลจิสติกส์', 'ACTIVE'],
    ['LOC021', 'อาคารวิศวกรรมโยธา', 'อาคารเรียน / ปฏิบัติการวิศวกรรมโยธา', 'ACTIVE'],
    ['LOC022', 'อาคาร ETB', 'อาคาร ETB / วิศวกรรมโยธา', 'ACTIVE'],
    ['LOC023', 'อาคารร้อยปีฝึกหัดครูไทย', 'อาคารร้อยปีฝึกหัดครูไทย', 'ACTIVE'],
    ['LOC024', 'อาคารทวารวดี', 'อาคารทวารวดี', 'ACTIVE'],
    ['LOC025', 'คณะมนุษยศาสตร์และสังคมศาสตร์', 'พื้นที่คณะมนุษยศาสตร์และสังคมศาสตร์', 'ACTIVE'],
    ['LOC026', 'คณะพยาบาลศาสตร์ / สำนักงานคณะพยาบาลศาสตร์', 'พื้นที่คณะพยาบาลศาสตร์', 'ACTIVE'],
    ['LOC027', 'อาคารศูนย์การจัดการ / โลจิสติกส์', 'อาคารศูนย์การจัดการ / โลจิสติกส์', 'ACTIVE'],
    ['LOC028', 'อาคารโรงเรียนสาธิต', 'อาคารโรงเรียนสาธิตมหาวิทยาลัยราชภัฏนครปฐม', 'ACTIVE'],
    ['LOC029', 'ศูนย์สาธิตการศึกษาปฐมวัย', 'ศูนย์สาธิตการศึกษาปฐมวัย', 'ACTIVE'],
    ['LOC030', 'โรงอาหาร / โรงกรองน้ำ', 'โรงอาหารและบริเวณโรงกรองน้ำ', 'ACTIVE'],
    ['LOC031', 'Sport Complex / โรงยิมพลศึกษา', 'พื้นที่กีฬาและโรงยิม', 'ACTIVE'],
    ['LOC032', 'สระว่ายน้ำ', 'สระว่ายน้ำของมหาวิทยาลัย', 'ACTIVE'],
    ['LOC033', 'อาคารศูนย์ศิลปวัฒนธรรม', 'พื้นที่ศูนย์ศิลปวัฒนธรรม', 'ACTIVE'],
    ['LOC034', 'อาคารกิจการนักศึกษา', 'อาคารกิจการนักศึกษา', 'ACTIVE'],
    ['LOC035', 'อาคาร UBI', 'อาคาร UBI', 'ACTIVE'],
    ['LOC036', 'อาคารโรงซ่อมครุภัณฑ์', 'งานอาคารสถานที่และภูมิสถาปัตย์ / โรงซ่อมครุภัณฑ์', 'ACTIVE'],
    ['LOC037', 'งานอาคารและภูมิทัศน์', 'พื้นที่งานอาคารและภูมิทัศน์', 'ACTIVE'],
    ['LOC038', 'งานยานพาหนะ', 'พื้นที่งานยานพาหนะ / โรงจอดยานพาหนะ', 'ACTIVE'],
    ['LOC039', 'สมาคมศิษย์เก่า', 'พื้นที่สมาคมศิษย์เก่า', 'ACTIVE'],
    ['LOC040', 'ลานกิจกรรม / ลานอเนกประสงค์', 'ลานกิจกรรมและพื้นที่อเนกประสงค์', 'ACTIVE'],
    ['LOC041', 'สนามกีฬา / สนามฟุตบอล', 'บริเวณสนามกีฬาและสนามฟุตบอล', 'ACTIVE'],
    ['LOC042', 'หอประชุม / ห้องประชุม', 'หอประชุมหรือพื้นที่จัดกิจกรรม', 'ACTIVE'],
    ['LOC043', 'หอพักนักศึกษา', 'บริเวณหอพักนักศึกษา', 'ACTIVE'],
    ['LOC044', 'ประตูทางเข้า-ออก / ลานจอดรถ', 'บริเวณทางเข้าออกและพื้นที่จอดรถ', 'ACTIVE'],
    ['LOC045', 'บริเวณรอบมหาวิทยาลัย', 'พื้นที่ภายนอกอาคารภายในมหาวิทยาลัย', 'ACTIVE'],
    ['LOC046', 'อื่น ๆ (ระบุเอง)', 'สถานที่อื่นที่ไม่มีในรายการ สามารถระบุเพิ่มเติมได้', 'ACTIVE']
  ];
}

function ensureDefaultMasterData() {
  const categoriesSheet = ensureSheet(CONFIG.SHEETS.CATEGORIES, SHEET_HEADERS.Categories);
  const locationsSheet = ensureSheet(CONFIG.SHEETS.LOCATIONS, SHEET_HEADERS.Locations);

  const categoryValues = categoriesSheet.getDataRange().getValues();
  const categoryNames = {};
  const categoryIds = {};
  for (let i = 1; i < categoryValues.length; i++) {
    categoryIds[String(categoryValues[i][0]).trim()] = true;
    categoryNames[String(categoryValues[i][1] || '').trim().toLowerCase()] = true;
  }

  getDefaultCategories_().forEach(function(item) {
    if (!categoryIds[item[0]] && !categoryNames[item[1].toLowerCase()]) {
      categoriesSheet.appendRow([item[0], item[1], item[2], new Date()]);
    }
  });

  const locationValues = locationsSheet.getDataRange().getValues();
  const locationNames = {};
  const locationIds = {};
  for (let i = 1; i < locationValues.length; i++) {
    locationIds[String(locationValues[i][0]).trim()] = true;
    locationNames[String(locationValues[i][1] || '').trim().toLowerCase()] = true;
  }

  getDefaultLocations_().forEach(function(item) {
    if (!locationIds[item[0]] && !locationNames[item[1].toLowerCase()]) {
      locationsSheet.appendRow([item[0], item[1], item[2], item[3], new Date()]);
    }
  });
}

function getCategories() {
  try {
    ensureDefaultMasterData();
    const result = getAllSheetData(CONFIG.SHEETS.CATEGORIES);
    if (!result.success) return result;

    return {
      success: true,
      categories: (result.data || []).filter(item => String(item.status || 'ACTIVE').toUpperCase() === 'ACTIVE')
    };
  } catch (error) {
    return { success: false, message: error.message, categories: [] };
  }
}

function getLocations() {
  try {
    ensureDefaultMasterData();
    const result = getAllSheetData(CONFIG.SHEETS.LOCATIONS);
    if (!result.success) return result;

    return {
      success: true,
      locations: (result.data || []).filter(item => String(item.status || 'ACTIVE').toUpperCase() === 'ACTIVE')
    };
  } catch (error) {
    return { success: false, message: error.message, locations: [] };
  }
}

function findOrCreateCategory_(name) {
  const clean = String(name || '').trim();
  if (!clean) throw new Error('กรุณาระบุหมวดหมู่เพิ่มเติม');

  const sheet = ensureSheet(CONFIG.SHEETS.CATEGORIES, SHEET_HEADERS.Categories);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][1] || '').trim().toLowerCase() === clean.toLowerCase()) {
      return String(values[i][0]);
    }
  }

  const id = generateId('CAT');
  sheet.appendRow([id, clean, 'ACTIVE', new Date()]);
  return id;
}

function findOrCreateLocation_(name) {
  const clean = String(name || '').trim();
  if (!clean) throw new Error('กรุณาระบุสถานที่เพิ่มเติม');

  const sheet = ensureSheet(CONFIG.SHEETS.LOCATIONS, SHEET_HEADERS.Locations);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][1] || '').trim().toLowerCase() === clean.toLowerCase()) {
      return String(values[i][0]);
    }
  }

  const id = generateId('LOC');
  sheet.appendRow([id, clean, 'สถานที่ที่ผู้ใช้งานระบุเพิ่มเติม', 'ACTIVE', new Date()]);
  return id;
}

/* =========================
   REGISTER + USER & SESSION
========================= */

function registerUser(userData) {
  try {
    if (!userData) throw new Error('ไม่พบข้อมูลสมัครสมาชิก');

    const fullName = String(userData.fullName || '').trim();
    const studentId = String(userData.studentId || '').trim();
    const email = String(userData.email || '').trim().toLowerCase();
    const phone = String(userData.phone || '').trim();
    const password = String(userData.password || '');
    const consent = userData.consent || {};

    if (!fullName || !studentId || !email || !phone || !password) {
      throw new Error('กรุณากรอกข้อมูลให้ครบทุกช่อง');
    }

    if (consent.privacyNoticeAccepted !== true || consent.consentAccepted !== true) {
      throw new Error('กรุณายินยอมให้เก็บและใช้ข้อมูลตาม Privacy Notice ก่อนสมัครสมาชิก');
    }

    const usersSheet = ensureSheet(CONFIG.SHEETS.USERS, SHEET_HEADERS.Users);
    const users = usersSheet.getDataRange().getValues();

    for (let i = 1; i < users.length; i++) {
      if (String(users[i][3] || '').trim().toLowerCase() === email) throw new Error('อีเมลนี้ถูกสมัครสมาชิกแล้ว');
      if (String(users[i][2] || '').trim() === studentId) throw new Error('รหัสนักศึกษานี้ถูกสมัครสมาชิกแล้ว');
      if (String(users[i][4] || '').trim() === phone) throw new Error('เบอร์โทรศัพท์นี้ถูกใช้งานแล้ว');
    }

    const userId = generateId('USER');
    const now = new Date();

    usersSheet.appendRow([userId, fullName, studentId, email, phone, password, 'USER', 'ACTIVE', now, now]);

    const consentSheet = ensureSheet(CONFIG.SHEETS.CONSENTS, SHEET_HEADERS.Consents);
    consentSheet.appendRow([generateId('CONSENT'), userId, fullName, email, true, true, now, PRIVACY_CONFIG.VERSION, now]);

    return {
      success: true,
      message: 'สมัครสมาชิกสำเร็จ',
      user: { userId, fullName, studentId, email, phone, role: 'USER', status: 'ACTIVE' }
    };
  } catch (error) {
    return { success: false, message: error.message || 'ไม่สามารถสมัครสมาชิกได้' };
  }
}

function loginUser(email, password) {
  try {
    email = String(email || '').trim().toLowerCase();
    password = String(password || '');

    if (!email || !password) throw new Error('กรุณากรอกอีเมลและรหัสผ่าน');

    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) throw new Error('ยังไม่มีข้อมูลผู้ใช้งานในระบบ');

    const headers = values[0];
    let foundUser = null;

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][headers.indexOf('email')]).trim().toLowerCase() === email) {
        foundUser = {
          userId: values[i][headers.indexOf('userId')],
          fullName: values[i][headers.indexOf('fullName')],
          studentId: values[i][headers.indexOf('studentId')],
          email: values[i][headers.indexOf('email')],
          phone: values[i][headers.indexOf('phone')],
          password: values[i][headers.indexOf('password')],
          role: values[i][headers.indexOf('role')],
          status: values[i][headers.indexOf('status')]
        };
        break;
      }
    }

    if (!foundUser) throw new Error('ไม่พบอีเมลนี้ในระบบ');
    if (String(foundUser.password) !== password) throw new Error('รหัสผ่านไม่ถูกต้อง');
    if (String(foundUser.status).toUpperCase() !== 'ACTIVE') throw new Error('บัญชีนี้ถูกระงับการใช้งาน');

    const token = createSession(foundUser.userId);

    return {
      success: true,
      message: 'เข้าสู่ระบบสำเร็จ',
      token: token,
      user: {
        userId: String(foundUser.userId),
        fullName: String(foundUser.fullName),
        studentId: String(foundUser.studentId),
        email: String(foundUser.email),
        phone: String(foundUser.phone),
        role: String(foundUser.role),
        status: String(foundUser.status)
      }
    };
  } catch (error) {
    return { success: false, message: error.message || 'เข้าสู่ระบบไม่สำเร็จ' };
  }
}

function createSession(userId) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('SESSION_' + token, String(userId), CONFIG.SESSION_SECONDS);
  return token;
}

function getUserFromSession(token) {
  if (!token) return null;
  const userId = CacheService.getScriptCache().get('SESSION_' + token);
  if (!userId) return null;

  const sheet = getSheet(CONFIG.SHEETS.USERS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][headers.indexOf('userId')]).trim() === String(userId).trim()) {
      return {
        userId: values[i][headers.indexOf('userId')],
        fullName: values[i][headers.indexOf('fullName')],
        studentId: values[i][headers.indexOf('studentId')],
        email: values[i][headers.indexOf('email')],
        phone: values[i][headers.indexOf('phone')],
        role: values[i][headers.indexOf('role')],
        status: values[i][headers.indexOf('status')]
      };
    }
  }
  return null;
}

/* =========================
   REPORT FUNCTIONS
========================= */

function createReport(token, reportData) {
  try {
    const user = getUserFromSession(token);
    if (!user) return { success: false, message: 'Session หมดอายุหรือไม่ถูกต้อง' };
    if (!reportData) return { success: false, message: 'ไม่พบข้อมูลรายการ' };

    const type = String(reportData.type || '').trim().toUpperCase();
    const itemName = String(reportData.itemName || '').trim();
    let categoryId = String(reportData.categoryId || '').trim();
    const categoryOther = String(reportData.categoryOther || '').trim();
    const description = String(reportData.description || '').trim();
    const distinguishingMark = String(reportData.distinguishingMark || '').trim();
    const date = String(reportData.date || '').trim();
    const time = String(reportData.time || '').trim();
    let locationId = String(reportData.locationId || '').trim();
    const locationOther = String(reportData.locationOther || '').trim();
    const locationDetail = String(reportData.locationDetail || '').trim();
    const imageUrl = String(reportData.imageUrl || '').trim();

    if (type !== 'LOST' && type !== 'FOUND') throw new Error('ประเภทต้องเป็น LOST หรือ FOUND');
    if (!itemName || !categoryId || !description || !date || !locationId) {
      throw new Error('กรุณากรอกข้อมูลสำคัญให้ครบถ้วน');
    }

    if (categoryId === '__OTHER__') categoryId = findOrCreateCategory_(categoryOther);
    if (locationId === '__OTHER__') locationId = findOrCreateLocation_(locationOther);

    const finalDescription = locationDetail 
      ? description + '\n[รายละเอียดจุดที่พบ: ' + locationDetail + ']' 
      : description;

    const sheet = getSheet(CONFIG.SHEETS.REPORTS);
    const reportId = generateId('REPORT');
    const now = new Date();

    sheet.appendRow([
      reportId, type, itemName, categoryId, finalDescription,
      distinguishingMark, date, time, locationId, imageUrl,
      user.userId, 'ACTIVE', now, now
    ]);

    const report = {
      reportId, type, itemName, categoryId, description: finalDescription,
      distinguishingMark, date, time, locationId, imageUrl,
      reporterId: user.userId, status: 'ACTIVE', createdAt: safeDate(now)
    };

    notifyAdminsOfNewReport(report, user);
    sendReportToLine(report, user);

    return {
      success: true,
      message: type === 'LOST' ? 'แจ้งของหายสำเร็จ' : 'แจ้งพบของสำเร็จ',
      report: report
    };
  } catch (error) {
    return { success: false, message: error.message || 'ไม่สามารถบันทึกรายการได้' };
  }
}

function uploadImage(token, imageData) {
  try {
    const user = getUserFromSession(token);
    if (!user) return { success: false, message: 'Session ไม่ถูกต้องหรือหมดอายุ' };
    if (!imageData || !imageData.base64) return { success: false, message: 'ไม่พบข้อมูลรูปภาพ' };

    let base64 = String(imageData.base64);
    if (base64.indexOf(',') !== -1) {
      base64 = base64.split(',')[1];
    }

    const decoded = Utilities.base64Decode(base64);
    const blob = Utilities.newBlob(
      decoded,
      imageData.mimeType || 'image/jpeg',
      imageData.fileName || ('upload-' + new Date().getTime() + '.jpg')
    );

    const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return {
      success: true,
      message: 'อัปโหลดรูปภาพสำเร็จ',
      imageUrl: 'https://drive.google.com/uc?export=view&id=' + file.getId()
    };
  } catch (error) {
    return { success: false, message: error.message || 'ไม่สามารถอัปโหลดรูปภาพได้' };
  }
}

/* =========================
   NOTIFICATIONS & LINE
========================= */

function notifyAdminsOfNewReport(report, reporter) {
  try {
    const usersResult = getAllSheetData(CONFIG.SHEETS.USERS);
    if (!usersResult.success) return;

    const admins = usersResult.data.filter(u => String(u.role).toUpperCase() === 'ADMIN');
    const notifSheet = ensureSheet(CONFIG.SHEETS.NOTIFICATIONS, SHEET_HEADERS.Notifications);

    const title = report.type === 'LOST' ? 'มีการแจ้งของหายใหม่' : 'มีการแจ้งพบของใหม่';
    const message = reporter.fullName + ' ได้แจ้ง ' + report.itemName;

    admins.forEach(admin => {
      notifSheet.appendRow([generateId('NOTIF'), admin.userId, title, message, 'NEW_REPORT', false, new Date()]);
    });
  } catch (e) {
    Logger.log('Notify Admins Error: ' + e.message);
  }
}

function notifyAdminsOfReturnedReport(report, reporter) {
  try {
    const usersResult = getAllSheetData(CONFIG.SHEETS.USERS);
    if (!usersResult.success) return;

    const admins = usersResult.data.filter(u => String(u.role).toUpperCase() === 'ADMIN');
    const notifSheet = ensureSheet(CONFIG.SHEETS.NOTIFICATIONS, SHEET_HEADERS.Notifications);
    const title = 'ผู้แจ้งยืนยันว่าได้รับของแล้ว';
    const message = reporter.fullName + ' ยืนยันว่าได้รับของรายการ ' + report.itemName + ' แล้ว';

    admins.forEach(admin => {
      notifSheet.appendRow([
        generateId('NOTIF'),
        admin.userId,
        title,
        message,
        'REPORT_RETURNED',
        false,
        new Date()
      ]);
    });
  } catch (e) {
    Logger.log('Notify Admins Returned Error: ' + e.message);
  }
}

function sendReportToLine(report, reporter) {
  try {
    const lineToken = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
    if (!lineToken) return;

    const typeText = report.type === 'LOST' ? '🔴 แจ้งของหาย' : '🟢 แจ้งพบของ';
    const textMsg = [
      typeText,
      'รายการ: ' + report.itemName,
      'วันที่: ' + report.date + ' ' + report.time,
      'ผู้แจ้ง: ' + reporter.fullName + ' (' + reporter.phone + ')',
      'รายละเอียด: ' + report.description
    ].join('\n');

    const payload = {
      to: CONFIG.LINE_ADMIN_USER_ID,
      messages: [{ type: 'text', text: textMsg }]
    };

    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + lineToken },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('LINE Push Error: ' + e.message);
  }
}

function getNotifications(token) {
  try {
    const user = getUserFromSession(token);
    if (!user) {
      return { success: false, message: 'Session หมดอายุหรือสิทธิ์ไม่ถูกต้อง', data: [] };
    }

    ensureSheet(CONFIG.SHEETS.NOTIFICATIONS, SHEET_HEADERS.Notifications);
    const dataResult = getAllSheetData(CONFIG.SHEETS.NOTIFICATIONS);
    if (!dataResult.success) return dataResult;

    const userNotifications = (dataResult.data || [])
      .filter(item => String(item.userId).trim() === String(user.userId).trim())
      .map(item => ({
        notificationId: item.notificationId || '',
        title: item.title || '',
        message: item.message || '',
        type: item.type || '',
        isRead: item.isRead === true || String(item.isRead).toLowerCase() === 'true',
        createdAt: item.createdAt || ''
      }))
      .reverse();

    return {
      success: true,
      data: userNotifications
    };
  } catch (error) {
    return {
      success: false,
      message: error.message || 'ไม่สามารถดึงข้อมูลแจ้งเตือนได้',
      data: []
    };
  }
}

/* ==========================================================================
   ADMIN FUNCTIONS SYSTEM
   ========================================================================== */

function checkAdminPermission_(token) {
  const user = getUserFromSession(token);
  if (!user) {
    throw new Error('Session หมดอายุหรือไม่พบข้อมูลล็อกอิน');
  }
  if (String(user.role).toUpperCase() !== 'ADMIN') {
    throw new Error('คุณไม่มีสิทธิ์ใช้งานส่วนนี้ (สำหรับ Admin เท่านั้น)');
  }
  return user;
}

function logAudit_(userId, action, target, targetId, description) {
  try {
    const sheet = ensureSheet(CONFIG.SHEETS.AUDIT_LOGS, SHEET_HEADERS.AuditLogs);
    sheet.appendRow([
      generateId('LOG'),
      userId,
      action,
      target,
      targetId,
      description,
      new Date()
    ]);
  } catch (e) {
    Logger.log('Audit Log Error: ' + e.message);
  }
}

function getAdminDashboardStats(token) {
  try {
    checkAdminPermission_(token);

    const users = getAllSheetData(CONFIG.SHEETS.USERS).data || [];
    const reports = getAllSheetData(CONFIG.SHEETS.REPORTS).data || [];
    const categories = getAllSheetData(CONFIG.SHEETS.CATEGORIES).data || [];

    const totalUsers = users.length;
    const totalLost = reports.filter(r => String(r.type).toUpperCase() === 'LOST').length;
    const totalFound = reports.filter(r => String(r.type).toUpperCase() === 'FOUND').length;
    const totalResolved = reports.filter(r => ['RESOLVED', 'RETURNED', 'CLOSED'].includes(String(r.status).toUpperCase())).length;

    return {
      success: true,
      stats: {
        totalUsers,
        totalLost,
        totalFound,
        totalResolved,
        totalCategories: categories.length
      }
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function getAdminUsers(token) {
  try {
    checkAdminPermission_(token);
    const users = getAllSheetData(CONFIG.SHEETS.USERS).data || [];
    
    const safeUsers = users.map(u => {
      const copy = Object.assign({}, u);
      delete copy.password;
      return copy;
    });

    return { success: true, data: safeUsers };
  } catch (error) {
    return { success: false, message: error.message, data: [] };
  }
}

function updateUserRoleStatus(token, targetUserId, newRole, newStatus) {
  try {
    const adminUser = checkAdminPermission_(token);
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const values = sheet.getDataRange().getValues();
    const headers = values[0];

    const uIdx = headers.indexOf('userId');
    const roleIdx = headers.indexOf('role');
    const statusIdx = headers.indexOf('status');
    const updateIdx = headers.indexOf('updatedAt');

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][uIdx]).trim() === String(targetUserId).trim()) {
        if (newRole) sheet.getRange(i + 1, roleIdx + 1).setValue(newRole);
        if (newStatus) sheet.getRange(i + 1, statusIdx + 1).setValue(newStatus);
        sheet.getRange(i + 1, updateIdx + 1).setValue(new Date());

        logAudit_(adminUser.userId, 'UPDATE_USER', 'Users', targetUserId, 'Updated role to ' + newRole + ', status to ' + newStatus);

        return { success: true, message: 'อัปเดตข้อมูลผู้ใช้งานสำเร็จ' };
      }
    }
    throw new Error('ไม่พบข้อมูลผู้ใช้งานที่ระบุ');
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function getAdminReports(token) {
  try {
    checkAdminPermission_(token);
    const reports = getAllSheetData(CONFIG.SHEETS.REPORTS).data || [];
    return { success: true, data: reports };
  } catch (error) {
    return { success: false, message: error.message, data: [] };
  }
}

function updateReportStatus(token, reportId, newStatus) {
  try {
    const adminUser = checkAdminPermission_(token);
    const sheet = getSheet(CONFIG.SHEETS.REPORTS);
    const values = sheet.getDataRange().getValues();
    const headers = values[0];

    const rIdx = headers.indexOf('reportId');
    const statusIdx = headers.indexOf('status');
    const updateIdx = headers.indexOf('updatedAt');

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][rIdx]).trim() === String(reportId).trim()) {
        sheet.getRange(i + 1, statusIdx + 1).setValue(newStatus);
        sheet.getRange(i + 1, updateIdx + 1).setValue(new Date());

        logAudit_(adminUser.userId, 'UPDATE_REPORT_STATUS', 'Reports', reportId, 'Updated status to ' + newStatus);

        return { success: true, message: 'อัปเดตสถานะรายการสำเร็จ' };
      }
    }
    throw new Error('ไม่พบรายการที่ระบุ');
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function confirmReportReturned(token, reportId) {
  try {
    const user = getUserFromSession(token);
    if (!user) {
      return { success: false, message: 'Session หมดอายุหรือไม่ถูกต้อง' };
    }
    if (!reportId) {
      return { success: false, message: 'ไม่พบรหัสรายการ' };
    }

    const sheet = getSheet(CONFIG.SHEETS.REPORTS);
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const rIdx = headers.indexOf('reportId');
    const userIdx = headers.indexOf('reporterId');
    const statusIdx = headers.indexOf('status');
    const updateIdx = headers.indexOf('updatedAt');
    const typeIdx = headers.indexOf('type');
    const itemIdx = headers.indexOf('itemName');

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][rIdx]).trim() === String(reportId).trim()) {
        if (String(values[i][userIdx]).trim() !== String(user.userId).trim()) {
          return { success: false, message: 'คุณไม่มีสิทธิ์ยืนยันรายการนี้' };
        }

        sheet.getRange(i + 1, statusIdx + 1).setValue('RETURNED');
        sheet.getRange(i + 1, updateIdx + 1).setValue(new Date());

        const report = {
          reportId: String(values[i][rIdx]),
          type: String(values[i][typeIdx] || ''),
          itemName: String(values[i][itemIdx] || '')
        };

        // แจ้งเฉพาะในระบบ Admin และจงใจไม่เรียก sendReportToLine()
        notifyAdminsOfReturnedReport(report, user);
        logAudit_(user.userId, 'CONFIRM_REPORT_RETURNED', 'Reports', reportId, 'User confirmed item returned');

        return { success: true, message: 'ยืนยันว่าได้รับของแล้ว และแจ้งผู้ดูแลระบบเรียบร้อยแล้ว' };
      }
    }

    return { success: false, message: 'ไม่พบรายการที่ระบุ' };
  } catch (error) {
    return { success: false, message: error.message || 'ไม่สามารถยืนยันการรับของได้' };
  }
}

function deleteReportByAdmin(token, reportId, isPermanent) {
  try {
    const adminUser = checkAdminPermission_(token);
    if (!reportId) throw new Error('กรุณาระบุรหัสโพสต์ที่ต้องการลบ');

    const sheet = getSheet(CONFIG.SHEETS.REPORTS);
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const rIdx = headers.indexOf('reportId');
    const statusIdx = headers.indexOf('status');
    const updateIdx = headers.indexOf('updatedAt');

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][rIdx]).trim() === String(reportId).trim()) {
        const itemName = values[i][headers.indexOf('itemName')];

        if (isPermanent === true) {
          sheet.deleteRow(i + 1);
          logAudit_(adminUser.userId, 'PERMANENT_DELETE_REPORT', 'Reports', reportId, 'Permanently deleted inappropriate post: ' + itemName);
          return { success: true, message: 'ลบโพสต์ที่ไม่เหมาะสมออกจากระบบถาวรเรียบร้อยแล้ว' };
        } else {
          sheet.getRange(i + 1, statusIdx + 1).setValue('DELETED');
          sheet.getRange(i + 1, updateIdx + 1).setValue(new Date());
          logAudit_(adminUser.userId, 'DELETE_REPORT', 'Reports', reportId, 'Marked post as DELETED due to inappropriate content: ' + itemName);
          return { success: true, message: 'ซ่อน/ลบโพสต์ที่ไม่เหมาะสมเรียบร้อยแล้ว' };
        }
      }
    }
    throw new Error('ไม่พบโพสต์ที่ต้องการลบ');
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function addCategory(token, categoryName) {
  try {
    const adminUser = checkAdminPermission_(token);
    if (!categoryName) throw new Error('กรุณาระบุชื่อหมวดหมู่');

    const catId = findOrCreateCategory_(categoryName);
    logAudit_(adminUser.userId, 'ADD_CATEGORY', 'Categories', catId, 'Added category: ' + categoryName);

    return { success: true, message: 'เพิ่มหมวดหมู่เรียบร้อย', categoryId: catId };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function addLocation(token, locationName, description) {
  try {
    const adminUser = checkAdminPermission_(token);
    if (!locationName) throw new Error('กรุณาระบุชื่อสถานที่');

    const locId = findOrCreateLocation_(locationName);
    logAudit_(adminUser.userId, 'ADD_LOCATION', 'Locations', locId, 'Added location: ' + locationName);

    return { success: true, message: 'เพิ่มสถานที่เรียบร้อย', locationId: locId };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function getAuditLogs(token) {
  try {
    checkAdminPermission_(token);
    const logs = getAllSheetData(CONFIG.SHEETS.AUDIT_LOGS).data || [];
    return { success: true, data: logs.reverse() };
  } catch (error) {
    return { success: false, message: error.message, data: [] };
  }
}

function exportSheetData(token, sheetName) {
  try {
    checkAdminPermission_(token);
    return getAllSheetData(sheetName);
  } catch (error) {
    return { success: false, message: error.message, data: [] };
  }
}
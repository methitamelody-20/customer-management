// ============================================================
// ระบบจัดการและติดตามเอกสารการสอน มสธ.
// Code.gs  v3  — เชื่อม MainSystem.html + CRM_Public.html จริง
// ============================================================

const SH_DATA     = 'ข้อมูลพัสดุ';
const SH_SETTINGS = 'ตั้งค่า';
const SH_USERS    = 'ผู้ใช้งาน';
const SH_CRM      = 'CRM แจ้งปัญหา';
const SH_TAGS     = 'Tags';
const SH_INVEST   = 'สอบสวนไปรษณีย์';
const R_SUPER  = 'superadmin';
const R_STAFF  = 'staff';
const R_VIEWER = 'viewer';

// ============================================================
// ENTRY POINTS
// ============================================================
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  // ถ้ามี ?page=crm → ให้หน้าสำหรับ นศ. แจ้งปัญหา
  if (params.page === 'crm') {
    const tpl = HtmlService.createTemplateFromFile('CRM_Public');
    return tpl.evaluate()
      .setTitle('แจ้งปัญหาเอกสารการสอน — มสธ.')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport','width=device-width,initial-scale=1');
  }
  // ถ้ามี ?page=ext_staff → Dashboard เจ้าหน้าที่ภายนอก
  if (params.page === 'ext_staff') {
    const tpl = HtmlService.createTemplateFromFile('ExternalStaffDashboard');
    return tpl.evaluate()
      .setTitle('Dashboard เจ้าหน้าที่ภายนอก — มสธ.')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport','width=device-width,initial-scale=1');
  }
  const tpl = HtmlService.createTemplateFromFile('Mainsystem');
  tpl.setpwToken = params.token || '';
  tpl.setpwEmail = params.email || '';
  return tpl.evaluate()
    .setTitle('ระบบจัดการและติดตามเอกสารการสอน มสธ.')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport','width=device-width,initial-scale=1');
}
function doGetPublic(e) {
  const tpl = HtmlService.createTemplateFromFile('CRM_Public');
  return tpl.evaluate()
    .setTitle('แจ้งปัญหาเอกสารการสอน — มสธ.')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport','width=device-width,initial-scale=1');
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// AUTH
// ============================================================
function login(email, password) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_USERS);
    if (!sheet) return { success:false, error:'ยังไม่ได้ setup กรุณารัน setupSystem() ก่อน' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const rowEmail  = (data[i][0]||'').toString().trim().toLowerCase();
      const rowHash   = (data[i][1]||'').toString().trim();
      const rowRole   = (data[i][2]||'').toString().trim();
      const rowName   = (data[i][3]||'').toString().trim();
      const rowActive = data[i][4];
      const rowPerms  = (data[i][6]||'').toString().trim();
      if (rowEmail === email.toLowerCase().trim()) {
        if (!rowActive) return { success:false, error:'บัญชีนี้ถูกระงับ' };
        if (hashPw(password) === rowHash) {
          // ใช้ ScriptProperties + key = email (ตรงกับ checkSession)
          const sessData = {email:rowEmail, role:rowRole, name:rowName, perms:rowPerms, ts:Date.now()};
          PropertiesService.getScriptProperties().setProperty('sess_'+rowEmail, JSON.stringify(sessData));
          sheet.getRange(i+1, 6).setValue(fmtDate(new Date()));
          Logger.log('Login success: '+rowEmail+' role='+rowRole);
          return { success:true, role:rowRole, name:rowName, email:rowEmail, perms:rowPerms };
        }
        return { success:false, error:'รหัสผ่านไม่ถูกต้อง' };
      }
    }
    return { success:false, error:'ไม่พบ Email นี้ในระบบ' };
  } catch(e) { return { success:false, error:e.message }; }
}

function logout() {
  PropertiesService.getScriptProperties().deleteProperty('sess_'+Session.getEffectiveUser().getEmail());
  return { success:true };
}

function checkSession() {
  try {
    const sp = PropertiesService.getScriptProperties();
    
    // ลอง effective user email ก่อน
    let userEmail = '';
    try { userEmail = Session.getEffectiveUser().getEmail() || ''; } catch(e){}
    
    if (userEmail) {
      const raw = sp.getProperty('sess_'+userEmail);
      if (raw) {
        try {
          const s = JSON.parse(raw);
          if (s && s.email && s.role && (Date.now()-s.ts < 8*60*60*1000)) {
            return {valid:true, role:s.role, name:s.name||s.email, email:s.email, perms:s.perms||''};
          }
          sp.deleteProperty('sess_'+userEmail);
        } catch(e) { sp.deleteProperty('sess_'+userEmail); }
      }
    }

    // ถ้าไม่มี session ของ effective user — ค้นหา session ที่ valid จากทุก key
    const allProps = sp.getProperties();
    for (const key in allProps) {
      if (!key.startsWith('sess_')) continue;
      try {
        const s = JSON.parse(allProps[key]);
        if (s && s.email && s.role && (Date.now()-s.ts < 8*60*60*1000)) {
          return {valid:true, role:s.role, name:s.name||s.email, email:s.email, perms:s.perms||''};
        }
      } catch(e) {}
    }
    
    return {valid:false};
  } catch(e) { return {valid:false}; }
}

function hashPw(pw) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw, Utilities.Charset.UTF_8);
  return b.map(x=>('0'+(x&0xff).toString(16)).slice(-2)).join('');
}
function _sess() { return checkSession(); }
// Helper: auto-refresh session จาก Google account
function _autoRefreshSession() {
  // ตรวจ session ปัจจุบัน
  const sess = checkSession();
  if (sess.valid) return true;
  
  // session หมด — refresh จาก Google effective user
  try {
    let email = '';
    try { email = Session.getEffectiveUser().getEmail(); } catch(e){}
    if (!email) return false;
    
    // หาใน Sheet ผู้ใช้
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_USERS);
    if (sh && sh.getLastRow() > 1) {
      const rows = sh.getDataRange().getValues();
      for (let i=1; i<rows.length; i++) {
        if ((rows[i][0]||'').toLowerCase().trim() === email.toLowerCase().trim()) {
          const newSess = {email:email, role:rows[i][2]||R_STAFF, name:rows[i][3]||email, ts:Date.now()};
          PropertiesService.getScriptProperties().setProperty('sess_'+Session.getEffectiveUser().getEmail(), JSON.stringify(newSess));
          Logger.log('Auto-refreshed session for: '+email+' role: '+newSess.role);
          return true; // session ใหม่ถูก set แล้ว ส่ง true เลย
        }
      }
      // email มี แต่ไม่อยู่ใน Sheet — สร้าง superadmin session
      const newSess = {email:email, role:R_SUPER, name:email, ts:Date.now()};
      PropertiesService.getScriptProperties().setProperty('sess_'+Session.getEffectiveUser().getEmail(), JSON.stringify(newSess));
      Logger.log('Auto-created superadmin session for: '+email);
      return true;
    }
  } catch(e) {
    Logger.log('_autoRefreshSession error: '+e.message);
  }
  return false;
}

function _can(roles) {
  const s = checkSession();
  return s.valid && roles.includes(s.role);
}

// ============================================================
// FIREBASE INTEGRATION
// ============================================================
const FIREBASE_CONFIG = {
  apiKey: "AIzaSyDpRytR1M8rLsckJXdJb3HTaHxP2S53HWc",
  authDomain: "management-crm-stoupost.firebaseapp.com",
  projectId: "management-crm-stoupost",
  databaseURL: "https://management-crm-stoupost-default-rtdb.asia-southeast1.firebasedatabase.app",
  storageBucket: "management-crm-stoupost.firebasestorage.app",
  messagingSenderId: "766285421505",
  appId: "1:766285421505:web:6fe74c474b386a6433ad40"
};

const USE_FIREBASE = false; // ตั้งเป็น true เมื่อต้องการใช้ Firebase

// Firebase REST API Helper
function firebaseCall(method, path, data = null) {
  try {
    const url = FIREBASE_CONFIG.databaseURL + path + '.json?auth=' + getFirebaseToken();
    const options = {
      method: method,
      contentType: 'application/json',
      muteHttpExceptions: true
    };
    if (data) {
      options.payload = JSON.stringify(data);
    }
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    if (response.getResponseCode() >= 400) {
      Logger.log('Firebase error: ' + response.getResponseCode() + ' ' + JSON.stringify(result));
      return null;
    }
    return result;
  } catch(e) {
    Logger.log('firebaseCall error: ' + e.message);
    return null;
  }
}

// Get Firebase token (service account or admin SDK would be better, but using current user for now)
function getFirebaseToken() {
  // Note: ในสภาพจริง ควรใช้ Firebase Admin SDK หรือ service account
  // สำหรับตอนนี้ใช้ API key แทน (จะต้องตั้ง security rules ให้เหมาะสม)
  return FIREBASE_CONFIG.apiKey;
}

// ============================================================
// FIREBASE - AUTH FUNCTIONS
// ============================================================
function loginWithFirebase(email, password) {
  if (!USE_FIREBASE) return login(email, password);
  try {
    // Firebase REST API สำหรับ login
    const url = 'https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=' + FIREBASE_CONFIG.apiKey;
    const payload = {
      email: email,
      password: password,
      returnSecureToken: true
    };
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const result = JSON.parse(response.getContentText());
    if (response.getResponseCode() !== 200) {
      return { success:false, error:'อีเมลหรือรหัสผ่านไม่ถูกต้อง' };
    }
    // ดึงข้อมูล user จาก Firebase Realtime Database
    const userData = firebaseCall('GET', '/users/' + result.localId);
    if (!userData) {
      return { success:false, error:'ไม่พบข้อมูลผู้ใช้งาน' };
    }
    const sessData = {
      email: email,
      role: userData.role || 'staff',
      name: userData.name || email,
      uid: result.localId,
      token: result.idToken,
      ts: Date.now()
    };
    PropertiesService.getScriptProperties().setProperty('sess_' + email, JSON.stringify(sessData));
    return {
      success: true,
      role: userData.role || 'staff',
      name: userData.name || email,
      email: email
    };
  } catch(e) {
    return { success:false, error:e.message };
  }
}

function checkSessionFirebase() {
  if (!USE_FIREBASE) return checkSession();
  try {
    const sp = PropertiesService.getScriptProperties();
    let userEmail = '';
    try { userEmail = Session.getEffectiveUser().getEmail() || ''; } catch(e){}

    if (userEmail) {
      const raw = sp.getProperty('sess_'+userEmail);
      if (raw) {
        try {
          const s = JSON.parse(raw);
          if (s && s.email && s.role && (Date.now()-s.ts < 8*60*60*1000)) {
            return {valid:true, role:s.role, name:s.name||s.email, email:s.email, uid:s.uid};
          }
          sp.deleteProperty('sess_'+userEmail);
        } catch(e) { sp.deleteProperty('sess_'+userEmail); }
      }
    }
    return {valid:false};
  } catch(e) { return {valid:false}; }
}

// ============================================================
// FIREBASE - USERS MANAGEMENT
// ============================================================
function getUsersFirebase() {
  if (!USE_FIREBASE) return getUsers();
  try {
    const usersData = firebaseCall('GET', '/users');
    if (!usersData) return [];
    const users = [];
    for (const uid in usersData) {
      const u = usersData[uid];
      users.push({
        uid: uid,
        email: u.email,
        role: u.role,
        name: u.name,
        active: u.active !== false,
        lastLogin: u.lastLogin || '-',
        perms: u.perms || ''
      });
    }
    return users;
  } catch(e) { return { error:e.message }; }
}

function addUserFirebase(email, password, role, name, perms) {
  if (!USE_FIREBASE) return addUser(email, password, role, name, perms);

  if (!_autoRefreshSession()) return { success:false, error:'ไม่มีสิทธิ์' };
  try {
    // สร้าง user ผ่าน Firebase Auth REST API
    const signUpUrl = 'https://identitytoolkit.googleapis.com/v1/accounts:signUp?key=' + FIREBASE_CONFIG.apiKey;
    const signUpPayload = {
      email: email,
      password: password,
      returnSecureToken: true
    };
    const signUpResponse = UrlFetchApp.fetch(signUpUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(signUpPayload),
      muteHttpExceptions: true
    });
    const signUpResult = JSON.parse(signUpResponse.getContentText());
    if (signUpResponse.getResponseCode() !== 200) {
      return { success:false, error:'ไม่สามารถสร้าง user ได้: ' + signUpResult.error.message };
    }
    const uid = signUpResult.localId;

    // เก็บข้อมูล user ใน Realtime Database
    const userData = {
      email: email,
      role: role,
      name: name,
      active: true,
      perms: perms || '',
      createdAt: new Date().toISOString()
    };
    const result = firebaseCall('PUT', '/users/' + uid, userData);
    if (!result) {
      return { success:false, error:'ไม่สามารถบันทึกข้อมูล user ได้' };
    }
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

// ============================================================
// FIREBASE - RECORDS/PACKAGES
// ============================================================
function addRecordFirebase(data) {
  if (!USE_FIREBASE) return addRecord(data);

  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const now = new Date();
    const sess = _sess();
    const pfx = {return:'P', lend:'L', special:'S'}[data.recType]||'P';
    const id = pfx + Utilities.formatDate(now,'Asia/Bangkok','yyyyMMdd') + '-' + Math.random().toString(36).substr(2, 9);

    let status = 'บันทึกแล้ว';
    if (data.send2Track) status = 'ส่งแล้วครั้งที่ 2';
    else if (data.send1Track) status = 'ส่งแล้วครั้งที่ 1';

    const record = {
      id: id,
      date: now.toISOString(),
      term: data.term || '',
      year: data.year || '',
      recType: data.recType || '',
      parcelType: data.parcelType || '',
      courseCode: data.courseCode || '',
      studentId: data.studentId || '',
      prefix: data.prefix || '',
      firstName: data.firstName || '',
      lastName: data.lastName || '',
      houseNo: data.houseNo || '',
      street: data.street || '',
      subDistrict: data.subDistrict || '',
      district: data.district || '',
      province: data.province || '',
      zipCode: data.zipCode || '',
      phone: data.phone || '',
      cause: data.cause || '',
      contactStatus: data.contactStatus || '',
      send1Track: data.send1Track || '',
      send1Date: data.send1Date || '',
      send2Track: data.send2Track || '',
      send2Date: data.send2Date || '',
      tags: data.tags || '',
      remark: data.remark || '',
      courses: data.courses || '[]',
      status: status,
      updatedAt: now.toISOString(),
      recorder: sess.name || sess.email || 'ผู้ใช้งาน'
    };

    const result = firebaseCall('PUT', '/records/' + id, record);
    if (!result) {
      return { success:false, error:'ไม่สามารถบันทึกข้อมูลได้' };
    }
    logAudit('บันทึกพัสดุ (Firebase)', id + ' | ' + data.recType + ' | นศ.' + data.studentId + ' | ' + data.courseCode);
    return { success:true, id:id };
  } catch(e) { return { success:false, error:e.message }; }
}

function getRecordsFirebase(filters) {
  if (!USE_FIREBASE) return getRecords(filters);

  try {
    const recordsData = firebaseCall('GET', '/records');
    if (!recordsData) return [];

    let rows = [];
    for (const id in recordsData) {
      const r = recordsData[id];
      rows.push({
        id: r.id || id,
        date: r.date || '',
        term: r.term || '',
        year: r.year || '',
        recType: r.recType || '',
        parcelType: r.parcelType || '',
        courseCode: r.courseCode || '',
        studentId: r.studentId || '',
        prefix: r.prefix || '',
        firstName: r.firstName || '',
        lastName: r.lastName || '',
        houseNo: r.houseNo || '',
        street: r.street || '',
        subDistrict: r.subDistrict || '',
        district: r.district || '',
        province: r.province || '',
        zipCode: r.zipCode || '',
        phone: r.phone || '',
        cause: r.cause || '',
        contactStatus: r.contactStatus || '',
        send1Track: r.send1Track || '',
        send1Date: r.send1Date || '',
        send2Track: r.send2Track || '',
        send2Date: r.send2Date || '',
        tags: r.tags || '',
        remark: r.remark || '',
        courses: r.courses || '[]',
        status: r.status || '',
        updatedAt: r.updatedAt || '',
        recorder: r.recorder || ''
      });
    }

    if (filters) {
      let filtered = rows;
      if (filters.recType) filtered = filtered.filter(function(r){ return r.recType === filters.recType; });
      if (filters.term) filtered = filtered.filter(function(r){ return r.term == filters.term; });
      if (filters.year) filtered = filtered.filter(function(r){ return r.year == filters.year; });
      if (filters.contactStatus) filtered = filtered.filter(function(r){ return r.contactStatus === filters.contactStatus; });
      if (filters.courseCode) filtered = filtered.filter(function(r){ return r.courseCode.indexOf(filters.courseCode) !== -1; });
      if (filters.search) {
        const q = String(filters.search).toLowerCase();
        filtered = filtered.filter(function(r) {
          return [r.studentId, r.firstName, r.lastName, r.province, r.district, r.zipCode, r.courseCode, r.tags, r.send1Track, r.send2Track]
            .some(function(v) { return v && v.toLowerCase().indexOf(q) !== -1; });
        });
      }
      return filtered;
    }
    return rows;
  } catch(e) {
    return { error: 'getRecordsFirebase error: ' + e.message };
  }
}

// ============================================================
// FIREBASE - CRM TICKETS
// ============================================================
function addCrmTicketFirebase(data) {
  if (!USE_FIREBASE) return addCrmTicket(data);

  try {
    const now = new Date();
    const sess = _sess();
    const id = 'CRM-' + Utilities.formatDate(now,'Asia/Bangkok','yyyyMMdd') + '-' + Math.random().toString(36).substr(2, 9);
    const assigneeName = data.assigneeName || (sess.valid ? sess.name||sess.email : '');
    const recorderName = data.recorderName || (sess.valid ? sess.name||sess.email : 'ผู้แจ้งออนไลน์');

    const ticket = {
      id: id,
      date: now.toISOString(),
      reporterName: data.reporterName || '',
      studentId: data.studentId || '',
      reporterEmail: data.reporterEmail || '',
      reporterPhone: data.reporterPhone || '',
      department: data.department || '',
      educationLevel: data.educationLevel || '',
      term: data.term || '',
      year: data.year || '',
      courses: data.courses || '',
      issueType: data.issueType || '',
      detail: data.detail || '',
      channel: data.channel || 'online',
      priority: data.priority || 'normal',
      tags: data.tags || '',
      status: 'open',
      assigneeName: assigneeName,
      assigneeEmail: '',
      replies: '[]',
      recorderName: recorderName
    };

    const result = firebaseCall('PUT', '/crm/' + id, ticket);
    if (!result) {
      return { success:false, error:'ไม่สามารถบันทึก CRM ได้' };
    }
    logAudit('รับเรื่อง CRM (Firebase)', id + ' | ' + data.reporterName + ' | ' + data.issueType + ' | ผู้รับ: ' + assigneeName);
    return { success:true, id:id };
  } catch(e) { return { success:false, error:e.message }; }
}

function getCrmTicketsFirebase(filters) {
  if (!USE_FIREBASE) return getCrmTickets(filters);

  try {
    const crmData = firebaseCall('GET', '/crm');
    if (!crmData) return [];

    let rows = [];
    for (const id in crmData) {
      const r = crmData[id];
      rows.push({
        id: r.id || id,
        date: r.date || '',
        reporterName: r.reporterName || '',
        studentId: r.studentId || '',
        reporterEmail: r.reporterEmail || '',
        reporterPhone: r.reporterPhone || '',
        department: r.department || '',
        educationLevel: r.educationLevel || '',
        term: r.term || '',
        year: r.year || '',
        courses: r.courses || '',
        issueType: r.issueType || '',
        detail: r.detail || '',
        channel: r.channel || '',
        priority: r.priority || '',
        tags: r.tags || '',
        status: r.status || '',
        assigneeName: r.assigneeName || '',
        assigneeEmail: r.assigneeEmail || '',
        replies: r.replies || '[]',
        recorderName: r.recorderName || ''
      });
    }

    let result = rows;
    if (filters) {
      if (filters.status) result = result.filter(function(r){ return r.status === filters.status; });
      if (filters.channel) result = result.filter(function(r){ return r.channel === filters.channel; });
      if (filters.search) {
        const q = String(filters.search).toLowerCase();
        result = result.filter(function(r) {
          return [r.reporterName, r.studentId, r.detail, r.issueType, r.department]
            .some(function(v){ return v && v.toLowerCase().indexOf(q) !== -1; });
        });
      }
    }
    return result.reverse();
  } catch(e) { return { error: 'getCrmTicketsFirebase: ' + e.message }; }
}

function replyCrmTicketFirebase(id, text, adminName, adminEmail) {
  if (!USE_FIREBASE) return replyCrmTicket(id, text, adminName);

  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const crmData = firebaseCall('GET', '/crm/' + id);
    if (!crmData) return { success:false, error:'ไม่พบ Ticket' };

    const replies = [];
    if (crmData.replies) {
      try {
        const parsed = JSON.parse(crmData.replies);
        replies.push(...parsed);
      } catch(e) {}
    }
    replies.push({
      name: adminName,
      email: adminEmail || '',
      text: text,
      time: new Date().toISOString()
    });

    const updates = {
      replies: JSON.stringify(replies)
    };
    if (crmData.status === 'open') {
      updates.status = 'inprogress';
    }

    const result = firebaseCall('PATCH', '/crm/' + id, updates);
    if (!result) {
      return { success:false, error:'ไม่สามารถบันทึก reply ได้' };
    }
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

function updateCrmStatusFirebase(id, status) {
  if (!USE_FIREBASE) return updateCrmStatus(id, status);

  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const result = firebaseCall('PATCH', '/crm/' + id, { status: status });
    if (!result) return { success:false, error:'ไม่สามารถอัปเดตได้' };
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

// ============================================================
// FIREBASE - INVESTIGATIONS
// ============================================================
function getInvestigationsFirebase(filters) {
  if (!USE_FIREBASE) return getInvestigations(filters);

  try {
    const investData = firebaseCall('GET', '/investigations');
    if (!investData) return [];

    let rows = [];
    for (const id in investData) {
      const r = investData[id];
      rows.push({
        investId: r.investId || id,
        refId: r.refId || '',
        createdAt: r.createdAt || '',
        itemNo: r.itemNo || '',
        barcode: r.barcode || '',
        sentDate: r.sentDate || '',
        courseCode: r.courseCode || '',
        courseName: r.courseName || '',
        weight: r.weight || '',
        fee: r.fee || '',
        recipientName: r.recipientName || '',
        recipientAddr: r.recipientAddr || '',
        cause: r.cause || '',
        notifyDate: r.notifyDate || '',
        notifyEmail: r.notifyEmail || '',
        notifyQty: r.notifyQty || '',
        replyDate: r.replyDate || '',
        replyDays: r.replyDays || '',
        result: r.result || '',
        resultDetail: r.resultDetail || '',
        status: r.status || 'รอส่ง',
        recorder: r.recorder || '',
        updatedAt: r.updatedAt || ''
      });
    }

    let result = rows;
    if (filters) {
      if (filters.status) result = result.filter(function(r){ return r.status === filters.status; });
      if (filters.search) {
        const q = String(filters.search).toLowerCase();
        result = result.filter(function(r) {
          return [r.investId, r.refId, r.barcode, r.itemNo, r.courseCode, r.courseName, r.recipientName]
            .some(function(v){ return v && v.toLowerCase().indexOf(q) !== -1; });
        });
      }
    }
    return result;
  } catch(e) { return { error: 'getInvestigationsFirebase: ' + e.message }; }
}

function addInvestigationFirebase(data) {
  if (!USE_FIREBASE) return addInvestigation(data);

  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sess = _sess();
    const now = new Date();
    const id = 'INV-' + Utilities.formatDate(now,'Asia/Bangkok','yyyyMMdd-HHmmss');

    const investigation = {
      investId: id,
      refId: data.refId || '',
      createdAt: now.toISOString(),
      itemNo: data.itemNo || '',
      barcode: data.barcode || '',
      sentDate: data.sentDate || '',
      courseCode: data.courseCode || '',
      courseName: data.courseName || '',
      weight: data.weight || '',
      fee: data.fee || '',
      recipientName: data.recipientName || '',
      recipientAddr: data.recipientAddr || '',
      cause: data.cause || '',
      notifyDate: '',
      notifyEmail: '',
      notifyQty: '',
      replyDate: '',
      replyDays: '',
      result: '',
      resultDetail: '',
      status: 'รอส่ง',
      recorder: sess.name || sess.email,
      updatedAt: now.toISOString()
    };

    const result = firebaseCall('PUT', '/investigations/' + id, investigation);
    if (!result) {
      return { success:false, error:'ไม่สามารถบันทึก investigation ได้' };
    }
    return { success:true, id:id };
  } catch(e) { return { success:false, error:e.message }; }
}

function updateInvestigationFirebase(investId, updates) {
  if (!USE_FIREBASE) return updateInvestigation(investId, updates);

  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const updateData = {};
    if (updates.itemNo !== undefined) updateData.itemNo = updates.itemNo;
    if (updates.barcode !== undefined) updateData.barcode = updates.barcode;
    if (updates.sentDate !== undefined) updateData.sentDate = updates.sentDate;
    if (updates.weight !== undefined) updateData.weight = updates.weight;
    if (updates.fee !== undefined) updateData.fee = updates.fee;
    if (updates.notifyDate !== undefined) updateData.notifyDate = updates.notifyDate;
    if (updates.notifyEmail !== undefined) updateData.notifyEmail = updates.notifyEmail;
    if (updates.notifyQty !== undefined) updateData.notifyQty = updates.notifyQty;
    if (updates.replyDate !== undefined) updateData.replyDate = updates.replyDate;
    if (updates.replyDays !== undefined) updateData.replyDays = updates.replyDays;
    if (updates.result !== undefined) updateData.result = updates.result;
    if (updates.resultDetail !== undefined) updateData.resultDetail = updates.resultDetail;
    if (updates.status !== undefined) updateData.status = updates.status;
    updateData.updatedAt = new Date().toISOString();

    const result = firebaseCall('PATCH', '/investigations/' + investId, updateData);
    if (!result) {
      return { success:false, error:'ไม่สามารถอัปเดตได้' };
    }
    logAudit('อัปเดตสอบสวน (Firebase)', investId + ' | สถานะ: ' + (updates.status||'-') + ' | ผล: ' + (updates.result||'-'));
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

function saveFollowUpRecordFirebase(data) {
  if (!USE_FIREBASE) return saveFollowUpRecord(data);
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sess = _sess();

    // Update CRM in Firebase
    const crmUpdates = {
      followUpType: data.type || '',
      followUpCause: data.cause || '',
      followUpDate: new Date().toISOString(),
      followUpNotes: data.notes || '',
      followUpStatus: 'บันทึกแล้ว'
    };

    const crmResult = firebaseCall('PATCH', '/crm/' + data.crmId, crmUpdates);
    if (!crmResult) {
      return { success:false, error:'ไม่สามารถอัปเดต CRM ได้' };
    }

    // Create follow-up record in Firebase records
    const followUpId = Utilities.getUuid();
    const followUpRecord = {
      id: followUpId,
      crmId: data.crmId,
      type: data.type,
      cause: data.cause,
      course: data.course,
      studentId: data.studentId,
      createdAt: new Date().toISOString(),
      status: 'บันทึกแล้ว',
      recorderName: sess.name || sess.email,
      notes: data.notes || ''
    };

    const recordResult = firebaseCall('PUT', '/followUps/' + followUpId, followUpRecord);
    if (!recordResult) {
      return { success:false, error:'ไม่สามารถบันทึก Follow-up ได้' };
    }

    logAudit('บันทึก Follow-up (Firebase)', data.crmId + ' | ' + data.type + ' | นศ.' + data.studentId);
    return { success:true, message:'บันทึก follow-up แล้ว', followUpId:followUpId };
  } catch(e) {
    Logger.log('saveFollowUpRecordFirebase error: ' + e.message);
    return { success:false, error:e.message };
  }
}

// ============================================================
// MIGRATION FUNCTIONS - Sheets → Firebase
// ============================================================
function migrateUsersToFirebase() {
  if (!USE_FIREBASE) return { success:false, error:'Firebase ยังไม่เปิดใช้' };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_USERS);
    if (!sheet) return { success:false, error:'ไม่พบ Sheet ผู้ใช้งาน' };

    const data = sheet.getDataRange().getValues();
    let count = 0;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const email = (row[0] || '').toString().trim();
      if (!email) continue; // skip empty rows

      const userData = {
        email: email,
        passwordHash: row[1],
        role: row[2] || 'staff',
        name: row[3] || '',
        active: row[4] !== false && row[4] !== 'FALSE',
        lastLogin: row[5] || '',
        perms: row[6] || ''
      };

      const safeKey = email.replace(/[@.]/g, '_');
      const result = firebaseCall('PUT', '/users/' + safeKey, userData);
      if (result) count++;
    }
    return { success:true, migratedCount:count, message:'โอนย้าย ' + count + ' ผู้ใช้งานแล้ว' };
  } catch(e) { return { success:false, error:e.message }; }
}

function migrateRecordsToFirebase() {
  if (!USE_FIREBASE) return { success:false, error:'Firebase ยังไม่เปิดใช้' };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_DATA);
    if (!sheet) return { success:false, error:'ไม่พบ Sheet ข้อมูลพัสดุ' };

    const data = sheet.getDataRange().getValues();
    let count = 0;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;

      const record = {
        id: row[0],
        date: row[1] instanceof Date ? row[1].toISOString() : row[1],
        term: row[2],
        year: row[3],
        recType: row[4],
        parcelType: row[5],
        courseCode: row[6],
        studentId: row[7],
        prefix: row[8],
        firstName: row[9],
        lastName: row[10],
        houseNo: row[11],
        street: row[12],
        subDistrict: row[13],
        district: row[14],
        province: row[15],
        zipCode: row[16],
        phone: row[17],
        cause: row[18],
        contactStatus: row[19],
        send1Track: row[20],
        send1Date: row[21] instanceof Date ? row[21].toISOString() : row[21],
        send2Track: row[22],
        send2Date: row[23] instanceof Date ? row[23].toISOString() : row[23],
        tags: row[24],
        remark: row[25],
        courses: row[26],
        status: row[27],
        updatedAt: row[28] instanceof Date ? row[28].toISOString() : row[28],
        recorder: row[29]
      };

      const result = firebaseCall('PUT', '/records/' + row[0], record);
      if (result) count++;
    }
    return { success:true, migratedCount:count, message:'โอนย้าย ' + count + ' พัสดุแล้ว' };
  } catch(e) { return { success:false, error:e.message }; }
}

function migrateCrmToFirebase() {
  if (!USE_FIREBASE) return { success:false, error:'Firebase ยังไม่เปิดใช้' };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CRM);
    if (!sheet) return { success:false, error:'ไม่พบ Sheet CRM' };

    const data = sheet.getDataRange().getValues();
    let count = 0;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;

      const ticket = {
        id: row[0],
        date: row[1] instanceof Date ? row[1].toISOString() : row[1],
        reporterName: row[2],
        studentId: row[3],
        reporterEmail: row[4],
        reporterPhone: row[5],
        department: row[6],
        educationLevel: row[7],
        term: row[8],
        year: row[9],
        courses: row[10],
        issueType: row[11],
        detail: row[12],
        channel: row[13],
        priority: row[14],
        tags: row[15],
        status: row[16],
        assigneeName: row[17],
        assigneeEmail: row[18],
        replies: row[19],
        recorderName: row[20]
      };

      const result = firebaseCall('PUT', '/crm/' + row[0], ticket);
      if (result) count++;
    }
    return { success:true, migratedCount:count, message:'โอนย้าย ' + count + ' CRM tickets แล้ว' };
  } catch(e) { return { success:false, error:e.message }; }
}

function migrateInvestigationsToFirebase() {
  if (!USE_FIREBASE) return { success:false, error:'Firebase ยังไม่เปิดใช้' };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_INVEST);
    if (!sheet) return { success:false, error:'ไม่พบ Sheet investigations' };

    const data = sheet.getDataRange().getValues();
    let count = 0;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;

      const investigation = {
        investId: row[0],
        refId: row[1],
        createdAt: row[2] instanceof Date ? row[2].toISOString() : row[2],
        itemNo: row[3],
        barcode: row[4],
        sentDate: row[5] instanceof Date ? row[5].toISOString() : row[5],
        courseCode: row[6],
        courseName: row[7],
        weight: row[8],
        fee: row[9],
        recipientName: row[10],
        recipientAddr: row[11],
        cause: row[12],
        notifyDate: row[13] instanceof Date ? row[13].toISOString() : row[13],
        notifyEmail: row[14],
        notifyQty: row[15],
        replyDate: row[16] instanceof Date ? row[16].toISOString() : row[16],
        replyDays: row[17],
        result: row[18],
        resultDetail: row[19],
        status: row[20],
        recorder: row[21],
        updatedAt: row[22] instanceof Date ? row[22].toISOString() : row[22]
      };

      const result = firebaseCall('PUT', '/investigations/' + row[0], investigation);
      if (result) count++;
    }
    return { success:true, migratedCount:count, message:'โอนย้าย ' + count + ' investigations แล้ว' };
  } catch(e) { return { success:false, error:e.message }; }
}

// Run all migrations at once
function migrateAllToFirebase() {
  if (!USE_FIREBASE) return { success:false, error:'Firebase ยังไม่เปิดใช้ - ตั้ง USE_FIREBASE = true ก่อน' };
  try {
    Logger.log('เริ่มการโอนย้ายข้อมูลทั้งหมดไปยัง Firebase...');
    const results = {
      users: migrateUsersToFirebase(),
      records: migrateRecordsToFirebase(),
      crm: migrateCrmToFirebase(),
      investigations: migrateInvestigationsToFirebase()
    };
    Logger.log('โอนย้ายเสร็จ: ' + JSON.stringify(results));
    return { success:true, results:results };
  } catch(e) {
    return { success:false, error:e.message };
  }
}

// Test Firebase connection
function testFirebaseConnection() {
  try {
    Logger.log('ทดสอบการเชื่อมต่อ Firebase...');
    const testData = { test: true, timestamp: new Date().toISOString() };
    const result = firebaseCall('PUT', '/test/connection', testData);
    if (result !== null) {
      Logger.log('✅ เชื่อมต่อ Firebase สำเร็จ');
      return { success:true, message:'เชื่อมต่อ Firebase สำเร็จ' };
    } else {
      Logger.log('❌ ไม่สามารถเชื่อมต่อ Firebase ได้');
      return { success:false, message:'ไม่สามารถเชื่อมต่อ Firebase ได้ - ตรวจสอบ Security Rules' };
    }
  } catch(e) {
    Logger.log('❌ Error: ' + e.message);
    return { success:false, error:e.message };
  }
}

// Get Firebase setup status
function getFirebaseStatus() {
  return {
    enabled: USE_FIREBASE,
    projectId: FIREBASE_CONFIG.projectId,
    databaseURL: FIREBASE_CONFIG.databaseURL,
    message: USE_FIREBASE
      ? '✅ Firebase เปิดใช้งาน - ระบบกำลังใช้ Firebase Realtime Database'
      : '⚠️ Firebase ปิดใช้งาน - ระบบใช้ Google Sheets แทน'
  };
}

// ============================================================
// SETTINGS
// ============================================================
var _settingsCache = null;
var _tagsCache = null;
var _yearsCache = null;

function getInitData() {
  try {
    // Load data in parallel for better performance
    const cache = CacheService.getUserCache();
    const cacheKey = 'initData_' + Session.getUser().getEmail();
    const cachedData = cache.get(cacheKey);

    // Return cached data if available (5 minute cache)
    if (cachedData) {
      return JSON.parse(cachedData);
    }

    const settings = getSettings();
    const tags = getTags();
    const years = getUniqueYears();
    const taskStats = getTaskStats();

    const result = {
      success: true,
      settings: settings,
      tags: tags,
      years: years,
      taskStats: taskStats
    };

    // Cache the result for 5 minutes
    try {
      cache.put(cacheKey, JSON.stringify(result), 300);
    } catch(e) {
      // Ignore cache errors - data will still be returned
    }

    return result;
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function getSettings() {
  try {
    if (_settingsCache) return _settingsCache;

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_SETTINGS);
    if (!sheet) {
      _settingsCache = { parcelTypes:[],returnCauses:[],prefixes:[],terms:[],branches:[],plans:[],issueTypes:[],assignees:[] };
      return _settingsCache;
    }
    const lr = sheet.getLastRow();
    if (lr < 2) {
      _settingsCache = { parcelTypes:[],returnCauses:[],prefixes:[],terms:[],branches:[],plans:[],issueTypes:[],assignees:[] };
      return _settingsCache;
    }
    const data = sheet.getRange(2, 1, lr-1, 2).getValues(); // ดึงเฉพาะ 2 คอลัมน์ที่ใช้
    const s = {};
    for (let i=0;i<data.length;i++) s[data[i][0]] = data[i][1]?data[i][1].toString().split(','):[];
    _settingsCache = {
      parcelTypes:  s['ประเภทพัสดุ']  || [],
      returnCauses: s['สาเหตุตีคืน'] || [],
      prefixes:     s['คำนำหน้า']    || [],
      terms:        s['ภาคการศึกษา'] || [],
      branches:     s['สาขาวิชา']    || [],
      plans:        s['แผนการศึกษา'] || [],
      issueTypes:   s['ประเภทปัญหา'] || [],
      assignees:    s['ผู้รับเรื่อง'] || [],
    };
    return _settingsCache;
  } catch(e) { return { error:e.message }; }
}

function getUniqueYears() {
  try {
    if (_yearsCache) return _yearsCache;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_DATA);
    if (!sheet) {
      _yearsCache = [];
      return _yearsCache;
    }
    const lr = sheet.getLastRow();
    if (lr < 2) {
      _yearsCache = [];
      return _yearsCache;
    }
    const years = sheet.getRange(2,4,lr-1,1).getValues().flat();
    _yearsCache = [...new Set(years.filter(y=>y!==''))].sort().reverse();
    return _yearsCache;
  } catch(e) { return []; }
}

// ============================================================
// TAGS
// ============================================================
function getTags() {
  try {
    if (_tagsCache) return _tagsCache;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_TAGS);
    if (!sheet) {
      _tagsCache = [];
      return _tagsCache;
    }
    const lr = sheet.getLastRow();
    if (lr < 2) {
      _tagsCache = [];
      return _tagsCache;
    }
    const data = sheet.getRange(2, 1, lr-1, 2).getValues();
    _tagsCache = data.map(r=>({name:r[0],color:r[1]||0})).filter(t=>t.name);
    return _tagsCache;
  } catch(e) { return []; }
}

function addTag(name, color) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_TAGS);
    const data  = sheet.getDataRange().getValues();
    if (data.slice(1).some(r=>r[0]===name)) return { success:false, error:'Tag นี้มีอยู่แล้ว' };
    sheet.appendRow([name, color||0]);
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

function deleteTag(name) {
  if (!_autoRefreshSession()) return { success:false, error:'ไม่มีสิทธิ์' };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_TAGS);
    const data  = sheet.getDataRange().getValues();
    for (let i=1;i<data.length;i++) {
      if (data[i][0]===name) { sheet.deleteRow(i+1); return { success:true }; }
    }
    return { success:false, error:'ไม่พบ Tag' };
  } catch(e) { return { success:false, error:e.message }; }
}

// ============================================================
// RECORDS (พัสดุตีคืน / ให้ยืม / พิเศษ)
// ============================================================
function addRecord(data) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    if (!data || !data.recType) return { success:false, error:'ข้อมูลไม่ครบ: recType' };
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_DATA);
    if (!sheet) return { success:false, error:'ไม่พบ Sheet ข้อมูลพัสดุ กรุณารัน setupSystem()' };
    const now   = new Date();
    const pfx   = {return:'P', lend:'L', special:'S'}[data.recType]||'P';
    const id    = pfx + Utilities.formatDate(now,'Asia/Bangkok','yyyyMMdd') + '-' + sheet.getLastRow();
    const sess  = _sess();
    let status  = 'บันทึกแล้ว';
    if (data.send2Track) status = 'ส่งแล้วครั้งที่ 2';
    else if (data.send1Track) status = 'ส่งแล้วครั้งที่ 1';
    // 30 columns: 0=id,1=date,2=term,3=year,4=recType,5=parcelType,6=courseCode,7=studentId,
    // 8=prefix,9=firstName,10=lastName,11=houseNo,12=street,13=subDistrict,14=district,
    // 15=province,16=zipCode,17=phone,18=cause,19=contactStatus,
    // 20=send1Track,21=send1Date,22=send2Track,23=send2Date,
    // 24=tags,25=remark,26=courses,27=status,28=updatedAt,29=recorder
    const row = [
      id, fmtDate(now), data.term, data.year, data.recType, data.parcelType||'',
      data.courseCode||'', data.studentId||'', data.prefix||'',
      data.firstName||'', data.lastName||'',
      data.houseNo||'', data.street||'', data.subDistrict||'',
      data.district||'', data.province||'', data.zipCode||'', data.phone||'',
      data.cause||'', data.contactStatus||'',
      data.send1Track||'', data.send1Date||'',
      data.send2Track||'', data.send2Date||'',
      data.tags||'', data.remark||'',
      data.courses||'[]',  // JSON array ของชุดวิชาทั้งหมด
      status, fmtDate(now), sess.name||sess.email||'ผู้ใช้งาน',
    ];
    sheet.appendRow(row);
    const lr = sheet.getLastRow();
    if (lr%2===0) sheet.getRange(lr,1,1,row.length).setBackground('#f0f4f8');
    logAudit('บันทึกพัสดุ', id+' | '+data.recType+' | นศ.'+data.studentId+' | '+data.courseCode);
    return { success:true, id:id };
  } catch(e) { return { success:false, error:e.message }; }
}

// กู้คืนข้อมูลที่ถูกบันทึกผิดคอลัมน์ (จาก bug ของ addRecord)
// ย้ายข้อมูลจาก AE→AA, AC→Y, AB→Z, AF→AB, AG→AC, AH→AD
function recoverShiftedColumns() {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_DATA);
    if (!sheet) return { success:false, error:'ไม่พบ Sheet' };
    const lr = sheet.getLastRow();
    if (lr < 2) return { success:true, recovered:0 };
    // อ่านคอลัมน์ AA-AH (27-34) ของทุกแถว
    const range = sheet.getRange(2, 25, lr-1, 10); // Y(25) ถึง AH(34)
    const values = range.getValues();
    let recovered = 0;
    for (let i = 0; i < values.length; i++) {
      const r = values[i];
      // r[0]=Y(25 tags), r[1]=Z(26 remark), r[2]=AA(27 courses), r[3]=AB(28 status),
      // r[4]=AC(29 updatedAt), r[5]=AD(30 recorder), r[6]=AE(31), r[7]=AF(32), r[8]=AG(33), r[9]=AH(34)
      // ถ้า AE(idx 6) มี JSON และ AA(idx 2) ว่าง = ข้อมูลผิดคอลัมน์
      const aeVal = String(r[6] || '');
      const aaVal = String(r[2] || '');
      if (aeVal && aeVal.startsWith('[') && !aaVal) {
        // ข้อมูลผิดคอลัมน์ — ย้ายกลับ
        const tags = String(r[0] || '');     // Y → ต้องเป็น tags
        const remark = String(r[1] || '');   // Z → ต้องเป็น remark
        const courses = aeVal;                // AE → AA (courses JSON)
        const status = String(r[7] || '');   // AF → AB (status)
        const updatedAt = r[8] || '';        // AG → AC (updatedAt)
        const recorder = String(r[9] || ''); // AH → AD (recorder)
        // เขียนคอลัมน์ที่ถูกต้อง Y(25) ถึง AD(30)
        sheet.getRange(i+2, 25).setValue(tags);      // Y = tags (เดิม)
        sheet.getRange(i+2, 26).setValue(remark);    // Z = remark (เดิม)
        sheet.getRange(i+2, 27).setValue(courses);   // AA = courses JSON
        sheet.getRange(i+2, 28).setValue(status);    // AB = status
        sheet.getRange(i+2, 29).setValue(updatedAt); // AC = updatedAt
        sheet.getRange(i+2, 30).setValue(recorder);  // AD = recorder
        // ล้าง AE(31) ถึง AH(34) ที่ไม่ใช้
        sheet.getRange(i+2, 31, 1, 4).clearContent();
        recovered++;
      }
    }
    logAudit('กู้คืนข้อมูลคอลัมน์', 'แก้ไข ' + recovered + ' แถว');
    return { success:true, recovered:recovered };
  } catch(e) {
    Logger.log('recoverShiftedColumns error: ' + e.message);
    return { success:false, error:e.message };
  }
}

var _recordsCache = null;
var _recordsCacheTime = 0;
const CACHE_TTL = 60000; // 60 seconds cache

function getRecords(filters) {
  // ไม่ตรวจ session — ใช้ login screen ฝั่ง HTML แทน
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_DATA);
    if (!sheet) return { error:'ไม่พบ Sheet กรุณารัน setupSystem()' };
    const lr    = sheet.getLastRow();
    if (lr < 2) return [];

    // ดึงทุก column ที่ใช้ (ต้อง 34 columns สำหรับข้อมูลพัสดุ)
    const numCols = Math.min(sheet.getLastColumn(), 35);
    const rawRows = sheet.getRange(2,1,lr-1,numCols).getValues();
    const rows = [];
    for (let i = 0; i < rawRows.length; i++) {
      const r = rawRows[i];
      if (!r[0]) continue; // skip empty rows
      
      // Safe string conversion
      const S = function(v) { return v == null ? '' : String(v); };
      // Safe number
      const N = function(v) { return v == null || v === '' ? '' : v; };
      // Safe date - handle both Date objects and strings
      const D = function(v) {
        if (!v) return '';
        try {
          if (v instanceof Date) {
            return Utilities.formatDate(v, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
          }
          return String(v);
        } catch(e) { return String(v); }
      };
      
      rows.push({
        id:          S(r[0]),
        date:        D(r[1]),
        term:        S(r[2]),
        year:        S(r[3]),
        recType:     S(r[4]),
        parcelType:  S(r[5]),
        courseCode:  S(r[6]),
        studentId:   S(r[7]),
        prefix:      S(r[8]),
        firstName:   S(r[9]),
        lastName:    S(r[10]),
        houseNo:     S(r[11]),
        street:      S(r[12]),
        subDistrict: S(r[13]),
        district:    S(r[14]),
        province:    S(r[15]),
        zipCode:     S(r[16]),
        phone:       S(r[17]),
        cause:       S(r[18]),
        contactStatus: S(r[19]),
        send1Track:  S(r[20]),
        send1Date:   D(r[21]),
        send2Track:  S(r[22]),
        send2Date:   D(r[23]),
        tags:        S(r[24]),
        remark:      S(r[25]),
        courses:     S(r[26]) || '[]',
        status:      S(r[27]),
        updatedAt:   D(r[28]),
        recorder:    S(r[29]),
      });
    }

    if (filters) {
      let filtered = rows;
      if (filters.recType)       filtered = filtered.filter(function(r){ return r.recType === filters.recType; });
      if (filters.term)          filtered = filtered.filter(function(r){ return r.term == filters.term; });
      if (filters.year)          filtered = filtered.filter(function(r){ return r.year == filters.year; });
      if (filters.contactStatus) filtered = filtered.filter(function(r){ return r.contactStatus === filters.contactStatus; });
      if (filters.courseCode)    filtered = filtered.filter(function(r){ return r.courseCode.indexOf(filters.courseCode) !== -1; });
      if (filters.province)      filtered = filtered.filter(function(r){ return r.province === filters.province; });
      if (filters.district)      filtered = filtered.filter(function(r){ return r.district === filters.district; });
      if (filters.zipCode)       filtered = filtered.filter(function(r){ return r.zipCode === filters.zipCode; });
      if (filters.tag)           filtered = filtered.filter(function(r){ return (r.tags||'').split(',').map(function(t){return t.trim();}).indexOf(filters.tag) !== -1; });
      if (filters.search) {
        const q = String(filters.search).toLowerCase();
        filtered = filtered.filter(function(r) {
          return [r.studentId, r.firstName, r.lastName, r.province, r.district, r.zipCode, r.courseCode, r.tags, r.send1Track, r.send2Track]
            .some(function(v) { return v && v.toLowerCase().indexOf(q) !== -1; });
        });
      }
      return filtered;
    }
    return rows;
  } catch(e) {
    return { error: 'getRecords error: ' + e.message + ' | stack: ' + (e.stack||'').substring(0,200) };
  }
}

function deleteRecord(id) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_DATA);
    const data  = sheet.getDataRange().getValues();
    for (let i=1;i<data.length;i++) {
      if (data[i][0]===id) { sheet.deleteRow(i+1); return { success:true }; }
    }
    return { success:false, error:'ไม่พบรายการ' };
  } catch(e) { return { success:false, error:e.message }; }
}

// ============================================================
// DASHBOARD
// ============================================================
function getDashboardData(filters) {
  // session check ผ่าน _autoRefreshSession อัตโนมัติ
  try {
    const records = getRecords(filters);
    if (records.error) return records;
    const crm = getCrmTickets({});
    const crmPending = Array.isArray(crm) ? crm.filter(c=>c.status==='open'||c.status==='inprogress').length : 0;
    const total      = records.length;
    const cnt = (key, val) => records.filter(r=>r[key]===val).length;
    const groupBy = (key) => {
      const m={};
      records.forEach(r=>{ if(r[key]) m[r[key]]=(m[r[key]]||0)+1; });
      return Object.entries(m).sort((a,b)=>b[1]-a[1]).map(([k,v])=>[k,v]);
    };
    const groupTag = () => {
      const m={};
      records.forEach(r=>{ (r.tags||'').split(',').filter(Boolean).forEach(t=>m[t]=(m[t]||0)+1); });
      return Object.entries(m).sort((a,b)=>b[1]-a[1]).map(([k,v])=>[k,v]);
    };
    return {
      total, records,
      countReturn:  cnt('recType','return'),
      countLend:    cnt('recType','lend'),
      countSpecial: cnt('recType','special'),
      countContact: cnt('contactStatus','yes'),
      crmPending,
      byCause:    groupBy('cause'),
      byType:     groupBy('parcelType'),
      byProvince: groupBy('province'),
      byDistrict: groupBy('district'),
      byTag:      groupTag(),
    };
  } catch(e) { return { error:e.message }; }
}

// ============================================================
// SEARCH
// ============================================================
function searchStudent(query) {
  // session check ผ่าน _autoRefreshSession อัตโนมัติ
  if (!query||query.trim().length<2) return { records:[], crm:[] };
  const q = query.trim().toLowerCase();
  const records = getRecords({ search:q });
  const crm = getCrmTickets({ search:q });
  return { records: Array.isArray(records)?records:[], crm: Array.isArray(crm)?crm:[] };
}

function searchStudentById(studentId, term) {
  // Search by student ID and optional term
  if (!_autoRefreshSession()) return { error:'SESSION_EXPIRED' };
  try {
    if (!studentId || studentId.trim().length < 2) {
      return { records:[], crm:[] };
    }

    const q = studentId.trim();
    const records = getRecords({ search:q });
    let crm = getCrmTickets({ search:q });

    // If term is provided, also filter by term
    if (term && term.trim()) {
      const t = term.trim();
      if (Array.isArray(records)) {
        records = records.filter(r => (r.term || '').toString() === t);
      }
    }

    return {
      records: Array.isArray(records) ? records : [],
      crm: Array.isArray(crm) ? crm : []
    };
  } catch(e) {
    return { error: e.message };
  }
}

// ============================================================
// CRM
// ============================================================
function addCrmTicket(data) {
  // CRM เปิดให้ทุกคนส่งได้ (รวมถึงหน้า public)
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CRM);
    const now   = new Date();
    const id    = 'CRM-' + Utilities.formatDate(now,'Asia/Bangkok','yyyyMMdd') + '-' + sheet.getLastRow();
    // ถ้าไม่มี assigneeName ให้ใช้ชื่อจาก session
    const sess = _sess();
    const assigneeName = data.assigneeName || (sess.valid ? sess.name||sess.email : '');
    const recorderName = data.recorderName || (sess.valid ? sess.name||sess.email : 'ผู้แจ้งออนไลน์');
    const source = data.source || (sess.valid ? 'admin' : 'external');
    const reportDate = data.reportDate ? fmtDate(new Date(data.reportDate)) : '';
    const row   = [
      id, fmtDate(now), data.reporterName||'', data.studentId||'',
      data.reporterEmail||'', data.reporterPhone||'', data.department||'',
      data.educationLevel||'', data.term||'', data.year||'',
      data.courses||'', data.issueType||'', data.detail||'',
      data.channel||'online', data.priority||'normal',
      data.tags||'', 'open', assigneeName, '', '[]',
      recorderName, source, data.plan||'', data.org||'', reportDate,
    ];
    sheet.appendRow(row);
    const lr = sheet.getLastRow();
    if (lr%2===0) sheet.getRange(lr,1,1,row.length).setBackground('#f0f8ff');
    logAudit('รับเรื่อง CRM', id+' | '+data.reporterName+' | '+data.issueType+' | ผู้รับ: '+assigneeName);
    return { success:true, id:id };
  } catch(e) { return { success:false, error:e.message }; }
}

function updateCrmTicket(crmId, data) {
  if (!_autoRefreshSession()) return { success:false, error:'ไม่มีสิทธิ์' };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CRM);
    if (!sheet) return { success:false, error:'ไม่พบ CRM Sheet' };

    const sheetData = sheet.getDataRange().getValues();
    for (let i = 1; i < sheetData.length; i++) {
      if (String(sheetData[i][0] || '') === String(crmId || '')) {
        // Update fields: col index 2=name, 3=sid, 4=email, 5=phone, 11=issueType, 12=detail, 13=channel, 14=priority
        if (data.reporterName) sheet.getRange(i+1, 3).setValue(data.reporterName);
        if (data.studentId !== undefined) sheet.getRange(i+1, 4).setValue(data.studentId);
        if (data.reporterEmail !== undefined) sheet.getRange(i+1, 5).setValue(data.reporterEmail);
        if (data.reporterPhone !== undefined) sheet.getRange(i+1, 6).setValue(data.reporterPhone);
        if (data.issueType !== undefined) sheet.getRange(i+1, 12).setValue(data.issueType);
        if (data.detail) sheet.getRange(i+1, 13).setValue(data.detail);
        if (data.channel) sheet.getRange(i+1, 14).setValue(data.channel);
        if (data.priority) sheet.getRange(i+1, 15).setValue(data.priority);

        logAudit('แก้ไข CRM', crmId + ' | ' + (data.reporterName||''));
        return { success:true };
      }
    }
    return { success:false, error:'ไม่พบรายการ ' + crmId };
  } catch(e) {
    Logger.log('updateCrmTicket error: ' + e.message);
    return { success:false, error:e.message };
  }
}

function bulkImportCrmTickets(dataList) {
  if (!_autoRefreshSession()) return { success:false, error:'ไม่มีสิทธิ์' };
  try {
    if (!Array.isArray(dataList) || dataList.length === 0) {
      return { success:false, error:'ข้อมูลว่างเปล่า' };
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CRM);
    const sess = _sess();
    const recorderName = sess.valid ? sess.name||sess.email : 'ผู้แจ้งออนไลน์';
    let successCount = 0;

    for (let i = 0; i < dataList.length; i++) {
      try {
        const data = dataList[i];
        if (!data.reporterName || !data.detail) continue;

        const now = new Date();
        const id = 'CRM-' + Utilities.formatDate(now,'Asia/Bangkok','yyyyMMdd') + '-' + (sheet.getLastRow() + i);
        const assigneeName = data.assigneeName || '';
        const reportDate = data.reportDate ? fmtDate(new Date(data.reportDate)) : '';
        const row = [
          id, fmtDate(now), data.reporterName||'', data.studentId||'',
          data.reporterEmail||'', data.reporterPhone||'', data.department||'',
          data.educationLevel||'', data.term||'', data.year||'',
          data.courses||'', data.issueType||'', data.detail||'',
          data.channel||'online', data.priority||'normal',
          data.tags||'', 'open', assigneeName, '', '[]',
          recorderName, 'admin', data.org||'', reportDate,
        ];
        sheet.appendRow(row);
        successCount++;
      } catch(rowErr) {
        // skip ถ้ารายการนี้ผิดพลาด
        Logger.log('Bulk import row error: ' + rowErr.message);
      }
    }

    logAudit('นำเข้า CRM จำนวนมาก', successCount + ' รายการ');
    return { success:true, count:successCount };
  } catch(e) { return { success:false, error:e.message }; }
}

function getCrmDashboardData(filters) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const allTickets = getCrmTickets({});
    if (allTickets.error) return { success:false, error:allTickets.error };

    // Filter
    let filtered = allTickets;
    if (filters.channel) filtered = filtered.filter(t => t.channel === filters.channel);
    if (filters.issueType) filtered = filtered.filter(t => t.issueType === filters.issueType);
    if (filters.term) filtered = filtered.filter(t => t.term === filters.term);
    if (filters.priority) filtered = filtered.filter(t => t.priority === filters.priority);
    if (filters.plan) filtered = filtered.filter(t => t.plan === filters.plan);
    if (filters.course) {
      filtered = filtered.filter(t =>
        t.courses && t.courses.split(',').some(c => c.trim().toLowerCase().includes(filters.course.toLowerCase()))
      );
    }

    // Calculate stats
    const stats = {
      total: filtered.length,
      open: filtered.filter(t => t.status === 'open').length,
      inprogress: filtered.filter(t => t.status === 'inprogress').length,
      resolved: filtered.filter(t => t.status === 'resolved').length,
      closed: filtered.filter(t => t.status === 'closed').length,
    };

    return { success:true, data:filtered, stats:stats };
  } catch(e) { return { success:false, error:e.message }; }
}

function getCrmTickets(filters) {
  // session check ผ่าน _autoRefreshSession อัตโนมัติ
  try {
    // Use cache for non-filtered requests
    if (!filters || Object.keys(filters).length === 0) {
      const cache = CacheService.getUserCache();
      const cacheKey = 'crmTickets_all';
      const cached = cache.get(cacheKey);
      if (cached) return JSON.parse(cached);
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CRM);
    if (!sheet) return [];
    const lr = sheet.getLastRow();
    if (lr < 2) return [];

    const S = function(v) { return v == null ? '' : String(v); };
    const D = function(v) {
      if (!v) return '';
      try {
        if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
        return String(v);
      } catch(e) { return String(v); }
    };

    const ncols = Math.min(sheet.getLastColumn(), 24);
    const rawRows = sheet.getRange(2,1,lr-1,ncols).getValues();
    const rows = [];
    for (let i = 0; i < rawRows.length; i++) {
      const r = rawRows[i];
      if (!r[0]) continue;
      const recorderName = S(r[20]);
      const storedSource = ncols >= 22 ? S(r[21]) : '';
      const source = storedSource || (recorderName === 'ผู้แจ้งออนไลน์' || recorderName === '' ? 'external' : 'admin');
      rows.push({
        id: S(r[0]), date: D(r[1]), reporterName: S(r[2]), studentId: S(r[3]),
        reporterEmail: S(r[4]), reporterPhone: S(r[5]), department: S(r[6]),
        educationLevel: S(r[7]), term: S(r[8]), year: S(r[9]),
        courses: S(r[10]), issueType: S(r[11]), detail: S(r[12]),
        channel: S(r[13]), priority: S(r[14]), tags: S(r[15]),
        status: S(r[16]), assigneeName: S(r[17]), assigneeEmail: S(r[18]),
        replies: S(r[19]) || '[]', recorderName: recorderName,
        source: source, plan: ncols >= 23 ? S(r[22]) : '', org: ncols >= 24 ? S(r[23]) : '',
      });
    }

    // Cache unfiltered results
    if (!filters || Object.keys(filters).length === 0) {
      try {
        const cache = CacheService.getUserCache();
        cache.put('crmTickets_all', JSON.stringify(rows), 300); // 5 minute cache
      } catch(e) {
        // Ignore cache errors
      }
    }

    let result = rows;
    if (filters) {
      if (filters.status)  result = result.filter(function(r){ return r.status === filters.status; });
      if (filters.channel) result = result.filter(function(r){ return r.channel === filters.channel; });
      if (filters.source)  result = result.filter(function(r){ return r.source === filters.source; });
      if (filters.org)     result = result.filter(function(r){ return r.org === filters.org; });
      if (filters.plan)    result = result.filter(function(r){ return r.plan === filters.plan; });
      if (filters.issueType) result = result.filter(function(r){ return r.issueType === filters.issueType; });
      if (filters.term)    result = result.filter(function(r){ return r.term === filters.term; });
      if (filters.priority) result = result.filter(function(r){ return r.priority === filters.priority; });
      if (filters.course)  result = result.filter(function(r){ return (r.courses||'').indexOf(filters.course) !== -1; });
      if (filters.dateFrom || filters.dateTo) {
        function parseFilterDate(s) {
          if (!s) return null;
          var m = String(s).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
          if (m) return new Date(parseInt(m[3]), parseInt(m[2])-1, parseInt(m[1]));
          return new Date(s);
        }
        var dFrom = parseFilterDate(filters.dateFrom);
        var dTo   = parseFilterDate(filters.dateTo);
        if (dTo) dTo.setHours(23,59,59,999);
        result = result.filter(function(r) {
          var d = r.date ? new Date(r.date) : null;
          if (!d || isNaN(d.getTime())) {
            // ลอง parse dd/MM/yyyy
            var m2 = String(r.date||'').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
            if (m2) d = new Date(parseInt(m2[3]), parseInt(m2[2])-1, parseInt(m2[1]));
          }
          if (!d || isNaN(d.getTime())) return true; // ถ้า parse ไม่ได้ให้ผ่าน
          if (dFrom && !isNaN(dFrom.getTime()) && d < dFrom) return false;
          if (dTo   && !isNaN(dTo.getTime())   && d > dTo)   return false;
          return true;
        });
      }
      if (filters.search) {
        const q = String(filters.search).toLowerCase();
        result = result.filter(function(r) {
          return [r.reporterName, r.studentId, r.detail, r.issueType, r.department, r.org]
            .some(function(v){ return v && v.toLowerCase().indexOf(q) !== -1; });
        });
      }
    }
    return result.reverse();
  } catch(e) { return { error: 'getCrmTickets: ' + e.message }; }
}

function assignCrmTicket(id, adminName) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  return _updateCrmField(id, 18, adminName);
}

function replyCrmTicket(id, text, adminName, adminEmail, sendEmail) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sess = _sess();
    const fromEmail = adminEmail || (sess && sess.email) || '';
    const fromName  = adminName  || (sess && sess.name)  || 'เจ้าหน้าที่';

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CRM);
    const data  = sheet.getDataRange().getValues();
    for (let i=1;i<data.length;i++) {
      if (data[i][0]===id) {
        const replies = JSON.parse(data[i][19]||'[]');
        replies.push({name:fromName, email:fromEmail, text:text, time:fmtDate(new Date())});
        sheet.getRange(i+1,20).setValue(JSON.stringify(replies));
        // อัปเดตสถานะเป็น inprogress ถ้ายังเป็น open
        if (data[i][16]==='open') sheet.getRange(i+1,17).setValue('inprogress');

        // ส่งอีเมลไปหา นศ./ผู้แจ้ง (ถ้า sendEmail === true)
        const reporterEmail = (data[i][4]||'').toString().trim();
        const reporterName  = (data[i][2]||'').toString().trim() || 'ผู้แจ้ง';
        const issueType     = (data[i][11]||'').toString().trim() || 'แจ้งปัญหา';

        let emailSent = false;
        let emailError = '';
        if (sendEmail && reporterEmail && /\S+@\S+\.\S+/.test(reporterEmail)) {
          try {
            const subject = '[มสธ.] ตอบกลับเรื่องที่ท่านแจ้ง: ' + issueType + ' (' + id + ')';
            const htmlBody = ''
              + '<div style="font-family:Sarabun,Arial,sans-serif;max-width:600px;margin:0 auto;background:#f4f6fb;padding:20px">'
              +   '<div style="background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 16px rgba(0,0,0,0.08)">'
              +     '<div style="background:linear-gradient(135deg,#0f2744,#1a4a8a);padding:24px;text-align:center;color:#fff">'
              +       '<div style="font-size:32px;margin-bottom:6px">📚</div>'
              +       '<h2 style="margin:0;font-size:18px">มหาวิทยาลัยสุโขทัยธรรมาธิราช</h2>'
              +       '<p style="margin:4px 0 0;font-size:13px;color:rgba(255,255,255,0.8)">สำนักบริการการศึกษา</p>'
              +     '</div>'
              +     '<div style="padding:24px">'
              +       '<p style="font-size:15px;color:#1a2840">เรียน ' + _esc(reporterName) + '</p>'
              +       '<p style="color:#1a2840">เจ้าหน้าที่ได้ตอบกลับเรื่องที่ท่านแจ้งไว้ดังนี้</p>'
              +       '<div style="background:#f0f4f8;border-left:4px solid #2d7dd2;padding:12px 16px;margin:16px 0;border-radius:4px">'
              +         '<div style="font-size:12px;color:#8899b4;margin-bottom:4px">รหัสเรื่อง: <strong>' + _esc(id) + '</strong></div>'
              +         '<div style="font-size:12px;color:#8899b4;margin-bottom:4px">ประเภท: <strong>' + _esc(issueType) + '</strong></div>'
              +         '<div style="font-size:12px;color:#8899b4">ผู้ตอบ: <strong>' + _esc(fromName) + '</strong></div>'
              +       '</div>'
              +       '<div style="background:#fff8e1;border-left:4px solid #f4a21e;padding:14px 18px;border-radius:4px;margin-bottom:16px">'
              +         '<div style="font-size:13px;color:#5c4b00;line-height:1.7;white-space:pre-wrap">' + _esc(text) + '</div>'
              +       '</div>'
              +       '<p style="color:#1a2840;font-size:13px">หากต้องการสอบถามเพิ่มเติม กรุณาตอบกลับอีเมลฉบับนี้ หรือโทร 02 504 7623, 7626</p>'
              +     '</div>'
              +     '<div style="background:#f4f6fb;padding:16px 24px;text-align:center;border-top:1px solid #e0e8f0">'
              +       '<p style="margin:0;font-size:11px;color:#8899b4">ขอแสดงความนับถือ<br>สำนักบริการการศึกษา มสธ.</p>'
              +     '</div>'
              +   '</div>'
              + '</div>';

            const textBody = 'เรียน ' + reporterName + '\n\n'
              + 'เจ้าหน้าที่ได้ตอบกลับเรื่องที่ท่านแจ้ง (' + id + ') ดังนี้\n\n'
              + 'ประเภท: ' + issueType + '\n'
              + 'ผู้ตอบ: ' + fromName + '\n\n'
              + '----------------------------------------\n'
              + text + '\n'
              + '----------------------------------------\n\n'
              + 'หากต้องการสอบถามเพิ่มเติม กรุณาตอบกลับอีเมลฉบับนี้ หรือโทร 02 504 7623, 7626\n\n'
              + 'ขอแสดงความนับถือ\n'
              + 'สำนักบริการการศึกษา มสธ.';

            const opts = {
              htmlBody: htmlBody,
              name: fromName + ' — สำนักบริการการศึกษา มสธ.'
            };
            if (fromEmail && /\S+@\S+\.\S+/.test(fromEmail)) opts.replyTo = fromEmail;

            GmailApp.sendEmail(reporterEmail, subject, textBody, opts);
            emailSent = true;
          } catch(mailErr) {
            emailError = mailErr.message;
          }
        } else {
          emailError = 'ไม่พบอีเมลผู้แจ้ง';
        }

        // บันทึก audit log
        logAudit('ตอบกลับ CRM', id + ' | ' + reporterName + ' | ' + (text||'').substring(0,80) + (emailSent ? ' | ส่งอีเมลแล้ว' : ''));

        return {
          success: true,
          emailSent: emailSent,
          emailError: emailError,
          sentTo: reporterEmail,
          sentFrom: fromEmail
        };
      }
    }
    return { success:false, error:'ไม่พบ Ticket' };
  } catch(e) { return { success:false, error:e.message }; }
}

// helper สำหรับ escape HTML
function _esc(s) {
  if (s == null) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function getAddressFromRecord(studentId, courseCode) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_DATA);
    if (!sheet) return { success:false, address:'' };

    const data = sheet.getDataRange().getValues();
    const courseStr = String(courseCode || '').trim();
    const studentStr = String(studentId || '').trim();

    for (let i = 1; i < data.length; i++) {
      const rowCourse = String(data[i][6] || '').trim();
      const rowStudent = String(data[i][7] || '').trim();

      if (rowCourse === courseStr && rowStudent === studentStr) {
        // Found matching record - construct address from components
        const houseNo = String(data[i][11] || '').trim();
        const street = String(data[i][12] || '').trim();
        const subDistrict = String(data[i][13] || '').trim();
        const district = String(data[i][14] || '').trim();
        const province = String(data[i][15] || '').trim();
        const zipCode = String(data[i][16] || '').trim();

        const addressParts = [houseNo, street, subDistrict, district, province, zipCode].filter(p => p);
        const fullAddress = addressParts.join(' ');

        return { success:true, address:fullAddress };
      }
    }

    return { success:false, address:'' };
  } catch(e) {
    Logger.log('getAddressFromRecord error: ' + e.message);
    return { success:false, address:'' };
  }
}

function getInvestigationByRefId(refId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_INVEST);
    if (!sheet) return { success:false, investigationId:null, rowIndex:-1 };

    const data = sheet.getDataRange().getValues();
    const refIdStr = String(refId || '').trim();

    for (let i = 1; i < data.length; i++) {
      const rowRefId = String(data[i][1] || '').trim();
      if (rowRefId === refIdStr) {
        return { success:true, investigationId:String(data[i][0] || ''), rowIndex:i };
      }
    }

    return { success:false, investigationId:null, rowIndex:-1 };
  } catch(e) {
    Logger.log('getInvestigationByRefId error: ' + e.message);
    return { success:false, investigationId:null, rowIndex:-1 };
  }
}

function updateRecordParcel(recordId, send1Track, send1Date, send2Track, send2Date, send3Track, send3Date) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_DATA);
    if (!sheet) return { success:false, error:'ไม่พบ Sheet ข้อมูลพัสดุ' };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === String(recordId || '').trim()) {
        // Found the record
        // Column indices (1-based for getRange): 21=send1Track, 22=send1Date, 23=send2Track, 24=send2Date, 25=send3Track, 26=send3Date
        if (send1Track) sheet.getRange(i + 1, 21).setValue(send1Track);
        if (send1Date) sheet.getRange(i + 1, 22).setValue(send1Date);
        if (send2Track) sheet.getRange(i + 1, 23).setValue(send2Track);
        if (send2Date) sheet.getRange(i + 1, 24).setValue(send2Date);
        if (send3Track) sheet.getRange(i + 1, 25).setValue(send3Track);
        if (send3Date) sheet.getRange(i + 1, 26).setValue(send3Date);

        // Update timestamp
        sheet.getRange(i + 1, 29).setValue(new Date());

        logAudit('แก้ไขเลขพัสดุ', recordId + ' | send1: ' + send1Track + ' | send2: ' + send2Track + ' | send3: ' + send3Track);
        return { success:true, message:'อัปเดตเลขพัสดุแล้ว' };
      }
    }

    return { success:false, error:'ไม่พบรายการ ' + recordId };
  } catch(e) {
    Logger.log('updateRecordParcel error: ' + e.message);
    return { success:false, error:e.message };
  }
}

function updateRecordCourseParcels(recordId, coursesJson) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_DATA);
    if (!sheet) return { success:false, error:'ไม่พบ Sheet ข้อมูลพัสดุ' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === String(recordId || '').trim()) {
        let courses = [];
        try { courses = JSON.parse(coursesJson || '[]'); } catch(e) {}
        // Save courses JSON to column 27 (index 26)
        sheet.getRange(i + 1, 27).setValue(JSON.stringify(courses));
        // Sync send1/send2 columns from first two courses
        if (courses[0]) {
          sheet.getRange(i + 1, 21).setValue(courses[0].track || '');
          sheet.getRange(i + 1, 22).setValue(courses[0].date || '');
        }
        if (courses[1]) {
          sheet.getRange(i + 1, 23).setValue(courses[1].track || '');
          sheet.getRange(i + 1, 24).setValue(courses[1].date || '');
        }
        sheet.getRange(i + 1, 29).setValue(fmtDate(new Date()));
        logAudit('แก้ไขเลขพัสดุ', recordId + ' | ' + (coursesJson || '').substring(0, 80));
        return { success:true };
      }
    }
    return { success:false, error:'ไม่พบรายการ ' + recordId };
  } catch(e) {
    Logger.log('updateRecordCourseParcels error: ' + e.message);
    return { success:false, error:e.message };
  }
}

function updateCrmStatus(id, status) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  return _updateCrmField(id, 17, status);
}

function _updateCrmField(id, col, val) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CRM);
    const data  = sheet.getDataRange().getValues();
    for (let i=1;i<data.length;i++) {
      if (data[i][0]===id) { sheet.getRange(i+1,col).setValue(val); return { success:true }; }
    }
    return { success:false, error:'ไม่พบรายการ' };
  } catch(e) { return { success:false, error:e.message }; }
}

function saveFollowUpRecord(data) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sess = _sess();
    const recorderName = sess.name || sess.email || 'เจ้าหน้าที่';

    // Update CRM Sheet
    const crmSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CRM);
    if (!crmSheet) return { success:false, error:'ไม่พบ CRM Sheet' };

    const crmData = crmSheet.getDataRange().getValues();
    let crmRowIndex = -1;
    for (let i = 1; i < crmData.length; i++) {
      if (crmData[i][0] === data.crmId) {
        crmRowIndex = i;
        break;
      }
    }

    if (crmRowIndex === -1) {
      return { success:false, error:'ไม่พบ CRM ID: ' + data.crmId };
    }

    // Add follow-up info to CRM (columns 25-29: followUpType, followUpCause, followUpDate, followUpNotes, followUpStatus)
    const followUpDate = fmtDate(new Date());
    const followUpStatus = 'บันทึกแล้ว';

    crmSheet.getRange(crmRowIndex + 1, 25).setValue(data.type || '');
    crmSheet.getRange(crmRowIndex + 1, 26).setValue(data.cause || '');
    crmSheet.getRange(crmRowIndex + 1, 27).setValue(followUpDate);
    crmSheet.getRange(crmRowIndex + 1, 28).setValue(data.notes || '');
    crmSheet.getRange(crmRowIndex + 1, 29).setValue(followUpStatus);

    // Update/Create record in ข้อมูลพัสดุ Sheet
    const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_DATA);
    if (!dataSheet) return { success:false, error:'ไม่พบ ข้อมูลพัสดุ Sheet' };

    const dataRows = dataSheet.getDataRange().getValues();
    let existingRowIndex = -1;

    // Find existing row by course + studentId
    const searchCourse = String(data.course || '').trim();
    const searchStudentId = String(data.studentId || '').trim();

    for (let i = 1; i < dataRows.length; i++) {
      const rowCourse = String(dataRows[i][6] || '').trim();
      const rowStudentId = String(dataRows[i][7] || '').trim();
      if (rowCourse === searchCourse && rowStudentId === searchStudentId) {
        existingRowIndex = i;
        break;
      }
    }

    let recordId = null;

    if (existingRowIndex !== -1) {
      // UPDATE existing row
      recordId = String(dataRows[existingRowIndex][0] || '');

      // Update courses JSON with new parcel data for the matching course
      if (data.parcelTrack || data.parcelDate) {
        let courses = [];
        try { courses = JSON.parse(String(dataRows[existingRowIndex][26] || '[]')); } catch(e) {}
        if (!Array.isArray(courses) || courses.length === 0) {
          // Build from courseCode
          const cc = String(dataRows[existingRowIndex][6] || '');
          courses = cc.split(',').map(function(s) { return { code: s.trim() }; }).filter(function(c) { return c.code; });
          if (courses.length === 0) courses = [{ code: cc }];
        }
        // Find course index matching data.course, or use first
        let courseIdx = courses.findIndex(function(c) { return (c.code || c) === (data.course||''); });
        if (courseIdx === -1) courseIdx = 0;
        courses[courseIdx] = Object.assign({}, courses[courseIdx], {
          track: data.parcelTrack || '',
          date:  data.parcelDate  || ''
        });
        dataSheet.getRange(existingRowIndex + 1, 27).setValue(JSON.stringify(courses));

        // Sync send1/send2 columns from first two courses
        if (courses[0]) { dataSheet.getRange(existingRowIndex + 1, 21).setValue(courses[0].track || ''); dataSheet.getRange(existingRowIndex + 1, 22).setValue(courses[0].date || ''); }
        if (courses[1]) { dataSheet.getRange(existingRowIndex + 1, 23).setValue(courses[1].track || ''); dataSheet.getRange(existingRowIndex + 1, 24).setValue(courses[1].date || ''); }
      }

      // Update status and timestamp (columns 27, 28, 29)
      dataSheet.getRange(existingRowIndex + 1, 28).setValue('ส่งแล้ว');
      dataSheet.getRange(existingRowIndex + 1, 29).setValue(new Date());

      logAudit('บันทึก Follow-up (อัปเดต)', data.crmId + ' | ' + data.type + ' | นศ.' + data.studentId);
    } else {
      // CREATE new row in ข้อมูลพัสดุ
      const now = new Date();
      const pfx = {return:'P', loan:'L', special_resend:'S'}[data.type] || 'P';
      recordId = pfx + Utilities.formatDate(now, 'Asia/Bangkok', 'yyyyMMdd') + '-' + (dataSheet.getLastRow() + 1);

      // Embed parcel data in courses JSON
      const coursesArr = [{ code: data.course || '', track: data.parcelTrack || '', date: data.parcelDate || '' }];

      // 30 columns: 0=id,1=date,2=term,3=year,4=recType,5=parcelType,6=courseCode,7=studentId,
      // 8=prefix,9=firstName,10=lastName,11=houseNo,12=street,13=subDistrict,14=district,
      // 15=province,16=zipCode,17=phone,18=cause,19=contactStatus,
      // 20=send1Track,21=send1Date,22=send2Track,23=send2Date,
      // 24=tags,25=remark,26=courses,27=status,28=updatedAt,29=recorder
      const newRow = [
        recordId, fmtDate(now), '', '', data.type, '',
        data.course || '', data.studentId || '', '',
        '', '',
        '', '', '', '',
        '', '', data.phone || '',
        data.cause || '', '',
        data.parcelTrack || '', data.parcelDate || '',
        '', '',
        '', '', JSON.stringify(coursesArr), 'ส่งแล้ว', fmtDate(now), recorderName
      ];

      dataSheet.appendRow(newRow);
      const lr = dataSheet.getLastRow();
      if (lr % 2 === 0) dataSheet.getRange(lr, 1, 1, newRow.length).setBackground('#f0f4f8');

      logAudit('บันทึก Follow-up (สร้างใหม่)', data.crmId + ' | ' + data.type + ' | นศ.' + data.studentId);
    }

    // For loan type, also create investigation record
    if (data.type === 'loan' && recordId) {
      try {
        const investCheck = getInvestigationByRefId(recordId);
        if (!investCheck.success) {
          // Investigation record doesn't exist - create one
          const addResult = addInvestigation({
            refId: recordId,
            itemNo: recordId,
            barcode: '',
            sentDate: '',
            courseCode: data.course || '',
            courseName: '',
            weight: '',
            fee: '',
            recipientName: data.name || '',
            recipientAddr: data.address || '',
            cause: data.cause || ''
          });

          if (addResult.success) {
            Logger.log('Created investigation record for loan follow-up: ' + addResult.id);
          }
        }
      } catch(investErr) {
        Logger.log('Error creating investigation for loan: ' + investErr.message);
      }
    }

    return { success:true, message:'บันทึก follow-up แล้ว' };
  } catch(e) {
    Logger.log('saveFollowUpRecord error: ' + e.message);
    return { success:false, error:e.message };
  }
}

// ============================================================
// USER MANAGEMENT
// ============================================================
function getUsers() {
  if (!_autoRefreshSession()) return { error:'ไม่มีสิทธิ์' };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_USERS);
    if (!sheet) return [];
    return sheet.getDataRange().getValues().slice(1)
      .map((r,i)=>({row:i+2,email:r[0],role:r[2],name:r[3],active:r[4]===true||r[4]==='TRUE',lastLogin:r[5]?fmtDate(r[5]):'-',perms:r[6]||''}))
      .filter(u=>u.email!=='');
  } catch(e) { return { error:e.message }; }
}

function addUser(email, password, role, name, perms) {
  if (!_autoRefreshSession()) return { success:false, error:'ไม่มีสิทธิ์' };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_USERS);
    const data  = sheet.getDataRange().getValues();
    for (let i=1;i<data.length;i++) {
      if ((data[i][0]||'').toLowerCase()===email.toLowerCase()) return { success:false, error:'Email นี้มีอยู่แล้ว' };
    }
    sheet.appendRow([email.toLowerCase(), hashPw(password), role, name, true, '', perms||'', '', '']);
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

function inviteUser(email, role, name, perms) {
  if (!_autoRefreshSession()) return { success:false, error:'ไม่มีสิทธิ์' };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_USERS);
    const data  = sheet.getDataRange().getValues();
    for (let i=1;i<data.length;i++) {
      if ((data[i][0]||'').toLowerCase()===email.toLowerCase()) return { success:false, error:'Email นี้มีอยู่แล้ว' };
    }
    const token = Utilities.getUuid();
    // columns: email, hashpw(empty), role, name, active(false until pw set), lastLogin, perms, invite_token, invite_expires
    const expires = new Date(); expires.setHours(expires.getHours()+72);
    sheet.appendRow([email.toLowerCase(), '', role, name, false, '', perms||'', token, fmtDate(expires)]);
    const scriptUrl = ScriptApp.getService().getUrl();
    const setpwUrl = scriptUrl + '?page=setpw&token=' + token + '&email=' + encodeURIComponent(email);
    logAudit('เชิญผู้ใช้', email + ' | ' + role);
    try {
      const subject = '[มสธ.] คำเชิญเข้าใช้งานระบบจัดการเอกสาร มสธ.';
      const body = 'เรียน ' + name + '\n\n'
        + 'ท่านได้รับสิทธิ์เข้าใช้งานระบบจัดการและติดตามเอกสารการสอน มสธ. ในบทบาท: ' + role + '\n\n'
        + 'กรุณาคลิกลิงก์ด้านล่างเพื่อตั้งรหัสผ่าน (ลิงก์ใช้ได้ 72 ชม.):\n\n'
        + setpwUrl + '\n\n'
        + 'หากท่านไม่ได้รับคำเชิญนี้ กรุณาเพิกเฉยต่ออีเมลนี้\n\n'
        + 'ขอแสดงความนับถือ\n'
        + 'ผู้ดูแลระบบ มสธ.';
      try {
        GmailApp.sendEmail(email, subject, body);
        return { success:true, emailSent:true };
      } catch(gmailErr) {
        // Fallback to MailApp if GmailApp fails
        MailApp.sendEmail(email, subject, body);
        return { success:true, emailSent:true };
      }
    } catch(emailErr) {
      return { success:true, emailSent:false, setupUrl:setpwUrl, message:'ไม่สามารถส่งอีเมลได้ กรุณาส่งลิงก์ด้านล่างให้ผู้ใช้: ' + setpwUrl };
    }
  } catch(e) { return { success:false, error:e.message }; }
}

function resendUserInvite(email) {
  if (!_autoRefreshSession()) return { success:false, error:'ไม่มีสิทธิ์' };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_USERS);
    const data  = sheet.getDataRange().getValues();
    for (let i=1;i<data.length;i++) {
      if ((data[i][0]||'').toLowerCase()===email.toLowerCase()) {
        const token = Utilities.getUuid();
        const expires = new Date(); expires.setHours(expires.getHours()+72);
        // ensure columns exist
        while (sheet.getLastColumn() < 9) sheet.getRange(1, sheet.getLastColumn()+1).setValue('');
        sheet.getRange(i+1, 8).setValue(token);
        sheet.getRange(i+1, 9).setValue(fmtDate(expires));
        const scriptUrl = ScriptApp.getService().getUrl();
        const setpwUrl = scriptUrl + '?page=setpw&token=' + token + '&email=' + encodeURIComponent(email);
        const name = data[i][3] || email;
        try {
          const subject = '[มสธ.] ลิงก์ตั้งรหัสผ่านใหม่ — ระบบจัดการเอกสาร มสธ.';
          const body = 'เรียน ' + name + '\n\n'
            + 'กรุณาคลิกลิงก์ด้านล่างเพื่อตั้งรหัสผ่านใหม่ (ลิงก์ใช้ได้ 72 ชม.):\n\n'
            + setpwUrl + '\n\n'
            + 'หากท่านไม่ได้ร้องขอ กรุณาเพิกเฉยต่ออีเมลนี้\n\n'
            + 'ขอแสดงความนับถือ\n'
            + 'ผู้ดูแลระบบ มสธ.';
          try {
            GmailApp.sendEmail(email, subject, body);
            return { success:true, emailSent:true };
          } catch(gmailErr) {
            MailApp.sendEmail(email, subject, body);
            return { success:true, emailSent:true };
          }
        } catch(emailErr) {
          return { success:true, emailSent:false, setupUrl:setpwUrl, message:'ไม่สามารถส่งอีเมลได้ กรุณาส่งลิงก์ด้านล่างให้ผู้ใช้: ' + setpwUrl };
        }
      }
    }
    return { success:false, error:'ไม่พบ Email' };
  } catch(e) { return { success:false, error:e.message }; }
}

function confirmInvitePassword(email, token, password) {
  try {
    if (!email || !token || !password) return { success:false, error:'ข้อมูลไม่ครบ' };
    if (password.length < 8) return { success:false, error:'รหัสผ่านต้องมีอย่างน้อย 8 ตัวอักษร' };
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_USERS);
    const data  = sheet.getDataRange().getValues();
    for (let i=1;i<data.length;i++) {
      if ((data[i][0]||'').toLowerCase()===email.toLowerCase()) {
        const storedToken = (data[i][7]||'').toString().trim();
        if (!storedToken || storedToken !== token) return { success:false, error:'ลิงก์ไม่ถูกต้องหรือหมดอายุแล้ว' };
        // check expiry (col 9, index 8)
        const expiry = data[i][8];
        if (expiry && new Date(expiry) < new Date()) return { success:false, error:'ลิงก์หมดอายุแล้ว กรุณาขอคำเชิญใหม่' };
        sheet.getRange(i+1, 2).setValue(hashPw(password)); // set hash
        sheet.getRange(i+1, 5).setValue(true);             // activate
        sheet.getRange(i+1, 8).setValue('');               // clear token
        sheet.getRange(i+1, 9).setValue('');               // clear expiry
        return { success:true };
      }
    }
    return { success:false, error:'ไม่พบบัญชีนี้' };
  } catch(e) { return { success:false, error:e.message }; }
}

function updateUserPassword(email, newPassword) {
  if (!_autoRefreshSession()) return { success:false, error:'ไม่มีสิทธิ์' };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_USERS);
    const data  = sheet.getDataRange().getValues();
    for (let i=1;i<data.length;i++) {
      if ((data[i][0]||'').toLowerCase()===email.toLowerCase()) {
        sheet.getRange(i+1,2).setValue(hashPw(newPassword));
        return { success:true };
      }
    }
    return { success:false, error:'ไม่พบ Email' };
  } catch(e) { return { success:false, error:e.message }; }
}

function toggleUserActive(email, active) {
  if (!_autoRefreshSession()) return { success:false, error:'ไม่มีสิทธิ์' };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_USERS);
    const data  = sheet.getDataRange().getValues();
    for (let i=1;i<data.length;i++) {
      if ((data[i][0]||'').toLowerCase()===email.toLowerCase()) {
        sheet.getRange(i+1,5).setValue(active);
        return { success:true };
      }
    }
    return { success:false, error:'ไม่พบ Email' };
  } catch(e) { return { success:false, error:e.message }; }
}

function updateUserRole(email, role, perms) {
  if (!_autoRefreshSession()) return { success:false, error:'ไม่มีสิทธิ์' };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_USERS);
    const data  = sheet.getDataRange().getValues();
    for (let i=1;i<data.length;i++) {
      if ((data[i][0]||'').toLowerCase()===email.toLowerCase()) {
        sheet.getRange(i+1,3).setValue(role);
        sheet.getRange(i+1,7).setValue(perms);
        logAudit('อัปเดต role/permissions', email + ' | ' + role);
        return { success:true };
      }
    }
    return { success:false, error:'ไม่พบ Email' };
  } catch(e) { return { success:false, error:e.message }; }
}

// ============================================================
// SETUP
// ============================================================
function setupSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Sheet: ข้อมูลพัสดุ (29 คอลัมน์)
  let ds = ss.getSheetByName(SH_DATA);
  if (!ds) ds = ss.insertSheet(SH_DATA);
  const dHdrs = ['รหัส','วันที่บันทึก','ภาค','ปี','ประเภทรายการ','ประเภทพัสดุ','รหัสชุดวิชา(หลัก)','รหัสนักศึกษา','คำนำหน้า','ชื่อ','นามสกุล','บ้านเลขที่','ถนน','ตำบล/แขวง','อำเภอ/เขต','จังหวัด','รหัสไปรษณีย์','เบอร์โทร','สาเหตุ','สถานะติดต่อ','เลขพัสดุส่งครั้งที่ 1','วันที่ส่งครั้งที่ 1','เลขพัสดุส่งครั้งที่ 2','วันที่ส่งครั้งที่ 2','Tags','หมายเหตุ','ชุดวิชา(JSON)','สถานะ','วันที่อัปเดต','ผู้บันทึก'];
  if (ds.getLastRow()===0) {
    ds.appendRow(dHdrs);
    const hr=ds.getRange(1,1,1,dHdrs.length);
    hr.setBackground('#1a3a5c');hr.setFontColor('#fff');hr.setFontWeight('bold');
    ds.setFrozenRows(1);
  }

  // Sheet: ตั้งค่า
  let st = ss.getSheetByName(SH_SETTINGS);
  if (!st) {
    st = ss.insertSheet(SH_SETTINGS);
    st.appendRow(['รายการ','ค่า']);
    st.appendRow(['ประเภทพัสดุ','จดหมาย,พัสดุธรรมดา,EMS,เอกสารราชการ,อื่นๆ']);
    st.appendRow(['สาเหตุตีคืน','จ่าหน้าไม่ชัดเจน,ไม่มีเลขที่บ้านตามจ่าหน้า,ไม่ยอมรับ,ไม่มีผู้รับตามจ่าหน้า,ไม่มารับภายในกำหนด,ไม่มีจ่าหน้าหรือจ่าหน้าสูญหาย,ย้ายไม่ทราบที่อยู่ใหม่,อื่นๆ']);
    st.appendRow(['คำนำหน้า','นาย,นาง,นางสาว,ว่าที่ร้อยตรี,อื่นๆ']);
    st.appendRow(['ภาคการศึกษา','1,2,3']);
    st.appendRow(['สาขาวิชา','สาขาศิลปศาสตร์,สาขาศึกษาศาสตร์,สาขาวิทยาการจัดการ,สาขานิติศาสตร์,สาขาวิทยาศาสตร์สุขภาพ,สาขาเศรษฐศาสตร์,สาขามนุษยนิเวศศาสตร์,สาขารัฐศาสตร์,สาขาเกษตรศาสตร์และสหกรณ์,สาขานิเทศศาสตร์,สาขาวิทยาศาสตร์และเทคโนโลยี,สาขาพยาบาลศาสตร์']);
    st.appendRow(['แผนการศึกษา','แผน ก1,แผน ก2,แผน ก3']);
    st.appendRow(['ประเภทปัญหา','พิมพ์เพิ่ม,ชุดปรับปรุง,ชุดผลิตใหม่,ปัญหาทวงถามหนังสือ (ไม่ได้สั่งซื้อ/แผน ก2-ก3),สอบถามทะเบียนและวัดผล (ลงทะเบียน/สอบ),สอบถามกิจกรรมประจำชุดวิชา,สอบถามอบรมเข้มเสริมประสบการณ์วิชาชีพ,สอบถามสอนเสริมออนไลน์,สอบถามโครงการสัมฤทธิบัตร,ไม่ได้รับเอกสาร,เอกสารชำรุด,ส่งผิดวิชา,อื่นๆ (ระบุเอง)']);
    st.appendRow(['ผู้รับเรื่อง','หทัย เริงเกษตรกิจ,วรรณี รัตนากร,สุพรรษา ช่อปทุมมา,เมธิตา สาไพรวัน']);
    const sh=st.getRange(1,1,1,2);sh.setBackground('#1a3a5c');sh.setFontColor('#fff');sh.setFontWeight('bold');
  }

  // Sheet: ผู้ใช้งาน (เพิ่มคอลัมน์ perms, token, expires สำหรับ invitation flow)
  let us = ss.getSheetByName(SH_USERS);
  if (!us) {
    us = ss.insertSheet(SH_USERS);
    us.appendRow(['email','password_hash','role','ชื่อ-สกุล','active','last_login','perms','invite_token','invite_expires']);
    const uh=us.getRange(1,1,1,9);uh.setBackground('#1a3a5c');uh.setFontColor('#fff');uh.setFontWeight('bold');
    us.appendRow(['admin@stou.ac.th',hashPw('admin1234'),'superadmin','ผู้ดูแลระบบ',true,'','dash,search,rec-return,rec-lend,rec-special,list,labels,crm,tags,users','','']);
    [220,200,100,160,70,150,300,250,200].forEach((w,i)=>us.setColumnWidth(i+1,w));
  } else {
    // If sheet exists but doesn't have invite_token column, add them
    const lastCol = us.getLastColumn();
    if (lastCol < 8) {
      us.getRange(1, 8).setValue('invite_token');
      us.getRange(1, 9).setValue('invite_expires');
      us.getRange(1, 8, 1, 2).setBackground('#1a3a5c').setFontColor('#fff').setFontWeight('bold');
    }
  }

  // Sheet: Tags
  let tg = ss.getSheetByName(SH_TAGS);
  if (!tg) {
    tg = ss.insertSheet(SH_TAGS);
    tg.appendRow(['name','color']);
    const th=tg.getRange(1,1,1,2);th.setBackground('#1a3a5c');th.setFontColor('#fff');th.setFontWeight('bold');
    [['ปัญหาซ้ำ',3],['ด่วน',4],['วิชาบังคับ',1],['ติดตามแล้ว',5],['สำคัญ',2]].forEach(row=>tg.appendRow(row));
  }

  // Sheet: CRM (22 คอลัมน์)
  let cm = ss.getSheetByName(SH_CRM);
  if (!cm) {
    cm = ss.insertSheet(SH_CRM);
    // col: 1-21=ข้อมูล, 22=แหล่งที่มา(source), 23=แผนการศึกษา(plan), 24=หน่วยงาน(org), 25=วันที่แจ้ง(reportDate)
    cm.appendRow(['รหัส','วันที่','ชื่อผู้แจ้ง','รหัสนักศึกษา','อีเมล','เบอร์โทร','หน่วยงาน/สาขา','ระดับการศึกษา','ภาค','ปี','ชุดวิชา','ประเภทปัญหา','รายละเอียด','ช่องทาง','ความเร่งด่วน','Tags','สถานะ','ผู้รับเรื่อง','อีเมลผู้รับเรื่อง','ประวัติการตอบ','ผู้บันทึก','แหล่งที่มา','แผนการศึกษา','หน่วยงานภายนอก','วันที่แจ้งปัญหา']);
    const ch=cm.getRange(1,1,1,25);ch.setBackground('#1a3a5c');ch.setFontColor('#fff');ch.setFontWeight('bold');
    cm.setFrozenRows(1);
  } else {
    // Migrate existing sheet: ensure correct column layout
    const headers = cm.getRange(1,1,1,cm.getLastColumn()).getValues()[0];
    // Col 22 (idx 21) = แหล่งที่มา (source), Col 23 (idx 22) = แผนการศึกษา, Col 24 (idx 23) = หน่วยงานภายนอก
    if (!headers[21] || headers[21] === 'แผนการศึกษา') {
      cm.getRange(1, 22).setValue('แหล่งที่มา');
      cm.getRange(1, 22).setBackground('#1a3a5c').setFontColor('#fff').setFontWeight('bold');
    }
    if (!headers[22] || headers[22] !== 'แผนการศึกษา') {
      cm.getRange(1, 23).setValue('แผนการศึกษา');
      cm.getRange(1, 23).setBackground('#1a3a5c').setFontColor('#fff').setFontWeight('bold');
    }
    if (!headers[23]) {
      cm.getRange(1, 24).setValue('หน่วยงานภายนอก');
      cm.getRange(1, 24).setBackground('#1a3a5c').setFontColor('#fff').setFontWeight('bold');
    }
    if (!headers[24]) {
      cm.getRange(1, 25).setValue('วันที่แจ้งปัญหา');
      cm.getRange(1, 25).setBackground('#1a3a5c').setFontColor('#fff').setFontWeight('bold');
    }
  }

  Logger.log('ALERT: '+
    'ตั้งค่าระบบสำเร็จ! ✅\n\n' +
    'บัญชี Super Admin:\n' +
    'Email: admin@stou.ac.th\n' +
    'Password: admin1234\n\n' +
    '⚠️ กรุณาเปลี่ยน Password ทันทีหลัง Login'
  );
}

// backward compat
function setupSpreadsheet() { setupSystem(); }

// ============================================================
// HELPERS
// ============================================================
function fmtDate(d) {
  if (!d) return '';
  try {
    return Utilities.formatDate(new Date(d), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
  } catch(e) { return d.toString(); }
}

// ============================================================
// POSTAL INVESTIGATION — ระบบสอบสวนไปรษณีย์ไทย
// ============================================================

function setupInvestSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SH_INVEST);
  if (!sh) {
    sh = ss.insertSheet(SH_INVEST);
    const hdrs = [
      'รหัสสอบสวน','รหัสพัสดุอ้างอิง','วันที่สร้าง',
      'เลขที่สิ่งของ','เลข Barcode','วันที่ฝากส่ง',
      'รหัสชุดวิชา','ชื่อชุดวิชา','น้ำหนัก','ค่าบริการ',
      'ชื่อ-นามสกุลผู้รับ','ที่อยู่ผู้รับ',
      'สาเหตุปัญหา','วันที่แจ้งไปรษณีย์','อีเมลที่แจ้ง','จำนวนรายการในอีเมล',
      'วันที่ได้รับผลตอบ','จำนวนวันตอบ','ผลสอบสวน','รายละเอียดผล',
      'สถานะ','ผู้บันทึก','วันที่อัปเดต'
    ];
    sh.appendRow(hdrs);
    const hr = sh.getRange(1,1,1,hdrs.length);
    hr.setBackground('#1a3a5c');hr.setFontColor('#fff');hr.setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function getInvestigations(filters) {
  try {
    const sh = setupInvestSheet();
    const lr = sh.getLastRow();
    if (lr < 2) return [];
    
    const S = function(v) { return v == null ? '' : String(v); };
    const D = function(v) {
      if (!v) return '';
      try {
        if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
        return String(v);
      } catch(e) { return String(v); }
    };
    
    const rawRows = sh.getRange(2,1,lr-1,23).getValues();
    const rows = [];
    for (let i = 0; i < rawRows.length; i++) {
      const r = rawRows[i];
      if (!r[0]) continue;
      rows.push({
        investId: S(r[0]), refId: S(r[1]), createdAt: D(r[2]),
        itemNo: S(r[3]), barcode: S(r[4]), sentDate: D(r[5]),
        courseCode: S(r[6]), courseName: S(r[7]), weight: S(r[8]), fee: S(r[9]),
        recipientName: S(r[10]), recipientAddr: S(r[11]),
        cause: S(r[12]), notifyDate: D(r[13]), notifyEmail: S(r[14]), notifyQty: S(r[15]),
        replyDate: D(r[16]), replyDays: S(r[17]), result: S(r[18]), resultDetail: S(r[19]),
        status: S(r[20]) || 'รอส่ง', recorder: S(r[21]), updatedAt: D(r[22])
      });
    }
    
    let result = rows;
    if (filters) {
      if (filters.status) result = result.filter(function(r){ return r.status === filters.status; });
      if (filters.search) {
        const q = String(filters.search).toLowerCase();
        result = result.filter(function(r) {
          return [r.investId, r.refId, r.barcode, r.itemNo, r.courseCode, r.courseName, r.recipientName]
            .some(function(v){ return v && v.toLowerCase().indexOf(q) !== -1; });
        });
      }
    }
    return result;
  } catch(e) { return { error: 'getInvestigations: ' + e.message }; }
}

function addInvestigation(data) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sh   = setupInvestSheet();
    const sess = _sess();
    const now  = new Date();
    const id   = 'INV-' + Utilities.formatDate(now,'Asia/Bangkok','yyyyMMdd-HHmmss');
    const row  = [
      id, data.refId||'', now,
      data.itemNo||'', data.barcode||'',
      data.sentDate ? new Date(data.sentDate) : '',  // 🆕 ถ้าไม่มี → เว้นว่าง (ไม่ fallback เป็น now)
      data.courseCode||'', data.courseName||'', data.weight||'', data.fee||'',
      data.recipientName||'', data.recipientAddr||'',
      data.cause||'', '', '', '',
      '', '', '', '',
      'รอส่ง', sess.name||sess.email, now
    ];
    sh.appendRow(row);
    return { success:true, id:id };
  } catch(e) { return { success:false, error:e.message }; }
}

function updateInvestigation(investId, updates) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sh   = setupInvestSheet();
    const data = sh.getDataRange().getValues();
    for (let i=1;i<data.length;i++) {
      if (data[i][0] === investId) {
        const row = i+1;
        // 🆕 อัปเดตข้อมูลพัสดุได้ด้วย
        if (updates.itemNo)     sh.getRange(row,4).setValue(updates.itemNo);
        if (updates.barcode)    sh.getRange(row,5).setValue(updates.barcode);
        if (updates.sentDate)    sh.getRange(row,6).setValue(new Date(updates.sentDate));
        if (updates.weight !== undefined)  sh.getRange(row,9).setValue(updates.weight);
        if (updates.fee !== undefined)     sh.getRange(row,10).setValue(updates.fee);
        // ข้อมูลการแจ้ง/สอบสวน (เดิม)
        if (updates.notifyDate)  sh.getRange(row,14).setValue(new Date(updates.notifyDate));
        if (updates.notifyEmail) sh.getRange(row,15).setValue(updates.notifyEmail);
        if (updates.notifyQty)   sh.getRange(row,16).setValue(updates.notifyQty);
        if (updates.replyDate)   sh.getRange(row,17).setValue(new Date(updates.replyDate));
        if (updates.replyDays !== undefined) sh.getRange(row,18).setValue(updates.replyDays);
        if (updates.result)      sh.getRange(row,19).setValue(updates.result);
        // 🆕 resultDetail รวม attachmentUrl ด้วย (append URL ไปท้าย detail)
        if (updates.resultDetail !== undefined) {
          var detailText = updates.resultDetail;
          if (updates.attachmentUrl) {
            detailText += '\n[ไฟล์แนบ: ' + updates.attachmentUrl + ']';
          }
          sh.getRange(row,20).setValue(detailText);
        } else if (updates.attachmentUrl) {
          // ถ้าแนบไฟล์อย่างเดียว ไม่แก้ detail — append ไปท้าย
          var curDetail = data[i][19] || '';
          sh.getRange(row,20).setValue(curDetail + '\n[ไฟล์แนบ: ' + updates.attachmentUrl + ']');
        }
        if (updates.status)      sh.getRange(row,21).setValue(updates.status);
        sh.getRange(row,23).setValue(new Date());
        logAudit('อัปเดตสอบสวน', investId+' | สถานะ: '+(updates.status||'-')+' | ผล: '+(updates.result||'-'));
        return { success:true };
      }
    }
    return { success:false, error:'ไม่พบรายการ' };
  } catch(e) { return { success:false, error:e.message }; }
}

// ============================================================
// ส่งอีเมลสอบสวนไปรษณีย์ (batch) - ส่งจากอีเมล admin โดยตรง
// ============================================================
function sendInvestigationBatchEmail(payload) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sess = _sess();
    const fromName  = (sess && sess.name)  || 'เจ้าหน้าที่';
    const fromEmail = (sess && sess.email) || '';

    const toEmail   = (payload.to || '').toString().trim();
    const ccEmail   = (payload.cc || '').toString().trim();
    const subject   = (payload.subject || 'ขอสอบสวนสิ่งของฝากส่งทางไปรษณีย์ — มสธ.').toString();
    const bodyText  = (payload.body || '').toString();
    const ids       = Array.isArray(payload.ids) ? payload.ids : [];

    if (!toEmail || !/\S+@\S+\.\S+/.test(toEmail)) {
      return { success:false, error:'อีเมลผู้รับไม่ถูกต้อง' };
    }
    if (!bodyText) return { success:false, error:'ไม่มีเนื้อหาอีเมล' };

    // แปลงเนื้อหาเป็น HTML
    const htmlBody = ''
      + '<div style="font-family:Sarabun,Arial,sans-serif;max-width:680px;margin:0 auto;background:#fff;padding:24px;color:#1a2840;line-height:1.7">'
      +   '<div style="border-bottom:3px solid #2d7dd2;padding-bottom:14px;margin-bottom:18px">'
      +     '<h2 style="margin:0;color:#0f2744">📮 ขอสอบสวนสิ่งของฝากส่งทางไปรษณีย์</h2>'
      +     '<p style="margin:4px 0 0;font-size:13px;color:#8899b4">สำนักบริการการศึกษา มหาวิทยาลัยสุโขทัยธรรมาธิราช</p>'
      +   '</div>'
      +   '<pre style="font-family:Sarabun,Arial,sans-serif;white-space:pre-wrap;font-size:14px;margin:0">' + _esc(bodyText) + '</pre>'
      +   '<div style="margin-top:24px;padding-top:14px;border-top:1px solid #e0e8f0;font-size:12px;color:#8899b4">'
      +     'ส่งโดย: ' + _esc(fromName) + (fromEmail ? ' &lt;' + _esc(fromEmail) + '&gt;' : '')
      +   '</div>'
      + '</div>';

    const opts = {
      htmlBody: htmlBody,
      name: fromName + ' — สำนักบริการการศึกษา มสธ.'
    };
    if (ccEmail) opts.cc = ccEmail;
    if (fromEmail && /\S+@\S+\.\S+/.test(fromEmail)) opts.replyTo = fromEmail;

    GmailApp.sendEmail(toEmail, subject, bodyText, opts);

    // อัปเดตสถานะของรายการที่ส่ง
    const today = Utilities.formatDate(new Date(),'Asia/Bangkok','yyyy-MM-dd');
    let updatedCount = 0;
    ids.forEach(function(id) {
      try {
        const res = updateInvestigation(id, {
          notifyDate: today,
          notifyEmail: toEmail,
          notifyQty: ids.length,
          status: 'ส่งแล้ว',
          result: 'รอประสาน'
        });
        if (res && res.success) updatedCount++;
      } catch(e) {}
    });

    logAudit('ส่งอีเมลสอบสวน', 'จำนวน '+ids.length+' รายการ ถึง '+toEmail);

    return {
      success: true,
      sentTo: toEmail,
      cc: ccEmail,
      sentFrom: fromEmail,
      itemsUpdated: updatedCount
    };
  } catch(e) {
    return { success:false, error:e.message };
  }
}

function addInvestigationsFromLend(recordIds) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_DATA);
    const data  = sheet.getDataRange().getValues();
    const results = [];
    recordIds.forEach(function(rid) {
      for (let i=1;i<data.length;i++) {
        if (data[i][0] === rid) {
          const r = data[i];
          let courses = [];
          try { courses = JSON.parse(r[26]||'[]'); } catch(e) {}
          if (!courses.length && r[6]) courses = [{code:r[6],b1:'',b2:'',track:'',date:''}];
          courses.forEach(function(c) {
            if (!c.code) return;
            const res = addInvestigation({
              refId: r[0],
              itemNo: r[0],
              barcode: c.track||r[20]||'',
              sentDate: '',  // 🆕 ไม่ดึงวันที่ — ให้เจ้าหน้าที่กรอกเองในหน้าสอบสวน
              courseCode: c.code,
              courseName: String(c.b1||c.b2||''),
              weight: '',
              fee: '',
              recipientName: (r[8]||'')+(r[9]||'')+' '+(r[10]||''),
              recipientAddr: [r[11],r[12],r[13],r[14],r[15],r[16]].filter(Boolean).join(' '),
              cause: r[18]||'',
            });
            results.push(res);
          });
          break;
        }
      }
    });
    return { success:true, created:results.length };
  } catch(e) { return { success:false, error:e.message }; }
}

function exportInvestigationsExcel() {
  // ส่งข้อมูลทั้งหมดกลับเป็น JSON สำหรับ export ฝั่ง client
  if (!_autoRefreshSession()) return { error:'SESSION_EXPIRED' };
  return getInvestigations({});
}

function testLogin() {
  // ทดสอบด้วย stou1234 (password ที่เพิ่ง reset)
  const result = login('methita.sap@stou.ac.th', 'stou1234');
  Logger.log('Login result: ' + JSON.stringify(result));
  
  // debug: แสดง hash ที่ควรจะเป็น vs ที่อยู่ใน Sheet
  const newHash = hashPw('stou1234');
  Logger.log('Hash of stou1234: ' + newHash);
  
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_USERS);
  const data = sh.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if ((data[i][0]||'').toString().toLowerCase().includes('methita')) {
      Logger.log('Sheet hash: ' + data[i][1]);
      Logger.log('Match: ' + (data[i][1] === newHash));
      Logger.log('Active: ' + data[i][4]);
    }
  }
}
function resetAdminHash() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ผู้ใช้งาน');
  const hash = hashPw('admin1234');
  sheet.getRange('B2').setValue(hash);
  if (!sheet.getRange('D2').getValue()) sheet.getRange('D2').setValue('ผู้ดูแลระบบ');
  Logger.log('Hash = ' + hash);
  Logger.log('ALERT: '+'เสร็จแล้ว! Hash = ' + hash.substring(0,20) + '...');
}

// อัปเดตสาเหตุตีคืนใน Sheet ตั้งค่า (รันครั้งเดียว)
// ดึง URL สำหรับ CRM Public (นศ. ใช้แจ้งปัญหา)
function getPublicUrl() {
  try {
    var url = ScriptApp.getService().getUrl();
    return { success:true, url: url + '?page=crm' };
  } catch(e) {
    return { success:false, error:e.message };
  }
}

function updateReturnCauses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const st = ss.getSheetByName(SH_SETTINGS);
  if (!st) { Logger.log('ALERT: '+'ไม่พบ Sheet ตั้งค่า'); return; }
  const data = st.getDataRange().getValues();
  for (let i=0; i<data.length; i++) {
    if (data[i][0] === 'สาเหตุตีคืน') {
      st.getRange(i+1,2).setValue('จ่าหน้าไม่ชัดเจน,ไม่มีเลขที่บ้านตามจ่าหน้า,ไม่ยอมรับ,ไม่มีผู้รับตามจ่าหน้า,ไม่มารับภายในกำหนด,ไม่มีจ่าหน้าหรือจ่าหน้าสูญหาย,ย้ายไม่ทราบที่อยู่ใหม่,อื่นๆ');
      Logger.log('ALERT: '+'อัปเดตสาเหตุตีคืนสำเร็จ ✅');
      return;
    }
  }
  Logger.log('ALERT: '+'ไม่พบแถว สาเหตุตีคืน');
}

// อัปเดต CRM settings ใน Sheet ที่มีอยู่แล้ว (รันครั้งเดียว)
function updateCrmSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const st = ss.getSheetByName(SH_SETTINGS);
  if (!st) { Logger.log('ALERT: '+'ไม่พบ Sheet ตั้งค่า'); return; }
  const data = st.getDataRange().getValues();
  const keys = data.map(r=>r[0]);
  const updates = {
    'สาขาวิชา': 'สาขาศิลปศาสตร์,สาขาศึกษาศาสตร์,สาขาวิทยาการจัดการ,สาขานิติศาสตร์,สาขาวิทยาศาสตร์สุขภาพ,สาขาเศรษฐศาสตร์,สาขามนุษยนิเวศศาสตร์,สาขารัฐศาสตร์,สาขาเกษตรศาสตร์และสหกรณ์,สาขานิเทศศาสตร์,สาขาวิทยาศาสตร์และเทคโนโลยี,สาขาพยาบาลศาสตร์',
    'แผนการศึกษา': 'แผน ก1,แผน ก2,แผน ก3',
    'ประเภทปัญหา': 'พิมพ์เพิ่ม,ชุดปรับปรุง,ชุดผลิตใหม่,ปัญหาทวงถามหนังสือ (ไม่ได้สั่งซื้อ/แผน ก2-ก3),สอบถามทะเบียนและวัดผล (ลงทะเบียน/สอบ),สอบถามกิจกรรมประจำชุดวิชา,สอบถามอบรมเข้มเสริมประสบการณ์วิชาชีพ,สอบถามสอนเสริมออนไลน์,สอบถามโครงการสัมฤทธิบัตร,ไม่ได้รับเอกสาร,เอกสารชำรุด,ส่งผิดวิชา,อื่นๆ (ระบุเอง)',
    'ผู้รับเรื่อง': 'หทัย เริงเกษตรกิจ,วรรณี รัตนากร,สุพรรษา ช่อปทุมมา,เมธิตา สาไพรวัน',
    'สาเหตุตีคืน': 'จ่าหน้าไม่ชัดเจน,ไม่มีเลขที่บ้านตามจ่าหน้า,ไม่ยอมรับ,ไม่มีผู้รับตามจ่าหน้า,ไม่มารับภายในกำหนด,ไม่มีจ่าหน้าหรือจ่าหน้าสูญหาย,ย้ายไม่ทราบที่อยู่ใหม่,อื่นๆ',
  };
  Object.keys(updates).forEach(function(key) {
    const idx = keys.indexOf(key);
    if (idx > -1) {
      st.getRange(idx+1, 2).setValue(updates[key]);
    } else {
      st.appendRow([key, updates[key]]);
    }
  });
  Logger.log('ALERT: '+'อัปเดตตั้งค่า CRM สำเร็จ ✅');
}

// ส่งอีเมลแจ้งผู้แจ้งเมื่อปิดเรื่อง CRM
function sendCrmClosedEmail(crmId) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CRM);
    const data = sh.getDataRange().getValues();
    for (let i=1;i<data.length;i++) {
      if (data[i][0] === crmId) {
        const reporterEmail = data[i][4];
        const reporterName  = data[i][2];
        const issueType     = data[i][11];
        const detail        = data[i][12];
        const assignee      = data[i][17];
        if (!reporterEmail) return { success:false, error:'ไม่พบอีเมลผู้แจ้ง' };
        const subject = 'สำนักบริการการศึกษา มสธ. — อัปเดตการดำเนินการ: '+issueType;
        const body = 'เรียน '+reporterName+'\n\n'
          +'เรื่องที่ท่านแจ้งไว้ ('+crmId+') ได้รับการดำเนินการแล้ว\n\n'
          +'ประเภทปัญหา: '+issueType+'\n'
          +'รายละเอียด: '+detail+'\n'
          +'ผู้ดำเนินการ: '+assignee+'\n\n'
          +'หากมีข้อสงสัยเพิ่มเติม กรุณาติดต่อสำนักบริการการศึกษา\n'
          +'โทร. 02 504 7623, 7626\n\n'
          +'ขอแสดงความนับถือ\n'
          +'สำนักบริการการศึกษา มหาวิทยาลัยสุโขทัยธรรมาธิราช';
        GmailApp.sendEmail(reporterEmail, subject, body);
        return { success:true };
      }
    }
    return { success:false, error:'ไม่พบรายการ CRM' };
  } catch(e) { return { success:false, error:e.message }; }
}

// ดึงข้อมูล CRM เพื่อ prefill ฟอร์ม ตีคืน/ให้ยืม/พิเศษ
function getCrmForPrefill(crmId) {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CRM);
    const data = sh.getDataRange().getValues();
    for (let i=1;i<data.length;i++) {
      if (data[i][0] === crmId) {
        const courses = (data[i][10]||'').toString().split(',').filter(Boolean);
        return {
          success: true,
          studentId:    data[i][3]||'',
          reporterName: data[i][2]||'',
          email:        data[i][4]||'',
          phone:        data[i][5]||'',
          term:         data[i][8]||'',
          year:         data[i][9]||'',
          courseCode:   courses[0]||'',
          courses:      JSON.stringify(courses.map(function(c){return {code:c,b1:'',b2:'',track:'',date:''};})),
          issueType:    data[i][11]||'',
        };
      }
    }
    return { error:'ไม่พบรายการ' };
  } catch(e) { return { error:e.message }; }
}

// ============================================================
// EXPORT — ส่งข้อมูลกลับสำหรับ export Excel (CSV)
// ============================================================
function exportRecordsData(filters) {
  // session check ผ่าน _autoRefreshSession อัตโนมัติ
  return getRecords(filters||{});
}

function exportCrmData() {
  if (!_autoRefreshSession()) return { error:'SESSION_EXPIRED' };
  return getCrmTickets({});
}

function getCrmTemplateOptions() {
  const DEFAULT_ISSUES = ['พิมพ์เพิ่ม','ชุดปรับปรุง','ชุดผลิตใหม่','ปัญหาทวงถามหนังสือ (ไม่ได้สั่งซื้อ/แผน ก2-ก3)','สอบถามทะเบียนและวัดผล (ลงทะเบียน/สอบ)','สอบถามกิจกรรมประจำชุดวิชา','สอบถามเลือกแผนการศึกษา','สอบถามอบรมเข้มเสริมประสบการณ์วิชาชีพ','สอบถามสอนเสริมออนไลน์','สอบถามโครงการสัมฤทธิบัตร','ไม่ได้รับเอกสาร','เอกสารชำรุด','ส่งผิดวิชา','อื่นๆ (ระบุเอง)'];
  const issues = (settings && settings.issueTypes && settings.issueTypes.length) ? settings.issueTypes : DEFAULT_ISSUES;
  return {
    issueTypes: issues,
    educationLevels: ['ปริญญาตรี','ปริญญาโท','ปริญญาเอก','หนังสือหมายเหตุ','อื่นๆ'],
    channels: ['อีเมล','โทรศัพท์','Line','Facebook','อื่นๆ'],
    priorities: ['normal','high','urgent'],
  };
}

function createCrmImportTemplate() {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const options = getCrmTemplateOptions();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let templateSheet = ss.getSheetByName('CRM_Import_Template');
    if (!templateSheet) {
      templateSheet = ss.insertSheet('CRM_Import_Template');
    } else {
      templateSheet.clear();
    }

    const headers = ['ชื่อ','อีเมล','เบอร์โทร','รหัสนักศึกษา','ประเภทปัญหา','รายละเอียด','ชุดวิชา','หน่วยงาน/สาขา','ระดับการศึกษา','ภาค','ปี','แผนการศึกษา','ช่องทาง','ลำดับความสำคัญ','หมายเหตุ'];
    templateSheet.appendRow(headers);
    const headerRange = templateSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1a3a5c').setFontColor('#fff').setFontWeight('bold');
    templateSheet.setFrozenRows(1);

    // ตั้ง column width
    templateSheet.setColumnWidth(1, 150);
    templateSheet.setColumnWidth(2, 180);
    templateSheet.setColumnWidth(3, 130);
    templateSheet.setColumnWidth(4, 110);
    templateSheet.setColumnWidth(5, 200);
    templateSheet.setColumnWidth(6, 250);
    templateSheet.setColumnWidth(7, 130);
    templateSheet.setColumnWidth(8, 150);
    templateSheet.setColumnWidth(9, 140);
    templateSheet.setColumnWidth(10, 70);
    templateSheet.setColumnWidth(11, 70);
    templateSheet.setColumnWidth(12, 150);
    templateSheet.setColumnWidth(13, 130);
    templateSheet.setColumnWidth(14, 110);
    templateSheet.setColumnWidth(15, 200);

    // เพิ่ม data validation สำหรับแถว 2-500
    const issueRange = templateSheet.getRange('E2:E500');
    const issueRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(options.issueTypes)
      .setAllowInvalid(false)
      .setHelpText('เลือกประเภทปัญหา')
      .build();
    issueRange.setDataValidation(issueRule);

    const educationRange = templateSheet.getRange('I2:I500');
    const educationRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(options.educationLevels)
      .setAllowInvalid(false)
      .setHelpText('เลือกระดับการศึกษา')
      .build();
    educationRange.setDataValidation(educationRule);

    const channelRange = templateSheet.getRange('M2:M500');
    const channelRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(options.channels)
      .setAllowInvalid(false)
      .setHelpText('เลือกช่องทาง')
      .build();
    channelRange.setDataValidation(channelRule);

    const priorityRange = templateSheet.getRange('N2:N500');
    const priorityRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(options.priorities)
      .setAllowInvalid(false)
      .setHelpText('เลือกลำดับความสำคัญ')
      .build();
    priorityRange.setDataValidation(priorityRule);

    const url = ss.getUrl();
    return { success:true, message:'สร้างเทมเพลตสำเร็จ', url:url };
  } catch(e) { return { success:false, error:e.message }; }
}

// ============================================================
// MONTHLY REPORT — ส่งรายงานประจำเดือน
// ============================================================
function sendMonthlyReport() {
  if (!_autoRefreshSession()) return { success:false, error:'เฉพาะ Super Admin' };
  try {
    const sess  = _sess();
    const now   = new Date();
    const month = Utilities.formatDate(now,'Asia/Bangkok','MMMM yyyy');
    const records = getRecords({});
    if (records.error) return { success:false, error:records.error };

    // สรุปสถิติ
    const total   = records.length;
    const ret     = records.filter(r=>r.recType==='return').length;
    const lend    = records.filter(r=>r.recType==='lend').length;
    const special = records.filter(r=>r.recType==='special').length;
    const crm     = getCrmTickets({});
    const crmTotal= Array.isArray(crm) ? crm.length : 0;

    // จังหวัดสูงสุด
    const provMap = {};
    records.forEach(r=>{ if(r.province) provMap[r.province]=(provMap[r.province]||0)+1; });
    const topProv = Object.entries(provMap).sort((a,b)=>b[1]-a[1]).slice(0,5);

    const subject = 'รายงานประจำเดือน ' + month + ' — ระบบจัดการเอกสารการสอน มสธ.';
    const body = `รายงานสรุปประจำเดือน ${month}
สำนักบริการการศึกษา มหาวิทยาลัยสุโขทัยธรรมาธิราช

━━━━━━━━━━━━━━━━━━━━
สถิติรายการพัสดุ
━━━━━━━━━━━━━━━━━━━━
รวมทั้งหมด:    ${total} รายการ
พัสดุตีคืน:    ${ret} รายการ
ให้ยืม:        ${lend} รายการ
ส่งใหม่พิเศษ:  ${special} รายการ
CRM แจ้งปัญหา: ${crmTotal} รายการ

จังหวัดที่มีปัญหาสูงสุด:
${topProv.map((p,i)=>`  ${i+1}. ${p[0]}: ${p[1]} รายการ`).join('\n')}

━━━━━━━━━━━━━━━━━━━━
จัดทำโดยระบบอัตโนมัติ
ผู้สั่งรายงาน: ${sess.name||sess.email}
วันที่: ${fmtDate(now)}`;

    MailApp.sendEmail(sess.email, subject, body);
    return { success:true, email:sess.email };
  } catch(e) { return { success:false, error:e.message }; }
}

// ============================================================
// AUDIT LOG
// ============================================================
const SH_AUDIT = 'Audit Log';

function logAudit(action, detail) {
  try {
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    let sh     = ss.getSheetByName(SH_AUDIT);
    if (!sh) {
      sh = ss.insertSheet(SH_AUDIT);
      sh.appendRow(['วันที่','ผู้ใช้','อีเมล','การกระทำ','รายละเอียด']);
      const hr=sh.getRange(1,1,1,5);hr.setBackground('#1a3a5c');hr.setFontColor('#fff');hr.setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    const sess = _sess();
    var recEmail = '';
    var recName  = '';
    try { recEmail = Session.getEffectiveUser().getEmail() || ''; } catch(ex) {}
    // Look up name from users sheet by email
    if (recEmail) {
      try {
        var uSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_USERS);
        if (uSh) {
          var uData = uSh.getDataRange().getValues();
          for (var ui = 1; ui < uData.length; ui++) {
            if ((uData[ui][0]||'').toLowerCase().trim() === recEmail.toLowerCase().trim()) {
              recName = uData[ui][3] || recEmail;
              break;
            }
          }
          if (!recName) recName = recEmail;
        }
      } catch(ex) { recName = recEmail; }
    }
    if (!recName) { recName = sess.name||'ระบบ'; recEmail = sess.email||''; }
    sh.appendRow([new Date(), recName, recEmail, action, detail||'']);
    const lr = sh.getLastRow();
    if (lr%2===0) sh.getRange(lr,1,1,5).setBackground('#f8f9fa');
  } catch(e) {}
}

function getAuditLog(limit) {
  // session check ผ่าน _autoRefreshSession อัตโนมัติ
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_AUDIT);
    if (!sh || sh.getLastRow()<2) return [];
    const n   = Math.min(limit||100, sh.getLastRow()-1);
    const rows= sh.getRange(sh.getLastRow()-n+1,1,n,5).getValues().reverse();
    return rows.map(r=>({
      date:r[0]?fmtDate(r[0]):'', name:r[1], email:r[2], action:r[3], detail:r[4]
    }));
  } catch(e) { return { error:e.message }; }
}

// เพิ่ม audit ใน addRecord
const _origAddRecord = typeof addRecord !== 'undefined' ? addRecord : null;

// ============================================================
// EMAIL NOTIFICATION — แจ้ง นศ. (HTML e-newsletter style)
// ============================================================
function sendStudentNotification(data) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  if (!data.toEmail) return { success:false, error:'ไม่มีอีเมลผู้รับ' };
  try {
    const sess    = _sess();
    const logoUrl = getOrgLogoUrl();
    const typeLabel = {return:'พัสดุตีคืน',lend:'ยืมชุดเอกสาร',special:'จัดส่งใหม่กรณีพิเศษ'}[data.recType]||'เอกสารการสอน';

    // สร้าง HTML body แบบ e-newsletter
    const htmlBody = `<!DOCTYPE html>
<html lang="th">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<style>
  body{margin:0;padding:0;background:#f4f6fb;font-family:'Sarabun',Arial,sans-serif}
  .wrap{max-width:600px;margin:24px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.10)}
  .header{background:linear-gradient(135deg,#0f2744 0%,#1a4a8a 100%);padding:32px 36px;text-align:center}
  .header img{height:56px;margin-bottom:14px;display:block;margin-left:auto;margin-right:auto}
  .header h1{color:#fff;font-size:20px;margin:0 0 4px;font-weight:700}
  .header p{color:rgba(255,255,255,0.75);font-size:13px;margin:0}
  .banner{background:#f4a21e;padding:10px 36px;text-align:center}
  .banner span{color:#0f2744;font-weight:700;font-size:14px}
  .body{padding:28px 36px}
  .greeting{font-size:16px;color:#1a2840;margin-bottom:18px}
  .info-box{background:#f0f4f8;border-radius:8px;padding:18px 20px;margin-bottom:20px}
  .info-row{display:flex;padding:5px 0;border-bottom:1px solid #e0e8f0;font-size:14px}
  .info-row:last-child{border-bottom:none}
  .info-label{color:#8899b4;width:140px;flex-shrink:0;font-size:13px}
  .info-val{color:#1a2840;font-weight:600}
  .message-box{background:#fff8e1;border-left:4px solid #f4a21e;padding:14px 18px;border-radius:0 8px 8px 0;margin-bottom:20px;font-size:14px;color:#5c4b00;line-height:1.7}
  .cta{text-align:center;margin:24px 0}
  .btn-cta{background:linear-gradient(135deg,#2d7dd2,#1a4a8a);color:#fff;padding:13px 32px;border-radius:8px;text-decoration:none;font-weight:700;font-size:15px;display:inline-block}
  .footer{background:#f4f6fb;padding:20px 36px;text-align:center;border-top:1px solid #e0e8f0}
  .footer p{color:#8899b4;font-size:12px;margin:2px 0;line-height:1.7}
</style></head>
<body>
<div class="wrap">
  <div class="header">
    ${logoUrl?'<img src="'+logoUrl+'" alt="STOU Logo">':'<div style="font-size:40px;margin-bottom:10px">📚</div>'}
    <h1>มหาวิทยาลัยสุโขทัยธรรมาธิราช</h1>
    <p>สำนักบริการการศึกษา — ศูนย์บริการการสอนทางไปรษณีย์</p>
  </div>
  <div class="banner"><span>แจ้งสถานะ: ${typeLabel}</span></div>
  <div class="body">
    <div class="greeting">เรียน <strong>${data.name||'นักศึกษา'}</strong></div>
    <div class="info-box">
      <div class="info-row"><span class="info-label">รหัสนักศึกษา</span><span class="info-val">${data.studentId||'—'}</span></div>
      <div class="info-row"><span class="info-label">ชุดวิชา</span><span class="info-val">${data.courseCode||'—'}</span></div>
      <div class="info-row"><span class="info-label">ประเภท</span><span class="info-val">${typeLabel}</span></div>
      <div class="info-row"><span class="info-label">สาเหตุ</span><span class="info-val">${data.cause||'—'}</span></div>
      <div class="info-row"><span class="info-label">เลขพัสดุ</span><span class="info-val">${data.trackNo||'—'}</span></div>
      <div class="info-row"><span class="info-label">วันที่บันทึก</span><span class="info-val">${fmtDate(new Date())}</span></div>
    </div>
    <div class="message-box">${data.message||'กรุณาตรวจสอบข้อมูลและติดต่อกลับมายังสำนักบริการการศึกษา หากมีข้อสงสัยหรือต้องการดำเนินการใดๆ'}</div>
    <div class="cta"><a href="https://www.stou.ac.th" class="btn-cta">ติดต่อสำนักบริการการศึกษา</a></div>
  </div>
  <div class="footer">
    <p><strong>สำนักบริการการศึกษา มสธ.</strong></p>
    <p>โทร. 02 504 7623, 7626 | อีเมล: oes@stou.ac.th</p>
    <p>9/9 หมู่ 9 ต.บางพูด อ.ปากเกร็ด จ.นนทบุรี 11120</p>
    <p style="color:#c5cfe0;margin-top:8px">อีเมลนี้ส่งโดยอัตโนมัติจากระบบจัดการเอกสารการสอน มสธ. — กรุณาอย่าตอบกลับอีเมลนี้โดยตรง</p>
  </div>
</div>
</body></html>`;

    GmailApp.sendEmail(data.toEmail, data.subject||('แจ้งสถานะ'+typeLabel+' — มสธ.'), '', {
      htmlBody: htmlBody,
      name: 'สำนักบริการการศึกษา มสธ.',
      replyTo: 'oes@stou.ac.th',
    });
    logAudit('ส่งอีเมล นศ.', data.toEmail+' | '+data.studentId+' | '+typeLabel);
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

// โลโก้หน่วยงาน — เก็บใน PropertiesService
function getOrgLogoUrl() {
  try { return PropertiesService.getScriptProperties().getProperty('ORG_LOGO_URL')||''; } catch(e){ return ''; }
}
function saveOrgLogoUrl(url) {
  if (!_autoRefreshSession()) return { success:false, error:'เฉพาะ Super Admin' };
  try { PropertiesService.getScriptProperties().setProperty('ORG_LOGO_URL', url); return { success:true }; } catch(e){ return { success:false, error:e.message }; }
}
function getOrgLogoUrlPublic() {
  return getOrgLogoUrl();
}

// ============================================================
// PERSONAL STATS — สถิติผลการปฏิบัติงานรายบุคคล
// ============================================================
function getPersonalStats(period) {
  try {
    const sess = _sess();
    const myEmail = sess.email||'';
    const myName  = sess.name||'';
    // ดึงข้อมูลทั้งหมด แล้ว filter ของตัวเอง
    const all = getRecords({});
    if (all.error) return all;
    const crm = getCrmTickets({});
    const inv = getInvestigations({});
    const audit = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_AUDIT);

    // คำนวณช่วงเวลา
    const now = new Date();
    const startWeek = new Date(now); startWeek.setDate(now.getDate() - now.getDay());
    startWeek.setHours(0,0,0,0);
    const startMonth = new Date(now.getFullYear(), now.getMonth(), 1);

    // filter audit log ของ user นี้
    let auditRows = [];
    if (audit && audit.getLastRow() > 1) {
      auditRows = audit.getRange(2,1,audit.getLastRow()-1,5).getValues()
        .filter(r => r[2]===myEmail || r[1]===myName);
    }

    const filterByDate = (rows, dateStr, startDate) => {
      return rows.filter(r => {
        try { return new Date(r[dateStr]||r[0]) >= startDate; } catch(e) { return false; }
      });
    };

    const countAuditAction = (rows, action, start) => {
      return rows.filter(r => {
        try {
          const d = new Date(r[0]);
          return d >= start && (r[3]||'').includes(action);
        } catch(e){ return false; }
      }).length;
    };

    return {
      success: true,
      user: myName||myEmail,
      thisWeek: {
        records: countAuditAction(auditRows, 'บันทึกพัสดุ', startWeek),
        crm:     countAuditAction(auditRows, 'รับเรื่อง CRM', startWeek),
        emails:  countAuditAction(auditRows, 'ส่งอีเมล', startWeek),
      },
      thisMonth: {
        records: countAuditAction(auditRows, 'บันทึกพัสดุ', startMonth),
        crm:     countAuditAction(auditRows, 'รับเรื่อง CRM', startMonth),
        emails:  countAuditAction(auditRows, 'ส่งอีเมล', startMonth),
      },
      allTime: {
        records: all.filter(r=>r.recorder&&(r.recorder.includes(myName)||r.recorder.includes(myEmail))).length,
        crm: Array.isArray(crm)?crm.filter(c=>c.assigneeName&&c.assigneeName.includes(myName)).length:0,
      },
    };
  } catch(e) { return { error:e.message }; }
}

// TASK BOARD — ระบบสั่งงานรายบุคคล
// ============================================================
const SH_TASKS = 'งานที่สั่ง';

function setupTaskSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SH_TASKS);
  if (!sh) {
    sh = ss.insertSheet(SH_TASKS);
    sh.appendRow(['รหัสงาน','วันที่สั่ง','ผู้สั่ง','ผู้รับผิดชอบ','หัวข้องาน','รายละเอียด',
      'ประเภทงาน','สถานะ','วันกำหนดส่ง','ไฟล์แนบ URL','หมายเหตุผู้รับ',
      'ต้องอนุมัติก่อน','สถานะอนุมัติ','ผู้อนุมัติ','วันที่อนุมัติ','วันที่อัปเดต']);
    const hr=sh.getRange(1,1,1,16);hr.setBackground('#1a3a5c');hr.setFontColor('#fff');hr.setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function getTasks(filters) {
  // session check ผ่าน _autoRefreshSession อัตโนมัติ
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_TASKS);
    if (!sh) return [];
    const lr = sh.getLastRow();
    if (lr < 2) return [];
    const sess = _sess();
    let rows = sh.getRange(2,1,lr-1,16).getValues().map(r=>({
      taskId:r[0], createdAt:r[1]?fmtDate(r[1]):'', createdBy:r[2], assignee:r[3],
      title:r[4], detail:r[5], taskType:r[6], status:r[7]||'รอดำเนินการ',
      dueDate:r[8]?fmtDate(r[8]):'', fileUrl:r[9], assigneeNote:r[10],
      needApproval:r[11], approvalStatus:r[12]||'', approvedBy:r[13],
      approvedAt:r[14]?fmtDate(r[14]):'', updatedAt:r[15]?fmtDate(r[15]):'',
    })).filter(r=>r.taskId!=='');

    // staff เห็นเฉพาะงานตัวเอง, superadmin เห็นทั้งหมด
    if (sess.role !== R_SUPER) {
      rows = rows.filter(r => r.assignee===sess.name || r.assignee===sess.email || r.createdBy===sess.name);
    }
    if (filters && filters.status) rows = rows.filter(r=>r.status===filters.status);
    if (filters && filters.assignee) rows = rows.filter(r=>r.assignee===filters.assignee);
    return rows;
  } catch(e) { return { error:e.message }; }
}

function addTask(data) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sh   = setupTaskSheet();
    const sess = _sess();
    const now  = new Date();
    const id   = 'TASK-'+Utilities.formatDate(now,'Asia/Bangkok','yyyyMMddHHmmss');
    sh.appendRow([
      id, now, sess.name||sess.email, data.assignee||'',
      data.title||'', data.detail||'', data.taskType||'ทั่วไป',
      'รอดำเนินการ', data.dueDate?new Date(data.dueDate):'',
      data.fileUrl||'', '', data.needApproval||false, 
      data.needApproval?'รออนุมัติ':'ไม่ต้องอนุมัติ', '', '', now
    ]);
    logAudit('สั่งงาน', id+' | '+data.title+' | ผู้รับ: '+data.assignee);
    return { success:true, id:id };
  } catch(e) { return { success:false, error:e.message }; }
}

function updateTask(taskId, updates) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sh   = setupTaskSheet();
    const data = sh.getDataRange().getValues();
    const sess = _sess();
    for (let i=1;i<data.length;i++) {
      if (data[i][0] === taskId) {
        const row = i+1;
        if (updates.status)       sh.getRange(row,8).setValue(updates.status);
        if (updates.assigneeNote) sh.getRange(row,11).setValue(updates.assigneeNote);
        if (updates.fileUrl)      sh.getRange(row,10).setValue(updates.fileUrl);
        // อนุมัติ (superadmin เท่านั้น)
        if (updates.approve !== undefined && sess.role === R_SUPER) {
          sh.getRange(row,13).setValue(updates.approve?'อนุมัติแล้ว':'ไม่อนุมัติ');
          sh.getRange(row,14).setValue(sess.name||sess.email);
          sh.getRange(row,15).setValue(new Date());
          if (updates.approve) sh.getRange(row,8).setValue('อนุมัติแล้ว — พร้อมโพส');
        }
        sh.getRange(row,16).setValue(new Date());
        logAudit('อัปเดตงาน', taskId+' | '+updates.status);
        return { success:true };
      }
    }
    return { success:false, error:'ไม่พบงาน' };
  } catch(e) { return { success:false, error:e.message }; }
}

function sendCrmResolutionEmail(id, emailData) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sess = _sess();
    const reporterEmail = emailData.to;
    const reporterName = emailData.reporterName || 'ผู้แจ้ง';
    const issueType = emailData.issueType || 'แจ้งปัญหา';
    const detail = emailData.detail || '';
    const replies = emailData.replies || '<p>ปัญหาได้รับการแก้ไขแล้ว</p>';
    const adminName = emailData.adminName || 'เจ้าหน้าที่';

    if (!reporterEmail || !/\S+@\S+\.\S+/.test(reporterEmail)) {
      return { success:false, error:'ไม่พบอีเมลผู้แจ้ง' };
    }

    const subject = '[มสธ.] ปัญหาของท่านได้รับการแก้ไข: ' + issueType + ' (' + id + ')';
    const htmlBody = ''
      + '<div style="font-family:Sarabun,Arial,sans-serif;max-width:600px;margin:0 auto;background:#f4f6fb;padding:20px">'
      +   '<div style="background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 16px rgba(0,0,0,0.08)">'
      +     '<div style="background:linear-gradient(135deg,#0f8a4a,#16a34a);padding:24px;text-align:center;color:#fff">'
      +       '<div style="font-size:32px;margin-bottom:6px">✅</div>'
      +       '<h2 style="margin:0;font-size:18px">ปัญหาได้รับการแก้ไขแล้ว</h2>'
      +       '<p style="margin:4px 0 0;font-size:13px;color:rgba(255,255,255,0.9)">มหาวิทยาลัยสุโขทัยธรรมาธิราช</p>'
      +     '</div>'
      +     '<div style="padding:24px">'
      +       '<p style="font-size:15px;color:#1a2840">เรียน ' + _esc(reporterName) + '</p>'
      +       '<p style="color:#1a2840">ปัญหาที่ท่านแจ้งหว่านได้รับการแก้ไขสำเร็จแล้ว</p>'
      +       '<div style="background:#f0f4f8;border-left:4px solid #2d7dd2;padding:12px 16px;margin:16px 0;border-radius:4px">'
      +         '<div style="font-size:12px;color:#8899b4;margin-bottom:4px">รหัสเรื่อง: <strong>' + _esc(id) + '</strong></div>'
      +         '<div style="font-size:12px;color:#8899b4;margin-bottom:4px">ประเภท: <strong>' + _esc(issueType) + '</strong></div>'
      +         '<div style="font-size:12px;color:#8899b4">สถานะ: <strong style="color:#16a34a">แก้ไขแล้ว ✅</strong></div>'
      +       '</div>'
      +       '<div style="background:#e8f8f0;border-left:4px solid #16a34a;padding:14px 18px;border-radius:4px;margin-bottom:16px">'
      +         '<div style="font-size:13px;color:#0a5a3a;line-height:1.7">'
      +           '<strong>รายละเอียดการแก้ไข:</strong><br>'
      +           replies
      +         '</div>'
      +       '</div>'
      +       '<p style="color:#1a2840;font-size:13px">หากต้องการสอบถามเพิ่มเติม กรุณาตอบกลับอีเมลฉบับนี้ หรือโทร 02 504 7623, 7626</p>'
      +     '</div>'
      +     '<div style="background:#f4f6fb;padding:16px 24px;text-align:center;border-top:1px solid #e0e8f0">'
      +       '<p style="margin:0;font-size:11px;color:#8899b4">ขอแสดงความนับถือ<br>สำนักบริการการศึกษา มสธ.</p>'
      +     '</div>'
      +   '</div>'
      + '</div>';

    const textBody = 'เรียน ' + reporterName + '\n\n'
      + 'ปัญหาที่ท่านแจ้งหว่านได้รับการแก้ไขสำเร็จแล้ว\n\n'
      + 'รหัสเรื่อง: ' + id + '\n'
      + 'ประเภท: ' + issueType + '\n'
      + 'สถานะ: แก้ไขแล้ว ✅\n\n'
      + '----------------------------------------\n'
      + 'รายละเอียดการแก้ไข:\n'
      + replies + '\n'
      + '----------------------------------------\n\n'
      + 'หากต้องการสอบถามเพิ่มเติม กรุณาตอบกลับอีเมลฉบับนี้ หรือโทร 02 504 7623, 7626\n\n'
      + 'ขอแสดงความนับถือ\n'
      + 'สำนักบริการการศึกษา มสธ.';

    const opts = {
      htmlBody: htmlBody,
      name: adminName + ' — สำนักบริการการศึกษา มสธ.'
    };

    GmailApp.sendEmail(reporterEmail, subject, textBody, opts);
    logAudit('ส่งอีเมล CRM', id + ' | ' + issueType + ' | ตรวจสอบปัญหา');
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

function sendCrmEmail(id, emailData) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sess = _sess();
    const reporterEmail = emailData.to;
    const reporterName = emailData.reporterName || 'ผู้แจ้ง';
    const issueType = emailData.issueType || 'แจ้งปัญหา';
    const detail = emailData.detail || '';
    const replies = emailData.replies || '<p>ข้อมูลได้รับการบันทึกแล้ว</p>';
    const adminName = emailData.adminName || 'เจ้าหน้าที่';

    if (!reporterEmail || !/\S+@\S+\.\S+/.test(reporterEmail)) {
      return { success:false, error:'ไม่พบอีเมลผู้แจ้ง' };
    }

    const subject = '[มสธ.] ตอบกลับเรื่องที่ท่านแจ้ง: ' + issueType + ' (' + id + ')';
    const htmlBody = ''
      + '<div style="font-family:Sarabun,Arial,sans-serif;max-width:600px;margin:0 auto;background:#f4f6fb;padding:20px">'
      +   '<div style="background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 16px rgba(0,0,0,0.08)">'
      +     '<div style="background:linear-gradient(135deg,#0f2744,#1a4a8a);padding:24px;text-align:center;color:#fff">'
      +       '<div style="font-size:32px;margin-bottom:6px">📚</div>'
      +       '<h2 style="margin:0;font-size:18px">มหาวิทยาลัยสุโขทัยธรรมาธิราช</h2>'
      +       '<p style="margin:4px 0 0;font-size:13px;color:rgba(255,255,255,0.8)">สำนักบริการการศึกษา</p>'
      +     '</div>'
      +     '<div style="padding:24px">'
      +       '<p style="font-size:15px;color:#1a2840">เรียน ' + _esc(reporterName) + '</p>'
      +       '<p style="color:#1a2840">เจ้าหน้าที่ได้รับและบันทึกเรื่องที่ท่านแจ้งไว้ดังนี้</p>'
      +       '<div style="background:#f0f4f8;border-left:4px solid #2d7dd2;padding:12px 16px;margin:16px 0;border-radius:4px">'
      +         '<div style="font-size:12px;color:#8899b4;margin-bottom:4px">รหัสเรื่อง: <strong>' + _esc(id) + '</strong></div>'
      +         '<div style="font-size:12px;color:#8899b4;margin-bottom:4px">ประเภท: <strong>' + _esc(issueType) + '</strong></div>'
      +         '<div style="font-size:12px;color:#8899b4">ผู้ดำเนินการ: <strong>' + _esc(adminName) + '</strong></div>'
      +       '</div>'
      +       '<div style="background:#fff8e1;border-left:4px solid #f4a21e;padding:14px 18px;border-radius:4px;margin-bottom:16px">'
      +         '<div style="font-size:13px;color:#5c4b00;line-height:1.7">'
      +           '<strong>รายละเอียดการดำเนินการ:</strong><br>'
      +           replies
      +         '</div>'
      +       '</div>'
      +       '<p style="color:#1a2840;font-size:13px">หากต้องการสอบถามเพิ่มเติม กรุณาตอบกลับอีเมลฉบับนี้ หรือโทร 02 504 7623, 7626</p>'
      +     '</div>'
      +     '<div style="background:#f4f6fb;padding:16px 24px;text-align:center;border-top:1px solid #e0e8f0">'
      +       '<p style="margin:0;font-size:11px;color:#8899b4">ขอแสดงความนับถือ<br>สำนักบริการการศึกษา มสธ.</p>'
      +     '</div>'
      +   '</div>'
      + '</div>';

    const textBody = 'เรียน ' + reporterName + '\n\n'
      + 'เจ้าหน้าที่ได้รับและบันทึกเรื่องที่ท่านแจ้งไว้ (' + id + ') ดังนี้\n\n'
      + 'ประเภท: ' + issueType + '\n'
      + 'ผู้ดำเนินการ: ' + adminName + '\n\n'
      + '----------------------------------------\n'
      + 'รายละเอียดการดำเนินการ:\n'
      + replies + '\n'
      + '----------------------------------------\n\n'
      + 'หากต้องการสอบถามเพิ่มเติม กรุณาตอบกลับอีเมลฉบับนี้ หรือโทร 02 504 7623, 7626\n\n'
      + 'ขอแสดงความนับถือ\n'
      + 'สำนักบริการการศึกษา มสธ.';

    const opts = {
      htmlBody: htmlBody,
      name: adminName + ' — สำนักบริการการศึกษา มสธ.'
    };

    GmailApp.sendEmail(reporterEmail, subject, textBody, opts);
    logAudit('ส่งอีเมล CRM', id + ' | ' + issueType + ' | ส่งอีเมลแจ้งผล');
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

function getTaskBadgeCount() {
  // นับงานที่รอดำเนินการของ user ปัจจุบัน
  if (!_autoRefreshSession()) return 0;
  try {
    const tasks = getTasks({status:'รอดำเนินการ'});
    const approvalTasks = getTasks({status:'รออนุมัติ'});
    return (Array.isArray(tasks)?tasks.length:0) + (Array.isArray(approvalTasks)?approvalTasks.length:0);
  } catch(e) { return 0; }
}

// ============================================================
// STAFF STATS — สถิติการปฏิบัติงานรายบุคคล
// ============================================================
function getStaffStats(days) {
  // session check ผ่าน _autoRefreshSession อัตโนมัติ
  try {
    const sess = _sess();
    const now  = new Date();
    days = parseInt(days) || 30;
    const fromDate = new Date(now);
    fromDate.setDate(now.getDate() - days + 1);
    fromDate.setHours(0,0,0,0);

    // Records
    const recSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_DATA);
    const lr = recSheet ? recSheet.getLastRow() : 0;
    const recData = lr < 2 ? [] : recSheet.getRange(2,1,lr-1,30).getValues();

    // กรองตาม period
    const periodRec = recData.filter(r => {
      if (!r[0]) return false;
      const d = r[1] ? new Date(r[1]) : null;
      return d && d >= fromDate;
    });

    // Audit log สำหรับนับอีเมล
    const auditSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Audit Log');
    const auditData  = auditSheet && auditSheet.getLastRow() > 1
      ? auditSheet.getRange(2,1,auditSheet.getLastRow()-1,5).getValues() : [];
    const emailCount = auditData.filter(r => {
      const d = r[0] ? new Date(r[0]) : null;
      return d && d >= fromDate && (r[1]===sess.name||r[2]===sess.email) && (r[3]||'').includes('ส่งอีเมล');
    }).length;

    // สถิติรายวัน — แสดงทุกวันใน range
    const daysToShow = Math.min(days, 30); // แสดงสูงสุด 30 bars
    const step = days <= 30 ? 1 : Math.ceil(days/30);
    const daily = [];
    for (let i = daysToShow-1; i >= 0; i--) {
      const d = new Date(now); d.setDate(now.getDate()-(i*step));
      const ds = Utilities.formatDate(d,'Asia/Bangkok','dd/MM');
      const dayStart = new Date(d); dayStart.setHours(0,0,0,0);
      const dayEnd   = new Date(d); dayEnd.setHours(23,59,59,999);
      const dayRec = recData.filter(r => {
        const rd = r[1] ? new Date(r[1]) : null;
        return rd && rd >= dayStart && rd <= dayEnd;
      });
      daily.push({ date:ds, total:dayRec.length });
    }

    return {
      days: days,
      totalRecords: periodRec.length,
      countReturn:  periodRec.filter(r=>r[4]==='return').length,
      countLend:    periodRec.filter(r=>r[4]==='lend').length,
      countSpecial: periodRec.filter(r=>r[4]==='special').length,
      countEmail:   emailCount,
      daily:        daily,
      recorder:     sess.name||sess.email,
    };
  } catch(e) { return { error:e.message }; }
}


// getTaskStats — นับงานค้างสำหรับ badge
function getTaskStats() {
  try {
    const tasks = getTasks({});
    if (!Array.isArray(tasks)) return { pending:0, total:0 };
    const pending = tasks.filter(function(t){ return t.status==='รอดำเนินการ'||t.status==='รออนุมัติ'; }).length;
    return { pending:pending, total:tasks.length };
  } catch(e) { return { pending:0, total:0, error:e.message }; }
}

// Reset รหัสผ่านทุก user เป็น stou1234 (รันเมื่อ login ไม่ได้)
function resetAllPasswords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SH_USERS);
  if (!sh) { Logger.log('ไม่พบ sheet ผู้ใช้งาน'); return; }
  const data = sh.getDataRange().getValues();
  const newHash = hashPw('stou1234');
  let count = 0;
  for (let i=1; i<data.length; i++) {
    if (data[i][0]) {
      sh.getRange(i+1, 2).setValue(newHash);
      // ทำให้ active
      sh.getRange(i+1, 5).setValue(true);
      Logger.log('Reset: ' + data[i][0]);
      count++;
    }
  }
  Logger.log('DONE: reset ' + count + ' users รหัสผ่านใหม่ = stou1234');
}
// ============================================================
// DEBUG / MAINTENANCE TOOLS
// ============================================================

// ทดสอบ connection โดยไม่ต้อง session
function testConnection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SH_DATA);
  const lr = sh ? sh.getLastRow() : 0;
  const email = Session.getEffectiveUser().getEmail();
  const sessRaw = PropertiesService.getScriptProperties().getProperty('sess_'+Session.getEffectiveUser().getEmail());
  let sessInfo = 'none';
  let sessValid = false;
  if (sessRaw) {
    try {
      const s = JSON.parse(sessRaw);
      if (s && s.email) {
        sessInfo = s.email + ' / ' + s.role + ' (valid: '+(Date.now()-s.ts < 8*60*60*1000)+')';
        sessValid = true;
      } else {
        sessInfo = 'MALFORMED: ' + sessRaw.substring(0,50);
      }
    } catch(e) {
      sessInfo = 'INVALID JSON: ' + sessRaw.substring(0,50);
    }
  }
  Logger.log('email='+email+' lastRow='+lr+' sessValid='+sessValid);
  Logger.log('ALERT: '+
    '🔍 Connection Test\n\n'+
    'Email: '+email+'\n'+
    'Data rows: '+(lr>0?lr-1:0)+' records\n'+
    'Session: '+sessInfo+'\n\n'+
    (sessValid ? '✅ พร้อมใช้งาน' : '❌ Session ไม่ valid — รัน forceResetAdminSession ก่อน')
  );
}

// รันฟังก์ชันนี้ใน Apps Script เพื่อ reset session ของ admin
function forceResetAdminSession() {
  try {
    const scriptEmail = Session.getEffectiveUser().getEmail();
    const sp = PropertiesService.getScriptProperties();
    const users = [];

    // Set session สำหรับทุก user ใน Sheet ผู้ใช้งาน
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_USERS);
    if (sh && sh.getLastRow() > 1) {
      const rows = sh.getDataRange().getValues();
      for (let i=1; i<rows.length; i++) {
        const uEmail = (rows[i][0]||'').trim();
        const uRole  = (rows[i][2]||'staff').toString().trim();
        const uName  = (rows[i][3]||uEmail).toString().trim();
        if (uEmail) {
          const sess = {email:uEmail, role:uRole, name:uName, ts:Date.now()};
          sp.setProperty('sess_'+uEmail, JSON.stringify(sess));
          users.push(uEmail + ' (' + uRole + ')');
        }
      }
    }

    // Set session สำหรับ script owner (stou.post) ด้วย
    const ownerSess = {email:scriptEmail, role:'superadmin', name:'Script Owner', ts:Date.now()};
    sp.setProperty('sess_'+scriptEmail, JSON.stringify(ownerSess));

    Logger.log('ALERT: '+
      '✅ Reset session สำเร็จ\n\n' +
      'Script runs as: ' + scriptEmail + '\n' +
      'Reset สำหรับ ' + users.length + ' user:\n' +
      users.join('\n') + '\n\n' +
      'กด Deploy → เวอร์ชันใหม่ → Deploy\nแล้ว reload หน้าเว็บ'
    );
  } catch(e) {
    Logger.log('ALERT: '+'❌ Error: ' + e.message);
  }
}

// ทดสอบว่า session ทำงานไหม
function testGetRecords() {
  const result = getRecords({});
  if (result.error) {
    Logger.log('ERROR: '+result.error);
    Logger.log('ALERT: '+'❌ getRecords error: '+result.error);
  } else {
    Logger.log('SUCCESS: '+result.length+' records');
    Logger.log('ALERT: '+'✅ getRecords สำเร็จ: '+result.length+' รายการ');
  }
}

// Simulate web app call — รันนี้แล้วดูผล
function testDirectLoad() {
  var result = getRecords({});
  Logger.log('Type: ' + typeof result);
  Logger.log('IsArray: ' + Array.isArray(result));
  if (Array.isArray(result)) {
    Logger.log('Length: ' + result.length);
    if (result.length > 0) {
      Logger.log('First record: ' + JSON.stringify(result[0]));
    }
  } else {
    Logger.log('Result: ' + JSON.stringify(result));
  }
}
// ============================================================
// FILE UPLOAD TO GOOGLE DRIVE (สำหรับผลสอบสวน)
// ============================================================
const INVEST_FOLDER_NAME = 'STOU_Invest_Attachments';

function uploadInvestAttachment(base64Data, filename, mimeType, investId) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    // หา/สร้างโฟลเดอร์
    let folder;
    const folders = DriveApp.getFoldersByName(INVEST_FOLDER_NAME);
    if (folders.hasNext()) folder = folders.next();
    else folder = DriveApp.createFolder(INVEST_FOLDER_NAME);
    
    // decode base64 → Blob
    const parts = base64Data.split(',');
    const data  = parts.length > 1 ? parts[1] : parts[0];
    const bytes = Utilities.base64Decode(data);
    const blob  = Utilities.newBlob(bytes, mimeType||'application/octet-stream', filename);
    
    // สร้างไฟล์ใน folder — ตั้งชื่อให้มี investId
    const safeFn = (investId||'') + '_' + (filename||'file');
    const file   = folder.createFile(blob).setName(safeFn);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return {
      success: true,
      url:   file.getUrl(),
      id:    file.getId(),
      name:  safeFn
    };
  } catch(e) {
    return { success:false, error:e.message };
  }
}

// ============================================================
// SEND INVESTIGATION EMAIL
// ============================================================
function sendInvestEmail(data) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sess = _sess();
    const toEmail = (data.to||'').toString().trim();
    const ccEmail = (data.cc||'').toString().trim();
    const subject = data.subject || 'ขอสอบสวน';
    const body = data.body || '';
    const fromName = data.fromName || (sess && sess.name) || 'เจ้าหน้าที่';
    const fromEmail = data.fromEmail || (sess && sess.email) || '';

    if (!toEmail || !/\S+@\S+\.\S+/.test(toEmail)) {
      return { success:false, error:'อีเมลผู้รับไม่ถูกต้อง' };
    }

    const opts = {
      name: fromName + ' — สำนักบริการการศึกษา มสธ.'
    };
    if (ccEmail && /\S+@\S+\.\S+/.test(ccEmail)) {
      opts.cc = ccEmail;
    }
    if (fromEmail && /\S+@\S+\.\S+/.test(fromEmail)) {
      opts.replyTo = fromEmail;
    }

    // ส่งอีเมล
    GmailApp.sendEmail(toEmail, subject, body, opts);

    // บันทึก audit log
    logAudit('ส่งอีเมลสอบสวน', 'ถึง: ' + toEmail + ' | ' + subject.substring(0,60));

    return { success:true };
  } catch(e) {
    return { success:false, error:e.message };
  }
}

// ============================================================
// SEND POST EMAIL (LEND EMAIL)
// ============================================================
function sendPostEmail(data) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sess = _sess();
    const toEmail = (data.to||'').toString().trim();
    const ccEmail = (data.cc||'').toString().trim();
    const subject = data.subject || '';
    const body = data.body || '';
    const fromName = data.fromName || (sess && sess.name) || 'เจ้าหน้าที่';
    const fromEmail = data.fromEmail || (sess && sess.email) || '';

    if (!toEmail || !/\S+@\S+\.\S+/.test(toEmail)) {
      return { success:false, error:'อีเมลผู้รับไม่ถูกต้อง' };
    }

    const opts = {
      name: fromName + ' — สำนักบริการการศึกษา มสธ.'
    };
    if (ccEmail && /\S+@\S+\.\S+/.test(ccEmail)) {
      opts.cc = ccEmail;
    }
    if (fromEmail && /\S+@\S+\.\S+/.test(fromEmail)) {
      opts.replyTo = fromEmail;
    }

    // ส่งอีเมล
    GmailApp.sendEmail(toEmail, subject, body, opts);

    // บันทึก audit log
    logAudit('ส่งอีเมลไปรษณีย์', 'ถึง: ' + toEmail + ' | ' + subject.substring(0,60));

    return { success:true };
  } catch(e) {
    return { success:false, error:e.message };
  }
}

// ============================================================
// CRM FILE UPLOAD
// ============================================================
function uploadCrmAttachment(filename, base64content, crmId) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const folder = DriveApp.getFoldersByName('CRM Attachments').hasNext()
      ? DriveApp.getFoldersByName('CRM Attachments').next()
      : DriveApp.getRootFolder().createFolder('CRM Attachments');

    const blob = Utilities.newBlob(Utilities.base64Decode(base64content), getMimeType(filename), filename);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    logAudit('อัปโหลดไฟล์ CRM', 'ID: ' + crmId + ' | ไฟล์: ' + filename);
    return { success:true, fileUrl:file.getUrl(), fileId:file.getId() };
  } catch(e) {
    return { success:false, error:e.message };
  }
}

function getMimeType(filename) {
  const ext = filename.substring(filename.lastIndexOf('.')).toLowerCase();
  const types = {
    '.pdf':'application/pdf',
    '.doc':'application/msword',
    '.docx':'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    '.xls':'application/vnd.ms-excel',
    '.xlsx':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    '.jpg':'image/jpeg', '.jpeg':'image/jpeg',
    '.png':'image/png',
    '.gif':'image/gif',
    '.txt':'text/plain'
  };
  return types[ext] || 'application/octet-stream';
}

// ============================================================
// UPDATE RECORD FIELDS (for list page edit modal)
// ============================================================
function updateRecordFields(id, updates) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_DATA);
    if (!sheet) return { success:false, error:'ไม่พบ Sheet' };
    const lr = sheet.getLastRow();
    if (lr < 2) return { success:false, error:'ไม่พบข้อมูล' };
    const ids = sheet.getRange(2,1,lr-1,1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === String(id).trim()) {
        const row = i + 2;
        // Column mapping based on field names (0-indexed in getRecords, 1-indexed in Sheets)
        const fieldColMap = {
          'date': 2,           // col 1 in sheets
          'term': 3,           // col 2
          'year': 4,           // col 3
          'recType': 5,        // col 4
          'parcelType': 6,     // col 5
          'courseCode': 7,     // col 6
          'studentId': 8,      // col 7
          'prefix': 9,         // col 8
          'firstName': 10,     // col 9
          'lastName': 11,      // col 10
          'houseNo': 12,       // col 11
          'street': 13,        // col 12
          'subDistrict': 14,   // col 13
          'district': 15,      // col 14
          'province': 16,      // col 15
          'zipCode': 17,       // col 16
          'phone': 18,         // col 17
          'cause': 19,         // col 18
          'contactStatus': 20, // col 19
          'send1Track': 21,    // col 20
          'send1Date': 22,     // col 21
          'send2Track': 23,    // col 22
          'send2Date': 24,     // col 23
          'tags': 25,          // col 24
          'remark': 26,        // col 25
          'courses': 27,       // col 26
          'status': 28         // col 27
        };

        // Update each field provided in updates
        for (const [field, value] of Object.entries(updates)) {
          const col = fieldColMap[field];
          if (col) {
            sheet.getRange(row, col).setValue(value || '');
          }
        }

        // Update timestamp
        sheet.getRange(row, 29).setValue(new Date());

        // Log audit
        const auditMsg = Object.entries(updates).map(([k,v]) => k + ': ' + (String(v)||'-').substring(0,30)).join(' | ');
        logAudit('แก้ไขรายการ', id + ' | ' + auditMsg);
        return { success:true };
      }
    }
    return { success:false, error:'ไม่พบรายการ ' + id };
  } catch(e) { return { success:false, error:e.message }; }
}

function updateRecordField(id, field, value) {
  return updateRecordFields(id, { [field]: value });
}

// ============================================================
// EXTERNAL STAFF — ลงทะเบียน / อนุมัติ / login / dashboard
// ============================================================
const SH_EXT = 'เจ้าหน้าที่ภายนอก';

function _getExtSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SH_EXT);
  if (!sh) {
    sh = ss.insertSheet(SH_EXT);
    sh.appendRow(['อีเมล','ชื่อ','หน่วยงาน','เบอร์โทร','สถานะ','รหัสผ่าน(hash)','token','วันที่สมัคร','วันที่อนุมัติ','อนุมัติโดย']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function registerExternalStaff(data) {
  try {
    const sh = _getExtSheet();
    const email = (data.email||'').toLowerCase().trim();
    if (!email || !data.name) return { success:false, error:'กรุณากรอกชื่อและอีเมล' };
    if (!data.password || data.password.length < 8) return { success:false, error:'รหัสผ่านต้องมีอย่างน้อย 8 ตัวอักษร' };
    // ตรวจซ้ำ
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][0]||'').toLowerCase().trim() === email) {
        return { success:false, error:'อีเมลนี้ลงทะเบียนไว้แล้ว' };
      }
    }
    // Hash the password and store it during registration
    const hashedPassword = hashPw(data.password);
    sh.appendRow([email, data.name||'', data.org||'', data.phone||'', 'pending', hashedPassword, '', new Date(), '', '']);

    // ส่งอีเมลยืนยันให้ผู้สมัคร
    try {
      GmailApp.sendEmail(email,
        '✅ ได้รับคำขอลงทะเบียนสำเร็จแล้ว',
        'สวัสดีค่ะ ' + (data.name||'') + '\n\n'
        + 'ขอขอบคุณที่ลงทะเบียนเป็นเจ้าหน้าที่ภายนอกกับระบบจัดการรายไปรษณีย์ของมสธ.\n\n'
        + '📋 ข้อมูลที่ลงทะเบียน:\n'
        + '• ชื่อ: ' + (data.name||'-') + '\n'
        + '• อีเมล: ' + email + '\n'
        + '• หน่วยงาน: ' + (data.org||'-') + '\n'
        + '• เบอร์โทร: ' + (data.phone||'-') + '\n\n'
        + '⏳ ขั้นตอนต่อไป:\n'
        + 'เจ้าหน้าที่ของระบบจะตรวจสอบและอนุมัติการลงทะเบียนของท่านโดยเร็วที่สุด\n'
        + 'เมื่ออนุมัติแล้ว ท่านจะสามารถเข้าระบบได้ทันที\n\n'
        + 'หากมีข้อสงสัยติดต่อเจ้าหน้าที่ระบบได้ที่ 02 504 7623, 7626\n\n'
        + 'ด้วยความเคารพ\n'
        + 'ระบบจัดการเอกสารการสอน มสธ.');
    } catch(e2) {}

    // แจ้ง admin ทาง email (ถ้ามี admin email ใน settings)
    try {
      const settings = getSettings();
      const adminEmail = (settings && settings.adminEmail) || Session.getEffectiveUser().getEmail();
      if (adminEmail) {
        GmailApp.sendEmail(adminEmail,
          '[มสธ.] คำขอลงทะเบียนเจ้าหน้าที่ภายนอก: ' + data.name,
          'มีคำขอลงทะเบียนใหม่จาก:\nชื่อ: ' + data.name + '\nอีเมล: ' + email + '\nหน่วยงาน: ' + (data.org||'-') + '\nเบอร์โทร: ' + (data.phone||'-') + '\n\nกรุณาเข้าระบบเพื่ออนุมัติการลงทะเบียน');
      }
    } catch(e2) {}
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

function approveExternalStaff(email) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sh = _getExtSheet();
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][0]||'').toLowerCase().trim() === email.toLowerCase().trim()) {
        if (rows[i][4] === 'active') return { success:false, error:'อนุมัติไปแล้ว' };
        const sess = _sess();
        // Password is already set during registration, so just activate the account
        sh.getRange(i+1, 5).setValue('active');
        sh.getRange(i+1, 9).setValue(new Date());
        sh.getRange(i+1, 10).setValue(sess.name||sess.email||'admin');
        // Send approval notification email
        try {
          GmailApp.sendEmail(email,
            '[มสธ.] ✅ อนุมัติการลงทะเบียนแล้ว',
            'ท่านได้รับการอนุมัติให้เข้าใช้ระบบ Dashboard เจ้าหน้าที่ภายนอก มสธ.\n\nท่านสามารถเข้าสู่ระบบได้ทันที โดยใช้อีเมลและรหัสผ่านที่ท่านตั้งไว้ตอนลงทะเบียน\n\nสนใจติดต่อระบบ:\nขอแสดงความนับถือ\nผู้ดูแลระบบ มสธ.');
        } catch(e2) {}
        logAudit('อนุมัติเจ้าหน้าที่ภายนอก', email);
        return { success:true };
      }
    }
    return { success:false, error:'ไม่พบอีเมลนี้' };
  } catch(e) { return { success:false, error:e.message }; }
}

function rejectExternalStaff(email) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sh = _getExtSheet();
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][0]||'').toLowerCase().trim() === email.toLowerCase().trim()) {
        sh.getRange(i+1, 5).setValue('rejected');
        logAudit('ปฏิเสธเจ้าหน้าที่ภายนอก', email);
        return { success:true };
      }
    }
    return { success:false, error:'ไม่พบอีเมลนี้' };
  } catch(e) { return { success:false, error:e.message }; }
}

function setExternalStaffPassword(email, token, newPassword) {
  try {
    const sh = _getExtSheet();
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][0]||'').toLowerCase().trim() === email.toLowerCase().trim()) {
        if (rows[i][6] !== token) return { success:false, error:'token ไม่ถูกต้อง' };
        if (!['approved','active'].includes(rows[i][4])) return { success:false, error:'บัญชีนี้ยังไม่ได้รับการอนุมัติ' };
        sh.getRange(i+1, 5).setValue('active');
        sh.getRange(i+1, 6).setValue(hashPw(newPassword));
        sh.getRange(i+1, 7).setValue(''); // clear token
        return { success:true };
      }
    }
    return { success:false, error:'ไม่พบอีเมลนี้' };
  } catch(e) { return { success:false, error:e.message }; }
}

function loginExternalStaff(email, password) {
  try {
    const sh = _getExtSheet();
    const rows = sh.getDataRange().getValues();
    const emailL = (email||'').toLowerCase().trim();
    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][0]||'').toLowerCase().trim() === emailL) {
        const status = rows[i][4]||'pending';
        if (status === 'pending') return { success:true, user:{email:emailL, name:rows[i][1], org:rows[i][2], status:'pending'} };
        if (status === 'rejected') return { success:false, error:'บัญชีนี้ถูกปฏิเสธ' };
        if (status !== 'active') return { success:false, error:'บัญชียังไม่ active' };
        if (!rows[i][5]) return { success:false, error:'ยังไม่ได้ตั้งรหัสผ่าน กรุณาตรวจสอบอีเมล' };
        if (hashPw(password) !== rows[i][5]) return { success:false, error:'รหัสผ่านไม่ถูกต้อง' };
        return { success:true, user:{email:emailL, name:rows[i][1], org:rows[i][2], status:'active'} };
      }
    }
    return { success:false, error:'ไม่พบอีเมลนี้ในระบบ' };
  } catch(e) { return { success:false, error:e.message }; }
}

function getExternalStaffTickets(email) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CRM);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const rows = sheet.getRange(2,1,sheet.getLastRow()-1,21).getValues();
    const result = [];
    const emailL = (email||'').toLowerCase().trim();
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      if ((r[4]||'').toLowerCase().trim() === emailL || (r[2]||'').toLowerCase().trim() === emailL) {
        result.push({
          id:String(r[0]), date:String(r[1]), reporterName:String(r[2]),
          issueType:String(r[11]), detail:String(r[12]),
          status:String(r[16]), replies:String(r[19])||'[]'
        });
      }
    }
    return result.reverse();
  } catch(e) { return []; }
}

function getPendingExternalRegistrations() {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sh = _getExtSheet();
    if (sh.getLastRow() < 2) return { success:true, rows:[] };
    const rows = sh.getRange(2,1,sh.getLastRow()-1,10).getValues();
    return { success:true, rows: rows.map(function(r){
      return { email:r[0], name:r[1], org:r[2], phone:r[3], status:r[4], registeredAt:r[7]?String(r[7]):'', approvedAt:r[8]?String(r[8]):'', approvedBy:r[9] };
    })};
  } catch(e) { return { success:false, error:e.message }; }
}

function getExternalStaffDetail(email) {
  if (!_autoRefreshSession()) return null;
  try {
    const sh = _getExtSheet();
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][0]||'').toLowerCase() === email.toLowerCase()) {
        return {
          email: rows[i][0],
          name: rows[i][1],
          org: rows[i][2],
          phone: rows[i][3],
          status: rows[i][4]
        };
      }
    }
    return null;
  } catch(e) { return null; }
}

function updateExternalStaff(email, data) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sh = _getExtSheet();
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][0]||'').toLowerCase() === email.toLowerCase()) {
        if (data.name) sh.getRange(i+1, 2).setValue(data.name);
        if (data.org) sh.getRange(i+1, 3).setValue(data.org);
        if (data.phone) sh.getRange(i+1, 4).setValue(data.phone);
        logAudit('อัปเดตเจ้าหน้าที่ภายนอก', email + ' | ' + (data.name||rows[i][1]));
        return { success:true };
      }
    }
    return { success:false, error:'ไม่พบผู้ใช้' };
  } catch(e) { return { success:false, error:e.message }; }
}

function deleteExternalStaff(email) {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sh = _getExtSheet();
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][0]||'').toLowerCase() === email.toLowerCase()) {
        sh.deleteRow(i+1);
        logAudit('ลบเจ้าหน้าที่ภายนอก', email);
        return { success:true };
      }
    }
    return { success:false, error:'ไม่พบผู้ใช้' };
  } catch(e) { return { success:false, error:e.message }; }
}

function diagnosticCheckRows315To329() {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_DATA);
    if (!sheet) return { success:false, error:'ไม่พบ Sheet' };

    const diagnosis = [];
    // Check rows 315-329 (which are rows 316-330 in the sheet, accounting for header)
    const startRow = 316; // row 315 in data (row 316 in sheet including header)
    const endRow = 330;

    for (let row = startRow; row <= endRow; row++) {
      const values = sheet.getRange(row, 1, 1, 35).getValues()[0];
      const rowNum = row - 1; // Convert to data row number

      // Get column data (0-indexed)
      const id = String(values[0] || '');
      const courseCode = String(values[6] || '');
      const studentId = String(values[7] || '');
      const firstName = String(values[9] || '');
      const lastName = String(values[10] || '');

      // Column AA (27 in 1-indexed) = index 26 in 0-indexed array
      const coursesAA = String(values[26] || '');

      // Columns that might have shifted data (AE-AH = indices 30-33)
      const columnAE = String(values[30] || '');
      const columnAF = String(values[31] || '');
      const columnAG = String(values[32] || '');
      const columnAH = String(values[33] || '');

      // Parse courses JSON if it exists
      let coursesInfo = 'ERROR: no JSON found';
      let coursesHasTrack = false;
      try {
        const cs = JSON.parse(coursesAA || '[]');
        if (cs.length > 0) {
          coursesInfo = 'JSON_OK: ' + cs.length + ' courses';
          // Check if courses have track property
          coursesHasTrack = cs.some(function(c) { return c.track; });
        } else {
          coursesInfo = 'JSON_EMPTY: empty array';
        }
      } catch(e) {
        coursesInfo = 'JSON_INVALID: ' + e.message;
      }

      diagnosis.push({
        row: rowNum,
        id: id,
        studentId: studentId,
        firstName: firstName,
        lastName: lastName,
        courseCode: courseCode,
        coursesAALength: coursesAA.length,
        coursesAAInfo: coursesInfo,
        coursesAAHasTrack: coursesHasTrack,
        coursesAAPreview: coursesAA.substring(0, 150),
        hasShiftedData: (columnAE.startsWith('[') && !coursesAA.trim()),
        columnAEPreview: columnAE.substring(0, 100),
        columnAFPreview: columnAF.substring(0, 50),
        columnAGPreview: columnAG.substring(0, 50),
        columnAHPreview: columnAH.substring(0, 50)
      });
    }

    return { success:true, rows: diagnosis };
  } catch(e) {
    Logger.log('diagnosticCheckRows315To329 error: ' + e.message);
    return { success:false, error:e.message };
  }
}

function runDiagnosticsAndRecover() {
  if (!_autoRefreshSession()) return { success:false, error:'SESSION_EXPIRED' };
  try {
    // First, run diagnostics
    const diag = diagnosticCheckRows315To329();
    if (!diag.success) return diag;

    // Count rows with shifted data
    const shiftedRows = diag.rows.filter(function(r) { return r.hasShiftedData; });
    const needsRecovery = shiftedRows.length > 0;

    let recoveryResult = { skipped: true };
    if (needsRecovery) {
      // Run recovery for shifted columns
      recoveryResult = recoverShiftedColumns();
    }

    return {
      success: true,
      totalRowsChecked: diag.rows.length,
      rowsWithShiftedData: shiftedRows.length,
      shiftedRowNumbers: shiftedRows.map(function(r) { return r.row; }),
      diagnostics: diag.rows,
      recoveryResult: recoveryResult
    };
  } catch(e) {
    Logger.log('runDiagnosticsAndRecover error: ' + e.message);
    return { success:false, error:e.message };
  }
}

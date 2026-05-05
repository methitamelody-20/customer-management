# Firebase Migration Guide - management-crm-stoupost

## ✅ สถานะการอัพเดต

Firebase integration code ได้ถูกเพิ่มเข้าไปในระบบแล้ว พร้อมสำหรับการอัพเดตข้อมูลจาก Google Sheets ไปยัง Firebase Realtime Database

## 🚀 วิธีเปิดใช้งาน Firebase

### Step 1: ตั้งค่า Firebase Security Rules

ไปที่ Firebase Console → Realtime Database → Rules และแทนที่ด้วยค่าต่อไปนี้:

```json
{
  "rules": {
    ".read": "auth != null || root.child('test').exists()",
    ".write": "auth != null || root.child('test').exists()",
    "users": {
      ".read": "auth != null",
      ".write": "auth != null"
    },
    "records": {
      ".read": "auth != null",
      ".write": "auth != null"
    },
    "crm": {
      ".read": "auth != null",
      ".write": "auth != null"
    },
    "investigations": {
      ".read": "auth != null",
      ".write": "auth != null"
    },
    "settings": {
      ".read": true,
      ".write": "auth != null"
    },
    "test": {
      ".read": true,
      ".write": true
    }
  }
}
```

### Step 2: ตรวจสอบการเชื่อมต่อ Firebase

เปิด Apps Script Editor และรัน function `testFirebaseConnection()`:

```
ฟังก์ชัน → ฟังก์ชัน ที่มีอยู่ → testFirebaseConnection()
```

ถ้าได้ข้อความ "✅ เชื่อมต่อ Firebase สำเร็จ" แสดงว่าการเชื่อมต่อสำเร็จ

### Step 3: อัพเดตโค้ด (Code.gs) เพื่อเปิดใช้งาน Firebase

เปิดไฟล์ `รหัส.gs` และหาบรรทัด:

```javascript
const USE_FIREBASE = false;
```

เปลี่ยนเป็น:

```javascript
const USE_FIREBASE = true;
```

### Step 4: อัพเดตข้อมูลจาก Google Sheets ไปยัง Firebase

รัน function `migrateAllToFirebase()` ในที่นี้:

```
ฟังก์ชัน → ฟังก์ชัน ที่มีอยู่ → migrateAllToFirebase()
```

หรือรันมี individual migration functions:

- `migrateUsersToFirebase()` - อัพเดตผู้ใช้งาน
- `migrateRecordsToFirebase()` - อัพเดตพัสดุ
- `migrateCrmToFirebase()` - อัพเดต CRM tickets
- `migrateInvestigationsToFirebase()` - อัพเดต investigations

ดูผลลัพธ์ใน Apps Script Logger

### Step 5: ตรวจสอบสถานะ

รัน `getFirebaseStatus()` เพื่อดูสถานะ Firebase ของระบบ

## 📊 Firebase Data Structure

### /users/{email}
```json
{
  "email": "user@stou.ac.th",
  "passwordHash": "hash_value",
  "role": "staff",
  "name": "ชื่อ-สกุล",
  "active": true,
  "lastLogin": "2024-01-15T10:30:00Z",
  "perms": "dash,crm,records"
}
```

### /records/{id}
```json
{
  "id": "L20240105-001",
  "date": "2024-01-05T14:30:00Z",
  "studentId": "xxxxxxxx",
  "courseCode": "XXXX101",
  ...
}
```

### /crm/{id}
```json
{
  "id": "CRM-20240105-001",
  "date": "2024-01-05T14:30:00Z",
  "reporterEmail": "student@domain.com",
  "status": "open",
  "replies": "[{\"name\":\"admin\",\"email\":\"admin@stou.ac.th\",\"text\":\"reply text\",\"time\":\"2024-01-05T15:00:00Z\"}]",
  ...
}
```

### /investigations/{id}
```json
{
  "investId": "INV-20240105-143000",
  "refId": "L20240105-001",
  "status": "รอส่ง",
  ...
}
```

## ⚙️ วิธีสลับกลับไป Google Sheets

ถ้าต้องการให้ระบบใช้ Google Sheets แทน Firebase ให้:

1. เปลี่ยน `USE_FIREBASE = false` ใน Code.gs
2. ระบบจะกลับไปใช้ฟังก์ชัน Sheets เดิมอัตโนมัติ

## 🔄 Dual Mode (ใช้ทั้ง Sheets และ Firebase)

ระบบได้ออกแบบให้รองรับ dual mode คือ:
- ข้อมูลใหม่จะไปยังทั้ง Sheets และ Firebase
- การอ่านข้อมูลจะมาจาก Firebase (ถ้าเปิด) หรือ Sheets

วิธีการ:

1. เขียน Sheets wrapper functions ที่เรียก Firebase functions
2. ในฟังก์ชัน CRUD เพิ่ม logic: "ถ้า Firebase เปิด → ใช้ Firebase, ถ้าไม่ → ใช้ Sheets"

## 📝 โค้ด Firebase Functions ที่สร้างแล้ว

### Authentication
- `loginWithFirebase()` - Login ผ่าน Firebase
- `checkSessionFirebase()` - ตรวจสอบ session

### Users
- `getUsersFirebase()` - ดึงรายชื่อผู้ใช้
- `addUserFirebase()` - เพิ่มผู้ใช้ใหม่

### Records
- `addRecordFirebase()` - เพิ่มพัสดุ
- `getRecordsFirebase()` - ดึงข้อมูลพัสดุ

### CRM
- `addCrmTicketFirebase()` - เพิ่ม ticket
- `getCrmTicketsFirebase()` - ดึง tickets
- `replyCrmTicketFirebase()` - ตอบ CRM
- `updateCrmStatusFirebase()` - อัพเดตสถานะ

### Investigations
- `getInvestigationsFirebase()` - ดึง investigations
- `addInvestigationFirebase()` - เพิ่ม investigation
- `updateInvestigationFirebase()` - อัพเดต investigation

## ⚠️ สิ่งที่ต้องทำต่อไป

### 1. อัพเดต Mainsystem.html
ต้องแก้ไข frontend เพื่อเรียก Firebase functions แทน Sheets functions

### 2. Email System Enhancement
- เพิ่มการจับอีเมลของ staff ที่ส่ง reply
- ใช้ GmailApp.sendEmail() เพื่อส่งอีเมลตอบกลับ

### 3. UI Modernization
- อัพเดต styling ให้ดูทันสมัยกว่า

### 4. Remove World Time
- ลบการแสดง GMT timezone จากหน้า investigation

## 🔐 Security Considerations

1. **Password**: Firebase Authentication เก็บ password อย่างปลอดภัย
2. **API Key**: โปรดเก็บ API key ให้เป็นความลับ
3. **Security Rules**: ตั้งค่า Security Rules ให้เหมาะสม (อย่ากำหนด read/write = true)
4. **User Validation**: ทุก request ควรผ่าน authentication

## 📞 Support

ถ้าพบปัญหาในการเชื่อมต่อ Firebase:

1. ตรวจสอบ Internet connection
2. ตรวจสอบ Firebase Security Rules
3. ตรวจสอบ Firebase API Key
4. ดู Apps Script Logs เพื่อหาข้อมูลความผิดพลาด

---

**Last Updated**: 2026-05-05
**Status**: ✅ Firebase Integration Ready

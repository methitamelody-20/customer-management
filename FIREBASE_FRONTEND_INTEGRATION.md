# Firebase Frontend Integration Guide

## Overview

The Mainsystem.html has been updated to support Firebase Realtime Database integration while maintaining fallback support for Google Sheets.

## What's Been Added

### 1. Firebase SDK (Module-based)
- Added Firebase Web SDK v10.7.0 imports in the `<head>` section
- Configured with your `management-crm-stoupost` project
- Exposes Firebase modules to global `window.firebase` object

### 2. Firebase API Wrapper (`firebaseAPI`)
- Provides unified interface for calling both Firebase and Apps Script functions
- Automatically detects which backend to use based on `useFirebase` flag
- Includes methods:
  - `checkSession()` - Verify user authentication
  - `addCrmTicket(data)` - Create new CRM ticket
  - `getCrmTickets(filters)` - Fetch CRM tickets
  - `replyCrmTicket(id, text, adminName, adminEmail)` - Add reply to ticket
  - `getRecords(filters)` - Fetch package records

### 3. Fallback Architecture
- All functions fall back to `google.script.run` if Firebase is unavailable
- No data loss if Firebase is disabled
- System works seamlessly with either backend

## How to Use

### Enable Firebase Frontend Integration

1. **In Firebase Console**:
   - Ensure Security Rules are published
   - Check that database has data from migration

2. **In Mainsystem.html** (around line 1882):
   ```javascript
   let useFirebase = false; // Change this to true
   ```

3. **Refresh the web app**

### Firebase API Usage Examples

#### Check Session
```javascript
const session = await firebaseAPI.checkSession();
if (session.valid) {
  console.log('User:', session.name);
}
```

#### Add CRM Ticket
```javascript
const result = await firebaseAPI.addCrmTicket({
  reporterName: 'ชื่อผู้แจ้ง',
  studentId: '12345678',
  reporterEmail: 'student@domain.com',
  issueType: 'ปัญหาทั่วไป',
  detail: 'รายละเอียดปัญหา',
  channel: 'online'
});
console.log('Created ticket:', result.id);
```

#### Get CRM Tickets
```javascript
const tickets = await firebaseAPI.getCrmTickets({});
console.log('Total tickets:', tickets.length);
```

#### Reply to CRM Ticket
```javascript
const reply = await firebaseAPI.replyCrmTicket(
  'CRM-20240105-001',
  'ข้อความตอบกลับ',
  'ชื่อผู้ตอบ',
  'admin@stou.ac.th'
);
if (reply.success) {
  console.log('Reply sent');
}
```

#### Get Records
```javascript
const records = await firebaseAPI.getRecords({
  recType: 'return',
  year: '2024'
});
console.log('Found records:', records.length);
```

## Data Format

### CRM Tickets (Firebase)
```json
{
  "id": "auto-generated",
  "date": "2024-01-05T14:30:00Z",
  "reporterName": "ชื่อผู้แจ้ง",
  "reporterEmail": "email@domain.com",
  "status": "open",
  "replies": "[{\"name\":\"admin\",\"email\":\"admin@stou.ac.th\",\"text\":\"reply\",\"time\":\"2024-01-05T15:00:00Z\"}]",
  ...
}
```

### Records (Firebase)
```json
{
  "id": "L20240105-001",
  "date": "2024-01-05T14:30:00Z",
  "studentId": "12345678",
  "courseCode": "XXXX101",
  "status": "บันทึกแล้ว",
  ...
}
```

## Important Notes

1. **Async/Await**: All firebaseAPI methods return Promises
   - Use `await` or `.then()` to handle results
   - Wrap in try/catch for error handling

2. **Error Handling**:
   ```javascript
   try {
     const result = await firebaseAPI.addCrmTicket(data);
     if (result.success) { ... }
   } catch(error) {
     console.error('Error:', error.message);
   }
   ```

3. **Email Parameter**: 
   - `replyCrmTicket()` now accepts `adminEmail` parameter
   - Used to show which staff member replied
   - Can be empty string if not available

4. **Fallback**: If `useFirebase = false`, all calls use `google.script.run`

## Update Existing Functions

To update existing Mainsystem.html functions to use Firebase wrapper:

### Before (Google Sheets only):
```javascript
function loadCRM() {
  google.script.run.getCrmTickets({}).then(tickets => {
    displayTickets(tickets);
  });
}
```

### After (Firebase + Sheets):
```javascript
async function loadCRM() {
  const tickets = await firebaseAPI.getCrmTickets({});
  displayTickets(tickets);
}
```

## Additional Functions to Add

The following functions in `firebaseAPI` are ready to be implemented:

```javascript
// These are skeleton ready - can be expanded:
// - addRecordFirebase(data)
// - updateCrmStatusFirebase(id, status)
// - getInvestigations(filters)
// - addInvestigation(data)
// - updateInvestigation(id, updates)
```

## Testing Firebase Integration

1. **Enable Firebase**:
   ```javascript
   let useFirebase = true;
   ```

2. **Open Browser Console** (F12)

3. **Test a function**:
   ```javascript
   firebaseAPI.getCrmTickets({}).then(t => console.log('Tickets:', t.length));
   ```

4. **Should see tickets logged** (from Firebase Realtime Database)

5. **Check Network Tab** to see Firebase calls (not Apps Script)

## Performance Considerations

- **Firebase**: ~100-300ms per call (depending on data size)
- **Google Sheets**: ~500-2000ms per call
- **Caching**: Consider caching data locally to reduce calls
- **Real-time**: Can add listeners for live updates

## Security

- ✅ Uses Firebase Security Rules
- ✅ User must be authenticated (session valid)
- ✅ API Key is in client code (OK - Security Rules protect data)
- ✅ No sensitive data exposed

## Troubleshooting

### Firebase Functions Not Working?

1. **Check console** (F12 → Console tab)
2. **Verify Security Rules** published in Firebase Console
3. **Ensure data exists** in Firebase Realtime Database
4. **Check useFirebase flag** is set to `true`
5. **Look for CORS errors** (shouldn't happen with Firebase SDK)

### Still Using Sheets Data?

- Verify migration completed (`migrateAllToFirebase()`)
- Check Firebase Console → Realtime Database for data
- Confirm `useFirebase = true` in Mainsystem.html

## Next Steps

1. ✅ Firebase SDK integrated
2. ✅ API wrapper created
3. ⏳ Update all function calls to use `firebaseAPI`
4. ⏳ Implement email reply system with Firebase
5. ⏳ Add real-time listeners for live updates

---

**Status**: 🟡 Frontend integration ready for enablement

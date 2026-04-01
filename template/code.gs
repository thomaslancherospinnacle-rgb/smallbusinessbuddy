function testCalendar() {
  var cal = CalendarApp.getCalendarById('{{CF_USERNAME}}@gmail.com');
  Logger.log(cal.getName());
}
function testWorkerCall() {
  var config = getConfig();
  Logger.log('Worker URL: ' + config.emailWorkerUrl);
  Logger.log('Noreply: ' + config.noreplyEmail);
  
  var response = UrlFetchApp.fetch(config.emailWorkerUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      from: '{{COMPANY_NAME}} <' + config.noreplyEmail + '>',
      to: '{{CF_USERNAME}}@gmail.com',
      subject: 'Test Auto-Reply',
      body: 'If you get this, the Worker connection works!',
      replyTo: config.noreplyEmail
    }),
    muteHttpExceptions: true
  });
  
  Logger.log('Status: ' + response.getResponseCode());
  Logger.log('Response: ' + response.getContentText());
}
function debugConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Config');
  var data = sheet.getDataRange().getValues();
  
  for (var i = 0; i < data.length; i++) {
    var key = data[i][0];
    var val = data[i][1];
    Logger.log('Row ' + (i+1) + ': key=[' + key + '] (length=' + String(key).length + ') val=[' + String(val).substring(0,50) + ']');
  }
}
/* ═══════════════════════════════════════════════════════════════
   {{COMPANY_SHORT}} — Google Apps Script Backend  (v3)
   ═══════════════════════════════════════════════════════════════
   
   WHAT'S NEW IN V3:
   - Google Calendar integration (create/update/cancel events)
   - Availability endpoint (checks real calendar for open slots)
   - Auto-reply confirmation emails via Cloudflare Worker
   - Unified "inquiry" action (replaces separate quote/contact)
   - Booking now creates calendar events automatically
   - Config-driven: business name, email, calendar ID stored in Config sheet
   
   SETUP INSTRUCTIONS:
   1. Go to https://script.google.com → New Project
   2. Paste this entire file into Code.gs
   3. Click Deploy → New Deployment → Web App
      - Execute as: Me
      - Who has access: Anyone
   4. Copy the Web App URL
   5. Paste it into BOTH:
      - index.html  → GAS_URL variable at the top of the <script>
      - admin.html  → Settings page (saves to localStorage)
   6. On first run, authorize the script when prompted
      (Grant Calendar + Gmail + Sheets permissions)
   7. In the Config sheet, set your calendarId to your Google Calendar ID
      (usually your Gmail address, or find it in Calendar Settings → Integrate)
   
   SHEETS CREATED AUTOMATICALLY:
   - "Appointments"  — All bookings from the website
   - "Quotes"        — All quote requests  
   - "Contacts"      — General contact form messages
   - "Clients"       — Client history / CRM
   - "Config"        — Business settings (editable)
   
   ═══════════════════════════════════════════════════════════════ */

// ─── SHEET NAMES ───
const SHEET_APPOINTMENTS = 'Appointments';
const SHEET_QUOTES       = 'Quotes';
const SHEET_CONTACTS     = 'Contacts';
const SHEET_CLIENTS      = 'Clients';
const SHEET_CONFIG       = 'Config';
const SHEET_REVENUE      = 'Revenue';
const SHEET_REVIEWS      = 'Reviews';
const SHEET_SERVICES     = 'Services';
const SHEET_ABOUT        = 'About';

// ─── HEADERS ───
const HEADERS = {
  [SHEET_APPOINTMENTS]: ['ID', 'Timestamp', 'Name', 'Email', 'Phone', 'Service', 'Date', 'Time', 'Address', 'Notes', 'Status', 'CalendarEventId'],
  [SHEET_QUOTES]:       ['ID', 'Timestamp', 'Name', 'Email', 'Phone', 'Service', 'Address', 'Description', 'Contact Method', 'Status'],
  [SHEET_CONTACTS]:     ['ID', 'Timestamp', 'Name', 'Email', 'Message', 'Status'],
  [SHEET_CLIENTS]:      ['ID', 'Timestamp', 'Name', 'Email', 'Phone', 'Service', 'Address', 'Notes', 'Status'],
  [SHEET_REVENUE]:      ['ID', 'Timestamp', 'Type', 'Status', 'Amount', 'ClientName', 'ClientEmail', 'Description', 'JobId', 'AppointmentId', 'StripeId', 'PaymentUrl', 'PaidAt', 'Notes'],
  [SHEET_REVIEWS]:      ['ID', 'Timestamp', 'Name', 'Rating', 'ReviewText', 'Status'],
  [SHEET_SERVICES]:     ['ID', 'Order', 'Name', 'Icon', 'Description', 'Price', 'Duration', 'Active'],
  [SHEET_ABOUT]:        ['Key', 'Value'],
};


/* ═══════════════════════════════════════
   ENTRY POINTS — doGet / doPost
   ═══════════════════════════════════════ */

function doGet(e) {
  const action = e.parameter.action || 'ping';
  
  switch (action) {
    case 'getAppointments':
      return jsonResponse(getSheetData(SHEET_APPOINTMENTS));
    
    case 'getQuotes':
      return jsonResponse(getSheetData(SHEET_QUOTES));
    
    case 'getContacts':
      return jsonResponse(getSheetData(SHEET_CONTACTS));
    
    case 'getClients':
      return jsonResponse(getSheetData(SHEET_CLIENTS));
    
    case 'getRevenue':
      return jsonResponse(getSheetData(SHEET_REVENUE));
    
    case 'getDashboard':
      return jsonResponse(getDashboardStats());
    
    case 'getConfig':
      return jsonResponse(getConfig());
    
    case 'getEmails':
      return jsonResponse(getEmails(e.parameter.view || 'inbox', parseInt(e.parameter.count) || 30));
    
    // NEW: Get available time slots for a given date
    case 'getAvailability':
      return jsonResponse(getAvailability(e.parameter.date, parseInt(e.parameter.duration) || 60));
    
    case 'getReviews':
      return jsonResponse(getReviews());

    case 'getServices':
      return jsonResponse(getServices());

    case 'getAbout':
      return jsonResponse(getAbout());

    case 'ping':
      return jsonResponse({ status: 'ok', version: 'v3', timestamp: new Date().toISOString() });
    
    default:
      return jsonResponse({ error: 'Unknown action: ' + action });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ error: 'Empty POST body' });
    }
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    Logger.log('doPost action: ' + action + ' | keys: ' + Object.keys(data).join(', '));
    
    switch (action) {
      case 'appointment':
        return jsonResponse(addAppointment(data));
      
      case 'quote':
        return jsonResponse(addQuote(data));
      
      case 'contact':
        return jsonResponse(addContact(data));
      
      // NEW: Unified inquiry (question / quote / general)
      case 'inquiry':
        return jsonResponse(addInquiry(data));
      
      case 'addClient':
        return jsonResponse(addClient(data));
      
      case 'addRevenue':
        return jsonResponse(addRevenue(data));
      
      case 'updateRevenue':
        return jsonResponse(updateRevenue(data));
      
      case 'deleteRevenue':
        return jsonResponse(deleteRevenue(data.id));
      
      case 'updateStatus':
        return jsonResponse(updateStatus(data.sheet, data.id, data.status));
      
      case 'updateRow':
        return jsonResponse(updateRow(data.sheet, data.id, data.fields));
      
      case 'deleteRow':
        return jsonResponse(deleteRow(data.sheet, data.id));
      
      case 'saveConfig':
        return jsonResponse(saveConfig(data.config));
      
      case 'sendEmail':
        return jsonResponse(sendEmailAction(data));
      
      // NEW: Cancel a calendar event linked to an appointment
      case 'cancelCalendarEvent':
        return jsonResponse(cancelCalendarEvent(data.id));
      
      // Stripe deposit confirmed — create calendar event, send confirmation, mark Confirmed
      case 'updateAppointmentStatus':
        return jsonResponse(updateAppointmentStatus(data));

      // ── Reviews ──
      case 'addReview':
        return jsonResponse(addReview(data));
      case 'updateReview':
        return jsonResponse(updateReviewStatus(data));
      case 'deleteReview':
        return jsonResponse(deleteReview(data.id));

      // ── Services (editable from settings) ──
      case 'addService':
        return jsonResponse(addService(data));
      case 'updateService':
        return jsonResponse(updateService(data));
      case 'deleteService':
        return jsonResponse(deleteService(data.id));
      case 'reorderServices':
        return jsonResponse(reorderServices(data.order));

      // ── About page content ──
      case 'saveAbout':
        return jsonResponse(saveAbout(data.fields));

      default:
        return jsonResponse({ error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}


/* ═══════════════════════════════════════
   WRITE OPERATIONS
   ═══════════════════════════════════════ */

function addAppointment(data) {
  const sheet = getOrCreateSheet(SHEET_APPOINTMENTS);

  // Use the appointmentId sent from the frontend (APPT-timestamp) so we can
  // look it up later when the deposit is confirmed. Fall back to generated ID.
  const id = data.appointmentId || generateId();

  // ── NO calendar event yet ──
  // Google Calendar event is created ONLY after the $100 deposit is paid.
  // booking-success.html calls action:'updateAppointmentStatus' with
  // depositPaid:true, which triggers createCalendarEvent + sends the
  // auto-reply confirmation email and sets status to "Confirmed".

  sheet.appendRow([
    id,
    data.timestamp || new Date().toISOString(),
    data.name || '',
    data.email || '',
    data.phone || '',
    data.service || '',
    data.date || '',
    data.time || '',
    data.address || '',
    data.notes || '',
    'Pending Payment',  // held here until Stripe deposit confirmed
    ''                  // CalendarEventId — filled in after payment
  ]);

  return { success: true, id: id, type: 'appointment' };
}

function addQuote(data) {
  const sheet = getOrCreateSheet(SHEET_QUOTES);
  const id = generateId();
  sheet.appendRow([
    id,
    data.timestamp || new Date().toISOString(),
    data.name || '',
    data.email || '',
    data.phone || '',
    data.service || '',
    data.address || '',
    data.desc || '',
    data.contactMethod || 'phone',
    'Pending'
  ]);
  
  // Send auto-reply
  var emailStatus = { sent: false, reason: 'skipped' };
  try {
    emailStatus = sendAutoReply('quote', data) || emailStatus;
  } catch (err) {
    emailStatus = { sent: false, reason: 'exception', error: err.message };
    Logger.log('Auto-reply failed: ' + err.message);
  }
  
  return { success: true, id: id, type: 'quote', email: emailStatus };
}

function addContact(data) {
  const sheet = getOrCreateSheet(SHEET_CONTACTS);
  const id = generateId();
  sheet.appendRow([
    id,
    data.timestamp || new Date().toISOString(),
    data.name || '',
    data.email || '',
    data.message || '',
    'Unread'
  ]);
  
  // Send auto-reply
  var emailStatus = { sent: false, reason: 'skipped' };
  try {
    emailStatus = sendAutoReply('contact', data) || emailStatus;
  } catch (err) {
    emailStatus = { sent: false, reason: 'exception', error: err.message };
    Logger.log('Auto-reply failed: ' + err.message);
  }
  
  return { success: true, id: id, type: 'contact', email: emailStatus };
}

/**
 * NEW: Unified inquiry handler
 * type can be: 'question', 'quote', 'callback'
 * Routes to the appropriate sheet based on type
 */
function addInquiry(data) {
  const inquiryType = (data.type || 'question').toLowerCase();
  
  if (inquiryType === 'quote') {
    // Route to Quotes sheet
    return addQuote({
      name: data.name,
      email: data.email,
      phone: data.phone,
      service: data.service || '',
      address: data.address || '',
      desc: data.message || data.desc || '',
      contactMethod: data.contactMethod || 'phone',
      timestamp: data.timestamp
    });
  } else {
    // Route to Contacts sheet (question, callback, general)
    const sheet = getOrCreateSheet(SHEET_CONTACTS);
    const id = generateId();
    var msg = data.message || '';
    if (inquiryType !== 'question') {
      msg = '[' + inquiryType.toUpperCase() + '] ' + msg;
    }
    if (data.phone) {
      msg = 'Phone: ' + data.phone + '\n' + msg;
    }
    sheet.appendRow([
      id,
      data.timestamp || new Date().toISOString(),
      data.name || '',
      data.email || '',
      msg,
      'Unread'
    ]);
    
    var emailStatus = { sent: false, reason: 'skipped' };
    try {
      emailStatus = sendAutoReply(inquiryType, data) || emailStatus;
    } catch (err) {
      emailStatus = { sent: false, reason: 'exception', error: err.message };
      Logger.log('Auto-reply failed: ' + err.message);
    }
    
    return { success: true, id: id, type: inquiryType, email: emailStatus };
  }
}

function addClient(data) {
  const sheet = getOrCreateSheet(SHEET_CLIENTS);
  const id = generateId();
  sheet.appendRow([
    id,
    data.timestamp || new Date().toISOString(),
    data.name || '',
    data.email || '',
    data.phone || '',
    data.service || '',
    data.address || '',
    data.notes || '',
    'Active'
  ]);
  return { success: true, id: id, type: 'client' };
}


/* ═══════════════════════════════════════
   REVENUE CRUD
   ═══════════════════════════════════════ */

/**
 * Add a transaction row to the Revenue sheet.
 * Called from admin.html whenever an invoice or deposit is created,
 * or when a manual transaction is added.
 *
 * Expected fields in data:
 *   type         — 'deposit' | 'invoice'
 *   status       — 'paid' | 'pending'
 *   amount       — number (dollars)
 *   clientName   — string
 *   clientEmail  — string
 *   description  — string
 *   jobId        — string (optional)
 *   appointmentId— string (optional)
 *   stripeId     — string (optional, Stripe link/price ID)
 *   paymentUrl   — string (optional, Stripe checkout URL)
 *   paidAt       — ISO string (optional, when status is 'paid')
 *   notes        — string (optional)
 */
function addRevenue(data) {
  const sheet = getOrCreateSheet(SHEET_REVENUE);
  const id = data.id || generateId();
  const now = new Date().toISOString();
  sheet.appendRow([
    id,
    data.timestamp || now,
    data.type    || 'invoice',
    data.status  || 'pending',
    parseFloat(data.amount) || 0,
    data.clientName   || '',
    data.clientEmail  || '',
    data.description  || '',
    data.jobId        || '',
    data.appointmentId|| '',
    data.stripeId     || '',
    data.paymentUrl   || '',
    data.paidAt       || (data.status === 'paid' ? now : ''),
    data.notes        || '',
  ]);
  return { success: true, id: id, type: 'revenue' };
}

/**
 * Update a Revenue row by ID.
 * Supports: status, paidAt, notes, amount, paymentUrl, stripeId
 */
function updateRevenue(data) {
  if (!data.id) return { error: 'Missing id' };
  const sheet = getOrCreateSheet(SHEET_REVENUE);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(data.id)) continue;

    const allowedFields = ['Type','Status','Amount','ClientName','ClientEmail','Description','JobId','AppointmentId','StripeId','PaymentUrl','PaidAt','Notes'];
    const updated = [];

    allowedFields.forEach(function(col) {
      const key = col.charAt(0).toLowerCase() + col.slice(1); // camelCase
      if (data[key] !== undefined) {
        const colIdx = headers.indexOf(col);
        if (colIdx !== -1) {
          sheet.getRange(i + 1, colIdx + 1).setValue(data[key]);
          updated.push(col);
        }
      }
    });

    // Shorthand: if status is being set to 'paid' and no paidAt provided, set it now
    if (data.status === 'paid') {
      const paidAtCol = headers.indexOf('PaidAt');
      if (paidAtCol !== -1 && !rows[i][paidAtCol]) {
        sheet.getRange(i + 1, paidAtCol + 1).setValue(new Date().toISOString());
        updated.push('PaidAt');
      }
    }

    return { success: true, id: data.id, updatedFields: updated };
  }
  return { error: 'Revenue ID not found: ' + data.id };
}

/**
 * Delete a Revenue row by ID.
 */
function deleteRevenue(id) {
  if (!id) return { error: 'Missing id' };
  const sheet = getOrCreateSheet(SHEET_REVENUE);
  const rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true, id: id };
    }
  }
  return { error: 'Revenue ID not found: ' + id };
}


/* ═══════════════════════════════════════
   UPDATE / DELETE
   ═══════════════════════════════════════ */

function updateStatus(sheetName, id, newStatus) {
  const sheet = getOrCreateSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const statusColIndex = headers.indexOf('Status');
  
  if (statusColIndex === -1) {
    return { error: 'No Status column found in ' + sheetName };
  }
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i + 1, statusColIndex + 1).setValue(newStatus);
      
      // If cancelling an appointment, also cancel the calendar event
      if (sheetName === SHEET_APPOINTMENTS && newStatus === 'Cancelled') {
        try {
          const eventIdCol = headers.indexOf('CalendarEventId');
          if (eventIdCol !== -1 && data[i][eventIdCol]) {
            deleteCalendarEvent(data[i][eventIdCol]);
          }
        } catch (err) {
          Logger.log('Failed to cancel calendar event: ' + err.message);
        }
      }
      
      return { success: true, id: id, status: newStatus };
    }
  }
  return { error: 'ID not found: ' + id };
}

function updateRow(sheetName, id, fields) {
  if (!fields || typeof fields !== 'object') {
    return { error: 'No fields provided' };
  }
  
  const sheet = getOrCreateSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      rowIndex = i;
      break;
    }
  }
  
  if (rowIndex === -1) {
    return { error: 'ID not found: ' + id };
  }
  
  const updatedFields = [];
  const fieldKeys = Object.keys(fields);
  
  for (let f = 0; f < fieldKeys.length; f++) {
    const fieldName = fieldKeys[f];
    const newValue = fields[fieldName];
    const colIndex = headers.indexOf(fieldName);
    if (colIndex === -1) continue;
    if (fieldName === 'ID') continue;
    
    sheet.getRange(rowIndex + 1, colIndex + 1).setValue(newValue);
    updatedFields.push(fieldName);
  }
  
  // If date/time/service changed on an appointment, update the calendar event
  if (sheetName === SHEET_APPOINTMENTS) {
    var dateChanged = updatedFields.indexOf('Date') !== -1 || updatedFields.indexOf('Time') !== -1 || updatedFields.indexOf('Service') !== -1;
    if (dateChanged) {
      try {
        var eventIdCol = headers.indexOf('CalendarEventId');
        var refreshed = sheet.getDataRange().getValues();
        var row = refreshed[rowIndex];
        var eventId = eventIdCol !== -1 ? row[eventIdCol] : '';
        if (eventId) {
          updateCalendarEvent(eventId, {
            name: row[headers.indexOf('Name')] || '',
            service: row[headers.indexOf('Service')] || '',
            date: row[headers.indexOf('Date')] || '',
            time: row[headers.indexOf('Time')] || '',
            phone: row[headers.indexOf('Phone')] || '',
            address: row[headers.indexOf('Address')] || '',
            notes: row[headers.indexOf('Notes')] || ''
          });
        }
      } catch (err) {
        Logger.log('Failed to update calendar event: ' + err.message);
      }
    }
  }
  
  return { 
    success: true, 
    id: id, 
    updatedFields: updatedFields,
    count: updatedFields.length 
  };
}

function deleteRow(sheetName, id) {
  const sheet = getOrCreateSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      // If appointment, cancel calendar event first
      if (sheetName === SHEET_APPOINTMENTS) {
        try {
          var eventIdCol = headers.indexOf('CalendarEventId');
          if (eventIdCol !== -1 && data[i][eventIdCol]) {
            deleteCalendarEvent(data[i][eventIdCol]);
          }
        } catch (err) {
          Logger.log('Failed to cancel calendar event: ' + err.message);
        }
      }
      sheet.deleteRow(i + 1);
      return { success: true, id: id };
    }
  }
  return { error: 'ID not found: ' + id };
}



/* ═══════════════════════════════════════
   UPDATE APPOINTMENT STATUS (called from booking-success.html)
   ═══════════════════════════════════════
   
   Triggered by booking-success.html after Stripe checkout completes.
   Payload: { appointmentId, status, depositPaid, stripeSession }
   
   When depositPaid === true:
     1. Flip status from "Pending Payment" → "Confirmed"
     2. Create the Google Calendar event now
     3. Write the CalendarEventId back to the sheet
     4. Send the customer their confirmation auto-reply email
   ═══════════════════════════════════════ */

function updateAppointmentStatus(data) {
  var appointmentId = data.appointmentId || '';
  var newStatus     = data.status || 'Confirmed';
  var depositPaid   = data.depositPaid === true;
  var stripeSession = data.stripeSession || '';

  if (!appointmentId) return { error: 'Missing appointmentId' };

  var sheet   = getOrCreateSheet(SHEET_APPOINTMENTS);
  var rows    = sheet.getDataRange().getValues();
  var headers = rows[0];

  var idCol        = 0; // ID is always col 0
  var statusCol    = headers.indexOf('Status');
  var calEventCol  = headers.indexOf('CalendarEventId');

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) !== String(appointmentId)) continue;

    // --- Found the row ---
    var row = rows[i];

    // Update status
    if (statusCol !== -1) {
      sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
    }

    // Store Stripe session ID in Notes if we have it (append, don't overwrite)
    if (stripeSession) {
      var notesCol = headers.indexOf('Notes');
      if (notesCol !== -1) {
        var existingNotes = row[notesCol] || '';
        var sessionNote   = '[Stripe: ' + stripeSession + ']';
        if (existingNotes.indexOf(sessionNote) === -1) {
          sheet.getRange(i + 1, notesCol + 1).setValue(
            existingNotes ? existingNotes + ' ' + sessionNote : sessionNote
          );
        }
      }
    }

    // Only create the calendar event if deposit was just paid
    var calEventId = '';
    if (depositPaid) {
      // Rebuild the appointment data object from the sheet row so we can
      // pass it to createCalendarEvent without needing the original POST body.
      var apptData = {
        name:    row[headers.indexOf('Name')]    || '',
        email:   row[headers.indexOf('Email')]   || '',
        phone:   row[headers.indexOf('Phone')]   || '',
        service: row[headers.indexOf('Service')] || '',
        date:    row[headers.indexOf('Date')]    || '',
        time:    row[headers.indexOf('Time')]    || '',
        address: row[headers.indexOf('Address')] || '',
        notes:   row[headers.indexOf('Notes')]   || '',
      };

      // Create the calendar event
      try {
        calEventId = createCalendarEvent(apptData) || '';
        if (calEventCol !== -1) {
          sheet.getRange(i + 1, calEventCol + 1).setValue(calEventId);
        }
        Logger.log('Calendar event created after deposit: ' + calEventId);
      } catch (err) {
        Logger.log('Calendar event creation failed: ' + err.message);
      }

      // Send the customer their confirmation email now that payment is done
      var emailStatus = { sent: false, reason: 'skipped' };
      try {
        emailStatus = sendAutoReply('appointment', apptData) || emailStatus;
      } catch (err) {
        emailStatus = { sent: false, reason: 'exception', error: err.message };
        Logger.log('Confirmation email failed: ' + err.message);
      }

      // Auto-log $100 deposit to Revenue sheet (idempotent — skip if already exists)
      try {
        var revSheet = getOrCreateSheet(SHEET_REVENUE);
        var revRows  = revSheet.getDataRange().getValues();
        var depExists = false;
        for (var ri = 1; ri < revRows.length; ri++) {
          if (String(revRows[ri][9]) === String(appointmentId) && String(revRows[ri][2]) === 'deposit') {
            depExists = true; break;
          }
        }
        if (!depExists) {
          addRevenue({
            type:          'deposit',
            status:        'paid',
            amount:        100,
            clientName:    apptData.name,
            clientEmail:   apptData.email,
            description:   'Booking deposit — ' + (apptData.service || 'Service Call'),
            appointmentId: appointmentId,
            stripeId:      stripeSession || '',
            paidAt:        new Date().toISOString(),
          });
        }
      } catch (err) {
        Logger.log('Revenue auto-log failed: ' + err.message);
      }

      return {
        success: true,
        id: appointmentId,
        status: newStatus,
        calendarEventId: calEventId,
        email: emailStatus,
        // Echo appointment fields so booking-success.html can render them
        appointment: {
          name:    apptData.name,
          email:   apptData.email,
          phone:   apptData.phone,
          service: apptData.service,
          date:    apptData.date,
          time:    apptData.time,
          address: apptData.address,
        }
      };
    }

    // Status-only update (no deposit flag)
    return { success: true, id: appointmentId, status: newStatus };
  }

  return { error: 'Appointment not found: ' + appointmentId };
}

/* ═══════════════════════════════════════
   GOOGLE CALENDAR INTEGRATION
   ═══════════════════════════════════════
   
   Uses CalendarApp to create/update/delete events.
   Calendar ID is stored in the Config sheet.
   Defaults to the user's primary calendar.
   
   ═══════════════════════════════════════ */

/**
 * Get the calendar ID from config, or default to primary
 */
function getCalendarId() {
  try {
    var config = getConfig();
    return config.calendarId || 'primary';
  } catch (e) {
    return 'primary';
  }
}

/**
 * Parse a date string like "Mar 15, {{COPYRIGHT_YEAR}}" and time like "2:30 PM"
 * into a JavaScript Date object
 */
function parseDateTime(dateStr, timeStr) {
  if (!dateStr) return null;
  
  // Try parsing the date string directly
  var datePart = new Date(dateStr);
  if (isNaN(datePart.getTime())) return null;
  
  if (timeStr) {
    // Parse time like "2:30 PM" or "14:30"
    var timeMatch = timeStr.match(/(\d{1,2}):(\d{2})\s*(AM|PM)?/i);
    if (timeMatch) {
      var hours = parseInt(timeMatch[1]);
      var minutes = parseInt(timeMatch[2]);
      var ampm = (timeMatch[3] || '').toUpperCase();
      
      if (ampm === 'PM' && hours < 12) hours += 12;
      if (ampm === 'AM' && hours === 12) hours = 0;
      
      datePart.setHours(hours, minutes, 0, 0);
    }
  }
  
  return datePart;
}

/**
 * Create a calendar event for a new appointment
 * Returns the event ID string
 */
function createCalendarEvent(data) {
  var calId = getCalendarId();
  var calendar = CalendarApp.getCalendarById(calId);
  if (!calendar) {
    calendar = CalendarApp.getDefaultCalendar();
  }
  if (!calendar) return '';
  
  var startTime = parseDateTime(data.date, data.time);
  if (!startTime) return '';
  
  // Default duration: 60 minutes
  var durationMin = parseInt(data.duration) || 60;
  var endTime = new Date(startTime.getTime() + durationMin * 60 * 1000);
  
  var title = '🔧 ' + (data.service || 'Service Call') + ' — ' + (data.name || 'Customer');
  
  var description = 'CUSTOMER INFO\n';
  description += '━━━━━━━━━━━━━━━━━━━━\n';
  description += 'Name: ' + (data.name || '—') + '\n';
  description += 'Phone: ' + (data.phone || '—') + '\n';
  description += 'Email: ' + (data.email || '—') + '\n';
  if (data.address) description += 'Address: ' + data.address + '\n';
  description += '\nSERVICE\n';
  description += '━━━━━━━━━━━━━━━━━━━━\n';
  description += 'Service: ' + (data.service || '—') + '\n';
  if (data.notes) description += 'Notes: ' + data.notes + '\n';
  description += '\n— Created via {{COMPANY_SHORT}} Booking';
  
  var event = calendar.createEvent(title, startTime, endTime, {
    description: description,
    location: data.address || ''
  });
  
  // Set reminder: 1 hour before and 15 min before
  event.removeAllReminders();
  event.addPopupReminder(60);
  event.addPopupReminder(15);
  
  return event.getId();
}

/**
 * Update an existing calendar event
 */
function updateCalendarEvent(eventId, data) {
  var calId = getCalendarId();
  var calendar = CalendarApp.getCalendarById(calId);
  if (!calendar) calendar = CalendarApp.getDefaultCalendar();
  if (!calendar || !eventId) return;
  
  var event = calendar.getEventById(eventId);
  if (!event) return;
  
  var startTime = parseDateTime(data.date, data.time);
  if (startTime) {
    var durationMin = parseInt(data.duration) || 60;
    var endTime = new Date(startTime.getTime() + durationMin * 60 * 1000);
    event.setTime(startTime, endTime);
  }
  
  var title = '🔧 ' + (data.service || 'Service Call') + ' — ' + (data.name || 'Customer');
  event.setTitle(title);
  
  if (data.address) event.setLocation(data.address);
  
  var description = 'CUSTOMER INFO\n';
  description += '━━━━━━━━━━━━━━━━━━━━\n';
  description += 'Name: ' + (data.name || '—') + '\n';
  description += 'Phone: ' + (data.phone || '—') + '\n';
  description += 'Email: ' + (data.email || '—') + '\n';
  if (data.address) description += 'Address: ' + data.address + '\n';
  description += '\nSERVICE\n';
  description += '━━━━━━━━━━━━━━━━━━━━\n';
  description += 'Service: ' + (data.service || '—') + '\n';
  if (data.notes) description += 'Notes: ' + data.notes + '\n';
  description += '\n— Updated via {{COMPANY_SHORT}} Admin';
  event.setDescription(description);
}

/**
 * Delete a calendar event
 */
function deleteCalendarEvent(eventId) {
  if (!eventId) return;
  var calId = getCalendarId();
  var calendar = CalendarApp.getCalendarById(calId);
  if (!calendar) calendar = CalendarApp.getDefaultCalendar();
  if (!calendar) return;
  
  try {
    var event = calendar.getEventById(eventId);
    if (event) event.deleteEvent();
  } catch (err) {
    Logger.log('Could not delete calendar event: ' + err.message);
  }
}

/**
 * Cancel a calendar event by appointment ID (looks up event ID from sheet)
 */
function cancelCalendarEvent(appointmentId) {
  if (!appointmentId) return { error: 'No appointment ID' };
  
  var sheet = getOrCreateSheet(SHEET_APPOINTMENTS);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var eventIdCol = headers.indexOf('CalendarEventId');
  
  if (eventIdCol === -1) return { error: 'No CalendarEventId column' };
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === appointmentId) {
      var eventId = data[i][eventIdCol];
      if (eventId) {
        deleteCalendarEvent(eventId);
        sheet.getRange(i + 1, eventIdCol + 1).setValue('');
        return { success: true, id: appointmentId };
      }
      return { error: 'No calendar event linked' };
    }
  }
  return { error: 'Appointment not found' };
}

/**
 * NEW: Get available time slots for a given date
 * Checks the Google Calendar for existing events and returns open 30-min slots
 * between business hours (8 AM - 5:30 PM, Mon-Sat)
 */
function getAvailability(dateStr, durationMinutes) {
  if (!dateStr) return { error: 'No date provided' };
  
  durationMinutes = durationMinutes || 60;
  
  var date = new Date(dateStr);
  if (isNaN(date.getTime())) return { error: 'Invalid date: ' + dateStr };
  
  // No Sunday appointments
  if (date.getDay() === 0) return { slots: [], message: 'Closed on Sundays' };
  
  var calId = getCalendarId();
  var calendar = CalendarApp.getCalendarById(calId);
  if (!calendar) calendar = CalendarApp.getDefaultCalendar();
  if (!calendar) return { slots: generateDefaultSlots(), source: 'default' };
  
  // Business hours: 8:00 AM to 5:30 PM
  var dayStart = new Date(date);
  dayStart.setHours(8, 0, 0, 0);
  var dayEnd = new Date(date);
  dayEnd.setHours(18, 0, 0, 0);
  
  // Get all events for the day
  var events = calendar.getEvents(dayStart, dayEnd);
  
  // Build list of busy periods
  var busy = [];
  for (var e = 0; e < events.length; e++) {
    busy.push({
      start: events[e].getStartTime().getTime(),
      end: events[e].getEndTime().getTime()
    });
  }
  
  // Generate 30-min slots from 8:00 to 17:30
  var slots = [];
  for (var h = 8; h <= 17; h++) {
    for (var m = 0; m < 60; m += 30) {
      if (h === 17 && m === 30) continue; // Last slot is 5:00 PM
      
      var slotStart = new Date(date);
      slotStart.setHours(h, m, 0, 0);
      var slotEnd = new Date(slotStart.getTime() + durationMinutes * 60 * 1000);
      
      // Check if this slot overlaps with any busy period
      var available = true;
      for (var b = 0; b < busy.length; b++) {
        if (slotStart.getTime() < busy[b].end && slotEnd.getTime() > busy[b].start) {
          available = false;
          break;
        }
      }
      
      // Don't show past time slots for today
      if (slotStart.getTime() < Date.now()) {
        available = false;
      }
      
      slots.push({
        hour: h,
        minute: m,
        available: available,
        label: formatTimeLabel(h, m)
      });
    }
  }
  
  return { slots: slots, date: dateStr, source: 'calendar' };
}

function generateDefaultSlots() {
  var slots = [];
  for (var h = 8; h <= 17; h++) {
    for (var m = 0; m < 60; m += 30) {
      if (h === 17 && m === 30) continue;
      slots.push({
        hour: h,
        minute: m,
        available: true,
        label: formatTimeLabel(h, m)
      });
    }
  }
  return slots;
}

function formatTimeLabel(h, m) {
  var ampm = h >= 12 ? 'PM' : 'AM';
  var hr = h % 12 || 12;
  return hr + ':' + (m < 10 ? '0' : '') + m + ' ' + ampm;
}


/* ═══════════════════════════════════════
   AUTO-REPLY CONFIRMATION EMAILS
   ═══════════════════════════════════════
   
   Sends a branded confirmation email after each submission.
   ONLY sends via Cloudflare Worker + Resend (never GmailApp).
   GmailApp.sendEmail always sends from the script owner's 
   personal Gmail — we never want that for customer-facing emails.
   
   Configure in Config sheet:
   - emailWorkerUrl: your Cloudflare Worker URL  (REQUIRED)
   - noreplyEmail: noreply@yourdomain.com        (REQUIRED)
   - businessName: your business name
   - businessPhone: your phone number
   
   ═══════════════════════════════════════ */

function sendAutoReply(type, data) {
  if (!data.email) return { sent: false, reason: 'no_recipient_email' };
  
  var config = getConfig();
  var bizName = config.businessName || '{{COMPANY_NAME}}';
  var bizPhone = config.businessPhone || '';
  var noreply = config.noreplyEmail || '';
  var workerUrl = config.emailWorkerUrl || '';
  
  // Must have BOTH worker URL and noreply address to send
  if (!workerUrl) {
    Logger.log('Auto-reply skipped: no emailWorkerUrl in Config sheet. Set it in Admin → Settings.');
    return { sent: false, reason: 'no_worker_url' };
  }
  if (!noreply) {
    Logger.log('Auto-reply skipped: no noreplyEmail in Config sheet. Set it in Admin → Settings.');
    return { sent: false, reason: 'no_noreply_email' };
  }
  
  var subject = getAutoReplySubject(type, bizName);
  var body = getAutoReplyBody(type, data, bizName, bizPhone);
  
  // Format: "{{COMPANY_NAME}} <noreply@{{COMPANY_SLUG}}.fit>"
  var fromFormatted = bizName + ' <' + noreply + '>';
  
  try {
    var payload = {
      from: fromFormatted,
      to: data.email,
      subject: subject,
      body: body,
      replyTo: noreply
    };
    
    Logger.log('Auto-reply sending: ' + JSON.stringify({
      workerUrl: workerUrl,
      from: fromFormatted,
      to: data.email,
      subject: subject
    }));
    
    var response = UrlFetchApp.fetch(workerUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    var code = response.getResponseCode();
    var responseText = response.getContentText();
    
    Logger.log('Auto-reply Worker response [' + code + ']: ' + responseText.substring(0, 500));
    
    if (code >= 200 && code < 300) {
      var result = {};
      try { result = JSON.parse(responseText); } catch(e) {}
      return { sent: true, code: code, workerId: result.id || '' };
    } else {
      return { sent: false, reason: 'worker_error', code: code, response: responseText.substring(0, 200) };
    }
  } catch (err) {
    Logger.log('Auto-reply fetch error: ' + err.message);
    return { sent: false, reason: 'fetch_exception', error: err.message };
  }
}

function getAutoReplySubject(type, bizName) {
  switch (type) {
    case 'appointment':
      return 'Appointment Confirmed — ' + bizName;
    case 'quote':
      return 'Quote Request Received — ' + bizName;
    case 'contact':
    case 'question':
      return 'We Received Your Message — ' + bizName;
    case 'callback':
      return 'Callback Request Received — ' + bizName;
    default:
      return 'Thank You — ' + bizName;
  }
}

function getAutoReplyBody(type, data, bizName, bizPhone) {
  var name = (data.name || 'there').split(' ')[0]; // First name
  var body = '';
  
  switch (type) {
    case 'appointment':
      body = 'Hi ' + name + ',\n\n';
      body += 'Your appointment has been confirmed! Here are the details:\n\n';
      body += '  Service: ' + (data.service || 'TBD') + '\n';
      body += '  Date: ' + (data.date || 'TBD') + '\n';
      body += '  Time: ' + (data.time || 'TBD') + '\n';
      if (data.address) body += '  Address: ' + data.address + '\n';
      body += '\nWe will call you within the hour to confirm. If you need to reschedule or cancel, please call us.\n';
      break;
      
    case 'quote':
      body = 'Hi ' + name + ',\n\n';
      body += 'Thank you for your quote request';
      if (data.service) body += ' for ' + data.service;
      body += '. We have received your details and will get back to you within 2 hours with a free estimate.\n\n';
      body += 'If your project is urgent, feel free to call us directly.\n';
      break;
      
    case 'callback':
      body = 'Hi ' + name + ',\n\n';
      body += 'We received your callback request. One of our team members will call you shortly';
      if (data.phone) body += ' at ' + data.phone;
      body += '.\n\nIf you need immediate assistance, please call us directly.\n';
      break;
      
    default:
      body = 'Hi ' + name + ',\n\n';
      body += 'Thank you for reaching out! We have received your message and will respond as soon as possible — typically within 2 hours during business hours.\n';
      break;
  }
  
  body += '\n';
  if (bizPhone) body += 'Call us: ' + bizPhone + '\n';
  body += '\nBest regards,\nThe ' + bizName + ' Team\n';
  body += '\n---\nThis is an automated confirmation. Please do not reply to this email.';
  
  return body;
}


/* ═══════════════════════════════════════
   READ OPERATIONS
   ═══════════════════════════════════════ */

function getSheetData(sheetName) {
  const sheet = getOrCreateSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = {};
    for (let j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    rows.push(row);
  }
  return rows;
}

function getDashboardStats() {
  const appointments = getSheetData(SHEET_APPOINTMENTS);
  const quotes = getSheetData(SHEET_QUOTES);
  const contacts = getSheetData(SHEET_CONTACTS);
  const clients = getSheetData(SHEET_CLIENTS);
  
  const now = new Date();
  const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  
  const thisWeekAppts = appointments.filter(a => new Date(a.Timestamp) >= weekAgo).length;
  const thisWeekQuotes = quotes.filter(q => new Date(q.Timestamp) >= weekAgo).length;
  
  const statusCounts = {};
  appointments.forEach(a => {
    const s = a.Status || 'New';
    statusCounts[s] = (statusCounts[s] || 0) + 1;
  });
  
  return {
    totalAppointments: appointments.length,
    totalQuotes: quotes.length,
    totalContacts: contacts.length,
    totalClients: clients.length,
    thisWeekAppointments: thisWeekAppts,
    thisWeekQuotes: thisWeekQuotes,
    pendingQuotes: quotes.filter(q => q.Status === 'Pending').length,
    unreadMessages: contacts.filter(c => c.Status === 'Unread').length,
    appointmentsByStatus: statusCounts,
  };
}


/* ═══════════════════════════════════════
   CONFIG MANAGEMENT
   ═══════════════════════════════════════ */

function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_CONFIG);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CONFIG);
    sheet.appendRow(['Key', 'Value']);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    sheet.setFrozenRows(1);
    
    // Default config — includes new v3 fields
    const defaults = [
      ['businessName', '{{COMPANY_NAME}}'],
      ['businessPhone', '{{BUSINESS_PHONE}}'],
      ['businessEmail', 'info@{{COMPANY_SLUG}}.com'],
      ['services', 'Emergency Repair,Drain Cleaning,Water Heater,Pipe Installation,Fixture Install,Inspection,Sewer Line,Gas Line'],
      ['businessHours', 'Mon-Sat: 7am-8pm'],
      ['serviceArea', '{{SERVICE_AREA}}'],
      ['calendarId', 'primary'],
      ['emailWorkerUrl', ''],
      ['noreplyEmail', ''],
    ];
    defaults.forEach(d => sheet.appendRow(d));
  }
  
  const data = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) {
    config[data[i][0]] = data[i][1];
  }
  return config;
}

function saveConfig(configObj) {
  if (!configObj || typeof configObj !== 'object') {
    return { error: 'Invalid config data' };
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_CONFIG);
  if (!sheet) {
    getConfig();
    sheet = ss.getSheetByName(SHEET_CONFIG);
  }
  
  const data = sheet.getDataRange().getValues();
  const keys = Object.keys(configObj);
  
  keys.forEach(key => {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(configObj[key]);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, configObj[key]]);
    }
  });
  
  return { success: true, keys: keys };
}


/* ═══════════════════════════════════════
   EMAIL — Read Gmail + Log Sends
   ═══════════════════════════════════════ */

function getEmails(view, count) {
  try {
    count = Math.min(count || 30, 50);
    var query = view === 'sent' ? 'in:sent' : 'in:inbox';
    var threads = GmailApp.search(query, 0, count);
    var results = [];
    
    for (var t = 0; t < threads.length; t++) {
      var thread = threads[t];
      var msgs = thread.getMessages();
      var lastMsg = msgs[msgs.length - 1];
      
      results.push({
        id: thread.getId(),
        subject: thread.getFirstMessageSubject() || '(no subject)',
        from: lastMsg.getFrom(),
        to: lastMsg.getTo(),
        date: lastMsg.getDate().toISOString(),
        snippet: lastMsg.getPlainBody().substring(0, 200),
        body: lastMsg.getPlainBody().substring(0, 2000),
        unread: thread.isUnread(),
        msgCount: msgs.length
      });
    }
    
    return results;
  } catch (err) {
    return { error: 'Gmail access error: ' + err.message };
  }
}

function sendEmailAction(data) {
  try {
    var to = data.to;
    var subject = data.subject || '(no subject)';
    var body = data.body || '';
    var from = data.from || '';
    
    if (!to) return { error: 'No recipient specified' };
    
    logSentEmail(to, from, subject, body);
    
    return { success: true, to: to, subject: subject, logged: true };
  } catch (err) {
    return { error: 'Log failed: ' + err.message };
  }
}

function logSentEmail(to, from, subject, body) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('SentEmails');
    if (!sheet) {
      sheet = ss.insertSheet('SentEmails');
      sheet.appendRow(['Timestamp', 'From', 'To', 'Subject', 'Body']);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([
      new Date().toISOString(),
      from || 'default',
      to,
      subject,
      body.substring(0, 500)
    ]);
  } catch (e) {
    Logger.log('Failed to log sent email: ' + e.message);
  }
}




/* ═══════════════════════════════════════
   REVIEWS
   ═══════════════════════════════════════
   Sheet: Reviews
   Columns: ID, Timestamp, Name, Rating, ReviewText, Status
   Status: 'Published' | 'Hidden' | 'Pending'
   ═══════════════════════════════════════ */

function addReview(data) {
  var sheet = getOrCreateSheet(SHEET_REVIEWS);
  var id = generateId();
  sheet.appendRow([
    id,
    data.timestamp || new Date().toISOString(),
    data.name || 'Anonymous',
    Math.min(5, Math.max(0, parseInt(data.rating) || 5)),
    data.reviewText || data.text || '',
    data.status || 'Pending'
  ]);
  return { success: true, id: id, type: 'review' };
}

function getReviews() {
  var sheet = getOrCreateSheet(SHEET_REVIEWS);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    rows.push(row);
  }
  return rows;
}

function updateReviewStatus(data) {
  if (!data.id) return { error: 'Missing id' };
  var sheet = getOrCreateSheet(SHEET_REVIEWS);
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var statusCol = headers.indexOf('Status');
  var ratingCol = headers.indexOf('Rating');
  var nameCol   = headers.indexOf('Name');
  var textCol   = headers.indexOf('ReviewText');

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(data.id)) continue;
    if (data.status !== undefined && statusCol !== -1)
      sheet.getRange(i + 1, statusCol + 1).setValue(data.status);
    if (data.rating !== undefined && ratingCol !== -1)
      sheet.getRange(i + 1, ratingCol + 1).setValue(Math.min(5, Math.max(0, parseInt(data.rating) || 0)));
    if (data.name !== undefined && nameCol !== -1)
      sheet.getRange(i + 1, nameCol + 1).setValue(data.name);
    if (data.reviewText !== undefined && textCol !== -1)
      sheet.getRange(i + 1, textCol + 1).setValue(data.reviewText);
    return { success: true, id: data.id };
  }
  return { error: 'Review not found: ' + data.id };
}

function deleteReview(id) {
  if (!id) return { error: 'Missing id' };
  var sheet = getOrCreateSheet(SHEET_REVIEWS);
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true, id: id };
    }
  }
  return { error: 'Review not found: ' + id };
}


/* ═══════════════════════════════════════
   SERVICES (editable from admin settings)
   ═══════════════════════════════════════
   Sheet: Services
   Columns: ID, Order, Name, Icon, Description, Price, Duration, Active
   Active: TRUE | FALSE
   ═══════════════════════════════════════ */

function getServices() {
  var sheet = getOrCreateSheet(SHEET_SERVICES);
  var data = sheet.getDataRange().getValues();
  // Seed defaults if sheet is brand new (only header row)
  if (data.length <= 1) {
    _seedDefaultServices(sheet);
    data = sheet.getDataRange().getValues();
  }
  if (data.length <= 1) return [];
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    // Normalise Active to boolean
    row.Active = (row.Active === true || row.Active === 'TRUE' || row.Active === 'true');
    rows.push(row);
  }
  // Sort by Order ascending
  rows.sort(function(a, b) { return (parseInt(a.Order) || 0) - (parseInt(b.Order) || 0); });
  return rows;
}

function _seedDefaultServices(sheet) {
  var defaults = [
    [generateId(), 1, 'Emergency Repair',  '🔧', '24/7 emergency {{INDUSTRY}} repairs for burst pipes, major leaks, and flooding.',        'From $89',  60,  true],
    [generateId(), 2, 'Drain Cleaning',    '🚿', 'Professional drain unclogging for kitchens, bathrooms, and main lines.',             'From $65',  45,  true],
    [generateId(), 3, 'Water Heater',      '🔥', 'Installation, repair, and maintenance of all water heater types.',                   'From $120', 90,  true],
    [generateId(), 4, 'Pipe Installation', '🔩', 'New pipe installation, repiping, and pipe replacement services.',                    'From $150', 120, true],
    [generateId(), 5, 'Fixture Install',   '🚰', 'Faucet, toilet, shower, and sink installation and replacement.',                    'From $75',  60,  true],
    [generateId(), 6, 'Inspection',        '🔍', 'Comprehensive {{INDUSTRY}} inspection with camera and detailed report.',                 'From $49',  45,  true],
    [generateId(), 7, 'Sewer Line',        '🏗️', 'Sewer line repair, replacement, and trenchless solutions.',                        'From $200', 180, true],
    [generateId(), 8, 'Gas Line',          '⚡', 'Gas line installation, repair, and leak detection services.',                       'From $175', 90,  true],
  ];
  defaults.forEach(function(row) { sheet.appendRow(row); });
}

function addService(data) {
  var sheet = getOrCreateSheet(SHEET_SERVICES);
  var rows = sheet.getDataRange().getValues();
  // Next order number = max existing + 1
  var maxOrder = 0;
  for (var i = 1; i < rows.length; i++) {
    var ord = parseInt(rows[i][1]) || 0;
    if (ord > maxOrder) maxOrder = ord;
  }
  var id = generateId();
  sheet.appendRow([
    id,
    maxOrder + 1,
    data.name || 'New Service',
    data.icon || '🔧',
    data.description || '',
    data.price || 'Call for pricing',
    parseInt(data.duration) || 60,
    data.active !== false  // default true
  ]);
  return { success: true, id: id };
}

function updateService(data) {
  if (!data.id) return { error: 'Missing id' };
  var sheet = getOrCreateSheet(SHEET_SERVICES);
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(data.id)) continue;
    var fields = ['Order', 'Name', 'Icon', 'Description', 'Price', 'Duration', 'Active'];
    fields.forEach(function(col) {
      var key = col.charAt(0).toLowerCase() + col.slice(1);
      if (data[key] !== undefined) {
        var colIdx = headers.indexOf(col);
        if (colIdx !== -1) sheet.getRange(i + 1, colIdx + 1).setValue(data[key]);
      }
    });
    return { success: true, id: data.id };
  }
  return { error: 'Service not found: ' + data.id };
}

function deleteService(id) {
  if (!id) return { error: 'Missing id' };
  var sheet = getOrCreateSheet(SHEET_SERVICES);
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true, id: id };
    }
  }
  return { error: 'Service not found: ' + id };
}

// Accepts array of IDs in desired order, re-numbers the Order column
function reorderServices(orderArray) {
  if (!Array.isArray(orderArray)) return { error: 'order must be an array of IDs' };
  var sheet = getOrCreateSheet(SHEET_SERVICES);
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var orderCol = headers.indexOf('Order');
  if (orderCol === -1) return { error: 'No Order column' };

  orderArray.forEach(function(id, idx) {
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        sheet.getRange(i + 1, orderCol + 1).setValue(idx + 1);
        break;
      }
    }
  });
  return { success: true, count: orderArray.length };
}


/* ═══════════════════════════════════════
   ABOUT PAGE CONTENT
   ═══════════════════════════════════════
   Sheet: About
   Columns: Key, Value  (same pattern as Config)

   Default keys:
     headline        — main tagline
     subheadline     — supporting line under headline
     story           — company origin / mission paragraph
     mission         — mission statement
     yearsInBusiness
     jobsCompleted
     satisfactionRate
     teamSize
     teamMember1Name / teamMember1Role / teamMember1Years / teamMember1Cert
     teamMember2Name / teamMember2Role / teamMember2Years / teamMember2Cert
     teamMember3Name / teamMember3Role / teamMember3Years / teamMember3Cert
     guarantee1Title / guarantee1Desc
     guarantee2Title / guarantee2Desc
     guarantee3Title / guarantee3Desc
     guarantee4Title / guarantee4Desc
   ═══════════════════════════════════════ */

function getAbout() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_ABOUT);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ABOUT);
    sheet.appendRow(['Key', 'Value']);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    sheet.setFrozenRows(1);
    _seedDefaultAbout(sheet);
  }

  var data = sheet.getDataRange().getValues();
  var about = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) about[data[i][0]] = data[i][1];
  }
  return about;
}

function _seedDefaultAbout(sheet) {
  var defaults = [
    ['headline',         '{{ABOUT_HEADLINE}}'],
    ['subheadline',      'Family-owned and serving {{SERVICE_AREA}} since {{YEAR_FOUNDED}}'],
    ['story',            '{{COMPANY_SHORT}} started as a one-truck operation out of Fort Lauderdale. Over 18 years we have grown into a team of licensed master plumbers who treat every home like their own. We believe in honest pricing, showing up on time, and never leaving a job half done.'],
    ['mission',          'To deliver reliable, fairly priced {{INDUSTRY}} services with the transparency and respect every homeowner deserves.'],
    ['yearsInBusiness',  '18+'],
    ['jobsCompleted',    '4,500+'],
    ['satisfactionRate', '99%'],
    ['teamSize',         '12'],
    ['teamMember1Name',  'Mike Henderson'],
    ['teamMember1Role',  'Master Plumber / Owner'],
    ['teamMember1Years', '18'],
    ['teamMember1Cert',  'Licensed & Insured'],
    ['teamMember2Name',  'Sarah Chen'],
    ['teamMember2Role',  'Lead Technician'],
    ['teamMember2Years', '12'],
    ['teamMember2Cert',  'Certified Pipefitter'],
    ['teamMember3Name',  'Tony Reeves'],
    ['teamMember3Role',  'Senior Plumber'],
    ['teamMember3Years', '9'],
    ['teamMember3Cert',  'Backflow Certified'],
    ['guarantee1Title',  'Licensed & Bonded'],
    ['guarantee1Desc',   'Full state licensing with comprehensive bonding coverage.'],
    ['guarantee2Title',  'Satisfaction Guarantee'],
    ['guarantee2Desc',   "Not happy? We'll make it right or your money back."],
    ['guarantee3Title',  '24/7 Emergency'],
    ['guarantee3Desc',   'Round-the-clock emergency response, 365 days a year.'],
    ['guarantee4Title',  'Upfront Pricing'],
    ['guarantee4Desc',   'No hidden fees. You approve the price before we start.'],
  ];
  defaults.forEach(function(row) { sheet.appendRow(row); });
}

function saveAbout(fields) {
  if (!fields || typeof fields !== 'object') return { error: 'Invalid fields' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_ABOUT);
  if (!sheet) {
    getAbout(); // creates + seeds the sheet
    sheet = ss.getSheetByName(SHEET_ABOUT);
  }

  var data = sheet.getDataRange().getValues();
  var keys = Object.keys(fields);

  keys.forEach(function(key) {
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(fields[key]);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, fields[key]]);
    }
  });

  return { success: true, keys: keys };
}

/* ═══════════════════════════════════════
   HELPERS
   ═══════════════════════════════════════ */

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (HEADERS[name]) {
      sheet.appendRow(HEADERS[name]);
      sheet.getRange(1, 1, 1, HEADERS[name].length).setFontWeight('bold');
      sheet.setFrozenRows(1);
      for (let i = 1; i <= HEADERS[name].length; i++) {
        sheet.setColumnWidth(i, 140);
      }
    }
  }
  
  return sheet;
}

function generateId() {
  return Utilities.getUuid().split('-')[0].toUpperCase();
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

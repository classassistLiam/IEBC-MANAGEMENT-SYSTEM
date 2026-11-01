
/**
 * HOTEL MANAGEMENT API
 * Google Apps Script + Google Sheets
 *
 * Actions:
 * GET  ?action=health
 * GET  ?action=listGuests
 * POST { action:"saveGuest", guest:{...} }
 * POST { action:"saveReservation", reservation:{...} }
 */

const SPREADSHEET_ID = "1WwJCWC_1x4g7ZerqBCX3dXbpbYcwyr0GOMa5BwMwH-E";

const RESERVATIONS_SHEET = "Reservations";
const GUESTS_SHEET = "Guests";
const ELECTION_RESULTS_SHEET = "ElectionResults";
const VOTER_REGISTRATIONS_SHEET = "VoterRegistrations";
const OPINION_POLLS_SHEET = "OpinionPolls";

const RESERVATION_HEADERS = [
  "Transaction ID", "Created At", "Updated At", "Device ID",
  "Guest Name", "Guest JSON", "Room Number", "Room Type",
  "Room JSON", "Discount JSON", "Check-in", "Check-out",
  "Days", "Room Rate", "Advance", "Subtotal", "Total",
  "Sync Status", "Server Received At"
];

const GUEST_HEADERS = [
  "Guest ID", "First Name", "Middle Name", "Last Name",
  "Address", "Contact", "Email", "Gender", "Status",
  "Registration Date", "Server Received At", "Guest JSON"
];

const ELECTION_HEADERS = [
  "Record Key", "Election ID", "Location ID", "Location Name",
  "Registered Voters", "Votes Cast", "Valid Votes", "Rejected Votes",
  "Candidates JSON", "Created At", "Updated At", "Updated By", "Change Log JSON",
  "GPS Latitude", "GPS Longitude", "GPS Accuracy Metres", "GPS Captured At"
];

const VOTER_REGISTRATION_HEADERS = [
  "Registration ID", "Name", "Sub-location ID", "Location ID", "Division ID",
  "District ID", "Created At", "Consent Confirmed"
];

const OPINION_POLL_HEADERS = [
  "Response ID", "Poll Station", "Presidential Candidate", "Governor Candidate",
  "Senator Candidate", "MP Candidate", "Woman Representative Candidate",
  "MCA Candidate", "Reason", "Consent Confirmed", "Submitted At"
];

/* =========================================
   HTTP ROUTES
========================================= */

function doGet(e) {
  try {
    const action = e && e.parameter ? e.parameter.action : "";

    if (action === "health") {
      return jsonOutput({
        success: true,
        service: "Hotel Management API",
        status: "ready",
        timestamp: new Date().toISOString()
      });
    }

    if (action === "listGuests") {
      return jsonOutput({
        success: true,
        guests: listGuests()
      });
    }

    if (action === "getElectionResult" || action === "get") {
      return jsonOutput(getElectionResult(
        e.parameter.electionId || "",
        e.parameter.locationId || ""
      ));
    }

    if (action === "checkElectionDuplicate" || action === "checkDuplicate") {
      return jsonOutput(checkElectionDuplicate(
        e.parameter.electionId || "",
        e.parameter.locationId || ""
      ));
    }

    if (action === "listElectionResults") {
      return jsonOutput(listElectionResults(e.parameter.electionId || ""));
    }

    if (action === "getNationwideOpinionPoll") {
      return jsonOutput(getNationwideOpinionPoll());
    }

    return jsonOutput({
      success: true,
      service: "Hotel Management API",
      supportedActions: ["health", "listGuests", "getElectionResult", "checkElectionDuplicate", "listElectionResults", "getNationwideOpinionPoll"],
      message: "Use a supported action."
    });

  } catch (error) {
    console.error(error);
    return jsonOutput({
      success: false,
      message: String(error.message || error)
    });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(
      e && e.postData && e.postData.contents
        ? e.postData.contents
        : "{}"
    );

    let result;

    if (body.action === "saveGuest") {
      result = saveGuest(body.guest || {});
    } else if (body.action === "saveReservation") {
      result = saveReservation(body.reservation || {});
    } else if (body.action === "saveElectionResult" || body.action === "save") {
      result = saveElectionResult(body.record || body);
    } else if (body.action === "overwriteElectionResult" || body.action === "overwrite") {
      result = overwriteElectionResult(body.record || body, body.password || "");
    } else if (body.action === "registerVoter") {
      result = registerVoter(body.data || body.registration || {});
    } else if (body.action === "submitOpinionPoll") {
      result = submitOpinionPoll(body.data || body.poll || {});
    } else {
      throw new Error("Unsupported POST action.");
    }

    return jsonOutput(result);

  } catch (error) {
    console.error(error);

    return jsonOutput({
      success: false,
      message: String(error.message || error)
    });
  }
}

/* =========================================
   SPREADSHEET HELPERS
========================================= */

function getSpreadsheet() {
  if (!SPREADSHEET_ID) {
    throw new Error("SPREADSHEET_ID is not configured.");
  }

  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getOrCreateSheet(name, headers) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold");
  } else {
    // Do not silently overwrite an existing sheet's headers.
    const existingHeaders = sheet
      .getRange(1, 1, 1, headers.length)
      .getDisplayValues()[0];

    if (existingHeaders[0] !== headers[0]) {
      throw new Error(
        "Unexpected headers in " + name +
        ". Check the sheet structure before writing."
      );
    }
  }

  return sheet;
}

function getGuestSheet() {
  return getOrCreateSheet(GUESTS_SHEET, GUEST_HEADERS);
}

function getReservationSheet() {
  return getOrCreateSheet(
    RESERVATIONS_SHEET,
    RESERVATION_HEADERS
  );
}

function safeCell(value) {
  const text = String(value == null ? "" : value);

  // Prevent user-supplied text from being interpreted as a formula.
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function jsonOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/* =========================================
   GUESTS: READ CENTRAL DATABASE
========================================= */

function listGuests() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sheet = getGuestSheet();
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) return [];

    const values = sheet
      .getRange(2, 1, lastRow - 1, GUEST_HEADERS.length)
      .getDisplayValues();

    return values
      .filter(row => row[0])
      .map(row => {
        let savedGuest = {};

        try {
          savedGuest = row[11] ? JSON.parse(row[11]) : {};
        } catch (error) {
          savedGuest = {};
        }

        return {
          ...savedGuest,
          guestId: row[0],
          first: row[1],
          middle: row[2],
          last: row[3],
          address: row[4],
          contact: row[5],
          email: row[6],
          gender: row[7],
          status: row[8],
          date: row[9]
        };
      });

  } finally {
    lock.releaseLock();
  }
}

/* =========================================
   GUESTS: SAVE + DEDUPLICATE
========================================= */

function saveGuest(g) {
  validateGuest(g);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sheet = getGuestSheet();
    const lastRow = sheet.getLastRow();

    const guestId = String(g.guestId).trim();

    if (lastRow >= 2) {
      const existingIds = sheet
        .getRange(2, 1, lastRow - 1, 1)
        .getDisplayValues();

      const foundIndex = existingIds.findIndex(
        row => String(row[0]).trim() === guestId
      );

      if (foundIndex !== -1) {
        return {
          success: true,
          duplicate: true,
          message: "Guest already exists; no duplicate row inserted.",
          guestId: guestId,
          row: foundIndex + 2
        };
      }
    }

    const now = new Date().toISOString();

    const row = [
      safeCell(guestId),
      safeCell(g.first),
      safeCell(g.middle || ""),
      safeCell(g.last),
      safeCell(g.address || ""),
      safeCell(g.contact),
      safeCell(g.email || ""),
      safeCell(g.gender || ""),
      safeCell(g.status || "Active"),
      safeCell(g.date || now),
      now,
      JSON.stringify(g)
    ];

    sheet.appendRow(row);
    SpreadsheetApp.flush();

    return {
      success: true,
      duplicate: false,
      message: "Guest saved to Google Sheets.",
      guestId: guestId,
      row: sheet.getLastRow()
    };

  } finally {
    lock.releaseLock();
  }
}

function validateGuest(g) {
  if (!g || typeof g !== "object") {
    throw new Error("Guest data is missing.");
  }

  if (!g.guestId || String(g.guestId).trim().length < 6) {
    throw new Error("A stable guestId is required.");
  }

  if (!String(g.first || "").trim()) {
    throw new Error("First name is required.");
  }

  if (!String(g.last || "").trim()) {
    throw new Error("Last name is required.");
  }

  if (!String(g.contact || "").trim()) {
    throw new Error("Contact number is required.");
  }
}

/* =========================================
   RESERVATIONS: SAVE + DEDUPLICATE
========================================= */

function saveReservation(r) {
  validateReservation(r);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sheet = getReservationSheet();
    const lastRow = sheet.getLastRow();

    if (lastRow >= 2) {
      const existingIds = sheet
        .getRange(2, 1, lastRow - 1, 1)
        .getDisplayValues();

      const foundIndex = existingIds.findIndex(
        row => String(row[0]).trim() === String(r.transID).trim()
      );

      if (foundIndex !== -1) {
        return {
          success: true,
          duplicate: true,
          message: "Transaction already exists; no row inserted.",
          transID: r.transID,
          row: foundIndex + 2
        };
      }
    }

    const now = new Date().toISOString();

    const row = [
      safeCell(r.transID),
      safeCell(r.createdAt || now),
      now,
      safeCell(r.deviceId || ""),
      safeCell(r.guestName || ""),
      JSON.stringify(r.guest || {}),
      safeCell(r.roomNumber || ""),
      safeCell(r.roomType || ""),
      JSON.stringify(r.room || {}),
      JSON.stringify(r.discount || null),
      safeCell(r.checkIn),
      safeCell(r.checkOut),
      Number(r.days),
      Number(r.roomRate),
      Number(r.advance),
      Number(r.subtotal),
      Number(r.total),
      "SYNCED",
      now
    ];

    sheet.appendRow(row);
    SpreadsheetApp.flush();

    return {
      success: true,
      duplicate: false,
      message: "Reservation saved.",
      transID: r.transID,
      row: sheet.getLastRow()
    };

  } finally {
    lock.releaseLock();
  }
}

function validateReservation(r) {
  if (!r || typeof r !== "object") {
    throw new Error("Reservation data is missing.");
  }

  if (!r.transID || String(r.transID).trim().length < 6) {
    throw new Error("A valid transaction ID is required.");
  }

  if (!r.guest || !r.room) {
    throw new Error("Guest and room information are required.");
  }

  if (
    !/^\d{4}-\d{2}-\d{2}$/.test(String(r.checkIn || "")) ||
    !/^\d{4}-\d{2}-\d{2}$/.test(String(r.checkOut || "")) ||
    r.checkOut <= r.checkIn
  ) {
    throw new Error("Invalid check-in or check-out dates.");
  }

  const days = Number(r.days);
  const rate = Number(r.roomRate);
  const advance = Number(r.advance);
  const subtotal = Number(r.subtotal);
  const total = Number(r.total);

  if (!Number.isInteger(days) || days < 1) {
    throw new Error("Reservation days must be at least one.");
  }

  if (
    !Number.isFinite(rate) || rate < 0 ||
    !Number.isFinite(advance) || advance < 0 ||
    !Number.isFinite(subtotal) || subtotal < 0 ||
    !Number.isFinite(total) || total < 0
  ) {
    throw new Error("Invalid reservation amounts.");
  }

  if (advance > subtotal + 0.01) {
    throw new Error("Advance cannot exceed subtotal.");
  }

  if (Math.abs(total - (subtotal - advance)) > 0.02) {
    throw new Error("Reservation total does not match amount due.");
  }
}


/* =========================================
   ELECTION RESULTS: READ, DEDUPLICATE, SAVE
   Unique key = Election ID + Location ID
========================================= */

function getElectionSheet() {
  const sheet = getOrCreateSheet(ELECTION_RESULTS_SHEET, ELECTION_HEADERS);
  // Migrate older ElectionResults sheets by adding the four GPS columns safely.
  const existing = sheet.getRange(1, 1, 1, ELECTION_HEADERS.length).getDisplayValues()[0];
  if (existing[0] !== ELECTION_HEADERS[0]) {
    throw new Error("Unexpected headers in " + ELECTION_RESULTS_SHEET + ". Check the sheet structure before writing.");
  }
  for (let i = 0; i < ELECTION_HEADERS.length; i++) {
    if (!existing[i]) sheet.getRange(1, i + 1).setValue(ELECTION_HEADERS[i]);
  }
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, ELECTION_HEADERS.length).setFontWeight("bold");
  return sheet;
}

function electionRecordKey(electionId, locationId) {
  const election = String(electionId || "").trim();
  const location = String(locationId || "").trim();

  if (!election || !location) {
    throw new Error("Election ID and unique Location ID are required.");
  }

  return election.toLowerCase() + "::" + location.toLowerCase();
}

function findElectionRecordRow(sheet, key) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const keys = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  const index = keys.findIndex(row => String(row[0]).trim() === key);
  return index === -1 ? -1 : index + 2;
}

function getElectionResult(electionId, locationId) {
  const key = electionRecordKey(electionId, locationId);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sheet = getElectionSheet();
    const rowNumber = findElectionRecordRow(sheet, key);

    if (rowNumber === -1) {
      return { success: true, found: false, record: null };
    }

    const row = sheet.getRange(rowNumber, 1, 1, ELECTION_HEADERS.length).getValues()[0];
    let candidates = {};
    let changeLog = [];

    try { candidates = row[8] ? JSON.parse(row[8]) : {}; } catch (err) {}
    try { changeLog = row[12] ? JSON.parse(row[12]) : []; } catch (err) {}

    return {
      success: true,
      found: true,
      row: rowNumber,
      record: {
        electionId: row[1],
        locationId: row[2],
        locationName: row[3],
        registered: Number(row[4]) || 0,
        cast: Number(row[5]) || 0,
        valid: Number(row[6]) || 0,
        rejected: Number(row[7]) || 0,
        candidates: candidates,
        createdAt: row[9],
        updatedAt: row[10],
        updatedBy: row[11],
        changeLog: changeLog,
        gpsLatitude: row[13] || "",
        gpsLongitude: row[14] || "",
        gpsAccuracyMetres: row[15] || "",
        gpsCapturedAt: row[16] || ""
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function checkElectionDuplicate(electionId, locationId) {
  const key = electionRecordKey(electionId, locationId);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sheet = getElectionSheet();
    const rowNumber = findElectionRecordRow(sheet, key);

    return {
      success: true,
      duplicate: rowNumber !== -1,
      found: rowNumber !== -1,
      row: rowNumber === -1 ? null : rowNumber,
      electionId: String(electionId).trim(),
      locationId: String(locationId).trim()
    };
  } finally {
    lock.releaseLock();
  }
}

function validateElectionRecord(record) {
  if (!record || typeof record !== "object") {
    throw new Error("Election result data is missing.");
  }

  electionRecordKey(record.electionId, record.locationId);

  ["registered", "cast", "valid", "rejected"].forEach(function(field) {
    const value = Number(record[field]);
    if (!Number.isInteger(value) || value < 0) {
      throw new Error(field + " must be a non-negative whole number.");
    }
  });

  if (Number(record.cast) > Number(record.registered)) {
    throw new Error("Votes cast cannot exceed registered voters.");
  }

  if (Number(record.valid) + Number(record.rejected) !== Number(record.cast)) {
    throw new Error("Valid votes plus rejected votes must equal votes cast.");
  }

  if (!record.candidates || typeof record.candidates !== "object" ||
      Array.isArray(record.candidates)) {
    throw new Error("Candidate vote totals are required.");
  }

  let candidateTotal = 0;
  Object.keys(record.candidates).forEach(function(name) {
    const votes = Number(record.candidates[name]);
    if (!Number.isInteger(votes) || votes < 0) {
      throw new Error("Candidate votes must be non-negative whole numbers.");
    }
    candidateTotal += votes;
  });

  if (candidateTotal !== Number(record.valid)) {
    throw new Error("Candidate votes must equal valid votes.");
  }
}

function electionRowValues(record, key, createdAt, updatedAt, updatedBy, changeLog) {
  return [
    safeCell(key),
    safeCell(String(record.electionId).trim()),
    safeCell(String(record.locationId).trim()),
    safeCell(record.locationName || ""),
    Number(record.registered),
    Number(record.cast),
    Number(record.valid),
    Number(record.rejected),
    JSON.stringify(record.candidates || {}),
    createdAt,
    updatedAt,
    safeCell(updatedBy || "unknown"),
    JSON.stringify(changeLog || []),
    record.gpsLatitude === "" || record.gpsLatitude == null ? "" : Number(record.gpsLatitude),
    record.gpsLongitude === "" || record.gpsLongitude == null ? "" : Number(record.gpsLongitude),
    record.gpsAccuracyMetres === "" || record.gpsAccuracyMetres == null ? "" : Number(record.gpsAccuracyMetres),
    safeCell(record.gpsCapturedAt || "")
  ];
}

function saveElectionResult(record) {
  validateElectionRecord(record);

  const key = electionRecordKey(record.electionId, record.locationId);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sheet = getElectionSheet();
    const existingRow = findElectionRecordRow(sheet, key);

    if (existingRow !== -1) {
      return {
        success: false,
        duplicate: true,
        message: "A record already exists for this election and location. Nothing was overwritten.",
        row: existingRow
      };
    }

    const now = new Date().toISOString();
    const log = [{
      action: "CREATE",
      at: now,
      updatedBy: String(record.updatedBy || record.deviceId || "unknown"),
      snapshot: {
        registered: Number(record.registered),
        cast: Number(record.cast),
        valid: Number(record.valid),
        rejected: Number(record.rejected),
        candidates: record.candidates
      }
    }];

    sheet.appendRow(electionRowValues(
      record, key, now, now,
      record.updatedBy || record.deviceId || "unknown", log
    ));
    SpreadsheetApp.flush();

    return {
      success: true,
      duplicate: false,
      message: "Election results saved successfully.",
      row: sheet.getLastRow()
    };
  } finally {
    lock.releaseLock();
  }
}

function overwriteElectionResult(record, suppliedPassword) {
  validateElectionRecord(record);

  // Configure this secret in Apps Script > Project Settings > Script Properties.
  // Property name: ELECTION_OVERWRITE_PASSWORD
  const expectedPassword = PropertiesService.getScriptProperties()
    .getProperty("ELECTION_OVERWRITE_PASSWORD");

  if (!expectedPassword) {
    throw new Error("Overwrite is disabled: ELECTION_OVERWRITE_PASSWORD is not configured in Script Properties.");
  }

  if (!suppliedPassword ||
      !constantTimeStringMatch(String(suppliedPassword), String(expectedPassword))) {
    throw new Error("Overwrite authorization failed. The password is incorrect.");
  }

  const key = electionRecordKey(record.electionId, record.locationId);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sheet = getElectionSheet();
    const rowNumber = findElectionRecordRow(sheet, key);

    if (rowNumber === -1) {
      throw new Error("No existing record was found to overwrite. Use Save New Results instead.");
    }

    const oldRow = sheet.getRange(rowNumber, 1, 1, ELECTION_HEADERS.length).getValues()[0];
    let changeLog = [];

    try { changeLog = oldRow[12] ? JSON.parse(oldRow[12]) : []; } catch (err) {}

    // Keep a copy of the prior values in the audit log before replacing them.
    changeLog.push({
      action: "OVERWRITE",
      at: new Date().toISOString(),
      updatedBy: String(record.updatedBy || record.deviceId || "unknown"),
      previous: {
        registered: Number(oldRow[4]) || 0,
        cast: Number(oldRow[5]) || 0,
        valid: Number(oldRow[6]) || 0,
        rejected: Number(oldRow[7]) || 0,
        candidates: (function() {
          try { return oldRow[8] ? JSON.parse(oldRow[8]) : {}; } catch (err) { return {}; }
        })()
      }
    });

    const createdAt = oldRow[9] || new Date().toISOString();
    const updatedAt = new Date().toISOString();

    sheet.getRange(rowNumber, 1, 1, ELECTION_HEADERS.length).setValues([electionRowValues(
      record, key, createdAt, updatedAt,
      record.updatedBy || record.deviceId || "unknown", changeLog
    )]);

    SpreadsheetApp.flush();

    return {
      success: true,
      overwritten: true,
      message: "Existing election results overwritten. Previous values were retained in the change log.",
      row: rowNumber
    };
  } finally {
    lock.releaseLock();
  }
}

function constantTimeStringMatch(a, b) {
  // Avoid returning early based on the first mismatched character.
  let mismatch = a.length ^ b.length;
  const maxLength = Math.max(a.length, b.length);

  for (let i = 0; i < maxLength; i++) {
    mismatch |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  }

  return mismatch === 0;
}


/* =========================================
   ELECTION RESULTS: NATIONWIDE LIST
========================================= */
function listElectionResults(electionId) {
  const sheet = getElectionSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, records: [] };
  const values = sheet.getRange(2, 1, lastRow - 1, ELECTION_HEADERS.length).getValues();
  const wanted = String(electionId || "").trim().toLowerCase();
  const records = values.filter(row => row[0] && (!wanted || String(row[1]).trim().toLowerCase() === wanted)).map(row => {
    let candidates = {};
    try { candidates = row[8] ? JSON.parse(row[8]) : {}; } catch (err) {}
    return {
      electionId: row[1], locationId: row[2], locationName: row[3],
      registered: Number(row[4]) || 0, cast: Number(row[5]) || 0,
      valid: Number(row[6]) || 0, rejected: Number(row[7]) || 0,
      candidates: candidates, createdAt: row[9], updatedAt: row[10], updatedBy: row[11],
      gpsLatitude: row[13] || "", gpsLongitude: row[14] || "",
      gpsAccuracyMetres: row[15] || "", gpsCapturedAt: row[16] || ""
    };
  });
  return { success: true, records: records, count: records.length };
}

/* =========================================
   VOTER REGISTRATION FORM
   Note: National ID is deliberately NOT stored by this endpoint.
========================================= */
function getVoterRegistrationSheet() {
  return getOrCreateSheet(VOTER_REGISTRATIONS_SHEET, VOTER_REGISTRATION_HEADERS);
}

function registerVoter(data) {
  if (!data || typeof data !== "object") throw new Error("Registration data is missing.");
  const name = String(data.name || "").trim();
  const subLocationId = String(data.subLocationId || "").trim();
  const locationId = String(data.locationId || "").trim();
  const divisionId = String(data.divisionId || "").trim();
  const districtId = String(data.districtId || "").trim();
  if (!name || !subLocationId || !locationId || !divisionId || !districtId) {
    throw new Error("Name, sub-location, location, division and district are required.");
  }
  if (data.consent !== true && String(data.consent).toLowerCase() !== "true") {
    throw new Error("Consent is required before submitting registration information.");
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getVoterRegistrationSheet();
    const registrationId = Utilities.getUuid();
    const now = new Date().toISOString();
    sheet.appendRow([
      registrationId, safeCell(name), safeCell(subLocationId), safeCell(locationId),
      safeCell(divisionId), safeCell(districtId), now, true
    ]);
    SpreadsheetApp.flush();
    return { success: true, message: "Registration information saved centrally.", registrationId: registrationId, nationalIdStored: false };
  } finally { lock.releaseLock(); }
}

/* =========================================
   2027 ELECTIONS OPINION POLL
   This is an independent opinion poll, not an official voter register or IEBC service.
========================================= */
function getOpinionPollSheet() {
  return getOrCreateSheet(OPINION_POLLS_SHEET, OPINION_POLL_HEADERS);
}

function submitOpinionPoll(data) {
  if (!data || typeof data !== "object") throw new Error("Opinion poll data is missing.");
  const station = String(data.pollStation || data.station || "").trim();
  const positions = [
    ["presidentialCandidate", "Presidential candidate"],
    ["governorCandidate", "Governor candidate"],
    ["senatorCandidate", "Senator candidate"],
    ["mpCandidate", "Member of Parliament candidate"],
    ["womanRepCandidate", "Woman Representative candidate"],
    ["mcaCandidate", "Member of County Assembly candidate"]
  ];
  if (!station) throw new Error("Please select a poll station.");
  positions.forEach(pair => { if (!String(data[pair[0]] || "").trim()) throw new Error(pair[1] + " is required."); });
  if (data.consent !== true && String(data.consent).toLowerCase() !== "true") {
    throw new Error("Explicit consent is required to submit political opinion data.");
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getOpinionPollSheet();
    const responseId = Utilities.getUuid();
    const submittedAt = new Date().toISOString();
    sheet.appendRow([
      responseId, safeCell(station), safeCell(data.presidentialCandidate),
      safeCell(data.governorCandidate), safeCell(data.senatorCandidate),
      safeCell(data.mpCandidate), safeCell(data.womanRepCandidate), safeCell(data.mcaCandidate),
      safeCell(String(data.reason || data.pollReason || "").trim().slice(0, 1000)), true, submittedAt
    ]);
    SpreadsheetApp.flush();
    return { success: true, message: "Opinion poll response saved centrally.", responseId: responseId, submittedAt: submittedAt };
  } finally { lock.releaseLock(); }
}

function getNationwideOpinionPoll() {
  const sheet = getOpinionPollSheet();
  const lastRow = sheet.getLastRow();
  const rows = lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, OPINION_POLL_HEADERS.length).getDisplayValues();
  const records = rows.filter(row => row[0]).map(row => ({
    responseId: row[0], pollStation: row[1], presidentialCandidate: row[2],
    governorCandidate: row[3], senatorCandidate: row[4], mpCandidate: row[5],
    womanRepCandidate: row[6], mcaCandidate: row[7], submittedAt: row[10]
  }));
  const fields = {
    presidential: "presidentialCandidate", governor: "governorCandidate",
    senator: "senatorCandidate", mp: "mpCandidate",
    womanRepresentative: "womanRepCandidate", mca: "mcaCandidate"
  };
  const totals = {};
  Object.keys(fields).forEach(position => {
    const counts = {};
    records.forEach(record => {
      const candidate = String(record[fields[position]] || "").trim();
      if (candidate) counts[candidate] = (counts[candidate] || 0) + 1;
    });
    const total = Object.keys(counts).reduce((sum, name) => sum + counts[name], 0);
    const results = Object.keys(counts).map(candidate => ({
      candidate: candidate, votes: counts[candidate],
      percentage: total ? Number((counts[candidate] * 100 / total).toFixed(3)) : 0
    })).sort((a, b) => b.votes - a.votes || a.candidate.localeCompare(b.candidate));
    totals[position] = { totalResponses: total, results: results };
  });
  const stationCounts = {};
  records.forEach(record => {
    const station = String(record.pollStation || "Unknown");
    stationCounts[station] = (stationCounts[station] || 0) + 1;
  });
  const stations = Object.keys(stationCounts).map(station => ({ station: station, responses: stationCounts[station], percentage: records.length ? Number((stationCounts[station] * 100 / records.length).toFixed(3)) : 0 })).sort((a, b) => a.station.localeCompare(b.station));
  return {
    success: true,
    pollName: "Independent 2027 Elections Opinion Poll",
    totalResponses: records.length,
    positions: totals,
    stations: stations,
    note: "Percentages are based on submitted opinion-poll responses, not official voter registration, turnout, or election results."
  };
}

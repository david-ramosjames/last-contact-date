// =============================================================================
// QUO CONTACTS SHEET SYNC — Google Apps Script
// =============================================================================
// Syncs all Quo (formerly OpenPhone) contacts and their communication activity
// into a single Google Sheet. One row per contact, updated in place.
//
// API base: https://api.openphone.com/v1
// Auth: Bearer token via API key
// =============================================================================

// ─── CONFIGURATION ───────────────────────────────────────────────────────────

var CONFIG = {
  // Quo / OpenPhone API key — set this in Script Properties instead of hardcoding.
  // Go to Project Settings → Script Properties → Add: QUO_API_KEY = your key
  API_KEY: PropertiesService.getScriptProperties().getProperty('QUO_API_KEY') || 'YOUR_API_KEY_HERE',

  API_BASE: 'https://api.openphone.com/v1',

  // Name of the sheet tab to use (created automatically if missing)
  SHEET_NAME: 'Quo Contacts',

  // Pagination page sizes (max allowed by API)
  CONTACTS_PAGE_SIZE: 50,   // API max for contacts is 50
  CALLS_PAGE_SIZE: 100,     // API max for calls is 100
  MESSAGES_PAGE_SIZE: 50,   // conservative default for messages

  // Retry / rate-limit settings
  MAX_RETRIES: 4,
  INITIAL_BACKOFF_MS: 1000,

  // Chunked execution — the Apps Script 6-minute limit makes it hard to
  // process thousands of contacts in one go, so we process a chunk per run
  // and schedule a continuation trigger to pick up where we left off.
  CHUNK_SIZE: 20,               // active contacts processed per run
  CHUNK_DELAY_SECONDS: 90,      // wait between chunks
  MAX_RUNTIME_MS: 4 * 60 * 1000, // stop processing after ~4 minutes, safely under the 6-min hard limit

  // How far back to look for calls/messages (ISO 8601). Set to null for all time.
  // Example: '2024-01-01T00:00:00Z'
  LOOKBACK_DATE: null
};

// Script Properties keys used for resumable chunked sync
var PROP_SYNC_OFFSET = 'QUO_SYNC_OFFSET';
var PROP_SYNC_STARTED_AT = 'QUO_SYNC_STARTED_AT';
var CONTINUATION_TRIGGER_FN = 'quoSyncContinue_';

// ─── COLUMN HEADERS (exact order as specified) ───────────────────────────────

var HEADERS = [
  'Contact ID',                     // A  (1)
  'Name',                           // B  (2)
  'Company',                        // C  (3)
  'Primary Phone',                  // D  (4)
  'Email',                          // E  (5)
  'Last Completed Inbound Call At', // F  (6)
  'Last Completed Outbound Call At',// G  (7)
  'Last Completed Human Call At',   // H  (8)
  'Last Inbound Text At',          // I  (9)
  'Last Outbound Text At',         // J  (10)
  'Last Text At',                  // K  (11)
  'Last Missed Call At',           // L  (12)
  'Last Missed Inbound Call At',   // M  (13)
  'Last Missed Outbound Call At',  // N  (14)
  'Last Sona Call At',             // O  (15)
  'Last Sona Inbound Call At',     // P  (16)
  'Last Sona Outbound Call At',    // Q  (17)
  'Last Voicemail At',             // R  (18)
  'Last Inbound Voicemail At',     // S  (19)
  'Last Outbound Voicemail At',    // T  (20)
  'Last Language Used',            // U  (21)
  'Last Outbound Call At',         // V  (22)
  'Last Any Call At',              // W  (23)
  'Last Any Contact At',           // X  (24)
  'Last Contact Type',             // Y  (25)
  'Last Contact Direction',        // Z  (26)
  'Days Since Last Contact',       // AA (27)
  'Last Sync At'                   // AB (28)
];

// Column index lookup (0-based)
var COL = {};
HEADERS.forEach(function(h, i) { COL[h] = i; });


// =============================================================================
// CUSTOM MENU
// =============================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Quo Sync')
    .addItem('Sync Contacts Now', 'syncQuoContactsToSheet')
    .addItem('Cancel In-Progress Sync', 'cancelQuoSync')
    .addItem('Setup Sheet', 'setupSheet')
    .addToUi();
}

/**
 * Cancels a chunked sync that's in progress — clears progress state and
 * removes the continuation trigger.
 */
function cancelQuoSync() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROP_SYNC_OFFSET);
  props.deleteProperty(PROP_SYNC_STARTED_AT);
  clearContinuationTriggers_();
  Logger.log('In-progress sync cancelled.');
}


// =============================================================================
// MAIN ENTRY POINT
// =============================================================================

/**
 * Main sync function. Call this manually or via a time-driven trigger.
 */
/**
 * Main entry point. Starts a fresh sync run (resets any in-progress state).
 */
function syncQuoContactsToSheet() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROP_SYNC_OFFSET);
  props.setProperty(PROP_SYNC_STARTED_AT, new Date().toISOString());
  clearContinuationTriggers_();
  runSyncChunk_();
}

/**
 * Continuation callback fired by the time-based trigger we schedule
 * between chunks. Resumes from the saved offset.
 */
function quoSyncContinue_() {
  runSyncChunk_();
}

/**
 * Processes one chunk of contacts, then either schedules the next chunk
 * or wraps up the sync.
 */
function runSyncChunk_() {
  var runStart = new Date();
  var props = PropertiesService.getScriptProperties();
  var offset = parseInt(props.getProperty(PROP_SYNC_OFFSET) || '0', 10);
  var syncStartedAt = props.getProperty(PROP_SYNC_STARTED_AT) || runStart.toISOString();
  var isFirstChunk = (offset === 0);

  Logger.log('=== Quo Sync chunk starting (offset=' + offset + ') ===');

  var sheet = setupSheet();
  var phoneNumberIds = fetchAllPhoneNumberIds();
  Logger.log('Quo phone numbers: ' + phoneNumberIds.length);

  var contacts = fetchAllContacts();
  Logger.log('Total contacts fetched: ' + contacts.length);

  // Write basic rows for all contacts on the first chunk so the sheet is
  // populated immediately — activity columns will be filled chunk by chunk.
  if (isFirstChunk) {
    writeBasicContactRows_(sheet, contacts, syncStartedAt);
  }

  // Build a set of normalized contact phones — we filter the conversation
  // map to these so we only fetch activity for contacts we care about.
  var contactPhoneSet = {};
  var contactsByPhone = {};
  for (var c = 0; c < contacts.length; c++) {
    var phone = normalizePhone(getPrimaryPhone((contacts[c].defaultFields || {}).phoneNumbers));
    if (phone) {
      contactPhoneSet[phone] = true;
      contactsByPhone[phone] = contacts[c];
    }
  }

  // Build conversation map filtered to contacts only
  var conversationMap = buildConversationMap(phoneNumberIds, contactPhoneSet);
  Logger.log('Active contact conversations: ' + Object.keys(conversationMap).length);

  // Deterministic ordering so offset is stable across runs
  var activePhones = Object.keys(conversationMap).sort();
  var total = activePhones.length;

  // Read existing rows once for in-place updates
  var existingMap = readExistingRows(sheet);

  var processed = 0;
  var cutoff = offset + CONFIG.CHUNK_SIZE;

  for (var i = offset; i < total && i < cutoff; i++) {
    // Respect the wall-clock safety budget too, in case fetches are slow
    if (new Date() - runStart > CONFIG.MAX_RUNTIME_MS) {
      Logger.log('Chunk hit runtime budget at offset ' + i);
      break;
    }

    var phone = activePhones[i];
    var contact = contactsByPhone[phone];
    if (!contact) { processed++; continue; }

    var pnIds = conversationMap[phone];
    var calls = [];
    var messages = [];

    for (var pn = 0; pn < pnIds.length; pn++) {
      calls = calls.concat(fetchAllPages('/calls', {
        phoneNumberId: pnIds[pn], participants: [phone]
      }, CONFIG.CALLS_PAGE_SIZE));

      messages = messages.concat(fetchAllPages('/messages', {
        phoneNumberId: pnIds[pn], participants: [phone]
      }, CONFIG.MESSAGES_PAGE_SIZE));
    }

    var voicemailCache = fetchVoicemailsForCalls(calls);
    var classified = classifyCalls(calls, voicemailCache);
    var textActivity = computeTextActivityFromMessages(messages);

    // Build updated row
    var df = contact.defaultFields || {};
    var name = ((df.firstName || '') + ' ' + (df.lastName || '')).trim();
    var row = buildContactRow(
      contact.id,
      name,
      df.company || '',
      getPrimaryPhone(df.phoneNumbers),
      getPrimaryEmail(df.emails),
      classified,
      textActivity,
      new Date()
    );

    // Update in place (row must already exist from the first-chunk basic write)
    if (existingMap[contact.id]) {
      sheet.getRange(existingMap[contact.id].rowIndex, 1, 1, HEADERS.length)
           .setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    processed++;
  }

  var newOffset = offset + processed;
  Logger.log('Processed ' + processed + ' contacts this chunk (' + offset + ' → ' + newOffset + ' of ' + total + ')');

  if (newOffset >= total) {
    // Done
    props.deleteProperty(PROP_SYNC_OFFSET);
    props.deleteProperty(PROP_SYNC_STARTED_AT);
    clearContinuationTriggers_();
    Logger.log('=== Quo Sync complete ===');
  } else {
    props.setProperty(PROP_SYNC_OFFSET, String(newOffset));
    scheduleContinuation_();
    Logger.log('Scheduled next chunk in ' + CONFIG.CHUNK_DELAY_SECONDS + 's; offset now ' + newOffset);
  }
}

/**
 * Writes a starter row for every contact with basic fields filled in.
 * Activity columns are left blank and get populated as each chunk runs.
 */
function writeBasicContactRows_(sheet, contacts, syncStartedAt) {
  var existingMap = readExistingRows(sheet);
  var rowsToUpdate = [];
  var rowsToAppend = [];

  for (var i = 0; i < contacts.length; i++) {
    var contact = contacts[i];
    if (!contact.id) continue;

    var df = contact.defaultFields || {};
    var name = ((df.firstName || '') + ' ' + (df.lastName || '')).trim();
    var phone = getPrimaryPhone(df.phoneNumbers);
    var email = getPrimaryEmail(df.emails);

    if (existingMap[contact.id]) {
      // Only overwrite the basic identity columns; preserve any existing
      // activity data from previous syncs until this run's activity fetch
      // replaces it.
      var existing = existingMap[contact.id].data.slice();
      existing[COL['Contact ID']] = contact.id;
      existing[COL['Name']] = name;
      existing[COL['Company']] = df.company || '';
      existing[COL['Primary Phone']] = phone;
      existing[COL['Email']] = email;
      existing[COL['Last Sync At']] = syncStartedAt;
      rowsToUpdate.push({ rowIndex: existingMap[contact.id].rowIndex, data: existing });
    } else {
      var row = new Array(HEADERS.length).fill('');
      row[COL['Contact ID']] = contact.id;
      row[COL['Name']] = name;
      row[COL['Company']] = df.company || '';
      row[COL['Primary Phone']] = phone;
      row[COL['Email']] = email;
      row[COL['Last Sync At']] = syncStartedAt;
      rowsToAppend.push(row);
    }
  }

  for (var u = 0; u < rowsToUpdate.length; u++) {
    sheet.getRange(rowsToUpdate[u].rowIndex, 1, 1, HEADERS.length)
         .setValues([rowsToUpdate[u].data]);
  }

  if (rowsToAppend.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, HEADERS.length)
         .setValues(rowsToAppend);
  }

  Logger.log('Basic rows written: ' + rowsToUpdate.length + ' updated, ' + rowsToAppend.length + ' appended');
}

/**
 * Creates a one-shot time-based trigger to resume the sync after a delay.
 */
function scheduleContinuation_() {
  clearContinuationTriggers_();
  ScriptApp.newTrigger(CONTINUATION_TRIGGER_FN)
    .timeBased()
    .after(CONFIG.CHUNK_DELAY_SECONDS * 1000)
    .create();
}

/**
 * Removes any previously-scheduled continuation triggers.
 */
function clearContinuationTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var t = 0; t < triggers.length; t++) {
    if (triggers[t].getHandlerFunction() === CONTINUATION_TRIGGER_FN) {
      ScriptApp.deleteTrigger(triggers[t]);
    }
  }
}


// =============================================================================
// SHEET SETUP & I/O
// =============================================================================

/**
 * Creates the sheet and headers if they don't exist. Returns the sheet.
 */
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    Logger.log('Created new sheet: ' + CONFIG.SHEET_NAME);
  }

  // Write headers if row 1 is empty or doesn't match
  var existingHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  var needsHeaders = !existingHeaders[0] || existingHeaders[0] !== HEADERS[0];

  if (needsHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    Logger.log('Headers written to sheet.');
  }

  return sheet;
}

/**
 * Reads all existing data rows into a map: { contactId: { rowIndex: N, data: [...] } }
 * rowIndex is 1-based (sheet row number).
 */
function readExistingRows(sheet) {
  var map = {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return map; // no data rows

  var data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

  for (var i = 0; i < data.length; i++) {
    var contactId = String(data[i][COL['Contact ID']]).trim();
    if (contactId) {
      map[contactId] = {
        rowIndex: i + 2,  // +2 because data starts at row 2, array is 0-based
        data: data[i]
      };
    }
  }

  return map;
}

// =============================================================================
// API FETCH HELPERS
// =============================================================================

/**
 * Makes a GET request to the Quo API with retry and backoff for rate limits.
 * Returns parsed JSON response or null on failure.
 */
function quoApiFetch(endpoint, queryParams) {
  var url = CONFIG.API_BASE + endpoint;

  // Build query string
  if (queryParams) {
    var parts = [];
    for (var key in queryParams) {
      if (queryParams[key] === null || queryParams[key] === undefined) continue;
      var val = queryParams[key];
      // Handle array parameters (e.g., participants[])
      if (Array.isArray(val)) {
        for (var a = 0; a < val.length; a++) {
          parts.push(encodeURIComponent(key) + '=' + encodeURIComponent(val[a]));
        }
      } else {
        parts.push(encodeURIComponent(key) + '=' + encodeURIComponent(val));
      }
    }
    if (parts.length > 0) {
      url += '?' + parts.join('&');
    }
  }

  var options = {
    method: 'get',
    headers: {
      // Quo / OpenPhone expects the API key directly — NO "Bearer " prefix
      'Authorization': CONFIG.API_KEY
    },
    muteHttpExceptions: true
  };

  var backoff = CONFIG.INITIAL_BACKOFF_MS;

  for (var attempt = 0; attempt <= CONFIG.MAX_RETRIES; attempt++) {
    try {
      var response = UrlFetchApp.fetch(url, options);
      var code = response.getResponseCode();

      if (code === 200) {
        return JSON.parse(response.getContentText());
      }

      // Rate limited (429) or server error (5xx) — retry with backoff
      if (code === 429 || code >= 500) {
        Logger.log('API returned ' + code + ' for ' + endpoint + ', retrying in ' + backoff + 'ms (attempt ' + (attempt + 1) + ')');
        Utilities.sleep(backoff);
        backoff *= 2;
        continue;
      }

      // Other client errors — don't retry
      Logger.log('API error ' + code + ' for ' + endpoint + ': ' + response.getContentText().substring(0, 200));
      return null;

    } catch (e) {
      Logger.log('Fetch exception for ' + endpoint + ': ' + e.message);
      if (attempt < CONFIG.MAX_RETRIES) {
        Utilities.sleep(backoff);
        backoff *= 2;
      }
    }
  }

  Logger.log('All retries exhausted for ' + endpoint);
  return null;
}

/**
 * Paginates through a Quo API list endpoint, collecting all items.
 * @param {string} endpoint - API path (e.g., '/contacts')
 * @param {object} baseParams - Query params (excluding pageToken)
 * @param {string} maxResultsKey - The param name for page size (usually 'maxResults')
 * @param {number} pageSize - Number of results per page
 * @returns {Array} All items across all pages
 */
function fetchAllPages(endpoint, baseParams, pageSize) {
  var allItems = [];
  var pageToken = null;

  do {
    var params = {};
    for (var k in baseParams) {
      params[k] = baseParams[k];
    }
    params['maxResults'] = pageSize;
    if (pageToken) {
      params['pageToken'] = pageToken;
    }

    var result = quoApiFetch(endpoint, params);
    if (!result || !result.data) break;

    allItems = allItems.concat(result.data);
    pageToken = result.nextPageToken || null;

  } while (pageToken);

  return allItems;
}


// =============================================================================
// DATA FETCHING FUNCTIONS
// =============================================================================

/**
 * Fetches all Quo phone number IDs (needed for calls/messages endpoints).
 * Returns array of phone number ID strings.
 */
function fetchAllPhoneNumberIds() {
  var result = quoApiFetch('/phone-numbers', {});
  if (!result || !result.data) return [];

  return result.data.map(function(pn) {
    return pn.id;
  });
}

/**
 * Fetches ALL contacts from Quo, handling pagination.
 */
function fetchAllContacts() {
  return fetchAllPages('/contacts', {}, CONFIG.CONTACTS_PAGE_SIZE);
}

/**
 * Fetches all conversations and builds a map of contactPhone → [phoneNumberId, ...]
 * Filtered to only include phones that match a known contact, so we don't
 * waste time fetching activity for people who aren't in the contacts list.
 *
 * The /conversations endpoint does NOT require a participants filter.
 */
function buildConversationMap(phoneNumberIds, contactPhoneSet) {
  var map = {};

  for (var p = 0; p < phoneNumberIds.length; p++) {
    var conversations = fetchAllPages('/conversations', {
      phoneNumbers: [phoneNumberIds[p]]
    }, 50);

    for (var c = 0; c < conversations.length; c++) {
      var conv = conversations[c];
      var participants = conv.participants || [];
      var pnId = conv.phoneNumberId || phoneNumberIds[p];

      for (var x = 0; x < participants.length; x++) {
        var normalized = normalizePhone(participants[x]);
        if (!normalized) continue;
        // Skip phones that don't belong to a known contact
        if (contactPhoneSet && !contactPhoneSet[normalized]) continue;

        if (!map[normalized]) map[normalized] = [];
        if (map[normalized].indexOf(pnId) === -1) {
          map[normalized].push(pnId);
        }
      }
    }
  }

  return map;
}

/**
 * Scans a contact's messages to find the most recent inbound/outbound timestamps
 * plus the content of the most recent inbound text (for language detection).
 */
function computeTextActivityFromMessages(messages) {
  var lastInbound = '';
  var lastOutbound = '';
  var lastInboundText = '';

  for (var m = 0; m < messages.length; m++) {
    var msg = messages[m];
    var ts = msg.createdAt || '';
    // ── FIELD MAPPING: direction is "incoming" or "outgoing" ──
    var dir = (msg.direction || '').toLowerCase();

    if (dir === 'incoming' && ts > lastInbound) {
      lastInbound = ts;
      // ── FIELD MAPPING: message body field is `text` ──
      lastInboundText = msg.text || '';
    } else if (dir === 'outgoing' && ts > lastOutbound) {
      lastOutbound = ts;
    }
  }

  return {
    lastInbound: lastInbound,
    lastOutbound: lastOutbound,
    lastInboundText: lastInboundText
  };
}

/**
 * Fetches voicemail details for missed calls.
 * Voicemails most commonly arrive on missed calls; limiting to those keeps
 * the request count manageable on large datasets.
 * Returns a cache: { callId: { transcript, duration, direction, createdAt } }
 */
function fetchVoicemailsForCalls(allCalls) {
  var cache = {};

  for (var i = 0; i < allCalls.length; i++) {
    var call = allCalls[i];
    var status = (call.status || '').toLowerCase();
    if (status !== 'missed') continue;

    var vmResult = quoApiFetch('/call-voicemails/' + call.id, {});
    if (vmResult && vmResult.data && vmResult.data.transcript) {
      cache[call.id] = {
        transcript: vmResult.data.transcript || '',
        duration: vmResult.data.duration || 0,
        direction: call.direction,
        createdAt: call.createdAt
      };
    }
  }

  return cache;
}


// =============================================================================
// CALL CLASSIFICATION
// =============================================================================

/**
 * Classifies an array of calls into categories.
 * Returns an object with the latest timestamp for each category,
 * plus voicemail transcript info.
 */
function classifyCalls(calls, voicemailCache) {
  var result = {
    lastCompletedInbound: '',
    lastCompletedOutbound: '',
    lastMissedInbound: '',
    lastMissedOutbound: '',
    lastSonaInbound: '',
    lastSonaOutbound: '',
    lastVoicemailInbound: '',
    lastVoicemailOutbound: '',
    // Most recent INBOUND voicemail transcript — used only for language detection.
    // We care about what the contact said, not what we said.
    lastInboundVoicemailTranscript: '',
    lastInboundVoicemailAt: ''
  };

  for (var i = 0; i < calls.length; i++) {
    var call = calls[i];
    var ts = call.createdAt || '';
    // ── FIELD MAPPING: direction values are "incoming" / "outgoing" ──
    var dir = (call.direction || '').toLowerCase();
    var status = (call.status || '').toLowerCase();
    // ── FIELD MAPPING: aiHandled is "ai-agent" when Sona handled, null otherwise ──
    var isSona = !!(call.aiHandled);
    var hasVoicemail = !!(voicemailCache[call.id]);

    // --- Voicemail tracking (any call with a voicemail) ---
    if (hasVoicemail) {
      var vmData = voicemailCache[call.id];
      if (dir === 'incoming' && ts > result.lastVoicemailInbound) {
        result.lastVoicemailInbound = ts;
        // Keep the most recent inbound transcript for language detection
        if (ts > result.lastInboundVoicemailAt) {
          result.lastInboundVoicemailAt = ts;
          result.lastInboundVoicemailTranscript = vmData.transcript || '';
        }
      }
      if (dir === 'outgoing' && ts > result.lastVoicemailOutbound) {
        result.lastVoicemailOutbound = ts;
      }
    }

    // --- Sona / AI-handled calls ---
    if (isSona) {
      if (dir === 'incoming' && ts > result.lastSonaInbound) {
        result.lastSonaInbound = ts;
      }
      if (dir === 'outgoing' && ts > result.lastSonaOutbound) {
        result.lastSonaOutbound = ts;
      }
      continue; // Sona calls don't count as completed human or missed
    }

    // --- Completed human calls (NOT Sona) ---
    if (status === 'completed') {
      if (dir === 'incoming' && ts > result.lastCompletedInbound) {
        result.lastCompletedInbound = ts;
      }
      if (dir === 'outgoing' && ts > result.lastCompletedOutbound) {
        result.lastCompletedOutbound = ts;
      }
    }

    // --- Missed calls (NOT Sona) ---
    if (status === 'missed') {
      if (dir === 'incoming' && ts > result.lastMissedInbound) {
        result.lastMissedInbound = ts;
      }
      if (dir === 'outgoing' && ts > result.lastMissedOutbound) {
        result.lastMissedOutbound = ts;
      }
    }
  }

  return result;
}


// =============================================================================
// ROW BUILDING & ROLLUPS
// =============================================================================

/**
 * Builds a complete row array for one contact.
 */
function buildContactRow(contactId, name, company, primaryPhone, email,
                         classified, textActivity, syncTime) {
  // Rollup: Last Completed Human Call At
  var lastCompletedHumanCall = latestOf([
    classified.lastCompletedInbound,
    classified.lastCompletedOutbound
  ]);

  // Rollup: Last Text At
  var lastText = latestOf([textActivity.lastInbound, textActivity.lastOutbound]);

  // Rollup: Last Missed Call At
  var lastMissedCall = latestOf([
    classified.lastMissedInbound,
    classified.lastMissedOutbound
  ]);

  // Rollup: Last Sona Call At
  var lastSonaCall = latestOf([
    classified.lastSonaInbound,
    classified.lastSonaOutbound
  ]);

  // Rollup: Last Voicemail At
  var lastVoicemail = latestOf([
    classified.lastVoicemailInbound,
    classified.lastVoicemailOutbound
  ]);

  // Rollup: Last Outbound Call At — any outbound call attempt
  // (completed, missed, Sona, or voicemail we left).
  var lastOutboundCall = latestOf([
    classified.lastCompletedOutbound,
    classified.lastMissedOutbound,
    classified.lastSonaOutbound,
    classified.lastVoicemailOutbound
  ]);

  // Rollup: Last Any Call At (completed human + missed + sona + voicemail)
  var lastAnyCall = latestOf([
    lastCompletedHumanCall,
    lastMissedCall,
    lastSonaCall,
    lastVoicemail
  ]);

  // Rollup: Last Any Contact At (calls + texts)
  var lastAnyContact = latestOf([lastAnyCall, lastText]);

  // Language detection — pick the most recent inbound sample from the contact.
  // Voicemail transcript vs inbound text; whichever is more recent wins.
  var languageSample = '';
  if (classified.lastInboundVoicemailAt > textActivity.lastInbound) {
    languageSample = classified.lastInboundVoicemailTranscript;
  } else {
    languageSample = textActivity.lastInboundText || classified.lastInboundVoicemailTranscript;
  }
  var lastLanguageUsed = detectLanguage(languageSample);

  // Determine Last Contact Type and Direction
  var lastContactInfo = determineLastContactType(classified, textActivity, lastAnyContact);

  // Days Since Last Contact
  var daysSince = '';
  if (lastAnyContact) {
    var lastDate = new Date(lastAnyContact);
    var now = new Date();
    daysSince = Math.floor((now - lastDate) / (1000 * 60 * 60 * 24));
  }

  // Build the row in exact column order
  var row = new Array(HEADERS.length);
  row[COL['Contact ID']]                      = contactId;
  row[COL['Name']]                             = name;
  row[COL['Company']]                          = company;
  row[COL['Primary Phone']]                    = primaryPhone;
  row[COL['Email']]                            = email;
  row[COL['Last Completed Inbound Call At']]   = classified.lastCompletedInbound;
  row[COL['Last Completed Outbound Call At']]  = classified.lastCompletedOutbound;
  row[COL['Last Completed Human Call At']]     = lastCompletedHumanCall;
  row[COL['Last Inbound Text At']]             = textActivity.lastInbound;
  row[COL['Last Outbound Text At']]            = textActivity.lastOutbound;
  row[COL['Last Text At']]                     = lastText;
  row[COL['Last Missed Call At']]              = lastMissedCall;
  row[COL['Last Missed Inbound Call At']]      = classified.lastMissedInbound;
  row[COL['Last Missed Outbound Call At']]     = classified.lastMissedOutbound;
  row[COL['Last Sona Call At']]                = lastSonaCall;
  row[COL['Last Sona Inbound Call At']]        = classified.lastSonaInbound;
  row[COL['Last Sona Outbound Call At']]       = classified.lastSonaOutbound;
  row[COL['Last Voicemail At']]                = lastVoicemail;
  row[COL['Last Inbound Voicemail At']]        = classified.lastVoicemailInbound;
  row[COL['Last Outbound Voicemail At']]       = classified.lastVoicemailOutbound;
  row[COL['Last Language Used']]               = lastLanguageUsed;
  row[COL['Last Outbound Call At']]            = lastOutboundCall;
  row[COL['Last Any Call At']]                 = lastAnyCall;
  row[COL['Last Any Contact At']]              = lastAnyContact;
  row[COL['Last Contact Type']]                = lastContactInfo.type;
  row[COL['Last Contact Direction']]           = lastContactInfo.direction;
  row[COL['Days Since Last Contact']]          = daysSince;
  row[COL['Last Sync At']]                     = syncTime.toISOString();

  return row;
}

/**
 * Returns the latest (max) ISO timestamp from an array of timestamps.
 * Ignores empty/falsy values. Returns '' if none.
 */
function latestOf(timestamps) {
  var latest = '';
  for (var i = 0; i < timestamps.length; i++) {
    if (timestamps[i] && timestamps[i] > latest) {
      latest = timestamps[i];
    }
  }
  return latest;
}

/**
 * Determines the type and direction of the most recent contact event.
 * Returns { type: 'completed_inbound_call', direction: 'inbound' }
 */
function determineLastContactType(classified, textActivity, lastAnyContact) {
  if (!lastAnyContact) return { type: '', direction: '' };

  // Build a list of all events with their timestamps, types, and directions
  var events = [
    { ts: classified.lastCompletedInbound,   type: 'completed_inbound_call',  direction: 'inbound'  },
    { ts: classified.lastCompletedOutbound,  type: 'completed_outbound_call', direction: 'outbound' },
    { ts: classified.lastMissedInbound,      type: 'missed_inbound_call',     direction: 'inbound'  },
    { ts: classified.lastMissedOutbound,     type: 'missed_outbound_call',    direction: 'outbound' },
    { ts: classified.lastSonaInbound,        type: 'sona_inbound_call',       direction: 'inbound'  },
    { ts: classified.lastSonaOutbound,       type: 'sona_outbound_call',      direction: 'outbound' },
    { ts: classified.lastVoicemailInbound,   type: 'inbound_voicemail',       direction: 'inbound'  },
    { ts: classified.lastVoicemailOutbound,  type: 'outbound_voicemail',      direction: 'outbound' },
    { ts: textActivity.lastInbound,          type: 'inbound_text',            direction: 'inbound'  },
    { ts: textActivity.lastOutbound,         type: 'outbound_text',           direction: 'outbound' }
  ];

  // Find the event matching lastAnyContact
  var best = { type: '', direction: '' };
  var bestTs = '';

  for (var i = 0; i < events.length; i++) {
    if (events[i].ts && events[i].ts > bestTs) {
      bestTs = events[i].ts;
      best.type = events[i].type;
      best.direction = events[i].direction;
    }
  }

  return best;
}


// =============================================================================
// UTILITY HELPERS
// =============================================================================

/**
 * Normalizes a phone number to E.164-ish format for matching.
 * Strips everything except digits, then prepends '+' if needed.
 * Returns '' if input is empty.
 */
function normalizePhone(phone) {
  if (!phone) return '';
  var digits = String(phone).replace(/[^\d+]/g, '');
  if (!digits) return '';
  // If it already starts with '+', keep it
  if (digits.charAt(0) === '+') return digits;
  // If 10 digits, assume US and prepend +1
  if (digits.length === 10) return '+1' + digits;
  // If 11 digits starting with 1, prepend +
  if (digits.length === 11 && digits.charAt(0) === '1') return '+' + digits;
  // Otherwise prepend +
  return '+' + digits;
}

/**
 * Gets the first phone number value from a contact's phoneNumbers array.
 * ── FIELD MAPPING: phoneNumbers is an array of { name, value, id } ──
 */
function getPrimaryPhone(phoneNumbers) {
  if (!phoneNumbers || !Array.isArray(phoneNumbers) || phoneNumbers.length === 0) return '';
  return phoneNumbers[0].value || '';
}

/**
 * Gets the first email value from a contact's emails array.
 * ── FIELD MAPPING: emails is an array of { name, value, id } ──
 */
function getPrimaryEmail(emails) {
  if (!emails || !Array.isArray(emails) || emails.length === 0) return '';
  return emails[0].value || '';
}

/**
 * Detects the language of a text sample.
 * Uses Google Cloud Translation API when GOOGLE_TRANSLATE_API_KEY is configured,
 * otherwise falls back to a lightweight keyword-based heuristic.
 * Returns 'English', 'Spanish', or '' if unclear.
 */
// Module-level state for language detection.
// Cache: same text → same result, so we don't re-detect duplicates.
// Disabled flag: once Google Translate returns 429, skip it for the rest of
// the run so we don't waste time on endless rate-limit hits.
var _langCache = {};
var _googleTranslateDisabled = false;

/**
 * Detects the language of a text sample. Returns 'English', 'Spanish', or ''.
 * Strategy: run the local keyword heuristic first (fast, no network). Only
 * fall back to the free Google Translate endpoint for ambiguous cases, and
 * disable Google entirely once it starts rate-limiting.
 */
function detectLanguage(text) {
  if (!text || typeof text !== 'string') return '';
  var trimmed = text.trim();
  if (trimmed.length < 2) return '';

  if (_langCache[trimmed] !== undefined) return _langCache[trimmed];

  // 1. Local heuristic first — good enough for most messages, zero network
  var result = detectLanguageHeuristic(trimmed);
  if (result) {
    _langCache[trimmed] = result;
    return result;
  }

  // 2. Fall back to Google only if it hasn't been disabled
  if (!_googleTranslateDisabled) {
    var googleResult = detectLanguageViaGoogleTranslate(trimmed);
    if (googleResult === null) {
      // Network error or 429 — disable for rest of run
      _googleTranslateDisabled = true;
      Logger.log('Google Translate disabled for this run (rate-limited or error)');
    } else {
      _langCache[trimmed] = googleResult;
      return googleResult;
    }
  }

  _langCache[trimmed] = '';
  return '';
}

/**
 * Uses the free Google Translate auto-detect endpoint.
 * No API key or billing required — same endpoint the Google Translate site uses.
 * Returns 'English', 'Spanish', '' for other languages, or null on failure.
 */
function detectLanguageViaGoogleTranslate(text) {
  // Cap at 500 chars — plenty for detection, keeps the URL safe
  var snippet = text.substring(0, 500);
  var url = 'https://translate.googleapis.com/translate_a/single'
          + '?client=gtx&sl=auto&tl=en&dt=t&q=' + encodeURIComponent(snippet);

  try {
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      Logger.log('Google Translate detect returned ' + response.getResponseCode());
      return null;
    }
    // Response is a nested JSON array. The detected language code is at index [2].
    var result = JSON.parse(response.getContentText());
    var lang = (result[2] || '').toLowerCase();
    if (lang === 'es') return 'Spanish';
    if (lang === 'en') return 'English';
    return '';
  } catch (e) {
    Logger.log('Google Translate detect exception: ' + e.message);
    return null;
  }
}

/**
 * Fallback keyword-based detector. Returns 'Spanish', 'English', or ''.
 */
function detectLanguageHeuristic(text) {
  var sample = text.toLowerCase();

  if (/[ñáéíóúü¿¡]/.test(sample)) return 'Spanish';

  var SPANISH = [
    ' el ', ' la ', ' los ', ' las ', ' de ', ' que ', ' no ', ' si ',
    ' una ', ' uno ', ' por ', ' para ', ' con ', ' sin ', ' pero ',
    ' hola ', ' gracias ', ' buenos ', ' buenas ', ' necesito ', ' puede '
  ];
  var ENGLISH = [
    ' the ', ' a ', ' an ', ' is ', ' are ', ' was ', ' and ', ' or ',
    ' but ', ' to ', ' of ', ' in ', ' for ', ' with ', ' you ', ' your ',
    ' hello ', ' hi ', ' thanks ', ' please ', ' call ', ' need '
  ];

  var padded = ' ' + sample.replace(/[^\w\s]/g, ' ') + ' ';
  var es = 0, en = 0;
  for (var i = 0; i < SPANISH.length; i++) { if (padded.indexOf(SPANISH[i]) !== -1) es++; }
  for (var j = 0; j < ENGLISH.length; j++) { if (padded.indexOf(ENGLISH[j]) !== -1) en++; }

  if (es === 0 && en === 0) return '';
  if (es > en) return 'Spanish';
  if (en > es) return 'English';
  return '';
}

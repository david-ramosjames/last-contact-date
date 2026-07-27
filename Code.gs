// =============================================================================
// QUO CONTACTS SHEET SYNC - Google Apps Script
// =============================================================================
// Faster version:
// - Uses /conversations as the fast activity index.
// - Only fetches expensive /calls and /messages when a contact is new or changed.
// - Adds a menu function to sync one contact by phone and write it to the sheet.
// =============================================================================

var CONFIG = {
  API_KEY: PropertiesService.getScriptProperties().getProperty('QUO_API_KEY') || 'YOUR_API_KEY_HERE',
  API_BASE: 'https://api.openphone.com/v1',

  SHEET_NAME: 'Quo Contacts',
  CONVERSATION_CACHE_SHEET_NAME: 'Quo Sync Cache',

  CONTACTS_PAGE_SIZE: 50,
  CALLS_PAGE_SIZE: 100,
  MESSAGES_PAGE_SIZE: 100,
  USERS_PAGE_SIZE: 50,
  CONVERSATIONS_PAGE_SIZE: 50,

  MAX_RETRIES: 4,
  INITIAL_BACKOFF_MS: 1000,

  CHUNK_SIZE: 40,
  CHUNK_DELAY_SECONDS: 30,
  MAX_RUNTIME_MS: 4 * 60 * 1000
};

var PROP_SYNC_OFFSET = 'QUO_SYNC_OFFSET';
var PROP_SYNC_STARTED_AT = 'QUO_SYNC_STARTED_AT';
var CONTINUATION_TRIGGER_FN = 'quoSyncContinue_';

var HEADERS = [
  'Contact ID',
  'Name',
  'Company',
  'Primary Phone',
  'Email',
  'Last Completed Inbound Call At',
  'Last Completed Outbound Call At',
  'Last Completed Human Call At',
  'Last Inbound Text At',
  'Last Outbound Text At',
  'Last Text At',
  'Last Missed Call At',
  'Last Missed Inbound Call At',
  'Last Missed Outbound Call At',
  'Last Sona Call At',
  'Last Sona Inbound Call At',
  'Last Sona Outbound Call At',
  'Last Voicemail At',
  'Last Inbound Voicemail At',
  'Last Outbound Voicemail At',
  'Last Language Used',
  'Last Outbound Call At',
  'Last Any Call At',
  'Last Any Contact At',
  'Last Contact Type',
  'Last Contact Direction',
  'Last Conversation Activity At',
  'Days Since Last Contact',
  'Last Sync At'
];

var COL = {};
HEADERS.forEach(function(h, i) {
  COL[h] = i;
});

// =============================================================================
// MENU
// =============================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Quo Sync')
    .addItem('Sync Contacts Now', 'syncQuoContactsToSheet')
    .addItem('Sync One Contact By Phone', 'syncOneContactByPhonePrompt')
    .addItem('Cancel In-Progress Sync', 'cancelQuoSync')
    .addSeparator()
    .addItem('Debug Contact Activity By Phone', 'debugContactActivityPrompt')
    .addItem('Debug Contact Record By Phone', 'debugContactRecordPrompt')
    .addItem('Setup Sheet', 'setupSheet')
    .addToUi();
}

function cancelQuoSync() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROP_SYNC_OFFSET);
  props.deleteProperty(PROP_SYNC_STARTED_AT);
  clearContinuationTriggers_();
  Logger.log('In-progress sync cancelled.');
}

// =============================================================================
// MAIN SYNC
// =============================================================================

function syncQuoContactsToSheet() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROP_SYNC_OFFSET);
  props.setProperty(PROP_SYNC_STARTED_AT, new Date().toISOString());
  clearContinuationTriggers_();
  runSyncChunk_();
}

function quoSyncContinue_() {
  runSyncChunk_();
}

function runSyncChunk_() {
  var runStart = new Date();
  var props = PropertiesService.getScriptProperties();
  var offset = parseInt(props.getProperty(PROP_SYNC_OFFSET) || '0', 10);
  var isFirstChunk = offset === 0;

  Logger.log('=== Quo Sync chunk starting, offset=' + offset + ' ===');

  var sheet = setupSheet();
  var phoneNumberIds = fetchAllPhoneNumberIds();
  var userIds = fetchAllUserIds();

  Logger.log('Quo inbox phone numbers found: ' + phoneNumberIds.length);
  Logger.log('Quo users found: ' + userIds.length);

  var contacts = fetchAllContacts();
  contacts.sort(function(a, b) {
    return String(a.id || '').localeCompare(String(b.id || ''));
  });

  Logger.log('Total contacts fetched: ' + contacts.length);

  if (isFirstChunk) {
    writeBasicContactRows_(sheet, contacts);
  }

  var contactPhoneSet = buildContactPhoneSet_(contacts);
  var conversationActivityMap = getConversationActivityMapForRun_(phoneNumberIds, contactPhoneSet, isFirstChunk);
  var existingMap = readExistingRows(sheet);

  var processed = 0;
  var total = contacts.length;
  var cutoff = Math.min(offset + CONFIG.CHUNK_SIZE, total);

  for (var i = offset; i < cutoff; i++) {
    if (new Date() - runStart > CONFIG.MAX_RUNTIME_MS) {
      Logger.log('Chunk hit runtime budget at contact index ' + i);
      break;
    }

    var contact = contacts[i];

    if (!contact || !contact.id) {
      processed++;
      continue;
    }

    var df = contact.defaultFields || {};
    var name = ((df.firstName || '') + ' ' + (df.lastName || '')).trim();
    var primaryPhone = getPrimaryPhone(df.phoneNumbers);
    var email = getPrimaryEmail(df.emails);
    var contactPhones = getAllPhones(df.phoneNumbers);
    var lastConversationActivityAt = latestConversationActivityForPhones_(contactPhones, conversationActivityMap);
    var existing = existingMap[contact.id] || null;
    var existingConversationActivityAt = existing
      ? valueToComparableString_(existing.data[COL['Last Conversation Activity At']])
      : '';

    Logger.log('Processing contact: ' + contact.id + ' | ' + name + ' | ' + contactPhones.join(', '));
    Logger.log('Last conversation activity: ' + lastConversationActivityAt);

    var row;

    if (existing && existingConversationActivityAt === lastConversationActivityAt) {
      row = buildSkippedExistingRow_(
        existing.data,
        contact.id,
        name,
        df.company || '',
        primaryPhone,
        email,
        lastConversationActivityAt,
        new Date()
      );

      Logger.log('Skipped expensive activity fetch because conversation activity is unchanged.');
    } else {
      row = buildFreshActivityRow_(
        contact.id,
        name,
        df.company || '',
        primaryPhone,
        email,
        contactPhones,
        phoneNumberIds,
        userIds,
        lastConversationActivityAt,
        new Date()
      );
    }

    if (existing) {
      sheet.getRange(existing.rowIndex, 1, 1, HEADERS.length).setValues([row]);
    } else {
      sheet.appendRow(row);
      existingMap[contact.id] = {
        rowIndex: sheet.getLastRow(),
        data: row
      };
    }

    processed++;
  }

  var newOffset = offset + processed;
  Logger.log('Processed ' + processed + ' contacts this chunk: ' + offset + ' -> ' + newOffset + ' of ' + total);

  if (newOffset >= total) {
    props.deleteProperty(PROP_SYNC_OFFSET);
    props.deleteProperty(PROP_SYNC_STARTED_AT);
    clearContinuationTriggers_();
    Logger.log('=== Quo Sync complete ===');
  } else {
    props.setProperty(PROP_SYNC_OFFSET, String(newOffset));
    scheduleContinuation_();
    Logger.log('Scheduled next chunk in ' + CONFIG.CHUNK_DELAY_SECONDS + ' seconds; offset now ' + newOffset);
  }
}

function buildFreshActivityRow_(contactId, name, company, primaryPhone, email, contactPhones, phoneNumberIds, userIds, lastConversationActivityAt, syncTime) {
  var activity = {
    calls: [],
    messages: []
  };

  var classified = emptyClassified_();
  var textActivity = emptyTextActivity_();

  if (lastConversationActivityAt || contactPhones.length > 0) {
    activity = fetchActivityForContactPhones_(contactPhones, phoneNumberIds, userIds);
    Logger.log('Activity fetched: calls=' + activity.calls.length + ', messages=' + activity.messages.length);

    var voicemailCache = fetchVoicemailsForCalls(activity.calls);
    classified = classifyCalls(activity.calls, voicemailCache);
    textActivity = computeTextActivityFromMessages(activity.messages);
  }

  Logger.log('Latest inbound completed: ' + classified.lastCompletedInbound);
  Logger.log('Latest any call: ' + classified.lastAnyCallAt);

  return buildContactRow(
    contactId,
    name,
    company,
    primaryPhone,
    email,
    classified,
    textActivity,
    lastConversationActivityAt,
    syncTime
  );
}

function buildSkippedExistingRow_(existingData, contactId, name, company, primaryPhone, email, lastConversationActivityAt, syncTime) {
  var row = existingData.slice();

  while (row.length < HEADERS.length) {
    row.push('');
  }

  row[COL['Contact ID']] = contactId;
  row[COL['Name']] = name;
  row[COL['Company']] = company;
  row[COL['Primary Phone']] = primaryPhone;
  row[COL['Email']] = email;
  row[COL['Last Conversation Activity At']] = lastConversationActivityAt;
  row[COL['Days Since Last Contact']] = calculateDaysSince_(row[COL['Last Any Contact At']]);
  row[COL['Last Sync At']] = syncTime.toISOString();

  return row;
}

function writeBasicContactRows_(sheet, contacts) {
  var existingMap = readExistingRows(sheet);
  var rowsToUpdate = [];
  var rowsToAppend = [];

  for (var i = 0; i < contacts.length; i++) {
    var contact = contacts[i];

    if (!contact.id) continue;

    var df = contact.defaultFields || {};
    var name = ((df.firstName || '') + ' ' + (df.lastName || '')).trim();

    if (existingMap[contact.id]) {
      var existing = existingMap[contact.id].data.slice();

      while (existing.length < HEADERS.length) {
        existing.push('');
      }

      existing[COL['Contact ID']] = contact.id;
      existing[COL['Name']] = name;
      existing[COL['Company']] = df.company || '';
      existing[COL['Primary Phone']] = getPrimaryPhone(df.phoneNumbers);
      existing[COL['Email']] = getPrimaryEmail(df.emails);

      rowsToUpdate.push({
        rowIndex: existingMap[contact.id].rowIndex,
        data: existing
      });
    } else {
      var row = new Array(HEADERS.length).fill('');
      row[COL['Contact ID']] = contact.id;
      row[COL['Name']] = name;
      row[COL['Company']] = df.company || '';
      row[COL['Primary Phone']] = getPrimaryPhone(df.phoneNumbers);
      row[COL['Email']] = getPrimaryEmail(df.emails);
      rowsToAppend.push(row);
    }
  }

  for (var u = 0; u < rowsToUpdate.length; u++) {
    sheet.getRange(rowsToUpdate[u].rowIndex, 1, 1, HEADERS.length).setValues([rowsToUpdate[u].data]);
  }

  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, HEADERS.length).setValues(rowsToAppend);
  }

  Logger.log('Basic rows written: ' + rowsToUpdate.length + ' updated, ' + rowsToAppend.length + ' appended');
}

function scheduleContinuation_() {
  clearContinuationTriggers_();

  ScriptApp.newTrigger(CONTINUATION_TRIGGER_FN)
    .timeBased()
    .after(CONFIG.CHUNK_DELAY_SECONDS * 1000)
    .create();
}

function clearContinuationTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();

  for (var t = 0; t < triggers.length; t++) {
    if (triggers[t].getHandlerFunction() === CONTINUATION_TRIGGER_FN) {
      ScriptApp.deleteTrigger(triggers[t]);
    }
  }
}

// =============================================================================
// ONE-OFF CONTACT SYNC
// =============================================================================

function syncOneContactByPhonePrompt() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Sync One Contact',
    'Enter the contact phone number, like +15124121624',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  syncOneContactByPhone_(response.getResponseText());
}

function syncOneContactByPhone_(phone) {
  var normalizedPhone = normalizePhone(phone);
  var sheet = setupSheet();
  var contacts = fetchAllContacts();
  var phoneNumberIds = fetchAllPhoneNumberIds();
  var userIds = fetchAllUserIds();
  var existingMap = readExistingRows(sheet);

  Logger.log('Syncing one contact by phone: ' + normalizedPhone);

  for (var i = 0; i < contacts.length; i++) {
    var contact = contacts[i];
    var df = contact.defaultFields || {};
    var contactPhones = getAllPhones(df.phoneNumbers);

    if (contactPhones.indexOf(normalizedPhone) === -1) continue;

    var name = ((df.firstName || '') + ' ' + (df.lastName || '')).trim();
    var company = df.company || '';
    var primaryPhone = getPrimaryPhone(df.phoneNumbers);
    var email = getPrimaryEmail(df.emails);

    var contactPhoneSet = {};
    for (var p = 0; p < contactPhones.length; p++) {
      contactPhoneSet[contactPhones[p]] = true;
    }

    var conversationActivityMap = buildConversationActivityMap_(phoneNumberIds, contactPhoneSet);
    var lastConversationActivityAt = latestConversationActivityForPhones_(contactPhones, conversationActivityMap);

    var row = buildFreshActivityRow_(
      contact.id,
      name,
      company,
      primaryPhone,
      email,
      contactPhones,
      phoneNumberIds,
      userIds,
      lastConversationActivityAt,
      new Date()
    );

    if (existingMap[contact.id]) {
      sheet.getRange(existingMap[contact.id].rowIndex, 1, 1, HEADERS.length).setValues([row]);
      Logger.log('Updated existing row for ' + name + ' at row ' + existingMap[contact.id].rowIndex);
    } else {
      sheet.appendRow(row);
      Logger.log('Appended new row for ' + name + ' at row ' + sheet.getLastRow());
    }

    return;
  }

  Logger.log('No Quo contact found with phone: ' + normalizedPhone);
}

// =============================================================================
// SHEET
// =============================================================================

function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }

  var existingHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  var needsHeaders = false;

  for (var i = 0; i < HEADERS.length; i++) {
    if (existingHeaders[i] !== HEADERS[i]) {
      needsHeaders = true;
      break;
    }
  }

  if (needsHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }

  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');

  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, COL['Primary Phone'] + 1, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
    sheet.getRange(2, COL['Days Since Last Contact'] + 1, sheet.getMaxRows() - 1, 1).setNumberFormat('0');
  }

  return sheet;
}

function readExistingRows(sheet) {
  var map = {};
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) return map;

  var data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

  for (var i = 0; i < data.length; i++) {
    var contactId = String(data[i][COL['Contact ID']] || '').trim();

    if (contactId) {
      map[contactId] = {
        rowIndex: i + 2,
        data: data[i]
      };
    }
  }

  return map;
}

// =============================================================================
// API
// =============================================================================

function quoApiFetch(endpoint, queryParams, options) {
  options = options || {};

  var url = CONFIG.API_BASE + endpoint;

  if (queryParams) {
    var parts = [];

    for (var key in queryParams) {
      if (queryParams[key] === null || queryParams[key] === undefined) continue;

      var val = queryParams[key];

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

  var fetchOptions = {
    method: 'get',
    headers: {
      Authorization: CONFIG.API_KEY
    },
    muteHttpExceptions: true
  };

  var backoff = CONFIG.INITIAL_BACKOFF_MS;

  for (var attempt = 0; attempt <= CONFIG.MAX_RETRIES; attempt++) {
    try {
      var response = UrlFetchApp.fetch(url, fetchOptions);
      var code = response.getResponseCode();

      if (code === 200) {
        return JSON.parse(response.getContentText());
      }

      if (options.silent403 && code === 403) {
        return null;
      }

      if (code === 429 || code >= 500) {
        Logger.log('API returned ' + code + ' for ' + endpoint + ', retrying in ' + backoff + 'ms');
        Utilities.sleep(backoff);
        backoff *= 2;
        continue;
      }

      Logger.log('API error ' + code + ' for ' + endpoint + ': ' + response.getContentText().substring(0, 500));
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

function fetchAllPages(endpoint, baseParams, pageSize, options) {
  var allItems = [];
  var pageToken = null;

  do {
    var params = {};

    for (var k in baseParams || {}) {
      params[k] = baseParams[k];
    }

    params.maxResults = pageSize;

    if (pageToken) {
      params.pageToken = pageToken;
    }

    var result = quoApiFetch(endpoint, params, options);

    if (!result || !result.data) break;

    allItems = allItems.concat(result.data);
    pageToken = result.nextPageToken || null;
  } while (pageToken);

  return allItems;
}

function fetchAllContacts() {
  return fetchAllPages('/contacts', {}, CONFIG.CONTACTS_PAGE_SIZE);
}

function fetchAllPhoneNumberIds() {
  var result = quoApiFetch('/phone-numbers', {});
  var ids = [];

  if (!result || !result.data) return ids;

  for (var i = 0; i < result.data.length; i++) {
    if (result.data[i].id && ids.indexOf(result.data[i].id) === -1) {
      ids.push(result.data[i].id);
    }
  }

  return ids;
}

function fetchAllUserIds() {
  var users = fetchAllPages('/users', {}, CONFIG.USERS_PAGE_SIZE);
  var ids = [];

  for (var i = 0; i < users.length; i++) {
    if (users[i].id && ids.indexOf(users[i].id) === -1) {
      ids.push(users[i].id);
    }
  }

  return ids;
}

// =============================================================================
// ACTIVITY FETCHING
// =============================================================================

function fetchActivityForContactPhones_(contactPhones, allPhoneNumberIds, userIds) {
  var calls = [];
  var messages = [];
  var seenCallIds = {};
  var seenMessageIds = {};

  if (!contactPhones || contactPhones.length === 0) {
    return {
      calls: calls,
      messages: messages
    };
  }

  userIds = userIds || [];

  for (var p = 0; p < contactPhones.length; p++) {
    var phone = contactPhones[p];

    for (var n = 0; n < allPhoneNumberIds.length; n++) {
      var pnId = allPhoneNumberIds[n];

      var baseParams = {
        phoneNumberId: pnId,
        participants: [phone]
      };

      var callParamSets = [baseParams];

      for (var u = 0; u < userIds.length; u++) {
        callParamSets.push({
          phoneNumberId: pnId,
          participants: [phone],
          userId: userIds[u]
        });
      }

      for (var cp = 0; cp < callParamSets.length; cp++) {
        var phoneCalls = fetchAllPages('/calls', callParamSets[cp], CONFIG.CALLS_PAGE_SIZE, {
          silent403: true
        });

        for (var c = 0; c < phoneCalls.length; c++) {
          var call = phoneCalls[c];
          var callKey = call.id || JSON.stringify(call);

          if (!seenCallIds[callKey]) {
            seenCallIds[callKey] = true;
            calls.push(call);
          }
        }
      }

      var messageParamSets = [baseParams];

      for (var mu = 0; mu < userIds.length; mu++) {
        messageParamSets.push({
          phoneNumberId: pnId,
          participants: [phone],
          userId: userIds[mu]
        });
      }

      for (var mp = 0; mp < messageParamSets.length; mp++) {
        var phoneMessages = fetchAllPages('/messages', messageParamSets[mp], CONFIG.MESSAGES_PAGE_SIZE, {
          silent403: true
        });

        for (var m = 0; m < phoneMessages.length; m++) {
          var msg = phoneMessages[m];
          var msgKey = msg.id || JSON.stringify(msg);

          if (!seenMessageIds[msgKey]) {
            seenMessageIds[msgKey] = true;
            messages.push(msg);
          }
        }
      }
    }
  }

  return {
    calls: calls,
    messages: messages
  };
}

function buildContactPhoneSet_(contacts) {
  var set = {};

  for (var i = 0; i < contacts.length; i++) {
    var df = contacts[i].defaultFields || {};
    var phones = getAllPhones(df.phoneNumbers);

    for (var p = 0; p < phones.length; p++) {
      set[phones[p]] = true;
    }
  }

  return set;
}

function getConversationActivityMapForRun_(phoneNumberIds, contactPhoneSet, forceRebuild) {
  if (!forceRebuild) {
    var cached = readConversationActivityCache_();

    if (Object.keys(cached).length > 0) {
      Logger.log('Loaded conversation activity cache: ' + Object.keys(cached).length + ' phones');
      return cached;
    }
  }

  Logger.log('Building conversation activity cache...');
  var map = buildConversationActivityMap_(phoneNumberIds, contactPhoneSet);
  writeConversationActivityCache_(map);
  Logger.log('Conversation activity cache built: ' + Object.keys(map).length + ' phones');

  return map;
}

function buildConversationActivityMap_(phoneNumberIds, contactPhoneSet) {
  var map = {};

  for (var p = 0; p < phoneNumberIds.length; p++) {
    var conversations = fetchAllPages('/conversations', {
      phoneNumbers: [phoneNumberIds[p]]
    }, CONFIG.CONVERSATIONS_PAGE_SIZE);

    for (var c = 0; c < conversations.length; c++) {
      var conv = conversations[c];
      var participants = conv.participants || [];
      var ts = conv.lastActivityAt || conv.updatedAt || '';

      if (!ts) continue;

      for (var x = 0; x < participants.length; x++) {
        var normalized = normalizePhone(participants[x]);

        if (!normalized) continue;
        if (contactPhoneSet && !contactPhoneSet[normalized]) continue;

        if (!map[normalized] || ts > map[normalized]) {
          map[normalized] = ts;
        }
      }
    }
  }

  return map;
}

function readConversationActivityCache_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cacheSheet = ss.getSheetByName(CONFIG.CONVERSATION_CACHE_SHEET_NAME);
  var map = {};

  if (!cacheSheet || cacheSheet.getLastRow() < 2) return map;

  var data = cacheSheet.getRange(2, 1, cacheSheet.getLastRow() - 1, 2).getValues();

  for (var i = 0; i < data.length; i++) {
    var phone = String(data[i][0] || '').trim();
    var ts = valueToComparableString_(data[i][1]);

    if (phone && ts) {
      map[phone] = ts;
    }
  }

  return map;
}

function writeConversationActivityCache_(map) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cacheSheet = ss.getSheetByName(CONFIG.CONVERSATION_CACHE_SHEET_NAME);

  if (!cacheSheet) {
    cacheSheet = ss.insertSheet(CONFIG.CONVERSATION_CACHE_SHEET_NAME);
  }

  cacheSheet.clear();
  cacheSheet.getRange(1, 1, 1, 2).setValues([['Phone', 'Last Conversation Activity At']]);

  var rows = [];

  for (var phone in map) {
    rows.push([phone, map[phone]]);
  }

  if (rows.length > 0) {
    cacheSheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }

  try {
    cacheSheet.hideSheet();
  } catch (e) {
    Logger.log('Could not hide cache sheet: ' + e.message);
  }
}

function latestConversationActivityForPhones_(contactPhones, conversationActivityMap) {
  var latest = '';

  for (var i = 0; i < contactPhones.length; i++) {
    var phone = normalizePhone(contactPhones[i]);
    var ts = conversationActivityMap[phone] || '';

    if (ts && ts > latest) {
      latest = ts;
    }
  }

  return latest;
}

function computeTextActivityFromMessages(messages) {
  var lastInbound = '';
  var lastOutbound = '';
  var lastInboundText = '';

  for (var m = 0; m < messages.length; m++) {
    var msg = messages[m];
    var ts = getMessageTimestamp_(msg);
    var dir = normalizeDirection_(msg.direction);

    if (!ts) continue;

    if (dir === 'inbound' && ts > lastInbound) {
      lastInbound = ts;
      lastInboundText = msg.text || '';
    }

    if (dir === 'outbound' && ts > lastOutbound) {
      lastOutbound = ts;
    }
  }

  return {
    lastInbound: lastInbound,
    lastOutbound: lastOutbound,
    lastInboundText: lastInboundText
  };
}

function fetchVoicemailsForCalls(allCalls) {
  var cache = {};

  for (var i = 0; i < allCalls.length; i++) {
    var call = allCalls[i];
    var status = String(call.status || '').toLowerCase();

    if (status !== 'missed' && status !== 'completed') continue;
    if (!call.id) continue;

    var vmResult = quoApiFetch('/call-voicemails/' + call.id, {}, {
      silent403: true
    });

    if (vmResult && vmResult.data && vmResult.data.transcript) {
      cache[call.id] = {
        transcript: vmResult.data.transcript || '',
        duration: vmResult.data.duration || 0,
        direction: call.direction,
        createdAt: getCallTimestamp_(call)
      };
    }
  }

  return cache;
}

// =============================================================================
// CLASSIFICATION
// =============================================================================

function emptyTextActivity_() {
  return {
    lastInbound: '',
    lastOutbound: '',
    lastInboundText: ''
  };
}

function emptyClassified_() {
  return {
    lastCompletedInbound: '',
    lastCompletedOutbound: '',
    lastMissedInbound: '',
    lastMissedOutbound: '',
    lastSonaInbound: '',
    lastSonaOutbound: '',
    lastVoicemailInbound: '',
    lastVoicemailOutbound: '',
    lastInboundVoicemailTranscript: '',
    lastInboundVoicemailAt: '',
    lastAnyInboundCall: '',
    lastAnyOutboundCall: '',
    lastAnyCallAt: '',
    lastAnyCallDirection: '',
    lastAnyCallStatus: '',
    lastAnyCallIsSona: false,
    lastAnyCallHasVoicemail: false
  };
}

function classifyCalls(calls, voicemailCache) {
  var result = emptyClassified_();

  for (var i = 0; i < calls.length; i++) {
    var call = calls[i];
    var ts = getCallTimestamp_(call);
    var dir = normalizeDirection_(call.direction);
    var status = String(call.status || '').toLowerCase();
    var isSona = !!call.aiHandled;
    var hasVoicemail = !!voicemailCache[call.id];

    if (!ts) continue;

    if (dir === 'inbound' && ts > result.lastAnyInboundCall) {
      result.lastAnyInboundCall = ts;
    }

    if (dir === 'outbound' && ts > result.lastAnyOutboundCall) {
      result.lastAnyOutboundCall = ts;
    }

    if (ts > result.lastAnyCallAt) {
      result.lastAnyCallAt = ts;
      result.lastAnyCallDirection = dir;
      result.lastAnyCallStatus = status || 'unknown';
      result.lastAnyCallIsSona = isSona;
      result.lastAnyCallHasVoicemail = hasVoicemail;
    }

    if (hasVoicemail) {
      var vmData = voicemailCache[call.id];

      if (dir === 'inbound' && ts > result.lastVoicemailInbound) {
        result.lastVoicemailInbound = ts;

        if (ts > result.lastInboundVoicemailAt) {
          result.lastInboundVoicemailAt = ts;
          result.lastInboundVoicemailTranscript = vmData.transcript || '';
        }
      }

      if (dir === 'outbound' && ts > result.lastVoicemailOutbound) {
        result.lastVoicemailOutbound = ts;
      }
    }

    if (isSona) {
      if (dir === 'inbound' && ts > result.lastSonaInbound) {
        result.lastSonaInbound = ts;
      }

      if (dir === 'outbound' && ts > result.lastSonaOutbound) {
        result.lastSonaOutbound = ts;
      }

      continue;
    }

    if (status === 'completed' || status === 'answered') {
      if (dir === 'inbound' && ts > result.lastCompletedInbound) {
        result.lastCompletedInbound = ts;
      }

      if (dir === 'outbound' && ts > result.lastCompletedOutbound) {
        result.lastCompletedOutbound = ts;
      }
    }

    if (status === 'missed' || status === 'no-answer' || status === 'no_answer') {
      if (dir === 'inbound' && ts > result.lastMissedInbound) {
        result.lastMissedInbound = ts;
      }

      if (dir === 'outbound' && ts > result.lastMissedOutbound) {
        result.lastMissedOutbound = ts;
      }
    }
  }

  return result;
}

// =============================================================================
// ROW BUILDING
// =============================================================================

function buildContactRow(contactId, name, company, primaryPhone, email, classified, textActivity, lastConversationActivityAt, syncTime) {
  var lastCompletedHumanCall = latestOf([
    classified.lastCompletedInbound,
    classified.lastCompletedOutbound
  ]);

  var lastText = latestOf([
    textActivity.lastInbound,
    textActivity.lastOutbound
  ]);

  var lastMissedCall = latestOf([
    classified.lastMissedInbound,
    classified.lastMissedOutbound
  ]);

  var lastSonaCall = latestOf([
    classified.lastSonaInbound,
    classified.lastSonaOutbound
  ]);

  var lastVoicemail = latestOf([
    classified.lastVoicemailInbound,
    classified.lastVoicemailOutbound
  ]);

  var lastOutboundCall = latestOf([
    classified.lastAnyOutboundCall,
    classified.lastCompletedOutbound,
    classified.lastMissedOutbound,
    classified.lastSonaOutbound,
    classified.lastVoicemailOutbound
  ]);

  var lastAnyCall = latestOf([
    classified.lastAnyCallAt,
    lastCompletedHumanCall,
    lastMissedCall,
    lastSonaCall,
    lastVoicemail
  ]);

  var lastAnyContact = latestOf([
    lastAnyCall,
    lastText,
    lastConversationActivityAt
  ]);

  var languageSample = '';

  if (classified.lastInboundVoicemailAt > textActivity.lastInbound) {
    languageSample = classified.lastInboundVoicemailTranscript;
  } else {
    languageSample = textActivity.lastInboundText || classified.lastInboundVoicemailTranscript;
  }

  var lastLanguageUsed = detectLanguage(languageSample);
  var lastContactInfo = determineLastContactType(classified, textActivity);

  if (lastConversationActivityAt && lastConversationActivityAt > latestOf([lastAnyCall, lastText])) {
    lastContactInfo = {
      type: 'conversation_activity',
      direction: ''
    };
  }

  var row = new Array(HEADERS.length).fill('');

  row[COL['Contact ID']] = contactId;
  row[COL['Name']] = name;
  row[COL['Company']] = company;
  row[COL['Primary Phone']] = primaryPhone;
  row[COL['Email']] = email;
  row[COL['Last Completed Inbound Call At']] = classified.lastCompletedInbound;
  row[COL['Last Completed Outbound Call At']] = classified.lastCompletedOutbound;
  row[COL['Last Completed Human Call At']] = lastCompletedHumanCall;
  row[COL['Last Inbound Text At']] = textActivity.lastInbound;
  row[COL['Last Outbound Text At']] = textActivity.lastOutbound;
  row[COL['Last Text At']] = lastText;
  row[COL['Last Missed Call At']] = lastMissedCall;
  row[COL['Last Missed Inbound Call At']] = classified.lastMissedInbound;
  row[COL['Last Missed Outbound Call At']] = classified.lastMissedOutbound;
  row[COL['Last Sona Call At']] = lastSonaCall;
  row[COL['Last Sona Inbound Call At']] = classified.lastSonaInbound;
  row[COL['Last Sona Outbound Call At']] = classified.lastSonaOutbound;
  row[COL['Last Voicemail At']] = lastVoicemail;
  row[COL['Last Inbound Voicemail At']] = classified.lastVoicemailInbound;
  row[COL['Last Outbound Voicemail At']] = classified.lastVoicemailOutbound;
  row[COL['Last Language Used']] = lastLanguageUsed;
  row[COL['Last Outbound Call At']] = lastOutboundCall;
  row[COL['Last Any Call At']] = lastAnyCall;
  row[COL['Last Any Contact At']] = lastAnyContact;
  row[COL['Last Contact Type']] = lastContactInfo.type;
  row[COL['Last Contact Direction']] = lastContactInfo.direction;
  row[COL['Last Conversation Activity At']] = lastConversationActivityAt;
  row[COL['Days Since Last Contact']] = calculateDaysSince_(lastAnyContact);
  row[COL['Last Sync At']] = syncTime.toISOString();

  return row;
}

function determineLastContactType(classified, textActivity) {
  var genericCallType = '';

  if (classified.lastAnyCallAt) {
    if (classified.lastAnyCallIsSona) {
      genericCallType = 'sona_' + classified.lastAnyCallDirection + '_call';
    } else if (classified.lastAnyCallHasVoicemail) {
      genericCallType = classified.lastAnyCallDirection + '_voicemail';
    } else {
      genericCallType = classified.lastAnyCallStatus + '_' + classified.lastAnyCallDirection + '_call';
    }
  }

  var events = [
    { ts: classified.lastCompletedInbound, type: 'completed_inbound_call', direction: 'inbound' },
    { ts: classified.lastCompletedOutbound, type: 'completed_outbound_call', direction: 'outbound' },
    { ts: classified.lastMissedInbound, type: 'missed_inbound_call', direction: 'inbound' },
    { ts: classified.lastMissedOutbound, type: 'missed_outbound_call', direction: 'outbound' },
    { ts: classified.lastSonaInbound, type: 'sona_inbound_call', direction: 'inbound' },
    { ts: classified.lastSonaOutbound, type: 'sona_outbound_call', direction: 'outbound' },
    { ts: classified.lastVoicemailInbound, type: 'inbound_voicemail', direction: 'inbound' },
    { ts: classified.lastVoicemailOutbound, type: 'outbound_voicemail', direction: 'outbound' },
    { ts: textActivity.lastInbound, type: 'inbound_text', direction: 'inbound' },
    { ts: textActivity.lastOutbound, type: 'outbound_text', direction: 'outbound' },
    { ts: classified.lastAnyCallAt, type: genericCallType, direction: classified.lastAnyCallDirection }
  ];

  var best = {
    ts: '',
    type: '',
    direction: ''
  };

  for (var i = 0; i < events.length; i++) {
    if (events[i].ts && events[i].ts > best.ts) {
      best = events[i];
    }
  }

  return {
    type: best.type,
    direction: best.direction
  };
}

function latestOf(timestamps) {
  var latest = '';

  for (var i = 0; i < timestamps.length; i++) {
    if (timestamps[i] && timestamps[i] > latest) {
      latest = timestamps[i];
    }
  }

  return latest;
}

// =============================================================================
// UTILITIES
// =============================================================================

function getCallTimestamp_(call) {
  return call.createdAt || call.answeredAt || call.completedAt || call.updatedAt || '';
}

function getMessageTimestamp_(msg) {
  return msg.createdAt || msg.updatedAt || '';
}

function normalizeDirection_(direction) {
  var dir = String(direction || '').toLowerCase();

  if (dir === 'incoming') return 'inbound';
  if (dir === 'outgoing') return 'outbound';

  return dir;
}

function normalizePhone(phone) {
  if (!phone) return '';

  var raw = String(phone).trim();
  var hasPlus = raw.charAt(0) === '+';
  var digits = raw.replace(/[^\d]/g, '');

  if (!digits) return '';

  if (hasPlus) return '+' + digits;
  if (digits.length === 10) return '+1' + digits;
  if (digits.length === 11 && digits.charAt(0) === '1') return '+' + digits;

  return '+' + digits;
}

function getAllPhones(phoneNumbers) {
  var phones = [];

  if (!phoneNumbers || !Array.isArray(phoneNumbers)) return phones;

  for (var i = 0; i < phoneNumbers.length; i++) {
    var normalized = normalizePhone(phoneNumbers[i].value);

    if (normalized && phones.indexOf(normalized) === -1) {
      phones.push(normalized);
    }
  }

  return phones;
}

function getPrimaryPhone(phoneNumbers) {
  if (!phoneNumbers || !Array.isArray(phoneNumbers) || phoneNumbers.length === 0) return '';
  return phoneNumbers[0].value || '';
}

function getPrimaryEmail(emails) {
  if (!emails || !Array.isArray(emails) || emails.length === 0) return '';
  return emails[0].value || '';
}

function calculateDaysSince_(timestamp) {
  var ts = valueToComparableString_(timestamp);

  if (!ts) return '';

  return Math.floor((new Date() - new Date(ts)) / (1000 * 60 * 60 * 24));
}

function valueToComparableString_(value) {
  if (!value) return '';

  if (Object.prototype.toString.call(value) === '[object Date]') {
    return value.toISOString();
  }

  return String(value).trim();
}

// =============================================================================
// LANGUAGE DETECTION
// =============================================================================

function detectLanguage(text) {
  if (!text || typeof text !== 'string') return '';

  return detectLanguageHeuristic(text);
}

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
  var es = 0;
  var en = 0;

  for (var i = 0; i < SPANISH.length; i++) {
    if (padded.indexOf(SPANISH[i]) !== -1) es++;
  }

  for (var j = 0; j < ENGLISH.length; j++) {
    if (padded.indexOf(ENGLISH[j]) !== -1) en++;
  }

  if (es === 0 && en === 0) return '';
  if (es > en) return 'Spanish';
  if (en > es) return 'English';

  return '';
}

// =============================================================================
// DEBUG HELPERS
// =============================================================================

function debugContactActivityPrompt() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Debug Contact Activity', 'Enter the contact phone number, like +15124121624', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) return;

  debugContactActivityByPhone_(response.getResponseText());
}

function debugContactActivityByPhone_(phone) {
  var normalizedPhone = normalizePhone(phone);
  var phoneNumberIds = fetchAllPhoneNumberIds();
  var userIds = fetchAllUserIds();

  Logger.log('Checking phone: ' + normalizedPhone);
  Logger.log('Quo inbox phone number IDs: ' + JSON.stringify(phoneNumberIds));
  Logger.log('Quo user IDs: ' + JSON.stringify(userIds));

  var activity = fetchActivityForContactPhones_([normalizedPhone], phoneNumberIds, userIds);

  Logger.log('Total deduped calls found: ' + activity.calls.length);
  Logger.log('Total deduped messages found: ' + activity.messages.length);

  activity.calls.sort(function(a, b) {
    return String(getCallTimestamp_(b)).localeCompare(String(getCallTimestamp_(a)));
  });

  for (var i = 0; i < activity.calls.length; i++) {
    var c = activity.calls[i];

    Logger.log('CALL ' + JSON.stringify({
      id: c.id,
      createdAt: c.createdAt,
      answeredAt: c.answeredAt,
      completedAt: c.completedAt,
      updatedAt: c.updatedAt,
      direction: c.direction,
      status: c.status,
      duration: c.duration,
      phoneNumberId: c.phoneNumberId,
      userId: c.userId,
      answeredBy: c.answeredBy,
      participants: c.participants,
      aiHandled: c.aiHandled
    }));
  }

  activity.messages.sort(function(a, b) {
    return String(getMessageTimestamp_(b)).localeCompare(String(getMessageTimestamp_(a)));
  });

  for (var m = 0; m < activity.messages.length; m++) {
    var msg = activity.messages[m];

    Logger.log('MESSAGE ' + JSON.stringify({
      id: msg.id,
      createdAt: msg.createdAt,
      updatedAt: msg.updatedAt,
      direction: msg.direction,
      userId: msg.userId,
      phoneNumberId: msg.phoneNumberId,
      text: msg.text
    }));
  }
}

function debugContactRecordPrompt() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Debug Contact Record', 'Enter the contact phone number, like +15124121624', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) return;

  debugContactRecordByPhone_(response.getResponseText());
}

function debugContactRecordByPhone_(phone) {
  var normalizedPhone = normalizePhone(phone);
  var contacts = fetchAllContacts();
  var found = false;

  Logger.log('Looking for contact record with phone: ' + normalizedPhone);
  Logger.log('Total contacts fetched: ' + contacts.length);

  for (var i = 0; i < contacts.length; i++) {
    var contact = contacts[i];
    var df = contact.defaultFields || {};
    var phones = getAllPhones(df.phoneNumbers);
    var name = ((df.firstName || '') + ' ' + (df.lastName || '')).trim();

    if (phones.indexOf(normalizedPhone) !== -1) {
      found = true;

      Logger.log('CONTACT FOUND ' + JSON.stringify({
        id: contact.id,
        name: name,
        company: df.company || '',
        primaryPhone: getPrimaryPhone(df.phoneNumbers),
        rawPhoneNumbers: df.phoneNumbers,
        normalizedPhones: phones,
        emails: df.emails
      }));
    }
  }

  if (!found) {
    Logger.log('NO CONTACT RECORD FOUND for ' + normalizedPhone);
  }
}

function sortByRecentTime() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange();

  range.sort({
    column: COL['Last Any Contact At'] + 1,
    ascending: false
  });
}

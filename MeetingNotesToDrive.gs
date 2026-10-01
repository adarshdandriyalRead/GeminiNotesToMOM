/**
 * Meeting notes from Gmail → Google Drive.
 *
 * Run while signed in as adandriyal@salesforce.com.
 *
 * Editor functions:
 *   setupMeetingNotesToDrive  — authorize, create the subjects sheet, hourly trigger, run once
 *   listConfiguredSubjects    — show the active allow-list
 *   listRecentNoteSubjects    — discover exact subjects from recent mail
 *   dryRunMeetingNotes        — log matches in the last hour, write nothing
 *   testLast12Hours           — log matches in the last 12 hours, write nothing
 *   fileLast12Hours           — copy matching mail from the last 12 hours into Drive
 *   diagnoseLast12Hours       — log why mail was or was not filed
 *   testSlackPost             — send a short test MOM message to the Slack group
 *   listGeminiNotesLast12Hours — dump From + Subject of Gemini notes (debug)
 *   runMeetingNotesToDrive    — process matching mail now (last hour)
 *   uninstallMeetingNotesTrigger
 */

var RUN_ = { needles: [], lookbackOverrideMinutes: null, locationOverride: null };

function setupMeetingNotesToDrive() {
  assertConfig_();
  var folder = DriveApp.getFolderById(CONFIG.folderId);
  getOrCreateLabel_(CONFIG.processedLabel);
  var ss = getOrCreateSubjectsSheet_(folder);
  installScheduleTrigger_();
  var summary = runMeetingNotesToDrive();
  Logger.log('Setup complete.');
  Logger.log('Add more subjects here: ' + ss.getUrl());
  Logger.log('Schedule: ' + describeSchedule_() + '. ' + summary);
  return summary;
}

function dryRunMeetingNotes() {
  return runWithLookback_(null, true);
}

/**
 * Test helper: search the last 12 hours (CONFIG.testLookbackMinutes).
 * Logs matches only — does not write to Drive, label, or archive.
 */
function testLast12Hours() {
  var minutes = Number(CONFIG.testLookbackMinutes) || 720;
  Logger.log('TEST: scanning the last ' + minutes + ' minutes. Dry run, nothing is written.');
  return runWithLookback_(minutes, true);
}

/**
 * One-time catch-up: file matching notes from the last 12 hours into Drive.
 * Hourly trigger stays at 1 hour.
 */
function fileLast12Hours() {
  var minutes = Number(CONFIG.testLookbackMinutes) || 720;
  Logger.log('CATCH-UP: filing matching mail from the last ' + minutes + ' minutes into Drive.');
  Logger.log('Catch-up searches all mail except Spam/Trash (Inbox-only applies to the hourly job).');
  return runWithLookback_(minutes, false, '-in:spam -in:trash');
}

/**
 * Debug: show spreadsheet titles, Gemini mail from the last 12 hours, and skip reasons.
 */
function diagnoseLast12Hours() {
  var minutes = Number(CONFIG.testLookbackMinutes) || 720;
  var cutoff = new Date(Date.now() - minutes * 60 * 1000);
  var folder = DriveApp.getFolderById(CONFIG.folderId);
  var ss = getOrCreateSubjectsSheet_(folder);
  var needles = loadEnabledSubjects_(ss);
  Logger.log('Drive folder: https://drive.google.com/drive/folders/' + CONFIG.folderId);
  Logger.log('Spreadsheet: ' + ss.getUrl());
  Logger.log('Enabled subjects (' + needles.length + '): ' + (needles.join(' | ') || '(none — add rows in the sheet)'));
  Logger.log('Allowed senders: ' + getAllowedSenders_().join(' | '));

  var days = Math.max(1, Math.ceil(minutes / (60 * 24)));
  var query = 'newer_than:' + days + 'd -in:spam -in:trash ' + buildFromClause_();
  Logger.log('Diagnose query: ' + query);
  var threads = GmailApp.search(query, 0, 50);
  Logger.log('Threads found: ' + threads.length);
  if (threads.length === 0) {
    Logger.log(
      'Nothing from ' +
        getAllowedSenders_().join(', ') +
        ' in the last ' +
        days +
        ' day(s). Check From on the email, or add that address to allowedSenders.'
    );
    return;
  }

  var previousNeedles = RUN_.needles;
  RUN_.needles = needles;
  threads.forEach(function (thread) {
    thread.getMessages().forEach(function (message) {
      if (message.getDate().getTime() < cutoff.getTime()) {
        return;
      }
      Logger.log(
        message.getDate() +
          ' | inbox=' +
          thread.isInInbox() +
          ' | ' +
          skipReason_(thread, message, getOrCreateLabel_(CONFIG.processedLabel)) +
          ' | From: ' +
          message.getFrom() +
          ' | Subject: ' +
          message.getSubject()
      );
    });
  });
  RUN_.needles = previousNeedles;
}

/**
 * Debug: list every recent Gemini-notes email so we can see the real From and Subject.
 */
function listGeminiNotesLast12Hours() {
  var minutes = Number(CONFIG.testLookbackMinutes) || 720;
  var cutoff = new Date(Date.now() - minutes * 60 * 1000);
  var days = Math.max(1, Math.ceil(minutes / (60 * 24)));
  var queries = [
    buildLocationClause_() + ' ' + buildFromClause_() + ' newer_than:' + days + 'd'
  ];
  var seen = {};
  queries.forEach(function (query) {
    Logger.log('Debug query: ' + query);
    GmailApp.search(query, 0, 50).forEach(function (thread) {
      thread.getMessages().forEach(function (message) {
        if (message.getDate().getTime() < cutoff.getTime()) {
          return;
        }
        var key = message.getId();
        if (seen[key]) {
          return;
        }
        seen[key] = true;
        Logger.log(
          message.getDate() +
            ' | From: ' +
            message.getFrom() +
            ' | Subject: ' +
            message.getSubject()
        );
      });
    });
  });
  if (Object.keys(seen).length === 0) {
    Logger.log('No messages in the last ' + minutes + ' minutes for those queries.');
  }
}

function runWithLookback_(lookbackMinutes, dryRun, locationOverride) {
  var previousDry = CONFIG.dryRun;
  var previousOverride = RUN_.lookbackOverrideMinutes;
  var previousLocation = RUN_.locationOverride;
  CONFIG.dryRun = dryRun;
  RUN_.lookbackOverrideMinutes = lookbackMinutes;
  RUN_.locationOverride = locationOverride || null;
  try {
    return runMeetingNotesToDrive();
  } finally {
    CONFIG.dryRun = previousDry;
    RUN_.lookbackOverrideMinutes = previousOverride;
    RUN_.locationOverride = previousLocation;
  }
}

function runMeetingNotesToDrive() {
  assertConfig_();
  var folder = DriveApp.getFolderById(CONFIG.folderId);
  var ss = getOrCreateSubjectsSheet_(folder);
  RUN_.needles = loadEnabledSubjects_(ss);
  Logger.log('Active subjects: ' + (RUN_.needles.join(' | ') || '(none)'));
  Logger.log('Allowed senders: ' + getAllowedSenders_().join(' | '));
  Logger.log('Lookback: ' + describeLookback_());

  if (RUN_.needles.length === 0) {
    throw new Error(
      'No enabled subjects. Add a row in "' +
        CONFIG.subjectsSpreadsheetName +
        '" (' +
        ss.getUrl() +
        ').'
    );
  }

  var label = getOrCreateLabel_(CONFIG.processedLabel);
  var query = buildSearchQuery_(label);
  Logger.log('Gmail query: ' + query);

  var threads = GmailApp.search(query, 0, CONFIG.maxThreadsPerRun);
  var filed = 0;
  var skipped = 0;
  var errors = 0;

  threads.forEach(function (thread) {
    var messages = thread.getMessages();
    for (var i = 0; i < messages.length; i++) {
      var message = messages[i];
      var reason = skipReason_(thread, message, label);
      if (reason !== 'ok') {
        skipped++;
        Logger.log('SKIP ' + reason + ' | ' + message.getDate() + ' | ' + message.getSubject());
        continue;
      }
      try {
        var result = fileMessage_(message, folder);
        if (result.wroteAnything || CONFIG.dryRun) {
          filed++;
          Logger.log(
            (CONFIG.dryRun ? '[dry-run] ' : 'FILED ') +
              message.getDate() +
              ' | ' +
              message.getSubject() +
              ' | ' +
              result.detail +
              ' | https://drive.google.com/drive/folders/' +
              CONFIG.folderId
          );
          if (!CONFIG.dryRun) {
            postMomToSlack_(message, result);
            thread.addLabel(label);
            if (CONFIG.archiveAfterSave) {
              thread.moveToArchive();
            }
          }
        } else {
          skipped++;
        }
      } catch (err) {
        errors++;
        Logger.log('ERROR ' + message.getSubject() + ': ' + err);
      }
    }
  });

  var summary =
    'threads=' +
    threads.length +
    ' filed=' +
    filed +
    ' skipped=' +
    skipped +
    ' errors=' +
    errors;
  Logger.log(summary);
  return summary;
}

function listConfiguredSubjects() {
  var folder = DriveApp.getFolderById(CONFIG.folderId);
  var ss = getOrCreateSubjectsSheet_(folder);
  var enabled = loadEnabledSubjects_(ss);
  Logger.log('Spreadsheet: ' + ss.getUrl());
  Logger.log('Enabled (' + enabled.length + '):');
  enabled.forEach(function (subject) {
    Logger.log('  - ' + subject);
  });
  return ss.getUrl();
}

function listRecentNoteSubjects() {
  var query =
    buildLocationClause_() +
    ' newer_than:21d ' +
    buildFromClause_();
  var threads = GmailApp.search(query, 0, 50);
  Logger.log('Found ' + threads.length + ' threads from allowed senders. Add useful subjects as new rows in the sheet:');
  var seen = {};
  threads.forEach(function (thread) {
    thread.getMessages().forEach(function (message) {
      var subject = message.getSubject() || '';
      if (seen[subject]) {
        return;
      }
      seen[subject] = true;
      Logger.log(formatDate_(message.getDate()) + ' | ' + subject + ' | ' + message.getFrom());
    });
  });
}

function uninstallMeetingNotesTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'runMeetingNotesToDrive') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function installScheduleTrigger_() {
  uninstallMeetingNotesTrigger();
  var builder = ScriptApp.newTrigger('runMeetingNotesToDrive').timeBased();
  if (CONFIG.schedule === 'frequent') {
    builder.everyMinutes(15).create();
  } else if (CONFIG.schedule === 'daily') {
    builder.atHour(Number(CONFIG.dailyHour) || 22).everyDays(1).create();
  } else {
    builder.everyHours(1).create();
  }
}

function describeSchedule_() {
  var window = describeLookback_();
  if (CONFIG.schedule === 'frequent') {
    return 'every 15 minutes, ' + window;
  }
  if (CONFIG.schedule === 'daily') {
    return (
      'daily at ' +
      (CONFIG.dailyHour || 22) +
      ':00 ' +
      Session.getScriptTimeZone() +
      ', ' +
      window
    );
  }
  return 'every hour, ' + window;
}

function describeLookback_() {
  var minutes = getLookbackMinutes_();
  return 'only mail from the last ' + minutes + ' minute' + (minutes === 1 ? '' : 's');
}

function getLookbackMinutes_() {
  if (RUN_.lookbackOverrideMinutes) {
    return Number(RUN_.lookbackOverrideMinutes);
  }
  return (Number(CONFIG.lookbackMinutes) || 60) + (Number(CONFIG.overlapMinutes) || 0);
}

function getLookbackCutoff_() {
  return new Date(Date.now() - getLookbackMinutes_() * 60 * 1000);
}

function isWithinLookback_(message) {
  return message.getDate().getTime() >= getLookbackCutoff_().getTime();
}

function skipReason_(thread, message, label) {
  if (!isWithinLookback_(message)) {
    return 'outside-lookback';
  }
  if (!RUN_.locationOverride && !thread.isInInbox()) {
    return 'not-in-inbox';
  }
  if (!RUN_.locationOverride && label && threadHasLabel_(thread, label)) {
    return 'already-labeled';
  }
  if (!isAllowedSender_(message)) {
    return 'sender-not-allowed:' + message.getFrom();
  }
  if (!shouldFileMessage_(message)) {
    return 'subject-not-in-sheet';
  }
  return 'ok';
}

function threadHasLabel_(thread, label) {
  var labels = thread.getLabels();
  for (var i = 0; i < labels.length; i++) {
    if (labels[i].getName() === label.getName()) {
      return true;
    }
  }
  return false;
}

function assertConfig_() {
  if (!CONFIG.folderId) {
    throw new Error('CONFIG.folderId is required.');
  }
  if (getAllowedSenders_().length === 0) {
    throw new Error('CONFIG.allowedSenders must include at least one address.');
  }
}

function getAllowedSenders_() {
  return (CONFIG.allowedSenders || [])
    .map(function (s) {
      return String(s || '').trim().toLowerCase();
    })
    .filter(function (s) {
      return s.length > 0;
    });
}

function getOrCreateSubjectsSheet_(folder) {
  var props = PropertiesService.getScriptProperties();
  var knownId = String(CONFIG.subjectsSpreadsheetId || '').trim() || props.getProperty('SUBJECTS_SHEET_ID');
  if (knownId) {
    try {
      return SpreadsheetApp.openById(knownId);
    } catch (err) {
      Logger.log('Stored subjects spreadsheet id is invalid; creating or finding a new one. ' + err);
    }
  }

  var existing = folder.getFilesByName(CONFIG.subjectsSpreadsheetName);
  while (existing.hasNext()) {
    var file = existing.next();
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      props.setProperty('SUBJECTS_SHEET_ID', file.getId());
      return SpreadsheetApp.openById(file.getId());
    }
  }

  var ss = SpreadsheetApp.create(CONFIG.subjectsSpreadsheetName);
  DriveApp.getFileById(ss.getId()).moveTo(folder);
  formatSubjectsSheet_(ss.getSheets()[0]);
  props.setProperty('SUBJECTS_SHEET_ID', ss.getId());
  Logger.log('Created subjects spreadsheet: ' + ss.getUrl());
  return ss;
}

function formatSubjectsSheet_(sheet) {
  sheet.setName('Subjects');
  sheet.getRange(1, 1, 1, 3).setValues([['Subject', 'Enabled', 'Notes']]);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 360);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 280);
  sheet.getRange('B2:B100').insertCheckboxes();
}

function ensureSubjectsHeader_(sheet) {
  var lastColumn = Math.max(sheet.getLastColumn(), 3);
  var header = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  if (String(header[0] || '').toLowerCase() === 'subject') {
    return;
  }
  sheet.insertRowBefore(1);
  sheet.getRange(1, 1, 1, 3).setValues([['Subject', 'Enabled', 'Notes']]);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function loadEnabledSubjects_(ss) {
  ensureSubjectsHeader_(ss.getSheets()[0]);
  var needles = [];
  var seen = {};
  loadAllSubjectRows_(ss.getSheets()[0]).forEach(function (row) {
    if (!row.enabled) {
      return;
    }
    var key = row.subject.toLowerCase();
    if (seen[key]) {
      return;
    }
    seen[key] = true;
    needles.push(row.subject);
  });
  return needles;
}

function loadAllSubjectRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  var values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    var subject = String(values[i][0] || '').trim();
    if (!subject) {
      continue;
    }
    rows.push({
      subject: subject,
      enabled: isEnabledCell_(values[i][1])
    });
  }
  return rows;
}

function isEnabledCell_(value) {
  if (value === '' || value === null) {
    return true;
  }
  if (value === true || value === 1) {
    return true;
  }
  var text = String(value).trim().toLowerCase();
  return text !== 'false' && text !== 'no' && text !== '0' && text !== 'n';
}

function getSubjectNeedles_() {
  return RUN_.needles || [];
}

function buildSearchQuery_(label) {
  var newerThanDays = Math.max(1, Math.ceil(getLookbackMinutes_() / (60 * 24)));
  var parts = [buildLocationClause_(), buildFromClause_(), 'newer_than:' + newerThanDays + 'd'];
  if (!RUN_.locationOverride) {
    parts.push('-label:' + label.getName());
  }
  return parts.join(' ');
}

function buildLocationClause_() {
  if (RUN_.locationOverride) {
    return String(RUN_.locationOverride).trim();
  }
  return String(CONFIG.gmailLocation || 'in:inbox').trim() || 'in:inbox';
}

function buildFromClause_() {
  var senders = getAllowedSenders_();
  if (senders.length === 1) {
    return 'from:' + senders[0];
  }
  return (
    '(' +
    senders
      .map(function (sender) {
        return 'from:' + sender;
      })
      .join(' OR ') +
    ')'
  );
}

function escapeQuery_(value) {
  return normalizeSubject_(value).replace(/"/g, '');
}

function shouldFileMessage_(message) {
  if (!isAllowedSender_(message)) {
    return false;
  }
  var subject = message.getSubject() || '';
  var needles = getSubjectNeedles_();
  for (var i = 0; i < needles.length; i++) {
    if (subjectMatchesNeedle_(subject, needles[i])) {
      return true;
    }
  }
  return false;
}

function subjectMatchesNeedle_(subject, needle) {
  var normSubject = normalizeSubject_(subject);
  var normNeedle = normalizeSubject_(needle);
  if (!normNeedle) {
    return false;
  }
  if (containsIgnoreCase_(normSubject, normNeedle)) {
    return true;
  }
  var title = extractNotesTitle_(subject);
  return !!(title && containsIgnoreCase_(title, normNeedle));
}

function extractNotesTitle_(subject) {
  var normalized = normalizeSubject_(subject);
  var quoted = normalized.match(/^notes:\s*"([^"]+)"/i);
  if (quoted) {
    return quoted[1].trim();
  }
  var withDate = normalized.match(
    /^notes:\s*(.+?)\s+[A-Za-z]{3,9} \d{1,2}, \d{4}\s*$/i
  );
  return withDate ? withDate[1].replace(/^"|"$/g, '').trim() : '';
}

function normalizeSubject_(value) {
  return String(value || '')
    .replace(/[\u201C\u201D\u201E\u201F\u00AB\u00BB]/g, '"')
    .replace(/[\u2018\u2019\u201A\u201B]/g, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

function isAllowedSender_(message) {
  var allowed = getAllowedSenders_();
  var fromAddresses = extractEmailAddresses_(message.getFrom());
  for (var i = 0; i < fromAddresses.length; i++) {
    if (allowed.indexOf(fromAddresses[i]) !== -1) {
      return true;
    }
  }
  return false;
}

function extractEmailAddresses_(header) {
  var text = String(header || '').toLowerCase();
  var matches = text.match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/g);
  return matches || [];
}

function fileMessage_(message, folder) {
  var names = [];
  var files = [];
  var wrote = false;

  if (CONFIG.saveLinkedGoogleFiles) {
    var fileIds = extractGoogleFileIds_(message.getBody() + ' ' + message.getPlainBody());
    fileIds.forEach(function (fileId) {
      try {
        var copy = copyDriveFile_(fileId, folder, message);
        if (copy) {
          names.push(copy.name);
          files.push(copy);
          wrote = true;
        }
      } catch (err) {
        Logger.log('Could not copy Drive file ' + fileId + ': ' + err);
      }
    });
  }

  if (CONFIG.saveAttachments) {
    message.getAttachments().forEach(function (attachment) {
      var destName = uniqueName_(folder, prefixName_(message, attachment.getName()));
      if (!CONFIG.dryRun) {
        var saved = folder.createFile(attachment.copyBlob().setName(destName));
        files.push(describeDriveFile_(saved));
      } else {
        files.push({ name: destName, id: '', url: '', mimeType: attachment.getContentType() });
      }
      names.push(destName);
      wrote = true;
    });
  }

  if (CONFIG.saveEmailHtmlIfNoDoc && !wrote) {
    var htmlName = uniqueName_(folder, prefixName_(message, 'email.html'));
    if (!CONFIG.dryRun) {
      var htmlFile = folder.createFile(
        Utilities.newBlob(wrapHtml_(message), 'text/html', htmlName)
      );
      files.push(describeDriveFile_(htmlFile));
    } else {
      files.push({ name: htmlName, id: '', url: '', mimeType: 'text/html' });
    }
    names.push(htmlName);
    wrote = true;
  }

  return {
    wroteAnything: wrote,
    detail: names.length ? names.join('; ') : 'no files',
    files: files
  };
}

function extractGoogleFileIds_(text) {
  var patterns = [
    /\/document\/d\/([a-zA-Z0-9_-]+)/g,
    /\/presentation\/d\/([a-zA-Z0-9_-]+)/g,
    /\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/g,
    /\/file\/d\/([a-zA-Z0-9_-]+)/g,
    /\/open\?id=([a-zA-Z0-9_-]+)/g
  ];
  var seen = {};
  var ids = [];
  patterns.forEach(function (re) {
    var match;
    while ((match = re.exec(text)) !== null) {
      var id = match[1];
      if (id && id !== CONFIG.folderId && !seen[id]) {
        seen[id] = true;
        ids.push(id);
      }
    }
  });
  return ids;
}

function copyDriveFile_(fileId, folder, message) {
  var source = DriveApp.getFileById(fileId);
  var destName = uniqueName_(folder, prefixName_(message, source.getName()));
  if (CONFIG.dryRun) {
    return {
      name: destName,
      id: '',
      url: '',
      mimeType: source.getMimeType()
    };
  }
  return describeDriveFile_(source.makeCopy(destName, folder));
}

function describeDriveFile_(file) {
  return {
    name: file.getName(),
    id: file.getId(),
    url: file.getUrl(),
    mimeType: file.getMimeType()
  };
}

function prefixName_(message, originalName) {
  var dateStamp = formatDate_(message.getDate());
  var subject = sanitizeFilename_(message.getSubject() || 'meeting-notes');
  var base = originalName || 'notes';
  if (base.indexOf(dateStamp) === 0) {
    return base;
  }
  return dateStamp + ' — ' + subject + ' — ' + base;
}

function uniqueName_(folder, name) {
  if (!folderHasFileNamed_(folder, name)) {
    return name;
  }
  var parts = splitExtension_(name);
  var n = 2;
  var candidate;
  do {
    candidate = parts.stem + ' (' + n + ')' + parts.ext;
    n++;
  } while (folderHasFileNamed_(folder, candidate));
  return candidate;
}

function folderHasFileNamed_(folder, name) {
  return folder.getFilesByName(name).hasNext();
}

function splitExtension_(name) {
  var idx = name.lastIndexOf('.');
  if (idx <= 0 || idx === name.length - 1) {
    return { stem: name, ext: '' };
  }
  return { stem: name.substring(0, idx), ext: name.substring(idx) };
}

function wrapHtml_(message) {
  return (
    '<html><head><meta charset="UTF-8"><title>' +
    escapeHtml_(message.getSubject()) +
    '</title></head><body>' +
    '<p><b>From:</b> ' +
    escapeHtml_(message.getFrom()) +
    '<br><b>Date:</b> ' +
    escapeHtml_(String(message.getDate())) +
    '<br><b>Subject:</b> ' +
    escapeHtml_(message.getSubject()) +
    '</p><hr>' +
    (message.getBody() || '<pre>' + escapeHtml_(message.getPlainBody()) + '</pre>') +
    '</body></html>'
  );
}

function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function containsIgnoreCase_(haystack, needle) {
  return String(haystack).toLowerCase().indexOf(String(needle).toLowerCase()) !== -1;
}

function sanitizeFilename_(value) {
  return String(value)
    .replace(/[\\/:*?"<>|]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 120);
}

function escapeHtml_(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function testSlackPost() {
  var token = getSlackBotToken_();
  if (!token) {
    throw new Error(
      'Set Script property SLACK_BOT_TOKEN (Project Settings → Script properties) to the Slack bot token (xoxb-...).'
    );
  }
  var ok = slackChatPost_(
    '*MOM bot test*\nIf you see this in <#' +
      CONFIG.slackChannelId +
      '>, Slack posting is working.'
  );
  if (!ok) {
    throw new Error(
      'Slack test post failed. Check Executions log, bot token, and that the bot is invited to the group.'
    );
  }
  Logger.log('Slack test post succeeded.');
}

function postMomToSlack_(message, result) {
  if (!CONFIG.slackEnabled) {
    return;
  }
  if (!getSlackBotToken_()) {
    Logger.log(
      'Slack skipped: add Script property SLACK_BOT_TOKEN, then invite the bot to ' +
        CONFIG.slackChannelId +
        '.'
    );
    return;
  }
  var text = buildMomSlackMessage_(message, result);
  if (CONFIG.dryRun) {
    Logger.log('[dry-run] Slack MOM:\n' + text);
    return;
  }
  var ok = slackChatPost_(text);
  Logger.log(ok ? 'Slack MOM posted.' : 'Slack MOM post failed. Drive file was still saved.');
}

function buildMomSlackMessage_(message, result) {
  var title =
    extractNotesTitle_(message.getSubject()) ||
    normalizeSubject_(message.getSubject()) ||
    'Meeting notes';
  var when = Utilities.formatDate(
    message.getDate(),
    Session.getScriptTimeZone(),
    'EEE d MMM yyyy, h:mm a'
  );
  var files = (result && result.files) || [];
  var driveLinks = [];
  for (var i = 0; i < files.length; i++) {
    if (files[i].url) {
      driveLinks.push('<' + files[i].url + '|' + escapeSlack_(files[i].name) + '>');
    }
  }
  var minutes = readMinutesText_(message, files);
  var maxChars = Number(CONFIG.slackMaxChars) || 12000;
  if (minutes.length > maxChars) {
    minutes = minutes.substring(0, maxChars) + '\n\n_[Truncated. Full MOM is in Drive.]_';
  }

  var lines = [
    '*MOM:* ' + escapeSlack_(title),
    '*When:* ' + escapeSlack_(when)
  ];
  if (driveLinks.length) {
    lines.push('*Drive:* ' + driveLinks.join('  ·  '));
  } else {
    lines.push(
      '*Drive folder:* <https://drive.google.com/drive/folders/' +
        CONFIG.folderId +
        '|Open folder>'
    );
  }
  lines.push('');
  lines.push('*Minutes*');
  lines.push('```');
  lines.push(minutes || '(No minutes text found in the email or linked Doc.)');
  lines.push('```');
  return lines.join('\n');
}

function readMinutesText_(message, files) {
  for (var i = 0; i < (files || []).length; i++) {
    var file = files[i];
    if (file.mimeType === MimeType.GOOGLE_DOCS && file.id) {
      var docText = readGoogleDocText_(file.id);
      if (docText) {
        return docText;
      }
    }
  }
  var plain = String(message.getPlainBody() || '').replace(/\r/g, '');
  plain = plain.replace(/https?:\/\/\S+/g, '');
  plain = plain.replace(/\n{3,}/g, '\n\n').trim();
  return plain;
}

function readGoogleDocText_(fileId) {
  try {
    return String(DocumentApp.openById(fileId).getBody().getText() || '').trim();
  } catch (err) {
    Logger.log('Could not read Google Doc ' + fileId + ': ' + err);
    return '';
  }
}

function getSlackBotToken_() {
  return String(
    PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN') || ''
  ).trim();
}

function slackChatPost_(text) {
  var token = getSlackBotToken_();
  var channel = CONFIG.slackChannelId;
  if (!token || !channel) {
    return false;
  }
  var response = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({
      channel: channel,
      text: text,
      unfurl_links: false,
      unfurl_media: false
    }),
    muteHttpExceptions: true
  });
  var body = {};
  try {
    body = JSON.parse(response.getContentText() || '{}');
  } catch (err) {
    Logger.log('Slack response was not JSON: ' + response.getContentText());
    return false;
  }
  if (!body.ok) {
    Logger.log('Slack API error: ' + (body.error || response.getContentText()));
    if (body.error === 'not_in_channel') {
      Logger.log('Invite the bot to the Slack group, then run testSlackPost.');
    }
    if (body.error === 'channel_not_found') {
      Logger.log('Channel ' + channel + ' was not found, or the bot cannot see it.');
    }
    if (body.error === 'missing_scope') {
      Logger.log('Bot needs chat:write. Reinstall the Slack app after adding that scope.');
    }
    return false;
  }
  return true;
}

function escapeSlack_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}


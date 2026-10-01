/**
 * Mailbox: adandriyal@salesforce.com  (run while signed into that account)
 * Drive folder: https://drive.google.com/drive/folders/1Zmkag1y2q5NHd24opvuH6PMaN0t4F3Co
 *
 * Subjects come only from the spreadsheet "Meeting note subjects" in that folder.
 * Add a row: meeting title (text inside the quotes) | Enabled | Notes
 * Email example:  Notes: “DRO: Sprint Grooming” Sep 9, 2026
 * Sheet row:      DRO: Sprint Grooming
 *
 * Search is Inbox only (not Spam, Trash, or other labels).
 * Sender is locked to gemini-notes@google.com.
 *
 * After Drive save, MOM is posted to Slack channel C0C0KU5DP37.
 * Bot token goes in Apps Script → Project Settings → Script properties:
 *   SLACK_BOT_TOKEN = xoxb-...
 */
var CONFIG = {
  folderId: '1Zmkag1y2q5NHd24opvuH6PMaN0t4F3Co',
  processedLabel: 'MeetingNotesFiledToDrive',
  lookbackMinutes: 60,
  overlapMinutes: 0,
  /** Used only by testLast12Hours / fileLast12Hours. Hourly job uses lookbackMinutes. */
  testLookbackMinutes: 720,
  maxThreadsPerRun: 40,

  subjectsSpreadsheetName: 'Meeting note subjects',
  subjectsSpreadsheetId: '',

  allowedSenders: [
    'gemini-notes@google.com'
  ],

  /** Inbox only. Do not search Spam, Trash, or other labels. */
  gmailLocation: 'in:inbox',

  slackEnabled: true,
  slackChannelId: 'C0C0KU5DP37',
  slackMaxChars: 12000,

  schedule: 'hourly',
  dailyHour: 22,

  saveLinkedGoogleFiles: true,
  saveAttachments: true,
  saveEmailHtmlIfNoDoc: true,
  archiveAfterSave: true,

  dryRun: false
};

# Meeting notes: Gmail → Drive

Reads **Inbox** on **adandriyal@salesforce.com** (not Spam, Trash, or other labels), matches subjects from the spreadsheet, and copies those emails into:

https://drive.google.com/drive/folders/1Zmkag1y2q5NHd24opvuH6PMaN0t4F3Co

A message is filed only if **all** of these are true:

- It is in **Inbox**
- **From** is `gemini-notes@google.com`
- **Subject** matches an enabled row in **Meeting note subjects** (the spreadsheet in that Drive folder)

There is no hardcoded meeting list in the script. Add every title in the sheet.

Example email: `Notes: “DRO: Sprint Grooming” Sep 9, 2026`  
Sheet row: `DRO: Sprint Grooming`

**Hourly job:** every 60 minutes, last **60 minutes** of Inbox only. Already-filed mail is skipped via `MeetingNotesFiledToDrive`. After Drive save, the MOM is posted to Slack group [`C0C0KU5DP37`](https://salesforce.enterprise.slack.com/archives/C0C0KU5DP37).

## Slack (MOM in the group)

Posting uses a Slack bot. The token is **not** stored in the script.

1. Create a Slack app at [api.slack.com/apps](https://api.slack.com/apps) on the Salesforce workspace. IT may need to approve it.
2. **OAuth & Permissions → Bot Token Scopes → `chat:write`**.
3. Install to workspace. Copy **Bot User OAuth Token** (`xoxb-...`).
4. In the group, invite the bot: `/invite @YourBotName`.
5. Apps Script → **Project Settings → Script properties**
   - `SLACK_BOT_TOKEN` = the `xoxb-...` token
6. Re-paste `Code.gs`, `Config.gs`, and `appsscript.json`, then run **`testSlackPost`**.

A real filing then posts:

- meeting title
- time
- Drive link
- minutes (from the Google Doc when present, otherwise the email body)

If Slack fails, the Drive file is still kept. Dry runs do not post.

To use a different group, change `slackChannelId` in `Config.gs`.

## Add subjects (spreadsheet only)

| Subject | Enabled | Notes |
|---|---|---|
| DRO: Sprint Grooming | ✓ | meeting title inside the quotes |
| DRO Daily Scrum Meeting | ✓ | add another row |

Uncheck Enabled to pause a row. Next hourly run uses the sheet as-is.

## What it saves

1. Linked Google Docs / Sheets / Slides from the email.
2. Attachments.
3. An `.html` copy of the email only if there was no linked file and no attachment.
4. Label `MeetingNotesFiledToDrive`, then archive (leaves Inbox).

## One-time setup

1. Open [script.google.com](https://script.google.com) signed in as `adandriyal@salesforce.com`.
2. Paste `MeetingNotesToDrive.gs` into `Code.gs` and `Config.gs` into a file named `Config`.
3. Run `setupMeetingNotesToDrive` and approve Gmail, Drive, and Sheets.
4. Open **Meeting note subjects** in the Drive folder and add your titles.
5. Catch-up: `listGeminiNotesLast12Hours` then `fileLast12Hours`.

To stop: `uninstallMeetingNotesTrigger`.

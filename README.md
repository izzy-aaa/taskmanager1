# Task Manager

A personal, Asana-style task manager for Izabel Arruda and team. It runs in the
browser and keeps its data in a Google Sheet through a Google Apps Script web
app. Work comes in from Slack (through a sheet tab that another automation
fills), from Jira (through the Jira API) and from meeting transcripts.

- **Front end:** Vite + TypeScript, plain DOM rendering, Iconify Lucide icons.
- **Back end:** `Code.gs`, a web app bound to the sheet. GET returns everything as JSON; POST saves a full snapshot.
- **AI (optional):** `@canva-ct/genai` summarises uploaded `.txt` / `.vtt` / `.srt` transcripts.

| File | What it does |
| --- | --- |
| `src/main.ts` | State, sync, rendering and events |
| `src/sheet.ts` | The sheet URL and the read/save calls |
| `src/richtext.ts` | Sanitising, linkifying and the rich text editor |
| `src/meetings.ts` | The AI transcript step |
| `src/style.css` | Styles |
| `Code.gs` | The Apps Script back end |
| `tests/` | Code.gs tests against an in-memory sheet, and a local mock of the web app |

## Setup

1. **Create the sheet.** Open Extensions → Apps Script, paste in `Code.gs` and save.
2. **Deploy it.** Deploy → New deployment → Web app, with Execute as: Me and Who has access: Anyone. Copy the `/exec` link.
3. **Point the app at it.** Put the link in `SHEETS_API_URL` in `src/sheet.ts`, or create `.env` with `VITE_SHEETS_API_URL=<link>`.
4. **Install and run.** Fill in the registry addresses in `.npmrc` (needed for `@canva-ct/genai`), then run `npm install` and `npm run dev`. `npm run build` writes the static site to `dist/`.

After every later script change, save the script, then go to Manage deployments, click the pencil and choose **New version**. A new deployment gives you a new link.

**Jira.** Reload the sheet, then choose Task manager → Connect Jira…. This asks for your Jira site, email and API token. You can also set `JIRA_SITE`, `JIRA_EMAIL` and `JIRA_API_TOKEN` in Project settings → Script properties. The default search is `project = L10N AND assignee = currentUser() AND statusCategory != Done`. Set a `JIRA_JQL` script property to change it.

**Slack.** The other automation writes rows to a tab named `Slack Tasks`. The script reads that tab but never writes to it. Each new row is copied once into `Slack Tasks (App)`.

**Meetings.** An external AI can write rows to the `Meetings` tab, with action items as JSON in `actions`. The script turns each action item into a task in the Meetings project. You can also upload a transcript in the app (the + next to Meetings).

## Security notes

- With "Anyone" access, anyone who has the `/exec` link can read and write the sheet. Keep the link private.
- The Jira token is stored only in Script properties. It is never in the code or the sheet.
- The in-app AI needs a valid Canva login. In the Flying Fox preview, an expired login gives a 401 until Flying Fox is restarted.

## How sync works

- **Saving.** The app sends a complete snapshot, debounced to 800 ms, with one save at a time. Edits made during a save go out with the next one. Each save also includes the IDs the app has seen (`seenTaskIds`, `seenProjectIds`, `seenMeetingIds`, `seenCommentIds`, `seenPeople`). The script deletes a row only when the app has seen it and left it out. It keeps rows the app has never seen.
- **Loading.** The app loads on start, when the window regains focus and every 60 s while the window is visible. It skips a load while you type, while a menu or dialog is open, or while a save is pending.
  - When there are unsaved local changes, those changes win, and sheet rows the browser never had are merged in.
  - If the script doesn't send a field (an older version), the browser keeps its own copy and refills the tab.
- **Never sending empty data.** On a first run, or when the browser's cached copy can't be read, the app starts empty. It sends nothing until it has read the sheet.
- **Feed tabs.** Rows in `Slack Tasks (App)` and `Jira Tickets (App)` are never deleted. Their `status` changes to `deleted`, `moved`, `duplicate` or `resolved`, and any status hides the row in the app.
- **Formula injection.** Text that starts with `=`, `+`, `-` or `@` is written with an apostrophe in front, as is text that Sheets would convert to a number or date.

## Testing

```sh
npm test                                  # Code.gs against an in-memory sheet
node tests/mock-server.mjs                # Code.gs behind http://localhost:8787/exec
VITE_SHEETS_API_URL=http://localhost:8787/exec npm run dev
```

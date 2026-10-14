/**
 * Task Manager — Google Apps Script back end.
 *
 * Paste into Extensions → Apps Script of the sheet, save, then
 * Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
 * After every later change: Manage deployments → pencil → Version: New version.
 *
 * GET  → imports Slack/Jira/meeting action items, then returns everything as JSON.
 * POST → saves a complete snapshot from the app (text/plain JSON body).
 *
 * Jira: Task manager → Connect Jira… (stores JIRA_SITE, JIRA_EMAIL and
 * JIRA_API_TOKEN in Script properties). JIRA_JQL overrides the default search.
 */

var OWNER = 'Izabel Arruda';

var T_PROJECTS = 'Projects';
var T_TASKS = 'Tasks';
var T_TAGS = 'Tags';
var T_PEOPLE = 'People';
var T_COMMENTS = 'Comments';
var T_MEETINGS = 'Meetings';
var T_SLACK_SRC = 'Slack Tasks';
var T_SLACK_APP = 'Slack Tasks (App)';
var T_JIRA_APP = 'Jira Tickets (App)';

var H_PROJECTS = ['id', 'name', 'color', 'kind'];
var H_TASKS = ['id', 'projectId', 'title', 'done', 'assignee', 'due', 'tags', 'description', 'createdAt', 'reporter', 'links', 'state'];
var H_TAGS = ['name', 'color'];
var H_PEOPLE = ['name', 'color'];
var H_COMMENTS = ['id', 'taskId', 'author', 'text', 'createdAt'];
var H_MEETINGS = ['id', 'title', 'date', 'attendees', 'notes', 'actions', 'fileName', 'createdAt'];
var H_FEED = ['id', 'sourceId', 'title', 'done', 'assignee', 'due', 'tags', 'description', 'status', 'createdAt', 'reporter', 'links', 'state'];

var SLACK_ID = 'slack';
var JIRA_ID = 'jira';
var MEETINGS_ID = 'meetings';
var INBOX_ID = 'inbox';
var BUILT_IN = {
  slack: { id: SLACK_ID, name: 'Slack Tasks', color: '#E01E5A', kind: 'project' },
  jira: { id: JIRA_ID, name: 'Jira Tickets', color: '#3D8BFF', kind: 'project' },
  meetings: { id: MEETINGS_ID, name: 'Meetings', color: '#14B8A6', kind: 'project' },
  inbox: { id: INBOX_ID, name: 'Inbox', color: '#64748B', kind: 'project' }
};

var TAG_COLORS = ['#ECE1FF', '#DCEFFF', '#D7F5E8', '#FFF2CC', '#FFE3D9', '#FCE0EE', '#D8F3F2', '#EAEAEE'];
var PERSON_COLORS = ['#8B3DFF', '#3D8BFF', '#14B8A6', '#F1A43C', '#F06A6A', '#E362B4', '#2FB36B', '#6366F1'];

var DEFAULT_JQL = 'project = L10N AND assignee = currentUser() AND statusCategory != Done';
var JIRA_MIN_INTERVAL_MS = 2 * 60 * 1000;
var JIRA_FIELDS = ['summary', 'description', 'duedate', 'created', 'reporter', 'assignee', 'status', 'priority', 'issuetype', 'labels'];

// =================================================================== web app

function doGet(e) {
  try {
    var result = withLock_(function () {
      // Imports run inside the lock and re-check what's already copied, so
      // overlapping loads can't copy the same row twice.
      safely_('Slack import', importSlack_);
      safely_('Jira import', function () { importJira_(false); });
      safely_('Meeting action items', createMeetingActionTasks_);
      return readAll_();
    });
    result.ok = true;
    return json_(result);
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    withLock_(function () { saveAll_(body); });
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

function safely_(label, fn) {
  try {
    fn();
  } catch (err) {
    console.error(label + ' failed: ' + (err && err.stack || err));
  }
}

// =================================================================== sheet menu

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Task manager')
    .addItem('Connect Jira…', 'connectJira')
    .addItem('Import from Jira now', 'importJiraNow')
    .addItem('Import from Slack tab now', 'importSlackNow')
    .addToUi();
}

function connectJira() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var site = ui.prompt('Connect Jira (1 of 3)', 'Your Jira site, e.g. yourcompany.atlassian.net' +
    (props.getProperty('JIRA_SITE') ? '\nCurrent: ' + props.getProperty('JIRA_SITE') + ' (leave blank to keep)' : ''), ui.ButtonSet.OK_CANCEL);
  if (site.getSelectedButton() !== ui.Button.OK) return;
  var email = ui.prompt('Connect Jira (2 of 3)', 'The email you sign in to Jira with' +
    (props.getProperty('JIRA_EMAIL') ? '\nCurrent: ' + props.getProperty('JIRA_EMAIL') + ' (leave blank to keep)' : ''), ui.ButtonSet.OK_CANCEL);
  if (email.getSelectedButton() !== ui.Button.OK) return;
  var token = ui.prompt('Connect Jira (3 of 3)', 'A Jira API token (create one at id.atlassian.com → Security → API tokens).' +
    (props.getProperty('JIRA_API_TOKEN') ? '\nLeave blank to keep the saved token.' : ''), ui.ButtonSet.OK_CANCEL);
  if (token.getSelectedButton() !== ui.Button.OK) return;

  var s = String(site.getResponseText() || '').trim().replace(/^https?:\/\//, '').replace(/\/.*$/, '');
  if (s) props.setProperty('JIRA_SITE', s);
  if (String(email.getResponseText()).trim()) props.setProperty('JIRA_EMAIL', String(email.getResponseText()).trim());
  if (String(token.getResponseText()).trim()) props.setProperty('JIRA_API_TOKEN', String(token.getResponseText()).trim());

  try {
    var me = jiraRequest_('get', '/rest/api/3/myself');
    ui.alert('Jira connected', 'Signed in as ' + (me.displayName || me.emailAddress || 'you') + '. Importing now.', ui.ButtonSet.OK);
    importJiraNow();
  } catch (err) {
    ui.alert('Jira connection failed', String(err && err.message || err), ui.ButtonSet.OK);
  }
}

function importJiraNow() {
  withLock_(function () { importJira_(true); });
}

function importSlackNow() {
  withLock_(importSlack_);
}

// =================================================================== table helpers

function normKey_(h) {
  return String(h == null ? '' : h).toLowerCase().replace(/[^a-z0-9]/g, '');
}

/**
 * Opens a tab as a table matched by heading. Missing headings are appended at
 * the end (unless noAppend), so hand-added columns are left alone.
 */
function table_(name, headers, opts) {
  opts = opts || {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    if (opts.noCreate) return null;
    sh = ss.insertSheet(name);
  }
  var lastCol = sh.getLastColumn();
  var lastRow = sh.getLastRow();
  var head = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) { return String(v || '').trim(); }) : [];
  while (head.length && !head[head.length - 1]) head.pop();
  var emptyHead = head.length === 0;
  var idx = {};
  head.forEach(function (h, i) {
    var k = normKey_(h);
    if (k && !(k in idx)) idx[k] = i;
  });
  if (!opts.noAppend || emptyHead) {
    var added = false;
    headers.forEach(function (h) {
      if (!(normKey_(h) in idx)) {
        head.push(h);
        idx[normKey_(h)] = head.length - 1;
        added = true;
      }
    });
    if (added) sh.getRange(1, 1, 1, head.length).setValues([head]);
  }
  var width = Math.max(head.length, lastCol, 1);
  var rows = [];
  if (lastRow > 1) {
    rows = sh.getRange(2, 1, lastRow - 1, width).getValues().filter(function (r) {
      return r.some(function (c) { return c !== '' && c !== null; });
    });
  }
  rows.forEach(function (r) { while (r.length < width) r.push(''); });
  return {
    name: name,
    sh: sh,
    head: head,
    idx: idx,
    width: width,
    rows: rows,
    oldRowCount: Math.max(lastRow - 1, 0),
    dirty: false,
    has: function (key) { return normKey_(key) in idx; },
    get: function (row, key) {
      var i = idx[normKey_(key)];
      return i === undefined ? '' : row[i];
    },
    set: function (row, key, value) {
      var i = idx[normKey_(key)];
      if (i === undefined) return;
      if (row[i] !== value) {
        row[i] = value;
        this.dirty = true;
      }
    },
    blank: function () {
      var r = [];
      for (var i = 0; i < width; i++) r.push('');
      return r;
    }
  };
}

function writeTable_(t) {
  var sh = t.sh;
  if (t.oldRowCount > 0) sh.getRange(2, 1, t.oldRowCount, t.width).clearContent();
  if (t.rows.length) sh.getRange(2, 1, t.rows.length, t.width).setValues(t.rows);
  t.oldRowCount = t.rows.length;
  t.dirty = false;
}

/** Text that Sheets would treat as a formula (= + - @) gets an apostrophe in front. */
function safe_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean' || v instanceof Date) return v;
  var s = String(v);
  if (/^[=+\-@]/.test(s)) return "'" + s;
  // Keep text that Sheets would turn into a number, date or boolean as text.
  if (/^\s*[\d.,:\/$%() -]+\s*$/.test(s) && /\d/.test(s)) return "'" + s;
  if (/^(true|false)$/i.test(s.trim())) return "'" + s;
  return s;
}

function tz_() {
  return Session.getScriptTimeZone() || 'Etc/UTC';
}

/** Date-only value as YYYY-MM-DD. */
function ymd_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  var s = String(v == null ? '' : v).trim();
  if (!s) return '';
  var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  var d = parseDate_(s);
  return d ? Utilities.formatDate(d, tz_(), 'yyyy-MM-dd') : '';
}

/** Date-time value as an ISO string. */
function iso_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v.toISOString();
  var d = parseDate_(v);
  return d ? d.toISOString() : '';
}

/** Parses dates from strings, epoch numbers (seconds or ms) or Date objects. */
function parseDate_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number' || /^\d{9,13}(\.\d+)?$/.test(String(v).trim())) {
    var n = Number(v);
    if (n > 1e12) return new Date(n);
    if (n > 1e9) return new Date(n * 1000);
    return null;
  }
  var s = String(v).trim();
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

/** Due dates are written as real dates (local midnight). */
function dueCell_(s) {
  var d = parseDate_(ymd_(s));
  return d || '';
}

function createdCell_(s) {
  return parseDate_(s) || '';
}

function bool_(v) {
  return v === true || /^(true|yes|1|x|done)$/i.test(String(v == null ? '' : v).trim());
}

function str_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return iso_(v);
  return String(v);
}

function list_(v, sep) {
  if (Array.isArray(v)) return v.map(function (x) { return String(x).trim(); }).filter(String);
  return String(v == null ? '' : v).split(sep).map(function (x) { return x.trim(); }).filter(String);
}

function peopleNames_() {
  var U = table_(T_PEOPLE, H_PEOPLE);
  return U.rows.map(function (r) { return String(U.get(r, 'name') || '').trim(); }).filter(String);
}

function uid_(prefix) {
  return prefix + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 16);
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function textToHtml_(s) {
  var t = String(s == null ? '' : s).trim();
  if (!t) return '';
  return '<p>' + esc_(t).replace(/\r?\n/g, '<br>') + '</p>';
}

function set_(arr) {
  var o = {};
  (arr || []).forEach(function (x) { o[String(x)] = true; });
  return o;
}

function lowerSet_(arr) {
  var o = {};
  (arr || []).forEach(function (x) { o[String(x).trim().toLowerCase()] = true; });
  return o;
}

function urlsIn_(text) {
  var out = [];
  String(text || '').replace(/\bhttps?:\/\/[^\s<>"'|]+[^\s<>"'|.,;:!?)\]}]/gi, function (m) { out.push(m); return m; });
  return out;
}

function uniq_(arr) {
  var seen = {};
  return arr.filter(function (x) {
    if (!x || seen[x]) return false;
    seen[x] = true;
    return true;
  });
}

// =================================================================== read

function readAll_() {
  var P = table_(T_PROJECTS, H_PROJECTS);
  var T = table_(T_TASKS, H_TASKS);
  var G = table_(T_TAGS, H_TAGS);
  var U = table_(T_PEOPLE, H_PEOPLE);
  var C = table_(T_COMMENTS, H_COMMENTS);
  var M = table_(T_MEETINGS, H_MEETINGS, { noAppend: true });
  var S = table_(T_SLACK_APP, H_FEED, { noCreate: !SpreadsheetApp.getActiveSpreadsheet().getSheetByName(T_SLACK_SRC) });
  var J = table_(T_JIRA_APP, H_FEED, { noCreate: true });

  // Rows added by hand without an ID get one.
  [[P, 'p'], [T, 't'], [C, 'c'], [M, 'm']].forEach(function (pair) {
    var tb = pair[0];
    if (!tb.has('id')) return;
    tb.rows.forEach(function (r) {
      if (!String(tb.get(r, 'id')).trim()) tb.set(r, 'id', uid_(pair[1]));
    });
  });
  [S, J].forEach(function (tb) {
    if (!tb) return;
    tb.rows.forEach(function (r) {
      if (!String(tb.get(r, 'id')).trim()) tb.set(r, 'id', uid_(tb === S ? 'sl' : 'jr'));
    });
  });

  var projects = P.rows.map(function (r) {
    return {
      id: str_(P.get(r, 'id')),
      name: str_(P.get(r, 'name')) || 'Untitled',
      color: str_(P.get(r, 'color')),
      kind: str_(P.get(r, 'kind')) === 'notes' ? 'notes' : 'project'
    };
  });
  var projectIds = set_(projects.map(function (p) { return p.id; }));
  function ensureProject(key) {
    var b = BUILT_IN[key];
    if (projectIds[b.id]) return;
    var row = P.blank();
    P.set(row, 'id', b.id);
    P.set(row, 'name', b.name);
    P.set(row, 'color', b.color);
    P.set(row, 'kind', b.kind);
    P.rows.push(row);
    projects.push({ id: b.id, name: b.name, color: b.color, kind: b.kind });
    projectIds[b.id] = true;
  }

  var tasks = [];
  T.rows.forEach(function (r) {
    var pid = str_(T.get(r, 'projectId')).trim();
    if (!projectIds[pid] && pid !== SLACK_ID && pid !== JIRA_ID && pid !== MEETINGS_ID) {
      // Tasks with an unknown project go into the inbox.
      ensureProject('inbox');
      T.set(r, 'projectId', INBOX_ID);
      pid = INBOX_ID;
    }
    if (pid === SLACK_ID) ensureProject('slack');
    if (pid === JIRA_ID) ensureProject('jira');
    if (pid === MEETINGS_ID) ensureProject('meetings');
    tasks.push(taskFromRow_(T, r, pid));
  });
  if (S) {
    ensureProject('slack');
    S.rows.forEach(function (r) {
      if (!String(S.get(r, 'status')).trim()) tasks.push(taskFromRow_(S, r, SLACK_ID));
    });
  }
  if (J) {
    ensureProject('jira');
    J.rows.forEach(function (r) {
      if (!String(J.get(r, 'status')).trim()) tasks.push(taskFromRow_(J, r, JIRA_ID));
    });
  }

  // People: names typed into the sheet that aren't on the list are added.
  var people = U.rows.map(function (r) { return { name: str_(U.get(r, 'name')).trim(), color: str_(U.get(r, 'color')) }; })
    .filter(function (p) { return p.name; });
  var known = lowerSet_(people.map(function (p) { return p.name; }));
  function addPerson(name) {
    name = String(name || '').trim();
    if (!name || known[name.toLowerCase()]) return;
    var color = PERSON_COLORS[people.length % PERSON_COLORS.length];
    var row = U.blank();
    U.set(row, 'name', safe_(name));
    U.set(row, 'color', color);
    U.rows.push(row);
    people.push({ name: name, color: color });
    known[name.toLowerCase()] = true;
  }
  addPerson(OWNER);
  tasks.forEach(function (t) { addPerson(t.assignee); });

  var tags = G.rows.map(function (r) { return { name: str_(G.get(r, 'name')).trim(), color: str_(G.get(r, 'color')) }; })
    .filter(function (t) { return t.name; });

  var comments = C.rows.map(function (r) {
    return {
      id: str_(C.get(r, 'id')),
      taskId: str_(C.get(r, 'taskId')),
      author: str_(C.get(r, 'author')) || OWNER,
      text: str_(C.get(r, 'text')),
      createdAt: iso_(C.get(r, 'createdAt'))
    };
  });

  var meetings = M.rows.map(function (r) { return meetingFromRow_(M, r); });

  [P, T, U, C, M, S, J].forEach(function (tb) { if (tb && tb.dirty) writeTable_(tb); });

  return { projects: projects, tasks: tasks, tags: tags, people: people, comments: comments, meetings: meetings };
}

function taskFromRow_(tb, r, projectId) {
  return {
    id: str_(tb.get(r, 'id')),
    projectId: projectId,
    title: str_(tb.get(r, 'title')),
    done: bool_(tb.get(r, 'done')),
    assignee: str_(tb.get(r, 'assignee')).trim(),
    due: ymd_(tb.get(r, 'due')),
    tags: list_(tb.get(r, 'tags'), ','),
    description: str_(tb.get(r, 'description')),
    createdAt: iso_(tb.get(r, 'createdAt')),
    reporter: str_(tb.get(r, 'reporter')),
    links: list_(tb.get(r, 'links'), /[\r\n]+/),
    state: normState_(tb.get(r, 'state'))
  };
}

function normState_(v) {
  var s = String(v == null ? '' : v).trim().toLowerCase();
  return s === 'urgent' ? 'urgent' : (s === 'progress' || s === 'in progress') ? 'progress' : '';
}

function parseActions_(v) {
  if (Array.isArray(v)) return v;
  var s = String(v == null ? '' : v).trim();
  if (!s) return [];
  try {
    var a = JSON.parse(s);
    return Array.isArray(a) ? a : [];
  } catch (err) {
    return [];
  }
}

function meetingFromRow_(M, r) {
  return {
    id: str_(M.get(r, 'id')),
    title: str_(M.get(r, 'title')) || 'Meeting',
    date: ymd_(M.get(r, 'date')),
    attendees: str_(M.get(r, 'attendees')),
    notes: str_(M.get(r, 'notes')),
    actions: parseActions_(M.get(r, 'actions')).map(function (a) {
      return {
        title: str_(a && a.title), owner: str_(a && a.owner), due: ymd_(a && a.due),
        details: str_(a && a.details), taskId: str_(a && a.taskId)
      };
    }),
    fileName: str_(M.get(r, 'fileName')),
    createdAt: iso_(M.get(r, 'createdAt'))
  };
}

// =================================================================== save

/**
 * Saves the app's snapshot. Rows the app has never seen (added by Slack, Jira,
 * an external AI or by hand) are kept; rows it has seen but left out were
 * deleted in the app. Feed rows are never physically deleted — they get a status.
 */
function saveAll_(body) {
  var seenTasks = set_(body.seenTaskIds);
  var seenProjects = set_(body.seenProjectIds);
  var seenMeetings = set_(body.seenMeetingIds);
  var seenComments = body.seenCommentIds ? set_(body.seenCommentIds) : null;
  var seenPeople = body.seenPeople ? lowerSet_(body.seenPeople) : null;

  // ---- projects
  if (Array.isArray(body.projects)) {
    var P = table_(T_PROJECTS, H_PROJECTS);
    upsert_(P, 'id', body.projects, function (p) { return str_(p.id); }, seenProjects, function (row, p) {
      P.set(row, 'id', safe_(p.id));
      P.set(row, 'name', safe_(p.name));
      P.set(row, 'color', safe_(p.color));
      P.set(row, 'kind', p.kind === 'notes' ? 'notes' : 'project');
    });
    writeTable_(P);
  }

  // ---- tasks (routed between the Tasks tab and the feed tabs)
  if (Array.isArray(body.tasks)) {
    var feeds = [
      { t: table_(T_SLACK_APP, H_FEED, { noCreate: true }), pid: SLACK_ID },
      { t: table_(T_JIRA_APP, H_FEED, { noCreate: true }), pid: JIRA_ID }
    ].filter(function (f) { return f.t; });
    var live = {};
    feeds.forEach(function (f) {
      f.t.rows.forEach(function (r) {
        if (!String(f.t.get(r, 'status')).trim()) live[str_(f.t.get(r, 'id'))] = { feed: f, row: r, used: false };
      });
    });

    var forTasksTab = [];
    body.tasks.forEach(function (task) {
      var id = str_(task.id);
      var hit = live[id];
      if (hit) {
        if (task.projectId === hit.feed.pid) {
          writeTaskFields_(hit.feed.t, hit.row, task);
          hit.used = true;
          return;
        }
        hit.feed.t.set(hit.row, 'status', 'moved'); // now lives in the Tasks tab
        hit.used = true;
      }
      forTasksTab.push(task);
    });
    Object.keys(live).forEach(function (id) {
      var hit = live[id];
      if (!hit.used && seenTasks[id]) hit.feed.t.set(hit.row, 'status', 'deleted');
    });
    feeds.forEach(function (f) { if (f.t.dirty) writeTable_(f.t); });

    var T = table_(T_TASKS, H_TASKS);
    upsert_(T, 'id', forTasksTab, function (t) { return str_(t.id); }, seenTasks, function (row, task) {
      T.set(row, 'id', safe_(task.id));
      T.set(row, 'projectId', safe_(task.projectId));
      writeTaskFields_(T, row, task);
    });
    writeTable_(T);
  }

  // ---- comments
  if (Array.isArray(body.comments)) {
    var C = table_(T_COMMENTS, H_COMMENTS);
    upsert_(C, 'id', body.comments, function (c) { return str_(c.id); }, seenComments || {}, function (row, c) {
      C.set(row, 'id', safe_(c.id));
      C.set(row, 'taskId', safe_(c.taskId));
      C.set(row, 'author', safe_(c.author));
      C.set(row, 'text', safe_(c.text));
      C.set(row, 'createdAt', createdCell_(c.createdAt));
    });
    writeTable_(C);
  }

  // ---- meetings (columns matched by heading; headings are not added)
  if (Array.isArray(body.meetings)) {
    var M = table_(T_MEETINGS, H_MEETINGS, { noAppend: true });
    upsert_(M, 'id', body.meetings, function (m) { return str_(m.id); }, seenMeetings, function (row, m) {
      M.set(row, 'id', safe_(m.id));
      M.set(row, 'title', safe_(m.title));
      M.set(row, 'date', dueCell_(m.date));
      M.set(row, 'attendees', safe_(Array.isArray(m.attendees) ? m.attendees.join(', ') : m.attendees));
      M.set(row, 'notes', safe_(m.notes));
      M.set(row, 'actions', safe_(typeof m.actions === 'string' ? m.actions : JSON.stringify(m.actions || [])));
      M.set(row, 'fileName', safe_(m.fileName));
      M.set(row, 'createdAt', createdCell_(m.createdAt));
    });
    writeTable_(M);
  }

  // ---- people (keyed by name)
  if (Array.isArray(body.people)) {
    var U = table_(T_PEOPLE, H_PEOPLE);
    upsert_(U, 'name', body.people, function (p) { return String(p.name || '').trim().toLowerCase(); }, seenPeople || {}, function (row, p) {
      U.set(row, 'name', safe_(String(p.name || '').trim()));
      U.set(row, 'color', safe_(p.color));
    }, true);
    writeTable_(U);
  }

  // ---- tags (keyed by name; never removed here, only recoloured/added)
  if (Array.isArray(body.tags)) {
    var G = table_(T_TAGS, H_TAGS);
    upsert_(G, 'name', body.tags, function (t) { return String(t.name || '').trim().toLowerCase(); }, {}, function (row, t) {
      G.set(row, 'name', safe_(String(t.name || '').trim()));
      G.set(row, 'color', safe_(t.color));
    }, true);
    writeTable_(G);
  }
}

function writeTaskFields_(tb, row, task) {
  tb.set(row, 'title', safe_(task.title));
  tb.set(row, 'done', !!task.done);
  tb.set(row, 'assignee', safe_(task.assignee));
  tb.set(row, 'due', dueCell_(task.due));
  tb.set(row, 'tags', safe_(list_(task.tags, ',').join(', ')));
  tb.set(row, 'description', safe_(task.description));
  if (task.createdAt) tb.set(row, 'createdAt', createdCell_(task.createdAt));
  tb.set(row, 'reporter', safe_(task.reporter));
  tb.set(row, 'links', safe_(list_(task.links, /[\r\n]+/).filter(function (l) { return /^https?:\/\//i.test(l); }).join('\n')));
  tb.set(row, 'state', task.state === 'urgent' || task.state === 'progress' ? task.state : '');
}

/**
 * Updates rows in place by key, appends new ones, removes rows whose key the
 * app has seen but no longer sends, and keeps rows it hasn't seen.
 */
function upsert_(tb, keyField, items, keyOf, seen, write, lowerKeys) {
  var incoming = {};
  items.forEach(function (it) {
    var k = keyOf(it);
    if (k) incoming[k] = it;
  });
  var kept = [];
  var placed = {};
  tb.rows.forEach(function (r) {
    var raw = String(tb.get(r, keyField)).trim();
    var k = lowerKeys ? raw.toLowerCase() : raw;
    if (k && incoming[k]) {
      if (placed[k]) return; // drop exact duplicates of a row the app sends
      placed[k] = true;
      write(r, incoming[k]);
      kept.push(r);
    } else if (k && seen[k]) {
      tb.dirty = true; // deleted in the app
    } else {
      kept.push(r);
    }
  });
  items.forEach(function (it) {
    var k = keyOf(it);
    if (!k || placed[k]) return;
    placed[k] = true;
    var row = tb.blank();
    write(row, it);
    kept.push(row);
    tb.dirty = true;
  });
  tb.rows = kept;
}

// =================================================================== Slack import

var SLACK_FIELDS = [
  ['id', [['id'], ['slack', 'id'], ['message', 'id'], ['ts']]],
  ['projectId', [['project', 'id'], ['project']]],
  ['title', [['title'], ['task'], ['summary'], ['subject']]],
  ['assignee', [['assignee'], ['assigned', 'to'], ['owner']]],
  ['reporter', [['reporter'], ['requested', 'by'], ['requester'], ['author']]],
  ['due', [['due'], ['deadline']]],
  ['description', [['description'], ['details'], ['body'], ['message']]],
  ['links', [['original', 'slack', 'links'], ['slack', 'links'], ['links'], ['link'], ['url'], ['permalink']]],
  ['createdAt', [['created', 'at'], ['created'], ['timestamp']]]
];

function words_(h) {
  return String(h || '').replace(/([a-z])([A-Z])/g, '$1 $2').toLowerCase().split(/[^a-z0-9]+/).filter(String);
}

function containsSeq_(arr, seq) {
  for (var i = 0; i + seq.length <= arr.length; i++) {
    var ok = true;
    for (var j = 0; j < seq.length; j++) if (arr[i + j] !== seq[j]) { ok = false; break; }
    if (ok) return true;
  }
  return false;
}

/** Maps source headings to fields: exact matches first, then whole-word matches (ignoring case). */
function mapHeadings_(head) {
  var w = head.map(words_);
  var map = {};
  var used = {};
  SLACK_FIELDS.forEach(function (f) {
    for (var s = 0; s < f[1].length && !(f[0] in map); s++) {
      for (var i = 0; i < w.length; i++) {
        if (!used[i] && w[i].join(' ') === f[1][s].join(' ')) { map[f[0]] = i; used[i] = true; break; }
      }
    }
  });
  SLACK_FIELDS.forEach(function (f) {
    for (var s = 0; s < f[1].length && !(f[0] in map); s++) {
      for (var i = 0; i < w.length; i++) {
        if (!used[i] && w[i].length && containsSeq_(w[i], f[1][s])) { map[f[0]] = i; used[i] = true; break; }
      }
    }
  });
  return map;
}

/**
 * Copies each new row of the "Slack Tasks" tab (filled by another automation)
 * once, by source ID, into "Slack Tasks (App)". Never writes to the source tab.
 */
function importSlack_() {
  var src = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(T_SLACK_SRC);
  if (!src || src.getLastRow() < 2 || src.getLastColumn() < 1) return;
  var values = src.getRange(1, 1, src.getLastRow(), src.getLastColumn()).getValues();
  var head = values[0].map(function (h) { return String(h || '').trim(); });
  var map = mapHeadings_(head);
  var app = table_(T_SLACK_APP, H_FEED);

  // Index existing copies by source ID (any status), marking duplicates.
  var bySource = {};
  app.rows.forEach(function (r) {
    var sid = str_(app.get(r, 'sourceId')).trim();
    if (!sid) return;
    var status = String(app.get(r, 'status')).trim();
    if (!bySource[sid]) {
      bySource[sid] = r;
    } else if (!status) {
      var first = bySource[sid];
      if (String(app.get(first, 'status')).trim()) bySource[sid] = r;
      else app.set(r, 'status', 'duplicate');
    }
  });

  function cell(row, field) {
    return field in map ? row[map[field]] : '';
  }

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row.some(function (c) { return c !== '' && c !== null; })) continue;
    var title = str_(cell(row, 'title')).trim();
    var description = str_(cell(row, 'description'));
    var sid = str_(cell(row, 'id')).trim() || ('row:' + title + '|' + str_(cell(row, 'createdAt')));
    var links = uniq_(list_(str_(cell(row, 'links')), /[\s,]+/).filter(function (l) { return /^https?:\/\//i.test(l); })
      .concat(urlsIn_(title)).concat(urlsIn_(description)));
    var reporter = str_(cell(row, 'reporter')).trim();
    var created = parseDate_(cell(row, 'createdAt'));

    var existing = bySource[sid];
    if (existing) {
      // Backfill rows copied earlier: fill empty cells only.
      if (!String(app.get(existing, 'reporter')).trim() && reporter) app.set(existing, 'reporter', safe_(reporter));
      if (!String(app.get(existing, 'links')).trim() && links.length) app.set(existing, 'links', safe_(links.join('\n')));
      if (!String(app.get(existing, 'createdAt')).trim() && created) app.set(existing, 'createdAt', created);
      continue;
    }

    var extra = [];
    head.forEach(function (h, c) {
      var mapped = false;
      for (var k in map) if (map[k] === c) mapped = true;
      if (!mapped && h && row[c] !== '' && row[c] !== null) extra.push('<p><b>' + esc_(h) + ':</b> ' + esc_(str_(row[c])) + '</p>');
    });

    var out = app.blank();
    app.set(out, 'id', uid_('sl'));
    app.set(out, 'sourceId', safe_(sid));
    app.set(out, 'title', safe_(title || '(untitled Slack task)'));
    app.set(out, 'done', false);
    app.set(out, 'assignee', safe_(str_(cell(row, 'assignee')).trim()));
    app.set(out, 'due', parseDate_(cell(row, 'due')) || '');
    app.set(out, 'tags', 'Slack');
    app.set(out, 'description', safe_(textToHtml_(description) + extra.join('')));
    app.set(out, 'status', '');
    app.set(out, 'createdAt', created || new Date());
    app.set(out, 'reporter', safe_(reporter));
    app.set(out, 'links', safe_(links.join('\n')));
    app.set(out, 'state', '');
    app.rows.push(out);
    app.dirty = true;
    bySource[sid] = out;
  }
  if (app.dirty) writeTable_(app);
}

// =================================================================== Jira import

function jiraRequest_(method, path, payload) {
  var props = PropertiesService.getScriptProperties();
  var site = props.getProperty('JIRA_SITE');
  var email = props.getProperty('JIRA_EMAIL');
  var token = props.getProperty('JIRA_API_TOKEN');
  if (!site || !email || !token) throw new Error('Jira is not connected');
  var res = UrlFetchApp.fetch('https://' + site + path, {
    method: method,
    contentType: 'application/json',
    headers: { Authorization: 'Basic ' + Utilities.base64Encode(email + ':' + token), Accept: 'application/json' },
    payload: payload ? JSON.stringify(payload) : undefined,
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  if (code < 200 || code >= 300) throw new Error('Jira ' + code + ': ' + res.getContentText().slice(0, 300));
  return JSON.parse(res.getContentText() || '{}');
}

function jiraSearch_(jql, fields, max) {
  var issues = [];
  var token = null;
  do {
    var body = { jql: jql, fields: fields, maxResults: 100 };
    if (token) body.nextPageToken = token;
    var res = jiraRequest_('post', '/rest/api/3/search/jql', body);
    issues = issues.concat(res.issues || []);
    token = res.isLast === false || res.nextPageToken ? res.nextPageToken : null;
  } while (token && issues.length < (max || 500));
  return issues;
}

/** Searches Jira (every 2 minutes at most) and copies new tickets into "Jira Tickets (App)". */
function importJira_(force) {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('JIRA_SITE') || !props.getProperty('JIRA_EMAIL') || !props.getProperty('JIRA_API_TOKEN')) return;
  var last = Number(props.getProperty('JIRA_LAST_FETCH') || 0);
  if (!force && Date.now() - last < JIRA_MIN_INTERVAL_MS) return;
  props.setProperty('JIRA_LAST_FETCH', String(Date.now()));

  var site = props.getProperty('JIRA_SITE');
  var jql = props.getProperty('JIRA_JQL') || DEFAULT_JQL;
  var issues = jiraSearch_(jql, JIRA_FIELDS, 500);

  var app = table_(T_JIRA_APP, H_FEED);
  var peopleSet = lowerSet_(peopleNames_().concat([OWNER]));

  var bySource = {};
  app.rows.forEach(function (r) {
    var sid = str_(app.get(r, 'sourceId')).trim();
    if (!sid) return;
    var status = String(app.get(r, 'status')).trim();
    if (!bySource[sid]) bySource[sid] = r;
    else if (!status) {
      if (String(app.get(bySource[sid], 'status')).trim() === '') app.set(r, 'status', 'duplicate');
      else bySource[sid] = r;
    }
  });

  var returned = {};
  issues.forEach(function (issue) {
    var key = issue.key;
    returned[key] = true;
    var existing = bySource[key];
    if (existing) {
      // A ticket that was resolved and reopened comes back.
      if (String(app.get(existing, 'status')).trim() === 'resolved') app.set(existing, 'status', '');
      return;
    }
    var f = issue.fields || {};
    var assignee = f.assignee && f.assignee.displayName ? f.assignee.displayName : '';
    var meta = [];
    if (f.status && f.status.name) meta.push('<b>Status:</b> ' + esc_(f.status.name));
    if (f.priority && f.priority.name) meta.push('<b>Priority:</b> ' + esc_(f.priority.name));
    if (f.issuetype && f.issuetype.name) meta.push('<b>Type:</b> ' + esc_(f.issuetype.name));
    if (f.labels && f.labels.length) meta.push('<b>Labels:</b> ' + esc_(f.labels.join(', ')));

    var row = app.blank();
    app.set(row, 'id', uid_('jr'));
    app.set(row, 'sourceId', safe_(key));
    app.set(row, 'title', safe_(key + ': ' + (f.summary || '')));
    app.set(row, 'done', false);
    app.set(row, 'assignee', peopleSet[assignee.toLowerCase()] ? safe_(assignee) : '');
    app.set(row, 'due', parseDate_(f.duedate) || '');
    app.set(row, 'tags', 'Jira');
    app.set(row, 'description', safe_(adfToHtml_(f.description) + (meta.length ? '<p>' + meta.join(' · ') + '</p>' : '')));
    app.set(row, 'status', '');
    app.set(row, 'createdAt', parseDate_(f.created) || new Date());
    app.set(row, 'reporter', safe_(f.reporter && f.reporter.displayName || ''));
    app.set(row, 'links', 'https://' + site + '/browse/' + encodeURIComponent(key));
    app.set(row, 'state', '');
    app.rows.push(row);
    app.dirty = true;
    bySource[key] = row;
  });

  // Live tickets the search no longer returns: mark resolved if Jira says they're done.
  var missing = Object.keys(bySource).filter(function (k) {
    return !returned[k] && !String(app.get(bySource[k], 'status')).trim() && /^[A-Z][A-Z0-9_]+-\d+$/.test(k);
  });
  for (var i = 0; i < missing.length; i += 50) {
    var chunk = missing.slice(i, i + 50);
    try {
      jiraSearch_('key in (' + chunk.join(',') + ')', ['status'], 100).forEach(function (issue) {
        var cat = issue.fields && issue.fields.status && issue.fields.status.statusCategory;
        if (cat && cat.key === 'done' && bySource[issue.key]) app.set(bySource[issue.key], 'status', 'resolved');
      });
    } catch (err) {
      console.warn('Jira status check failed: ' + err);
    }
  }
  if (app.dirty) writeTable_(app);
}

/** Flattens Jira's rich-text format (ADF) into the app's basic HTML. */
function adfToHtml_(node) {
  if (!node) return '';
  if (typeof node === 'string') return textToHtml_(node);
  function inner(n) {
    return (n.content || []).map(walk).join('');
  }
  function walk(n) {
    if (!n) return '';
    switch (n.type) {
      case 'doc': return inner(n);
      case 'paragraph': return '<p>' + inner(n) + '</p>';
      case 'heading': return '<p><b>' + inner(n) + '</b></p>';
      case 'bulletList': return '<ul>' + inner(n) + '</ul>';
      case 'orderedList': return '<ol>' + inner(n) + '</ol>';
      case 'listItem': return '<li>' + inner(n).replace(/^<p>([\s\S]*)<\/p>$/, '$1') + '</li>';
      case 'blockquote': return '<blockquote>' + inner(n) + '</blockquote>';
      case 'codeBlock': return '<p><code>' + inner(n) + '</code></p>';
      case 'hardBreak': return '<br>';
      case 'rule': return '<br>';
      case 'mention': return esc_((n.attrs && n.attrs.text) || '@someone');
      case 'emoji': return esc_((n.attrs && (n.attrs.text || n.attrs.shortName)) || '');
      case 'inlineCard':
      case 'blockCard': {
        var url = n.attrs && n.attrs.url;
        return url && /^https?:/i.test(url) ? '<a href="' + esc_(url) + '">' + esc_(url) + '</a>' : '';
      }
      case 'text': {
        var t = esc_(n.text || '');
        (n.marks || []).forEach(function (m) {
          if (m.type === 'strong') t = '<b>' + t + '</b>';
          else if (m.type === 'em') t = '<i>' + t + '</i>';
          else if (m.type === 'underline') t = '<u>' + t + '</u>';
          else if (m.type === 'strike') t = '<s>' + t + '</s>';
          else if (m.type === 'code') t = '<code>' + t + '</code>';
          else if (m.type === 'link' && m.attrs && /^(https?:|mailto:)/i.test(m.attrs.href || '')) t = '<a href="' + esc_(m.attrs.href) + '">' + t + '</a>';
        });
        return t;
      }
      default: return inner(n);
    }
  }
  return walk(node);
}

// =================================================================== meeting action items

/**
 * Meetings written into the sheet by an external AI may list action items
 * without tasks. Each gets a real task in the Meetings project, tagged
 * "Meeting", and the meeting records the task's ID.
 */
function createMeetingActionTasks_() {
  var M = table_(T_MEETINGS, H_MEETINGS, { noAppend: true });
  if (!M.rows.length || !M.has('actions')) return;
  var T = null;
  var people = null;
  M.rows.forEach(function (r) {
    if (M.has('id') && !String(M.get(r, 'id')).trim()) M.set(r, 'id', uid_('m'));
    var actions = parseActions_(M.get(r, 'actions'));
    var changed = false;
    actions.forEach(function (a) {
      if (!a || !String(a.title || '').trim() || a.taskId) return;
      if (!T) {
        T = table_(T_TASKS, H_TASKS);
        people = peopleNames_().concat([OWNER]);
      }
      var title = str_(M.get(r, 'title')) || 'Meeting';
      var date = ymd_(M.get(r, 'date'));
      var owner = String(a.owner || '').trim();
      var assignee = '';
      people.forEach(function (p) { if (p && owner && p.toLowerCase() === owner.toLowerCase()) assignee = p; });
      var when = date ? ' (' + Utilities.formatDate(parseDate_(date), tz_(), 'MMM d') + ')' : '';
      var task = T.blank();
      var id = uid_('t');
      T.set(task, 'id', id);
      T.set(task, 'projectId', MEETINGS_ID);
      T.set(task, 'title', safe_(String(a.title).trim()));
      T.set(task, 'done', false);
      T.set(task, 'assignee', safe_(assignee));
      T.set(task, 'due', dueCell_(a.due));
      T.set(task, 'tags', 'Meeting');
      T.set(task, 'description', safe_('<p>From meeting: ' + esc_(title) + esc_(when) + '</p>' + (a.details ? textToHtml_(a.details) : '')));
      T.set(task, 'createdAt', new Date());
      T.rows.push(task);
      T.dirty = true;
      a.taskId = id;
      changed = true;
    });
    if (changed) M.set(r, 'actions', JSON.stringify(actions));
  });
  if (T && T.dirty) writeTable_(T);
  if (M.dirty) writeTable_(M);
  if (T) {
    // Make sure the Meetings tag has a colour.
    var G = table_(T_TAGS, H_TAGS);
    var has = G.rows.some(function (g) { return String(G.get(g, 'name')).trim().toLowerCase() === 'meeting'; });
    if (!has) {
      var row = G.blank();
      G.set(row, 'name', 'Meeting');
      G.set(row, 'color', TAG_COLORS[G.rows.length % TAG_COLORS.length]);
      G.rows.push(row);
      writeTable_(G);
    }
  }
}

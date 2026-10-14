// Tests for Code.gs sync logic against an in-memory sheet. Run: npm test
import assert from 'node:assert/strict';
import { createGas, Formula } from './gas-mock.mjs';

process.env.TZ = 'UTC';
let passed = 0;
const tests = [];
const test = (name, fn) => tests.push([name, fn]);

const project = (id, name = id, kind = 'project') => ({ id, name, color: '#8B3DFF', kind });
const task = (id, projectId, extra = {}) => ({
  id, projectId, title: `Task ${id}`, done: false, assignee: '', due: '', tags: [], description: '',
  createdAt: '2026-10-01T10:00:00.000Z', reporter: '', links: [], state: '', ...extra,
});
const snapshot = (over = {}) => ({
  projects: [], tasks: [], tags: [], people: [{ name: 'Izabel Arruda', color: '#8B3DFF' }], comments: [], meetings: [],
  seenTaskIds: [], seenProjectIds: [], seenMeetingIds: [], seenCommentIds: [], seenPeople: [], ...over,
});

test('empty sheet loads with the owner and creates tabs with headings', () => {
  const g = createGas();
  const r = g.get();
  assert.equal(r.ok, true);
  assert.deepEqual(r.tasks, []);
  assert.deepEqual(r.people.map((p) => p.name), ['Izabel Arruda']);
  assert.deepEqual(g.sheet('Tasks').data[0], ['id', 'projectId', 'title', 'done', 'assignee', 'due', 'tags', 'description', 'createdAt', 'reporter', 'links', 'state']);
  assert.ok(!g.sheets.has('Slack Tasks (App)'), 'no Slack app tab without a Slack source tab');
});

test('round trip keeps types: done, due, tags, links, state, kind, comments, meetings', () => {
  const g = createGas();
  g.get();
  g.post(snapshot({
    projects: [project('p1', 'Work'), project('n1', 'Ideas', 'notes')],
    tasks: [
      task('t1', 'p1', { done: true, due: '2026-10-20', tags: ['A', 'B'], links: ['https://x.com/a', 'javascript:alert(1)'], state: 'urgent', assignee: 'Sam Lee', description: '<p>Hi</p>' }),
      task('t2', 'n1', { title: 'A note' }),
    ],
    comments: [{ id: 'c1', taskId: 't1', author: 'Izabel Arruda', text: '<b>yo</b>', createdAt: '2026-10-02T09:30:00.000Z' }],
    meetings: [{ id: 'm1', title: 'Sync', date: '2026-10-03', attendees: 'A, B', notes: '<p>n</p>', actions: JSON.stringify([{ title: 'Do', taskId: 't1' }]), fileName: 'x.vtt', createdAt: '2026-10-03T01:00:00.000Z' }],
    tags: [{ name: 'A', color: '#ECE1FF' }],
  }));
  const r = g.get();
  const t1 = r.tasks.find((t) => t.id === 't1');
  assert.equal(t1.done, true);
  assert.equal(t1.due, '2026-10-20');
  assert.deepEqual(t1.tags, ['A', 'B']);
  assert.deepEqual(t1.links, ['https://x.com/a']);
  assert.equal(t1.state, 'urgent');
  assert.equal(t1.createdAt, '2026-10-01T10:00:00.000Z');
  assert.equal(r.projects.find((p) => p.id === 'n1').kind, 'notes');
  assert.equal(r.comments[0].text, '<b>yo</b>');
  assert.equal(r.comments[0].createdAt, '2026-10-02T09:30:00.000Z');
  assert.equal(r.meetings[0].date, '2026-10-03');
  assert.equal(r.meetings[0].actions[0].taskId, 't1');
  assert.ok(r.people.some((p) => p.name === 'Sam Lee'), 'assignee added to People');
  assert.deepEqual(r.tags.map((t) => t.name), ['A']);
});

test('formula injection and coercion: text stays text', () => {
  const g = createGas();
  g.get();
  g.post(snapshot({ projects: [project('p1')], tasks: [
    task('t1', 'p1', { title: '=HYPERLINK("http://evil","x")' }),
    task('t2', 'p1', { title: '+1 for this' }),
    task('t3', 'p1', { title: '2026-01-01' }),
    task('t4', 'p1', { title: '@here please' }),
    task('t5', 'p1', { title: 'TRUE' }),
    task('t6', 'p1', { title: '1/2' }),
  ] }));
  const cells = g.sheet('Tasks').data.slice(1).map((r) => r[2]);
  assert.ok(cells.every((c) => typeof c === 'string' && !(c instanceof Formula)), JSON.stringify(cells));
  const r = g.get();
  assert.deepEqual(r.tasks.map((t) => t.title), ['=HYPERLINK("http://evil","x")', '+1 for this', '2026-01-01', '@here please', 'TRUE', '1/2']);
});

test('unseen rows are kept; seen rows missing from the snapshot are deleted', () => {
  const g = createGas();
  g.get();
  g.post(snapshot({ projects: [project('p1')], tasks: [task('t1', 'p1')] }));
  // Someone adds a row by hand, without an ID.
  const sh = g.sheet('Tasks');
  sh.data.push(['', 'p1', 'Hand added', false, '', '', '', '', '', '', '', '']);
  const r1 = g.get();
  const hand = r1.tasks.find((t) => t.title === 'Hand added');
  assert.ok(hand && hand.id, 'hand-added row got an ID');
  // A save from a browser that hasn't seen it must keep it.
  g.post(snapshot({ projects: [project('p1')], tasks: [task('t1', 'p1', { title: 'edited' })], seenTaskIds: ['t1'], seenProjectIds: ['p1'] }));
  let r = g.get();
  assert.deepEqual(r.tasks.map((t) => t.title).sort(), ['Hand added', 'edited']);
  // Once seen, leaving it out deletes it.
  g.post(snapshot({ projects: [project('p1')], tasks: [task('t1', 'p1')], seenTaskIds: ['t1', hand.id], seenProjectIds: ['p1'] }));
  r = g.get();
  assert.deepEqual(r.tasks.map((t) => t.id), ['t1']);
});

test('older app payload (no seen lists, no meetings) deletes nothing', () => {
  const g = createGas();
  g.get();
  g.post(snapshot({ projects: [project('p1')], tasks: [task('t1', 'p1'), task('t2', 'p1')],
    meetings: [{ id: 'm1', title: 'M', date: '2026-10-01', attendees: '', notes: '', actions: '[]', fileName: '', createdAt: '' }] }));
  g.post({ projects: [project('p1')], tasks: [task('t1', 'p1')], people: [{ name: 'Izabel Arruda', color: '#000' }] });
  const r = g.get();
  assert.deepEqual(r.tasks.map((t) => t.id).sort(), ['t1', 't2']);
  assert.equal(r.meetings.length, 1);
});

test('tasks with an unknown project go into an inbox project', () => {
  const g = createGas();
  g.get();
  g.sheet('Tasks').data.push(['t9', 'nope', 'Lost', '', '', '', '', '', '', '', '', '']);
  const r = g.get();
  assert.equal(r.tasks[0].projectId, 'inbox');
  assert.ok(r.projects.some((p) => p.id === 'inbox' && p.name === 'Inbox'));
});

function slackSource(g) {
  g.sheet('Slack Tasks').setRaw([
    ['id', 'projectId', 'title', 'assignee', 'Reporter', 'due', 'description', 'originalSlackLinks', 'createdAt', 'Channel'],
    ['S1', 'x', 'Review deck https://canva.com/d/1', 'Izabel Arruda', 'Ana', '2026-10-20', 'See https://docs.x/y and thanks', 'https://slack.com/archives/1 https://canva.com/d/1', '2026-10-01T08:00:00Z', '#l10n'],
    ['S2', 'x', 'Second', 'New Person', '', '', 'line1\nline2', '', '', ''],
  ]);
}

test('Slack import: copies once by ID, maps headings, links, tag, extra columns, never writes the source', () => {
  const g = createGas();
  slackSource(g);
  const before = JSON.stringify(g.sheet('Slack Tasks').data);
  const r = g.get();
  g.get(); // second load must not copy again
  assert.equal(JSON.stringify(g.sheet('Slack Tasks').data), before, 'source tab untouched');
  const app = g.sheet('Slack Tasks (App)').objects();
  assert.equal(app.length, 2);
  const t = r.tasks.find((x) => x.title.startsWith('Review'));
  assert.equal(t.projectId, 'slack');
  assert.equal(t.reporter, 'Ana');
  assert.deepEqual(t.tags, ['Slack']);
  assert.equal(t.due, '2026-10-20');
  assert.equal(t.createdAt, '2026-10-01T08:00:00.000Z');
  assert.deepEqual(t.links, ['https://slack.com/archives/1', 'https://canva.com/d/1', 'https://docs.x/y']);
  assert.match(t.description, /<b>Channel:<\/b> #l10n/);
  assert.ok(r.projects.some((p) => p.id === 'slack'));
  assert.ok(r.people.some((p) => p.name === 'New Person'));
  assert.match(r.tasks.find((x) => x.title === 'Second').description, /line1<br>line2/);
});

test('Slack import: duplicates by source ID are marked; backfill fills empty cells only', () => {
  const g = createGas();
  slackSource(g);
  g.get();
  const app = g.sheet('Slack Tasks (App)');
  const head = app.data[0];
  const col = (k) => head.indexOf(k);
  // Simulate an overlapping copy and an older copy without reporter/links.
  const dup = app.data[1].slice();
  dup[col('id')] = 'sl-dup';
  app.data.push(dup);
  app.data[1][col('reporter')] = '';
  app.data[1][col('links')] = '';
  app.data[2][col('reporter')] = 'Kept';
  const r = g.get();
  const rows = app.objects();
  assert.equal(rows.find((x) => x.id === 'sl-dup').status, 'duplicate');
  assert.equal(rows[0].reporter, 'Ana');
  assert.match(rows[0].links, /slack\.com/);
  assert.equal(rows[1].reporter, 'Kept');
  assert.equal(r.tasks.filter((t) => t.projectId === 'slack').length, 2);
});

test('feed rows: edits write back, moves mark "moved", deletes mark "deleted", rows never removed', () => {
  const g = createGas();
  slackSource(g);
  const r = g.get();
  const [a, b] = r.tasks;
  const seenTaskIds = r.tasks.map((t) => t.id);
  g.post(snapshot({ projects: [...r.projects, project('p1')], seenTaskIds, seenProjectIds: r.projects.map((p) => p.id),
    tasks: [{ ...a, done: true, title: 'Edited in app' }, { ...b, projectId: 'p1' }] }));
  let rows = g.sheet('Slack Tasks (App)').objects();
  assert.equal(rows[0].title, 'Edited in app');
  assert.equal(rows[0].done, true);
  assert.equal(rows[1].status, 'moved');
  let back = g.get();
  assert.equal(back.tasks.find((t) => t.id === b.id).projectId, 'p1');
  assert.equal(back.tasks.length, 2);
  // Delete the Slack one.
  g.post(snapshot({ projects: back.projects, seenTaskIds, tasks: back.tasks.filter((t) => t.id !== a.id) }));
  rows = g.sheet('Slack Tasks (App)').objects();
  assert.equal(rows.length, 2);
  assert.equal(rows[0].status, 'deleted');
  back = g.get();
  assert.deepEqual(back.tasks.map((t) => t.id), [b.id]);
  // The source row must not be copied again after delete.
  g.get();
  assert.equal(g.sheet('Slack Tasks (App)').objects().length, 2);
});

test('meetings written by an external AI get real tasks once', () => {
  const g = createGas();
  g.get();
  const M = g.sheet('Meetings');
  M.data.push(['', 'Kickoff', '2026-10-05', 'A, B', 'Notes here', JSON.stringify([{ title: 'Send recap', owner: 'izabel arruda', due: '2026-10-07' }, { title: 'Book room', owner: 'Stranger' }]), '', '']);
  const r = g.get();
  g.get();
  const tasks = r.tasks.filter((t) => t.projectId === 'meetings');
  assert.equal(tasks.length, 2);
  assert.equal(tasks[0].assignee, 'Izabel Arruda');
  assert.equal(tasks[1].assignee, '');
  assert.deepEqual(tasks[0].tags, ['Meeting']);
  assert.match(tasks[0].description, /From meeting: Kickoff \(Oct 5\)/);
  assert.equal(r.meetings[0].actions.every((a) => a.taskId), true);
  assert.ok(r.meetings[0].id, 'meeting got an ID');
  assert.ok(r.projects.some((p) => p.id === 'meetings'));
  assert.equal(g.get().tasks.filter((t) => t.projectId === 'meetings').length, 2, 'no duplicates on reload');
});

test('Meetings tab is matched by heading and gets no new headings', () => {
  const g = createGas();
  g.sheet('Meetings').setRaw([['Title', 'ID', 'Notes', 'Date']]);
  g.get();
  g.post(snapshot({ meetings: [{ id: 'm1', title: 'T', date: '2026-10-01', attendees: 'x', notes: 'n', actions: '[]', fileName: 'f', createdAt: '' }] }));
  const sh = g.sheet('Meetings');
  assert.deepEqual(sh.data[0], ['Title', 'ID', 'Notes', 'Date']);
  assert.deepEqual(sh.objects()[0].Title, 'T');
  assert.deepEqual(sh.objects()[0].ID, 'm1');
});

test('Jira import: copies new tickets, filters assignee, throttles, marks resolved', () => {
  let searches = 0;
  const issues = [
    { key: 'L10N-1', fields: { summary: 'Translate', duedate: '2026-10-30', created: '2026-10-01T00:00:00.000+0000',
      reporter: { displayName: 'Rep' }, assignee: { displayName: 'Izabel Arruda' }, status: { name: 'To Do' }, priority: { name: 'High' },
      issuetype: { name: 'Task' }, labels: ['es'],
      description: { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Hello ' }, { type: 'text', text: 'bold', marks: [{ type: 'strong' }] }] }] } } },
    { key: 'L10N-2', fields: { summary: 'Other', assignee: { displayName: 'Not Listed' } } },
  ];
  let open = issues;
  const g = createGas({
    fetch: (url, opts) => {
      const body = JSON.parse(opts.payload || '{}');
      if (/search\/jql/.test(url)) {
        if (/^key in/.test(body.jql)) return { code: 200, body: { issues: [{ key: 'L10N-2', fields: { status: { statusCategory: { key: 'done' } } } }], isLast: true } };
        searches++;
        return { code: 200, body: { issues: open, isLast: true } };
      }
      return { code: 404, body: {} };
    },
  });
  g.props.set('JIRA_SITE', 'example.atlassian.net');
  g.props.set('JIRA_EMAIL', 'a@b.c');
  g.props.set('JIRA_API_TOKEN', 'tok');
  let r = g.get();
  const t = r.tasks.find((x) => x.title.startsWith('L10N-1'));
  assert.equal(t.title, 'L10N-1: Translate');
  assert.equal(t.projectId, 'jira');
  assert.equal(t.assignee, 'Izabel Arruda');
  assert.equal(t.reporter, 'Rep');
  assert.equal(t.due, '2026-10-30');
  assert.deepEqual(t.links, ['https://example.atlassian.net/browse/L10N-1']);
  assert.deepEqual(t.tags, ['Jira']);
  assert.match(t.description, /Hello <b>bold<\/b>/);
  assert.match(t.description, /Priority:<\/b> High/);
  assert.equal(r.tasks.find((x) => x.title.startsWith('L10N-2')).assignee, '');
  g.get();
  assert.equal(searches, 1, 'second load within 2 minutes does not search');
  open = [issues[0]];
  g.props.set('JIRA_LAST_FETCH', '0');
  r = g.get();
  assert.equal(searches, 2);
  assert.ok(!r.tasks.some((x) => x.title.startsWith('L10N-2')), 'resolved ticket hidden');
  assert.equal(g.sheet('Jira Tickets (App)').objects().find((x) => x.sourceId === 'L10N-2').status, 'resolved');
});

test('comments and people deletions only affect seen rows', () => {
  const g = createGas();
  g.get();
  g.post(snapshot({ projects: [project('p1')], tasks: [task('t1', 'p1')],
    comments: [{ id: 'c1', taskId: 't1', author: 'Izabel Arruda', text: 'a', createdAt: '' }],
    people: [{ name: 'Izabel Arruda', color: '#000' }, { name: 'Bo', color: '#111' }] }));
  g.sheet('Comments').data.push(['c2', 't1', 'Izabel Arruda', 'by hand', '']);
  g.sheet('People').data.push(['Hand Person', '#222']);
  g.post(snapshot({ projects: [project('p1')], tasks: [task('t1', 'p1')], comments: [],
    people: [{ name: 'Izabel Arruda', color: '#000' }],
    seenTaskIds: ['t1'], seenCommentIds: ['c1'], seenPeople: ['izabel arruda', 'bo'] }));
  const r = g.get();
  assert.deepEqual(r.comments.map((c) => c.id), ['c2']);
  assert.deepEqual(r.people.map((p) => p.name).sort(), ['Hand Person', 'Izabel Arruda']);
});

for (const [name, fn] of tests) {
  try {
    fn();
    passed++;
    console.log(`  ✓ ${name}`);
  } catch (err) {
    console.log(`  ✗ ${name}\n    ${err.stack.split('\n').slice(0, 4).join('\n    ')}`);
    process.exitCode = 1;
  }
}
console.log(`\n${passed}/${tests.length} passed`);

// Task Manager: state, rendering and events.
import 'iconify-icon';
import './style.css';
import { SHEETS_API_URL, loadSheet, saveSheet, type SheetResponse } from './sheet';
import {
  editorHtml,
  escapeHtml as h,
  installRichText,
  isEmptyHtml,
  normalizeLinkInput,
  safeUrl,
  sanitizeHtml,
  toDisplayHtml,
} from './richtext';
import { readTranscript, summariseTranscript, type MeetingDraft } from './meetings';

// ================================================================ types

type Kind = 'project' | 'notes';
type TaskState = '' | 'urgent' | 'progress';
type Status = 'todo' | 'urgent' | 'progress' | 'done';

interface Project { id: string; name: string; color: string; kind: Kind }
interface Task {
  id: string; projectId: string; title: string; done: boolean; assignee: string; due: string;
  tags: string[]; description: string; createdAt: string; reporter: string; links: string[]; state: TaskState;
}
interface Tag { name: string; color: string }
interface Person { name: string; color: string }
interface Comment { id: string; taskId: string; author: string; text: string; createdAt: string }
interface MeetingAction { title: string; owner: string; due: string; details: string; taskId: string }
interface Meeting {
  id: string; title: string; date: string; attendees: string[]; notes: string;
  actions: MeetingAction[]; fileName: string; createdAt: string;
}
interface Data {
  projects: Project[]; tasks: Task[]; tags: Tag[]; people: Person[]; comments: Comment[]; meetings: Meeting[];
}
interface Seen { tasks: Set<string>; projects: Set<string>; meetings: Set<string>; comments: Set<string>; people: Set<string> }

type View = { type: 'project' | 'tag' | 'person' | 'meeting'; id: string };
type ColKey = 'name' | 'assignee' | 'due' | 'tags' | 'project' | 'reporter' | 'meeting' | 'created';
type MenuKind = 'status' | 'assignee' | 'due' | 'tags' | 'project';
interface MenuState { kind: MenuKind; taskId: string; x: number; y: number; query: string }
type SyncStatus = 'loading' | 'saving' | 'saved' | 'error' | 'local';

// ================================================================ constants

const OWNER = 'Izabel Arruda';
const SLACK_ID = 'slack';
const JIRA_ID = 'jira';
const MEETINGS_ID = 'meetings';
const INBOX_ID = 'inbox';
const FEED_IDS = new Set([SLACK_ID, JIRA_ID]);
const BUILT_IN_IDS = new Set([SLACK_ID, JIRA_ID, MEETINGS_ID]);

const TAG_PALETTE = [
  { bg: '#ECE1FF', fg: '#5B2BB5' },
  { bg: '#DCEFFF', fg: '#1E5F8A' },
  { bg: '#D7F5E8', fg: '#1F7A55' },
  { bg: '#FFF2CC', fg: '#8A6400' },
  { bg: '#FFE3D9', fg: '#A6432A' },
  { bg: '#FCE0EE', fg: '#9C2F68' },
  { bg: '#D8F3F2', fg: '#1D6F72' },
  { bg: '#EAEAEE', fg: '#4A4A55' },
];
const PROJECT_COLORS = ['#8B3DFF', '#F06A6A', '#F1A43C', '#2FB36B', '#3D8BFF', '#E362B4', '#14B8A6', '#6366F1', '#64748B'];
const PERSON_COLORS = ['#8B3DFF', '#3D8BFF', '#14B8A6', '#F1A43C', '#F06A6A', '#E362B4', '#2FB36B', '#6366F1'];

const STATUS_LABEL: Record<Status, string> = { todo: 'To do', urgent: 'Urgent', progress: 'In progress', done: 'Done' };
const COL_LABEL: Record<ColKey, string> = {
  name: 'Task name', assignee: 'Assignee', due: 'Due date', tags: 'Tags', project: 'Project',
  reporter: 'Reporter', meeting: 'Meeting', created: 'Created',
};
const DEFAULT_W: Record<ColKey, number> = {
  name: 380, assignee: 160, due: 110, tags: 190, project: 160, reporter: 160, meeting: 210, created: 100,
};
const MIN_W = 60;
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

const SAVE_DEBOUNCE_MS = 800;
const POLL_MS = 60_000;

const LS = {
  cache: 'tm.cache.v1',
  view: 'tm.view',
  collapsed: 'tm.collapsed',
  colw: 'tm.colw',
};

const TICK_SVG = '<svg viewBox="0 0 12 12" width="10" height="10" aria-hidden="true"><path d="M2.5 6.3l2.3 2.2 4.7-5" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>';

// ================================================================ small utils

const ic = (name: string, cls = ''): string => `<iconify-icon icon="lucide:${name}"${cls ? ` class="${cls}"` : ''}></iconify-icon>`;

function uid(prefix: string): string {
  const rnd = typeof crypto !== 'undefined' && 'randomUUID' in crypto
    ? crypto.randomUUID().replace(/-/g, '').slice(0, 16)
    : Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
  return `${prefix}-${rnd}`;
}

function lsGet<T>(key: string, fallback: T): T {
  try {
    const raw = localStorage.getItem(key);
    return raw == null ? fallback : (JSON.parse(raw) as T);
  } catch {
    return fallback;
  }
}
function lsSet(key: string, value: unknown): void {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch {
    /* storage full or blocked: the sheet is the real store */
  }
}

const pad = (n: number): string => String(n).padStart(2, '0');
function ymd(d: Date): string {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}
function parseYmd(s: string): Date | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s || '');
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
}
const todayYmd = (): string => ymd(new Date());
function addDays(base: string, n: number): string {
  const d = parseYmd(base) || new Date();
  d.setDate(d.getDate() + n);
  return ymd(d);
}
function shortDate(d: Date): string {
  const s = `${MONTHS[d.getMonth()]} ${d.getDate()}`;
  return d.getFullYear() === new Date().getFullYear() ? s : `${s}, ${d.getFullYear()}`;
}
function fmtDue(due: string, done: boolean): { text: string; cls: string } {
  const d = parseYmd(due);
  if (!d) return { text: '', cls: '' };
  const t = todayYmd();
  if (due === t) return { text: 'Today', cls: done ? '' : 'due-soon' };
  if (due === addDays(t, 1)) return { text: 'Tomorrow', cls: done ? '' : 'due-soon' };
  if (due === addDays(t, -1)) return { text: 'Yesterday', cls: done ? '' : 'due-over' };
  return { text: shortDate(d), cls: !done && due < t ? 'due-over' : '' };
}
function isoToDate(iso: string): Date | null {
  if (!iso) return null;
  const d = parseYmd(iso) && iso.length <= 10 ? parseYmd(iso) : new Date(iso);
  return d && !isNaN(d.getTime()) ? d : null;
}
function fmtCreated(iso: string): string {
  const d = isoToDate(iso);
  return d ? shortDate(d) : '';
}
function fmtDateTime(iso: string): string {
  const d = isoToDate(iso);
  if (!d) return '—';
  return `${shortDate(d)} at ${d.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })}`;
}
function timeAgo(iso: string): string {
  const d = isoToDate(iso);
  if (!d) return '';
  const s = Math.round((Date.now() - d.getTime()) / 1000);
  if (s < 45) return 'just now';
  const units: [number, string][] = [[60, 'minute'], [3600, 'hour'], [86400, 'day'], [604800, 'week'], [2629800, 'month'], [31557600, 'year']];
  let label = '';
  for (let i = units.length - 1; i >= 0; i--) {
    if (s >= units[i][0]) {
      const n = Math.floor(s / units[i][0]);
      label = `${n} ${units[i][1]}${n === 1 ? '' : 's'} ago`;
      break;
    }
  }
  return label || 'just now';
}
function fmtMeetingDate(date: string): string {
  const d = parseYmd(date);
  return d ? shortDate(d) : '';
}
function initials(name: string): string {
  const parts = name.replace(/\(.*?\)/g, '').trim().split(/\s+/).filter(Boolean);
  return ((parts[0]?.[0] || '?') + (parts.length > 1 ? parts[parts.length - 1][0] : '')).toUpperCase();
}
const eqName = (a: string, b: string): boolean => a.trim().toLowerCase() === b.trim().toLowerCase();

// ================================================================ state

const emptyData = (): Data => ({ projects: [], tasks: [], tags: [], people: [], comments: [], meetings: [] });
const emptySeen = (): Seen => ({ tasks: new Set(), projects: new Set(), meetings: new Set(), comments: new Set(), people: new Set() });

let data: Data = emptyData();
let seen: Seen = emptySeen();

const ui = {
  view: null as View | null,
  selectedId: null as string | null,
  collapsed: { projects: false, notes: false, meetings: false, tags: true, people: true } as Record<string, boolean>,
  colW: {} as Partial<Record<ColKey, number>>,
  menu: null as MenuState | null,
  dialogOpen: false,
  adding: null as null | 'project' | 'notebook' | 'person',
  aiBusy: null as string | null,
  sync: 'loading' as SyncStatus,
  syncError: '',
};

const commentDrafts = new Map<string, string>();
let draggingTaskId: string | null = null;

// ---------------------------------------------------------------- normalising

const str = (v: unknown): string => (v == null ? '' : String(v));
const toBool = (v: unknown): boolean => v === true || /^(true|yes|1|x|done)$/i.test(str(v).trim());
function toList(v: unknown, sep: RegExp): string[] {
  if (Array.isArray(v)) return v.map((x) => str(x).trim()).filter(Boolean);
  return str(v).split(sep).map((x) => x.trim()).filter(Boolean);
}
function toState(v: unknown): TaskState {
  const s = str(v).toLowerCase();
  return s === 'urgent' ? 'urgent' : s === 'progress' || s === 'in progress' ? 'progress' : '';
}

function normProject(p: any, prev?: Project): Project {
  const kind = p?.kind === undefined || p?.kind === '' ? prev?.kind ?? 'project' : p.kind === 'notes' ? 'notes' : 'project';
  return { id: str(p?.id), name: str(p?.name) || 'Untitled', color: str(p?.color) || PROJECT_COLORS[0], kind };
}
function normTask(t: any, prev?: Task): Task {
  return {
    id: str(t?.id),
    projectId: str(t?.projectId),
    title: str(t?.title),
    done: toBool(t?.done),
    assignee: str(t?.assignee).trim(),
    due: /^\d{4}-\d{2}-\d{2}/.test(str(t?.due)) ? str(t.due).slice(0, 10) : '',
    tags: toList(t?.tags, /,/),
    description: str(t?.description),
    createdAt: str(t?.createdAt),
    reporter: str(t?.reporter),
    links: toList(t?.links, /[\n\r]+/).filter((l) => safeUrl(l)),
    state: t?.state === undefined ? prev?.state ?? '' : toState(t.state),
  };
}
function normComment(c: any): Comment {
  return { id: str(c?.id), taskId: str(c?.taskId), author: str(c?.author) || OWNER, text: str(c?.text), createdAt: str(c?.createdAt) };
}
function normMeeting(m: any): Meeting {
  let actions: any = m?.actions;
  if (typeof actions === 'string') {
    try {
      actions = JSON.parse(actions);
    } catch {
      actions = [];
    }
  }
  return {
    id: str(m?.id),
    title: str(m?.title) || 'Meeting',
    date: /^\d{4}-\d{2}-\d{2}/.test(str(m?.date)) ? str(m.date).slice(0, 10) : '',
    attendees: toList(m?.attendees, /[,;\n]+/),
    notes: str(m?.notes),
    actions: (Array.isArray(actions) ? actions : []).map((a: any) => ({
      title: str(a?.title), owner: str(a?.owner), due: str(a?.due), details: str(a?.details), taskId: str(a?.taskId),
    })),
    fileName: str(m?.fileName),
    createdAt: str(m?.createdAt),
  };
}
const normPerson = (p: any): Person => ({ name: str(p?.name).trim(), color: str(p?.color) || PERSON_COLORS[0] });
const normTag = (t: any): Tag => ({ name: str(t?.name).trim(), color: str(t?.color) || TAG_PALETTE[0].bg });

function ensureOwner(): void {
  const i = data.people.findIndex((p) => eqName(p.name, OWNER));
  if (i === 0) return;
  const owner = i > 0 ? data.people.splice(i, 1)[0] : { name: OWNER, color: PERSON_COLORS[0] };
  data.people.unshift(owner);
}

/** Makes sure every tag used on a task has a colour entry. */
function ensureTagColors(): void {
  for (const t of data.tasks) for (const name of t.tags) ensureTag(name);
}

// ---------------------------------------------------------------- cache

function persistCache(): void {
  lsSet(LS.cache, {
    data,
    seen: {
      tasks: [...seen.tasks], projects: [...seen.projects], meetings: [...seen.meetings],
      comments: [...seen.comments], people: [...seen.people],
    },
  });
}

function readCache(): boolean {
  try {
    const raw = localStorage.getItem(LS.cache);
    if (!raw) return false;
    const c = JSON.parse(raw);
    if (!c || !c.data || !Array.isArray(c.data.tasks)) return false;
    data = {
      projects: (c.data.projects || []).map((p: any) => normProject(p)),
      tasks: (c.data.tasks || []).map((t: any) => normTask(t)),
      tags: (c.data.tags || []).map(normTag),
      people: (c.data.people || []).map(normPerson),
      comments: (c.data.comments || []).map(normComment),
      meetings: (c.data.meetings || []).map(normMeeting),
    };
    const s = c.seen || {};
    seen = {
      tasks: new Set(s.tasks || []), projects: new Set(s.projects || []), meetings: new Set(s.meetings || []),
      comments: new Set(s.comments || []), people: new Set(s.people || []),
    };
    return true;
  } catch {
    // Can't read the saved copy: start empty and load from the sheet.
    data = emptyData();
    seen = emptySeen();
    return false;
  }
}

// ================================================================ derived helpers

const projectById = (id: string): Project | undefined => data.projects.find((p) => p.id === id);
const taskById = (id: string): Task | undefined => data.tasks.find((t) => t.id === id);
const meetingById = (id: string): Meeting | undefined => data.meetings.find((m) => m.id === id);
const kindOfTask = (t: Task): Kind => projectById(t.projectId)?.kind ?? 'project';
const isNote = (t: Task): boolean => kindOfTask(t) === 'notes';

function statusOf(t: Task): Status {
  if (t.done) return 'done';
  return t.state === 'urgent' ? 'urgent' : t.state === 'progress' ? 'progress' : 'todo';
}

function meetingOfTask(taskId: string): Meeting | undefined {
  return data.meetings.find((m) => m.actions.some((a) => a.taskId === taskId));
}

function tagStyle(name: string): string {
  const tag = data.tags.find((t) => eqName(t.name, name));
  const pal = TAG_PALETTE.find((p) => p.bg.toLowerCase() === (tag?.color || '').toLowerCase()) || TAG_PALETTE[7];
  return `background:${pal.bg};color:${pal.fg}`;
}

function ensureTag(name: string): string {
  const existing = data.tags.find((t) => eqName(t.name, name));
  if (existing) return existing.name;
  data.tags.push({ name, color: TAG_PALETTE[data.tags.length % TAG_PALETTE.length].bg });
  return name;
}

function personColor(name: string): string {
  return data.people.find((p) => eqName(p.name, name))?.color || '#9CA0AA';
}
const displayName = (name: string): string => (eqName(name, OWNER) ? `${OWNER} (you)` : name);

/** Finds a person on the Assignees list by full name, or by a first name only one person has. */
function matchPerson(name: string): string {
  const n = name.trim();
  if (!n) return '';
  const exact = data.people.find((p) => eqName(p.name, n));
  if (exact) return exact.name;
  const first = data.people.filter((p) => eqName(p.name.split(/\s+/)[0], n.split(/\s+/)[0]));
  return first.length === 1 && n.split(/\s+/).length === 1 ? first[0].name : '';
}

function tagsInUse(): { name: string; open: number }[] {
  const counts = new Map<string, { name: string; open: number }>();
  for (const t of data.tasks) {
    for (const name of t.tags) {
      const key = name.toLowerCase();
      const c = counts.get(key) || { name, open: 0 };
      if (!t.done && !isNote(t)) c.open++;
      counts.set(key, c);
    }
  }
  const order = data.tags.map((t) => t.name.toLowerCase());
  return [...counts.entries()]
    .sort((a, b) => {
      const ia = order.indexOf(a[0]);
      const ib = order.indexOf(b[0]);
      return (ia < 0 ? 1e9 : ia) - (ib < 0 ? 1e9 : ib) || a[0].localeCompare(b[0]);
    })
    .map(([, v]) => v);
}

function openActionCount(m: Meeting): number {
  return m.actions.filter((a) => {
    const t = taskById(a.taskId);
    return t && !t.done;
  }).length;
}

function firstTaskProjectId(): string {
  const p = data.projects.find((x) => x.kind === 'project' && !BUILT_IN_IDS.has(x.id));
  if (p) return p.id;
  return ensureProject(INBOX_ID, 'Inbox', '#64748B').id;
}

function ensureProject(id: string, name: string, color: string): Project {
  let p = projectById(id);
  if (!p) {
    p = { id, name, color, kind: 'project' };
    data.projects.push(p);
  }
  return p;
}

// ================================================================ sync engine

let localVersion = 0; // bumps on every local change
let syncedVersion = 0; // the version the last successful save included
let saveGen = 0; // bumps when a save lands
let saveTimer: number | null = null;
let saving = false;
let loading = false;
let loadedOnce = false;
let retryDelay = 0;
let retryTimer: number | null = null;

const hasUnsaved = (): boolean => localVersion !== syncedVersion;

function setSync(s: SyncStatus, err = ''): void {
  ui.sync = s;
  ui.syncError = err;
  const el = document.getElementById('sync-status');
  if (el) el.outerHTML = syncHtml();
}

function syncHtml(): string {
  const text: Record<SyncStatus, string> = {
    loading: 'Loading…',
    saving: 'Saving…',
    saved: 'Saved to Google Sheet',
    error: "Can't reach the sheet, retrying",
    local: 'Saved in this browser only',
  };
  const icon: Record<SyncStatus, string> = {
    loading: 'loader-circle', saving: 'loader-circle', saved: 'cloud-check', error: 'cloud-off', local: 'hard-drive',
  };
  const spin = ui.sync === 'loading' || ui.sync === 'saving' ? ' spin' : '';
  const title = ui.sync === 'local' ? 'Add SHEETS_API_URL in src/sheet.ts to sync with Google Sheets' : ui.syncError;
  return `<div class="sb-sync sync-${ui.sync}" id="sync-status" title="${h(title)}">${ic(icon[ui.sync], spin)}<span>${text[ui.sync]}</span></div>`;
}

/** Record a local change: cache it, schedule a save and (by default) redraw. */
function commit(opts: { render?: boolean } = {}): void {
  localVersion++;
  persistCache();
  scheduleSave();
  if (opts.render !== false) render();
}

function scheduleSave(delay = SAVE_DEBOUNCE_MS): void {
  if (!SHEETS_API_URL) {
    setSync('local');
    return;
  }
  if (!loadedOnce) return; // never send anything before the sheet has been read
  if (retryTimer) return; // a retry is already queued and will pick up the latest snapshot
  if (saveTimer) clearTimeout(saveTimer);
  setSync('saving');
  saveTimer = window.setTimeout(() => {
    saveTimer = null;
    void runSave();
  }, delay);
}

function buildPayload() {
  return {
    projects: data.projects,
    tasks: data.tasks,
    tags: data.tags,
    people: data.people,
    comments: data.comments,
    meetings: data.meetings.map((m) => ({ ...m, attendees: m.attendees.join(', '), actions: JSON.stringify(m.actions) })),
    seenTaskIds: [...seen.tasks],
    seenProjectIds: [...seen.projects],
    seenMeetingIds: [...seen.meetings],
    seenCommentIds: [...seen.comments],
    seenPeople: [...seen.people],
  };
}

async function runSave(): Promise<void> {
  if (saving) return; // the finally block below sends whatever changed meanwhile
  if (!hasUnsaved()) {
    setSync('saved');
    return;
  }
  saving = true;
  const version = localVersion;
  const sent = {
    tasks: data.tasks.map((t) => t.id), projects: data.projects.map((p) => p.id), meetings: data.meetings.map((m) => m.id),
    comments: data.comments.map((c) => c.id), people: data.people.map((p) => p.name.toLowerCase()),
  };
  let failed = false;
  try {
    await saveSheet(buildPayload());
    syncedVersion = version;
    saveGen++;
    retryDelay = 0;
    // The sheet now has these rows, so a later local delete should delete them there too.
    sent.tasks.forEach((id) => seen.tasks.add(id));
    sent.projects.forEach((id) => seen.projects.add(id));
    sent.meetings.forEach((id) => seen.meetings.add(id));
    sent.comments.forEach((id) => seen.comments.add(id));
    sent.people.forEach((n) => seen.people.add(n));
    persistCache();
  } catch (err) {
    failed = true;
    retryDelay = Math.min(retryDelay ? retryDelay * 2 : 2000, 60000);
    setSync('error', String((err as Error)?.message || err));
    retryTimer = window.setTimeout(() => {
      retryTimer = null;
      void runSave();
    }, retryDelay);
  } finally {
    saving = false;
  }
  if (!failed) {
    if (hasUnsaved()) scheduleSave(0); // edits made while this save was in flight
    else setSync('saved');
  }
}

function isTyping(): boolean {
  const a = document.activeElement as HTMLElement | null;
  if (!a || a === document.body) return false;
  if (a.isContentEditable) return true;
  if (a instanceof HTMLTextAreaElement) return true;
  if (a instanceof HTMLInputElement) return !['button', 'checkbox', 'radio', 'file', 'submit', 'reset'].includes(a.type);
  return a instanceof HTMLSelectElement;
}

const busyForLoad = (): boolean =>
  isTyping() || !!ui.menu || ui.dialogOpen || saving || saveTimer !== null || retryTimer !== null || hasUnsaved();

async function runLoad(): Promise<void> {
  if (!SHEETS_API_URL || loading) return;
  if (loadedOnce && busyForLoad()) return;
  loading = true;
  if (!loadedOnce) setSync('loading');
  const v0 = localVersion;
  const g0 = saveGen;
  try {
    const remote = await loadSheet();
    // Typing or a menu opened while we were fetching: drop this copy rather than
    // redraw under the person. Nothing is marked seen, so nothing is lost.
    if (loadedOnce && (isTyping() || ui.menu || ui.dialogOpen)) return;
    const merge = localVersion !== v0 || saveGen !== g0 || hasUnsaved() || saving;
    const needsSave = applyRemote(remote, merge);
    const first = !loadedOnce;
    loadedOnce = true;
    retryDelay = 0;
    persistCache();
    if (needsSave) localVersion++;
    const pending = needsSave || hasUnsaved();
    if (!pending && !saving) ui.sync = 'saved';
    render();
    if (pending) scheduleSave(first ? 0 : SAVE_DEBOUNCE_MS);
  } catch (err) {
    setSync('error', String((err as Error)?.message || err));
    if (!loadedOnce) {
      render();
      retryDelay = Math.min(retryDelay ? retryDelay * 2 : 3000, 60000);
      window.setTimeout(() => void runLoad(), retryDelay);
    }
  } finally {
    loading = false;
  }
}

/**
 * Applies the sheet's copy. In merge mode (there are unsaved local changes)
 * local rows win and only rows the browser never had are added. Otherwise
 * the sheet wins, except for local rows that were never synced.
 * Returns true when the browser holds rows the sheet still needs.
 */
function applyRemote(remote: SheetResponse, merge: boolean): boolean {
  let needsSave = false;

  function mergeById<T extends { id: string }>(
    local: T[], remoteRaw: unknown[] | undefined, seenSet: Set<string>, norm: (r: any, prev?: T) => T,
  ): T[] {
    if (remoteRaw === undefined) {
      // Older script without this field: keep the browser's copy and refill the tab.
      if (local.length) needsSave = true;
      return local;
    }
    const localById = new Map(local.map((x) => [x.id, x]));
    const rows = remoteRaw.map((r: any) => norm(r, localById.get(str(r?.id)))).filter((r) => r.id);
    const remoteIds = new Set(rows.map((r) => r.id));
    let next: T[];
    if (merge) {
      next = local.slice();
      for (const r of rows) if (!localById.has(r.id) && !seenSet.has(r.id)) next.push(r);
    } else {
      next = rows;
      for (const l of local) {
        if (!remoteIds.has(l.id) && !seenSet.has(l.id)) {
          next.push(l); // never synced yet
          needsSave = true;
        }
      }
    }
    rows.forEach((r) => seenSet.add(r.id));
    return next;
  }

  data.projects = mergeById(data.projects, remote.projects, seen.projects, normProject);
  data.tasks = mergeById(data.tasks, remote.tasks, seen.tasks, normTask);
  data.comments = mergeById(data.comments, remote.comments, seen.comments, normComment);
  data.meetings = mergeById(data.meetings, remote.meetings, seen.meetings, normMeeting);

  if (remote.people === undefined) {
    if (data.people.length) needsSave = true;
  } else {
    const rows = remote.people.map(normPerson).filter((p) => p.name);
    const has = (list: Person[], n: string) => list.some((p) => eqName(p.name, n));
    let next: Person[];
    if (merge) {
      next = data.people.slice();
      for (const r of rows) if (!has(next, r.name) && !seen.people.has(r.name.toLowerCase())) next.push(r);
    } else {
      next = rows;
      for (const l of data.people) {
        if (!has(next, l.name) && !seen.people.has(l.name.toLowerCase())) {
          next.push(l);
          needsSave = true;
        }
      }
    }
    rows.forEach((r) => seen.people.add(r.name.toLowerCase()));
    data.people = next;
  }

  if (remote.tags !== undefined) {
    for (const r of remote.tags.map(normTag).filter((t) => t.name)) {
      const l = data.tags.find((t) => eqName(t.name, r.name));
      if (!l) data.tags.push(r);
      else if (!merge) l.color = r.color;
    }
  }

  ensureOwner();
  ensureTagColors();
  return needsSave;
}

// ================================================================ mutations

function newTask(fields: Partial<Task> & { projectId: string }): Task {
  const t: Task = {
    id: uid('t'), title: '', done: false, assignee: '', due: '', tags: [], description: '',
    createdAt: new Date().toISOString(), reporter: '', links: [], state: '', ...fields,
  };
  t.tags.forEach(ensureTag);
  data.tasks.push(t);
  return t;
}

function setStatus(t: Task, s: Status): void {
  if (s === 'done') {
    t.done = true;
  } else {
    t.done = false;
    t.state = s === 'todo' ? '' : s;
  }
}

function deleteTasks(ids: Set<string>): void {
  data.tasks = data.tasks.filter((t) => !ids.has(t.id));
  data.comments = data.comments.filter((c) => !ids.has(c.taskId));
  for (const m of data.meetings) m.actions = m.actions.filter((a) => !ids.has(a.taskId));
  if (ui.selectedId && ids.has(ui.selectedId)) ui.selectedId = null;
}

function moveTask(t: Task, projectId: string): boolean {
  const target = projectById(projectId);
  if (!target || target.kind !== kindOfTask(t) || t.projectId === projectId) return false;
  t.projectId = projectId;
  return true;
}

function createMeetingFromDraft(draft: MeetingDraft, fileName: string): Meeting {
  const m: Meeting = {
    id: uid('m'), title: draft.title, date: draft.date || todayYmd(), attendees: draft.attendees,
    notes: sanitizeHtml(draft.notesHtml), actions: [], fileName, createdAt: new Date().toISOString(),
  };
  data.meetings.push(m);
  for (const a of draft.actions) addMeetingAction(m, a.title, a.owner, a.due, a.details);
  return m;
}

function addMeetingAction(m: Meeting, title: string, owner: string, due: string, details: string): Task {
  ensureProject(MEETINGS_ID, 'Meetings', '#14B8A6');
  const when = m.date ? ` (${fmtMeetingDate(m.date)})` : '';
  const t = newTask({
    projectId: MEETINGS_ID,
    title,
    assignee: matchPerson(owner),
    due: /^\d{4}-\d{2}-\d{2}$/.test(due) ? due : '',
    tags: [ensureTag('Meeting')],
    description: `<p>From meeting: ${h(m.title)}${h(when)}</p>${details ? `<p>${h(details)}</p>` : ''}`,
  });
  m.actions.push({ title, owner, due: t.due, details, taskId: t.id });
  return t;
}

// ================================================================ dialogs (built in; confirm() is blocked in embedded previews)

let dialogResolve: ((ok: boolean) => void) | null = null;

function confirmDialog(o: { title: string; message: string; confirm?: string; cancel?: string | null; danger?: boolean }): Promise<boolean> {
  closeMenu();
  return new Promise((resolve) => {
    dialogResolve?.(false);
    dialogResolve = resolve;
    ui.dialogOpen = true;
    const root = document.getElementById('dialog-root')!;
    root.innerHTML = `<div class="dialog-backdrop" data-dialog="cancel">
      <div class="dialog" role="dialog" aria-modal="true" aria-labelledby="dialog-title">
        <h2 id="dialog-title">${h(o.title)}</h2>
        <p>${h(o.message)}</p>
        <div class="dialog-actions">
          ${o.cancel === null ? '' : `<button class="btn" data-dialog="cancel">${h(o.cancel || 'Cancel')}</button>`}
          <button class="btn ${o.danger === false ? 'btn-primary' : 'btn-danger'}" data-dialog="ok">${h(o.confirm || 'Delete')}</button>
        </div>
      </div>
    </div>`;
    requestAnimationFrame(() => root.querySelector<HTMLElement>('[data-dialog="ok"]')?.focus());
  });
}

function alertDialog(title: string, message: string): Promise<boolean> {
  return confirmDialog({ title, message, confirm: 'OK', cancel: null, danger: false });
}

function closeDialog(ok: boolean): void {
  const r = dialogResolve;
  dialogResolve = null;
  ui.dialogOpen = false;
  document.getElementById('dialog-root')!.innerHTML = '';
  r?.(ok);
}

// ================================================================ rendering: shell

function currentView(): View | null {
  const v = ui.view;
  if (v) {
    if (v.type === 'project' && projectById(v.id)) return v;
    if (v.type === 'meeting' && meetingById(v.id)) return v;
    if (v.type === 'person' && data.people.some((p) => eqName(p.name, v.id))) return v;
    if (v.type === 'tag' && tagsInUse().some((t) => eqName(t.name, v.id))) return v;
  }
  const p = data.projects.find((x) => x.kind === 'project') || data.projects[0];
  return p ? { type: 'project', id: p.id } : null;
}

interface FocusSnap { fk: string; value?: string; start?: number | null; end?: number | null; keep: boolean }

function captureFocus(): FocusSnap | null {
  const a = document.activeElement as HTMLElement | null;
  if (!a || !a.dataset?.fk) return null;
  const snap: FocusSnap = { fk: a.dataset.fk, keep: a.hasAttribute('data-keep') };
  if (a instanceof HTMLInputElement || a instanceof HTMLTextAreaElement) {
    snap.value = a.value;
    try {
      snap.start = a.selectionStart;
      snap.end = a.selectionEnd;
    } catch {
      /* not a text input */
    }
  }
  return snap;
}

function restoreFocus(snap: FocusSnap | null): void {
  if (!snap) return;
  const el = document.querySelector<HTMLElement>(`[data-fk="${CSS.escape(snap.fk)}"]`);
  if (!el) return;
  el.focus({ preventScroll: true });
  if (el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement) {
    if (snap.keep && snap.value !== undefined) el.value = snap.value;
    try {
      if (snap.start != null) el.setSelectionRange(snap.start, snap.end ?? snap.start);
    } catch {
      /* not a text input */
    }
  } else if (el.isContentEditable) {
    const r = document.createRange();
    r.selectNodeContents(el);
    r.collapse(false);
    const sel = window.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(r);
  }
}

function render(): void {
  const focus = captureFocus();
  const scrolls = ['.list-scroll', '.panel-body', '.sb-scroll', '.meeting-scroll'].map((sel) => {
    const el = document.querySelector(sel);
    return { sel, top: el?.scrollTop ?? 0, left: el?.scrollLeft ?? 0 };
  });

  if (ui.selectedId && !taskById(ui.selectedId)) ui.selectedId = null;

  document.getElementById('sidebar')!.innerHTML = renderSidebar();
  document.getElementById('main')!.innerHTML = renderMain();
  const panel = document.getElementById('panel')!;
  panel.innerHTML = renderPanel();
  document.getElementById('app')!.classList.toggle('panel-open', !!ui.selectedId);

  for (const s of scrolls) {
    const el = document.querySelector(s.sel);
    if (el) {
      el.scrollTop = s.top;
      el.scrollLeft = s.left;
    }
  }
  restoreFocus(focus);
  autoGrowAll();
  if (ui.menu) renderMenu();
}

function autoGrowAll(): void {
  document.querySelectorAll<HTMLTextAreaElement>('textarea.autogrow').forEach(autoGrow);
}
function autoGrow(el: HTMLTextAreaElement): void {
  el.style.height = 'auto';
  el.style.height = `${el.scrollHeight}px`;
}

// ================================================================ rendering: sidebar

function sbSection(key: string, title: string, addAction: string | null, addTitle: string, body: string): string {
  const collapsed = !!ui.collapsed[key];
  return `<section class="sb-section${collapsed ? ' collapsed' : ''}">
    <div class="sb-head">
      <button class="sb-toggle" data-action="toggle-section" data-key="${key}" aria-expanded="${!collapsed}">
        ${ic('chevron-down', 'chev')}<span>${h(title)}</span>
      </button>
      ${addAction ? `<button class="sb-add" data-action="${addAction}" title="${h(addTitle)}" aria-label="${h(addTitle)}">${ic('plus')}</button>` : ''}
    </div>
    ${collapsed ? '' : `<div class="sb-body">${body}</div>`}
  </section>`;
}

function sbAddRow(kind: 'project' | 'notebook' | 'person', placeholder: string): string {
  if (ui.adding !== kind) return '';
  return `<div class="sb-adding"><input class="sb-input" data-keep data-fk="add-${kind}" data-action-key="add-${kind}" placeholder="${h(placeholder)}" /></div>`;
}

function renderSidebar(): string {
  const v = currentView();
  const active = (type: string, id: string) => (v && v.type === type && (type === 'person' || type === 'tag' ? eqName(v.id, id) : v.id === id) ? ' active' : '');

  const projectItem = (p: Project) => {
    const open = data.tasks.filter((t) => t.projectId === p.id && (p.kind === 'notes' || !t.done)).length;
    const icon = p.id === SLACK_ID ? ic('hash') : p.id === JIRA_ID ? ic('ticket') : p.id === MEETINGS_ID ? ic('sparkles') : p.kind === 'notes' ? ic('notebook') : '';
    return `<button class="sb-item${active('project', p.id)}" data-action="nav" data-type="project" data-id="${h(p.id)}" data-drop-project="${h(p.id)}">
      <span class="sb-dot" style="background:${h(p.color)}">${icon}</span><span class="sb-name">${h(p.name)}</span>${open ? `<span class="sb-count">${open}</span>` : ''}
    </button>`;
  };

  const projects = data.projects.filter((p) => p.kind === 'project');
  const notebooks = data.projects.filter((p) => p.kind === 'notes');
  const meetings = data.meetings.slice().sort((a, b) => (b.date || b.createdAt).localeCompare(a.date || a.createdAt));
  const tags = tagsInUse();

  const projectsBody = projects.map(projectItem).join('') + sbAddRow('project', 'Project name') +
    (!projects.length && ui.adding !== 'project' ? `<p class="sb-empty">No projects yet. Use + to add one.</p>` : '');
  const notesBody = notebooks.map(projectItem).join('') + sbAddRow('notebook', 'Notebook name') +
    (!notebooks.length && ui.adding !== 'notebook' ? `<p class="sb-empty">Keep notes in notebooks. Use + to add one.</p>` : '');
  const meetingsBody =
    (ui.aiBusy ? `<div class="sb-item ai-busy">${ic('sparkles', 'sparkle-pulse')}<span class="sb-name">Summarising ${h(ui.aiBusy)}…</span></div>` : '') +
    meetings.map((m) => {
      const open = openActionCount(m);
      return `<button class="sb-item${active('meeting', m.id)}" data-action="nav" data-type="meeting" data-id="${h(m.id)}">
        <span class="sb-dot meeting-dot">${ic('sparkles')}</span><span class="sb-name">${h(m.title)}</span>${open ? `<span class="sb-count">${open}</span>` : ''}
      </button>`;
    }).join('') +
    (!meetings.length && !ui.aiBusy ? `<p class="sb-empty">Upload a transcript with +, or let your automation fill the Meetings tab.</p>` : '');
  const tagsBody = tags.length
    ? tags.map((t) => `<button class="sb-item${active('tag', t.name)}" data-action="nav" data-type="tag" data-id="${h(t.name)}">
        <span class="sb-tagdot" style="${tagStyle(t.name)}"></span><span class="sb-name">${h(t.name)}</span>${t.open ? `<span class="sb-count">${t.open}</span>` : ''}
      </button>`).join('')
    : `<p class="sb-empty">Tags you add to tasks show up here.</p>`;
  const peopleBody = data.people.map((p) => {
    const open = data.tasks.filter((t) => eqName(t.assignee, p.name) && !t.done && !isNote(t)).length;
    return `<button class="sb-item${active('person', p.name)}" data-action="nav" data-type="person" data-id="${h(p.name)}">
      ${avatar(p.name, 'sm')}<span class="sb-name">${h(displayName(p.name))}</span>${open ? `<span class="sb-count">${open}</span>` : ''}
    </button>`;
  }).join('') + sbAddRow('person', 'Name');

  return `<div class="sb-brand"><span class="sb-logo">${ic('check-check')}</span><span>Task Manager</span></div>
    <nav class="sb-scroll">
      ${sbSection('projects', 'Projects', 'add-project', 'Add project', projectsBody)}
      ${sbSection('notes', 'Notes', 'add-notebook', 'Add notebook', notesBody)}
      ${sbSection('meetings', 'Meetings', 'upload-transcript', 'Upload a transcript', meetingsBody)}
      ${sbSection('tags', 'Tags', null, '', tagsBody)}
      ${sbSection('people', 'Assignees', 'add-person', 'Add person', peopleBody)}
    </nav>
    ${syncHtml()}`;
}

function avatar(name: string, size: 'sm' | 'md' = 'md'): string {
  if (!name) return `<span class="avatar avatar-${size} avatar-empty">${ic('user')}</span>`;
  return `<span class="avatar avatar-${size}" style="background:${h(personColor(name))}" title="${h(displayName(name))}">${h(initials(name))}</span>`;
}

// ================================================================ rendering: main

function emptyState(icon: string, title: string, text: string, extra = ''): string {
  return `<div class="empty"><div class="empty-icon">${ic(icon)}</div><h3>${h(title)}</h3><p>${h(text)}</p>${extra}</div>`;
}

function renderMain(): string {
  const v = currentView();
  if (!v) {
    const loadingNow = !!SHEETS_API_URL && !loadedOnce && ui.sync !== 'error';
    return `<div class="main-inner">${loadingNow
      ? emptyState('loader-circle', 'Loading your tasks…', 'Reading the Google Sheet.')
      : emptyState('folder-plus', 'Create your first project', 'Use the + next to Projects in the sidebar, then add tasks to it.',
        `<button class="btn btn-primary" data-action="add-project">${ic('plus')} New project</button>`)}</div>`;
  }
  if (v.type === 'meeting') return renderMeeting(meetingById(v.id)!);

  let tasks: Task[];
  let header: string;
  let notes = false;
  let cols: ColKey[];
  if (v.type === 'project') {
    const p = projectById(v.id)!;
    notes = p.kind === 'notes';
    tasks = data.tasks.filter((t) => t.projectId === p.id);
    const icon = p.id === SLACK_ID ? 'hash' : p.id === JIRA_ID ? 'ticket' : p.id === MEETINGS_ID ? 'sparkles' : notes ? 'notebook' : 'list-checks';
    const tileCls = p.id === MEETINGS_ID ? ' tile-gradient' : '';
    header = viewHeader({
      tile: `<span class="tile${tileCls}" style="${tileCls ? '' : `background:${h(p.color)}`}">${ic(icon)}</span>`,
      title: p.name, rename: 'rename-project', id: p.id,
      sub: notes ? `${tasks.length} note${tasks.length === 1 ? '' : 's'}` : completeText(tasks),
      addLabel: notes ? 'Add note' : 'Add task',
      canDelete: !BUILT_IN_IDS.has(p.id), deleteLabel: notes ? 'Delete notebook' : 'Delete project',
    });
    cols = notes ? ['name', 'tags', 'created']
      : FEED_IDS.has(p.id) ? ['name', 'assignee', 'due', 'tags', 'reporter', 'created']
        : p.id === MEETINGS_ID ? ['name', 'assignee', 'due', 'tags', 'meeting', 'created']
          : ['name', 'assignee', 'due', 'tags', 'project', 'created'];
  } else if (v.type === 'tag') {
    const name = tagsInUse().find((t) => eqName(t.name, v.id))!.name;
    tasks = data.tasks.filter((t) => !isNote(t) && t.tags.some((x) => eqName(x, name)));
    header = viewHeader({
      tile: `<span class="tile" style="${tagStyle(name)}">${ic('tag')}</span>`,
      title: name, rename: 'rename-tag', id: name, sub: completeText(tasks), addLabel: 'Add task',
      canDelete: true, deleteLabel: 'Delete tag',
    });
    cols = ['name', 'assignee', 'due', 'tags', 'project', 'created'];
  } else {
    const person = data.people.find((p) => eqName(p.name, v.id))!;
    const owner = eqName(person.name, OWNER);
    tasks = data.tasks.filter((t) => !isNote(t) && eqName(t.assignee, person.name));
    header = viewHeader({
      tile: `<span class="tile tile-avatar" style="background:${h(person.color)}">${h(initials(person.name))}</span>`,
      title: owner ? displayName(person.name) : person.name, rename: owner ? null : 'rename-person', id: person.name,
      sub: completeText(tasks), addLabel: 'Add task', canDelete: !owner, deleteLabel: 'Remove person',
    });
    cols = ['name', 'assignee', 'due', 'tags', 'project', 'created'];
  }

  return `<div class="main-inner">${header}${renderList(tasks, cols, notes, v)}</div>`;
}

function completeText(tasks: Task[]): string {
  return `${tasks.filter((t) => t.done).length} of ${tasks.length} complete`;
}

function viewHeader(o: {
  tile: string; title: string; rename: string | null; id: string; sub: string; addLabel: string;
  canDelete: boolean; deleteLabel: string;
}): string {
  const title = o.rename
    ? `<input class="view-title" value="${h(o.title)}" data-action-change="${o.rename}" data-id="${h(o.id)}" data-fk="view-title" aria-label="Name" />`
    : `<h1 class="view-title static">${h(o.title)}</h1>`;
  return `<header class="view-header">
    ${o.tile}
    <div class="view-titles">${title}<div class="view-sub">${h(o.sub)}</div></div>
    <div class="view-actions">
      <button class="btn btn-primary" data-action="focus-add">${ic('plus')} ${h(o.addLabel)}</button>
      ${o.canDelete ? `<button class="btn btn-icon" data-action="delete-view" title="${h(o.deleteLabel)}" aria-label="${h(o.deleteLabel)}">${ic('trash-2')}</button>` : ''}
    </div>
  </header>`;
}

function colStyle(): string {
  return (Object.keys(DEFAULT_W) as ColKey[]).map((k) => `--w-${k}:${ui.colW[k] ?? DEFAULT_W[k]}px`).join(';');
}

function renderList(tasks: Task[], cols: ColKey[], notes: boolean, v: View): string {
  const head = cols.map((c) => `<div class="cell c-${c} hcell">
      <span>${c === 'name' && notes ? 'Note' : COL_LABEL[c]}</span>
      <span class="resizer" data-col="${c}" title="Drag to resize, double-click to reset"></span>
    </div>`).join('');

  const rows = tasks.map((t) => renderRow(t, cols, notes)).join('');
  let empty = '';
  if (!tasks.length) {
    empty = v.type === 'project'
      ? notes
        ? emptyState('notebook-pen', 'No notes yet', 'Type a title below and press Enter to start a note.')
        : v.id === SLACK_ID
          ? emptyState('hash', 'No Slack tasks yet', 'Rows your Slack automation adds to the "Slack Tasks" tab show up here.')
          : v.id === JIRA_ID
            ? emptyState('ticket', 'No Jira tickets', 'Connect Jira from the sheet (Task manager → Connect Jira…). Open tickets assigned to you appear here.')
            : emptyState('circle-check-big', 'No tasks yet', 'Type a task below and press Enter to add it.')
      : v.type === 'tag'
        ? emptyState('tag', 'Nothing tagged', 'Add this tag to a task to see it here.')
        : emptyState('user', 'Nothing assigned', 'Assign a task to this person, or add one below.');
  }
  const addRow = `<div class="lrow ladd">
      <div class="cell c-name">${ic('plus', 'add-icon')}
        <input class="add-input" data-keep data-fk="add-task" data-action-key="add-task" placeholder="${notes ? 'Add note…' : 'Add task…'}" aria-label="${notes ? 'Add note' : 'Add task'}" />
      </div>
    </div>`;

  return `<div class="list${notes ? ' list-notes' : ''}" style="${colStyle()}">
    <div class="list-scroll">
      <div class="lrow lhead">${head}<div class="cell c-fill"></div></div>
      ${rows}
      ${addRow}
      ${empty}
    </div>
  </div>`;
}

function renderRow(t: Task, cols: ColKey[], notes: boolean): string {
  const sel = ui.selectedId === t.id ? ' selected' : '';
  const cells = cols.map((c) => {
    switch (c) {
      case 'name':
        return `<div class="cell c-name">
          ${notes ? `<span class="note-icon">${ic('file-text')}</span>` : ballHtml(t)}
          <span class="task-title${t.done && !notes ? ' is-done' : ''}">${t.title ? h(t.title) : `<span class="muted">Untitled</span>`}</span>
          ${commentCount(t.id)}
        </div>`;
      case 'assignee':
        return `<div class="cell c-assignee clickable" data-action="assignee-menu" data-id="${h(t.id)}">
          ${t.assignee ? `${avatar(t.assignee, 'sm')}<span class="ellipsis">${h(displayName(t.assignee))}</span>` : `<span class="placeholder">${ic('user-plus')}</span>`}
        </div>`;
      case 'due': {
        const d = fmtDue(t.due, t.done);
        return `<div class="cell c-due clickable" data-action="due-menu" data-id="${h(t.id)}">
          ${d.text ? `<span class="${d.cls}">${h(d.text)}</span>` : `<span class="placeholder">${ic('calendar')}</span>`}
        </div>`;
      }
      case 'tags':
        return `<div class="cell c-tags clickable" data-action="tags-menu" data-id="${h(t.id)}">
          ${t.tags.length ? t.tags.map((x) => `<span class="chip" style="${tagStyle(x)}">${h(x)}</span>`).join('') : `<span class="placeholder">${ic('tag')}</span>`}
        </div>`;
      case 'project': {
        const p = projectById(t.projectId);
        return `<div class="cell c-project clickable" data-action="project-menu" data-id="${h(t.id)}">
          ${p ? `<span class="pdot" style="background:${h(p.color)}"></span><span class="ellipsis">${h(p.name)}</span>` : ''}
        </div>`;
      }
      case 'reporter':
        return `<div class="cell c-reporter"><span class="ellipsis">${h(t.reporter)}</span></div>`;
      case 'meeting': {
        const m = meetingOfTask(t.id);
        return `<div class="cell c-meeting">${m
          ? `<button class="link-btn ellipsis" data-action="nav" data-type="meeting" data-id="${h(m.id)}">${h(m.title)}${m.date ? ` · ${h(fmtMeetingDate(m.date))}` : ''}</button>`
          : ''}</div>`;
      }
      case 'created':
        return `<div class="cell c-created muted">${h(fmtCreated(t.createdAt))}</div>`;
    }
  }).join('');
  return `<div class="lrow${sel}${t.done && !notes ? ' done' : ''}" data-action="open-task" data-id="${h(t.id)}" draggable="true">${cells}<div class="cell c-fill"></div></div>`;
}

function commentCount(taskId: string): string {
  const n = data.comments.filter((c) => c.taskId === taskId).length;
  return n ? `<span class="comment-count" title="${n} comment${n === 1 ? '' : 's'}">${ic('message-circle')}${n}</span>` : '';
}

function ballHtml(t: Task, extraCls = ''): string {
  const st = statusOf(t);
  return `<button class="ball st-${st}${extraCls}" data-action="status-menu" data-id="${h(t.id)}" title="${STATUS_LABEL[st]}" aria-label="Status: ${STATUS_LABEL[st]}">${st === 'urgent' ? '<b>!</b>' : st === 'done' ? TICK_SVG : ''}</button>`;
}

function statusBall(st: Status): string {
  return `<span class="ball st-${st} static">${st === 'urgent' ? '<b>!</b>' : st === 'done' ? TICK_SVG : ''}</span>`;
}

// ================================================================ rendering: meeting page

function renderMeeting(m: Meeting): string {
  const items = m.actions
    .map((a) => ({ a, t: taskById(a.taskId) }))
    .filter((x): x is { a: MeetingAction; t: Task } => !!x.t)
    .sort((x, y) => Number(x.t.done) - Number(y.t.done));
  const open = items.filter((x) => !x.t.done).length;
  const taskProjects = data.projects.filter((p) => p.kind === 'project');
  const meta = [
    m.date ? `${ic('calendar')} ${h(fmtMeetingDate(m.date))}` : '',
    m.attendees.length ? `${ic('users')} ${h(m.attendees.join(', '))}` : '',
    m.fileName ? `${ic('file-text')} ${h(m.fileName)}` : '',
  ].filter(Boolean).map((x) => `<span>${x}</span>`).join('');

  const rows = items.map(({ a, t }) => {
    const d = fmtDue(t.due, t.done);
    return `<div class="action-row${t.done ? ' done' : ''}${ui.selectedId === t.id ? ' selected' : ''}">
      ${ballHtml(t)}
      <button class="action-title link-btn" data-action="open-task" data-id="${h(t.id)}">${h(t.title || a.title)}</button>
      <span class="action-meta">
        ${t.assignee ? `${avatar(t.assignee, 'sm')}<span class="ellipsis">${h(displayName(t.assignee))}</span>` : a.owner ? `<span class="muted ellipsis" title="Not on the Assignees list">${h(a.owner)}</span>` : ''}
        ${d.text ? `<span class="${d.cls}">${h(d.text)}</span>` : ''}
      </span>
      <select class="select-sm" data-action-change="action-project" data-id="${h(t.id)}" aria-label="Project">
        ${taskProjects.map((p) => `<option value="${h(p.id)}"${p.id === t.projectId ? ' selected' : ''}>${h(p.name)}</option>`).join('')}
      </select>
      <button class="btn btn-sm" data-action="open-task" data-id="${h(t.id)}">Open</button>
      <button class="btn btn-icon btn-sm" data-action="delete-action" data-id="${h(t.id)}" title="Delete action item" aria-label="Delete action item">${ic('trash-2')}</button>
    </div>`;
  }).join('');

  return `<div class="main-inner meeting-scroll">
    <header class="view-header">
      <span class="tile tile-gradient">${ic('sparkles')}</span>
      <div class="view-titles">
        <input class="view-title" value="${h(m.title)}" data-action-change="rename-meeting" data-id="${h(m.id)}" data-fk="view-title" aria-label="Meeting title" />
        <div class="view-sub meeting-meta">${meta || 'No details'}</div>
      </div>
      <div class="view-actions">
        <button class="btn btn-icon" data-action="delete-view" title="Delete meeting" aria-label="Delete meeting">${ic('trash-2')}</button>
      </div>
    </header>
    <div class="meeting-body">
      <section class="card">
        <h2 class="card-title">Action items <span class="muted">${open} open</span></h2>
        ${rows || `<p class="muted empty-inline">No action items yet. Add one below.</p>`}
        <div class="action-add">${ic('plus', 'add-icon')}<input class="add-input" data-keep data-fk="add-action" data-action-key="add-action" data-id="${h(m.id)}" placeholder="Add action item…" /></div>
      </section>
      <section class="card">
        <h2 class="card-title">Meeting notes</h2>
        ${editorHtml(`meeting:${m.id}`, m.notes, { placeholder: 'Write meeting notes…', minHeight: 160 })}
      </section>
    </div>
  </div>`;
}

// ================================================================ rendering: task panel

function renderPanel(): string {
  const t = ui.selectedId ? taskById(ui.selectedId) : undefined;
  if (!t) return '';
  const note = isNote(t);
  const p = projectById(t.projectId);
  const m = meetingOfTask(t.id);
  const st = statusOf(t);
  const showFeedFields = FEED_IDS.has(t.projectId) || !!t.reporter || t.links.length > 0;
  const d = fmtDue(t.due, t.done);

  const field = (label: string, body: string) => `<div class="field"><div class="field-label">${h(label)}</div><div class="field-value">${body}</div></div>`;

  const fields = [
    note ? '' : field('Assignee', `<button class="field-btn" data-action="assignee-menu" data-id="${h(t.id)}">${t.assignee ? `${avatar(t.assignee, 'sm')} ${h(displayName(t.assignee))}` : `<span class="muted">No assignee</span>`}</button>`),
    note ? '' : field('Due date', `<button class="field-btn" data-action="due-menu" data-id="${h(t.id)}">${d.text ? `${ic('calendar')} <span class="${d.cls}">${h(d.text)}</span>` : `<span class="muted">No due date</span>`}</button>`),
    field(note ? 'Notebook' : 'Project', `<button class="field-btn" data-action="project-menu" data-id="${h(t.id)}">${p ? `<span class="pdot" style="background:${h(p.color)}"></span> ${h(p.name)}` : '—'}</button>`),
    note ? '' : field('Status', `<button class="field-btn" data-action="status-menu" data-id="${h(t.id)}">${statusBall(st)} ${STATUS_LABEL[st]}</button>`),
    m ? field('Meeting', `<button class="field-btn" data-action="nav" data-type="meeting" data-id="${h(m.id)}">${ic('sparkles')} ${h(m.title)}${m.date ? ` · ${h(fmtMeetingDate(m.date))}` : ''}</button>`) : '',
    field('Created', `<span class="field-text">${h(fmtDateTime(t.createdAt))}${t.createdAt ? ` <span class="muted">· ${h(timeAgo(t.createdAt))}</span>` : ''}</span>`),
    field('Tags', `<div class="tag-input">
      ${t.tags.map((x, i) => `<span class="chip" style="${tagStyle(x)}">${h(x)}<button class="chip-x" data-action="remove-tag" data-id="${h(t.id)}" data-index="${i}" aria-label="Remove ${h(x)}">${ic('x')}</button></span>`).join('')}
      <input data-keep data-fk="panel-tag" data-action-key="panel-tag" data-id="${h(t.id)}" list="tag-options" placeholder="${t.tags.length ? '' : 'Add a tag…'}" aria-label="Add tag" />
      <datalist id="tag-options">${data.tags.filter((x) => !t.tags.some((y) => eqName(x.name, y))).map((x) => `<option value="${h(x.name)}"></option>`).join('')}</datalist>
    </div>`),
    !note && showFeedFields ? field('Reporter', `<input class="field-input" value="${h(t.reporter)}" data-action-change="reporter" data-id="${h(t.id)}" data-fk="reporter" placeholder="Who asked for this" />`) : '',
    !note && showFeedFields ? field('Links', `<div class="links">
      ${t.links.map((l, i) => `<div class="link-row"><a href="${h(l)}" target="_blank" rel="noopener noreferrer" class="ellipsis">${ic('external-link')} ${h(l.replace(/^https?:\/\//, ''))}</a><button class="btn btn-icon btn-xs" data-action="remove-link" data-id="${h(t.id)}" data-index="${i}" aria-label="Remove link">${ic('x')}</button></div>`).join('')}
      <input class="field-input" data-keep data-fk="link-add" data-action-key="add-link" data-id="${h(t.id)}" placeholder="Add a link (http or https)…" />
    </div>`) : '',
  ].join('');

  const comments = data.comments
    .filter((c) => c.taskId === t.id)
    .sort((a, b) => a.createdAt.localeCompare(b.createdAt));

  const draft = commentDrafts.get(t.id) || '';
  return `<div class="panel-top">
      ${note ? '<span class="panel-kind">' + ic('file-text') + ' Note</span>' : `<button class="btn btn-sm${t.done ? ' btn-done' : ''}" data-action="toggle-done" data-id="${h(t.id)}">${ic('check')} ${t.done ? 'Completed' : 'Mark complete'}</button>`}
      <span class="spacer"></span>
      <button class="btn btn-icon" data-action="delete-task" data-id="${h(t.id)}" title="Delete" aria-label="Delete">${ic('trash-2')}</button>
      <button class="btn btn-icon" data-action="close-panel" title="Close (Esc)" aria-label="Close">${ic('arrow-right-to-line')}</button>
    </div>
    <div class="panel-body">
      <textarea class="panel-title autogrow" rows="1" data-action-input="title" data-id="${h(t.id)}" data-fk="panel-title" placeholder="${note ? 'Note title' : 'Task name'}">${h(t.title)}</textarea>
      <div class="fields">${fields}</div>
      <div class="panel-section">
        <h3>${note ? 'Note' : 'Description'}</h3>
        ${editorHtml(`desc:${t.id}`, t.description, { placeholder: note ? 'Write your note…' : 'What is this task about?', minHeight: note ? 240 : 110 })}
      </div>
      <div class="panel-section">
        <h3>Comments ${comments.length ? `<span class="muted">${comments.length}</span>` : ''}</h3>
        <div class="comments">
          ${comments.map((c) => `<div class="comment">
            ${avatar(c.author)}
            <div class="comment-main">
              <div class="comment-head"><b>${h(displayName(c.author))}</b><span class="muted" title="${h(fmtDateTime(c.createdAt))}">${h(timeAgo(c.createdAt))}</span>
                <button class="btn btn-icon btn-xs comment-del" data-action="delete-comment" data-id="${h(c.id)}" title="Delete comment" aria-label="Delete comment">${ic('trash-2')}</button></div>
              <div class="rt-content comment-text">${toDisplayHtml(c.text)}</div>
            </div>
          </div>`).join('') || '<p class="muted empty-inline">No comments yet.</p>'}
        </div>
        <div class="comment-box">
          ${editorHtml(`comment:${t.id}`, draft, { placeholder: 'Write a comment… (⌘+Enter to post)', compact: true })}
          <div class="comment-actions"><button class="btn btn-primary btn-sm" data-action="post-comment" data-id="${h(t.id)}">Comment</button></div>
        </div>
      </div>
    </div>`;
}

// ================================================================ menus

function openMenu(kind: MenuKind, anchor: HTMLElement, taskId: string): void {
  const r = anchor.getBoundingClientRect();
  ui.menu = { kind, taskId, x: r.left, y: r.bottom + 4, query: '' };
  renderMenu();
  const search = document.querySelector<HTMLInputElement>('#menu-root .menu-search');
  search?.focus();
}

function closeMenu(): void {
  if (!ui.menu) return;
  ui.menu = null;
  document.getElementById('menu-root')!.innerHTML = '';
}

function menuItem(value: string, label: string, opts: { checked?: boolean; icon?: string; cls?: string } = {}): string {
  return `<button class="menu-item${opts.cls ? ' ' + opts.cls : ''}" data-action="menu-pick" data-value="${h(value)}">
    ${opts.icon || ''}<span class="menu-label">${label}</span>${opts.checked ? ic('check', 'menu-check') : ''}
  </button>`;
}

function renderMenu(): void {
  const root = document.getElementById('menu-root')!;
  const m = ui.menu;
  const t = m ? taskById(m.taskId) : undefined;
  if (!m || !t) {
    ui.menu = null;
    root.innerHTML = '';
    return;
  }
  let body = '';
  switch (m.kind) {
    case 'status': {
      const st = statusOf(t);
      body = (['todo', 'urgent', 'progress', 'done'] as Status[])
        .map((s) => menuItem(s, STATUS_LABEL[s], { checked: s === st, icon: statusBall(s) })).join('');
      break;
    }
    case 'assignee':
      body = data.people.map((p) => menuItem(p.name, h(displayName(p.name)), { checked: eqName(p.name, t.assignee), icon: avatar(p.name, 'sm') })).join('') +
        (t.assignee ? `<div class="menu-sep"></div>${menuItem('', 'Unassign', { icon: ic('user-x') })}` : '');
      break;
    case 'due': {
      const today = todayYmd();
      const nextWeek = addDays(today, 7);
      body = menuItem(today, 'Today', { icon: ic('sun'), checked: t.due === today }) +
        menuItem(addDays(today, 1), 'Tomorrow', { icon: ic('sunrise'), checked: t.due === addDays(today, 1) }) +
        menuItem(nextWeek, `Next week <span class="muted">${h(shortDate(parseYmd(nextWeek)!))}</span>`, { icon: ic('calendar-plus'), checked: t.due === nextWeek }) +
        `<label class="menu-item menu-date">${ic('calendar-days')}<span class="menu-label">Pick a date</span>
          <input type="date" value="${h(t.due)}" data-action-change="menu-date" /></label>` +
        (t.due ? `<div class="menu-sep"></div>${menuItem('', 'Clear', { icon: ic('x') })}` : '');
      break;
    }
    case 'tags': {
      const q = m.query.trim();
      const list = data.tags.filter((x) => !q || x.name.toLowerCase().includes(q.toLowerCase()));
      const exact = data.tags.some((x) => eqName(x.name, q));
      body = `<input class="menu-search" placeholder="Search or create a tag" value="${h(m.query)}" data-action-key="menu-tag-search" data-action-input="menu-tag-search" />` +
        `<div class="menu-scroll">` +
        list.map((x) => menuItem(x.name, `<span class="chip" style="${tagStyle(x.name)}">${h(x.name)}</span>`, { checked: t.tags.some((y) => eqName(y, x.name)) })).join('') +
        (q && !exact ? menuItem(`+${q}`, `Create “${h(q)}”`, { icon: ic('plus') }) : '') +
        (!list.length && !q ? `<p class="menu-empty">Type to create your first tag.</p>` : '') +
        `</div>`;
      break;
    }
    case 'project': {
      const kind = kindOfTask(t);
      body = data.projects.filter((p) => p.kind === kind)
        .map((p) => menuItem(p.id, h(p.name), { checked: p.id === t.projectId, icon: `<span class="pdot" style="background:${h(p.color)}"></span>` })).join('');
      break;
    }
  }
  root.innerHTML = `<div class="menu menu-${m.kind}" role="menu" style="left:${m.x}px;top:${m.y}px">${body}</div>`;
  const el = root.firstElementChild as HTMLElement;
  const rect = el.getBoundingClientRect();
  if (rect.right > window.innerWidth - 8) el.style.left = `${Math.max(8, window.innerWidth - rect.width - 8)}px`;
  if (rect.bottom > window.innerHeight - 8) el.style.top = `${Math.max(8, m.y - rect.height - 36)}px`;
}

function pickMenu(value: string): void {
  const m = ui.menu;
  const t = m ? taskById(m.taskId) : undefined;
  if (!m || !t) return closeMenu();
  switch (m.kind) {
    case 'status':
      setStatus(t, value as Status);
      closeMenu();
      break;
    case 'assignee':
      t.assignee = value;
      closeMenu();
      break;
    case 'due':
      t.due = value;
      closeMenu();
      break;
    case 'project':
      moveTask(t, value);
      closeMenu();
      break;
    case 'tags': {
      if (value.startsWith('+')) {
        const name = ensureTag(value.slice(1).trim());
        if (name && !t.tags.some((x) => eqName(x, name))) t.tags.push(name);
        m.query = '';
      } else {
        const i = t.tags.findIndex((x) => eqName(x, value));
        if (i >= 0) t.tags.splice(i, 1);
        else t.tags.push(value);
      }
      commit();
      document.querySelector<HTMLInputElement>('#menu-root .menu-search')?.focus();
      return;
    }
  }
  commit();
}

// ================================================================ actions

function addTaskFromInput(input: HTMLInputElement): void {
  const title = input.value.trim();
  if (!title) return;
  const v = currentView();
  if (!v || v.type === 'meeting') return;
  input.value = '';
  if (v.type === 'project') newTask({ projectId: v.id, title });
  else if (v.type === 'tag') newTask({ projectId: firstTaskProjectId(), title, tags: [v.id] });
  else newTask({ projectId: firstTaskProjectId(), title, assignee: data.people.find((p) => eqName(p.name, v.id))?.name || '' });
  commit();
}

function addTagToTask(t: Task, raw: string): void {
  for (const part of raw.split(',')) {
    const name = part.trim();
    if (!name) continue;
    const real = ensureTag(name);
    if (!t.tags.some((x) => eqName(x, real))) t.tags.push(real);
  }
}

function addFromSidebar(kind: 'project' | 'notebook' | 'person', name: string): void {
  name = name.trim();
  ui.adding = null;
  if (!name) return render();
  if (kind === 'person') {
    if (!data.people.some((p) => eqName(p.name, name))) {
      data.people.push({ name, color: PERSON_COLORS[data.people.length % PERSON_COLORS.length] });
    }
    commit();
    return;
  }
  const notes = kind === 'notebook';
  const count = data.projects.filter((p) => p.kind === (notes ? 'notes' : 'project')).length;
  const p: Project = { id: uid('p'), name, color: PROJECT_COLORS[(count + (notes ? 4 : 0)) % PROJECT_COLORS.length], kind: notes ? 'notes' : 'project' };
  data.projects.push(p);
  ui.view = { type: 'project', id: p.id };
  lsSet(LS.view, ui.view);
  commit();
}

async function deleteCurrentView(): Promise<void> {
  const v = currentView();
  if (!v) return;
  if (v.type === 'project') {
    const p = projectById(v.id)!;
    const ids = new Set(data.tasks.filter((t) => t.projectId === p.id).map((t) => t.id));
    const what = p.kind === 'notes' ? 'notebook' : 'project';
    const items = p.kind === 'notes' ? 'note' : 'task';
    const ok = await confirmDialog({
      title: `Delete ${what} “${p.name}”?`,
      message: ids.size ? `This also deletes its ${ids.size} ${items}${ids.size === 1 ? '' : 's'}. This can't be undone.` : "This can't be undone.",
      confirm: `Delete ${what}`,
    });
    if (!ok) return;
    deleteTasks(ids);
    data.projects = data.projects.filter((x) => x.id !== p.id);
    ui.view = null;
  } else if (v.type === 'tag') {
    const ok = await confirmDialog({ title: `Delete tag “${v.id}”?`, message: 'The tag is removed from every task that has it.', confirm: 'Delete tag' });
    if (!ok) return;
    for (const t of data.tasks) t.tags = t.tags.filter((x) => !eqName(x, v.id));
    data.tags = data.tags.filter((x) => !eqName(x.name, v.id));
    ui.view = null;
  } else if (v.type === 'person') {
    if (eqName(v.id, OWNER)) return;
    const n = data.tasks.filter((t) => eqName(t.assignee, v.id)).length;
    const ok = await confirmDialog({
      title: `Remove ${v.id}?`,
      message: n ? `Their ${n} task${n === 1 ? '' : 's'} will become unassigned.` : 'They have no tasks assigned.',
      confirm: 'Remove',
    });
    if (!ok) return;
    for (const t of data.tasks) if (eqName(t.assignee, v.id)) t.assignee = '';
    data.people = data.people.filter((p) => !eqName(p.name, v.id));
    ui.view = null;
  } else if (v.type === 'meeting') {
    const m = meetingById(v.id)!;
    const ok = await confirmDialog({
      title: `Delete meeting “${m.title}”?`,
      message: m.actions.length ? 'Its notes are deleted. Its action items stay as tasks in your projects.' : 'Its notes are deleted.',
      confirm: 'Delete meeting',
    });
    if (!ok) return;
    data.meetings = data.meetings.filter((x) => x.id !== m.id);
    ui.view = null;
  }
  lsSet(LS.view, null);
  commit();
}

function renameView(kind: string, id: string, value: string): void {
  const name = value.trim();
  if (!name) return render();
  if (kind === 'rename-project') {
    const p = projectById(id);
    if (p && p.name !== name) p.name = name;
  } else if (kind === 'rename-meeting') {
    const m = meetingById(id);
    if (m) m.title = name;
  } else if (kind === 'rename-tag') {
    if (eqName(id, name) && id === name) return;
    const existing = data.tags.find((t) => eqName(t.name, name) && !eqName(t.name, id));
    for (const t of data.tasks) {
      const i = t.tags.findIndex((x) => eqName(x, id));
      if (i < 0) continue;
      t.tags.splice(i, 1);
      if (!t.tags.some((x) => eqName(x, existing ? existing.name : name))) t.tags.splice(i, 0, existing ? existing.name : name);
    }
    const tag = data.tags.find((t) => eqName(t.name, id));
    if (tag && !existing) tag.name = name;
    else if (tag) data.tags = data.tags.filter((t) => t !== tag);
    ui.view = { type: 'tag', id: existing ? existing.name : name };
  } else if (kind === 'rename-person') {
    if (eqName(id, OWNER) || data.people.some((p) => eqName(p.name, name) && !eqName(p.name, id))) return render();
    const p = data.people.find((x) => eqName(x.name, id));
    if (!p) return;
    for (const t of data.tasks) if (eqName(t.assignee, id)) t.assignee = name;
    p.name = name;
    ui.view = { type: 'person', id: name };
  }
  lsSet(LS.view, ui.view);
  commit();
}

async function uploadTranscript(file: File): Promise<void> {
  if (ui.aiBusy) return;
  ui.aiBusy = file.name;
  ui.collapsed.meetings = false;
  render();
  try {
    const text = await readTranscript(file);
    const draft = await summariseTranscript(text, file.name, data.people.map((p) => p.name), todayYmd());
    const m = createMeetingFromDraft(draft, file.name);
    ui.aiBusy = null;
    ui.view = { type: 'meeting', id: m.id };
    lsSet(LS.view, ui.view);
    commit();
  } catch (err) {
    ui.aiBusy = null;
    render();
    await alertDialog("Couldn't summarise that transcript", String((err as Error)?.message || err));
  }
}

// ================================================================ events

function onClick(e: MouseEvent): void {
  const target = e.target as HTMLElement;
  const dlg = target.closest<HTMLElement>('[data-dialog]');
  if (dlg) {
    if (dlg.dataset.dialog === 'cancel' && dlg.classList.contains('dialog-backdrop') && target !== dlg) return;
    closeDialog(dlg.dataset.dialog === 'ok');
    return;
  }
  const el = target.closest<HTMLElement>('[data-action]');
  if (!el) return;
  if (target.closest('a[href]') && !target.closest('.rt-body')) return; // plain links just open
  const action = el.dataset.action!;
  const id = el.dataset.id || '';
  const t = id ? taskById(id) : undefined;

  switch (action) {
    case 'nav': {
      const type = el.dataset.type as View['type'];
      ui.view = { type, id };
      lsSet(LS.view, ui.view);
      if (type !== 'meeting' && ui.selectedId) {
        const sel = taskById(ui.selectedId);
        if (!sel) ui.selectedId = null;
      }
      render();
      break;
    }
    case 'toggle-section': {
      const key = el.dataset.key!;
      ui.collapsed[key] = !ui.collapsed[key];
      const persisted = { projects: ui.collapsed.projects, notes: ui.collapsed.notes, meetings: ui.collapsed.meetings };
      lsSet(LS.collapsed, persisted);
      render();
      break;
    }
    case 'add-project':
    case 'add-notebook':
    case 'add-person': {
      const kind = action.slice(4) as 'project' | 'notebook' | 'person';
      ui.adding = kind;
      ui.collapsed[kind === 'project' ? 'projects' : kind === 'notebook' ? 'notes' : 'people'] = false;
      render();
      document.querySelector<HTMLInputElement>(`[data-fk="add-${kind}"]`)?.focus();
      break;
    }
    case 'upload-transcript':
      (document.getElementById('transcript-input') as HTMLInputElement).click();
      break;
    case 'open-task':
      if (!t) break;
      ui.selectedId = t.id;
      render();
      break;
    case 'close-panel':
      ui.selectedId = null;
      render();
      break;
    case 'status-menu':
    case 'assignee-menu':
    case 'due-menu':
    case 'tags-menu':
    case 'project-menu':
      if (!t) break;
      e.stopPropagation();
      if (ui.menu && ui.menu.taskId === id && ui.menu.kind === action.replace('-menu', '')) closeMenu();
      else openMenu(action.replace('-menu', '') as MenuKind, el, id);
      break;
    case 'menu-pick':
      pickMenu(el.dataset.value || '');
      break;
    case 'toggle-done':
      if (!t) break;
      setStatus(t, t.done ? 'todo' : 'done');
      commit();
      break;
    case 'delete-task':
      if (!t) break;
      void confirmDialog({
        title: `Delete ${isNote(t) ? 'note' : 'task'}?`,
        message: `“${t.title || 'Untitled'}” and its comments will be deleted.`,
        confirm: 'Delete',
      }).then((ok) => {
        if (!ok) return;
        deleteTasks(new Set([t.id]));
        commit();
      });
      break;
    case 'delete-action':
      if (!t) break;
      void confirmDialog({ title: 'Delete action item?', message: `“${t.title}” is removed from the meeting and its task is deleted.`, confirm: 'Delete' })
        .then((ok) => {
          if (!ok) return;
          deleteTasks(new Set([t.id]));
          commit();
        });
      break;
    case 'remove-tag':
      if (!t) break;
      t.tags.splice(Number(el.dataset.index), 1);
      commit();
      break;
    case 'remove-link':
      if (!t) break;
      t.links.splice(Number(el.dataset.index), 1);
      commit();
      break;
    case 'delete-comment': {
      const c = data.comments.find((x) => x.id === id);
      if (!c) break;
      void confirmDialog({ title: 'Delete comment?', message: "This can't be undone.", confirm: 'Delete' }).then((ok) => {
        if (!ok) return;
        data.comments = data.comments.filter((x) => x.id !== id);
        commit();
      });
      break;
    }
    case 'post-comment':
      postComment(id);
      break;
    case 'focus-add': {
      const input = document.querySelector<HTMLInputElement>('[data-fk="add-task"]');
      input?.scrollIntoView({ block: 'nearest' });
      input?.focus();
      break;
    }
    case 'delete-view':
      void deleteCurrentView();
      break;
  }
}

function postComment(taskId: string): void {
  const draft = commentDrafts.get(taskId) || '';
  const text = sanitizeHtml(draft);
  if (isEmptyHtml(text)) return;
  data.comments.push({ id: uid('c'), taskId, author: OWNER, text, createdAt: new Date().toISOString() });
  commentDrafts.delete(taskId);
  commit();
  const box = document.querySelector<HTMLElement>(`[data-fk="rt:comment:${CSS.escape(taskId)}"]`);
  box?.focus();
}

function onKeyDown(e: KeyboardEvent): void {
  const target = e.target as HTMLElement;

  if (ui.dialogOpen) {
    if (e.key === 'Escape') {
      e.preventDefault();
      closeDialog(false);
    } else if (e.key === 'Enter' && !(target instanceof HTMLButtonElement)) {
      e.preventDefault();
      closeDialog(true);
    }
    return;
  }

  const key = target.dataset?.actionKey;
  if (key) {
    const input = target as HTMLInputElement;
    if (key === 'add-task' && e.key === 'Enter') {
      e.preventDefault();
      addTaskFromInput(input);
      return;
    }
    if (key === 'add-action' && e.key === 'Enter') {
      e.preventDefault();
      const m = meetingById(input.dataset.id || '');
      const title = input.value.trim();
      if (m && title) {
        input.value = '';
        addMeetingAction(m, title, '', '', '');
        commit();
      }
      return;
    }
    if (key.startsWith('add-') && ['add-project', 'add-notebook', 'add-person'].includes(key)) {
      if (e.key === 'Enter') {
        e.preventDefault();
        const v = input.value;
        input.value = '';
        addFromSidebar(key.slice(4) as 'project' | 'notebook' | 'person', v);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        input.value = '';
        ui.adding = null;
        render();
      }
      return;
    }
    if (key === 'panel-tag') {
      const t = taskById(input.dataset.id || '');
      if (!t) return;
      if ((e.key === 'Enter' || e.key === ',') && input.value.trim()) {
        e.preventDefault();
        addTagToTask(t, input.value);
        input.value = '';
        commit();
      } else if (e.key === 'Backspace' && !input.value && t.tags.length) {
        e.preventDefault();
        t.tags.pop();
        commit();
      }
      if (e.key !== 'Escape') return;
    }
    if (key === 'add-link' && e.key === 'Enter') {
      e.preventDefault();
      const t = taskById(input.dataset.id || '');
      const url = normalizeLinkInput(input.value, false);
      if (!t) return;
      if (!url) {
        input.classList.add('invalid');
        return;
      }
      if (!t.links.includes(url)) t.links.push(url);
      input.value = '';
      commit();
      return;
    }
    if (key === 'menu-tag-search' && e.key === 'Enter') {
      e.preventDefault();
      const q = input.value.trim();
      if (!q || !ui.menu) return;
      const existing = data.tags.find((x) => eqName(x.name, q));
      pickMenu(existing ? existing.name : `+${q}`);
      return;
    }
  }

  if (target.matches?.('input.view-title, input[data-action-change="reporter"]') && e.key === 'Enter') {
    e.preventDefault();
    target.blur();
    return;
  }
  if (target.matches?.('textarea.panel-title') && e.key === 'Enter') {
    e.preventDefault();
    target.blur();
    return;
  }

  if (e.key === 'Escape') {
    if (ui.menu) {
      e.preventDefault();
      closeMenu();
      return;
    }
    if (ui.adding) {
      ui.adding = null;
      render();
      return;
    }
    if (ui.selectedId) {
      e.preventDefault();
      (document.activeElement as HTMLElement | null)?.blur();
      ui.selectedId = null;
      render();
    }
  }
}

function onInput(e: Event): void {
  const target = e.target as HTMLElement;
  const kind = target.dataset?.actionInput;
  if (target.classList.contains('invalid')) target.classList.remove('invalid');
  if (!kind) return;
  if (kind === 'title') {
    const t = taskById(target.dataset.id || '');
    if (!t) return;
    t.title = (target as HTMLTextAreaElement).value.replace(/\n/g, ' ');
    autoGrow(target as HTMLTextAreaElement);
    const rowTitle = document.querySelector(`.lrow[data-id="${CSS.escape(t.id)}"] .task-title`);
    if (rowTitle) rowTitle.textContent = t.title || 'Untitled';
    commit({ render: false });
  } else if (kind === 'menu-tag-search' && ui.menu) {
    ui.menu.query = (target as HTMLInputElement).value;
    const pos = (target as HTMLInputElement).selectionStart;
    renderMenu();
    const s = document.querySelector<HTMLInputElement>('#menu-root .menu-search');
    if (s) {
      s.focus();
      s.setSelectionRange(pos, pos);
    }
  }
}

function onChange(e: Event): void {
  const target = e.target as HTMLInputElement | HTMLSelectElement;
  if (target.id === 'transcript-input') {
    const file = (target as HTMLInputElement).files?.[0];
    (target as HTMLInputElement).value = '';
    if (file) void uploadTranscript(file);
    return;
  }
  const kind = target.dataset?.actionChange;
  if (!kind) return;
  const id = target.dataset.id || '';
  if (kind.startsWith('rename-')) {
    renameView(kind, id, target.value);
  } else if (kind === 'reporter') {
    const t = taskById(id);
    if (t) {
      t.reporter = target.value.trim();
      commit();
    }
  } else if (kind === 'action-project') {
    const t = taskById(id);
    if (t && moveTask(t, target.value)) commit();
  } else if (kind === 'menu-date') {
    if (target.value) pickMenu(target.value);
  }
}

function onFocusOut(e: FocusEvent): void {
  const target = e.target as HTMLElement;
  if (target.matches?.('textarea.panel-title')) {
    // Title changes are already saved as you type; redraw so the list catches up.
    setTimeout(() => {
      if (!isTyping() || document.activeElement !== target) render();
    }, 0);
    return;
  }
  const key = target.dataset?.actionKey;
  if (key && ['add-project', 'add-notebook', 'add-person'].includes(key) && ui.adding) {
    const value = (target as HTMLInputElement).value;
    setTimeout(() => {
      if (!ui.adding) return;
      addFromSidebar(key.slice(4) as 'project' | 'notebook' | 'person', value);
    }, 0);
  }
}

function onMouseDown(e: MouseEvent): void {
  const target = e.target as HTMLElement;
  if (ui.menu && !target.closest('#menu-root') && !target.closest('[data-action$="-menu"]')) closeMenu();

  const resizer = target.closest<HTMLElement>('.resizer');
  if (resizer) startResize(e, resizer);
}

function startResize(e: MouseEvent, resizer: HTMLElement): void {
  e.preventDefault();
  e.stopPropagation();
  const col = resizer.dataset.col as ColKey;
  const list = resizer.closest<HTMLElement>('.list')!;
  const startX = e.clientX;
  const startW = ui.colW[col] ?? DEFAULT_W[col];
  document.body.classList.add('resizing');
  const move = (ev: MouseEvent) => {
    const w = Math.max(MIN_W, Math.round(startW + ev.clientX - startX));
    ui.colW[col] = w;
    list.style.setProperty(`--w-${col}`, `${w}px`);
  };
  const up = () => {
    document.removeEventListener('mousemove', move);
    document.removeEventListener('mouseup', up);
    document.body.classList.remove('resizing');
    lsSet(LS.colw, ui.colW);
  };
  document.addEventListener('mousemove', move);
  document.addEventListener('mouseup', up);
}

function onDblClick(e: MouseEvent): void {
  const resizer = (e.target as HTMLElement).closest<HTMLElement>('.resizer');
  if (!resizer) return;
  delete ui.colW[resizer.dataset.col as ColKey];
  lsSet(LS.colw, ui.colW);
  render();
}

function onDragStart(e: DragEvent): void {
  const row = (e.target as HTMLElement).closest<HTMLElement>('.lrow[data-id]');
  if (!row || !e.dataTransfer) return;
  draggingTaskId = row.dataset.id!;
  e.dataTransfer.effectAllowed = 'move';
  e.dataTransfer.setData('text/plain', draggingTaskId);
  row.classList.add('dragging');
}

function dropTarget(e: DragEvent): HTMLElement | null {
  const el = (e.target as HTMLElement).closest<HTMLElement>('[data-drop-project]');
  if (!el || !draggingTaskId) return null;
  const t = taskById(draggingTaskId);
  const p = projectById(el.dataset.dropProject!);
  return t && p && p.kind === kindOfTask(t) && p.id !== t.projectId ? el : null;
}

function onDragOver(e: DragEvent): void {
  const el = dropTarget(e);
  document.querySelectorAll('.drop-hover').forEach((x) => x !== el && x.classList.remove('drop-hover'));
  if (!el) return;
  e.preventDefault();
  if (e.dataTransfer) e.dataTransfer.dropEffect = 'move';
  el.classList.add('drop-hover');
}

function onDrop(e: DragEvent): void {
  const el = dropTarget(e);
  if (!el) return;
  e.preventDefault();
  const t = taskById(draggingTaskId!);
  draggingTaskId = null;
  if (t && moveTask(t, el.dataset.dropProject!)) commit();
}

function onDragEnd(): void {
  draggingTaskId = null;
  document.querySelectorAll('.drop-hover, .dragging').forEach((x) => x.classList.remove('drop-hover', 'dragging'));
}

// ================================================================ boot (kept last: everything it uses is defined above)

function boot(): void {
  ui.view = lsGet<View | null>(LS.view, null);
  const savedCollapsed = lsGet<Record<string, boolean>>(LS.collapsed, {});
  for (const k of ['projects', 'notes', 'meetings']) if (typeof savedCollapsed[k] === 'boolean') ui.collapsed[k] = savedCollapsed[k];
  ui.collapsed.tags = true; // always start collapsed
  ui.collapsed.people = true;
  ui.colW = lsGet<Partial<Record<ColKey, number>>>(LS.colw, {});

  readCache();
  ensureOwner();
  ui.sync = SHEETS_API_URL ? 'loading' : 'local';

  installRichText({
    onChange(key, html) {
      const [kind, id] = [key.slice(0, key.indexOf(':')), key.slice(key.indexOf(':') + 1)];
      if (kind === 'comment') {
        commentDrafts.set(id, html);
        return;
      }
      if (kind === 'desc') {
        const t = taskById(id);
        if (!t || t.description === html) return;
        t.description = html;
      } else if (kind === 'meeting') {
        const m = meetingById(id);
        if (!m || m.notes === html) return;
        m.notes = html;
      } else return;
      commit({ render: false });
    },
    onSubmit(key) {
      if (key.startsWith('comment:')) postComment(key.slice('comment:'.length));
    },
  });

  document.addEventListener('click', onClick);
  document.addEventListener('keydown', onKeyDown);
  document.addEventListener('input', onInput);
  document.addEventListener('change', onChange);
  document.addEventListener('focusout', onFocusOut);
  document.addEventListener('mousedown', onMouseDown);
  document.addEventListener('dblclick', onDblClick);
  document.addEventListener('dragstart', onDragStart);
  document.addEventListener('dragover', onDragOver);
  document.addEventListener('drop', onDrop);
  document.addEventListener('dragend', onDragEnd);
  window.addEventListener('resize', () => {
    closeMenu();
    autoGrowAll();
  });
  // The panel slides open; measure the title again once it has its real width.
  document.getElementById('app')!.addEventListener('transitionend', (e) => {
    if (e.target === e.currentTarget) autoGrowAll();
  });
  document.addEventListener('scroll', (e) => {
    if (!(e.target instanceof Element && e.target.closest('#menu-root'))) closeMenu();
  }, true);

  render();

  if (SHEETS_API_URL) {
    void runLoad();
    window.addEventListener('focus', () => void runLoad());
    setInterval(() => {
      if (document.visibilityState === 'visible') void runLoad();
    }, POLL_MS);
  }
  // Keep "x ago" and due labels fresh.
  setInterval(() => {
    if (!isTyping() && !ui.menu && !ui.dialogOpen) render();
  }, 5 * 60_000);
}

boot();

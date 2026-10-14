// The AI transcript step: read a .txt/.vtt/.srt file and have
// @canva-ct/genai (Agentic chat, signed in with the Canva login) summarise it
// through a single tool call, save_meeting_notes.

export interface ActionDraft {
  title: string;
  owner: string;
  due: string; // YYYY-MM-DD or ''
  details: string;
}

export interface MeetingDraft {
  title: string;
  date: string; // YYYY-MM-DD
  attendees: string[];
  notesHtml: string;
  actions: ActionDraft[];
}

export const SAVE_TOOL = {
  name: 'save_meeting_notes',
  description: 'Save the summary of a meeting transcript: its title, date, attendees, notes and action items.',
  input_schema: {
    type: 'object',
    properties: {
      title: { type: 'string', description: 'Short meeting title' },
      date: { type: 'string', description: 'Meeting date as YYYY-MM-DD, or empty if unknown' },
      attendees: { type: 'array', items: { type: 'string' }, description: 'Full names of people who spoke or attended' },
      notes_html: {
        type: 'string',
        description:
          'Meeting notes as simple HTML using only p, ul, ol, li, b, i, br. Cover decisions, key discussion points and open questions.',
      },
      action_items: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            title: { type: 'string', description: 'What needs to be done, starting with a verb' },
            owner: { type: 'string', description: 'Full name of the person responsible, or empty' },
            due: { type: 'string', description: 'Due date as YYYY-MM-DD, or empty' },
            details: { type: 'string', description: 'Any useful context, or empty' },
          },
          required: ['title'],
        },
      },
    },
    required: ['title', 'notes_html', 'action_items'],
  },
} as const;

const MAX_CHARS = 180_000;

/** Reads a transcript file and strips cue numbers and timestamps from .vtt/.srt. */
export async function readTranscript(file: File): Promise<string> {
  const raw = await file.text();
  const name = file.name.toLowerCase();
  if (!name.endsWith('.vtt') && !name.endsWith('.srt')) return raw.trim();
  const out: string[] = [];
  let last = '';
  for (const line of raw.split(/\r?\n/)) {
    const l = line.trim();
    if (!l) continue;
    if (/^WEBVTT/i.test(l) || /^(NOTE|STYLE|REGION)\b/.test(l)) continue;
    if (/^\d+$/.test(l)) continue;
    if (/-->/.test(l)) continue;
    const text = l.replace(/<v\s+([^>]+)>/gi, '$1: ').replace(/<[^>]+>/g, '').trim();
    if (text && text !== last) out.push(text);
    last = text;
  }
  return out.join('\n');
}

function systemPrompt(people: string[], today: string): string {
  return [
    'You summarise meeting transcripts for a personal task manager.',
    `Today is ${today}. If the transcript doesn't say when the meeting happened, use today's date.`,
    'Resolve relative due dates ("by Friday") against the meeting date.',
    people.length
      ? `Known team members (use these exact spellings when they match): ${people.join(', ')}.`
      : '',
    'Only list real, concrete action items. Always reply by calling save_meeting_notes exactly once.',
  ]
    .filter(Boolean)
    .join('\n');
}

/** Finds the save_meeting_notes call in whatever shape the chat result takes. */
function findToolInput(value: unknown, depth = 0): Record<string, unknown> | null {
  if (!value || typeof value !== 'object' || depth > 8) return null;
  const v = value as Record<string, unknown>;
  const name = v.name ?? v.toolName ?? v.tool_name;
  if (name === SAVE_TOOL.name) {
    let input = v.input ?? v.arguments ?? v.args ?? v.parameters;
    if (typeof input === 'string') {
      try {
        input = JSON.parse(input);
      } catch {
        input = null;
      }
    }
    if (input && typeof input === 'object') return input as Record<string, unknown>;
  }
  for (const child of Array.isArray(value) ? value : Object.values(v)) {
    const found = findToolInput(child, depth + 1);
    if (found) return found;
  }
  return null;
}

const str = (v: unknown): string => (typeof v === 'string' ? v.trim() : v == null ? '' : String(v).trim());

function isoDate(v: unknown, fallback = ''): string {
  const s = str(v);
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : fallback;
}

function toDraft(input: Record<string, unknown>, today: string): MeetingDraft {
  const actions = Array.isArray(input.action_items) ? input.action_items : [];
  return {
    title: str(input.title) || 'Meeting',
    date: isoDate(input.date, today),
    attendees: (Array.isArray(input.attendees) ? input.attendees : []).map(str).filter(Boolean),
    notesHtml: str(input.notes_html ?? input.notes),
    actions: actions
      .map((a: any) => ({ title: str(a?.title), owner: str(a?.owner), due: isoDate(a?.due), details: str(a?.details) }))
      .filter((a) => a.title),
  };
}

/**
 * The one place that talks to @canva-ct/genai. If the package's API differs
 * from what's assumed here, this is the only function to adjust: it must send
 * the prompt with the tool and return the raw result.
 */
async function callAgent(system: string, prompt: string): Promise<unknown> {
  const mod: any = await import('@canva-ct/genai');
  if (mod.__stub) {
    throw new Error(
      'In-app AI isn\'t set up: install @canva-ct/genai (see .npmrc) and sign in with your Canva login. You can still add meetings through the sheet\'s Meetings tab.',
    );
  }
  const Chat = mod.AgenticChat ?? mod.default?.AgenticChat ?? mod.default;
  if (!Chat) throw new Error('@canva-ct/genai: AgenticChat not found');
  const chat = typeof Chat === 'function' && /^class\b/.test(Function.prototype.toString.call(Chat))
    ? new Chat({ system, tools: [SAVE_TOOL], toolChoice: { type: 'tool', name: SAVE_TOOL.name } })
    : await Chat({ system, tools: [SAVE_TOOL], toolChoice: { type: 'tool', name: SAVE_TOOL.name } });
  const send = chat.send ?? chat.sendMessage ?? chat.run;
  if (typeof send !== 'function') throw new Error('@canva-ct/genai: no send method on the chat');
  return await send.call(chat, prompt);
}

/** Summarises a transcript into a meeting draft. Throws a readable error on failure. */
export async function summariseTranscript(
  transcript: string,
  fileName: string,
  people: string[],
  today: string,
): Promise<MeetingDraft> {
  if (!transcript.trim()) throw new Error('That file is empty.');
  const text = transcript.length > MAX_CHARS ? transcript.slice(0, MAX_CHARS) + '\n[transcript truncated]' : transcript;
  let result: unknown;
  try {
    result = await callAgent(systemPrompt(people, today), `Transcript file: ${fileName}\n\n${text}`);
  } catch (err: any) {
    const status = err?.status ?? err?.response?.status;
    if (status === 401 || /\b401\b/.test(String(err?.message))) {
      throw new Error('Your Canva login has expired (401). Sign in again — in the Flying Fox preview, restart Flying Fox.');
    }
    throw err instanceof Error ? err : new Error(String(err));
  }
  const input = findToolInput(result);
  if (!input) throw new Error('The AI didn\'t return meeting notes. Try again.');
  return toDraft(input, today);
}

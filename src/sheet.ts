// The Google Sheet back end: the web app URL and the read/save calls.
//
// Paste the /exec link of your Apps Script deployment below (or set
// VITE_SHEETS_API_URL in a .env file). Leave it empty to run browser-only.
export const SHEETS_API_URL: string =
  (import.meta.env.VITE_SHEETS_API_URL as string | undefined) || '';

export interface SheetPayload {
  projects?: unknown[];
  tasks?: unknown[];
  tags?: unknown[];
  people?: unknown[];
  comments?: unknown[];
  meetings?: unknown[];
}

/** What the script returns. A field is undefined when an older script doesn't send it. */
export interface SheetResponse extends SheetPayload {
  ok?: boolean;
  error?: string;
}

const TIMEOUT_MS = 30000;

async function request(init?: RequestInit): Promise<SheetResponse> {
  if (!SHEETS_API_URL) throw new Error('No sheet URL configured');
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), TIMEOUT_MS);
  try {
    const url = init?.method === 'POST' ? SHEETS_API_URL : `${SHEETS_API_URL}${SHEETS_API_URL.includes('?') ? '&' : '?'}t=${Date.now()}`;
    const res = await fetch(url, { redirect: 'follow', cache: 'no-store', ...init, signal: ctrl.signal });
    if (!res.ok) throw new Error(`Sheet responded ${res.status}`);
    const text = await res.text();
    let json: SheetResponse;
    try {
      json = JSON.parse(text);
    } catch {
      throw new Error('The sheet did not return JSON. Is the web app deployed with access "Anyone"?');
    }
    if (json && json.ok === false) throw new Error(json.error || 'The sheet reported an error');
    return json;
  } finally {
    clearTimeout(timer);
  }
}

/** Reads everything from the sheet. */
export function loadSheet(): Promise<SheetResponse> {
  return request({ method: 'GET' });
}

/**
 * Saves a complete snapshot. Sent as text/plain so the browser doesn't need a
 * CORS preflight, which Apps Script can't answer.
 */
export async function saveSheet(payload: SheetPayload & Record<string, unknown>): Promise<void> {
  await request({
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(payload),
  });
}

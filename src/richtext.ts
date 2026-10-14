// Rich text: sanitising, plain-text conversion, linkifying and the editor.
// Stored format is a small set of basic HTML tags. Everything is cleaned
// before it is shown or saved.

const ALLOWED = new Set([
  'B', 'STRONG', 'I', 'EM', 'U', 'S', 'A', 'UL', 'OL', 'LI', 'P', 'BR', 'DIV', 'CODE', 'BLOCKQUOTE',
]);

// Removed entirely, contents included.
const DROPPED = new Set([
  'SCRIPT', 'STYLE', 'IFRAME', 'IMG',
  'OBJECT', 'EMBED', 'TEMPLATE', 'HEAD', 'TITLE', 'META', 'LINK', 'SVG', 'MATH', 'NOSCRIPT',
]);

// Block-level tags we don't keep but whose line breaks we want to preserve.
const AS_DIV = new Set([
  'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'PRE', 'SECTION', 'ARTICLE', 'HEADER', 'FOOTER',
  'TR', 'TABLE', 'FIGURE', 'ASIDE', 'NAV', 'MAIN', 'DL', 'DT', 'DD',
]);

// http(s)/www links, plus bare domains on common TLDs ("canva.com/brief"), but not emails.
const URL_RE = new RegExp(
  String.raw`(?<![@\w.\/-])((?:https?:\/\/|www\.)[^\s<>"']*[^\s<>"'.,;:!?)\]}]` +
    String.raw`|[a-z0-9-]+(?:\.[a-z0-9-]+)*\.(?:com|org|net|io|co|ai|dev|app|design|me|nz|au|uk|edu|gov|so|ly|gl|ca|de|fr|es|br)\b(?![@\w-])(?:\/[^\s<>"']*[^\s<>"'.,;:!?)\]}])?)`,
  'gi',
);

export function escapeHtml(s: unknown): string {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/** Returns the URL if it uses http, https or mailto; otherwise null. */
export function safeUrl(href: string): string | null {
  const s = (href || '').trim();
  if (!/^(https?:|mailto:)/i.test(s)) return null;
  try {
    const u = new URL(s);
    if (!['http:', 'https:', 'mailto:'].includes(u.protocol)) return null;
    return s;
  } catch {
    return null;
  }
}

/** Accepts what a person types ("canva.com", "https://…", "name@x.com") and returns a safe URL or null. */
export function normalizeLinkInput(input: string, allowMailto = true): string | null {
  let s = (input || '').trim();
  if (!s) return null;
  if (/^[a-z][a-z0-9+.-]*:/i.test(s)) {
    const ok = safeUrl(s);
    if (!ok) return null;
    if (!allowMailto && /^mailto:/i.test(ok)) return null;
    return ok;
  }
  if (allowMailto && /^[^\s@/]+@[^\s@/]+\.[^\s@/]+$/.test(s)) return safeUrl('mailto:' + s);
  if (/\s/.test(s) || !/\./.test(s)) return null;
  s = 'https://' + s.replace(/^\/+/, '');
  return safeUrl(s);
}

function cleanInto(src: Node, dst: Node): void {
  for (const n of Array.from(src.childNodes)) {
    if (n.nodeType === Node.TEXT_NODE) {
      dst.appendChild(document.createTextNode(n.textContent || ''));
      continue;
    }
    if (n.nodeType !== Node.ELEMENT_NODE) continue;
    const el = n as Element;
    const tag = el.tagName.toUpperCase();
    if (DROPPED.has(tag)) continue;
    if (ALLOWED.has(tag)) {
      if (tag === 'A') {
        const url = safeUrl(el.getAttribute('href') || '');
        if (!url) {
          cleanInto(el, dst);
          continue;
        }
        const a = document.createElement('a');
        a.setAttribute('href', url);
        a.setAttribute('target', '_blank');
        a.setAttribute('rel', 'noopener noreferrer');
        cleanInto(el, a);
        dst.appendChild(a);
        continue;
      }
      const c = document.createElement(tag.toLowerCase());
      cleanInto(el, c);
      dst.appendChild(c);
    } else if (AS_DIV.has(tag)) {
      const c = document.createElement('div');
      cleanInto(el, c);
      dst.appendChild(c);
    } else {
      cleanInto(el, dst);
    }
  }
}

/** Keeps only the allowed tags; links keep only http(s)/mailto hrefs and always open in a new tab. */
export function sanitizeHtml(html: string): string {
  if (!html) return '';
  const doc = new DOMParser().parseFromString(`<!doctype html><body>${html}</body>`, 'text/html');
  const out = document.createElement('div');
  cleanInto(doc.body, out);
  return out.innerHTML;
}

export function looksLikeHtml(s: string): boolean {
  return /<\/?[a-z][\s\S]*?>/i.test(s);
}

function linkifyText(escaped: string): string {
  return escaped.replace(URL_RE, (m) => {
    const raw = m.replace(/&amp;/g, '&');
    const url = normalizeLinkInput(raw);
    return url ? `<a href="${escapeHtml(url)}">${m}</a>` : m;
  });
}

/** Converts stored text (HTML or old plain text) into clean HTML for display or editing. */
export function toDisplayHtml(stored: string | null | undefined): string {
  const s = String(stored ?? '');
  if (!s.trim()) return '';
  if (looksLikeHtml(s)) return sanitizeHtml(s);
  const html = linkifyText(escapeHtml(s)).replace(/\r?\n/g, '<br>');
  return sanitizeHtml(html);
}

export function htmlToText(html: string): string {
  const doc = new DOMParser().parseFromString(`<body>${html || ''}</body>`, 'text/html');
  return (doc.body.textContent || '').replace(/ /g, ' ');
}

export function isEmptyHtml(html: string): boolean {
  if (!html) return true;
  if (/<li[\s>]/i.test(html)) return false;
  return htmlToText(html).trim() === '';
}

/** Turns plain URLs inside an element (outside existing links) into links. Returns true if anything changed. */
export function linkifyElement(root: HTMLElement): boolean {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const targets: Text[] = [];
  let node: Node | null;
  while ((node = walker.nextNode())) {
    const t = node as Text;
    if ((t.parentElement && t.parentElement.closest('a')) || !t.data) continue;
    URL_RE.lastIndex = 0;
    if (URL_RE.test(t.data)) targets.push(t);
  }
  for (const t of targets) {
    const frag = document.createDocumentFragment();
    let last = 0;
    URL_RE.lastIndex = 0;
    let m: RegExpExecArray | null;
    while ((m = URL_RE.exec(t.data))) {
      const url = normalizeLinkInput(m[0]);
      if (!url) continue;
      frag.appendChild(document.createTextNode(t.data.slice(last, m.index)));
      const a = document.createElement('a');
      a.href = url;
      a.target = '_blank';
      a.rel = 'noopener noreferrer';
      a.textContent = m[0];
      frag.appendChild(a);
      last = m.index + m[0].length;
    }
    frag.appendChild(document.createTextNode(t.data.slice(last)));
    t.replaceWith(frag);
  }
  return targets.length > 0;
}

// ---------------------------------------------------------------- editor

const TOOLBAR: { cmd: string; icon: string; title: string }[] = [
  { cmd: 'bold', icon: 'lucide:bold', title: 'Bold (⌘B)' },
  { cmd: 'italic', icon: 'lucide:italic', title: 'Italic (⌘I)' },
  { cmd: 'underline', icon: 'lucide:underline', title: 'Underline (⌘U)' },
  { cmd: 'strikeThrough', icon: 'lucide:strikethrough', title: 'Strikethrough' },
  { cmd: 'insertUnorderedList', icon: 'lucide:list', title: 'Bulleted list' },
  { cmd: 'insertOrderedList', icon: 'lucide:list-ordered', title: 'Numbered list' },
  { cmd: 'link', icon: 'lucide:link', title: 'Link (⌘K)' },
];

export interface EditorOptions {
  placeholder?: string;
  compact?: boolean;
  minHeight?: number;
}

/** Markup for a rich text editor. `key` identifies it to the change handlers. */
export function editorHtml(key: string, html: string, opts: EditorOptions = {}): string {
  const body = toDisplayHtml(html);
  const empty = isEmptyHtml(body);
  const buttons = TOOLBAR.map(
    (b) =>
      `<button type="button" class="rt-btn" data-rt-cmd="${b.cmd}" title="${escapeHtml(b.title)}" tabindex="-1"><iconify-icon icon="${b.icon}"></iconify-icon></button>`,
  ).join('');
  const style = opts.minHeight ? ` style="min-height:${opts.minHeight}px"` : '';
  return `<div class="rt${opts.compact ? ' rt-compact' : ''}" data-rt="${escapeHtml(key)}">
    <div class="rt-toolbar">${buttons}</div>
    <div class="rt-body rt-content${empty ? ' is-empty' : ''}" contenteditable="true" spellcheck="true"
      data-fk="rt:${escapeHtml(key)}" data-placeholder="${escapeHtml(opts.placeholder || '')}"${style}>${body}</div>
  </div>`;
}

export interface RichTextHandlers {
  /** Called as the person types (final=false) and when they leave the box (final=true). */
  onChange(key: string, html: string, final: boolean): void;
  /** ⌘/Ctrl+Enter inside an editor. */
  onSubmit?(key: string): void;
}

let savedRange: Range | null = null;

function editorOf(el: EventTarget | null): { wrap: HTMLElement; body: HTMLElement; key: string } | null {
  if (!(el instanceof Element)) return null;
  const wrap = el.closest<HTMLElement>('.rt');
  if (!wrap) return null;
  const body = wrap.querySelector<HTMLElement>('.rt-body');
  if (!body) return null;
  return { wrap, body, key: wrap.dataset.rt || '' };
}

function updateEmpty(body: HTMLElement): void {
  body.classList.toggle('is-empty', isEmptyHtml(body.innerHTML));
}

function emit(h: RichTextHandlers, ed: { body: HTMLElement; key: string }, final: boolean): void {
  updateEmpty(ed.body);
  h.onChange(ed.key, sanitizeHtml(ed.body.innerHTML), final);
}

function closeLinkBox(wrap: HTMLElement): void {
  wrap.querySelector('.rt-linkbox')?.remove();
}

function openLinkBox(wrap: HTMLElement, body: HTMLElement): void {
  closeLinkBox(wrap);
  const sel = window.getSelection();
  savedRange = sel && sel.rangeCount && body.contains(sel.getRangeAt(0).commonAncestorContainer)
    ? sel.getRangeAt(0).cloneRange()
    : null;
  const box = document.createElement('div');
  box.className = 'rt-linkbox';
  box.innerHTML = `<iconify-icon icon="lucide:link"></iconify-icon>
    <input type="text" placeholder="Paste or type a link, e.g. canva.com" />
    <button type="button" class="btn btn-primary btn-sm" data-rt-link-apply>Add</button>`;
  wrap.appendChild(box);
  const input = box.querySelector('input')!;
  if (savedRange && !savedRange.collapsed) {
    const txt = savedRange.toString().trim();
    if (normalizeLinkInput(txt)) input.value = txt;
  }
  input.focus();
}

function applyLink(h: RichTextHandlers, wrap: HTMLElement): void {
  const ed = editorOf(wrap);
  const input = wrap.querySelector<HTMLInputElement>('.rt-linkbox input');
  if (!ed || !input) return;
  const url = normalizeLinkInput(input.value);
  if (!url) {
    input.classList.add('invalid');
    input.focus();
    return;
  }
  closeLinkBox(wrap);
  ed.body.focus();
  const sel = window.getSelection();
  if (sel) {
    sel.removeAllRanges();
    if (savedRange) sel.addRange(savedRange);
    else {
      const r = document.createRange();
      r.selectNodeContents(ed.body);
      r.collapse(false);
      sel.addRange(r);
    }
  }
  if (!savedRange || savedRange.collapsed) {
    document.execCommand('insertHTML', false, `<a href="${escapeHtml(url)}">${escapeHtml(input.value.trim())}</a>&nbsp;`);
  } else {
    document.execCommand('createLink', false, url);
  }
  ed.body.querySelectorAll('a').forEach((a) => {
    a.setAttribute('target', '_blank');
    a.setAttribute('rel', 'noopener noreferrer');
  });
  savedRange = null;
  emit(h, ed, true);
}

/** Installs document-level listeners that make every `.rt` editor work. Call once. */
export function installRichText(h: RichTextHandlers): void {
  document.addEventListener('mousedown', (e) => {
    const btn = (e.target as Element).closest?.('[data-rt-cmd]') as HTMLElement | null;
    if (!btn) return;
    e.preventDefault(); // keep the selection in the editor
    const ed = editorOf(btn);
    if (!ed) return;
    const cmd = btn.dataset.rtCmd!;
    if (document.activeElement !== ed.body && !ed.body.contains(document.activeElement)) ed.body.focus();
    if (cmd === 'link') {
      openLinkBox(ed.wrap, ed.body);
      return;
    }
    document.execCommand(cmd, false);
    emit(h, ed, false);
  });

  document.addEventListener('click', (e) => {
    const t = e.target as Element;
    if (t.closest?.('[data-rt-link-apply]')) {
      const wrap = t.closest<HTMLElement>('.rt');
      if (wrap) applyLink(h, wrap);
      return;
    }
    const a = t.closest?.('a[href]') as HTMLAnchorElement | null;
    if (a && a.closest('.rt-body')) {
      e.preventDefault();
      const url = safeUrl(a.getAttribute('href') || '');
      if (url) window.open(url, '_blank', 'noopener,noreferrer');
    }
  });

  document.addEventListener('input', (e) => {
    const t = e.target as HTMLElement;
    if (!t.classList?.contains('rt-body')) return;
    const ed = editorOf(t);
    if (ed) emit(h, ed, false);
  });

  document.addEventListener('focusout', (e) => {
    const t = e.target as HTMLElement;
    if (!t.classList?.contains('rt-body')) return;
    const ed = editorOf(t);
    if (!ed) return;
    // Leaving for the link box isn't leaving the editor.
    const next = e.relatedTarget as Element | null;
    if (next && ed.wrap.contains(next)) return;
    linkifyElement(ed.body);
    emit(h, ed, true);
  });

  document.addEventListener('paste', (e) => {
    const t = e.target as Element;
    const body = t.closest?.('.rt-body') as HTMLElement | null;
    if (!body) return;
    e.preventDefault();
    const text = e.clipboardData?.getData('text/plain') ?? '';
    document.execCommand('insertText', false, text);
  });

  document.addEventListener('keydown', (e) => {
    const t = e.target as HTMLElement;
    if (t.matches?.('.rt-linkbox input')) {
      const wrap = t.closest<HTMLElement>('.rt')!;
      if (e.key === 'Enter') {
        e.preventDefault();
        applyLink(h, wrap);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        closeLinkBox(wrap);
        wrap.querySelector<HTMLElement>('.rt-body')?.focus();
      }
      return;
    }
    if (!t.classList?.contains('rt-body')) return;
    const mod = e.metaKey || e.ctrlKey;
    if (mod && e.key.toLowerCase() === 'k') {
      e.preventDefault();
      const ed = editorOf(t);
      if (ed) openLinkBox(ed.wrap, ed.body);
    } else if (mod && e.key === 'Enter') {
      e.preventDefault();
      const ed = editorOf(t);
      if (ed) {
        emit(h, ed, true);
        h.onSubmit?.(ed.key);
      }
    }
  }, true);
}

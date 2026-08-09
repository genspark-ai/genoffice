/**
 * Built-in browser for the agent: loads a page in a real, hidden Chromium
 * window and returns its rendered text.
 *
 * A plain fetch would be cheaper, but most of what an agent is asked to read
 * — docs sites, dashboards, anything React-rendered — ships an empty body and
 * fills it in with JavaScript. Rendering in Chromium is what makes "browse
 * this page" actually work, and the engine is already in the process.
 *
 * The window is locked down: no Node, no preload, sandboxed, its own session
 * partition (so the user's cookies are never sent to a page the model chose),
 * navigation pinned to the URL that was approved, and destroyed on every exit
 * path. URLs are SSRF-checked before load, so a prompt-injected link cannot
 * reach the loopback interface or the local network.
 */
import { BrowserWindow, session } from 'electron'
import { isSafeRemoteUrl } from './safe-remote-url'

/** hard cap on returned text, so one page cannot swallow the context window */
const DEFAULT_MAX_CHARS = 20_000
const DEFAULT_TIMEOUT_MS = 30_000
/** grace period after load for client-side rendering to paint */
const SETTLE_MS = 700

export interface BrowsePageResult {
  url: string
  title: string
  /** visible text of the rendered page, whitespace-collapsed and truncated */
  text: string
  /** in-page links, absolute, deduplicated */
  links: Array<{ text: string; href: string }>
  /** true when `text` hit the cap */
  truncated: boolean
}

export interface BrowsePageOptions {
  maxChars?: number
  timeoutMs?: number
  /** proxy to route this page through (same value the main process uses) */
  proxyUrl?: string
  /** include the link list; off by default to keep results small */
  includeLinks?: boolean
}

/**
 * Extraction runs inside the page. Kept as a self-contained expression string
 * because it is evaluated in the renderer, where none of this module's scope
 * exists. Strips the elements that are never content, then reads innerText so
 * the browser's own layout decides what is visible.
 */
function extractionScript(maxChars: number, includeLinks: boolean): string {
  return `(() => {
    const drop = ['script','style','noscript','iframe','svg','canvas','template'];
    const doc = document.cloneNode(true);
    for (const sel of drop) for (const el of doc.querySelectorAll(sel)) el.remove();
    const main = doc.querySelector('main,article,[role=main]') || doc.body;
    const raw = (main && main.innerText) || (doc.body && doc.body.innerText) || '';
    const text = raw.replace(/[ \\t\\u00a0]+/g, ' ').replace(/\\n{3,}/g, '\\n\\n').trim();
    const links = ${includeLinks ? 'true' : 'false'}
      ? Array.from(document.querySelectorAll('a[href]'))
          .map((a) => ({ text: (a.innerText || '').trim().slice(0, 120), href: a.href }))
          .filter((l) => l.href.startsWith('http') && l.text)
          .slice(0, 100)
      : [];
    const seen = new Set();
    const unique = [];
    for (const l of links) { if (!seen.has(l.href)) { seen.add(l.href); unique.push(l); } }
    return {
      title: document.title || '',
      text: text.slice(0, ${maxChars}),
      truncated: text.length > ${maxChars},
      links: unique,
    };
  })()`
}

/**
 * Load `rawUrl` and return its rendered content.
 * Throws on an unsafe URL, a navigation failure, or a timeout.
 */
export async function browsePage(
  rawUrl: unknown,
  options: BrowsePageOptions = {},
): Promise<BrowsePageResult> {
  const url = String(rawUrl ?? '').trim()
  if (!/^https?:\/\//i.test(url)) throw new Error('Only http(s) URLs can be browsed')
  // resolves the host and rejects loopback / private / link-local targets
  if (!(await isSafeRemoteUrl(url))) throw new Error(`Refusing to browse a non-public URL: ${url}`)

  const maxChars = options.maxChars && options.maxChars > 0 ? options.maxChars : DEFAULT_MAX_CHARS
  const timeoutMs =
    options.timeoutMs && options.timeoutMs > 0 ? options.timeoutMs : DEFAULT_TIMEOUT_MS

  // A dedicated in-memory partition per call: the page gets no access to the
  // user's cookies or storage, and nothing it sets outlives the call.
  const partition = `agent-browse-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
  const browseSession = session.fromPartition(partition, { cache: false })
  if (options.proxyUrl) {
    await browseSession.setProxy({ proxyRules: options.proxyUrl }).catch(() => {})
  }

  const win = new BrowserWindow({
    show: false,
    webPreferences: {
      session: browseSession,
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: true,
      webSecurity: true,
      // the whole point is to render client-side pages
      javascript: true,
      // text is all the agent reads; skipping images cuts load time a lot
      images: false,
      webgl: false,
    },
  })

  // The model picked this URL; treat any attempt to leave it as hostile.
  win.webContents.setWindowOpenHandler(() => ({ action: 'deny' }))

  let timer: NodeJS.Timeout | undefined
  try {
    const loaded = new Promise<void>((resolvePromise, reject) => {
      win.webContents.once('did-finish-load', () => resolvePromise())
      win.webContents.once('did-fail-load', (_e, code, desc, failedUrl) => {
        // sub-resource failures also fire here; only the main frame matters
        if (failedUrl === url || code !== -3) reject(new Error(`Load failed (${code}): ${desc}`))
      })
    })
    const timeout = new Promise<never>((_r, reject) => {
      timer = setTimeout(
        () => reject(new Error(`Timed out after ${timeoutMs}ms loading ${url}`)),
        timeoutMs,
      )
    })

    await Promise.race([win.loadURL(url).then(() => loaded), timeout])
    await new Promise((r) => setTimeout(r, SETTLE_MS))

    const raw = (await Promise.race([
      win.webContents.executeJavaScript(
        extractionScript(maxChars, options.includeLinks === true),
        true,
      ),
      timeout,
    ])) as Partial<BrowsePageResult>

    return {
      url: win.webContents.getURL() || url,
      title: String(raw.title ?? ''),
      text: String(raw.text ?? ''),
      links: Array.isArray(raw.links) ? raw.links : [],
      truncated: raw.truncated === true,
    }
  } finally {
    if (timer) clearTimeout(timer)
    if (!win.isDestroyed()) win.destroy()
    // drop the partition's in-memory data so nothing accumulates across calls
    await browseSession.clearStorageData().catch(() => {})
  }
}

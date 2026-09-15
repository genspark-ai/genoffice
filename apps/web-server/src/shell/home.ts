/**
 * Home (start page) channels — recents, starred, theme/language, account
 * status, analytics, update-channel, onboarding, github stars, stat paths,
 * browse, new-{doc/sheet/slide/markdown/pdf}, delete/rename/duplicate, reveal,
 * open-trash, cloud-projects. The recents and starred sets share the maps
 * declared in `common/state.ts` (`DOCS_RECENT`, `DOCS_STARRED`).
 */
import { existsSync, readFileSync, renameSync, statSync, unlinkSync, writeFileSync } from 'node:fs'
/** Map a DocInfo record to the RecentEntry shape the home renderer expects,
 *  stat-ing the file for size/mtime. Files that fail to stat are flagged
 *  `missing` instead of being dropped (mirrors the desktop behaviour). */
function toRecentEntry(d: {
  id: string
  path: string
  name: string
  openedAt?: number
  modified?: boolean
}): {
  path: string
  name: string
  ext: string
  mtimeMs: number
  sizeBytes: number
  starred: boolean
  missing?: boolean
} {
  const ext =
    d.path
      .split(/[\\./]/)
      .pop()
      ?.toLowerCase() ?? ''
  try {
    if (existsSync(d.path)) {
      const s = statSync(d.path)
      return {
        path: d.path,
        name: d.name,
        ext,
        mtimeMs: s.mtimeMs,
        sizeBytes: s.size,
        starred: DOCS_STARRED.has(d.path),
      }
    }
  } catch {}
  return {
    path: d.path,
    name: d.name,
    ext,
    mtimeMs: d.openedAt ?? 0,
    sizeBytes: 0,
    starred: DOCS_STARRED.has(d.path),
    missing: true,
  }
}

import { basename, dirname, extname, join } from 'node:path'
import { DATA_DIR, DOCS_RECENT, DOCS_STARRED, registerHandle } from '../common/index'

export function registerHomeHandlers(): void {
  registerHandle('home:get-app-version', () => '1.0.0')

  registerHandle('home:get-theme', () => 'light')
  registerHandle('home:set-theme', (_event: unknown, theme: unknown) => ({ ok: true, theme }))

  registerHandle('home:get-language', () => 'zh-CN')
  registerHandle('home:set-language', (_event: unknown, lang: unknown) => ({
    ok: true,
    language: lang,
  }))

  registerHandle('home:recents', (_event: unknown, args: unknown) => {
    const {
      offset = 0,
      limit = 50,
      ext,
    } = (args || {}) as { offset?: number; limit?: number; ext?: string }
    const all = [...DOCS_RECENT.values()]
    const filtered = ext ? all.filter((d) => d.path.toLowerCase().endsWith('.' + ext)) : all
    const sliced = filtered.slice(offset, offset + limit)
    const entries = sliced.map((d) => toRecentEntry(d))
    return {
      entries,
      total: filtered.length,
      totalAll: all.length,
    }
  })

  registerHandle('home:starred', (_event: unknown, args: unknown) => {
    const {
      offset = 0,
      limit = 50,
      ext,
    } = (args || {}) as { offset?: number; limit?: number; ext?: string }
    const all = [...DOCS_STARRED]
      .map((p) => DOCS_RECENT.get(p))
      .filter((d): d is NonNullable<typeof d> => Boolean(d))
    const filtered = ext ? all.filter((d) => d.path.toLowerCase().endsWith('.' + ext)) : all
    const sliced = filtered.slice(offset, offset + limit)
    const entries = sliced.map((d) => toRecentEntry(d))
    return {
      entries,
      total: filtered.length,
      totalAll: all.length,
    }
  })

  registerHandle('home:toggle-star', (_event: unknown, path: unknown) => {
    if (typeof path !== 'string' || !path) return { starred: false }
    if (DOCS_STARRED.has(path)) {
      DOCS_STARRED.delete(path)
      return { starred: false }
    }
    DOCS_STARRED.add(path)
    return { starred: true }
  })

  registerHandle('home:open-path', async (_event: unknown, path: unknown) => {
    if (typeof path !== 'string') return { ok: false, error: 'invalid path' }
    return { ok: true, path, opened: true }
  })

  registerHandle('home:remove-recent', (_event: unknown, paths: unknown) => {
    if (!Array.isArray(paths)) return { ok: false, removed: 0 }
    paths.forEach((p) => DOCS_RECENT.delete(p))
    return { ok: true, removed: paths.length }
  })

  registerHandle('home:delete-files', async (_event: unknown, paths: unknown) => {
    if (!Array.isArray(paths)) return { ok: false, deleted: 0 }
    paths.forEach((p) => {
      if (existsSync(p)) unlinkSync(p)
    })
    return { ok: true, deleted: paths.length }
  })

  registerHandle('home:duplicate-file', async (_event: unknown, path: unknown) => {
    if (typeof path !== 'string' || !existsSync(path)) return { ok: false, error: 'File not found' }
    const dir = dirname(path)
    const ext = extname(path)
    const base = basename(path, ext)
    const newPath = join(dir, `${base}-copy${ext}`)
    writeFileSync(newPath, readFileSync(path))
    return { ok: true, path: newPath }
  })

  registerHandle('home:rename-file', async (_event: unknown, path: unknown, newName: unknown) => {
    if (typeof path !== 'string' || typeof newName !== 'string' || !existsSync(path)) {
      return { ok: false, error: 'File not found' }
    }
    const dir = dirname(path)
    const newPath = join(dir, newName)
    renameSync(path, newPath)
    return { ok: true, path: newPath }
  })

  registerHandle('home:reveal-path', (_event: unknown, path: unknown) => {
    return { ok: true, path }
  })

  registerHandle('home:open-trash', () => ({ ok: true }))

  registerHandle('home:new-doc', () => {
    const id = `doc-${Date.now()}`
    return { id, path: join(DATA_DIR, `${id}.docx`) }
  })

  registerHandle('home:new-sheet', () => {
    const id = `sheet-${Date.now()}`
    return { id, path: join(DATA_DIR, `${id}.xlsx`) }
  })

  registerHandle('home:new-slide', () => {
    const id = `slide-${Date.now()}`
    return { id, path: join(DATA_DIR, `${id}.pptx`) }
  })

  registerHandle('home:new-markdown', () => {
    const id = `md-${Date.now()}`
    return { id, path: join(DATA_DIR, `${id}.md`) }
  })

  registerHandle('home:new-pdf', () => {
    const id = `pdf-${Date.now()}`
    const path = join(DATA_DIR, `${id}.pdf`)
    const objects = [
      '1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n',
      '2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n',
      '3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595.28 841.89] /Contents 4 0 R /Resources << >> >>\nendobj\n',
      '4 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n',
    ]
    let pdf = '%PDF-1.4\n'
    const offsets = [0]
    for (const object of objects) {
      offsets.push(Buffer.byteLength(pdf, 'binary'))
      pdf += object
    }
    const xrefOffset = Buffer.byteLength(pdf, 'binary')
    pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`
    for (const offset of offsets.slice(1)) pdf += `${String(offset).padStart(10, '0')} 00000 n \n`
    pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF\n`
    writeFileSync(path, Buffer.from(pdf, 'binary'))
    return { id, path }
  })

  registerHandle('home:new-html', () => {
    const id = `html-${Date.now()}`
    const path = join(DATA_DIR, `${id}.html`)
    const html = `<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8" />
  <title>新 HTML 文档</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; max-width: 720px; margin: 60px auto; padding: 0 24px; color: #1a1a1a; line-height: 1.7; }
    h1 { font-size: 32px; font-weight: 600; }
    .meta { color: #6b7280; font-size: 14px; }
  </style>
</head>
<body>
  <h1>新 HTML 文档</h1>
  <p class="meta">由 GenOffice 创建</p>
  <p>开始编辑你的 HTML 内容...</p>
</body>
</html>
`
    writeFileSync(path, html, 'utf-8')
    return { id, path }
  })

  /* Web build ships without a real login backend; surface a friendly
   * placeholder identity so the user chip + sidebar account block have
   * something to render. The renderer displays `email.split('@')[0]` as the
   * username, so the local-part is the visible name. Loaded from a small
   * JSON file when present so the user can rename themselves across
   * sessions; falls back to a sensible default. */
  const WEB_ACCOUNT_FILE = join(DATA_DIR, 'web-account.json')
  function loadWebAccount(): typeof WEB_ACCOUNT {
    try {
      if (existsSync(WEB_ACCOUNT_FILE)) {
        const raw = JSON.parse(readFileSync(WEB_ACCOUNT_FILE, 'utf-8'))
        if (raw && typeof raw.email === 'string') return { ...WEB_ACCOUNT, ...raw }
      }
    } catch {}
    return WEB_ACCOUNT
  }
  const WEB_ACCOUNT = {
    loggedIn: true,
    email: 'godlinchong@genoffice.ai',
    displayName: 'godlinchong',
    plan: 'pro',
    creditBalance: 1000,
  }
  registerHandle('home:account-status', () => loadWebAccount())
  registerHandle('home:account-set-name', (_event: unknown, name: unknown) => {
    if (typeof name !== 'string' || !name.trim()) return { ok: false }
    const clean = name.trim().slice(0, 64)
    const next = { ...loadWebAccount(), displayName: clean, email: clean + '@genoffice.ai' }
    try {
      writeFileSync(
        WEB_ACCOUNT_FILE,
        JSON.stringify(
          {
            email: next.email,
            displayName: next.displayName,
            plan: next.plan,
            creditBalance: next.creditBalance,
          },
          null,
          2,
        ),
      )
    } catch {}
    return { ok: true, account: next }
  })

  registerHandle('home:account-login', async (_event: unknown, _args: unknown) => {
    const acc = loadWebAccount()
    return { ok: true, email: acc.email, displayName: acc.displayName }
  })

  registerHandle('home:account-login-open-url', () => ({
    url: 'https://account.genspark.ai/login',
  }))

  registerHandle('home:account-logout', () => ({ ok: true }))

  /* Real GitHub star count for the About pane. Returns null (not a
   * placeholder number) when the API is unreachable, so the UI stays honest. */
  let cachedGithubStars: number | null = null
  registerHandle('home:github-stars', async () => {
    if (cachedGithubStars !== null) return cachedGithubStars
    try {
      const response = await fetch('https://api.github.com/repos/genspark-ai/genoffice', {
        headers: { Accept: 'application/vnd.github+json' },
        signal: AbortSignal.timeout(5000),
      })
      if (!response.ok) return null
      const body: unknown = await response.json()
      const count = (body as { stargazers_count?: unknown }).stargazers_count
      if (typeof count !== 'number' || !Number.isFinite(count)) return null
      cachedGithubStars = count
      return count
    } catch {
      return null
    }
  })

  registerHandle('home:get-analytics-enabled', () => true)
  registerHandle('home:set-analytics-enabled', (_event: unknown, enabled: unknown) => ({
    ok: true,
    enabled,
  }))

  registerHandle('home:get-default-save-dir', () => DATA_DIR)
  registerHandle('home:pick-default-save-dir', () => DATA_DIR)

  registerHandle('home:get-update-channel', () => 'stable')
  registerHandle('home:set-update-channel', (_event: unknown, channel: unknown) => ({
    ok: true,
    channel,
  }))

  /* ── Onboarding flag ─────────────────────────────────────────────
   * The web build has no real login / first-run backend, but the renderer
   * still wants to know whether to show the welcome overlay. Track the flag
   * in a small JSON file so onboarding shows once per fresh install and
   * stays dismissed thereafter. */
  const ONBOARDING_FILE = join(DATA_DIR, 'onboarding.json')
  function loadOnboarding(): boolean {
    try {
      if (existsSync(ONBOARDING_FILE)) {
        const raw = JSON.parse(readFileSync(ONBOARDING_FILE, 'utf-8'))
        return Boolean(raw?.seen)
      }
    } catch {}
    return false
  }
  registerHandle('home:onboarding-seen', () => loadOnboarding())
  registerHandle('home:set-onboarding-seen', (_event: unknown, seen: unknown) => {
    /* Renderer can call this with no args (the common "I'm done with
     * onboarding" path) or with an explicit boolean. Default to marking
     * the onboarding seen when called without arguments so the welcome
     * overlay stays dismissed after the user clicks skip / next. */
    const value = seen === undefined ? true : Boolean(seen)
    try {
      writeFileSync(ONBOARDING_FILE, JSON.stringify({ seen: value, setAt: Date.now() }, null, 2))
    } catch {}
    return value
  })

  registerHandle('home:star-prompt-should-show', () => ({ shouldShow: false }))
  registerHandle('home:star-prompt-action', (_event: unknown, action: unknown) => {
    return { ok: true, action }
  })

  /* ── Cloud projects ────────────────────────────────────────────────────
   * The Electron build syncs the signed-in user's Genspark projects through
   * the bundled `gsk` CLI. The standalone web server has no account session,
   * so it reports `available: false` — the home pane then renders its
   * sign-in / empty state instead of fabricated project rows. */
  const buildCloudSnapshot = () => ({
    available: false,
    projects: [],
    syncedAt: 0,
  })
  registerHandle('home:cloud-projects', () => buildCloudSnapshot())
  registerHandle('home:cloud-projects-cached', () => buildCloudSnapshot())

  registerHandle('home:open-cloud-project', (_event: unknown, projectUrl: unknown) => {
    return { ok: true, url: projectUrl }
  })

  registerHandle('home:open-gen-team', () => ({ ok: true }))
  registerHandle('home:open-credit-usage', () => ({ ok: true }))
  registerHandle('home:open-github-repo', () => ({ ok: true }))

  registerHandle('home:stat-paths', (_event: unknown, paths: unknown) => {
    return (Array.isArray(paths) ? paths : []).map((p) => ({
      path: p,
      exists: existsSync(p),
      size: existsSync(p) ? statSync(p).size : 0,
    }))
  })

  registerHandle('home:browse', () => ({ canceled: false, filePaths: [] }))
}

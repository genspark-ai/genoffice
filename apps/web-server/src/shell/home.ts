/**
 * Home (start page) channels — recents, starred, theme/language, account
 * status, analytics, update-channel, onboarding, github stars, stat paths,
 * browse, new-{doc/sheet/slide/markdown/pdf}, delete/rename/duplicate, reveal,
 * open-trash, cloud-projects. The recents and starred sets share the maps
 * declared in `common/state.ts` (`DOCS_RECENT`, `DOCS_STARRED`).
 */
import { existsSync, readFileSync, renameSync, statSync, unlinkSync, writeFileSync } from 'node:fs'
import { basename, dirname, extname, join } from 'node:path'
import {
  DATA_DIR,
  DOCS_RECENT,
  DOCS_STARRED,
  registerHandle,
} from '../common/index.js'

export function registerHomeHandlers(): void {
  registerHandle('home:get-app-version', () => '1.0.0')

  registerHandle('home:get-theme', () => 'light')
  registerHandle('home:set-theme', (_event: unknown, theme: unknown) => ({ ok: true, theme }))

  registerHandle('home:get-language', () => 'zh-CN')
  registerHandle('home:set-language', (_event: unknown, lang: unknown) => ({ ok: true, language: lang }))

  registerHandle('home:recents', (_event: unknown, args: unknown) => {
    const { limit = 20 } = (args || {}) as { limit?: number }
    return {
      items: [...DOCS_RECENT.values()].slice(0, limit),
      total: DOCS_RECENT.size,
    }
  })

  registerHandle('home:starred', (_event: unknown, args: unknown) => {
    const { limit = 20 } = (args || {}) as { limit?: number }
    return {
      items: [...DOCS_STARRED].slice(0, limit),
      total: DOCS_STARRED.size,
    }
  })

  registerHandle('home:toggle-star', (_event: unknown, args: unknown) => {
    const { path } = (args || {}) as { path: string }
    if (DOCS_STARRED.has(path)) {
      DOCS_STARRED.delete(path)
      return { starred: false }
    }
    DOCS_STARRED.add(path)
    return { starred: true }
  })

  registerHandle('home:open-path', async (_event: unknown, args: unknown) => {
    const { path } = (args || {}) as { path: string }
    return { ok: true, path, opened: true }
  })

  registerHandle('home:remove-recent', (_event: unknown, args: unknown) => {
    const { paths } = (args || {}) as { paths: string[] }
    paths?.forEach(p => DOCS_RECENT.delete(p))
    return { ok: true, removed: paths?.length || 0 }
  })

  registerHandle('home:delete-files', async (_event: unknown, args: unknown) => {
    const { paths } = (args || {}) as { paths: string[] }
    paths?.forEach(p => {
      if (existsSync(p)) unlinkSync(p)
    })
    return { ok: true, deleted: paths?.length || 0 }
  })

  registerHandle('home:duplicate-file', async (_event: unknown, args: unknown) => {
    const { path } = (args || {}) as { path: string }
    if (!existsSync(path)) return { ok: false, error: 'File not found' }
    const dir = dirname(path)
    const ext = extname(path)
    const base = basename(path, ext)
    const newPath = join(dir, `${base}-copy${ext}`)
    writeFileSync(newPath, readFileSync(path))
    return { ok: true, path: newPath }
  })

  registerHandle('home:rename-file', async (_event: unknown, args: unknown) => {
    const { path, newName } = (args || {}) as { path: string; newName: string }
    if (!existsSync(path)) return { ok: false, error: 'File not found' }
    const dir = dirname(path)
    const newPath = join(dir, newName)
    renameSync(path, newPath)
    return { ok: true, path: newPath }
  })

  registerHandle('home:reveal-path', (_event: unknown, args: unknown) => {
    const { path } = (args || {}) as { path: string }
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
    return { id, path: join(DATA_DIR, `${id}.pdf`) }
  })

  registerHandle('home:account-status', () => ({
    loggedIn: true,
    email: 'web-user@genoffice.ai',
    plan: 'pro',
  }))

  registerHandle('home:account-login', async (_event: unknown, _args: unknown) => ({
    ok: true,
    email: 'web-user@genoffice.ai',
  }))

  registerHandle('home:account-login-open-url', () => ({
    url: 'https://account.genspark.ai/login',
  }))

  registerHandle('home:account-logout', () => ({ ok: true }))

  registerHandle('home:github-stars', () => ({ stars: 128 }))

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

  registerHandle('home:onboarding-seen', () => true)
  registerHandle('home:set-onboarding-seen', (_event: unknown, seen: unknown) => ({
    ok: true,
    seen,
  }))

  registerHandle('home:star-prompt-should-show', () => ({ shouldShow: false }))
  registerHandle('home:star-prompt-action', (_event: unknown, args: unknown) => {
    const { action } = (args || {}) as { action: string }
    return { ok: true, action }
  })

  registerHandle('home:cloud-projects', () => ({ projects: [] }))
  registerHandle('home:cloud-projects-cached', () => ({ projects: [], cached: true }))

  registerHandle('home:open-cloud-project', (_event: unknown, args: unknown) => {
    const { projectUrl } = (args || {}) as { projectUrl: string }
    return { ok: true, url: projectUrl }
  })

  registerHandle('home:open-gen-team', () => ({ ok: true }))
  registerHandle('home:open-credit-usage', () => ({ ok: true }))
  registerHandle('home:open-github-repo', () => ({ ok: true }))

  registerHandle('home:stat-paths', (_event: unknown, args: unknown) => {
    const { paths } = (args || {}) as { paths: string[] }
    return (paths || []).map(p => ({
      path: p,
      exists: existsSync(p),
      size: existsSync(p) ? statSync(p).size : 0,
    }))
  })

  registerHandle('home:browse', () => ({ canceled: false, filePaths: [] }))
}

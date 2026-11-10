import { createRequire } from 'node:module'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'
import { packagedResourcesDir, repoRoot } from './resources'
import { CliError, EXIT } from './result'

let installed = false

/**
 * The docs and markdown editors are ProseMirror documents; they need a DOM
 * to exist, not to be seen. jsdom provides one. It is installed once, before
 * any editor module is imported (ProseMirror sniffs the environment at load).
 */
export async function ensureDom(): Promise<void> {
  if (installed) return
  const { JSDOM, VirtualConsole } = await loadJsdom()
  // a listener-less console swallows jsdom's "not implemented" notes (canvas
  // getContext from the editors' layout probes) that would otherwise hit stderr
  const dom = new JSDOM('<!doctype html><html><body></body></html>', {
    pretendToBeVisual: true,
    virtualConsole: new VirtualConsole(),
  })
  const w = dom.window as unknown as Record<string, unknown>
  const g = globalThis as unknown as Record<string, unknown>
  for (const key of ['window', 'document', 'navigator']) {
    Object.defineProperty(g, key, {
      value: key === 'window' ? w : w[key],
      configurable: true,
      writable: true,
    })
  }
  for (const key of Object.getOwnPropertyNames(w)) {
    if (key in g) continue
    try {
      g[key] = w[key]
    } catch {}
  }
  g.ResizeObserver ??= class {
    observe() {}
    unobserve() {}
    disconnect() {}
  }
  const noMedia = () => ({
    matches: false,
    addEventListener() {},
    removeEventListener() {},
    addListener() {},
    removeListener() {},
  })
  g.matchMedia ??= noMedia
  w.matchMedia ??= noMedia
  g.requestAnimationFrame ??= (cb: (t: number) => void) => setTimeout(() => cb(Date.now()), 0)
  g.cancelAnimationFrame ??= (id: ReturnType<typeof setTimeout>) => clearTimeout(id)
  const range = (w.Range as { prototype: Record<string, unknown> }).prototype
  range.getClientRects ??= () => []
  range.getBoundingClientRect ??= () => ({
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    width: 0,
    height: 0,
  })
  const doc = w.document as Record<string, unknown>
  doc.getSelection ??= () => null
  installed = true
}

async function loadJsdom(): Promise<typeof import('jsdom')> {
  const candidates: string[] = []
  const packaged = packagedResourcesDir()
  if (packaged) candidates.push(join(packaged, 'cli', 'node_modules', 'jsdom', 'package.json'))
  const root = repoRoot()
  if (root) candidates.push(join(root, 'package.json'))
  for (const from of candidates) {
    try {
      const entry = createRequire(from).resolve('jsdom')
      return (await import(pathToFileURL(entry).href)) as typeof import('jsdom')
    } catch {}
  }
  throw new CliError(EXIT.conversion, 'jsdom not found (needed for Word and Markdown documents)', {
    hint: 'run from a checkout with node_modules installed',
  })
}

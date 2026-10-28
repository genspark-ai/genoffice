import { decodeFrame, encodeFrame } from '../serialize'

/**
 * The one WebSocket per browser tab. Connects to the URL the page was served
 * from (so the server can tell which app the tab is), moves IPC frames in both
 * directions, hosts the dialog overlay that stands in for native dialogs, and
 * keeps stray OS-file drops from navigating the tab away.
 */

interface PendingInvoke {
  resolve: (value: unknown) => void
  reject: (err: Error) => void
}

interface DialogRequest {
  id: number
  options: {
    kind: 'save' | 'open' | 'message'
    title?: string
    defaultPath?: string
    buttonLabel?: string
    filters?: { name: string; extensions: string[] }[]
    message?: string
    detail?: string
    buttons?: string[]
  }
}

interface DirEntry {
  name: string
  dir: boolean
}

const pendingInvokes = new Map<number, PendingInvoke>()
const pendingDialogList = new Map<number, (entries: DirEntry[]) => void>()
const pendingDialogValidate = new Map<number, (path: string | null) => void>()
const pendingDialogStarts = new Map<number, (dirs: string[]) => void>()
let dialogResolver: ((result: unknown) => void) | null = null
let nextId = 1
let ws: WebSocket | null = null
let bootedAt = Date.now()

function wsUrl(): string {
  const proto = location.protocol === 'https:' ? 'wss:' : 'ws:'
  return `${proto}//${location.host}${location.pathname}${location.search}`
}

function call(frame: Record<string, unknown>): void {
  if (ws && ws.readyState === WebSocket.OPEN) ws.send(encodeFrame(frame))
}

function invoke(channel: string, args: unknown[]): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const id = nextId++
    pendingInvokes.set(id, { resolve, reject })
    call({ t: 'invoke', id, channel, args })
  })
}

function send(channel: string, args: unknown[]): void {
  call({ t: 'send', channel, args })
}

function failAllPending(err: Error): void {
  for (const p of pendingInvokes.values()) p.reject(err)
  pendingInvokes.clear()
  dialogResolver?.(null)
  dialogResolver = null
}

function onMessage(text: string): void {
  let frame: Record<string, unknown>
  try {
    frame = decodeFrame(text) as Record<string, unknown>
  } catch {
    return
  }
  const t = frame.t
  if (t === 'reply') {
    const p = pendingInvokes.get(Number(frame.id))
    if (!p) return
    pendingInvokes.delete(Number(frame.id))
    if (frame.ok === true) p.resolve(frame.value)
    else p.reject(new Error(String(frame.error)))
  } else if (t === 'event') {
    const channel = String(frame.channel)
    const args = Array.isArray(frame.args) ? frame.args : []
    if (channel === 'web:dialog') {
      const req = args[0] as DialogRequest | undefined
      if (req) showDialogOverlay(req)
      return
    }
    const dispatch = (window as unknown as Record<string, unknown>).__genofficeWebDispatch as
      ((channel: string, args: unknown[]) => void) | undefined
    dispatch?.(channel, args)
  } else if (t === 'dialog-list-reply') {
    pendingDialogList.get(Number(frame.id))?.((frame.entries as DirEntry[]) ?? [])
    pendingDialogList.delete(Number(frame.id))
  } else if (t === 'dialog-validate-reply') {
    pendingDialogValidate.get(Number(frame.id))?.((frame.path as string | null) ?? null)
    pendingDialogValidate.delete(Number(frame.id))
  } else if (t === 'dialog-start-dirs-reply') {
    pendingDialogStarts.get(Number(frame.id))?.((frame.dirs as string[]) ?? [])
    pendingDialogStarts.delete(Number(frame.id))
  } else if (t === 'open-url' && typeof frame.url === 'string') {
    window.open(frame.url, '_blank')
  }
}

function connect(): void {
  bootedAt = Date.now()
  const socket = new WebSocket(wsUrl())
  ws = socket
  socket.onmessage = (ev) => onMessage(String(ev.data))
  socket.onclose = () => {
    ws = null
    failAllPending(new Error('web bridge closed'))
    // a fresh load re-seeds the tab's pending state cleanly; a boot that never
    // connected (server gone) would loop, so only reload a page that worked
    if (Date.now() - bootedAt > 5000) {
      location.reload()
    } else {
      showLostOverlay()
    }
  }
}

// ---- the dialog overlay (native dialogs live on the server machine) ----

function overlayHost(): HTMLDivElement {
  let host = document.getElementById('genoffice-web-dialog') as HTMLDivElement | null
  if (!host) {
    host = document.createElement('div')
    host.id = 'genoffice-web-dialog'
    host.style.cssText =
      'position:fixed;inset:0;background:rgba(15,18,22,.45);z-index:2147483646;display:flex;align-items:center;justify-content:center;font-family:-apple-system,system-ui,sans-serif'
    document.body.appendChild(host)
  }
  return host
}

function toast(text: string): void {
  const el = document.createElement('div')
  el.textContent = text
  el.style.cssText =
    'position:fixed;left:50%;bottom:28px;transform:translateX(-50%);background:#1f2328;color:#fff;padding:8px 14px;border-radius:8px;z-index:2147483647;font:13px -apple-system,system-ui,sans-serif'
  document.body.appendChild(el)
  setTimeout(() => el.remove(), 3200)
}

function requestDialogList(dir: string, exts: string[]): Promise<DirEntry[]> {
  return new Promise((resolve) => {
    const id = nextId++
    pendingDialogList.set(id, resolve)
    call({ t: 'dialog-list', id, dir, exts })
    setTimeout(() => {
      if (pendingDialogList.delete(id)) resolve([])
    }, 8000)
  })
}

function requestDialogStarts(defaultPath?: string): Promise<string[]> {
  return new Promise((resolve) => {
    const id = nextId++
    pendingDialogStarts.set(id, resolve)
    call({ t: 'dialog-start-dirs', id, defaultPath })
    setTimeout(() => {
      if (pendingDialogStarts.delete(id)) resolve([])
    }, 8000)
  })
}

function validateDialogPath(path: string, kind: 'save' | 'open'): Promise<string | null> {
  return new Promise((resolve) => {
    const id = nextId++
    pendingDialogValidate.set(id, resolve)
    call({ t: 'dialog-validate', id, path, kind })
    setTimeout(() => {
      if (pendingDialogValidate.delete(id)) resolve(null)
    }, 8000)
  })
}

function showDialogOverlay(req: DialogRequest): void {
  if (req.options.kind === 'message') {
    showMessageOverlay(req)
    return
  }
  const kind = req.options.kind
  const exts = req.options.filters?.[0]?.extensions ?? []
  const host = overlayHost()
  const panel = document.createElement('div')
  panel.style.cssText =
    'background:#fff;border-radius:12px;min-width:460px;max-width:620px;max-height:70vh;display:flex;flex-direction:column;box-shadow:0 18px 50px rgba(0,0,0,.3);overflow:hidden'
  panel.innerHTML = `
    <div style="padding:12px 16px;border-bottom:1px solid #e4e7ec;font-weight:600;font-size:14px" data-role="title"></div>
    <div style="padding:8px 16px 0;font-size:12px;color:#59636e" data-role="crumbs"></div>
    <div data-role="list" style="flex:1;overflow:auto;margin:8px 16px;border:1px solid #e4e7ec;border-radius:8px;min-height:220px"></div>
    <div style="padding:0 16px 12px;display:flex;gap:8px;align-items:center" data-role="namer">
      <input data-role="name" style="flex:1;padding:7px 10px;border:1px solid #d1d9e0;border-radius:8px;font-size:13px" />
    </div>
    <div style="padding:12px 16px;border-top:1px solid #e4e7ec;display:flex;justify-content:flex-end;gap:8px">
      <button data-role="cancel" style="padding:7px 14px;border:1px solid #d1d9e0;background:#fff;border-radius:8px;cursor:pointer;font-size:13px">Cancel</button>
      <button data-role="ok" style="padding:7px 14px;border:1px solid #0969da;background:#0969da;color:#fff;border-radius:8px;cursor:pointer;font-size:13px"></button>
    </div>`
  host.appendChild(panel)

  const q = (role: string): HTMLElement =>
    panel.querySelector(`[data-role="${role}"]`) as HTMLElement
  let cwd = ''
  let entries: DirEntry[] = []
  const nameInput = q('name') as HTMLInputElement

  q('title').textContent = req.options.title ?? (kind === 'save' ? 'Save' : 'Open')
  q('ok').textContent = req.options.buttonLabel ?? (kind === 'save' ? 'Save' : 'Open')
  if (kind === 'open') q('namer').style.display = 'none'
  else if (req.options.defaultPath) nameInput.value = req.options.defaultPath.split('/').pop() ?? ''

  async function showDir(dir: string): Promise<void> {
    cwd = dir
    entries = await requestDialogList(dir, exts)
    q('crumbs').textContent = dir
    const list = q('list')
    list.innerHTML = ''
    const parent = document.createElement('div')
    parent.textContent = '…'
    parent.style.cssText = 'padding:4px 10px;cursor:pointer;font-size:13px'
    parent.onclick = () => void showDir(dir.replace(/\/[^/]*\/?$/, '') || '/')
    list.appendChild(parent)
    for (const entry of entries) {
      const row = document.createElement('div')
      row.textContent = (entry.dir ? '📁 ' : '　 ') + entry.name
      row.style.cssText = 'padding:4px 10px;cursor:pointer;font-size:13px;white-space:nowrap'
      row.onmouseenter = () => (row.style.background = '#f3f4f6')
      row.onmouseleave = () => (row.style.background = '')
      row.onclick = () => {
        if (entry.dir) void showDir(dir.replace(/\/$/, '') + '/' + entry.name)
        else if (kind === 'open') nameInput.value = entry.name
        else nameInput.value = entry.name
      }
      row.ondblclick = () => {
        if (!entry.dir && kind === 'open') void confirmSelection()
      }
      list.appendChild(row)
    }
  }

  function respond(path: string | null): void {
    host.remove()
    dialogResolver?.(
      path === null ? { canceled: true } : { canceled: false, filePath: path, filePaths: [path] },
    )
    dialogResolver = null
  }

  async function confirmSelection(): Promise<void> {
    const name = nameInput.value.trim()
    if (!name) return
    const path = cwd.replace(/\/$/, '') + '/' + name
    const valid = await validateDialogPath(path, kind)
    if (!valid) {
      toast(kind === 'open' ? 'That file does not exist' : 'That directory does not exist')
      return
    }
    respond(valid)
  }

  q('cancel').onclick = () => respond(null)
  q('ok').onclick = () => void confirmSelection()
  nameInput.addEventListener('keydown', (ev) => {
    if (ev.key === 'Enter') void confirmSelection()
  })

  dialogResolver = () => respond(null)
  void requestDialogStarts(req.options.defaultPath).then((dirs) => {
    void showDir(dirs[0] ?? '/')
  })
}

function showMessageOverlay(req: DialogRequest): void {
  const host = overlayHost()
  const panel = document.createElement('div')
  panel.style.cssText =
    'background:#fff;border-radius:12px;min-width:380px;max-width:520px;padding:20px;box-shadow:0 18px 50px rgba(0,0,0,.3);font-size:14px'
  const buttons = req.options.buttons?.length ? req.options.buttons : ['OK']
  panel.innerHTML = `
    <div style="font-weight:600;margin-bottom:6px">${escapeHtml(req.options.message ?? '')}</div>
    <div style="color:#59636e;margin-bottom:16px;white-space:pre-wrap">${escapeHtml(req.options.detail ?? '')}</div>
    <div style="display:flex;justify-content:flex-end;gap:8px" data-role="buttons"></div>`
  host.appendChild(panel)
  const row = panel.querySelector('[data-role="buttons"]') as HTMLElement
  buttons.forEach((label, index) => {
    const btn = document.createElement('button')
    btn.textContent = label
    btn.style.cssText = `padding:7px 14px;border-radius:8px;cursor:pointer;font-size:13px;border:1px solid #d1d9e0;background:#fff;${index === 0 ? '' : ''}`
    if (index === buttons.length - 1) {
      btn.style.background = '#0969da'
      btn.style.borderColor = '#0969da'
      btn.style.color = '#fff'
    }
    btn.onclick = () => {
      host.remove()
      const resolve = dialogResolver
      dialogResolver = null
      resolve?.({ canceled: false, response: index })
    }
    row.appendChild(btn)
  })
  dialogResolver = () => {
    const resolve = dialogResolver
    dialogResolver = null
    host.remove()
    resolve?.({ canceled: false, response: 0 })
  }
}

function escapeHtml(text: string): string {
  return text.replace(
    /[&<>"]/g,
    (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c] ?? c,
  )
}

function showLostOverlay(): void {
  const host = overlayHost()
  host.style.background = 'rgba(15,18,22,.85)'
  host.innerHTML =
    '<div style="background:#fff;border-radius:12px;padding:24px 28px;font:14px -apple-system,system-ui,sans-serif;max-width:420px">' +
    '<b>GenOffice service disconnected.</b><br><br>' +
    'The local service stopped or restarted. Reload this tab once it is reachable again.</div>'
}

// ---- OS-file drops: a browser tab cannot read dropped local paths, and the
// default navigation would discard the open document ----

function installDropGuard(): void {
  window.addEventListener(
    'dragover',
    (ev) => {
      const e = ev as DragEvent
      if (!e.defaultPrevented && e.dataTransfer?.types.includes('Files')) e.preventDefault()
    },
    false,
  )
  window.addEventListener(
    'drop',
    (ev) => {
      const e = ev as DragEvent
      if (!e.defaultPrevented && e.dataTransfer?.types.includes('Files')) {
        e.preventDefault()
        toast('Opening local files by drag-and-drop needs the desktop app here')
      }
    },
    false,
  )
}

connect()
installDropGuard()

;(window as unknown as Record<string, unknown>).__genofficeWebClient = { invoke, send }

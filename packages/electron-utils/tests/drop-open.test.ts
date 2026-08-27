import { beforeEach, describe, expect, it, vi } from 'vitest'

import {
  DROP_OPEN_CHANNEL,
  droppableFilePaths,
  installDropOpenBridge,
  partitionDropPayload,
} from '../src/drop-open'

// ---- electron mock (module under test imports ipcRenderer/webUtils eagerly) ----

const electronMocks = vi.hoisted(() => ({
  send: vi.fn(),
  getPathForFile: vi.fn(),
}))

vi.mock('electron', () => ({
  ipcRenderer: { send: electronMocks.send },
  webUtils: { getPathForFile: electronMocks.getPathForFile },
}))

// ---- tiny DOM fakes ----

interface RecordedEvent {
  type: string
  handler: (ev: unknown) => void
}

/** minimal stand-in for `window` recording every listener registration */
function fakeWindow(): {
  recorded: RecordedEvent[]
  fire(type: string, ev: Record<string, unknown>): void
} {
  const recorded: RecordedEvent[] = []
  const target = {
    recorded,
    addEventListener(type: string, handler: (ev: unknown) => void) {
      recorded.push({ type, handler })
    },
    removeEventListener() {},
    fire(type: string, ev: Record<string, unknown>) {
      for (const r of recorded.filter((r) => r.type === type)) r.handler(ev)
    },
  }
  return target
}

type GlobalWithWindow = { window?: unknown }

/** the idempotence flag lives on globalThis; clear it between installs */
function resetInstallFlag(): void {
  delete (globalThis as Record<symbol | string, unknown>)[
    Symbol.for('genoffice.drop-open-installed')
  ]
}

/** DragEvent-shaped literal; only the fields the bridge touches are provided */
function fileDrag(paths: Array<string | ''>, prevented = false): Record<string, unknown> {
  return {
    defaultPrevented: prevented,
    preventDefault: vi.fn(),
    stopPropagation: vi.fn(),
    dataTransfer: {
      types: ['Files'],
      files: paths.map((p) => ({ name: p.split('/').pop() ?? '?' })),
    },
  }
}

beforeEach(() => {
  electronMocks.send.mockClear()
  electronMocks.getPathForFile.mockReset()
  resetInstallFlag()
})

describe('droppableFilePaths', () => {
  const resolveByName =
    (map: Record<string, string>) =>
    (file: File): string =>
      map[(file as { name: string }).name] ?? ''

  it('returns null when there is no dataTransfer or no Files kind', () => {
    expect(droppableFilePaths({ dataTransfer: null }, () => '')).toBeNull()
    expect(
      droppableFilePaths(
        { dataTransfer: { types: ['text/plain'], files: [] } as unknown as DataTransfer },
        () => '',
      ),
    ).toBeNull()
  })

  it('maps every dropped file through the resolver in order', () => {
    const result = droppableFilePaths(fileDrag(['a.docx', 'b.pdf']) as never, () => '/tmp')
    expect(result).toEqual(['/tmp', '/tmp'])
  })

  it('drops empty resolutions from virtual entries instead of failing', () => {
    const result = droppableFilePaths(
      fileDrag(['shot.png', 'doc.docx']) as never,
      resolveByName({ 'shot.png': '', 'doc.docx': '/tmp/doc.docx' }),
    )
    expect(result).toEqual(['/tmp/doc.docx'])
  })
})

describe('partitionDropPayload', () => {
  it('rejects non-array payloads outright', () => {
    expect(partitionDropPayload(null)).toEqual({ supported: [], unsupportedExts: [] })
    expect(partitionDropPayload('a.docx')).toEqual({ supported: [], unsupportedExts: [] })
    expect(partitionDropPayload({ 0: 'a.docx' })).toEqual({ supported: [], unsupportedExts: [] })
  })

  it('ignores non-string entries, whitespace-only strings, and duplicates', () => {
    const result = partitionDropPayload([' a.docx ', 'a.docx', 42, null, '', '  ', 'b.xlsx'])
    expect(result.supported).toEqual(['a.docx', 'b.xlsx'])
    expect(result.unsupportedExts).toEqual([])
  })

  it('classifies openable extensions case-insensitively', () => {
    const result = partitionDropPayload(['REPORT.DOCX', 'data.CSV', 'notes.MarkDown'])
    expect(result.supported).toEqual(['REPORT.DOCX', 'data.CSV', 'notes.MarkDown'])
  })

  it('collects known-unsupported extensions uniquely, first-seen order', () => {
    const result = partitionDropPayload(['old.doc', 'x.pages', 'y.rtf', 'z.doc'])
    expect(result.supported).toEqual([])
    expect(result.unsupportedExts).toEqual(['doc', 'pages', 'rtf'])
  })

  it('caps the number of openable files at 20', () => {
    const many = Array.from({ length: 30 }, (_, i) => `/tmp/f${i}.docx`)
    expect(partitionDropPayload(many).supported.length).toBe(20)
  })

  it('keeps the original order so the last file wins activation', () => {
    const result = partitionDropPayload(['/a.docx', '/b/c.csv', '/d.md'])
    expect(result.supported).toEqual(['/a.docx', '/b/c.csv', '/d.md'])
  })
})

describe('installDropOpenBridge', () => {
  function install(): ReturnType<typeof fakeWindow> {
    const win = fakeWindow()
    ;(globalThis as GlobalWithWindow).window = win
    installDropOpenBridge()
    return win
  }

  function uninstall(win: ReturnType<typeof fakeWindow>): void {
    delete (globalThis as GlobalWithWindow).window
    void win
    resetInstallFlag()
  }

  it('registers exactly one dragover and one drop listener', () => {
    const win = install()
    try {
      expect(win.recorded.map((r) => r.type)).toEqual(['dragover', 'drop'])
    } finally {
      uninstall(win)
    }
  })

  it('is idempotent within the same process', () => {
    const win = install()
    try {
      installDropOpenBridge()
      expect(win.recorded.length).toBe(2)
    } finally {
      uninstall(win)
    }
  })

  it('cancels dragover for file drags so Chromium fires drop', () => {
    const win = install()
    try {
      const ev = fileDrag(['a.docx'])
      win.fire('dragover', ev)
      expect(ev.preventDefault).toHaveBeenCalledOnce()
    } finally {
      uninstall(win)
    }
  })

  it('leaves dragover alone for non-file drags and page-handled events', () => {
    const win = install()
    try {
      const text = {
        defaultPrevented: false,
        preventDefault: vi.fn(),
        dataTransfer: { types: ['text/plain'], files: [] },
      }
      win.fire('dragover', text)
      expect(text.preventDefault).not.toHaveBeenCalled()

      const claimed = fileDrag(['a.docx'], true)
      win.fire('dragover', claimed)
      expect(claimed.preventDefault).not.toHaveBeenCalled()
    } finally {
      uninstall(win)
    }
  })

  it('sends resolved document paths over the channel and cancels the native drop', async () => {
    electronMocks.getPathForFile.mockImplementation((f: { name: string }) => `/tmp/${f.name}`)
    const win = install()
    try {
      const ev = fileDrag(['report.docx', 'plan.md', 'photo.png'])
      win.fire('drop', ev)
      await vi.waitFor(() =>
        expect(electronMocks.send).toHaveBeenCalledWith(DROP_OPEN_CHANNEL, [
          '/tmp/report.docx',
          '/tmp/plan.md',
        ]),
      )
      expect(ev.preventDefault).toHaveBeenCalled()
    } finally {
      uninstall(win)
    }
  })

  it('stays silent when the page already handled the drop', async () => {
    electronMocks.getPathForFile.mockImplementation((f: { name: string }) => `/tmp/${f.name}`)
    const win = install()
    try {
      const ev = fileDrag(['report.docx'], true)
      win.fire('drop', ev)
      await Promise.resolve()
      expect(electronMocks.send).not.toHaveBeenCalled()
    } finally {
      uninstall(win)
    }
  })

  it('swallows unresolvable/non-document drops without sending anything', async () => {
    electronMocks.getPathForFile.mockReturnValue('')
    const win = install()
    try {
      const virtual = fileDrag(['remote.html'])
      win.fire('drop', virtual)
      const images = fileDrag(['a.png', 'b.jpg'])
      win.fire('drop', images)
      await Promise.resolve()
      expect(electronMocks.send).not.toHaveBeenCalled()
      expect(images.preventDefault).toHaveBeenCalled()
    } finally {
      uninstall(win)
    }
  })

  it('forwards known-unsupported docs too so the shell can explain why', async () => {
    electronMocks.getPathForFile.mockImplementation((f: { name: string }) => `/tmp/${f.name}`)
    const win = install()
    try {
      const ev = fileDrag(['legacy.doc'])
      win.fire('drop', ev)
      await vi.waitFor(() =>
        expect(electronMocks.send).toHaveBeenCalledWith(DROP_OPEN_CHANNEL, ['/tmp/legacy.doc']),
      )
    } finally {
      uninstall(win)
    }
  })
})

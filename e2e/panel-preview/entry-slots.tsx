import '@genoffice/ui/tokens.css'
import '../../apps/shell/src/renderer/src/home.css'
import { createRoot } from 'react-dom/client'
import { LocaleProvider } from '../../apps/shell/src/renderer/src/locale'
import { SlotHost, SlotPanel } from '../../apps/shell/src/renderer/src/slots/SlotView'
import { OverlayLayer } from '../../apps/shell/src/renderer/src/slots/OverlayLayer'
import { refreshFeatures } from '../../apps/shell/src/renderer/src/slots/features-state'

/**
 * Slot gallery: what a feature draws, in the three places open today. Same
 * harness contract as the panel preview — the real components and the real
 * stylesheet against a stub bridge, so the picture is the product's rendering.
 *
 * The file sidebar is the interesting one: the slot itself carries no files, it
 * names the recent-files source and the host fills it.
 */

/** One appearance feature, so the harness exercises the real apply path. */
const appearanceFeature = {
  id: 'night-shift',
  name: 'Night shift',
  description: 'softer chrome after dark',
  kind: 'appearance',
  appearance: { background: 'surface-subtle', accent: 'violet' },
  createdAt: 'x',
  updatedAt: '2099-01-01T00:00:00.000Z',
}

/** A browser panel, the native view slot: the host frame is real, the page inside
 * is mounted by the shell (in this harness there is no main process, so the body
 * shows the state before the view attaches). */
const browserFeature = {
  id: 'headlines',
  name: 'Headlines',
  description: 'A news page, next to the files.',
  kind: 'slot',
  slot: {
    at: 'home.panel',
    element: {
      type: 'view',
      view: 'browser',
      url: 'https://news.ycombinator.com',
      title: 'News',
      minHeight: 320,
    },
  },
  createdAt: 'x',
  updatedAt: 'x',
}

/** The editor place: the same declarations, drawn in an editor's own dock. */
const editorRailFeature = {
  id: 'doc-tools',
  name: 'Doc tools',
  description: 'A rail of features inside the document editor.',
  kind: 'slot',
  slot: {
    at: 'editor.side',
    element: {
      type: 'group',
      children: [
        { type: 'text', value: 'Doc tools', size: 'caption', tone: 'muted' },
        {
          type: 'button',
          label: 'Tidy headings',
          tone: 'primary',
          action: { kind: 'ask-assistant', prompt: 'Tidy the headings in this document' },
        },
        {
          type: 'file-list',
          source: 'recent-files',
          limit: 3,
          emptyText: 'No recent files',
          action: { kind: 'open', path: '$path' },
        },
      ],
    },
  },
  createdAt: 'x',
  updatedAt: 'x',
}

const plain = new URLSearchParams(location.search).get('plain') === '1'

const features = [
  {
    id: 'file-sidebar',
    name: 'Files',
    description: 'Jump to a recent file from the home screen.',
    kind: 'slot',
    slot: {
      at: 'home.panel',
      element: {
        type: 'group',
        children: [
          { type: 'text', value: 'Recent files', size: 'caption', tone: 'muted' },
          {
            type: 'file-list',
            source: 'recent-files',
            limit: 6,
            emptyText: 'No recent files',
            action: { kind: 'open', path: '$path' },
          },
          { type: 'divider' },
          { type: 'text', value: 'Starred', size: 'caption', tone: 'muted' },
          {
            type: 'file-list',
            source: 'starred-files',
            limit: 3,
            action: { kind: 'ask-assistant', prompt: 'Summarise $name from $folder' },
          },
        ],
      },
    },
    createdAt: 'x',
    updatedAt: 'x',
  },
  {
    id: 'files-entry',
    name: 'Files',
    description: 'the same feature, opened from the sidebar',
    kind: 'slot',
    slot: {
      at: 'home.sidebar',
      element: {
        type: 'group',
        children: [
          {
            type: 'button',
            label: 'Files',
            action: { kind: 'panel', id: 'file-sidebar' },
          },
        ],
      },
    },
    createdAt: 'x',
    updatedAt: 'x',
  },
  {
    id: 'house-style',
    name: 'House style',
    description: 'Apply the team style to this report.',
    kind: 'slot',
    slot: {
      at: 'home.card',
      element: {
        type: 'group',
        children: [
          { type: 'text', value: 'House style', size: 'caption', tone: 'muted' },
          { type: 'text', value: 'Heading levels, table borders and the footer.' },
          {
            type: 'button',
            label: 'Apply to this report',
            tone: 'primary',
            action: { kind: 'ask-assistant', prompt: 'Apply the house style to this report' },
          },
        ],
      },
    },
    createdAt: 'x',
    updatedAt: 'x',
  },
]

const page = (entries: Array<{ path: string; name: string }>) => ({
  entries: entries.map((entry, index) => ({
    ...entry,
    ext: entry.name.split('.').pop() ?? 'docx',
    mtimeMs: Date.now() - index * 3_600_000,
    sizeBytes: 20_000 + index * 1_000,
    starred: false,
  })),
  total: entries.length,
  totalAll: entries.length,
})

const recent = [
  { path: '/Users/me/Documents/GenOffice/weekly-report.docx', name: 'weekly-report.docx' },
  { path: '/Users/me/Documents/GenOffice/budget-2026.xlsx', name: 'budget-2026.xlsx' },
  { path: '/Users/me/Documents/GenOffice/kickoff.pptx', name: 'kickoff.pptx' },
]
const starred = [
  { path: '/Users/me/Documents/GenOffice/contract-lease.docx', name: 'contract-lease.docx' },
  { path: '/Users/me/Documents/GenOffice/brand-notes.md', name: 'brand-notes.md' },
]

/** An overlay: what a mod can point at without touching anything. */
const overlayFeature = {
  id: 'review-guide',
  name: 'Review guide',
  description: 'Points at the parts of a report that need a reader.',
  kind: 'slot',
  slot: {
    at: 'home.overlay',
    element: {
      type: 'overlay',
      shapes: [
        {
          kind: 'frame',
          x: 0.03,
          y: 0.06,
          width: 0.32,
          height: 0.3,
          tone: 'accent',
          label: 'Sidebar slot',
        },
        { kind: 'line', x1: 0.42, y1: 0, x2: 0.42, y2: 1, tone: 'muted', dashed: true },
        {
          kind: 'frame',
          x: 0.5,
          y: 0.08,
          width: 0.47,
          height: 0.24,
          tone: 'success',
          dashed: true,
        },
        { kind: 'dot', x: 0.5, y: 0.62, tone: 'danger', label: 'A page the mod added' },
        {
          kind: 'label',
          x: 0.52,
          y: 0.95,
          value: 'drawn by a feature, over the screen only',
          tone: 'muted',
        },
      ],
    },
  },
  createdAt: 'x',
  updatedAt: 'x',
}

;(window as unknown as { aiOffice: unknown }).aiOffice = {
  featuresList: async () =>
    plain
      ? [...features, browserFeature, overlayFeature, editorRailFeature]
      : [...features, browserFeature, overlayFeature, editorRailFeature, appearanceFeature],
  recents: async () => page(recent),
  starred: async () => page(starred),
  openPath: async () => {},
  getAiSettings: async () => ({}),
  onAiStream: () => () => {},
  aiStream: async () => {},
  aiStreamCancel: async () => {},
  setLanguage: async () => {},
}

const limited = new URLSearchParams(location.search).get('host') === 'limited'
/** An editor without a file-opener: an open action must draw disabled, not dead. */
const limitedActions = { ask: () => {}, panel: () => {} }

const actions = {
  open: () => {},
  ask: () => {},
  panel: () => {},
  // the harness plays the host bridge the shell provides for a file-list
  listSource: async (source: string) => (source === 'starred-files' ? starred : recent),
  mountView: () => {},
  unmountView: () => {},
}

function Gallery(): React.JSX.Element {
  return (
    <div className="gallery content--overlay">
      <div className="slot-overlay-host">
        <OverlayLayer />
      </div>
      <div className="gallery-col gallery-sidebar">
        <SlotHost place="home.sidebar" actions={actions} />
      </div>
      <div className="gallery-col gallery-cards">
        <SlotHost place="home.card" actions={limited ? limitedActions : actions} />
        {/* the editor place, drawn by the same renderer the editors bind */}
        <SlotHost place="editor.side" actions={limited ? limitedActions : actions} />
      </div>
      <div className="gallery-col gallery-panel">
        <SlotPanel featureId="file-sidebar" actions={actions} onClose={() => {}} />
        <SlotPanel featureId="headlines" actions={actions} onClose={() => {}} />
      </div>
    </div>
  )
}

void refreshFeatures().then(() => {
  createRoot(document.getElementById('root')!).render(
    <LocaleProvider initial="en">
      <Gallery />
    </LocaleProvider>,
  )
})

import '@genoffice/ui/tokens.css'
import '../../apps/shell/src/renderer/src/home.css'
import { createRoot } from 'react-dom/client'
import { LocaleProvider } from '../../apps/shell/src/renderer/src/locale'
import { ModsAssistant } from '../../apps/shell/src/renderer/src/ModsAssistant'

/**
 * Design preview harness: renders the real ModsAssistant with a stub bridge so
 * the panel can be looked at without booting the Electron shell. Not part of the
 * product build; it exists so the layout can be reviewed as a picture.
 */

const withFeatures = [
  {
    id: 'weekly-report-formatting',
    name: 'Weekly report formatting',
    description: 'Apply the house style to a weekly report: heading levels, table borders, footer.',
    kind: 'slot',
    createdAt: '2026-10-11T00:00:00.000Z',
    updatedAt: '2026-10-11T00:00:00.000Z',
  },
  {
    id: 'file-switcher',
    name: 'File switcher',
    description: 'Jump to any recent file from the keyboard.',
    kind: '',
    createdAt: '2026-10-10T00:00:00.000Z',
    updatedAt: '2026-10-10T00:00:00.000Z',
  },
  {
    id: 'tidy-headings',
    name: 'Tidy headings',
    description: 'Unfold every heading that was wrapped across lines.',
    kind: 'script',
    script: { source: 'var blocks = $.docs.blocks()\n$.docs.propose([])' },
    createdAt: '2026-10-09T00:00:00.000Z',
    updatedAt: '2026-10-09T00:00:00.000Z',
  },
]

/** What the host would answer: a plan for a dry run, and a write for an apply. */
const outcome = {
  applied: 3,
  plan: [
    '0: setHeadingLevel: heading 2',
    '4: findReplace: replaced 2 runs',
    '7: insertField: page number',
  ],
  failures: [],
  log: ['looked at 24 blocks'],
  ms: 6,
}

const empty = new URLSearchParams(location.search).get('state') === 'empty'

;(window as unknown as { aiOffice: unknown }).aiOffice = {
  getAiSettings: async () => ({}),
  featuresList: async () => (empty ? [] : withFeatures),
  featuresGet: async () => null,
  featuresCreate: async (draft: unknown) => draft,
  featuresUpdate: async () => null,
  featuresDelete: async () => true,
  featuresReadDocument: async () => ({ name: '', text: '', chars: 0 }),
  featuresRun: async () => outcome,
  // what the host records when a hook observes an event
  featuresPackExport: async () => '/tmp/genoffice-features.json',
  // the import review: what a pack declares, before anything is written
  featuresPackPreview: async () => ({
    path: '/tmp/genoffice-features.json',
    description: {
      version: 1,
      exportedAt: '2026-10-11T00:00:00.000Z',
      count: 3,
      lines: [
        {
          id: 'tidy-headings',
          name: 'Tidy headings',
          kind: 'script',
          declares: ['script, 96 characters'],
        },
        {
          id: 'no-untitled-reports',
          name: 'No untitled reports',
          kind: 'hook',
          declares: ['observes document.will-save'],
        },
        {
          id: 'house-style',
          name: 'House style',
          kind: 'knowledge',
          declares: ['reads folder /Users/me/Documents/house-style'],
        },
      ],
      warnings: [
        'No untitled reports can stop a save',
        'Tidy headings runs code (read-only, sandboxed)',
        'House style reads from /Users/me/Documents/house-style',
      ],
    },
  }),
  featuresPackInstall: async () => ({ installed: ['tidy-headings'], replaced: [], skipped: [] }),
  featuresHookActivity: async () => [
    {
      at: new Date(Date.now() - 600_000).toISOString(),
      hookId: 'note-opened-reports',
      event: 'document.closed',
      path: '/Users/me/Documents/GenOffice/kickoff.pptx',
      log: [],
      error: 'the sandbox refused this run',
    },
    {
      at: new Date(Date.now() - 30_000).toISOString(),
      hookId: 'note-opened-reports',
      event: 'document.opened',
      path: '/Users/me/Documents/GenOffice/weekly-report.docx',
      log: ['opened weekly-report.docx', 'blocks=24'],
    },
    {
      at: new Date(Date.now() - 10_000).toISOString(),
      hookId: 'no-untitled-reports',
      event: 'document.will-save',
      path: '/Users/me/Documents/GenOffice/untitled.docx',
      log: ['checking the title'],
      vetoes: ['the report has no title'],
    },
  ],
  featuresApply: async () => ({ ...outcome, written: '/tmp/weekly-report.docx', bytes: 24_510 }),
  recents: async () => ({
    entries: [
      {
        path: '/Users/me/Documents/GenOffice/weekly-report.docx',
        name: 'weekly-report.docx',
        ext: 'docx',
        mtimeMs: Date.now(),
        sizeBytes: 24_510,
        starred: false,
      },
      {
        path: '/Users/me/Documents/GenOffice/kickoff.pptx',
        name: 'kickoff.pptx',
        ext: 'pptx',
        mtimeMs: Date.now(),
        sizeBytes: 12_000,
        starred: false,
      },
    ],
    total: 2,
    totalAll: 2,
  }),
  onAiStream: () => () => {},
  aiStream: async () => {},
  aiStreamCancel: async () => {},
  setLanguage: async () => {},
}

createRoot(document.getElementById('root')!).render(
  <LocaleProvider initial="en">
    <ModsAssistant />
  </LocaleProvider>,
)

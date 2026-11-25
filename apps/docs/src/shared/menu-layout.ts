import { isPlanned, type CommandId } from './commands'

/**
 * The Docs menu bar as data: which command sits where, in Word for Mac's
 * order. The main process turns it into the native menu; reordering or
 * regrouping happens here, without touching how items are built.
 */

/** items the host builds itself: dynamic lists, shell-dependent or platform-only entries */
export type MenuSlot =
  | 'new-window'
  | 'window-head'
  | 'window-tail'
  | 'open-recent'
  | 'extra-file-items'
  | 'close'
  | 'name-with-ai'
  | 'zoom-presets'
  | 'settings-off-mac'
  | 'dev-tools'

export type MenuRole =
  'cut' | 'copy' | 'paste' | 'pasteAndMatchStyle' | 'selectAll' | 'togglefullscreen'

/** `K` is the menu dictionary's key type; the layout itself only needs strings */
export type MenuNode<K extends string = string> =
  | CommandId
  | '---'
  /** a command shown under a label other than its own, with the pick it stands for */
  | { readonly cmd: CommandId; readonly label: K; readonly payload?: string }
  | { readonly submenu: K; readonly items: readonly MenuNode<K>[] }
  | { readonly role: MenuRole; readonly label: K }
  | { readonly slot: MenuSlot }

export interface MenuTop<K extends string = string> {
  readonly label: K
  /** the menu macOS adds its window list to */
  readonly role?: 'window'
  readonly items: readonly MenuNode<K>[]
}

/** drops the commands that are not built yet, and the separators left with nothing between them */
function withoutPlanned<K extends string>(nodes: readonly MenuNode<K>[]): MenuNode<K>[] {
  const kept: MenuNode<K>[] = []
  for (const node of nodes) {
    if (node === '---') {
      if (kept.length > 0 && kept.at(-1) !== '---') kept.push(node)
    } else if (typeof node === 'string') {
      if (!isPlanned(node)) kept.push(node)
    } else if ('submenu' in node) {
      const items = withoutPlanned(node.items)
      if (items.length > 0) kept.push({ ...node, items })
    } else if (!('cmd' in node) || !isPlanned(node.cmd)) {
      kept.push(node)
    }
  }
  if (kept.at(-1) === '---') kept.pop()
  return kept
}

function shownMenu<K extends string>(tops: readonly MenuTop<K>[]): readonly MenuTop<K>[] {
  return tops.map((top) => ({ ...top, items: withoutPlanned(top.items) }))
}

/** the menu bar with every command Word has in its place, built or not */
const WORD_MENU = [
  {
    label: 'menuFile',
    items: [
      'new',
      'file-new-from-template',
      'open',
      { slot: 'open-recent' },
      '---',
      { slot: 'close' },
      'save',
      'save-as',
      'file-save-as-template',
      'file-move',
      'file-rename',
      '---',
      'file-always-read-only',
      '---',
      'file-reduce-size',
      '---',
      'page-setup',
      // routed through the renderer: it opens the pagination preview first so each
      // printed sheet is exactly one editor page (WYSIWYG), then invokes docs:print
      'print',
      '---',
      'file-properties',
      // not in Word's File menu: kept below its items so their order stays Word's
      '---',
      // also in the native File menu, the only entry on macOS: the ribbon's
      // File tab is Windows/Linux-only
      { slot: 'name-with-ai' },
      'export-pdf',
      'export-html',
      'export-markdown',
      'export-images',
      { slot: 'extra-file-items' },
    ],
  },
  {
    label: 'menuEdit',
    items: [
      'undo',
      'repeat',
      '---',
      'cut',
      'copy',
      '---',
      'paste',
      'edit-paste-special',
      'paste-match-style',
      '---',
      { submenu: 'menuClear', items: ['clear-formatting', 'edit-clear-contents'] },
      'select-all',
      '---',
      {
        submenu: 'menuFindMenu',
        items: [
          'find',
          'edit-find-next',
          'edit-find-previous',
          '---',
          'replace',
          'goto',
          'edit-advanced-find',
        ],
      },
      '---',
      'edit-links',
      '---',
      'edit-select-data',
      '---',
      'edit-toggle-drawing',
    ],
  },
  {
    label: 'menuView',
    items: [
      'view-print-layout',
      'view-web-layout',
      'view-outline',
      'view-draft',
      '---',
      'view-focus',
      '---',
      'view-ribbon',
      {
        submenu: 'menuSidebar',
        items: [
          'view-thumbnails',
          'toggle-navigation',
          'review-pane',
          { cmd: 'find', label: 'menuSidebarSearch' },
        ],
      },
      '---',
      'toggle-ruler',
      'view-header-footer',
      'view-footnotes',
      'view-markup',
      'view-reveal-formatting',
      // Word has one Zoom… item here; ours also carries the steps and the fit presets
      {
        submenu: 'menuZoom',
        items: [
          'zoom-dialog',
          '---',
          'zoom-in',
          'zoom-out',
          'zoom-100',
          '---',
          { slot: 'zoom-presets' },
          'zoom-page-width',
          'zoom-whole-page',
          'zoom-multi-page',
        ],
      },
      '---',
      'toggle-grid',
      { role: 'togglefullscreen', label: 'menuFullscreen' },
      // not in Word's View menu: kept below its items so their order stays Word's
      '---',
      'toggle-ai',
      'toggle-dark',
      { slot: 'dev-tools' },
    ],
  },
  {
    label: 'menuInsert',
    items: [
      // Word's Photo Browser, Stock Images and Online Pictures are not offered
      {
        submenu: 'menuPictures',
        items: [{ cmd: 'insert-image', label: 'menuPictureFromFile' }],
      },
      {
        submenu: 'menuShape',
        items: [
          { cmd: 'insert-shape', label: 'menuShapeRectangle', payload: 'rect' },
          { cmd: 'insert-shape', label: 'menuShapeRoundedRectangle', payload: 'roundRect' },
          { cmd: 'insert-shape', label: 'menuShapeTriangle', payload: 'triangle' },
          { cmd: 'insert-shape', label: 'menuShapeOval', payload: 'ellipse' },
          { cmd: 'insert-shape', label: 'menuShapeLine', payload: 'line' },
        ],
      },
      // Word's chart types, in its order; ours writes a column, a line and a pie chart
      {
        submenu: 'menuChart',
        items: [
          { cmd: 'insert-chart', label: 'menuChartColumn', payload: 'bar' },
          'insert-chart-bar',
          { cmd: 'insert-chart', label: 'menuChartLine', payload: 'line' },
          'insert-chart-area',
          { cmd: 'insert-chart', label: 'menuChartPie', payload: 'pie' },
          'insert-chart-treemap',
          'insert-chart-sunburst',
          'insert-chart-histogram',
          'insert-chart-pareto',
          'insert-chart-box-whisker',
          'insert-chart-scatter',
          'insert-chart-waterfall',
          'insert-chart-funnel',
          'insert-chart-stock',
          'insert-chart-surface',
          'insert-chart-radar',
        ],
      },
      'insert-table',
      '---',
      'insert-icons',
      '---',
      'insert-link',
      'insert-bookmark',
      'insert-cross-reference',
      '---',
      'insert-comment',
      '---',
      {
        submenu: 'menuBreak',
        items: [
          // no CmdOrCtrl+Enter accelerator: a menu accelerator would intercept
          // the key before renderer inputs (comments panel / prompt modal use
          // Cmd+Enter to submit); the editor keymap handles it instead
          'insert-page-break',
          'insert-column-break',
          '---',
          { cmd: 'insert-section-break', label: 'menuSectionNextPage', payload: 'nextPage' },
          { cmd: 'insert-section-break', label: 'menuSectionContinuous', payload: 'continuous' },
          { cmd: 'insert-section-break', label: 'menuSectionOddPage', payload: 'oddPage' },
          { cmd: 'insert-section-break', label: 'menuSectionEvenPage', payload: 'evenPage' },
        ],
      },
      'insert-footnote-dialog',
      'insert-caption',
      'insert-index-tables',
      'insert-watermark',
      'insert-page-numbers',
      '---',
      {
        submenu: 'menuTextBox',
        items: [
          { cmd: 'insert-text-box', label: 'menuTextBox', payload: 'draw' },
          'insert-vertical-text-box',
        ],
      },
      { submenu: 'menuAutoText', items: ['insert-autotext', 'insert-autotext-new'] },
      'insert-date-time',
      'insert-object',
      'insert-file',
      'insert-field-dialog',
      '---',
      'insert-equation',
      'insert-advanced-symbol',
    ],
  },
  {
    label: 'menuFormat',
    items: [
      'font-dialog',
      'format-text-effects',
      'paragraph-dialog',
      // Word's Document dialog holds the margins and the layout: our Layout tab
      'format-document',
      '---',
      'format-bullets-numbering',
      'format-borders-shading',
      '---',
      'format-columns',
      'format-tabs',
      'format-drop-cap',
      'format-text-direction',
      'format-change-case',
      '---',
      {
        submenu: 'menuAsianLayout',
        items: [
          'format-phonetic-guide',
          'format-enclose-characters',
          'format-horizontal-in-vertical',
          'format-combine-characters',
          'format-two-lines-in-one',
          'format-fit-text',
        ],
      },
      '---',
      // Word opens its Style dialog; ours is the Styles pane
      { cmd: 'styles-pane', label: 'menuStyle' },
      '---',
      'format-font-substitution',
      '---',
      'format-object',
      '---',
      'format-equation-options',
      '---',
      'format-selection-pane',
      // not in Word's Format menu: kept below its items so their order stays Word's
      '---',
      'bold',
      'italic',
      'underline',
      {
        submenu: 'menuAlign',
        items: ['align-left', 'align-center', 'align-right', 'align-justify'],
      },
    ],
  },
  {
    // Word for Mac keeps Word Count in the Tools menu, not on the ribbon
    label: 'menuTools',
    items: [
      {
        submenu: 'menuSpellingAndGrammar',
        items: [
          'tools-spelling',
          // the same AI proofread as Review > Editor (the renderer shows the one-time ack)
          { cmd: 'ai-proofread', label: 'menuEditor' },
          '---',
          'tools-hide-spelling',
          'tools-hide-grammar',
          '---',
          'tools-reset-ignored',
        ],
      },
      'tools-thesaurus',
      'tools-hyphenation',
      {
        submenu: 'menuTranslate',
        items: [
          'tools-translate-selection',
          'tools-translate-document',
          '---',
          'tools-translation-language',
        ],
      },
      'tools-language',
      '---',
      'word-count',
      'autocorrect-options',
      // Word's Check Accessibility is not offered
      '---',
      {
        submenu: 'menuTrackChanges',
        items: ['tools-highlight-changes', 'compare-documents'],
      },
      'tools-combine-documents',
      '---',
      'protect-document',
      '---',
      'tools-envelopes',
      'tools-labels',
      '---',
      'tools-customize-keyboard',
      // Word for Mac keeps Preferences in the application menu
      { slot: 'settings-off-mac' },
    ],
  },
  {
    // Word's Table menu; the renderer answers with a hint when the caret is outside a table
    label: 'menuTable',
    items: [
      {
        submenu: 'menuTableInsert',
        items: [
          { cmd: 'insert-table', label: 'menuTableInsertTable' },
          '---',
          'table-insert-cols-left',
          'table-insert-cols-right',
          '---',
          'table-insert-rows-above',
          'table-insert-rows-below',
          '---',
          'table-insert-cells',
        ],
      },
      {
        submenu: 'menuTableDelete',
        items: [
          'table-delete-table',
          'table-delete-columns',
          'table-delete-rows',
          'table-delete-cells',
        ],
      },
      {
        submenu: 'menuTableSelect',
        items: [
          'table-select-table',
          'table-select-column',
          'table-select-row',
          'table-select-cell',
        ],
      },
      'table-merge-cells',
      'table-split-cells',
      'table-split-table',
      '---',
      {
        submenu: 'menuTableAutoFit',
        items: [
          'table-autofit-contents',
          'table-autofit-window',
          'table-autofit-fixed',
          '---',
          'table-distribute-rows',
          'table-distribute-columns',
        ],
      },
      'table-repeat-header',
      '---',
      {
        submenu: 'menuTableConvert',
        items: ['convert-text-to-table', 'table-convert-to-text'],
      },
      { cmd: 'sort', label: 'menuTableSort' },
      'table-formula',
      'table-gridlines',
      '---',
      'table-properties',
    ],
  },
  {
    // Word's Window menu: the system's items come first and last, around New Window, Arrange All and Split
    label: 'menuWindow',
    role: 'window',
    items: [
      { slot: 'window-head' },
      '---',
      { slot: 'new-window' },
      'view-arrange-all',
      // no Word counterpart: Word has no tab strip to return windows to
      'view-merge-windows',
      'split-view',
      '---',
      { slot: 'window-tail' },
    ],
  },
] as const satisfies readonly MenuTop[]

export const DOCS_MENU = shownMenu(WORD_MENU)

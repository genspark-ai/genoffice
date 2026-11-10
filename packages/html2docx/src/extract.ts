/**
 * The in-page extractor is one self-contained function: every core fragment
 * is a slice of its body (they share one scope), assembled here at build time
 * from raw sources. The helper scripts install `__html2docx*` globals the body
 * calls into, so they are injected first.
 */
import stylesSource from './browser/core/styles.js?raw'
import runsScreenshotsSource from './browser/core/runs-screenshots.js?raw'
import tablesListsSource from './browser/core/tables-lists.js?raw'
import layoutsControlsSource from './browser/core/layouts-controls.js?raw'
import decoratorsSource from './browser/core/decorators.js?raw'
import classifierSource from './browser/core/classifier.js?raw'
import paragraphsDocumentSource from './browser/core/paragraphs-document.js?raw'
import tableHelperSource from './browser/table.js?raw'
import formsMediaHelperSource from './browser/forms-media.js?raw'
import visualEffectsHelperSource from './browser/visual-effects.js?raw'
import pagesHelperSource from './browser/pages.js?raw'

const EXTRACTOR_BODY = [
  stylesSource,
  runsScreenshotsSource,
  tablesListsSource,
  layoutsControlsSource,
  decoratorsSource,
  classifierSource,
  paragraphsDocumentSource,
].join('\n')

/** Source of `function extractIR() { … }`; evaluate `${EXTRACTOR_SOURCE}` then call it. */
export const EXTRACTOR_SOURCE = `function extractIR() {\n${EXTRACTOR_BODY}\n}`

/** Expression that runs the extractor in the page and yields the IR. */
export const EXTRACTOR_CALL = `(${EXTRACTOR_SOURCE})()`

/** Classic scripts to inject before running the extractor, in order. */
export const BROWSER_HELPER_SCRIPTS: readonly string[] = [
  tableHelperSource,
  formsMediaHelperSource,
  visualEffectsHelperSource,
  pagesHelperSource,
]

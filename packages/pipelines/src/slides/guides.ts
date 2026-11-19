import designMd from './guides/design.md?raw'
import specMd from './guides/spec.md?raw'

/** Agent-facing guides for building a deck from a spec; the app and the genoffice CLI print the same text. */
export const SLIDES_GUIDES: Readonly<
  Record<'design' | 'spec', { description: string; content: string }>
> = {
  design: {
    description:
      'Deck design workflow: style sheet first, then a page-by-page outline with layout variants, text sizing rules, anti-patterns and the QC loop',
    content: designMd,
  },
  spec: {
    description:
      'The deck spec JSON that `create --type pptx --spec` turns into a pptx: canvas, element types, fields and limits',
    content: specMd,
  },
}

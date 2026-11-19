import { Extension } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import { decodeEntities } from '@genoffice/docx-engine'
import type { AnchorBlock } from './margin-annotations'

type EndpointSeed = { id: string; block: number; paragraph: number; offset: number }
interface Endpoints {
  source?: AnchorBlock[]
  seeds: EndpointSeed[]
  positions: Map<string, number>
}

const key = new PluginKey<Endpoints>('commentEndpoints')

function importedEndpoints(source: AnchorBlock[]): EndpointSeed[] {
  const seeds: EndpointSeed[] = []
  for (const block of source) {
    if (block.docxIndex == null || !block.originalXml) continue
    let paragraph = -1
    let offset = 0
    const starts = new Set<string>()
    const tokens =
      /<w:p(?:\s[^>]*)?>|<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:(?:tab|br|cr|drawing|pict|object)\b[^>]*>|<w:commentRange(?:Start|End)\b[^>]*>/g
    for (const match of block.originalXml.matchAll(tokens)) {
      if (/^<w:p[\s>]/.test(match[0])) {
        paragraph++
        offset = 0
        starts.clear()
      } else if (match[1] !== undefined) offset += decodeEntities(match[1]).length
      else if (match[0].startsWith('<w:commentRange')) {
        const id = /\bw:id\s*=\s*["']([^"']+)["']/.exec(match[0])?.[1]
        if (id && match[0].startsWith('<w:commentRangeStart')) starts.add(id)
        else if (id && !starts.has(id))
          seeds.push({ id, block: block.docxIndex, paragraph, offset })
      } else offset++
    }
  }
  return seeds
}

function resolveEndpoints(
  doc: PmNode,
  seeds: EndpointSeed[],
  positions: Map<string, number>,
): void {
  const unresolved = seeds.filter((seed) => !positions.has(seed.id))
  if (!unresolved.length) return
  const byIndex = new Map<number, Array<{ node: PmNode; pos: number }>>()
  doc.forEach((node, pos) => {
    if (node.attrs.docxIndex == null) return
    const paragraphs: Array<{ node: PmNode; pos: number }> = []
    if (node.isTextblock) paragraphs.push({ node, pos })
    else
      node.descendants((child, offset) => {
        if (child.isTextblock) paragraphs.push({ node: child, pos: pos + 1 + offset })
      })
    byIndex.set(node.attrs.docxIndex as number, paragraphs)
  })
  for (const seed of unresolved) {
    const target = byIndex.get(seed.block)?.[seed.paragraph]
    if (
      target &&
      seed.offset <= target.node.content.size &&
      (target.node.attrs.commentEnds as string[] | null)?.includes(seed.id)
    ) {
      positions.set(seed.id, target.pos + 1 + seed.offset)
    }
  }
}

/** Screen-only imported cross-paragraph endpoints. Inline ranges already have
 * live marks; these positions preserve a precise end inside the final paragraph
 * without changing any document markers or serialized content. */
export const CommentEndpointsExtension = Extension.create({
  name: 'commentEndpoints',
  addProseMirrorPlugins() {
    return [
      new Plugin<Endpoints>({
        key,
        state: {
          init: () => ({ positions: new Map(), seeds: [] }),
          apply(tr, previous) {
            const seeded = tr.getMeta(key) as Endpoints | undefined
            if (seeded) return seeded
            if (!tr.docChanged) return previous
            const positions = new Map(
              [...previous.positions].map(([id, pos]) => [id, tr.mapping.map(pos, -1)]),
            )
            // Large documents append their later paragraphs in phases. Resolve
            // those once they arrive, without resetting already-mapped endpoints.
            resolveEndpoints(tr.doc, previous.seeds, positions)
            return { ...previous, positions }
          },
        },
      }),
    ]
  },
})

/** Seed once per imported document; subsequent edits use transaction mapping. */
export function commentEndpoints(
  view: EditorView,
  source: AnchorBlock[] | undefined,
): Map<string, number> {
  const previous = key.getState(view.state)
  if (!previous || previous.source === source) return previous?.positions ?? new Map()
  const seeds = importedEndpoints(source ?? [])
  const positions = new Map<string, number>()
  resolveEndpoints(view.state.doc, seeds, positions)
  view.dispatch(
    view.state.tr.setMeta(key, { source, seeds, positions }).setMeta('addToHistory', false),
  )
  return positions
}

// Modified for ForaOffice: section-addressed header/footer access shared by every host.
import type { AiHeaderFooterAccess } from './tools'
import { applyHfText, hfEditText, hfTextboxes } from '../editor/hf-text'
import { editKey, resolveHf, withHfEdit, type HfKind, type HfSectionState } from '../hf-sections'
import type { HfView } from '../doc-state'

export function createHeaderFooterAccess(bindings: {
  current(): HfSectionState
  commit(state: HfSectionState): void
  locked(): boolean
  evenOdd(): boolean
  enableFirst(section: number): string | null
  enableEven(): void
}): AiHeaderFooterAccess {
  const views: HfView[] = ['default', 'first', 'even']
  const kinds: HfKind[] = ['header', 'footer']
  return {
    read() {
      const state = bindings.current()
      const source = (index: number, kind: HfKind, view: HfView) => {
        const resolved = resolveHf(state, index, kind, view)
        const owner = state.sections[resolved.owner]
        return resolved.value?.sectionOnly
          ? `isolated:${resolved.owner}:${kind}:${view}`
          : (owner?.[kind === 'header' ? 'headerRefs' : 'footerRefs'][view] ??
              `section:${resolved.owner}:${kind}:${view}`)
      }
      const text = (kind: HfKind, view: HfView) =>
        hfEditText(resolveHf(state, 0, kind, view).value ?? { text: '' })
      const titlePg = state.sections[0]?.titlePg ?? false
      return {
        header: text('header', 'default'),
        footer: text('footer', 'default'),
        headerFirst: titlePg ? text('header', 'first') : null,
        footerFirst: titlePg ? text('footer', 'first') : null,
        headerEven: bindings.evenOdd() ? text('header', 'even') : null,
        footerEven: bindings.evenOdd() ? text('footer', 'even') : null,
        titlePg,
        evenOddHf: bindings.evenOdd(),
        multiSection: state.sections.length > 1,
        sections: state.sections.map((section, index) => ({
          index,
          titlePg: section.titlePg,
          strips: kinds.flatMap((kind) =>
            views.map((view) => {
              const resolved = resolveHf(state, index, kind, view)
              const textboxes = hfTextboxes(resolved.value)
              return {
                kind,
                view,
                text: hfEditText(resolved.value ?? { text: '' }),
                textEditable:
                  !resolved.value?.textReadOnly &&
                  textboxes.every((box) => box.editable) &&
                  !resolved.value?.paras?.some(
                    (para) =>
                      para.textReadOnly || ((para.box || para.boxAnchored) && !para.textbox),
                  ),
                textboxes,
                owner: resolved.owner,
                inherited: resolved.owner !== index,
                sharedWith: state.sections.flatMap((_, other) =>
                  source(index, kind, view) === source(other, kind, view) ? [other] : [],
                ),
                sharedTargets: state.sections.flatMap((_, other) =>
                  views.flatMap((v) =>
                    source(index, kind, view) === source(other, kind, v)
                      ? [{ section: other, view: v }]
                      : [],
                  ),
                ),
              }
            }),
          ),
        })),
      }
    },
    set(kind, view, text, section = 0, scope = 'section', textboxIndex) {
      if (bindings.locked()) return 'the document is read-only; headers/footers cannot be edited'
      const state = bindings.current()
      if (!Number.isInteger(section) || !state.sections[section])
        return `section ${section} does not exist`
      const current = resolveHf(state, section, kind, view).value
      if (current?.textReadOnly)
        return 'This header/footer has unsupported textbox correspondence and is read-only; no changes were made.'
      const textboxes = hfTextboxes(current)
      if (
        textboxIndex !== undefined &&
        (!Number.isInteger(textboxIndex) || !textboxes.some((box) => box.index === textboxIndex))
      )
        return 'The textboxIndex is not in the current part; read fresh header/footer context.'
      if (
        textboxes.some(
          (box) => !box.editable && (textboxIndex === undefined || box.index === textboxIndex),
        )
      )
        return 'This header/footer textbox has unsupported structured content and is read-only; no changes were made.'
      if (
        textboxIndex === undefined &&
        current?.paras?.some(
          (para) => para.textReadOnly || ((para.box || para.boxAnchored) && !para.textbox),
        )
      )
        return 'This header/footer contains unmodelled text in a drawing paragraph; target a supported textbox or leave the part unchanged.'
      const next = applyHfText(current, text, textboxIndex)
      if (scope === 'section' || current?.sectionOnly) next.sectionOnly = true
      if (view === 'first') {
        const error = bindings.enableFirst(section)
        if (error) return error
      }
      if (view === 'even') bindings.enableEven()
      const key = editKey(state.sections[section].lastBlockIndex, kind, view)
      const maps =
        scope === 'section'
          ? {
              edits: { ...state.edits, [key]: next },
              links: state.links,
            }
          : withHfEdit(state, section, kind, view, next)
      if (!maps) return 'the header/footer target is unavailable'
      bindings.commit({
        ...state,
        ...maps,
        sections:
          view === 'first'
            ? state.sections.map((s, i) => (i === section ? { ...s, titlePg: true } : s))
            : state.sections,
      })
      return null
    },
  }
}

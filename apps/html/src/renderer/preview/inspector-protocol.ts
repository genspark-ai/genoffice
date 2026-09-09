/** messages between the app (parent) and the inspector script inside the sandboxed preview frame */

/** inspect = full editing; browse = the page behaves as published */
export type InspectorMode = 'inspect' | 'browse'

export type ToInspector =
  | { type: 'gx:select'; sid: number | null }
  | { type: 'gx:highlight'; sids: number[] }
  | { type: 'gx:clearHighlight' }
  | { type: 'gx:setMode'; mode: InspectorMode }
  | { type: 'gx:beginTextEdit'; sid: number }
  | { type: 'gx:theme'; dark: boolean }
  /** poke inline styles for a live preview; the host commits them to the source as one set_style op */
  | { type: 'gx:previewStyle'; sid: number; styles: Record<string, string | null> }
  /** numbered pins on elements with queued AI edits (replaces the previous set) */
  | { type: 'gx:mark'; marks: Array<{ sid: number; label: string }> }

/** effective styles of the selected element (px numbers without unit, colours as #rrggbb, '' when transparent) */
export interface ComputedSnapshot {
  color: string
  fontSize: string
  fontWeight: string
  fontStyle: string
  textAlign: string
  background: string
  /** computed `background-image` ('none', gradients, url(...) layers) */
  backgroundImage: string
  width: string
  height: string
  borderRadius: string
  padding: string
  opacity: string
  /** authored inline values (what set_style writes back), not the resolved matrix / px */
  transform: string
  marginLeft: string
  marginRight: string
  objectFit: string
  /** only on <img> */
  image?: { src: string; alt: string; naturalWidth: number; naturalHeight: number }
}

export interface ElementRect {
  x: number
  y: number
  width: number
  height: number
}

export interface ElementFingerprint {
  sid: number
  tag: string
  className: string
  /** direct child element count and text content, compared against the source to spot script-mutated nodes */
  childElementCount: number
  text: string
  rect: ElementRect
  computed: ComputedSnapshot
  /** the element's single direct text run, when it has exactly one (inline-editable) */
  textRun: string | null
  /** index of that run among the element's direct text nodes (set_text_node's index), -1 without a run */
  textRunIndex: number
}

/** every message carries the parse-map version the frame was instrumented with */
export type FromInspector = { version: number } & FromInspectorBody

export type FromInspectorBody =
  | { type: 'gx:ready'; title: string; docHeight: number }
  | { type: 'gx:hover'; sid: number | null }
  | {
      type: 'gx:rect'
      sid: number
      rect: ElementRect
      computed: ComputedSnapshot
      textRun: string | null
      textRunIndex: number
    }
  | { type: 'gx:select'; element: ElementFingerprint | null; dynamic: boolean }
  | {
      type: 'gx:textSelect'
      sid: number
      /** index among the element's direct child text nodes */
      textNodeIndex: number
      start: number
      end: number
      text: string
    }
  | { type: 'gx:textEditCommit'; sid: number; textNodeIndex: number; newText: string }
  | { type: 'gx:textEditCancel' }
  | {
      type: 'gx:keyCommand'
      command:
        | 'delete'
        | 'parent'
        | 'next'
        | 'prev'
        | 'child'
        | 'askAi'
        | 'bold'
        | 'italic'
        | 'escape'
        | 'undo'
        | 'redo'
        | 'zoomIn'
        | 'zoomOut'
        | 'zoomReset'
    }
  /** ctrl/meta + wheel inside the frame (trackpad pinch); the host owns the zoom level */
  | { type: 'gx:zoom'; delta: number }
  | { type: 'gx:navigateBlocked'; href: string }
  | { type: 'gx:markClick'; sid: number }

export function isFromInspector(data: unknown): data is FromInspector {
  return (
    typeof data === 'object' &&
    data !== null &&
    typeof (data as { type?: unknown }).type === 'string' &&
    (data as { type: string }).type.startsWith('gx:')
  )
}

import type { FeatureRecord } from './features-skill'
import { validateOverlayShapes, type OverlayShape } from './slot-overlay'

/**
 * The slot kind's contract: where a slot may be drawn, the small element set it
 * may use, the limits that keep a registry entry from becoming a page, and the
 * declared actions a press may run.
 *
 * It is deliberately data, not code: the host renders these nodes with its own
 * components, so theme, dark mode, text scaling and i18n stay with the host and
 * nothing here can reach a document.
 */

/** Places a slot may occupy. The editor places land once every app panel shares the surface contract. */
export const SLOT_PLACES = [
  'home.sidebar',
  'home.card',
  'home.panel',
  'home.overlay',
  /** a section in an editor's own side dock, next to its assistant */
  'editor.side',
] as const
export type SlotPlace = (typeof SLOT_PLACES)[number]

/**
 * Data the host will hand a slot. A slot never reads a file itself: it names a
 * source, the host fills the list, and only three fields travel into an action
 * as $path / $name / $folder. That is what lets a mod build a file sidebar out
 * of declared data instead of code.
 */
export const SLOT_SOURCES = ['recent-files', 'starred-files'] as const
export type SlotSource = (typeof SLOT_SOURCES)[number]

/** Fields a press substitutes into an action; nothing else is substituted. */
export const SLOT_FILL_FIELDS = ['path', 'name', 'folder'] as const

/**
 * What a native view may open (K5). https anywhere; http only on the loopback,
 * for a local server the user runs. file:, data:, about:, javascript: and any
 * custom scheme are refused before a view exists — a mod may show a page, never
 * reach the machine.
 */
export function isAllowedViewUrl(url: string): boolean {
  let parsed: URL
  try {
    parsed = new URL(url)
  } catch {
    return false
  }
  if (parsed.protocol === 'https:') return true
  if (parsed.protocol !== 'http:') return false
  return ['localhost', '127.0.0.1', '[::1]', '::1'].includes(parsed.hostname)
}

/** The host shown in a view's frame while the page attaches. */
export function viewHost(url: string): string {
  try {
    return new URL(url).host
  } catch {
    return url
  }
}

export const SLOT_LIMITS = {
  /** nesting levels under the root element */
  maxDepth: 4,
  /** every node in the tree, root included */
  maxElements: 24,
  /** characters in any one text, label, placeholder or markdown value */
  maxTextChars: 200,
  /** entries in one select or list */
  maxItems: 32,
  /** rows a file-list may ask the host for */
  maxSourceRows: 50,
} as const

export type SlotTextTone = 'default' | 'muted' | 'danger'
export type SlotTextSize = 'body' | 'caption'

export type SlotAction =
  | { kind: 'open'; path: string }
  | { kind: 'ask-assistant'; prompt: string }
  /** open that feature's own home.panel */
  | { kind: 'panel'; id: string }
  /** accepted now, but nothing executes a feature yet: the assistant is asked instead */
  | { kind: 'run-feature'; id: string }

export interface SlotListItem {
  label: string
  /** reported back with the select action; defaults to the label */
  value?: string
  description?: string
}

export type SlotNode =
  | { type: 'text'; value: string; tone?: SlotTextTone; size?: SlotTextSize }
  | { type: 'button'; label: string; tone?: 'default' | 'primary'; action: SlotAction }
  | { type: 'input'; name?: string; placeholder?: string; submitLabel?: string; action: SlotAction }
  | { type: 'select'; name?: string; options: SlotListItem[]; action?: SlotAction }
  | { type: 'list'; items: SlotListItem[]; action?: SlotAction }
  /**
   * A list the host fills from a named source. Pressing an item substitutes
   * $path, $name and $folder into the action, so one declaration covers every
   * file the source returns.
   */
  | {
      type: 'file-list'
      source: SlotSource
      /** rows to ask for (default 8, at most SLOT_LIMITS.maxSourceRows) */
      limit?: number
      action?: SlotAction
      emptyText?: string
    }
  /**
   * A real view the host mounts over this rectangle: today a browser. The slot
   * says where and which address; the host owns the view, the frame and the
   * rules above.
   */
  | { type: 'view'; view: 'browser'; url: string; title?: string; minHeight?: number }
  /**
   * A screen-only layer drawn over the content area (K4). Coordinates are
   * fractions, tones are host names, and the layer takes no pointer events: it
   * guides the eye and nothing else. A document never sees it.
   */
  | { type: 'overlay'; shapes: OverlayShape[] }
  | { type: 'divider' }
  | { type: 'markdown'; value: string }
  /** a container: the only node that nests */
  | { type: 'group'; children: SlotNode[] }

export interface SlotDeclaration {
  at: SlotPlace
  element: SlotNode
}

export interface SlotValidationError {
  /** dotted path to the offending field, e.g. "element.children[0].label" */
  field: string
  message: string
}

export type SlotValidation =
  { ok: true; slot: SlotDeclaration } | { ok: false; errors: SlotValidationError[] }

const ACTIONS = ['open', 'ask-assistant', 'panel', 'run-feature'] as const

function text(value: unknown): value is string {
  return typeof value === 'string' && value.trim().length > 0
}

function shortText(
  value: unknown,
  field: string,
  errors: SlotValidationError[],
  optional = false,
): void {
  if (value === undefined && optional) return
  if (!text(value)) {
    errors.push({ field, message: 'needs a non-empty string' })
    return
  }
  if (value.length > SLOT_LIMITS.maxTextChars) {
    errors.push({
      field,
      message: 'is longer than ' + SLOT_LIMITS.maxTextChars + ' characters',
    })
  }
}

/** Validate a declared action. Exported because a command has the same rules. */
export function validateSlotAction(
  value: unknown,
  field: string,
  errors: SlotValidationError[],
): void {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    errors.push({ field, message: 'needs an action object' })
    return
  }
  const action = value as Record<string, unknown>
  const kind = action.kind
  if (typeof kind !== 'string' || !(ACTIONS as readonly string[]).includes(kind)) {
    errors.push({ field: field + '.kind', message: 'must be one of ' + ACTIONS.join(', ') })
    return
  }
  if (kind === 'open') shortText(action.path, field + '.path', errors)
  if (kind === 'ask-assistant') shortText(action.prompt, field + '.prompt', errors)
  if (kind === 'panel' || kind === 'run-feature') shortText(action.id, field + '.id', errors)
}

function validateItems(value: unknown, field: string, errors: SlotValidationError[]): void {
  if (!Array.isArray(value) || value.length === 0) {
    errors.push({ field, message: 'needs at least one entry' })
    return
  }
  if (value.length > SLOT_LIMITS.maxItems) {
    errors.push({ field, message: 'has more than ' + SLOT_LIMITS.maxItems + ' entries' })
  }
  value.forEach((item, index) => {
    const at = field + '[' + index + ']'
    if (!item || typeof item !== 'object' || Array.isArray(item)) {
      errors.push({ field: at, message: 'needs an object' })
      return
    }
    const entry = item as Record<string, unknown>
    shortText(entry.label, at + '.label', errors)
    shortText(entry.value, at + '.value', errors, true)
    shortText(entry.description, at + '.description', errors, true)
  })
}

function validateNode(
  value: unknown,
  field: string,
  depth: number,
  errors: SlotValidationError[],
): number {
  if (depth > SLOT_LIMITS.maxDepth) {
    errors.push({ field, message: 'nests deeper than ' + SLOT_LIMITS.maxDepth + ' levels' })
    return 0
  }
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    errors.push({ field, message: 'needs an element object' })
    return 0
  }
  const node = value as Record<string, unknown>
  const type = node.type
  if (typeof type !== 'string') {
    errors.push({ field: field + '.type', message: 'needs a type' })
    return 0
  }
  let count = 1
  switch (type) {
    case 'text':
      shortText(node.value, field + '.value', errors)
      if (node.tone !== undefined && !['default', 'muted', 'danger'].includes(String(node.tone))) {
        errors.push({ field: field + '.tone', message: 'must be default, muted or danger' })
      }
      if (node.size !== undefined && !['body', 'caption'].includes(String(node.size))) {
        errors.push({ field: field + '.size', message: 'must be body or caption' })
      }
      break
    case 'button':
      shortText(node.label, field + '.label', errors)
      if (node.tone !== undefined && !['default', 'primary'].includes(String(node.tone))) {
        errors.push({ field: field + '.tone', message: 'must be default or primary' })
      }
      validateSlotAction(node.action, field + '.action', errors)
      break
    case 'input':
      shortText(node.name, field + '.name', errors, true)
      shortText(node.placeholder, field + '.placeholder', errors, true)
      shortText(node.submitLabel, field + '.submitLabel', errors, true)
      validateSlotAction(node.action, field + '.action', errors)
      break
    case 'select':
      shortText(node.name, field + '.name', errors, true)
      validateItems(node.options, field + '.options', errors)
      if (node.action !== undefined) validateSlotAction(node.action, field + '.action', errors)
      break
    case 'list':
      validateItems(node.items, field + '.items', errors)
      if (node.action !== undefined) validateSlotAction(node.action, field + '.action', errors)
      break
    case 'file-list': {
      if (
        typeof node.source !== 'string' ||
        !(SLOT_SOURCES as readonly string[]).includes(node.source)
      ) {
        errors.push({
          field: field + '.source',
          message: 'must be a known source (' + SLOT_SOURCES.join(', ') + ')',
        })
      }
      if (node.limit !== undefined) {
        const limit = Number(node.limit)
        if (!Number.isInteger(limit) || limit < 1 || limit > SLOT_LIMITS.maxSourceRows) {
          errors.push({
            field: field + '.limit',
            message: 'must be a whole number between 1 and ' + SLOT_LIMITS.maxSourceRows,
          })
        }
      }
      shortText(node.emptyText, field + '.emptyText', errors, true)
      if (node.action !== undefined) validateSlotAction(node.action, field + '.action', errors)
      break
    }
    case 'view': {
      if (node.view !== 'browser') {
        errors.push({
          field: field + '.view',
          message: 'must be "browser": it is the only native view the host mounts',
        })
      }
      if (typeof node.url !== 'string' || !isAllowedViewUrl(node.url)) {
        errors.push({
          field: field + '.url',
          message: 'must be an https address (or http on localhost)',
        })
      }
      shortText(node.title, field + '.title', errors, true)
      if (node.minHeight !== undefined) {
        const minHeight = Number(node.minHeight)
        if (!Number.isInteger(minHeight) || minHeight < 80 || minHeight > 1200) {
          errors.push({
            field: field + '.minHeight',
            message: 'must be a whole number of pixels between 80 and 1200',
          })
        }
      }
      break
    }
    case 'overlay':
      validateOverlayShapes(node.shapes, field + '.shapes', errors)
      break
    case 'divider':
      break
    case 'markdown':
      shortText(node.value, field + '.value', errors)
      break
    case 'group': {
      const children = node.children
      if (!Array.isArray(children) || children.length === 0) {
        errors.push({ field: field + '.children', message: 'needs at least one child' })
        break
      }
      children.forEach((child, index) => {
        count += validateNode(child, field + '.children[' + index + ']', depth + 1, errors)
      })
      break
    }
    default:
      errors.push({
        field: field + '.type',
        message:
          'unknown element type "' +
          type +
          '" (text, button, input, select, list, file-list, view, overlay, divider, markdown, group)',
      })
  }
  return count
}

/** Validate a slot declaration from the assistant or a mod file. */
export function validateSlot(value: unknown): SlotValidation {
  const errors: SlotValidationError[] = []
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    return { ok: false, errors: [{ field: 'slot', message: 'needs a slot object' }] }
  }
  const slot = value as Record<string, unknown>
  if (typeof slot.at !== 'string' || !(SLOT_PLACES as readonly string[]).includes(slot.at)) {
    errors.push({
      field: 'slot.at',
      message: 'must be a known place (' + SLOT_PLACES.join(', ') + ')',
    })
  }
  const elements = validateNode(slot.element, 'slot.element', 1, errors)
  if (elements > SLOT_LIMITS.maxElements) {
    errors.push({
      field: 'slot.element',
      message:
        'draws ' + elements + ' elements, more than the ' + SLOT_LIMITS.maxElements + ' allowed',
    })
  }
  if (errors.length) return { ok: false, errors }
  return { ok: true, slot: slot as unknown as SlotDeclaration }
}

/** The slots a registry holds for one place, in registry order. */
export function slotsAt(features: readonly FeatureRecord[], place: SlotPlace): FeatureRecord[] {
  return features.filter((feature) => feature.slot?.at === place)
}

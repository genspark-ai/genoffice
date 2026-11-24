/**
 * Keyboard customizations a document carries: the wne:keymaps of its
 * Customizations part (word/customizations.xml, root wne:tcg; [MS-OE376]
 * 3.1.2.1.1). Only key maps that run a fixed command (wne:fci) or mask a
 * built-in key are understood; every other entry, and everything else in the
 * part (toolbars, allocated commands, key maps for other keyboards), is kept
 * as it is written.
 */
import type JSZip from 'jszip'
import { parseRels, resolveRelationshipTargetPath } from './parse-package'

export const KEY_CUSTOMIZATIONS_REL_TYPE =
  'http://schemas.microsoft.com/office/2006/relationships/keyMapCustomizations'
export const KEY_CUSTOMIZATIONS_CONTENT_TYPE = 'application/vnd.ms-word.keyMapCustomizations+xml'
export const KEY_CUSTOMIZATIONS_PART_PATH = 'word/customizations.xml'

const WNE_NS = 'http://schemas.microsoft.com/office/word/2006/wordml'
const DOCUMENT_PART = 'word/document.xml'

/** modifier bits of a KCM's high byte */
export const KCM_SHIFT = 0x01
export const KCM_CONTROL = 0x02
export const KCM_ALT = 0x04
export const KCM_COMMAND = 0x08

export interface KeymapEntry {
  /** the entry as the part writes it; absent for one made here */
  raw?: string
  /** modifier byte and Windows virtual key code of the first key; null when the entry names no key code */
  kcmPrimary: number | null
  /** set for a two-key sequence */
  kcmSecondary: number | null
  /** the built-in key is blocked and runs nothing */
  mask: boolean
  /** the fixed command the key runs; null for a macro, a character, an allocated command, or none */
  fciName: string | null
}

const ATTR = /([\w.:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g

function attrs(tag: string): Map<string, string> {
  const out = new Map<string, string>()
  for (const m of tag.matchAll(ATTR)) out.set(m[1].replace(/^[^:]*:/, ''), m[2] ?? m[3] ?? '')
  return out
}

const hex = (text: string | undefined): number | null =>
  text !== undefined && /^[0-9a-f]{1,4}$/i.test(text) ? Number.parseInt(text, 16) : null

const isOn = (text: string | undefined): boolean =>
  text !== undefined && !['0', 'false', 'off'].includes(text.toLowerCase())

const escapeRe = (text: string): string => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

/** the prefix the part binds to the wne namespace, or null when its root is not a prefixed tcg */
function prefixOf(xml: string): string | null {
  const root = /<([\w.-]+):tcg\b[^>]*>/.exec(xml)
  if (!root) return null
  const bound = new RegExp(`\\bxmlns:${escapeRe(root[1])}\\s*=\\s*["']${escapeRe(WNE_NS)}["']`)
  return bound.test(root[0]) ? root[1] : null
}

const keymapsBlock = (p: string): RegExp =>
  new RegExp(`<${p}:keymaps\\b[^>]*?(?:/>|>([\\s\\S]*?)</${p}:keymaps>)`)

/** the key maps of a Customizations part for this keyboard, in file order */
export function readKeymaps(xml: string): KeymapEntry[] {
  const p = prefixOf(xml)
  if (!p) return []
  const prefix = escapeRe(p)
  const inner = keymapsBlock(prefix).exec(xml)?.[1] ?? ''
  const entry = new RegExp(`<${prefix}:keymap\\b[^>]*?(?:/>|>([\\s\\S]*?)</${prefix}:keymap>)`, 'g')
  const out: KeymapEntry[] = []
  for (const m of inner.matchAll(entry)) {
    const own = attrs(/^<[^>]*>/.exec(m[0])?.[0] ?? '')
    const action = new RegExp(`<${prefix}:(\\w+)\\b[^>]*>`).exec(m[1] ?? '')
    const fci = action?.[1] === 'fci' ? attrs(action[0]).get('fciName') : undefined
    out.push({
      raw: m[0],
      kcmPrimary: hex(own.get('kcmPrimary')),
      kcmSecondary: hex(own.get('kcmSecondary')),
      mask: isOn(own.get('mask')),
      fciName: fci || null,
    })
  }
  return out
}

const hex4 = (value: number): string => value.toString(16).toUpperCase().padStart(4, '0')

const escapeAttr = (text: string): string =>
  text.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')

function entryXml(p: string, entry: KeymapEntry): string {
  if (entry.raw !== undefined) return entry.raw
  const key = `${p}:kcmPrimary="${hex4(entry.kcmPrimary ?? 0)}"`
  if (entry.mask) return `<${p}:keymap ${p}:mask="1" ${key}/>`
  return (
    `<${p}:keymap ${key}>` +
    `<${p}:fci ${p}:fciName="${escapeAttr(entry.fciName ?? '')}" ${p}:swArg="0000"/>` +
    `</${p}:keymap>`
  )
}

/**
 * The Customizations part with `entries` as its key maps for this keyboard.
 * An entry that carries `raw` is written as it was; the rest of the part is
 * untouched. `xml` is null for a document without the part. Returns null when
 * the part cannot be edited (its root is not a tcg in the wne namespace), or
 * when there is no part and nothing to write.
 */
export function writeKeymaps(xml: string | null, entries: readonly KeymapEntry[]): string | null {
  if (xml === null) {
    if (entries.length === 0) return null
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      `<wne:tcg xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wne="${WNE_NS}">` +
      `<wne:keymaps>${entries.map((e) => entryXml('wne', e)).join('')}</wne:keymaps></wne:tcg>`
    )
  }
  const p = prefixOf(xml)
  if (!p) return null
  const prefix = escapeRe(p)
  // a keymaps element needs at least one keymap
  const block =
    entries.length > 0
      ? `<${p}:keymaps>${entries.map((e) => entryXml(p, e)).join('')}</${p}:keymaps>`
      : ''
  const existing = keymapsBlock(prefix)
  if (existing.test(xml)) return xml.replace(existing, () => block)
  if (!block) return xml
  // keymaps comes first in a tcg
  const empty = new RegExp(`<${prefix}:tcg\\b([^>]*?)/>`)
  if (empty.test(xml)) {
    return xml.replace(empty, (_, rest: string) => `<${p}:tcg${rest}>${block}</${p}:tcg>`)
  }
  return xml.replace(new RegExp(`<${prefix}:tcg\\b[^>]*>`), (open) => open + block)
}

/** where the main document's Customizations part is, or null when it has none */
export async function keyCustomizationsPath(zip: JSZip): Promise<string | null> {
  for (const rel of (await parseRels(zip, 'word/_rels/document.xml.rels')).values()) {
    if (rel.type !== KEY_CUSTOMIZATIONS_REL_TYPE || rel.targetMode === 'External') continue
    const part = resolveRelationshipTargetPath(DOCUMENT_PART, rel.target)
    if (part && zip.file(part)) return part
  }
  return null
}

export async function readKeyCustomizations(zip: JSZip): Promise<string | undefined> {
  const part = await keyCustomizationsPath(zip)
  return part ? zip.file(part)!.async('string') : undefined
}

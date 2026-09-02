// Package-level parts: main document path, relationships, comments,
// protection settings and numbering definitions.
import JSZip from 'jszip'

import {
  attrsOf,
  childrenOf,
  findChild,
  findChildren,
  nameOf,
  textOf,
  xmlParser,
  type XNode,
} from './xml-utils'
import { customEnumItems } from './list-markers'
import { decodeNumericCharRefs, type RelInfo } from './parse-xml-text'
import type {
  CommentInfo,
  DocProtection,
  NumberingDef,
  NumberingLevel,
  WriteProtection,
} from './types'

/**
 * Main document part: word/document.xml when present, else the package-level
 * officeDocument relationship target (LO corpus has e.g. word/trial.xml).
 */
export async function resolveMainDocumentPath(zip: JSZip): Promise<string | null> {
  if (zip.file('word/document.xml')) return 'word/document.xml'
  const rels = await parseRels(zip, '_rels/.rels')
  for (const rel of rels.values()) {
    if (!/\/officeDocument$/.test(rel.type) || rel.targetMode === 'External') continue
    const target = rel.target.replace(/^\//, '')
    if (zip.file(target)) return target
  }
  return null
}

export async function parseRels(zip: JSZip, path: string): Promise<Map<string, RelInfo>> {
  const rels = new Map<string, RelInfo>()
  const file = zip.file(path)
  if (!file) return rels
  // fast-xml-parser rejects a DOCTYPE declaring external entities; drop the
  // prologue instead of failing the whole document (entities never resolve —
  // XXE-safe — and Relationship elements carry everything in attributes)
  const relsXml = (await file.async('string')).replace(/<!DOCTYPE(?:[^>[]|\[[\s\S]*?\])*>/i, '')
  const parsed = xmlParser.parse(relsXml) as XNode[]
  const root = parsed.find((n) => nameOf(n) === 'Relationships')
  if (!root) return rels
  for (const relNode of findChildren(root, 'Relationship')) {
    const attrs = attrsOf(relNode)
    if (!attrs['Id']) continue
    rels.set(attrs['Id'], {
      target: attrs['Target'] ?? '',
      type: attrs['Type'] ?? '',
      targetMode: attrs['TargetMode'],
    })
  }
  return rels
}

/** word/comments.xml (+ reply/resolved relations from commentsExtended) -> display list, file order */
export async function parseComments(zip: JSZip): Promise<CommentInfo[]> {
  const file = zip.file('word/comments.xml')
  if (!file) return []
  const parsed = xmlParser.parse(await file.async('string')) as XNode[]
  const root = parsed.find((n) => nameOf(n) === 'w:comments')
  if (!root) return []
  const out: CommentInfo[] = []
  for (const node of findChildren(root, 'w:comment')) {
    const attrs = attrsOf(node)
    if (!attrs['w:id']) continue
    const paras = findChildren(node, 'w:p')
    const paraId = paras.length > 0 ? attrsOf(paras[paras.length - 1])['w14:paraId'] : undefined
    out.push({
      id: attrs['w:id'],
      author: attrs['w:author'] ?? '',
      initials: attrs['w:initials'],
      date: attrs['w:date'],
      text: paras.map((p) => textOf(p)).join('\n'),
      ...(paraId ? { paraId } : {}),
    })
  }
  // commentsExtended.xml: paraId → parent paraId / done (Word 2013+ replies and resolution)
  const extFile = zip.file('word/commentsExtended.xml')
  if (extFile) {
    const extXml = await extFile.async('string')
    const byParaId = new Map(out.filter((c) => c.paraId).map((c) => [c.paraId!, c]))
    for (const m of extXml.match(/<w15:commentEx [^>]*\/>/g) ?? []) {
      const paraId = /w15:paraId="([^"]+)"/.exec(m)?.[1]
      const parentParaId = /w15:paraIdParent="([^"]+)"/.exec(m)?.[1]
      const done = /w15:done="(?:1|true)"/.test(m)
      const c = paraId ? byParaId.get(paraId) : undefined
      if (!c) continue
      if (done) c.done = true
      if (parentParaId) {
        const parent = byParaId.get(parentParaId)
        if (parent) c.parentId = parent.id
      }
    }
  }
  return out
}

/** w:documentProtection from word/settings.xml (editing restriction) */
export async function parseProtection(zip: JSZip): Promise<DocProtection | null> {
  const file = zip.file('word/settings.xml')
  if (!file) return null
  const xml = await file.async('string')
  const tag = /<w:documentProtection[^>]*\/>/.exec(xml)?.[0]
  if (!tag) return null
  const edit = /w:edit="([^"]+)"/.exec(tag)?.[1]
  if (!edit || edit === 'none') return null
  const enforcement = /w:enforcement="([^"]+)"/.exec(tag)?.[1]
  const hash = /w:hash="([^"]+)"/.exec(tag)?.[1]
  const salt = /w:salt="([^"]+)"/.exec(tag)?.[1]
  const spin = /w:cryptSpinCount="(\d+)"/.exec(tag)?.[1]
  const sid = /w:cryptAlgorithmSid="(\d+)"/.exec(tag)?.[1]
  return {
    edit,
    enforced: enforcement === '1' || enforcement === 'true',
    ...(hash ? { hash } : {}),
    ...(salt ? { salt } : {}),
    ...(spin ? { spinCount: parseInt(spin, 10) } : {}),
    ...(sid ? { algorithmSid: parseInt(sid, 10) } : {}),
  }
}

/** w:writeProtection from word/settings.xml (password to modify / read-only recommended) */
export async function parseWriteProtection(zip: JSZip): Promise<WriteProtection | null> {
  const file = zip.file('word/settings.xml')
  if (!file) return null
  const tag = /<w:writeProtection[^>]*\/>/.exec(await file.async('string'))?.[0]
  if (!tag) return null
  const recommended = /w:recommended="(?:1|true|on)"/.test(tag)
  const hash = /w:hash="([^"]+)"/.exec(tag)?.[1]
  const salt = /w:salt="([^"]+)"/.exec(tag)?.[1]
  const spin = /w:cryptSpinCount="(\d+)"/.exec(tag)?.[1]
  const sid = /w:cryptAlgorithmSid="(\d+)"/.exec(tag)?.[1]
  if (!recommended && !hash) return null
  return {
    ...(recommended ? { recommended } : {}),
    ...(hash ? { hash } : {}),
    ...(salt ? { salt } : {}),
    ...(spin ? { spinCount: parseInt(spin, 10) } : {}),
    ...(sid ? { algorithmSid: parseInt(sid, 10) } : {}),
  }
}

/** settings.xml removePersonalInformation (namespace prefix and quote style are arbitrary). */
export async function parseRemovePersonalInfo(zip: JSZip): Promise<boolean> {
  const file = zip.file('word/settings.xml')
  if (!file) return false
  const xml = await file.async('string')
  const prefixes = new Set<string>()
  const namespace = /\bxmlns(?::([A-Za-z_][\w.-]*))?\s*=\s*(["'])([^"']*)\2/g
  let declaration: RegExpExecArray | null
  while ((declaration = namespace.exec(xml)) !== null) {
    if (
      declaration[3] === 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' ||
      declaration[3] === 'http://purl.oclc.org/ooxml/wordprocessingml/main'
    ) {
      prefixes.add(declaration[1] ?? '')
    }
  }
  const escapedPrefixes = [...prefixes]
    .filter(Boolean)
    .map((prefix) => prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))
  for (const prefix of prefixes) {
    const qName = prefix ? `${prefix}:removePersonalInformation` : 'removePersonalInformation'
    const escapedName = qName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    const tag = new RegExp(`<${escapedName}\\b[^>]*(?:\\/\\s*>|>\\s*<\\/${escapedName}\\s*>)`).exec(
      xml,
    )?.[0]
    if (!tag) continue
    const valPrefix = escapedPrefixes.length > 0 ? `(?:${escapedPrefixes.join('|')}):` : ''
    const val = new RegExp(`(?:^|\\s)(?:${valPrefix})?val\\s*=\\s*(["'])(0|false)\\1`, 'i')
    return !val.test(tag)
  }
  return false
}

/** w:numFmt of a w:lvl, including w14 custom formats hidden in mc:AlternateContent
 *  (mc:Choice carries val="custom" + w:format; mc:Fallback the standard substitute) */
function numFmtOfLevel(lvlNode: XNode): { numFmt?: string; customFormat?: string } {
  const direct = findChild(lvlNode, 'w:numFmt')
  if (direct) return { numFmt: attrsOf(direct)['w:val'] }
  const alt = childrenOf(lvlNode).find((c) => nameOf(c)?.endsWith(':AlternateContent'))
  if (!alt) return {}
  const pick = (local: string) => childrenOf(alt).find((c) => nameOf(c)?.endsWith(`:${local}`))
  const choice = attrsOf(findChild(pick('Choice') ?? {}, 'w:numFmt') ?? {})
  const format = choice['w:format']
  if (choice['w:val'] === 'custom' && format && customEnumItems(format))
    return { numFmt: 'custom', customFormat: format }
  if (choice['w:val'] && choice['w:val'] !== 'custom') return { numFmt: choice['w:val'] }
  return { numFmt: attrsOf(findChild(pick('Fallback') ?? {}, 'w:numFmt') ?? {})['w:val'] }
}

function parseNumberingLevel(lvlNode: XNode): NumberingLevel {
  // ECMA-376: a w:lvl without w:start starts at 0 (Word renders "0.")
  const start = parseInt(attrsOf(findChild(lvlNode, 'w:start') ?? {})['w:val'] ?? '0', 10)
  const { numFmt, customFormat } = numFmtOfLevel(lvlNode)
  const level: NumberingLevel = {
    numFmt: numFmt ?? 'decimal',
    lvlText: decodeNumericCharRefs(attrsOf(findChild(lvlNode, 'w:lvlText') ?? {})['w:val'] ?? ''),
    start: Number.isFinite(start) ? start : 0,
  }
  if (customFormat) level.customFormat = customFormat
  const suff = attrsOf(findChild(lvlNode, 'w:suff') ?? {})['w:val']
  if (suff === 'space' || suff === 'nothing' || suff === 'tab') level.suff = suff
  const lvlPPr = findChild(lvlNode, 'w:pPr')
  const ind = lvlPPr ? findChild(lvlPPr, 'w:ind') : undefined
  if (ind) {
    const attrs = attrsOf(ind)
    const left = parseInt(attrs['w:left'] ?? attrs['w:start'] ?? '', 10)
    if (left > 0) level.indentLeft = left
    const hanging = parseInt(attrs['w:hanging'] ?? '', 10)
    if (hanging > 0) level.hanging = hanging
    const firstLine = parseInt(attrs['w:firstLine'] ?? '', 10)
    if (!level.hanging && firstLine > 0) level.firstLine = firstLine
  }
  const lvlRPr = findChild(lvlNode, 'w:rPr')
  const sz = lvlRPr ? parseInt(attrsOf(findChild(lvlRPr, 'w:sz') ?? {})['w:val'] ?? '', 10) : NaN
  if (sz > 0) level.szHalfPoints = sz
  const fonts = lvlRPr ? attrsOf(findChild(lvlRPr, 'w:rFonts') ?? {}) : {}
  const font = fonts['w:ascii'] ?? fonts['w:hAnsi'] ?? fonts['w:eastAsia']
  if (font) level.font = font
  return level
}

/** word/numbering.xml -> per-numId level definitions + the bullet/ordered classification */
export async function parseNumbering(
  zip: JSZip,
): Promise<{ formats: Map<string, 'bullet' | 'ordered'>; defs: Map<string, NumberingDef> }> {
  const formats = new Map<string, 'bullet' | 'ordered'>()
  const defs = new Map<string, NumberingDef>()
  const file = zip.file('word/numbering.xml')
  if (!file) return { formats, defs }
  const parsed = xmlParser.parse(await file.async('string')) as XNode[]
  const root = parsed.find((n) => nameOf(n) === 'w:numbering')
  if (!root) return { formats, defs }

  const absLevels = new Map<string, Record<number, NumberingLevel>>()
  const numStyleLinks = new Map<string, string>()
  const styleLinkAbs = new Map<string, string>()
  for (const abs of findChildren(root, 'w:abstractNum')) {
    const absId = attrsOf(abs)['w:abstractNumId']
    if (!absId) continue
    const levels: Record<number, NumberingLevel> = {}
    for (const lvl of findChildren(abs, 'w:lvl')) {
      const ilvl = parseInt(attrsOf(lvl)['w:ilvl'] ?? '', 10)
      if (Number.isFinite(ilvl)) levels[ilvl] = parseNumberingLevel(lvl)
    }
    absLevels.set(absId, levels)
    const numStyleLink = attrsOf(findChild(abs, 'w:numStyleLink') ?? {})['w:val']
    if (numStyleLink) numStyleLinks.set(absId, numStyleLink)
    const styleLink = attrsOf(findChild(abs, 'w:styleLink') ?? {})['w:val']
    if (styleLink) styleLinkAbs.set(styleLink, absId)
  }
  // w:numStyleLink indirection: the abstractNum is a reference to a numbering style;
  // the real levels live on the abstractNum carrying the matching w:styleLink
  for (const absId of numStyleLinks.keys()) {
    const seen = new Set([absId])
    let target = absId
    for (;;) {
      const styleId = numStyleLinks.get(target)
      const next = styleId !== undefined ? styleLinkAbs.get(styleId) : undefined
      if (next === undefined || seen.has(next)) break
      seen.add(next)
      target = next
    }
    if (target !== absId) {
      absLevels.set(absId, { ...absLevels.get(target), ...absLevels.get(absId) })
    }
  }
  for (const num of findChildren(root, 'w:num')) {
    const numId = attrsOf(num)['w:numId']
    const absId = attrsOf(findChild(num, 'w:abstractNumId') ?? {})['w:val']
    if (!numId || !absId) continue
    const levels: Record<number, NumberingLevel> = { ...(absLevels.get(absId) ?? {}) }
    const startOverrides: Record<number, number> = {}
    for (const over of findChildren(num, 'w:lvlOverride')) {
      const ilvl = parseInt(attrsOf(over)['w:ilvl'] ?? '', 10)
      if (!Number.isFinite(ilvl)) continue
      const startVal = attrsOf(findChild(over, 'w:startOverride') ?? {})['w:val']
      if (startVal !== undefined) {
        const n = parseInt(startVal, 10)
        if (Number.isFinite(n)) startOverrides[ilvl] = n
      }
      const lvl = findChild(over, 'w:lvl')
      if (lvl) levels[ilvl] = parseNumberingLevel(lvl)
    }
    defs.set(numId, { numId, abstractNumId: absId, levels, startOverrides })
    formats.set(numId, levels[0]?.numFmt === 'bullet' ? 'bullet' : 'ordered')
  }
  return { formats, defs }
}

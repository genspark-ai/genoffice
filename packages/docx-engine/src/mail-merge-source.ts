import JSZip from 'jszip'

import { loadDocxZip } from './zip-load'

/**
 * A merged document carries the merged values, not the data source it came
 * from. Word keeps the data source in word/settings.xml: a w:mailMerge block
 * whose w:dataSource / w:headerSource are relationship ids, plus the ODSO
 * connection string (w:udl) and query the block records, and
 * word/_rels/settings.xml.rels carries the external relationship — the data
 * file's absolute path. None of it belongs in a merged letter, so a save that
 * writes a merge product removes it. genoffice#2097.
 *
 * The template itself keeps its block: only merge output is stripped, and only
 * through the save's `exportAsMerged` switch.
 */

const SETTINGS_PATH = 'word/settings.xml'
const SETTINGS_RELS_PATH = 'word/_rels/settings.xml.rels'

/** a whole <w:mailMerge> element, prefixed or not, paired or self-closing */
const MAIL_MERGE_RE =
  /<(?:[A-Za-z_][\w.-]*:)?mailMerge\b[^>]*(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?mailMerge\s*>)/

/** relationship-id attributes (r:id on dataSource / headerSource / srcRel) */
const REL_ID_RE = /\sr:id\s*=\s*(?:"([^"]*)"|'([^']*)')/g

/** one Relationship tag, paired or self-closing, any prefix */
const RELATIONSHIP_RE =
  /<(?:[A-Za-z_][\w.-]*:)?Relationship\b[^>]*(?:\/\s*>|>\s*<\/(?:[A-Za-z_][\w.-]*:)?Relationship\s*>)/g

function attribute(tag: string, localName: string): string | undefined {
  const escaped = localName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  const match = new RegExp(
    `(?:^|\\s)(?:[A-Za-z_][\\w.-]*:)?${escaped}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`,
  ).exec(tag)
  return match?.[1] ?? match?.[2]
}

/**
 * Drop the mail merge data-source reference from a package: the w:mailMerge
 * block in settings.xml and the relationships it names. Returns whether
 * anything was removed.
 */
export async function stripMailMergeSourceFromZip(zip: JSZip): Promise<boolean> {
  const settings = zip.file(SETTINGS_PATH)
  if (!settings) return false
  const xml = await settings.async('string')
  const block = MAIL_MERGE_RE.exec(xml)?.[0]
  if (!block) return false

  const relIds = new Set<string>()
  for (const match of block.matchAll(REL_ID_RE)) {
    const id = match[1] ?? match[2]
    if (id) relIds.add(id)
  }

  zip.file(SETTINGS_PATH, xml.replace(MAIL_MERGE_RE, ''), { date: settings.date })

  if (relIds.size === 0) return true
  const rels = zip.file(SETTINGS_RELS_PATH)
  if (!rels) return true
  const relsXml = await rels.async('string')
  const next = relsXml.replace(RELATIONSHIP_RE, (tag) => {
    const id = attribute(tag, 'Id')
    return id && relIds.has(id) ? '' : tag
  })
  if (next !== relsXml) zip.file(SETTINGS_RELS_PATH, next, { date: rels.date })
  return true
}

/**
 * Strip the mail merge data-source reference from saved .docx bytes. A
 * document without a w:mailMerge block comes back byte-identical.
 */
export async function stripMailMergeSource(bytes: Uint8Array): Promise<Uint8Array> {
  const zip = await loadDocxZip(bytes)
  const removed = await stripMailMergeSourceFromZip(zip)
  if (!removed) return bytes
  return zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
}

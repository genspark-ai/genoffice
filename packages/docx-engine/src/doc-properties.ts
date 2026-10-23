import JSZip from 'jszip'
import {
  CUSTOM_PROPERTIES_CONTENT_TYPE,
  CUSTOM_PROPERTIES_PATH,
  CUSTOM_PROPERTIES_REL_TYPE,
  PROPERTY_RE,
  attrValue,
  decodeXmlText,
  propertyText,
} from './zotero-doc-props'
import { escapeXmlAttr, escapeXmlText } from './xml-utils'

/** docProps/core.xml — the core properties part. */
export const CORE_PROPERTIES_PATH = 'docProps/core.xml'
/** docProps/app.xml — the extended properties part. */
export const APP_PROPERTIES_PATH = 'docProps/app.xml'

export const CORE_PROPERTIES_CONTENT_TYPE =
  'application/vnd.openxmlformats-package.core-properties+xml'
export const CORE_PROPERTIES_REL_TYPE =
  'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
export const APP_PROPERTIES_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.extended-properties+xml'
export const APP_PROPERTIES_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'

/** the fmtid every Word custom property carries */
const CUSTOM_PROPERTIES_FMTID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}'

/** friendly key -> the element of docProps/core.xml that holds it */
const CORE_PROPERTIES: Record<string, { prefix: string; local: string }> = {
  title: { prefix: 'dc', local: 'title' },
  subject: { prefix: 'dc', local: 'subject' },
  creator: { prefix: 'dc', local: 'creator' },
  keywords: { prefix: 'cp', local: 'keywords' },
  description: { prefix: 'dc', local: 'description' },
  lastModifiedBy: { prefix: 'cp', local: 'lastModifiedBy' },
  revision: { prefix: 'cp', local: 'revision' },
  created: { prefix: 'dcterms', local: 'created' },
  modified: { prefix: 'dcterms', local: 'modified' },
  category: { prefix: 'cp', local: 'category' },
  contentStatus: { prefix: 'cp', local: 'contentStatus' },
  identifier: { prefix: 'dc', local: 'identifier' },
  language: { prefix: 'dc', local: 'language' },
  version: { prefix: 'cp', local: 'version' },
}

/** friendly key -> the element of docProps/app.xml that holds it */
const APP_PROPERTIES: Record<string, string> = {
  template: 'Template',
  totalTime: 'TotalTime',
  pages: 'Pages',
  words: 'Words',
  characters: 'Characters',
  application: 'Application',
  docSecurity: 'DocSecurity',
  lines: 'Lines',
  paragraphs: 'Paragraphs',
  scaleCrop: 'ScaleCrop',
  company: 'Company',
  linksUpToDate: 'LinksUpToDate',
  charactersWithSpaces: 'CharactersWithSpaces',
  sharedDoc: 'SharedDoc',
  hyperlinksChanged: 'HyperlinksChanged',
  appVersion: 'AppVersion',
  manager: 'Manager',
}

/**
 * The schema order of CT_Properties: app.xml is a sequence, so an element that
 * is not there yet has to be written where the schema puts it.
 */
const APP_PROPERTY_ORDER = [
  'Template',
  'Manager',
  'Company',
  'Pages',
  'Words',
  'Characters',
  'PresentationFormat',
  'Lines',
  'Paragraphs',
  'Slides',
  'Notes',
  'TotalTime',
  'HiddenSlides',
  'MMClips',
  'ScaleCrop',
  'HeadingPairs',
  'TitlesOfParts',
  'LinksUpToDate',
  'CharactersWithSpaces',
  'SharedDoc',
  'HyperlinkBase',
  'HLinks',
  'HyperlinksChanged',
  'DigSig',
  'Application',
  'AppVersion',
  'DocSecurity',
]

/** the core property keys, in the order the part lists them */
export const CORE_PROPERTY_KEYS = Object.keys(CORE_PROPERTIES)
/** the extended (app) property keys */
export const APP_PROPERTY_KEYS = Object.keys(APP_PROPERTIES)

/** the document properties of a .docx, read from its three docProps parts */
export interface DocumentProperties {
  /** docProps/core.xml by friendly key (title, creator, lastModifiedBy, …) */
  core: Record<string, string>
  /** docProps/app.xml by friendly key (company, template, application, …) */
  app: Record<string, string>
  /** docProps/custom.xml by property name */
  custom: Record<string, string>
}

/** what a set_properties op asks for: a value per key, or null to remove it */
export interface PropertyChanges {
  core?: Record<string, string | null>
  app?: Record<string, string | null>
  custom?: Record<string, string | null>
  /** remove every property of all three parts (a full metadata scrub) */
  clear?: 'all'
}

/**
 * `<dc:title>`, `<dc:title/>` and `<title>`: the element by local name, the
 * prefix optional because a part may bind it to another prefix.
 */
function elementRe(local: string): RegExp {
  return new RegExp(
    `<([A-Za-z_][\\w.-]*:)?${local}(?=[\\s/>])([^>]*?)(?:/>|>([\\s\\S]*?)</([A-Za-z_][\\w.-]*:)?${local}\\s*>)`,
    'g',
  )
}

/** the text of an element, null when the part does not carry it */
function elementText(xml: string, local: string): string | null {
  const match = elementRe(local).exec(xml)
  return match ? decodeXmlText(match[3] ?? '') : null
}

/** insert an element before the root's closing tag */
function insertElement(xml: string, element: string): string {
  const at = xml.lastIndexOf('</')
  return at === -1 ? xml + element : xml.slice(0, at) + element + xml.slice(at)
}

/** replace the text of the element, keeping its attributes; insert it when missing */
function setElementText(xml: string, local: string, value: string, element: string): string {
  let found = false
  const out = xml.replace(elementRe(local), (_match, openPrefix, attrs) => {
    found = true
    const name = `${openPrefix ?? ''}${local}`
    return `<${name}${attrs}>${escapeXmlText(value)}</${name}>`
  })
  return found ? out : insertElement(xml, element)
}

/** the core properties of a core.xml part */
export function parseCorePropertiesXml(xml: string): Record<string, string> {
  const out: Record<string, string> = {}
  for (const [key, { local }] of Object.entries(CORE_PROPERTIES)) {
    const text = elementText(xml, local)
    if (text !== null) out[key] = text
  }
  return out
}

/** the extended properties of an app.xml part */
export function parseAppPropertiesXml(xml: string): Record<string, string> {
  const out: Record<string, string> = {}
  for (const [key, local] of Object.entries(APP_PROPERTIES)) {
    const text = elementText(xml, local)
    if (text !== null) out[key] = text
  }
  return out
}

const VT_VALUE_RE =
  /<(?:[A-Za-z_][\w.-]*:)?(lpwstr|lpstr|bstr|i1|i2|i4|i8|ui1|ui2|ui4|ui8|int|uint|r4|r8|bool|filetime|date|decimal)\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?\1\s*>/

function customPropertyValue(property: string): string {
  const match = VT_VALUE_RE.exec(property)
  return match ? decodeXmlText(match[2]) : propertyText(property)
}

/** the custom properties of a custom.xml part, by property name */
export function parseCustomPropertiesXml(xml: string): Record<string, string> {
  const out: Record<string, string> = {}
  for (const property of xml.match(PROPERTY_RE) ?? []) {
    const name = attrValue(property, 'name')
    if (name) out[name] = customPropertyValue(property)
  }
  return out
}

/** the three docProps parts of a package, empty when the part is missing */
export async function readDocumentProperties(zip: JSZip): Promise<DocumentProperties> {
  const core = zip.file(CORE_PROPERTIES_PATH)
  const app = zip.file(APP_PROPERTIES_PATH)
  const custom = zip.file(CUSTOM_PROPERTIES_PATH)
  return {
    core: core ? parseCorePropertiesXml(await core.async('string')) : {},
    app: app ? parseAppPropertiesXml(await app.async('string')) : {},
    custom: custom ? parseCustomPropertiesXml(await custom.async('string')) : {},
  }
}

/** the properties of a package that carries no docProps part at all */
export function emptyDocumentProperties(): DocumentProperties {
  return { core: {}, app: {}, custom: {} }
}

function blankCorePropertiesXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"' +
    ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"' +
    ' xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    '</cp:coreProperties>'
  )
}

function blankAppPropertiesXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"' +
    ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    '</Properties>'
  )
}

function blankCustomPropertiesXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"' +
    ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    '</Properties>'
  )
}

/** every child of the root goes, known or not; the root itself stays */
function clearRootChildrenXml(xml: string, root: string): string {
  const re = new RegExp(
    `(<(?:[A-Za-z_][\\w.-]*:)?${root}(?=[\\s>])[^>]*>)[\\s\\S]*(</(?:[A-Za-z_][\\w.-]*:)?${root}\\s*>)`,
  )
  return xml.replace(re, '$1$2')
}

function clearCorePropertiesXml(xml: string): string {
  return clearRootChildrenXml(xml, 'coreProperties')
}

function clearAppPropertiesXml(xml: string): string {
  return clearRootChildrenXml(xml, 'Properties')
}

/** every custom property goes; the part itself stays */
function clearCustomPropertiesXml(xml: string): string {
  return xml.replace(PROPERTY_RE, '')
}

/** set or remove the core properties of a core.xml part */
export function patchCorePropertiesXml(
  xml: string,
  changes: Record<string, string | null>,
): string {
  let out = xml
  for (const [key, value] of Object.entries(changes)) {
    const spec = CORE_PROPERTIES[key]
    if (!spec) continue
    if (value === null) {
      out = out.replace(elementRe(spec.local), '')
      continue
    }
    const name = out.includes(`xmlns:${spec.prefix}=`) ? `${spec.prefix}:${spec.local}` : spec.local
    const attrs =
      (spec.local === 'created' || spec.local === 'modified') && out.includes('xmlns:xsi=')
        ? ' xsi:type="dcterms:W3CDTF"'
        : ''
    out = setElementText(
      out,
      spec.local,
      value,
      `<${name}${attrs}>${escapeXmlText(value)}</${name}>`,
    )
  }
  return out
}

/** set or remove the extended properties of an app.xml part */
export function patchAppPropertiesXml(xml: string, changes: Record<string, string | null>): string {
  let out = xml
  for (const [key, value] of Object.entries(changes)) {
    const local = APP_PROPERTIES[key]
    if (!local) continue
    if (value === null) {
      out = out.replace(elementRe(local), '')
      continue
    }
    out = setAppElement(out, local, value)
  }
  return out
}

/** an app.xml element goes where the schema sequence puts it */
function setAppElement(xml: string, local: string, value: string): string {
  let found = false
  const out = xml.replace(elementRe(local), (_match, openPrefix, attrs) => {
    found = true
    const name = `${openPrefix ?? ''}${local}`
    return `<${name}${attrs}>${escapeXmlText(value)}</${name}>`
  })
  if (found) return out
  const element = `<${local}>${escapeXmlText(value)}</${local}>`
  const at = APP_PROPERTY_ORDER.indexOf(local)
  for (const other of APP_PROPERTY_ORDER) {
    if (APP_PROPERTY_ORDER.indexOf(other) <= at) continue
    const match = elementRe(other).exec(xml)
    if (match) return xml.slice(0, match.index) + element + xml.slice(match.index)
  }
  return insertElement(xml, element)
}

/** set or remove the custom properties of a custom.xml part */
export function patchCustomPropertiesXml(
  xml: string,
  changes: Record<string, string | null>,
): string {
  let out = xml
  let maxPid = 1
  for (const property of out.match(PROPERTY_RE) ?? []) {
    const pid = Number(attrValue(property, 'pid'))
    if (Number.isSafeInteger(pid)) maxPid = Math.max(maxPid, pid)
  }
  for (const [name, value] of Object.entries(changes)) {
    if (value === null) {
      out = out.replace(PROPERTY_RE, (property) =>
        attrValue(property, 'name') === name ? '' : property,
      )
      continue
    }
    let replaced = false
    out = out.replace(PROPERTY_RE, (property) => {
      if (attrValue(property, 'name') !== name) return property
      replaced = true
      const fmtid = attrValue(property, 'fmtid') ?? CUSTOM_PROPERTIES_FMTID
      const pid = attrValue(property, 'pid') ?? String(maxPid + 1)
      return (
        `<property fmtid="${escapeXmlAttr(fmtid)}" pid="${escapeXmlAttr(pid)}"` +
        ` name="${escapeXmlAttr(name)}"><vt:lpwstr>${escapeXmlText(value)}</vt:lpwstr></property>`
      )
    })
    if (replaced) continue
    maxPid += 1
    out = insertElement(
      out,
      `<property fmtid="${CUSTOM_PROPERTIES_FMTID}" pid="${maxPid}"` +
        ` name="${escapeXmlAttr(name)}"><vt:lpwstr>${escapeXmlText(value)}</vt:lpwstr></property>`,
    )
  }
  return out
}

function maxRelId(relsXml: string): number {
  let max = 0
  for (const match of relsXml.matchAll(/\bId\s*=\s*["']rId(\d+)["']/g)) {
    const n = parseInt(match[1]!, 10)
    if (n > max) max = n
  }
  return max
}

async function addContentTypeOverride(
  zip: JSZip,
  partName: string,
  contentType: string,
): Promise<void> {
  const file = zip.file('[Content_Types].xml')
  if (!file) return
  const xml = await file.async('string')
  if (xml.includes(`PartName="${partName}"`)) return
  zip.file(
    '[Content_Types].xml',
    xml.replace(
      '</Types>',
      `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`,
    ),
    { date: file.date },
  )
}

async function addRootRelationship(zip: JSZip, relType: string, target: string): Promise<void> {
  const path = '_rels/.rels'
  const file = zip.file(path)
  const xml = file
    ? await file.async('string')
    : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
  if (xml.includes(`Type="${relType}"`)) return
  const id = `rId${maxRelId(xml) + 1}`
  const out = xml.replace(
    '</Relationships>',
    `<Relationship Id="${id}" Type="${relType}" Target="${target}"/></Relationships>`,
  )
  zip.file(path, out, file ? { date: file.date } : undefined)
}

/** write a docProps part, creating it (with its content type and relationship) when missing */
async function writePropertiesPart(
  zip: JSZip,
  path: string,
  xml: string,
  contentType: string,
  relType: string,
): Promise<void> {
  const entry = zip.file(path)
  if (entry) {
    zip.file(path, xml, { date: entry.date })
    return
  }
  zip.file(path, xml)
  await addContentTypeOverride(zip, `/${path}`, contentType)
  await addRootRelationship(zip, relType, path)
}

const hasSetValue = (values: Record<string, string | null> | undefined): boolean =>
  !!values && Object.values(values).some((value) => value !== null)

/**
 * Write the requested property changes into the package: a string sets a
 * property, null removes it, clear: "all" empties all three parts. A part that
 * is not there yet is created, so a document without docProps can still take
 * properties.
 */
export async function applyPropertyChanges(zip: JSZip, changes: PropertyChanges): Promise<void> {
  const clearAll = changes.clear === 'all'
  if (!clearAll && !changes.core && !changes.app && !changes.custom) return

  if (clearAll || changes.core) {
    const entry = zip.file(CORE_PROPERTIES_PATH)
    const values = changes.core ?? {}
    if (entry || hasSetValue(values)) {
      const xml = entry ? await entry.async('string') : blankCorePropertiesXml()
      // the clear comes first: a value given in the same op still lands
      const out = patchCorePropertiesXml(clearAll ? clearCorePropertiesXml(xml) : xml, values)
      await writePropertiesPart(
        zip,
        CORE_PROPERTIES_PATH,
        out,
        CORE_PROPERTIES_CONTENT_TYPE,
        CORE_PROPERTIES_REL_TYPE,
      )
    }
  }

  if (clearAll || changes.app) {
    const entry = zip.file(APP_PROPERTIES_PATH)
    const values = changes.app ?? {}
    if (entry || hasSetValue(values)) {
      const xml = entry ? await entry.async('string') : blankAppPropertiesXml()
      const out = patchAppPropertiesXml(clearAll ? clearAppPropertiesXml(xml) : xml, values)
      await writePropertiesPart(
        zip,
        APP_PROPERTIES_PATH,
        out,
        APP_PROPERTIES_CONTENT_TYPE,
        APP_PROPERTIES_REL_TYPE,
      )
    }
  }

  if (clearAll || changes.custom) {
    const entry = zip.file(CUSTOM_PROPERTIES_PATH)
    const values = changes.custom ?? {}
    if (entry || hasSetValue(values)) {
      const xml = entry ? await entry.async('string') : blankCustomPropertiesXml()
      const out = patchCustomPropertiesXml(clearAll ? clearCustomPropertiesXml(xml) : xml, values)
      await writePropertiesPart(
        zip,
        CUSTOM_PROPERTIES_PATH,
        out,
        CUSTOM_PROPERTIES_CONTENT_TYPE,
        CUSTOM_PROPERTIES_REL_TYPE,
      )
    }
  }
}

/**
 * The set_properties op: core / app / custom hold the values to set (null
 * removes), clear: "all" empties every part. Throws on anything else.
 */
export function parsePropertyChanges(input: Record<string, unknown>): PropertyChanges {
  const changes: PropertyChanges = {}
  if (input.clear !== undefined) {
    if (input.clear !== 'all') throw new Error('set_properties: clear must be "all"')
    changes.clear = 'all'
  }
  for (const section of ['core', 'app', 'custom'] as const) {
    const values = input[section]
    if (values === undefined) continue
    if (!values || typeof values !== 'object' || Array.isArray(values)) {
      throw new Error(`set_properties: ${section} must be an object of property values`)
    }
    const known =
      section === 'core' ? CORE_PROPERTY_KEYS : section === 'app' ? APP_PROPERTY_KEYS : []
    const out: Record<string, string | null> = {}
    for (const [key, value] of Object.entries(values as Record<string, unknown>)) {
      if (known.length > 0 && !known.includes(key)) {
        throw new Error(
          `set_properties: unknown ${section} property "${key}" (known: ${known.join(', ')})`,
        )
      }
      if (value === null) {
        out[key] = null
        continue
      }
      if (typeof value !== 'string') {
        throw new Error(`set_properties: ${section}.${key} must be a string or null`)
      }
      out[key] = value
    }
    changes[section] = out
  }
  if (!changes.clear && isEmpty(changes.core) && isEmpty(changes.app) && isEmpty(changes.custom)) {
    throw new Error('set_properties: nothing to change (give core, app, custom or clear)')
  }
  return changes
}

function isEmpty(values: Record<string, string | null> | undefined): boolean {
  return !values || Object.keys(values).length === 0
}

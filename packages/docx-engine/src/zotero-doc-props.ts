import JSZip from 'jszip'
import { escapeXmlAttr, escapeXmlText } from './xml-utils'

export const CUSTOM_PROPERTIES_PATH = 'docProps/custom.xml'
export const CUSTOM_PROPERTIES_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'
export const CUSTOM_PROPERTIES_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.custom-properties+xml'

const ZOTERO_PREF_NAME = /^ZOTERO_PREF(?:_(\d+))?$/
const PROPERTY_RE =
  /<(?:[A-Za-z_][\w.-]*:)?property\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?property\s*>/g

function decodeXmlText(value: string): string {
  return value
    .replace(/&#x([0-9a-f]+);/gi, (_match, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_match, decimal) => String.fromCodePoint(parseInt(decimal, 10)))
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

function attrValue(xml: string, name: string): string | null {
  const match = new RegExp(`\\b${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(xml)
  return match ? decodeXmlText(match[1] ?? match[2] ?? '') : null
}

function propertyText(xml: string): string {
  const match =
    /<(?:[A-Za-z_][\w.-]*:)?lpwstr\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?lpwstr\s*>/.exec(xml)
  return decodeXmlText(match?.[1] ?? '')
}

export function parseZoteroDocumentDataXml(xml: string): string {
  const chunks: Array<{ index: number; value: string }> = []
  let legacy = ''
  for (const property of xml.match(PROPERTY_RE) ?? []) {
    const name = attrValue(property, 'name')
    const match = name ? ZOTERO_PREF_NAME.exec(name) : null
    if (!match) continue
    const value = propertyText(property)
    if (match[1]) chunks.push({ index: Number(match[1]), value })
    else legacy = value
  }
  if (chunks.length === 0) return legacy
  return chunks
    .sort((a, b) => a.index - b.index)
    .map((chunk) => chunk.value)
    .join('')
}

export async function readZoteroDocumentData(zip: JSZip): Promise<string> {
  const file = zip.file(CUSTOM_PROPERTIES_PATH)
  return file ? parseZoteroDocumentDataXml(await file.async('string')) : ''
}

export function patchZoteroDocumentDataXml(xml: string | null, data: string): string {
  const original =
    xml ??
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" ' +
      'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"></Properties>'
  let maxPid = 1
  for (const property of original.match(PROPERTY_RE) ?? []) {
    const pid = Number(attrValue(property, 'pid'))
    if (Number.isSafeInteger(pid)) maxPid = Math.max(maxPid, pid)
  }
  const withoutOld = original.replace(PROPERTY_RE, (property) => {
    const name = attrValue(property, 'name')
    return name && ZOTERO_PREF_NAME.test(name) ? '' : property
  })
  const chunks: string[] = []
  for (let offset = 0; offset < data.length; offset += 255)
    chunks.push(data.slice(offset, offset + 255))
  const properties = chunks
    .map((value, index) => {
      const name = `ZOTERO_PREF_${index + 1}`
      return (
        `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${++maxPid}"` +
        ` name="${escapeXmlAttr(name)}"><vt:lpwstr>${escapeXmlText(value)}</vt:lpwstr></property>`
      )
    })
    .join('')
  return withoutOld.replace(/<\/(?:[A-Za-z_][\w.-]*:)?Properties\s*>/, `${properties}$&`)
}

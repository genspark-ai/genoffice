import JSZip from 'jszip'
import {
  THEME_CONTENT_TYPE,
  THEME_REL_TYPE,
  applyThemeFormat,
  completeThemeElements,
  editedThemeXml,
  readThemeColors,
  readThemeFonts,
  readThemeFormat,
  type ThemeEdit,
  type ThemeFormat,
} from './theme'
import type { ThemeColors, ThemeFonts } from './types'
import { loadDocxZip } from './zip-load'

/**
 * A theme as a file of its own. An Office theme file (.thmx) is a package whose
 * main part is an empty theme manager pointing at the theme part; a document or
 * a template carries the same theme part beside its main part.
 */

const TYPES = '[Content_Types].xml'
const ROOT_RELS = '_rels/.rels'
const MANAGER_PART = 'theme/theme/themeManager.xml'
const MANAGER_RELS = 'theme/theme/_rels/themeManager.xml.rels'
const THMX_THEME_PART = 'theme/theme/theme1.xml'
const OFFICE_DOCUMENT_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
const MANAGER_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.themeManager+xml'
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

/** what a theme file gives a document: the parts of a theme pick */
export interface ThemeFileTheme {
  name: string
  fonts: ThemeFonts
  colors: ThemeColors
  format: ThemeFormat
}

const SOLID_FILL = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'

/** a picture fill names a picture of the package it came from: a plain fill stands in for it */
const withoutPictureFills = (xml: string): string =>
  xml.replace(/<a:blipFill\b[^>]*?(?:\/>|>[\s\S]*?<\/a:blipFill>)/g, SOLID_FILL)

/** the theme part as the values a pick carries; null when it is not a theme with colours and faces */
export function themeOfXml(themeXml: string): ThemeFileTheme | null {
  if (!/<a:theme[\s>]/.test(themeXml)) return null
  const xml = completeThemeElements(themeXml)
  const fonts = readThemeFonts(xml)
  const colors = readThemeColors(xml)
  const format = readThemeFormat(xml)
  if (!fonts || !colors || !format) return null
  const fontsName = /<a:fontScheme name="([^"]*)"/.exec(xml)?.[1]
  return {
    name: format.name,
    // an empty slot is the theme's choice too: the document's own face does not stay in it
    fonts: {
      eastAsia: '',
      majorCs: '',
      minorCs: '',
      ...fonts,
      ...(fontsName ? { name: fontsName } : {}),
    },
    colors,
    format: { name: format.name, xml: withoutPictureFills(format.xml) },
  }
}

/** the theme part a document of this theme would be saved with */
export const themeXmlOf = (theme: ThemeEdit): string => editedThemeXml(undefined, theme)

const targetsOf = (relsXml: string, type: string): string[] =>
  [...relsXml.matchAll(/<Relationship\b[^>]*?\/?>/g)]
    .filter((m) => m[0].includes(`Type="${type}"`) && !/TargetMode="External"/.test(m[0]))
    .flatMap((m) => /Target="([^"]*)"/.exec(m[0])?.[1] ?? [])

/** a relationship target as a part path, from the part that owns the relationship */
function resolveTarget(owner: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const out = owner.split('/').slice(0, -1)
  for (const piece of target.split('/')) {
    if (piece === '..') out.pop()
    else if (piece !== '.' && piece !== '') out.push(piece)
  }
  return out.join('/')
}

const relsOf = (part: string): string => {
  const at = part.lastIndexOf('/')
  return `${part.slice(0, at + 1)}_rels/${part.slice(at + 1)}.rels`
}

/**
 * The theme part of a theme file, a document or a template: the one its main
 * part points at. null when the package has none.
 */
export async function readThemeFileXml(bytes: Uint8Array): Promise<string | null> {
  const zip = await loadDocxZip(bytes)
  const text = (part: string) => zip.file(part)?.async('string')
  const rootRels = await text(ROOT_RELS)
  for (const main of rootRels ? targetsOf(rootRels, OFFICE_DOCUMENT_REL) : []) {
    const mainPart = resolveTarget('', main)
    const rels = await text(relsOf(mainPart))
    for (const target of rels ? targetsOf(rels, THEME_REL_TYPE) : []) {
      const xml = await text(resolveTarget(mainPart, target))
      if (xml !== undefined) return xml
    }
  }
  return null
}

/** the document's theme part as a save of `edit` leaves it; one is made for a document without */
export async function readDocxThemeXml(bytes: Uint8Array, edit: ThemeEdit = {}): Promise<string> {
  return editedThemeXml((await readThemeFileXml(bytes)) ?? undefined, edit)
}

const relationships = (items: Array<[id: string, type: string, target: string]>): string =>
  `${XML_HEAD}<Relationships xmlns="${RELS_NS}">` +
  items
    .map(([id, type, target]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`)
    .join('') +
  '</Relationships>'

/** an Office theme file holding `themeXml`, under `name` when one is given */
export function buildThmx(themeXml: string, name?: string): Promise<Uint8Array> {
  const complete = completeThemeElements(themeXml)
  const format = readThemeFormat(complete)
  const theme = name && format ? applyThemeFormat(complete, { name, xml: format.xml }) : complete
  const zip = new JSZip()
  zip.file(
    TYPES,
    `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      `<Override PartName="/${MANAGER_PART}" ContentType="${MANAGER_CONTENT_TYPE}"/>` +
      `<Override PartName="/${THMX_THEME_PART}" ContentType="${THEME_CONTENT_TYPE}"/>` +
      '</Types>',
  )
  zip.file(ROOT_RELS, relationships([['rId1', OFFICE_DOCUMENT_REL, MANAGER_PART]]))
  zip.file(
    MANAGER_PART,
    `${XML_HEAD}<a:themeManager xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>`,
  )
  zip.file(MANAGER_RELS, relationships([['rId1', THEME_REL_TYPE, 'theme1.xml']]))
  zip.file(THMX_THEME_PART, theme)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

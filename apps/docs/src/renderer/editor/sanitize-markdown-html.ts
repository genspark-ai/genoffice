/**
 * Sanitiser for the HTML `marked` produced from a plain-text Markdown paste.
 *
 * `marked` is a converter, not a sanitiser: it copies link targets and raw
 * inline HTML through untouched, so whatever the clipboard carried becomes
 * markup. CSP covers inline event handlers, but a `javascript:` href is a
 * navigation, not an inline handler, so no script-src rule stops it — and the
 * target is stored in the document model and written back out to .docx. The
 * converted HTML is therefore cleaned here, before it reaches the ProseMirror
 * node view, so every consumer of the paste path shares one gate.
 *
 * URL-bearing attributes go through `safeExternalUrl`, the protocol allowlist
 * the rest of the suite already applies to `shell.openExternal`: a target
 * this app would refuse to open is never stored in the first place.
 */
import { safeExternalUrl } from '@genoffice/electron-utils/safe-external-url'
import { PERSISTABLE_IMAGE_URL } from './persistable-image'

/// Protocol allowlist for a link target, matching the open-external gate.
const LINK_PROTOCOLS = ['http:', 'https:', 'mailto:'] as const

/**
 * Elements that execute, embed a foreign document, or take over URL
 * resolution. Dropped with their subtree: their content is not prose the
 * author pasted. Anything else (`h1`, `table`, `input[type=checkbox]`, ...)
 * stays, so the conversion still reproduces the Markdown structure.
 */
const DROPPED_TAGS = [
  'applet',
  'base',
  'embed',
  'form',
  'frame',
  'frameset',
  'iframe',
  'link',
  'math',
  'meta',
  'noscript',
  'object',
  'script',
  'style',
  'svg',
  'template',
]

/**
 * Attributes that carry a URL. Checking the element is not enough: the same
 * scheme hole opens through `src`, `action` and `poster` as through `href`.
 */
const URL_ATTRIBUTES = new Set([
  'action',
  'background',
  'cite',
  'data',
  'formaction',
  'href',
  'longdesc',
  'ping',
  'poster',
  'src',
  'xlink:href',
])

/// Attributes that execute or load a document inline whatever their value.
const DROPPED_ATTRIBUTES = new Set(['srcdoc'])

/** An in-document anchor has no protocol, so it cannot execute. */
const isFragment = (value: string): boolean => value.startsWith('#')

function cleanAttributes(el: Element): void {
  for (const { name, value } of Array.from(el.attributes)) {
    const attr = name.toLowerCase()
    // an inline handler is the one thing CSP would have covered; drop it here
    // so the markup is safe for any consumer, not just this editor
    if (DROPPED_ATTRIBUTES.has(attr) || attr.startsWith('on')) {
      el.removeAttribute(name)
      continue
    }
    if (!URL_ATTRIBUTES.has(attr)) continue
    const url = value.trim()
    if (isFragment(url) || safeExternalUrl(url, { allowedProtocols: LINK_PROTOCOLS })) continue
    // an inline picture is the one data: URL the image node persists; a data:
    // document behind a link or any other attribute stays out
    if (attr === 'src' && el.tagName === 'IMG' && PERSISTABLE_IMAGE_URL.test(url)) continue
    el.removeAttribute(name)
  }
  for (const child of Array.from(el.children)) cleanAttributes(child)
}

/**
 * The input with executable markup and unsafe link targets removed. A target
 * that is not allowed is dropped as an attribute rather than as the link, so
 * the pasted text stays in the document and only the destination is lost.
 */
export function sanitizeMarkdownHtml(html: string): string {
  const body = new DOMParser().parseFromString(html, 'text/html').body
  for (const el of Array.from(body.querySelectorAll(DROPPED_TAGS.join(',')))) el.remove()
  for (const el of Array.from(body.querySelectorAll('*'))) cleanAttributes(el)
  return body.innerHTML
}

/**
 * @vitest-environment jsdom
 */
/// <reference lib="dom" />
// The package's own tsconfig deliberately omits the DOM lib — it is a headless
// engine — so the two globals this file needs are pulled in here rather than by
// changing what the package compiles against.

import { describe, expect, it } from 'vitest'
import { REDACT_NS, REDACT_PREFIX, redactExtXml, setRedactExt } from '../src/redaction-xml'

/**
 * The saved slide has to be XML PowerPoint can open.
 *
 * The mark rides in an `<a:ext>` whose child is an element in a namespace of
 * ours. An undeclared prefix is not a warning, it is a parse error — PowerPoint
 * offers to repair the deck, and repair drops the run carrying the mark, so the
 * span quietly becomes readable by a model again on the next open.
 *
 * This file exists separately from `redaction-xml.test.ts` for one reason: that
 * suite runs in node and reads the fragment back with `fast-xml-parser`, which is
 * lenient about undeclared prefixes. That leniency is why a deck no
 * presentation would open passed every round-trip test in the package.
 */

const RPR = '<a:rPr lang="en-US" sz="1800"><a:latin typeface="Calibri"/></a:rPr>'

function parseStrictly(xml: string): Document {
  return new DOMParser().parseFromString(xml, 'application/xml')
}

/** a slide-shaped document wrapping the given run properties */
function asSlide(rPr: string): string {
  return (
    '<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:sp><p:txBody>' +
    `<a:p><a:r>${rPr}<a:t>Call now</a:t></a:r></a:p>` +
    '</p:txBody></p:sp></p:spTree></p:cSld></p:sld>'
  )
}

describe('the marker is XML a presentation can open', () => {
  it('is well-formed with the mark in it', () => {
    const doc = parseStrictly(asSlide(setRedactExt(RPR, 'client phone')))
    expect(doc.querySelector('parsererror')).toBeNull()
  })

  it('declares the namespace on the extension, so the prefix resolves', () => {
    // wrapped, not parsed bare: an `<a:ext>` on its own has no `a` either, and
    // that would be a third unbound prefix saying nothing about ours
    const ext = parseStrictly(asSlide(`<a:rPr>${redactExtXml('client phone')}</a:rPr>`))
    expect(ext.querySelector('parsererror')).toBeNull()
    // Declared on the `a:ext` rather than on the part root: `a:extLst` is
    // OOXML's own extension point and a consumer is required to skip the
    // content of an `<a:ext>` whose uri it does not know, so the prefix
    // resolving is the whole requirement.
    expect(redactExtXml('x')).toContain(`xmlns:${REDACT_PREFIX}="${REDACT_NS}"`)
  })

  it('leaves the label where a namespace-aware reader can find it', () => {
    const doc = parseStrictly(asSlide(setRedactExt(RPR, 'client phone')))
    expect(doc.getElementsByTagNameNS(REDACT_NS, 'redact').length).toBe(1)
  })
})

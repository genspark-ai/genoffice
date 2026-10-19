/**
 * Recorded narration (Slide Show > Record), stored the way PowerPoint stores it so PowerPoint
 * plays it back too: per slide, an audio picture in the bottom-right corner that starts with the
 * slide, under a `<p:audio isNarration="1">` media node whose `showWhenStopped="0"` hides it
 * during the show, and presProps' `<p:showPr showNarration="1">` for the whole deck.
 */
import { elementSpid, getSlideAnimations, setSlideAnimations } from './animation'
import type { OpenedPptx } from './index'
import { deleteElement } from './insert'
import { addMedia } from './media-insert'
import { addContentTypeOverride, appendRelationship } from './notes'
import type { Slide } from './types'
import { resolveTarget } from './zip'

const PRES_PATH = 'ppt/presentation.xml'
const PRES_PROPS_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps'
const PRES_PROPS_CT = 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml'
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
/** The size PowerPoint gives a recorded narration's icon: a third of an inch. */
const NARRATION_ICON_EMU = 304800
export const NARRATION_NAME = 'Narration'

const NARRATION_NODE = /<p:audio\b[^>]*\bisNarration="(?:1|true)"[^>]*>[\s\S]*?<\/p:audio>/g

export function narrationSpids(slide: Slide): number[] {
  const spids: number[] = []
  for (const node of slide.bodySuffix.matchAll(NARRATION_NODE)) {
    const spid = /<p:spTgt spid="(\d+)"/.exec(node[0])
    if (spid) spids.push(Number(spid[1]))
  }
  return spids
}

/** Removes the slide's narration: its pictures, their timing and their media parts. */
export function clearNarration(opened: OpenedPptx, slideIndex: number): number {
  const slide = opened.deck.slides[slideIndex]
  if (!slide) return 0
  const spids = new Set(narrationSpids(slide))
  const targets = slide.elements.filter((el) => spids.has(elementSpid(el) ?? -1))
  for (const el of targets) deleteElement(opened, slide, el.id)
  return targets.length
}

function markNarration(bodySuffix: string, spid: number): string {
  return bodySuffix.replace(/<p:audio>([\s\S]*?)<\/p:audio>/g, (node, inner: string) => {
    if (!inner.includes(`<p:spTgt spid="${spid}"/>`)) return node
    const media = inner.replace(/<p:cMediaNode\b/, '<p:cMediaNode showWhenStopped="0"')
    return `<p:audio isNarration="1">${media}</p:audio>`
  })
}

export interface NarrationAudio {
  bytes: Uint8Array
  /** wav, m4a or mp3 */
  ext: string
  durationMs: number
}

/**
 * Replaces the slide's narration with `audio`, which plays from the slide's start, before the
 * slide's own animations. The slide's advance time is not touched: see setSlideAdvanceTime.
 */
export function addNarration(
  opened: OpenedPptx,
  slideIndex: number,
  audio: NarrationAudio,
): { slide: Slide; elementId: string } | null {
  clearNarration(opened, slideIndex)
  const { cx, cy } = opened.deck.size
  const added = addMedia(opened, slideIndex, {
    kind: 'audio',
    bytes: audio.bytes,
    ext: audio.ext,
    name: NARRATION_NAME,
    offset: {
      x: cx - NARRATION_ICON_EMU,
      y: cy - NARRATION_ICON_EMU,
      cx: NARRATION_ICON_EMU,
      cy: NARRATION_ICON_EMU,
    },
  })
  if (!added) return null
  const { slide } = added
  const el = slide.elements.find((e) => e.id === added.elementId)
  const spid = el ? elementSpid(el) : null
  if (spid == null) return null
  setSlideAnimations(slide, [
    {
      spid,
      effect: 'mediaPlay',
      trigger: 'afterPrev',
      durationMs: audio.durationMs,
      delayMs: 0,
      mediaKind: 'audio',
    },
    ...getSlideAnimations(slide),
  ])
  slide.bodySuffix = markNarration(slide.bodySuffix, spid)
  return added
}

/** The deck's presProps part, created and registered when missing. */
function presPropsPath(opened: OpenedPptx): string | null {
  const { archive } = opened
  for (const rel of archive.readRels(PRES_PATH).values()) {
    if (rel.type === PRES_PROPS_REL) return resolveTarget(PRES_PATH, rel.target)
  }
  if (!archive.has(PRES_PATH)) return null
  const path = 'ppt/presProps.xml'
  archive.entries.set(
    path,
    Buffer.from(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
        `<p:presentationPr xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"/>`,
      'utf8',
    ),
  )
  addContentTypeOverride(archive, path, PRES_PROPS_CT)
  appendRelationship(archive, PRES_PATH, PRES_PROPS_REL, 'presProps.xml')
  return path
}

const SHOW_PR_FOLLOWERS = /<p:(?:clrMru|extLst)\b|<\/p:presentationPr>/

/**
 * Sets presProps' `showNarration` (its schema default is off, so PowerPoint would not play the
 * recorded narration) and `useTimings`, keeping every other show property.
 */
export function setShowNarration(opened: OpenedPptx): void {
  const path = presPropsPath(opened)
  const xml = path && opened.archive.readText(path)
  if (!path || !xml) return
  const flags = { showNarration: '1', useTimings: '1' }
  let next: string
  const showPr = /<p:showPr\b[^>]*?(\/?)>/.exec(xml)
  if (showPr) {
    let tag = showPr[0]
    for (const [name, value] of Object.entries(flags)) {
      tag = new RegExp(`\\s${name}="[^"]*"`).test(tag)
        ? tag.replace(new RegExp(`\\s${name}="[^"]*"`), ` ${name}="${value}"`)
        : tag.replace(/^<p:showPr/, `<p:showPr ${name}="${value}"`)
    }
    next = xml.replace(showPr[0], tag)
  } else {
    const element = `<p:showPr showNarration="1" useTimings="1"/>`
    const selfClosed = xml.replace(/<p:presentationPr\b([^>]*?)\/>/, (_m, attrs: string) => {
      return `<p:presentationPr${attrs}>${element}</p:presentationPr>`
    })
    next =
      selfClosed !== xml
        ? selfClosed
        : xml.replace(SHOW_PR_FOLLOWERS, (follower) => `${element}${follower}`)
  }
  if (next !== xml) opened.archive.entries.set(path, Buffer.from(next, 'utf8'))
}

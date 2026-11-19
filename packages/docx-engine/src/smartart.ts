/**
 * SmartArt for Word — a real diagram, not a picture and not a shape group.
 *
 * Word keeps an editable diagram in five parts:
 *   word/diagrams/data1.xml       the data model (the tree)
 *   word/diagrams/layout1.xml     the layout definition (the algorithm)
 *   word/diagrams/quickStyle1.xml the style labels
 *   word/diagrams/colors1.xml     the colour variants
 *   word/diagrams/drawing1.xml    the pre-rendered shapes Word shows first
 *
 * document.xml points at the first four through <dgm:relIds>; the drawing part
 * is reached through the host part's relationship, which is what Office writes
 * (the relId in data1.xml's <dsp:dataModelExt> names a relationship of
 * document.xml, not one of its own). Word recomputes the shapes from the
 * layout whenever the user moves a node, so the result stays a diagram the
 * reader can edit rather than a frozen picture.
 *
 * The layout is Word's built-in "Organization Chart" definition
 * (urn:microsoft.com/office/officeart/2005/8/layout/hierarchy1), copied
 * verbatim from a document Word wrote: Word silently replaces a layout
 * definition it cannot match with "Basic Block List", while the definition it
 * wrote itself is re-flowed on every edit. buildDataXml follows the shape Word
 * writes around that layout — the doc point is empty, the title is a regular
 * node point under it, every layout node gets a presentation point, and the
 * connectors are presented from the parent transition point.
 *
 * The pre-rendered drawing is laid out by the org-chart interpreter in
 * @genoffice/pptx-engine/dgm-hier from the same constraint facts the layout
 * carries, so the first picture and Word's re-layout share their geometry.
 */

import {
  layoutHierTree,
  parseHierConstraints,
  type HierTreeNode,
} from '@genoffice/pptx-engine/dgm-hier'

import { HIERARCHY1_LAYOUT_XML } from './smartart-layout'

/** A node of the tree to draw. */
export interface SmartArtNode {
  label: string
  children?: SmartArtNode[]
}

export interface SmartArtOptions {
  /** frame width in px; defaults to the widest a page text column holds */
  widthPx?: number
  /** frame height in px; defaults to the drawing's own aspect */
  heightPx?: number
  /** the widest the frame may get, in px; a larger diagram scales down to fit.
   *  A reader clips a diagram that is wider than the page text column, so the
   *  default (600 px, 6.25 in) fits both A4 and Letter at 1 in margins. */
  maxSizePx?: number
}

export interface SmartArtParts {
  dataXml: string
  layoutXml: string
  quickStyleXml: string
  colorsXml: string
  drawingXml: string
  /** the <w:p> to place in document.xml */
  paragraphXml: string
  widthPx: number
  heightPx: number
}

// ---- constants ----

/** Word's built-in Organization Chart layout: the definition a re-layout uses. */
export const SMARTART_LAYOUT_ID = 'urn:microsoft.com/office/officeart/2005/8/layout/hierarchy1'
export const SMARTART_QUICK_STYLE_ID = 'urn:genoffice/officeart/quickstyle/smartart'
export const SMARTART_COLORS_ID = 'urn:genoffice/officeart/colors/smartart'

const DGM_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const DSP_NS = 'http://schemas.microsoft.com/office/drawing/2008/diagram'

/** EMU per px at 96 dpi. */
const EMU_PER_PX = 9525

/** The default page limit: 600 px is 6.25 in, just inside the 6.27 in A4
 *  and 6.5 in Letter text columns at 1 in margins (700 px does not fit A4). */
const DEFAULT_MAX_SIZE_PX = 600

/** The frame width before the page limit applies. */
const DEFAULT_WIDTH_PX = 800

/** hierarchy1 recurses through four levels and then repeats level four. */
const LAST_LEVEL = 4

/** Relationship types of the five diagram parts. */
export const SMARTART_REL_TYPES = {
  data: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData',
  layout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout',
  quickStyle:
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle',
  colors: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors',
  drawing: 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing',
} as const

/** Content types of the five parts. */
export const SMARTART_CONTENT_TYPES = {
  data: 'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml',
  layout: 'application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml',
  quickStyle: 'application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml',
  colors: 'application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml',
  drawing: 'application/vnd.ms-office.drawingml.diagramDrawing+xml',
} as const

// ---- small helpers ----

function escapeXmlText(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

/**
 * A deterministic GUID-shaped id. Word accepts any unique string, and a stable
 * one keeps two saves of the same tree byte-identical — the round-trip tests
 * rely on it.
 */
function modelId(seed: string): string {
  // FNV-1a over the seed, expanded into four GUID groups.
  let h1 = 0x811c9dc5
  let h2 = 0x01000193
  for (let i = 0; i < seed.length; i++) {
    h1 = Math.imul(h1 ^ seed.charCodeAt(i), 0x01000193) >>> 0
    h2 = Math.imul(h2 + seed.charCodeAt(i) + i, 0x85ebca6b) >>> 0
  }
  const hex = (n: number, width: number): string =>
    (n >>> 0).toString(16).padStart(width, '0').slice(-width)
  return `{${hex(h1, 8)}-${hex(h2, 8).slice(0, 4)}-${hex(h1, 8).slice(0, 4)}-${hex(h2, 8).slice(0, 4)}-${hex(
    h1 ^ h2,
    12,
  )}}`.toUpperCase()
}

/** The run language a label declares: CJK-bearing labels say so, so consumers
 *  pick an east-asian face for them instead of guessing from en-US. */
const CJK_RE = /[\u2e80-\u9fff\uf900-\ufaff\uac00-\ud7af\uff00-\uffef]/

function runLang(label: string): string {
  return CJK_RE.test(label) ? 'zh-CN' : 'en-US'
}

/** Flatten the tree into the order the parts are written in. */
interface FlatNode {
  label: string
  /** Word's depth: the doc point is above the root, so the root node is 1. */
  depth: number
  children: FlatNode[]
  /** index among its siblings */
  ord: number
}

function flatten(root: SmartArtNode): FlatNode {
  const walk = (node: SmartArtNode, depth: number, ord: number): FlatNode => ({
    label: node.label,
    depth,
    ord,
    children: (node.children ?? []).map((kid, i) => walk(kid, depth + 1, i)),
  })
  return walk(root, 1, 0)
}

function countNodes(node: FlatNode): number {
  return 1 + node.children.reduce((sum, kid) => sum + countNodes(kid), 0)
}

function textBody(label: string): string {
  const lang = runLang(label)
  return (
    '<dgm:t><a:bodyPr/><a:lstStyle/><a:p>' +
    (label ? `<a:r><a:rPr lang="${lang}"/><a:t>${escapeXmlText(label)}</a:t></a:r>` : '') +
    `<a:endParaRPr lang="${lang}"/>` +
    '</a:p></dgm:t>'
  )
}

function emptyText(): string {
  return '<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t>'
}

// ---- the data model ----

interface Ids {
  /** data point of the node */
  data: string
  /** the parent-child connection into this node */
  cxn: string
  /** the parent transition point of that connection */
  parTrans: string
  /** the sibling transition point of that connection */
  sibTrans: string
  /** presentation point of the hierRoot layout node (the node's presentation root) */
  hierRoot: string
  /** presentation point of the composite layout node (box + text) */
  composite: string
  /** presentation point of the background layout node (the box) */
  background: string
  /** presentation point of the text layout node (the label) */
  text: string
  /** presentation point of the hierChild layout node (the row of children) */
  hierChild: string
  /** presentation point of the connector layout node (the line into this node) */
  connectorPres: string
}

function buildIds(root: FlatNode): Map<FlatNode, Ids> {
  const ids = new Map<FlatNode, Ids>()
  let n = 0
  const walk = (node: FlatNode): void => {
    const key = String(n++)
    ids.set(node, {
      data: modelId(`data:${key}:${node.label}`),
      cxn: modelId(`cxn:${key}`),
      parTrans: modelId(`par:${key}`),
      sibTrans: modelId(`sib:${key}`),
      hierRoot: modelId(`hierRoot:${key}`),
      composite: modelId(`composite:${key}`),
      background: modelId(`background:${key}`),
      text: modelId(`text:${key}`),
      hierChild: modelId(`hierChild:${key}`),
      connectorPres: modelId(`connector:${key}`),
    })
    for (const kid of node.children) walk(kid)
  }
  walk(root)
  return ids
}

// ---- layout node names ----

/** hierarchy1 names its layout nodes per level: suffix 1..4, level four repeats. */
function levelOf(depth: number): string {
  return String(Math.min(depth, LAST_LEVEL))
}

function hierRootName(depth: number): string {
  return `hierRoot${levelOf(depth)}`
}

function compositeName(depth: number): string {
  return depth === 1 ? 'composite' : `composite${levelOf(depth)}`
}

function backgroundName(depth: number): string {
  return depth === 1 ? 'background' : `background${levelOf(depth)}`
}

function textName(depth: number): string {
  return depth === 1 ? 'text' : `text${levelOf(depth)}`
}

function hierChildName(depth: number): string {
  return `hierChild${Math.min(depth + 1, LAST_LEVEL + 1)}`
}

/** The connector layout node depends on the parent level: the first two levels
 *  have their own node, everything deeper repeats the level-three one. */
function connectorName(parentDepth: number): string {
  return parentDepth === 1 ? 'Name10' : parentDepth === 2 ? 'Name17' : 'Name23'
}

/** Style labels the quick style below defines, by node level. Word keys the
 *  labels by the depth of the node in the tree (level one is label 0). */
function levelLabel(prefix: string, depth: number): string {
  return `${prefix}${depth === 1 ? 0 : Math.min(depth, LAST_LEVEL)}`
}

function backgroundStyle(depth: number): string {
  return ` presStyleLbl="${levelLabel('node', depth)}" presStyleIdx="0" presStyleCnt="1"`
}

function textStyle(depth: number): string {
  return ` presStyleLbl="${levelLabel('fgAcc', depth)}" presStyleIdx="0" presStyleCnt="1"`
}

function connectorStyle(childDepth: number): string {
  // Word writes one label per child level; the layout's repeated level-four
  // connector (Name23) keeps that label for every deeper level.
  const label = `parChTrans1D${Math.min(childDepth, LAST_LEVEL)}`
  return ` presStyleLbl="${label}" presStyleIdx="0" presStyleCnt="2"`
}

/**
 * The data model.
 *
 * Word resolves a diagram by walking the layout definition over the data
 * points, so the data model has to have the shape the layout expects: an empty
 * doc point (no layout presents the doc point's text), the title in a regular
 * node point connected from it, a parent/sibling transition pair per
 * connection, and a presentation point per layout node. The presentation
 * points are named exactly after the layout nodes they stand for and are
 * wired into the same parent-child tree; a presentation point that does not
 * match a shape-owning layout node makes Word drop the layout.
 */
function buildDataXml(root: FlatNode, ids: Map<FlatNode, Ids>, drawingRelId: string): string {
  const docId = modelId('doc')
  const rootIds = ids.get(root)!
  const pts: string[] = []
  const cxns: string[] = []
  const presOf: string[] = []
  const presParOf: string[] = []

  // the doc point: an empty container. The title lives in the root node below.
  pts.push(
    `<dgm:pt modelId="${docId}" type="doc">` +
      `<dgm:prSet loTypeId="${SMARTART_LAYOUT_ID}" loCatId="hierarchy" ` +
      `qsTypeId="${SMARTART_QUICK_STYLE_ID}" qsCatId="simple" ` +
      `csTypeId="${SMARTART_COLORS_ID}" csCatId="accent1" phldr="1"/>` +
      '<dgm:spPr/>' +
      emptyText() +
      '</dgm:pt>',
  )

  const walkData = (node: FlatNode): void => {
    const nodeIds = ids.get(node)!
    pts.push(
      `<dgm:pt modelId="${nodeIds.data}"><dgm:prSet/><dgm:spPr/>${textBody(node.label)}</dgm:pt>`,
    )
    pts.push(
      `<dgm:pt modelId="${nodeIds.parTrans}" type="parTrans" cxnId="${nodeIds.cxn}">` +
        `<dgm:prSet/><dgm:spPr/>${emptyText()}</dgm:pt>`,
    )
    pts.push(
      `<dgm:pt modelId="${nodeIds.sibTrans}" type="sibTrans" cxnId="${nodeIds.cxn}">` +
        `<dgm:prSet/><dgm:spPr/>${emptyText()}</dgm:pt>`,
    )
    for (const kid of node.children) walkData(kid)
  }
  walkData(root)

  // the parent-child connections, each carrying its transition pair
  const connect = (parent: FlatNode): void => {
    const parentIds = ids.get(parent)!
    parent.children.forEach((kid, i) => {
      const kidIds = ids.get(kid)!
      cxns.push(
        `<dgm:cxn modelId="${kidIds.cxn}" srcId="${parentIds.data}" destId="${kidIds.data}" ` +
          `srcOrd="${i}" destOrd="0" parTransId="${kidIds.parTrans}" sibTransId="${kidIds.sibTrans}"/>`,
      )
      connect(kid)
    })
  }
  connect(root)
  // document -> root
  cxns.push(
    `<dgm:cxn modelId="${rootIds.cxn}" srcId="${docId}" destId="${rootIds.data}" ` +
      `srcOrd="0" destOrd="0" parTransId="${rootIds.parTrans}" sibTransId="${rootIds.sibTrans}"/>`,
  )

  const presPoints: string[] = []
  const emitPres = (
    assocId: string,
    presId: string,
    name: string,
    attrs: string,
    inner: string,
    body: string,
    presentsData = false,
  ): void => {
    presPoints.push(
      `<dgm:pt modelId="${presId}" type="pres">` +
        `<dgm:prSet presAssocID="${assocId}" presName="${name}"${attrs}>${inner}</dgm:prSet>` +
        body +
        '</dgm:pt>',
    )
    // Only the point that presents an object's content carries the presOf
    // link — the text point of a node, the connector of a connection, the
    // document's root point — while the structural points (composite,
    // background, hierChild) hang off it through presParOf.
    if (presentsData) {
      presOf.push(
        `<dgm:cxn modelId="${modelId(`presOf:${presId}`)}" type="presOf" ` +
          `srcId="${assocId}" destId="${presId}" srcOrd="0" destOrd="0" presId="${SMARTART_LAYOUT_ID}"/>`,
      )
    }
  }
  const linkPres = (parentPres: string, childPres: string, srcOrd: number): void => {
    presParOf.push(
      `<dgm:cxn modelId="${modelId(`presPar:${parentPres}:${childPres}`)}" type="presParOf" ` +
        `srcId="${parentPres}" destId="${childPres}" srcOrd="${srcOrd}" destOrd="0" presId="${SMARTART_LAYOUT_ID}"/>`,
    )
  }

  // The document's presentation point is the layout's root node (hierChild1).
  const docPres = modelId('pres:doc')
  emitPres(
    docId,
    docPres,
    'hierChild1',
    ' presStyleCnt="0"',
    '<dgm:presLayoutVars><dgm:chPref val="1"/><dgm:dir/><dgm:animOne val="branch"/>' +
      '<dgm:animLvl val="lvl"/><dgm:resizeHandles/></dgm:presLayoutVars>',
    '<dgm:spPr/>' + emptyText(),
    true,
  )
  linkPres(docPres, rootIds.hierRoot, 0)

  const walkPres = (node: FlatNode): void => {
    const nodeIds = ids.get(node)!
    // hierRoot presents the node and owns the subtree's layout; the point that
    // carries the node's content is the text one, which takes the presOf link
    emitPres(
      nodeIds.data,
      nodeIds.hierRoot,
      hierRootName(node.depth),
      ' presStyleCnt="0"',
      '',
      '<dgm:spPr/>',
    )
    // composite holds the box: a background shape and a text shape
    emitPres(
      nodeIds.data,
      nodeIds.composite,
      compositeName(node.depth),
      ' presStyleCnt="0"',
      '',
      '<dgm:spPr/>',
    )
    emitPres(
      nodeIds.data,
      nodeIds.background,
      backgroundName(node.depth),
      backgroundStyle(node.depth) || ' presStyleCnt="0"',
      '',
      '<dgm:spPr/>',
    )
    // the node's data point is presented through its text point, as in Word's
    // own hierarchy1 data models
    emitPres(
      nodeIds.data,
      nodeIds.text,
      textName(node.depth),
      textStyle(node.depth) || ' presStyleCnt="0"',
      '<dgm:presLayoutVars><dgm:chPref val="3"/></dgm:presLayoutVars>',
      '<dgm:spPr/>' + emptyText(),
      true,
    )
    // hierChild lays out the row of children under this node
    emitPres(
      nodeIds.data,
      nodeIds.hierChild,
      hierChildName(node.depth),
      ' presStyleCnt="0"',
      '',
      '<dgm:spPr/>',
    )
    linkPres(nodeIds.hierRoot, nodeIds.composite, 0)
    linkPres(nodeIds.hierRoot, nodeIds.hierChild, 1)
    linkPres(nodeIds.composite, nodeIds.background, 0)
    linkPres(nodeIds.composite, nodeIds.text, 1)
    node.children.forEach((kid, i) => {
      const kidIds = ids.get(kid)!
      // A hierarchy connector is presented from the parent transition point of
      // the connection it draws — not from the sibling transition.
      emitPres(
        kidIds.parTrans,
        kidIds.connectorPres,
        connectorName(node.depth),
        connectorStyle(kid.depth) || ' presStyleCnt="0"',
        '',
        '<dgm:spPr/>' + emptyText(),
        true,
      )
      walkPres(kid)
      linkPres(nodeIds.hierChild, kidIds.connectorPres, i * 2)
      linkPres(nodeIds.hierChild, kidIds.hierRoot, i * 2 + 1)
    })
  }
  walkPres(root)

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<dgm:dataModel xmlns:dgm="${DGM_NS}" xmlns:a="${A_NS}">` +
    `<dgm:ptLst>${[...pts, ...presPoints].join('')}</dgm:ptLst>` +
    `<dgm:cxnLst>${[...cxns, ...presOf, ...presParOf].join('')}</dgm:cxnLst>` +
    '<dgm:bg/><dgm:whole/>' +
    // the pre-rendered drawing, reached through the host part's relationship
    '<dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram">' +
    `<dsp:dataModelExt xmlns:dsp="${DSP_NS}" relId="${drawingRelId}" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>` +
    '</a:ext></dgm:extLst>' +
    '</dgm:dataModel>'
  )
}

// ---- the quick style and the colours ----

/** The style labels the layout definition references: the box (node0), the
 *  text (fgAcc0, and one per deeper level) and the connector lines. */
function styleLabels(): Array<{ name: string; ln: number; fill: number; font: string }> {
  const textLevels = Array.from({ length: LAST_LEVEL + 1 }, (_, i) => ({
    name: `fgAcc${i}`,
    ln: 0,
    fill: 0,
    font: 'lt1',
  }))
  const boxLevels = Array.from({ length: LAST_LEVEL + 1 }, (_, i) => ({
    name: `node${i}`,
    ln: 2,
    fill: 1,
    font: 'lt1',
  }))
  return [
    ...boxLevels,
    ...textLevels,
    { name: 'parChTrans1D2', ln: 1, fill: 0, font: 'tx1' },
    { name: 'parChTrans1D3', ln: 1, fill: 0, font: 'tx1' },
    { name: 'parChTrans1D4', ln: 1, fill: 0, font: 'tx1' },
  ]
}

function buildQuickStyleXml(): string {
  const labels = styleLabels()
    .map(
      ({ name, ln, fill, font }) =>
        `<dgm:styleLbl name="${name}">` +
        '<dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d>' +
        '<dgm:sp3d/><dgm:txPr/>' +
        '<dgm:style>' +
        `<a:lnRef idx="${ln}"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef>` +
        `<a:fillRef idx="${fill}"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef>` +
        '<a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef>' +
        `<a:fontRef idx="minor"><a:schemeClr val="${font}"/></a:fontRef>` +
        '</dgm:style>' +
        '</dgm:styleLbl>',
    )
    .join('')
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<dgm:styleDef xmlns:dgm="${DGM_NS}" xmlns:a="${A_NS}" uniqueId="${SMARTART_QUICK_STYLE_ID}">` +
    '<dgm:title val="SmartArt"/><dgm:desc val=""/>' +
    '<dgm:catLst><dgm:cat type="simple" pri="10100"/></dgm:catLst>' +
    '<dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d>' +
    labels +
    '</dgm:styleDef>'
  )
}

function buildColorsXml(): string {
  const fill = (val: string): string =>
    val
      ? `<dgm:fillClrLst meth="repeat"><a:schemeClr val="${val}"/></dgm:fillClrLst>`
      : '<dgm:fillClrLst/>'
  const line = (val: string): string =>
    val
      ? `<dgm:linClrLst meth="repeat"><a:schemeClr val="${val}"/></dgm:linClrLst>`
      : '<dgm:linClrLst/>'
  const text = (val: string): string =>
    val
      ? `<dgm:txFillClrLst meth="repeat"><a:schemeClr val="${val}"/></dgm:txFillClrLst>`
      : '<dgm:txFillClrLst/>'
  const label = (name: string, fillClr: string, ln: string, tx: string): string =>
    `<dgm:styleLbl name="${name}">` +
    fill(fillClr) +
    line(ln) +
    '<dgm:effectClrLst/><dgm:txLinClrLst/>' +
    text(tx) +
    '<dgm:txEffectClrLst/>' +
    '</dgm:styleLbl>'
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<dgm:colorsDef xmlns:dgm="${DGM_NS}" xmlns:a="${A_NS}" uniqueId="${SMARTART_COLORS_ID}">` +
    '<dgm:title val="SmartArt"/><dgm:desc val=""/>' +
    '<dgm:catLst><dgm:cat type="accent1" pri="11100"/></dgm:catLst>' +
    styleLabels()
      .filter((entry) => entry.name.startsWith('node'))
      .map((entry) => label(entry.name, 'accent1', 'lt1', 'lt1'))
      .join('') +
    styleLabels()
      .filter((entry) => entry.name.startsWith('fgAcc'))
      .map((entry) => label(entry.name, '', '', 'lt1'))
      .join('') +
    styleLabels()
      .filter((entry) => entry.name.startsWith('parChTrans'))
      .map((entry) => label(entry.name, '', 'accent1', ''))
      .join('') +
    '</dgm:colorsDef>'
  )
}

// ---- the pre-rendered drawing ----

/** The boxes and connectors of a layout, in px. */
interface DrawingGeometry {
  boxes: Array<{
    label: string
    depth: number
    x: number
    y: number
    w: number
    h: number
    /** presentation point of the box shape, owned by the background node */
    backgroundPresId: string
    /** presentation point of the label shape, owned by the text node */
    textPresId: string
  }>
  lines: Array<{
    x1: number
    y1: number
    x2: number
    y2: number
    /** presentation point of the connection this segment draws */
    presId: string
  }>
}

/** The base font of a label, in px, by depth. */
function labelFontPx(depth: number): number {
  return depth === 1 ? 13 : depth === 2 ? 11.5 : 10.5
}

/** The width one label character renders at, in ems of the font size: CJK and
 *  the fullwidth forms run the whole em, latin text about half. */
function textUnitsOf(label: string): number {
  let units = 0
  for (const ch of label) {
    units += ch.charCodeAt(0) > 0x2e7f ? 1 : 0.55
  }
  return units
}

/** The rendered font size of a label, in DrawingML hundredths of a point: the
 *  depth's base size, shrunk when the box it ended up in got too narrow. */
function labelFontSize(label: string, boxW: number, depth: number): number {
  const units = textUnitsOf(label)
  const cap = units > 0 ? (boxW - 22) / units : labelFontPx(depth)
  const px = Math.max(7, Math.min(labelFontPx(depth), cap))
  return Math.round(px * 75)
}

/**
 * The pre-rendered shapes.
 *
 * Word shows these before it recomputes anything, so the diagram is correct
 * the moment the document opens. The coordinates come from the same
 * interpreter that reads the layout's constraint list, so the picture matches
 * the geometry Word produces when it re-lays the diagram out.
 *
 * Every shape is keyed the way Office keys its own cached drawings: the box to
 * the background point, the label to the text point, and each connector
 * segment to the connector point of the connection it draws. hierRoot,
 * composite and hierChild own no shape, so they never appear as a modelId.
 */
function buildDrawingXml(geo: DrawingGeometry): string {
  const sps: string[] = []

  const nodeColor = (depth: number): string =>
    depth === 1 ? 'E8663F' : depth === 2 ? '4472C4' : '5B9BD5'
  const textColor = 'FFFFFF'

  /** the shape transform every shape of a box shares */
  const transform = (x: number, y: number, w: number, h: number): string =>
    `<a:off x="${Math.round(x * EMU_PER_PX)}" y="${Math.round(y * EMU_PER_PX)}"/>` +
    `<a:ext cx="${Math.round(w * EMU_PER_PX)}" cy="${Math.round(h * EMU_PER_PX)}"/>`

  for (const box of geo.boxes) {
    const { label, depth, x, y, w, h, backgroundPresId, textPresId } = box
    // the text follows the box: a scaled diagram keeps its proportions
    const fontSize = labelFontSize(label, w, depth)
    // Office keys every shape to the presentation point that owns it: the box
    // to the background node, the label to the text node (hierRoot and the
    // composite own no shape).
    sps.push(
      `<dsp:sp modelId="${backgroundPresId}">` +
        '<dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>' +
        '<dsp:spPr>' +
        `<a:xfrm>${transform(x, y, w, h)}</a:xfrm>` +
        '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>' +
        `<a:solidFill><a:srgbClr val="${nodeColor(depth)}"/></a:solidFill>` +
        '<a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>' +
        '</dsp:spPr>' +
        '<dsp:style>' +
        '<a:lnRef idx="2"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef>' +
        '<a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef>' +
        '<a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef>' +
        '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>' +
        '</dsp:style>' +
        '</dsp:sp>',
    )
    sps.push(
      `<dsp:sp modelId="${textPresId}">` +
        '<dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>' +
        '<dsp:spPr>' +
        `<a:xfrm>${transform(x, y, w, h)}</a:xfrm>` +
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
        '<a:noFill/><a:ln><a:noFill/></a:ln>' +
        '</dsp:spPr>' +
        '<dsp:style>' +
        '<a:lnRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef>' +
        '<a:fillRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef>' +
        '<a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef>' +
        '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>' +
        '</dsp:style>' +
        '<dsp:txBody>' +
        '<a:bodyPr spcFirstLastPara="0" vert="horz" wrap="square" lIns="45720" tIns="45720" ' +
        'rIns="45720" bIns="45720" numCol="1" spcCol="1270" anchor="ctr" anchorCtr="0">' +
        '<a:noAutofit/></a:bodyPr>' +
        '<a:lstStyle/>' +
        '<a:p><a:pPr lvl="0" algn="ctr" defTabSz="1111250"/>' +
        `<a:r><a:rPr lang="${runLang(label)}" sz="${fontSize}" b="1" kern="1200"><a:solidFill><a:srgbClr val="${textColor}"/></a:solidFill></a:rPr>` +
        `<a:t>${escapeXmlText(label)}</a:t></a:r>` +
        '</a:p>' +
        '</dsp:txBody>' +
        `<dsp:txXfrm>${transform(x, y, w, h)}</dsp:txXfrm>` +
        '</dsp:sp>',
    )
  }

  // the connectors: the interpreter's segments, each keyed to the connection's
  // presentation point (the connector node owns the shape, not hierRoot)
  for (const line of geo.lines) {
    const x = Math.min(line.x1, line.x2)
    const y = Math.min(line.y1, line.y2)
    const cx = Math.abs(line.x2 - line.x1)
    const cy = Math.abs(line.y2 - line.y1)
    if (cx < 1 && cy < 1) continue
    // the line preset draws the top left to the bottom right diagonal; a
    // segment that runs the other diagonal needs the vertical flip
    const flipV = line.x2 < line.x1 !== line.y2 < line.y1
    sps.push(
      `<dsp:sp modelId="${line.presId}">` +
        '<dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>' +
        '<dsp:spPr>' +
        `<a:xfrm${flipV ? ' flipV="1"' : ''}>` +
        `<a:off x="${Math.round(x * EMU_PER_PX)}" y="${Math.round(y * EMU_PER_PX)}"/>` +
        `<a:ext cx="${Math.round(cx * EMU_PER_PX)}" cy="${Math.round(cy * EMU_PER_PX)}"/></a:xfrm>` +
        '<a:prstGeom prst="line"><a:avLst/></a:prstGeom>' +
        '<a:noFill/>' +
        '<a:ln w="12700"><a:solidFill><a:srgbClr val="9AA4B2"/></a:solidFill></a:ln>' +
        '</dsp:spPr>' +
        '<dsp:style>' +
        '<a:lnRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef>' +
        '<a:fillRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef>' +
        '<a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef>' +
        '<a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef>' +
        '</dsp:style>' +
        '<dsp:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></dsp:txBody>' +
        '</dsp:sp>',
    )
  }

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<dsp:drawing xmlns:dsp="${DSP_NS}" xmlns:a="${A_NS}">` +
    `<dsp:spTree>` +
    '<dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr>' +
    '<dsp:grpSpPr/>' +
    sps.join('') +
    '</dsp:spTree>' +
    '</dsp:drawing>'
  )
}

// ---- the document paragraph ----

/**
 * The paragraph Word puts in the body: an inline drawing that points at the
 * four diagram parts (the cached drawing hangs off the host part as well).
 */
function buildParagraphXml(
  relIds: { data: string; layout: string; quickStyle: string; colors: string; drawing: string },
  widthPx: number,
  heightPx: number,
  docPrId: number,
): string {
  const cx = Math.round(widthPx * EMU_PER_PX)
  const cy = Math.round(heightPx * EMU_PER_PX)
  return (
    '<w:p><w:r><w:drawing>' +
    '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    `<wp:docPr id="${docPrId}" name="SmartArt ${docPrId}"/>` +
    '<wp:cNvGraphicFramePr/>' +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">' +
    `<dgm:relIds xmlns:dgm="${DGM_NS}" xmlns:r="${R_NS}" ` +
    `r:dm="${relIds.data}" r:lo="${relIds.layout}" r:qs="${relIds.quickStyle}" r:cs="${relIds.colors}"/>` +
    '</a:graphicData></a:graphic>' +
    '</wp:inline>' +
    '</w:drawing></w:r></w:p>'
  )
}

// ---- the entry point ----

/** Turn the tree into the shape the org-chart interpreter lays out. */
function toHierTree(node: FlatNode, nodes: Map<HierTreeNode, FlatNode>): HierTreeNode {
  const hier: HierTreeNode = {
    texts: node.label ? [node.label] : [],
    children: node.children.map((kid) => toHierTree(kid, nodes)),
    // Word's org chart never hangs deep leaf branches; that is a PowerPoint
    // variant, so pin every node to the standard hierarchy algorithm.
    hierBranch: 'std',
  }
  nodes.set(hier, node)
  return hier
}

/**
 * Build every part a Word diagram needs.
 *
 * `relIds` are the relationship ids the caller puts in document.xml; the
 * drawing is reached through the host part's relationship (relIds.drawing),
 * so it needs no id of its own.
 */
export function buildSmartArtParts(
  root: SmartArtNode,
  options: SmartArtOptions = {},
  relIds: { data: string; layout: string; quickStyle: string; colors: string; drawing: string } = {
    data: 'rId1',
    layout: 'rId2',
    quickStyle: 'rId3',
    colors: 'rId4',
    drawing: 'rId5',
  },
  docPrId = 1,
): SmartArtParts {
  const flat = flatten(root)
  const ids = buildIds(flat)
  const cons = parseHierConstraints(HIERARCHY1_LAYOUT_XML)

  const nodes = new Map<HierTreeNode, FlatNode>()
  const hierRoot = toHierTree(flat, nodes)

  // A probe pass measures the tree, a second pass fills the frame with it.
  // The interpreter scales the tree uniformly into the frame, so the probe
  // yields the width-to-height ratio rather than a size.
  const probe = layoutHierTree([hierRoot], cons, 1000, 1000)
  const probed = probe?.boxes ?? []
  const spanW = probed.length
    ? Math.max(...probed.map((b) => b.x + b.w)) - Math.min(...probed.map((b) => b.x))
    : 0
  const spanH = probed.length
    ? Math.max(...probed.map((b) => b.y + b.h)) - Math.min(...probed.map((b) => b.y))
    : 0
  const ratio = spanH > 0 ? spanW / spanH : 1

  // The frame: default to 600 px (inside both the A4 and Letter text columns)
  // and scale the whole diagram down so neither side outgrows the page limit.
  const limit = options.maxSizePx ?? DEFAULT_MAX_SIZE_PX
  let width = options.widthPx ?? Math.min(DEFAULT_WIDTH_PX, limit)
  let height = options.heightPx ?? Math.round(width / ratio)
  const overflow = Math.max(width / limit, height / limit, 1)
  if (overflow > 1) {
    width /= overflow
    height /= overflow
  }
  width = Math.max(1, Math.round(width))
  height = Math.max(1, Math.round(height))

  const laid = layoutHierTree([hierRoot], cons, width, height)
  const geo: DrawingGeometry = {
    boxes: (laid?.boxes ?? []).flatMap((b) => {
      const node = nodes.get(b.node)
      const nodeIds = node ? ids.get(node) : undefined
      if (!node || !nodeIds) return []
      return [
        {
          label: b.node.texts.join(' '),
          depth: node.depth,
          x: b.x,
          y: b.y,
          w: b.w,
          h: b.h,
          backgroundPresId: nodeIds.background,
          textPresId: nodeIds.text,
        },
      ]
    }),
    lines: (laid?.lines ?? []).flatMap((l) => {
      const node = l.to ? nodes.get(l.to) : undefined
      const nodeIds = node ? ids.get(node) : undefined
      if (!nodeIds) return []
      return [
        {
          x1: l.x,
          y1: l.y,
          x2: l.x + l.cx,
          y2: l.y + l.cy,
          presId: nodeIds.connectorPres,
        },
      ]
    }),
  }

  return {
    dataXml: buildDataXml(flat, ids, relIds.drawing),
    layoutXml: HIERARCHY1_LAYOUT_XML,
    quickStyleXml: buildQuickStyleXml(),
    colorsXml: buildColorsXml(),
    drawingXml: buildDrawingXml(geo),
    paragraphXml: buildParagraphXml(relIds, width, height, docPrId),
    widthPx: width,
    heightPx: height,
  }
}

/** The number of nodes a tree holds — the shape count the drawing gets. */
export function smartArtNodeCount(root: SmartArtNode): number {
  return countNodes(flatten(root))
}

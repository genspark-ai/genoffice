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
 * document.xml points at the first four through <dgm:relIds>; the drawing is
 * reached from data1.xml through its own rels. Word recomputes the shapes
 * from the layout whenever the user moves a node, so the result stays a
 * diagram the reader can edit rather than a frozen picture.
 *
 * The layout definition here is written in the same constraint dialect the
 * org-chart interpreter in @genoffice/pptx-engine reads (see dgm-hier.ts), so
 * the pre-rendered drawing and the layout Word applies agree on the geometry.
 */

import {
  layoutHierTree,
  type HierConstraints,
  type HierTreeNode,
} from '@genoffice/pptx-engine/dgm-hier'

/** A node of the tree to draw. */
export interface SmartArtNode {
  label: string
  children?: SmartArtNode[]
}

/** How the tree is arranged. */
export type SmartArtOrientation = 'top-down' | 'left-right' | 'radial'

/**
 * The two layout definitions. `hierarchy` is the org-chart shape — the topic
 * on top, the branches below it. `radial` puts the topic in the middle and
 * the branches around it.
 */
export type SmartArtLayoutKind = 'hierarchy' | 'radial'

export interface SmartArtOptions {
  /** defaults to `left-right` — the root in the middle, branches to both sides */
  orientation?: SmartArtOrientation
  /** frame width in px; defaults to a width that fits the tree */
  widthPx?: number
  /** frame height in px; defaults to a height that fits the tree */
  heightPx?: number
  /** the widest the frame may get, in px; a larger map scales down to fit.
   *  A reader drops a diagram that is wider than the page, so the default is
   *  a little under the width of A4. */
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

/** The layout this module writes. Word shows it as "Mind Map". */
export const SMARTART_LAYOUT_ID = 'urn:genoffice/officeart/layout/mindMap'
export const SMARTART_RADIAL_LAYOUT_ID = 'urn:genoffice/officeart/layout/radialMindMap'
export const SMARTART_QUICK_STYLE_ID = 'urn:genoffice/officeart/quickstyle/mindMap'
export const SMARTART_COLORS_ID = 'urn:genoffice/officeart/colors/mindMap'

const DGM_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const DSP_NS = 'http://schemas.microsoft.com/office/drawing/2008/diagram'

/** the width a node box is drawn at before the map is scaled to fit */
const BASE_NODE_WIDTH = 150

/** EMU per px at 96 dpi. */
const EMU_PER_PX = 9525

/** Relationship types of the four diagram parts. */
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

/** Flatten the tree into the order the parts are written in. */
interface FlatNode {
  label: string
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
  return walk(root, 0, 0)
}

function countNodes(node: FlatNode): number {
  return 1 + node.children.reduce((sum, kid) => sum + countNodes(kid), 0)
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
  /** the presentation point of this node */
  pres: string
  /** the presentation point of the connector below this node */
  connectorPres: string
}

function buildIds(root: FlatNode): Map<FlatNode, Ids> {
  const ids = new Map<FlatNode, Ids>()
  let n = 0
  const walk = (node: FlatNode): void => {
    const key = String(n++)
    ids.set(node, {
      data: modelId(`data:${key}:${node.label}`),
      cxn: modelId(`cxn:${key}:${node.label}`),
      parTrans: modelId(`par:${key}:${node.label}`),
      sibTrans: modelId(`sib:${key}:${node.label}`),
      pres: modelId(`pres:${key}:${node.label}`),
      connectorPres: modelId(`conn:${key}:${node.label}`),
    })
    for (const kid of node.children) walk(kid)
  }
  walk(root)
  return ids
}

/** The presentation name of a node's shape, matching the layout's node names. */
function presNameOf(depth: number, kind: SmartArtLayoutKind): string {
  if (depth === 0) return 'diagram'
  if (kind === 'radial') return depth === 1 ? 'radialNode' : 'node'
  return depth === 1 ? 'rootComposite1' : 'hierRoot1'
}

/** The presentation name of a connector, matching the layout's node names. */
function connectorPresNameOf(kind: SmartArtLayoutKind): string {
  return kind === 'radial' ? 'radialConnector' : 'rootConnector1'
}

function textBody(label: string): string {
  return (
    '<dgm:t><a:bodyPr/><a:lstStyle/><a:p>' +
    (label ? `<a:r><a:rPr lang="en-US"/><a:t>${escapeXmlText(label)}</a:t></a:r>` : '') +
    '<a:endParaRPr lang="en-US"/>' +
    '</a:p></dgm:t>'
  )
}

/**
 * The data model: one point per tree node, a transition pair per connection,
 * and a presentation point per shape. Word rewrites the presentation tree
 * when the reader moves something, so it has to mirror the layout.
 */
function buildDataXml(
  root: FlatNode,
  ids: Map<FlatNode, Ids>,
  drawingRelId: string,
  kind: SmartArtLayoutKind,
): string {
  const rootIds = ids.get(root)!
  // the data model points at the layout it was written for
  const layoutId = kind === 'radial' ? SMARTART_RADIAL_LAYOUT_ID : SMARTART_LAYOUT_ID
  const pts: string[] = []

  // the document point carries the ids of the layout, style and colours
  pts.push(
    `<dgm:pt modelId="${rootIds.data}" type="doc">` +
      `<dgm:prSet loTypeId="${layoutId}" loCatId="hierarchy" ` +
      `qsTypeId="${SMARTART_QUICK_STYLE_ID}" qsCatId="simple" ` +
      `csTypeId="${SMARTART_COLORS_ID}" csCatId="accent1"/>` +
      '<dgm:spPr/>' +
      textBody(root.label) +
      '</dgm:pt>',
  )

  // data points, transition points and presentation points
  const walk = (node: FlatNode): void => {
    const nodeIds = ids.get(node)!
    if (node.depth > 0) {
      pts.push(
        `<dgm:pt modelId="${nodeIds.data}">` +
          '<dgm:prSet/><dgm:spPr/>' +
          textBody(node.label) +
          '</dgm:pt>',
      )
      // the transition pair belongs to the connection that brings the node in
      pts.push(
        `<dgm:pt modelId="${nodeIds.parTrans}" type="parTrans" cxnId="${nodeIds.cxn}">` +
          '<dgm:prSet/><dgm:spPr/>' +
          '<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t>' +
          '</dgm:pt>',
      )
      pts.push(
        `<dgm:pt modelId="${nodeIds.sibTrans}" type="sibTrans" cxnId="${nodeIds.cxn}">` +
          '<dgm:prSet/><dgm:spPr/>' +
          '<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t>' +
          '</dgm:pt>',
      )
    }
    if (node.depth > 0) {
      // the connector of the link into this node hangs off its sibling transition
      pts.push(
        `<dgm:pt modelId="${nodeIds.connectorPres}" type="pres">` +
          `<dgm:prSet presAssocID="${nodeIds.sibTrans}" presName="${connectorPresNameOf(kind)}" presStyleCnt="0"/>` +
          '<dgm:spPr/>' +
          '</dgm:pt>',
      )
    }
    // the presentation point of the node's shape
    pts.push(
      `<dgm:pt modelId="${nodeIds.pres}" type="pres">` +
        `<dgm:prSet presAssocID="${nodeIds.data}" presName="${presNameOf(node.depth, kind)}" ` +
        'presStyleLbl="node1" presStyleCnt="0"/>' +
        '<dgm:spPr/>' +
        '</dgm:pt>',
    )
    for (const kid of node.children) walk(kid)
  }
  for (const kid of root.children) walk(kid)
  // the root's own presentation point
  pts.push(
    `<dgm:pt modelId="${rootIds.pres}" type="pres">` +
      `<dgm:prSet presAssocID="${rootIds.data}" presName="diagram" presStyleCnt="0">` +
      '<dgm:presLayoutVars><dgm:dir/><dgm:animLvl val="lvl"/><dgm:resizeHandles val="exact"/></dgm:presLayoutVars>' +
      '</dgm:prSet>' +
      '<dgm:spPr/>' +
      '</dgm:pt>',
  )

  const cxns: string[] = []

  // the data connections, each carrying its transition pair
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

  // presentation mappings
  const mapPres = (node: FlatNode): void => {
    const nodeIds = ids.get(node)!
    cxns.push(
      `<dgm:cxn modelId="${modelId(`presOf:${nodeIds.pres}`)}" type="presOf" ` +
        `srcId="${nodeIds.data}" destId="${nodeIds.pres}" srcOrd="0" destOrd="0" ` +
        `presId="${layoutId}"/>`,
    )
    node.children.forEach((kid) => {
      const kidIds = ids.get(kid)!
      cxns.push(
        `<dgm:cxn modelId="${modelId(`presOfConn:${kidIds.connectorPres}`)}" type="presOf" ` +
          `srcId="${kidIds.sibTrans}" destId="${kidIds.connectorPres}" srcOrd="0" destOrd="0" ` +
          `presId="${layoutId}"/>`,
      )
    })
    for (const kid of node.children) mapPres(kid)
  }
  mapPres(root)

  // the presentation tree: a shape hangs under its parent's shape, and the
  // connector sits between a node and the row below it
  const mapPresTree = (node: FlatNode): void => {
    const nodeIds = ids.get(node)!
    node.children.forEach((kid, i) => {
      const kidIds = ids.get(kid)!
      // the shape, then the connector that reaches it: a connector is the
      // sibling right after the shape it leads to
      cxns.push(
        `<dgm:cxn modelId="${modelId(`presPar:${kidIds.pres}`)}" type="presParOf" ` +
          `srcId="${nodeIds.pres}" destId="${kidIds.pres}" srcOrd="${i * 2}" destOrd="0" ` +
          `presId="${layoutId}"/>`,
      )
      cxns.push(
        `<dgm:cxn modelId="${modelId(`presParConn:${kidIds.connectorPres}`)}" type="presParOf" ` +
          `srcId="${nodeIds.pres}" destId="${kidIds.connectorPres}" ` +
          `srcOrd="${i * 2 + 1}" destOrd="0" presId="${layoutId}"/>`,
      )
      mapPresTree(kid)
    })
  }
  mapPresTree(root)

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<dgm:dataModel xmlns:dgm="${DGM_NS}" xmlns:a="${A_NS}">` +
    `<dgm:ptLst>${pts.join('')}</dgm:ptLst>` +
    `<dgm:cxnLst>${cxns.join('')}</dgm:cxnLst>` +
    '<dgm:bg/><dgm:whole/>' +
    // the pre-rendered drawing, reached through the host part's relationship
    '<dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram">' +
    `<dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" ` +
    `relId="${drawingRelId}" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>` +
    '</a:ext></dgm:extLst>' +
    '</dgm:dataModel>'
  )
}

// ---- the layout ----

/**
 * The layout definition.
 *
 * The constraint names and factors are the ones the org-chart interpreter in
 * @genoffice/pptx-engine reads, so the pre-rendered drawing below and the
 * layout Word applies describe the same geometry: a box as wide as it is
 * tall times 0.5, a sibling gap of 0.21 and a row gap of 0.21.
 */
function buildLayoutXml(kind: SmartArtLayoutKind): string {
  const constraints =
    '<dgm:constrLst>' +
    '<dgm:constr type="h" for="des" forName="rootComposite1" refType="h" refForName="rootComposite1" fact="0.5"/>' +
    '<dgm:constr type="w" for="des" forName="rootComposite1" refType="w"/>' +
    '<dgm:constr type="sibSp" refType="w" refForName="rootComposite1" fact="0.21"/>' +
    '<dgm:constr type="sp" for="des" forName="hierRoot1" refType="h" refForName="rootComposite1" fact="0.21"/>' +
    '<dgm:constr type="w" for="ch" forName="rootConnector1" refType="w" refForName="rootComposite1" fact="0.1"/>' +
    '<dgm:constr type="h" for="ch" forName="rootConnector1" refType="h" refForName="rootComposite1" fact="0.2"/>' +
    '<dgm:constr type="primFontSz" for="des" forName="rootComposite1" op="equ" val="65"/>' +
    '<dgm:constr type="primFontSz" for="des" forName="hierRoot1" op="equ" val="60"/>' +
    '</dgm:constrLst>'
  const emptyDataModel = '<dgm:dataModel><dgm:ptLst/><dgm:bg/><dgm:whole/></dgm:dataModel>'

  if (kind === 'radial') {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      `<dgm:layoutDef xmlns:dgm="${DGM_NS}" xmlns:a="${A_NS}" uniqueId="${SMARTART_RADIAL_LAYOUT_ID}">` +
      '<dgm:title val="Radial Mind Map"/>' +
      '<dgm:desc val="A mind map: the topic in the middle, the branches around it."/>' +
      '<dgm:catLst><dgm:cat type="radial" pri="1000"/></dgm:catLst>' +
      `<dgm:sampData useDef="1">${emptyDataModel}</dgm:sampData>` +
      `<dgm:styleData useDef="1">${emptyDataModel}</dgm:styleData>` +
      `<dgm:clrData useDef="1">${emptyDataModel}</dgm:clrData>` +
      '<dgm:layoutNode name="diagram">' +
      '<dgm:varLst><dgm:animLvl val="lvl"/><dgm:resizeHandles val="exact"/></dgm:varLst>' +
      // the branches sit on the ring: the cycle arranges the children of this
      // node, which are the branches the forEach below produces
      '<dgm:alg type="cycle"><dgm:param type="stAng" val="0"/><dgm:param type="spanAng" val="360"/></dgm:alg>' +
      `<dgm:shape xmlns:r="${R_NS}" r:blip=""><dgm:adjLst/></dgm:shape>` +
      '<dgm:presOf/>' +
      '<dgm:constrLst>' +
      '<dgm:constr type="w" for="des" forName="radialNode" refType="w" fact="0.26"/>' +
      '<dgm:constr type="h" for="des" forName="radialNode" refType="w" fact="0.13"/>' +
      '<dgm:constr type="sibSp" refType="w" fact="0.12"/>' +
      '<dgm:constr type="sp" refType="h" fact="0.6"/>' +
      '<dgm:constr type="primFontSz" for="des" forName="radialNode" op="equ" val="55"/>' +
      '<dgm:constr type="primFontSz" for="des" forName="node" op="equ" val="50"/>' +
      '</dgm:constrLst>' +
      '<dgm:ruleLst/>' +
      '<dgm:forEach name="Name0" axis="ch" ptType="node">' +
      '<dgm:layoutNode name="radialNode">' +
      // inside a branch the children run outwards in a line
      '<dgm:alg type="lin"><dgm:param type="linDir" val="fromT"/></dgm:alg>' +
      `<dgm:shape xmlns:r="${R_NS}" r:blip=""><dgm:adjLst/></dgm:shape>` +
      '<dgm:presOf/>' +
      '<dgm:constrLst/><dgm:ruleLst/>' +
      '<dgm:layoutNode name="radialConnector">' +
      '<dgm:alg type="conn"><dgm:param type="dim" val="1D"/><dgm:param type="endSty" val="noArr"/></dgm:alg>' +
      `<dgm:shape xmlns:r="${R_NS}" type="conn" r:blip=""><dgm:adjLst/></dgm:shape>` +
      '<dgm:presOf axis="self"/>' +
      '<dgm:constrLst><dgm:constr type="begPad"/><dgm:constr type="endPad"/></dgm:constrLst>' +
      '<dgm:ruleLst/>' +
      '</dgm:layoutNode>' +
      '<dgm:forEach name="Name1" axis="ch" ptType="node">' +
      '<dgm:layoutNode name="node">' +
      '<dgm:alg type="space"/>' +
      `<dgm:shape xmlns:r="${R_NS}" type="roundRect" r:blip=""><dgm:adjLst/></dgm:shape>` +
      '<dgm:presOf axis="self" ptType="node"/>' +
      '<dgm:constrLst>' +
      '<dgm:constr type="lMarg" refType="primFontSz" fact="0.3"/>' +
      '<dgm:constr type="rMarg" refType="primFontSz" fact="0.3"/>' +
      '</dgm:constrLst>' +
      '<dgm:ruleLst><dgm:rule type="primFontSz" val="5" fact="NaN" max="NaN"/></dgm:ruleLst>' +
      '</dgm:layoutNode>' +
      '</dgm:forEach>' +
      '</dgm:layoutNode>' +
      '</dgm:forEach>' +
      '</dgm:layoutNode>' +
      '</dgm:layoutDef>'
    )
  }

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<dgm:layoutDef xmlns:dgm="${DGM_NS}" xmlns:a="${A_NS}" uniqueId="${SMARTART_LAYOUT_ID}">` +
    '<dgm:title val="Mind Map"/>' +
    '<dgm:desc val="A mind map: the topic in the middle, the branches to both sides."/>' +
    '<dgm:catLst><dgm:cat type="hierarchy" pri="1000"/></dgm:catLst>' +
    `<dgm:sampData useDef="1">${emptyDataModel}</dgm:sampData>` +
    `<dgm:styleData useDef="1">${emptyDataModel}</dgm:styleData>` +
    `<dgm:clrData useDef="1">${emptyDataModel}</dgm:clrData>` +
    '<dgm:layoutNode name="diagram">' +
    '<dgm:varLst><dgm:dir/><dgm:animLvl val="lvl"/><dgm:resizeHandles val="exact"/></dgm:varLst>' +
    '<dgm:alg type="hierRoot"/>' +
    `<dgm:shape xmlns:r="${R_NS}" r:blip=""><dgm:adjLst/></dgm:shape>` +
    '<dgm:presOf/>' +
    constraints +
    '<dgm:ruleLst/>' +
    '<dgm:forEach name="Name0" axis="ch" ptType="node">' +
    '<dgm:layoutNode name="rootComposite1">' +
    '<dgm:alg type="composite"/>' +
    `<dgm:shape xmlns:r="${R_NS}" r:blip=""><dgm:adjLst/></dgm:shape>` +
    '<dgm:presOf/>' +
    '<dgm:constrLst/><dgm:ruleLst/>' +
    '<dgm:layoutNode name="rootConnector1">' +
    '<dgm:alg type="conn"><dgm:param type="dim" val="1D"/><dgm:param type="endSty" val="noArr"/></dgm:alg>' +
    `<dgm:shape xmlns:r="${R_NS}" type="conn" r:blip=""><dgm:adjLst/></dgm:shape>` +
    '<dgm:presOf axis="self"/>' +
    '<dgm:constrLst><dgm:constr type="begPad"/><dgm:constr type="endPad"/></dgm:constrLst>' +
    '<dgm:ruleLst/>' +
    '</dgm:layoutNode>' +
    '<dgm:forEach name="Name1" axis="ch" ptType="node">' +
    '<dgm:layoutNode name="hierRoot1">' +
    '<dgm:alg type="hierChild"/>' +
    `<dgm:shape xmlns:r="${R_NS}" r:blip=""><dgm:adjLst/></dgm:shape>` +
    '<dgm:presOf axis="self" ptType="node"/>' +
    '<dgm:constrLst><dgm:constr type="tMarg" refType="primFontSz" fact="0.3"/>' +
    '<dgm:constr type="bMarg" refType="primFontSz" fact="0.3"/>' +
    '<dgm:constr type="lMarg" refType="primFontSz" fact="0.3"/>' +
    '<dgm:constr type="rMarg" refType="primFontSz" fact="0.3"/></dgm:constrLst>' +
    '<dgm:ruleLst/>' +
    '</dgm:layoutNode>' +
    '</dgm:forEach>' +
    '</dgm:layoutNode>' +
    '<dgm:layoutNode name="text">' +
    '<dgm:alg type="tx"/>' +
    `<dgm:shape xmlns:r="${R_NS}" type="rect" r:blip=""><dgm:adjLst/></dgm:shape>` +
    '<dgm:presOf axis="self" ptType="node"/>' +
    '<dgm:constrLst>' +
    '<dgm:constr type="lMarg" refType="primFontSz" fact="0.3"/>' +
    '<dgm:constr type="rMarg" refType="primFontSz" fact="0.3"/>' +
    '<dgm:constr type="tMarg" refType="primFontSz" fact="0.3"/>' +
    '<dgm:constr type="bMarg" refType="primFontSz" fact="0.3"/>' +
    '</dgm:constrLst>' +
    '<dgm:ruleLst><dgm:rule type="primFontSz" val="5" fact="NaN" max="NaN"/></dgm:ruleLst>' +
    '</dgm:layoutNode>' +
    '</dgm:forEach>' +
    '</dgm:layoutNode>' +
    '</dgm:layoutDef>'
  )
}

// ---- the quick style and the colours ----

function buildQuickStyleXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<dgm:styleDef xmlns:dgm="${DGM_NS}" xmlns:a="${A_NS}" uniqueId="${SMARTART_QUICK_STYLE_ID}">` +
    '<dgm:title val="Mind Map"/><dgm:desc val=""/>' +
    '<dgm:catLst><dgm:cat type="simple" pri="10100"/></dgm:catLst>' +
    '<dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d>' +
    '<dgm:styleLbl name="node1">' +
    '<dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d>' +
    '<dgm:sp3d/><dgm:txPr/>' +
    '<dgm:style>' +
    '<a:lnRef idx="2"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef>' +
    '<a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef>' +
    '<a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef>' +
    '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>' +
    '</dgm:style>' +
    '</dgm:styleLbl>' +
    '<dgm:styleLbl name="rootConnector1">' +
    '<dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d>' +
    '<dgm:sp3d/><dgm:txPr/>' +
    '<dgm:style>' +
    '<a:lnRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef>' +
    '<a:fillRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef>' +
    '<a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef>' +
    '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>' +
    '</dgm:style>' +
    '</dgm:styleLbl>' +
    '</dgm:styleDef>'
  )
}

function buildColorsXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<dgm:colorsDef xmlns:dgm="${DGM_NS}" xmlns:a="${A_NS}" uniqueId="${SMARTART_COLORS_ID}">` +
    '<dgm:title val="Mind Map"/><dgm:desc val=""/>' +
    '<dgm:catLst><dgm:cat type="accent1" pri="11100"/></dgm:catLst>' +
    '<dgm:styleLbl name="node1">' +
    '<dgm:fillClrLst meth="repeat"><a:schemeClr val="accent1"/></dgm:fillClrLst>' +
    '<dgm:linClrLst meth="repeat"><a:schemeClr val="lt1"/></dgm:linClrLst>' +
    '<dgm:effectClrLst/><dgm:txLinClrLst/>' +
    '<dgm:txFillClrLst meth="repeat"><a:schemeClr val="lt1"/></dgm:txFillClrLst>' +
    '<dgm:txEffectClrLst/>' +
    '</dgm:styleLbl>' +
    '<dgm:styleLbl name="rootConnector1">' +
    '<dgm:fillClrLst meth="repeat"><a:schemeClr val="accent1"/></dgm:fillClrLst>' +
    '<dgm:linClrLst meth="repeat"><a:schemeClr val="accent1"/></dgm:linClrLst>' +
    '<dgm:effectClrLst/><dgm:txLinClrLst/><dgm:txFillClrLst/><dgm:txEffectClrLst/>' +
    '</dgm:styleLbl>' +
    '</dgm:colorsDef>'
  )
}

// ---- radial layout ----

/** The metrics of the radial layout, in px. */
const RADIAL_METRICS = {
  nodeWidth: 150,
  nodeHeight: 30,
  /** the radius the first ring sits at, in px: a box wide plus the gap, so
   *  the topic and the branch beside it do not overlap */
  firstRing: 168,
  /** the distance between two rings, in px: a box wide plus the gap, so a
   *  parent and the child below it do not overlap either */
  ringStep: 168,
  /** the share of a branch's slice its own children get */
  childSpread: 0.72,
  /** the degrees between two children on the same arm */
  fanDeg: 9,
  /** the room kept between two boxes */
  boxGap: 18,
} as const

/** The boxes and connectors of the radial layout, in px. */
interface RadialGeometry {
  boxes: Array<{ node: FlatNode; x: number; y: number; w: number; h: number }>
  lines: Array<{
    x1: number
    y1: number
    x2: number
    y2: number
    /** the nodes the connector joins */
    from: FlatNode
    to: FlatNode
  }>
}

/**
 * The radial layout: the topic in the middle, the branches around it.
 *
 * Every branch owns a slice of the circle. The children of a branch sit on
 * the next ring, spread over a share of that slice — the ring is pushed out
 * until the boxes fit, so neighbours never overlap.
 */
function layoutRadial(root: FlatNode): RadialGeometry {
  const { nodeWidth: W, nodeHeight: H, firstRing, childSpread, fanDeg, boxGap } = RADIAL_METRICS
  const boxes: RadialGeometry['boxes'] = []
  const lines: RadialGeometry['lines'] = []
  // the map is laid out around the origin; the frame is measured from the
  // result, so a ring that had to be pushed out is never cut off
  const cx = 0
  const cy = 0

  const push = (
    node: FlatNode,
    x: number,
    y: number,
  ): { x: number; y: number; w: number; h: number } => {
    const box = { node, x: x - W / 2, y: y - H / 2, w: W, h: H }
    boxes.push(box)
    return box
  }

  push(root, cx, cy)

  /**
   * The children of a node sit on the next ring, spread over their share of
   * the parent's slice. The ring is pushed out until the boxes of that ring
   * fit without touching.
   */
  const place = (
    parent: { x: number; y: number; w: number; h: number },
    node: FlatNode,
    angleCenter: number,
    angleSpan: number,
    prevRing: number,
  ): void => {
    const kids = node.children
    if (kids.length === 0) return
    const slice = angleSpan / kids.length
    kids.forEach((kid, i) => {
      // the children walk out along the arm their parent opened, one box
      // apart: boxes on a ring would sit on top of each other wherever the
      // ring runs sideways
      const deg = angleCenter + (i - (kids.length - 1) / 2) * fanDeg
      const rad = (deg * Math.PI) / 180
      const step = prevRing + (i + 1) * (W + boxGap)
      const x = cx + Math.cos(rad) * step
      const y = cy + Math.sin(rad) * step
      const box = push(kid, x, y)
      lines.push({
        x1: parent.x + parent.w / 2,
        y1: parent.y + parent.h / 2,
        x2: x,
        y2: y,
        from: node,
        to: kid,
      })
      place(box, kid, deg, slice * childSpread, step)
    })
  }

  const branches = root.children
  const count = branches.length
  if (count > 0) {
    // the branches fill the circle, the first one at the top
    const slice = 360 / count
    const ring = Math.max(firstRing, (count * (W + boxGap)) / (2 * Math.PI))
    branches.forEach((branch, i) => {
      const deg = -90 + slice * (i + 0.5)
      const rad = (deg * Math.PI) / 180
      const x = cx + Math.cos(rad) * ring
      const y = cy + Math.sin(rad) * ring
      const box = push(branch, x, y)
      lines.push({
        x1: cx,
        y1: cy,
        x2: x,
        y2: y,
        from: root,
        to: branch,
      })
      place(box, branch, deg, slice * childSpread, ring)
    })
  }
  return { boxes, lines }
}

// ---- the pre-rendered drawing ----

/** The geometry the drawing is written from, in px. */
/** The boxes and connectors of a layout, in px. */
interface DrawingGeometry {
  boxes: Array<{
    label: string
    depth: number
    x: number
    y: number
    w: number
    h: number
    /** the node the box stands for, when the layout knows it */
    node?: FlatNode
    /** the presentation point the shape stands for: the drawing is matched to
     *  the data model through it */
    presId?: string
  }>
  lines: Array<{
    x1: number
    y1: number
    x2: number
    y2: number
    /** the boxes the connector joins, so it can meet their borders */
    from?: FlatNode
    to?: FlatNode
    /** the presentation points the connector stands for */
    fromPres?: string
    toPres?: string
  }>
}

/** Turn the tree into the shape the org-chart interpreter lays out. */
function toHierTree(node: FlatNode, depth = 0, depths?: Map<HierTreeNode, number>): HierTreeNode {
  const hier: HierTreeNode = {
    texts: node.label ? [node.label] : [],
    children: node.children.map((kid) => toHierTree(kid, depth + 1, depths)),
  }
  depths?.set(hier, depth)
  return hier
}

/** The box a node is drawn in, when the layout knows the node. */
function boxOfNode(
  boxes: DrawingGeometry['boxes'],
  node: FlatNode,
): DrawingGeometry['boxes'][number] | null {
  for (const box of boxes) {
    if (box.node === node) return box
  }
  return null
}

/**
 * Where the segment towards (tx, ty) leaves the box: the border point in
 * that direction, so a connector starts at the edge of the shape.
 */
function borderPoint(
  box: { x: number; y: number; w: number; h: number },
  tx: number,
  ty: number,
): { x: number; y: number } {
  const cx = box.x + box.w / 2
  const cy = box.y + box.h / 2
  const dx = tx - cx
  const dy = ty - cy
  if (dx === 0 && dy === 0) return { x: cx, y: cy }
  const sx = dx === 0 ? Infinity : box.w / 2 / Math.abs(dx)
  const sy = dy === 0 ? Infinity : box.h / 2 / Math.abs(dy)
  const s = Math.min(sx, sy)
  return { x: cx + dx * s, y: cy + dy * s }
}

/**
 * The pre-rendered shapes.
 *
 * Word shows these before it recomputes anything, so the diagram is correct
 * the moment the document opens. The coordinates come from the same
 * interpreter that reads a real org chart, so the picture matches what the
 * layout produces when Word re-lays it out.
 */
function buildDrawingXml(geo: DrawingGeometry): string {
  const { boxes, lines } = geo
  const sps: string[] = []

  const nodeColor = (depth: number): string =>
    depth === 0 ? 'E8663F' : depth === 1 ? '4472C4' : '5B9BD5'
  const textColor = 'FFFFFF'

  for (const box of boxes) {
    const depth = box.depth
    const x = box.x
    const y = box.y
    const w = box.w
    const h = box.h
    const label = box.label
    // the text follows the box: a scaled map keeps its proportions
    const fontSize = Math.max(700, Math.round((1200 * box.w) / BASE_NODE_WIDTH))
    sps.push(
      `<dsp:sp modelId="${box.presId ?? modelId(`dsp:${label}:${x}:${y}`)}">` +
        '<dsp:nvSpPr>' +
        `<dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/>` +
        '</dsp:nvSpPr>' +
        '<dsp:spPr>' +
        `<a:xfrm><a:off x="${Math.round(x * EMU_PER_PX)}" y="${Math.round(y * EMU_PER_PX)}"/>` +
        `<a:ext cx="${Math.round(w * EMU_PER_PX)}" cy="${Math.round(h * EMU_PER_PX)}"/></a:xfrm>` +
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
        '<dsp:txBody>' +
        '<a:bodyPr spcFirstLastPara="0" vert="horz" wrap="square" lIns="45720" tIns="45720" ' +
        'rIns="45720" bIns="45720" numCol="1" spcCol="1270" anchor="ctr" anchorCtr="0">' +
        '<a:noAutofit/></a:bodyPr>' +
        '<a:lstStyle/>' +
        '<a:p><a:pPr lvl="0" algn="ctr" defTabSz="1111250"/>' +
        `<a:r><a:rPr lang="en-US" sz="${fontSize}" b="1" kern="1200"><a:solidFill><a:srgbClr val="${textColor}"/></a:solidFill></a:rPr>` +
        `<a:t>${escapeXmlText(label)}</a:t></a:r>` +
        '</a:p>' +
        '</dsp:txBody>' +
        `<dsp:txXfrm><a:off x="${Math.round(x * EMU_PER_PX)}" y="${Math.round(y * EMU_PER_PX)}"/>` +
        `<a:ext cx="${Math.round(w * EMU_PER_PX)}" cy="${Math.round(h * EMU_PER_PX)}"/></dsp:txXfrm>` +
        '</dsp:sp>',
    )
  }

  // the connectors: stroke-only lines from the border of one box to the
  // border of the next, so a line never runs through a shape
  for (const line of lines) {
    const from = line.from ? boxOfNode(boxes, line.from) : null
    const to = line.to ? boxOfNode(boxes, line.to) : null
    const start = from ? borderPoint(from, line.x2, line.y2) : { x: line.x1, y: line.y1 }
    const end = to ? borderPoint(to, line.x1, line.y1) : { x: line.x2, y: line.y2 }

    const x = Math.min(start.x, end.x)
    const y = Math.min(start.y, end.y)
    const cx = Math.abs(end.x - start.x)
    const cy = Math.abs(end.y - start.y)
    if (cx < 1 && cy < 1) continue
    // the line preset draws the top left to the bottom right diagonal
    const flipV = end.y < start.y
    sps.push(
      `<dsp:sp modelId="${line.toPres ?? modelId(`dspLine:${x}:${y}:${cx}:${cy}`)}">` +
        '<dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>' +
        '<dsp:spPr>' +
        // a line preset draws the diagonal of its box, from the top left to the
        // bottom right: a connector that climbs needs the box flipped
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

/** The depth of a node inside the tree, for the colour ramp. */

// ---- the document paragraph ----

/**
 * The paragraph Word puts in the body: an inline drawing that points at the
 * four diagram parts.
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
    `<wp:docPr id="${docPrId}" name="Mind map ${docPrId}"/>` +
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

/**
 * Build every part a Word diagram needs.
 *
 * `relIds` are the relationship ids the caller puts in document.xml; the
 * drawing is reached from data1.xml through its own rels, so it needs no id
 * here.
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
  const kind: SmartArtLayoutKind = options.orientation === 'radial' ? 'radial' : 'hierarchy'

  // the geometry: the org-chart interpreter places the tree inside the frame
  // and hands back pixel coordinates, so the frame is the size of the diagram
  const depths = new Map<HierTreeNode, number>()
  const hierRoot = toHierTree(flat, 0, depths)
  const cons: HierConstraints = {
    boxAspect: 0.5,
    sibSp: 0.21,
    sp: 0.21,
    trunkOff: 0.1,
    hangIndent: 0.25,
    fontMax: 65,
  }
  let geo: DrawingGeometry
  let widthPx: number
  let heightPx: number

  if (options.orientation === 'radial') {
    // the topic in the middle, the branches around it
    const radial = layoutRadial(flat)
    // the topic sits in the middle, the branches reach out from it
    const reach = radial.boxes.length
      ? Math.max(
          ...radial.boxes.map(
            (b) => Math.hypot(b.x + b.w / 2, b.y + b.h / 2) + Math.max(b.w, b.h) / 2,
          ),
        )
      : 0
    const pad = 12
    const wanted = Math.ceil(reach * 2 + pad * 2)
    // a map wider than the page scales down as a whole, so it still fits
    const limit = options.maxSizePx ?? 700
    const scale = Math.min(1, limit / Math.max(1, wanted))
    widthPx = options.widthPx ?? Math.ceil(wanted * scale)
    heightPx = options.heightPx ?? Math.ceil(wanted * scale)
    const shiftX = widthPx / 2
    const shiftY = heightPx / 2
    geo = {
      boxes: radial.boxes.map((b) => ({
        label: b.node.label,
        depth: b.node.depth,
        x: b.x * scale + shiftX,
        y: b.y * scale + shiftY,
        w: b.w * scale,
        h: b.h * scale,
        node: b.node,
        presId: ids.get(b.node)?.pres,
      })),
      lines: radial.lines.map((l) => ({
        x1: l.x1 * scale + shiftX,
        y1: l.y1 * scale + shiftY,
        x2: l.x2 * scale + shiftX,
        y2: l.y2 * scale + shiftY,
        from: l.from,
        to: l.to,
        toPres: l.to ? ids.get(l.to)?.connectorPres : undefined,
      })),
    }
  } else {
    // A probe pass measures the tree, the real pass fills the frame with it.
    // The interpreter scales the tree until it fills the frame, so the probe
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
    widthPx = options.widthPx ?? 800
    heightPx = options.heightPx ?? Math.max(140, Math.round(widthPx / ratio))
    const laid = layoutHierTree([hierRoot], cons, widthPx, heightPx) ?? probe
    geo = {
      boxes: (laid?.boxes ?? []).map((b) => ({
        label: b.node.texts.join(' '),
        depth: depths.get(b.node) ?? 0,
        x: b.x,
        y: b.y,
        w: b.w,
        h: b.h,
      })),
      lines: (laid?.lines ?? []).map((l) => ({
        x1: l.x,
        y1: l.y,
        x2: l.x + l.cx,
        y2: l.y + l.cy,
      })),
    }
  }

  return {
    dataXml: buildDataXml(flat, ids, relIds.drawing, kind),
    layoutXml: buildLayoutXml(kind),
    quickStyleXml: buildQuickStyleXml(),
    colorsXml: buildColorsXml(),
    drawingXml: buildDrawingXml(geo),
    paragraphXml: buildParagraphXml(relIds, widthPx, heightPx, docPrId),
    widthPx,
    heightPx,
  }
}

/** The number of nodes a tree holds — the shape count the drawing gets. */
export function smartArtNodeCount(root: SmartArtNode): number {
  return countNodes(flatten(root))
}

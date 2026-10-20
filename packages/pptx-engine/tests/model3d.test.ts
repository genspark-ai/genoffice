/**
 * 3D models as PowerPoint stores them: am3d:model3d in an mc:Choice, the model's rendering as
 * the Fallback picture the editor works with, both kept in step through move, resize, turn
 * and save.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { validatePptx, xmllintAvailable } from '../../../tools/ooxml-validate/validate-pptx.mjs'
import {
  MODEL3D_REL_TYPE,
  addModel3d,
  createBlankPptx,
  deleteElement,
  glbBounds,
  openPptx,
  savePptx,
  setModel3dView,
  solidPng,
  type OpenedPptx,
  type PictureElement,
} from '../src/index'
import { model3dFraming, setModel3dRotationXml, readModel3dRotation } from '../src/model3d'

const schemaGate = xmllintAvailable() || !!process.env.CI
const OFF = { x: 1000000, y: 500000, cx: 3000000, cy: 2000000 }

/** A .glb whose JSON chunk holds `gltf` (no binary chunk: only accessor bounds are read). */
function glb(gltf: object): Uint8Array {
  let json = Buffer.from(JSON.stringify(gltf), 'utf8')
  json = Buffer.concat([json, Buffer.alloc((4 - (json.length % 4)) % 4, 0x20)])
  const head = Buffer.alloc(20)
  head.writeUInt32LE(0x46546c67, 0)
  head.writeUInt32LE(2, 4)
  head.writeUInt32LE(20 + json.length, 8)
  head.writeUInt32LE(json.length, 12)
  head.writeUInt32LE(0x4e4f534a, 16)
  return new Uint8Array(Buffer.concat([head, json]))
}

/** A unit cube mesh (corners 0..1) under the given nodes. */
const cubeScene = (nodes: object[], sceneNodes = [0]) =>
  glb({
    scene: 0,
    scenes: [{ nodes: sceneNodes }],
    nodes,
    meshes: [{ primitives: [{ attributes: { POSITION: 0 } }] }],
    accessors: [{ min: [0, 0, 0], max: [1, 1, 1] }],
  })

const lastPicture = (opened: OpenedPptx) => opened.deck.slides[0]!.elements.at(-1) as PictureElement

async function slideXml(bytes: Uint8Array): Promise<string> {
  return (await JSZip.loadAsync(bytes)).file('ppt/slides/slide1.xml')!.async('string')
}

describe('glbBounds', () => {
  it('carries mesh bounds through node translation, scale and children', () => {
    const bounds = glbBounds(
      cubeScene([
        { translation: [10, 0, 0], scale: [2, 2, 2], children: [1] },
        { mesh: 0, translation: [0, 1, 0] },
      ]),
    )
    expect(bounds).toEqual({ center: [11, 3, 1], size: [2, 2, 2] })
  })

  it('reads nothing from bytes that are not glTF binary', () => {
    expect(glbBounds(new Uint8Array([1, 2, 3, 4]))).toBeNull()
    expect(glbBounds(cubeScene([{}]))).toBeNull()
  })
})

describe('3D model', () => {
  it('is written as PowerPoint writes it, and reads back as its picture with the model behind it', async () => {
    const opened = await openPptx(await createBlankPptx())
    const bytes = cubeScene([{ mesh: 0, scale: [4, 2, 2] }])
    addModel3d(opened, 0, { bytes, ext: 'glb', offset: OFF })
    const saved = await savePptx(opened)
    const xml = await slideXml(saved)
    const zip = await JSZip.loadAsync(saved)

    expect(xml).toMatch(
      /<mc:AlternateContent [^>]*><mc:Choice xmlns:am3d="[^"]+" Requires="am3d"><p:graphicFrame>/,
    )
    const modelRId = /<am3d:model3d r:embed="([^"]+)">/.exec(xml)![1]
    const rasterRId = /<am3d:blip r:embed="([^"]+)"\/>/.exec(xml)![1]
    const fallbackRId = /<mc:Fallback>[\s\S]*<a:blip r:embed="([^"]+)"\/>/.exec(xml)![1]
    expect(rasterRId).toBe(fallbackRId)
    const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')
    expect(rels).toContain(
      `<Relationship Id="${modelRId}" Type="${MODEL3D_REL_TYPE}" Target="../media/model3d1.glb"/>`,
    )
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain(
      '<Default Extension="glb" ContentType="model/gltf.binary"/>',
    )

    // a 4 m model: framed at one meter across, centred on (2, 1, 1) → preTrans -(0.5, 0.25, 0.25) m
    expect(xml).toContain('<am3d:meterPerModelUnit n="1000000" d="4000000"/>')
    expect(xml).toContain('<am3d:preTrans dx="-18000000" dy="-9000000" dz="-9000000"/>')
    expect(xml).toContain(`<am3d:pos x="0" y="0" z="${model3dFraming(glbBounds(bytes)).cameraZ}"/>`)
    expect(xml).toContain('<am3d:objViewport viewportSz="3000000"/>')

    const back = await openPptx(saved)
    const el = lastPicture(back)
    expect(el.model3d).toEqual({ target: 'ppt/media/model3d1.glb', rot: { ax: 0, ay: 0, az: 0 } })
    expect(el.transform.offset).toEqual(OFF)
    if (schemaGate) expect(await validatePptx(saved)).toEqual([])
  })

  it('moving or resizing the picture moves and resizes the model PowerPoint shows', async () => {
    const opened = await openPptx(await createBlankPptx())
    addModel3d(opened, 0, { bytes: cubeScene([{ mesh: 0 }]), ext: 'glb', offset: OFF })
    const back = await openPptx(await savePptx(opened))
    const el = lastPicture(back)
    el.transform = { ...el.transform, offset: { x: 200, y: 300, cx: 4000000, cy: 5000000 } }
    el.dirtyTransform = true
    const saved = await savePptx(back)
    const xml = await slideXml(saved)

    expect(xml).toContain(
      '<p:xfrm><a:off x="200" y="300"/><a:ext cx="4000000" cy="5000000"/></p:xfrm>',
    )
    expect(xml).toContain(
      '<am3d:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="5000000"/>',
    )
    expect(xml).toContain('<am3d:objViewport viewportSz="5000000"/>')
    expect(xml).toContain(
      '<p:spPr><a:xfrm><a:off x="200" y="300"/><a:ext cx="4000000" cy="5000000"/></a:xfrm>',
    )
    if (schemaGate) expect(await validatePptx(saved)).toEqual([])
  })

  it('turning the model stores its rotation and its new rendering in both branches', async () => {
    const opened = await openPptx(await createBlankPptx())
    addModel3d(opened, 0, { bytes: cubeScene([{ mesh: 0 }]), ext: 'glb', offset: OFF })
    const slide = opened.deck.slides[0]!
    const el = lastPicture(opened)
    const firstPoster = el.mediaRef
    const rendering = solidPng(4, 4, [200, 10, 10])

    expect(
      setModel3dView(opened, slide, el.id, {
        rot: { ax: 1800000, ay: -2700000, az: 0 },
        poster: { bytes: rendering, ext: 'png' },
      }),
    ).toBe(true)
    const saved = await savePptx(opened)
    const xml = await slideXml(saved)
    const back = await openPptx(saved)
    const turned = lastPicture(back)

    expect(xml).toContain('<am3d:rot ax="1800000" ay="-2700000" az="0"/>')
    expect(turned.model3d?.rot).toEqual({ ax: 1800000, ay: -2700000, az: 0 })
    expect(Buffer.from(back.archive.readBytes(turned.mediaRef)!)).toEqual(Buffer.from(rendering))
    expect(/<am3d:blip r:embed="([^"]+)"/.exec(xml)![1]).toBe(
      /<mc:Fallback>[\s\S]*<a:blip r:embed="([^"]+)"/.exec(xml)![1],
    )
    expect(back.archive.entries.has(firstPoster)).toBe(false)
    if (schemaGate) expect(await validatePptx(saved)).toEqual([])
  })

  it('a rotation is added where the schema places it when the model has none', () => {
    const xml =
      '<am3d:trans><am3d:meterPerModelUnit n="1" d="1"/><am3d:postTrans dx="0" dy="0" dz="0"/></am3d:trans>'
    expect(readModel3dRotation(xml)).toEqual({ ax: 0, ay: 0, az: 0 })
    const turned = setModel3dRotationXml(xml, { ax: 60000, ay: 0, az: -60000 })
    expect(turned).toBe(
      '<am3d:trans><am3d:meterPerModelUnit n="1" d="1"/><am3d:rot ax="60000" ay="0" az="-60000"/>' +
        '<am3d:postTrans dx="0" dy="0" dz="0"/></am3d:trans>',
    )
    expect(readModel3dRotation(turned)).toEqual({ ax: 60000, ay: 0, az: -60000 })
  })

  it('deleting the model drops its model part and its rendering', async () => {
    const opened = await openPptx(await createBlankPptx())
    const added = addModel3d(opened, 0, {
      bytes: cubeScene([{ mesh: 0 }]),
      ext: 'glb',
      offset: OFF,
    })!
    const media = () => [...opened.archive.entries.keys()].filter((p) => p.startsWith('ppt/media/'))
    expect(media()).toHaveLength(2)
    expect(deleteElement(opened, added.slide, added.elementId)).toBe(true)
    expect(media()).toEqual([])
  })

  it('only glTF binary is embedded, as PowerPoint does', async () => {
    const opened = await openPptx(await createBlankPptx())
    expect(
      addModel3d(opened, 0, { bytes: new Uint8Array([123]), ext: 'gltf', offset: OFF }),
    ).toBeNull()
  })
})

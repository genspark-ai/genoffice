import { describe, expect, it } from 'vitest'
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { ImageEditLayer, type LocalImageEdit } from '../src/renderer/ImageEditLayer'

const rect: [number, number, number, number] = [100, 100, 300, 250]

function render(layer: 'belowText' | 'aboveText'): Document {
  const edit: LocalImageEdit = {
    id: 'op1',
    input: { kind: 'transformImage', pageIndex: 0, oldRect: rect, rect, layer },
    png: 'iVBORw0KGgo=',
  }
  const html = renderToStaticMarkup(
    createElement(ImageEditLayer, {
      geom: { pw: 612, ph: 792, rot: 0 },
      scale: 1,
      edits: [edit],
      existing: [],
      selectedId: null,
      selectedKey: null,
      editHint: 'hint',
      onSelectEdit: () => {},
      onSelectExisting: () => {},
    }),
  )
  return new DOMParser().parseFromString(html, 'text/html')
}

describe('ImageEditLayer', () => {
  it('gives an edited picture below the text a hit target outside the click-through layer', () => {
    const doc = render('belowText')
    expect(doc.querySelector('.pdf-imgedit-under .pdf-imgedit-img')).not.toBeNull()
    const hits = [...doc.querySelectorAll('.pdf-imgedit-hit')]
    expect(hits).toHaveLength(1)
    expect(hits[0]!.closest('.pdf-imgedit-under')).toBeNull()
  })

  it('lets an edited picture above the text take clicks itself', () => {
    const doc = render('aboveText')
    expect(doc.querySelector('.pdf-imgedit-under')).toBeNull()
    expect(doc.querySelector('.pdf-imgedit-img')).not.toBeNull()
    expect(doc.querySelector('.pdf-imgedit-hit')).toBeNull()
  })
})

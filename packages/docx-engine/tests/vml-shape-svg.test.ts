import { describe, expect, it } from 'vitest'
import {
  vmlFloatAnchor,
  vmlFraction,
  vmlLengthPx,
  vmlRotationDeg,
  vmlShapeSvg,
} from '../src/parse-vml'
import { readWatermarkShape } from '../src/watermark'

const svgOf = (pict: string, types?: Map<string, Record<string, string>>): string =>
  decodeURIComponent(vmlShapeSvg(pict, types)!.dataUrl.replace('data:image/svg+xml,', ''))

describe('VML style readers', () => {
  it('lengths accept pt (default), in, cm, mm and px', () => {
    expect(vmlLengthPx('72pt')).toBe(96)
    expect(vmlLengthPx('16in')).toBe(1536)
    expect(vmlLengthPx('2.54cm')).toBe(96)
    expect(vmlLengthPx('25.4mm')).toBeCloseTo(96, 6)
    expect(vmlLengthPx('10px')).toBe(10)
    expect(vmlLengthPx('9')).toBe(12)
    expect(vmlLengthPx('auto')).toBeUndefined()
  })

  it('rotation normalizes negatives and 1/65536-degree fd values; 0 reads as none', () => {
    expect(vmlRotationDeg('rotation:315')).toBe(315)
    expect(vmlRotationDeg('rotation:-45')).toBe(315)
    expect(vmlRotationDeg('width:1pt;rotation:41366637fd')).toBeCloseTo(271.2, 1)
    expect(vmlRotationDeg('rotation:0')).toBeUndefined()
    expect(vmlRotationDeg('width:1pt')).toBeUndefined()
  })

  it('fractions read fixed-point "f" suffixes', () => {
    expect(vmlFraction('32768f')).toBe(0.5)
    expect(vmlFraction('.5')).toBe(0.5)
    expect(vmlFraction(undefined)).toBeUndefined()
  })

  it('anchor: keyword alignment wins, offsets otherwise; text-relative vertical is paragraph-relative', () => {
    const a: Record<string, unknown> = {}
    vmlFloatAnchor(
      'mso-position-horizontal:center;mso-position-horizontal-relative:margin;' +
        'margin-top:10pt;mso-position-vertical:absolute;mso-position-vertical-relative:text',
      a,
    )
    expect(a).toEqual({ posH: 'center', posHRel: 'margin', posYPx: 13, posVRel: 'paragraph' })
    const b: Record<string, unknown> = {}
    vmlFloatAnchor(
      'margin-left:1in;mso-position-horizontal-relative:page;mso-position-vertical:bottom',
      b,
    )
    expect(b).toEqual({ posXPx: 96, posHRel: 'page', posV: 'bottom', posVRel: 'margin' })
  })
})

describe('vmlShapeSvg', () => {
  it('draws a rounded rectangle with default arcsize and default black 0.75pt stroke', () => {
    const svg = svgOf(
      '<w:pict><v:roundrect style="position:absolute;width:60pt;height:30pt" fillcolor="#ffff00"/></w:pict>',
    )
    expect(svg).toContain('viewBox="0 0 80 40"')
    expect(svg).toContain('<rect x="0.5" y="0.5" width="79" height="39" rx="3.9"')
    expect(svg).toContain('fill="#ffff00"')
    expect(svg).toContain('stroke="#000000" stroke-width="1"')
  })

  it('resolves a straight path through the shapetype the shape points at, honors filled="f"', () => {
    const svg = svgOf(
      '<w:pict><v:shapetype id="_x0000_t4" coordsize="21600,21600" path="m10800,l,10800,10800,21600,21600,10800xe"/>' +
        '<v:shape type="#_x0000_t4" style="position:absolute;width:100pt;height:50pt" filled="f" strokecolor="blue" strokeweight="3pt"/></w:pict>',
    )
    expect(svg).toContain('<path d="M 66.5 2 L 2 33.5 L 66.5 65 L 131 33.5 Z"')
    expect(svg).toContain('fill="none"')
    expect(svg).toContain('stroke="#0000FF" stroke-width="4"')
  })

  it('bails on curved paths, groups, textboxes and pictures', () => {
    const shape = (inner: string) => `<w:pict>${inner}</w:pict>`
    expect(
      vmlShapeSvg(
        shape(
          '<v:shape style="width:10pt;height:10pt" coordsize="100,100" path="m0,0c10,10,20,20,30,30e"/>',
        ),
      ),
    ).toBeNull()
    expect(
      vmlShapeSvg(shape('<v:group style="width:10pt;height:10pt"><v:oval/></v:group>')),
    ).toBeNull()
    expect(
      vmlShapeSvg(shape('<v:rect style="width:10pt;height:10pt"><v:textbox/></v:rect>')),
    ).toBeNull()
    expect(
      vmlShapeSvg(
        shape('<v:shape style="width:10pt;height:10pt"><v:imagedata r:id="rId1"/></v:shape>'),
      ),
    ).toBeNull()
  })

  it('a bare oval keeps VML defaults: white fill, black 1px stroke', () => {
    const svg = svgOf(
      '<w:pict><v:oval style="position:absolute;width:62.25pt;height:46.5pt"/></w:pict>',
    )
    expect(svg).toContain(
      '<ellipse cx="41.5" cy="31" rx="41" ry="30.5" fill="#FFFFFF" stroke="#000000" stroke-width="1"/>',
    )
  })

  it('gradient fills become a linearGradient: angle 0 runs color2 (top) to fillcolor (bottom), |focus| 100% swaps', () => {
    const plain = svgOf(
      '<w:pict><v:rect style="width:10pt;height:20pt" fillcolor="white" stroked="f"><v:fill color2="#d6e3bc" type="gradient"/></v:rect></w:pict>',
    )
    expect(plain).toContain('<linearGradient id="g" x1="0.5" y1="0" x2="0.5" y2="1">')
    expect(plain).toContain(
      '<stop offset="0" stop-color="#d6e3bc"/><stop offset="1" stop-color="#FFFFFF"/>',
    )
    expect(plain).toContain('fill="url(#g)"')
    const focused = svgOf(
      '<w:pict><v:rect style="width:10pt;height:20pt" fillcolor="white" stroked="f"><v:fill color2="#d6e3bc" focus="100%" type="gradient"/></v:rect></w:pict>',
    )
    expect(focused).toContain(
      '<stop offset="0" stop-color="#FFFFFF"/><stop offset="1" stop-color="#d6e3bc"/>',
    )
    const axial = svgOf(
      '<w:pict><v:rect style="width:10pt;height:20pt" fillcolor="white" stroked="f"><v:fill color2="#d6e3bc" focus="50%" type="gradient"/></v:rect></w:pict>',
    )
    expect(axial).toContain(
      '<stop offset="0" stop-color="#FFFFFF"/><stop offset="0.5" stop-color="#d6e3bc"/><stop offset="1" stop-color="#FFFFFF"/>',
    )
    // 25%..75% approximates as axial like ±50% (LibreOffice import rule)
    expect(
      svgOf(
        '<w:pict><v:rect style="width:10pt;height:20pt" fillcolor="white" stroked="f"><v:fill color2="#d6e3bc" focus="-30%" type="gradient"/></v:rect></w:pict>',
      ),
    ).toContain(
      '<stop offset="0" stop-color="#d6e3bc"/><stop offset="0.5" stop-color="#FFFFFF"/><stop offset="1" stop-color="#d6e3bc"/>',
    )
    // a gradient with no second color paints the fillcolor solid
    expect(
      svgOf(
        '<w:pict><v:oval style="width:10pt;height:10pt"><v:fill type="gradient"/></v:oval></w:pict>',
      ),
    ).toContain('fill="#FFFFFF"')
  })

  it('resolves geometry and paint from a shapetype declared elsewhere in the document', () => {
    const types = new Map([
      [
        '_x0000_t110',
        {
          id: '_x0000_t110',
          coordsize: '21600,21600',
          'o:spt': '110',
          path: 'm10800,l,10800,10800,21600,21600,10800xe',
          strokecolor: 'red',
        },
      ],
    ])
    const svg = svgOf(
      '<w:pict><v:shape type="#_x0000_t110" style="width:1in;height:48pt" fillcolor="white [3201]"/></w:pict>',
      types,
    )
    expect(svg).toContain(
      '<path d="M 48 0.5 L 0.5 32 L 48 63.5 L 95.5 32 Z" fill="#FFFFFF" stroke="#FF0000"',
    )
    expect(
      vmlShapeSvg('<w:pict><v:shape type="#_x0000_t110" style="width:1in;height:48pt"/></w:pict>'),
    ).toBeNull()
  })

  it('falls back to the well-known o:spt geometry and honors <v:stroke on="f"/>', () => {
    const diamond = svgOf(
      '<w:pict><v:shape o:spt="4" style="width:20pt;height:20pt" fillcolor="blue"><v:stroke on="f"/></v:shape></w:pict>',
    )
    expect(diamond).toContain('<path d="M 13.5 0 L 0 13.5 L 13.5 27 L 27 13.5 Z" fill="#0000FF"/>')
    const ellipse = svgOf('<w:pict><v:shape o:spt="3" style="width:20pt;height:20pt"/></w:pict>')
    expect(ellipse).toContain('<ellipse ')
    // Word's ellipse shapetype carries a curved (qx/qy) path: the preset wins over it
    const typed = svgOf(
      '<w:pict><v:shapetype id="_x0000_t3" coordsize="21600,21600" o:spt="3" path="m10800,qx,10800,10800,21600,21600,10800,10800,xe"/>' +
        '<v:shape type="#_x0000_t3" style="width:20pt;height:10pt"/></w:pict>',
    )
    expect(typed).toContain('<ellipse cx="13.5" cy="6.5"')
  })
})

describe('readWatermarkShape', () => {
  it('keeps a zero fill opacity (fully transparent WordArt) instead of defaulting to opaque', () => {
    const pict =
      '<w:pict><v:shape type="#_x0000_t136" style="position:absolute;width:100pt;height:50pt" fillcolor="red">' +
      '<v:fill opacity="0"/><v:textpath string="GHOST"/></v:shape></w:pict>'
    expect(readWatermarkShape(pict)?.wordArt?.opacity).toBe(0)
    const noFill =
      '<w:pict><v:shape type="#_x0000_t136" style="position:absolute;width:100pt;height:50pt" fillcolor="red">' +
      '<v:textpath string="GHOST"/></v:shape></w:pict>'
    expect(readWatermarkShape(noFill)?.wordArt?.opacity).toBe(1)
  })
})

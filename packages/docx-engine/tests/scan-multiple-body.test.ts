import { describe, expect, it } from 'vitest'
import { scanBody } from '../src/scan'

describe('scanBody with multiple <w:body> elements', () => {
  it('keeps scanning past </w:body> into later sibling bodies (POI MultipleBodyBug)', () => {
    const xml =
      '<w:document>' +
      '<w:body><w:p><w:r><w:t>BODY 1</w:t></w:r></w:p></w:body>' +
      '<w:body><w:p><w:r><w:t>BODY 2</w:t></w:r></w:p></w:body>' +
      '<w:body><w:p><w:r><w:t>BODY 3</w:t></w:r></w:p></w:body>' +
      '</w:document>'
    const scan = scanBody(xml)
    expect(scan.elements.map((e) => e.name)).toEqual(['w:p', 'w:p', 'w:p'])
    expect(xml.slice(scan.elements[2].start, scan.elements[2].end)).toContain('BODY 3')
    expect(scan.innerEnd).toBe(scan.elements[2].end)
  })

  it('single body still terminates at its closing tag', () => {
    const xml = '<w:document><w:body><w:p/></w:body></w:document>'
    expect(scanBody(xml).elements.map((e) => e.name)).toEqual(['w:p'])
  })
})

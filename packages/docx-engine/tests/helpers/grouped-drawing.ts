// Synthetic small-coordinate drawing; no private document content.
export const GROUPED_DRAWING =
  '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="9525000" cy="4762500"/>' +
  '<wp:docPr id="1" name="Diagram"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
  '<wpg:wgp xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9525000" cy="4762500"/>' +
  '<a:chOff x="1000" y="2000"/><a:chExt cx="10000" cy="5000"/></a:xfrm></wpg:grpSpPr>' +
  box('10', 1200, 2200, 'Source') +
  box('11', 4200, 2200, 'Destination') +
  line('12', 3200, 2400, 1000, 0, '<a:tailEnd type="triangle" w="med" len="med"/>') +
  line('13', 2200, 2600, 0, 1000, '<a:prstDash val="dash"/>') +
  line('14', 2200, 3600, 1000, 600, '', ' cmpd="thickThin"') +
  '</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'

function box(id: string, x: number, y: number, text: string): string {
  return (
    `<wps:wsp><wps:cNvPr id="${id}" name="Box ${id}"/><wps:cNvSpPr txBox="1"/>` +
    `<wps:spPr bwMode="auto"><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="2000" cy="400"/></a:xfrm>` +
    '<a:prstGeom prst="rect"/><a:noFill/><a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>' +
    `<wps:txbx><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
    '<wps:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"><a:noAutofit/></wps:bodyPr></wps:wsp>'
  )
}

function line(
  id: string,
  x: number,
  y: number,
  w: number,
  h: number,
  ends: string,
  attrs = '',
): string {
  return (
    `<wps:wsp><wps:cNvPr id="${id}" name="Line ${id}"/><wps:cNvCnPr/>` +
    `<wps:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>` +
    `<a:prstGeom prst="line"/><a:noFill/><a:ln w="9525"${attrs}><a:solidFill><a:srgbClr val="000000"/></a:solidFill>${ends}</a:ln>` +
    '</wps:spPr><wps:bodyPr/></wps:wsp>'
  )
}

export const GROUPED_DRAWING_WITH_FALLBACK = GROUPED_DRAWING.replace(
  '<w:r>',
  '<w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><mc:Choice Requires="wpg">',
).replace(
  /<\/w:r><\/w:p>$/,
  '</mc:Choice><mc:Fallback><w:pict><v:group xmlns:v="urn:schemas-microsoft-com:vml" style="width:750pt;height:375pt" coordorigin="1000,2000" coordsize="10000,5000">' +
    '<v:shape id="Box 10" style="position:absolute;left:1200;top:2200;width:2000;height:400"><v:textbox><w:txbxContent><w:p><w:r><w:t>Source</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape>' +
    '<v:shape id="Box 11" style="position:absolute;left:4200;top:2200;width:2000;height:400"><v:textbox><w:txbxContent><w:p><w:r><w:t>Destination</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape>' +
    '<v:line id="Line 12" style="position:absolute" from="3200,2400" to="4200,2400"/>' +
    '</v:group></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>',
)
